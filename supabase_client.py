"""
Supabase PostgreSQL Integration & Auto-Sync Engine for LAMP Packing System
Provides real-time cloud synchronization between local SQLite and Supabase PostgreSQL.
Ensures zero data loss, offline resilience on factory floor, and real-time cloud backup.
"""

import os
import threading
import time
import datetime
import sqlite3

try:
    import psycopg2
    from psycopg2 import extras
    HAS_PSYCOPG2 = True
except ImportError:
    HAS_PSYCOPG2 = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SQLITE_DB_PATH = os.path.join(BASE_DIR, "lamp_system.db")

# Default Supabase PostgreSQL Pooler URI (IPv4 compatible, Port 5432 session or 6543 transaction)
DEFAULT_SUPABASE_URL = "postgresql://postgres.wwqijexgkicowmzvwqvy:KOCH-LAMP123@aws-0-ap-southeast-1.pooler.supabase.com:5432/postgres"

SUPABASE_URL = os.environ.get("SUPABASE_DB_URL") or os.environ.get("DATABASE_URL") or DEFAULT_SUPABASE_URL

# In case someone passes direct 'db.xxx.supabase.co:5432' which has IPv6 resolution issues on Windows,
# automatically convert to pooler host if needed:
if "db.wwqijexgkicowmzvwqvy.supabase.co" in SUPABASE_URL:
    SUPABASE_URL = SUPABASE_URL.replace("db.wwqijexgkicowmzvwqvy.supabase.co", "aws-0-ap-southeast-1.pooler.supabase.com")
    if "postgres:" in SUPABASE_URL and "postgres.wwqijexgkicowmzvwqvy" not in SUPABASE_URL:
        SUPABASE_URL = SUPABASE_URL.replace("postgres:", "postgres.wwqijexgkicowmzvwqvy:", 1)

# Global status tracking
_last_sync_time = None
_last_sync_status = "Not initialized"
_sync_lock = threading.Lock()


def get_supabase_conn(timeout=6):
    """Establishes a connection to Supabase PostgreSQL."""
    if not HAS_PSYCOPG2:
        raise RuntimeError("psycopg2-binary is not installed")
    return psycopg2.connect(SUPABASE_URL, connect_timeout=timeout)


def test_supabase_connection():
    """Tests connection to Supabase and returns (is_ok, message)."""
    if not HAS_PSYCOPG2:
        return False, "psycopg2 library not available"
    try:
        conn = get_supabase_conn(timeout=5)
        cur = conn.cursor()
        cur.execute("SELECT 1;")
        cur.close()
        conn.close()
        return True, "เชื่อมต่อ Supabase สำเร็จ (Connected)"
    except Exception as e:
        return False, f"เชื่อมต่อไม่ได้: {str(e)}"


def get_sync_status():
    """Returns the current synchronization status and record counts."""
    global _last_sync_time, _last_sync_status

    status = {
        "configured": bool(SUPABASE_URL),
        "connected": False,
        "message": "",
        "local": {"receive_batches": 0, "pack_scans": 0},
        "cloud": {"receive_batches": 0, "pack_scans": 0},
        "in_sync": False,
        "last_sync": _last_sync_time.strftime("%Y-%m-%d %H:%M:%S") if _last_sync_time else "-",
        "last_status": _last_sync_status
    }

    # 1. Local counts
    try:
        s_conn = sqlite3.connect(SQLITE_DB_PATH)
        s_cur = s_conn.cursor()
        s_cur.execute("SELECT COUNT(*) FROM receive_batches;")
        status["local"]["receive_batches"] = s_cur.fetchone()[0]
        s_cur.execute("SELECT COUNT(*) FROM pack_scans;")
        status["local"]["pack_scans"] = s_cur.fetchone()[0]
        s_conn.close()
    except Exception as e:
        status["message"] = f"Local DB error: {e}"
        return status

    # 2. Cloud counts
    try:
        p_conn = get_supabase_conn(timeout=5)
        p_cur = p_conn.cursor()
        p_cur.execute("SELECT COUNT(*) FROM receive_batches;")
        status["cloud"]["receive_batches"] = p_cur.fetchone()[0]
        p_cur.execute("SELECT COUNT(*) FROM pack_scans;")
        status["cloud"]["pack_scans"] = p_cur.fetchone()[0]
        p_conn.close()

        status["connected"] = True
        status["in_sync"] = (
            status["local"]["receive_batches"] == status["cloud"]["receive_batches"] and
            status["local"]["pack_scans"] == status["cloud"]["pack_scans"]
        )
        status["message"] = "ข้อมูลตรงกัน 100%" if status["in_sync"] else "มียอดต่างกัน รอการซิงค์"
    except Exception as e:
        status["connected"] = False
        status["message"] = f"Cloud connection failed: {e}"

    return status


def sync_all(direction="bidirectional"):
    """
    Synchronizes receive_batches and pack_scans between SQLite and Supabase.
    direction: 'bidirectional', 'push_to_cloud', or 'pull_from_cloud'
    """
    global _last_sync_time, _last_sync_status

    with _sync_lock:
        try:
            s_conn = sqlite3.connect(SQLITE_DB_PATH)
            s_conn.row_factory = sqlite3.Row
            s_cur = s_conn.cursor()

            p_conn = get_supabase_conn(timeout=10)
            p_cur = p_conn.cursor()

            pushed_batches = 0
            pushed_scans = 0
            pulled_batches = 0
            pulled_scans = 0

            # ------------------------------------------------------------------
            # 1. Sync receive_batches
            # ------------------------------------------------------------------
            s_cur.execute("SELECT id FROM receive_batches;")
            local_rb_ids = set(r[0] for r in s_cur.fetchall())

            p_cur.execute("SELECT id FROM receive_batches;")
            cloud_rb_ids = set(r[0] for r in p_cur.fetchall())

            # Push missing to cloud
            if direction in ("bidirectional", "push_to_cloud"):
                missing_in_cloud_rb = local_rb_ids - cloud_rb_ids
                if missing_in_cloud_rb:
                    placeholders = ",".join("?" for _ in missing_in_cloud_rb)
                    s_cur.execute(f"""
                        SELECT id, timestamp, grn, receive_date, invoice_no, po_no, line, 
                               part_no, part_name, qty_received, kanban, lot, rack_no, raw_scan 
                        FROM receive_batches WHERE id IN ({placeholders})
                    """, list(missing_in_cloud_rb))
                    rows_to_push = s_cur.fetchall()
                    if rows_to_push:
                        p_sql = """
                            INSERT INTO receive_batches 
                            (id, timestamp, grn, receive_date, invoice_no, po_no, line, 
                             part_no, part_name, qty_received, kanban, lot, rack_no, raw_scan)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                            ON CONFLICT (id) DO UPDATE SET
                                timestamp = EXCLUDED.timestamp,
                                grn = EXCLUDED.grn,
                                receive_date = EXCLUDED.receive_date,
                                invoice_no = EXCLUDED.invoice_no,
                                po_no = EXCLUDED.po_no,
                                line = EXCLUDED.line,
                                part_no = EXCLUDED.part_no,
                                part_name = EXCLUDED.part_name,
                                qty_received = EXCLUDED.qty_received,
                                kanban = EXCLUDED.kanban,
                                lot = EXCLUDED.lot,
                                rack_no = EXCLUDED.rack_no,
                                raw_scan = EXCLUDED.raw_scan;
                        """
                        extras.execute_batch(p_cur, p_sql, [tuple(r) for r in rows_to_push])
                        pushed_batches = len(rows_to_push)

            # Pull missing to local
            if direction in ("bidirectional", "pull_from_cloud"):
                missing_in_local_rb = cloud_rb_ids - local_rb_ids
                if missing_in_local_rb:
                    p_cur.execute(f"""
                        SELECT id, timestamp, grn, receive_date, invoice_no, po_no, line, 
                               part_no, part_name, qty_received, kanban, lot, rack_no, raw_scan 
                        FROM receive_batches WHERE id = ANY(%s)
                    """, (list(missing_in_local_rb),))
                    rows_to_pull = p_cur.fetchall()
                    if rows_to_pull:
                        s_cur.executemany("""
                            INSERT OR REPLACE INTO receive_batches 
                            (id, timestamp, grn, receive_date, invoice_no, po_no, line, 
                             part_no, part_name, qty_received, kanban, lot, rack_no, raw_scan)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, rows_to_pull)
                        s_conn.commit()
                        pulled_batches = len(rows_to_pull)

            # ------------------------------------------------------------------
            # 2. Sync pack_scans
            # ------------------------------------------------------------------
            s_cur.execute("SELECT id FROM pack_scans;")
            local_ps_ids = set(r[0] for r in s_cur.fetchall())

            p_cur.execute("SELECT id FROM pack_scans;")
            cloud_ps_ids = set(r[0] for r in p_cur.fetchall())

            # Push missing to cloud
            if direction in ("bidirectional", "push_to_cloud"):
                missing_in_cloud_ps = local_ps_ids - cloud_ps_ids
                if missing_in_cloud_ps:
                    placeholders = ",".join("?" for _ in missing_in_cloud_ps)
                    s_cur.execute(f"""
                        SELECT id, timestamp, invoice_no, po_no, line, part_scan, 
                               part_no, part_name, box_scan, box_expected, is_box_valid, 
                               rack_no, qty, remark 
                        FROM pack_scans WHERE id IN ({placeholders})
                    """, list(missing_in_cloud_ps))
                    rows_to_push = s_cur.fetchall()
                    if rows_to_push:
                        p_sql = """
                            INSERT INTO pack_scans 
                            (id, timestamp, invoice_no, po_no, line, part_scan, 
                             part_no, part_name, box_scan, box_expected, is_box_valid, 
                             rack_no, qty, remark)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                            ON CONFLICT (id) DO UPDATE SET
                                timestamp = EXCLUDED.timestamp,
                                invoice_no = EXCLUDED.invoice_no,
                                po_no = EXCLUDED.po_no,
                                line = EXCLUDED.line,
                                part_scan = EXCLUDED.part_scan,
                                part_no = EXCLUDED.part_no,
                                part_name = EXCLUDED.part_name,
                                box_scan = EXCLUDED.box_scan,
                                box_expected = EXCLUDED.box_expected,
                                is_box_valid = EXCLUDED.is_box_valid,
                                rack_no = EXCLUDED.rack_no,
                                qty = EXCLUDED.qty,
                                remark = EXCLUDED.remark;
                        """
                        extras.execute_batch(p_cur, p_sql, [tuple(r) for r in rows_to_push])
                        pushed_scans = len(rows_to_push)

            # Pull missing to local
            if direction in ("bidirectional", "pull_from_cloud"):
                missing_in_local_ps = cloud_ps_ids - local_ps_ids
                if missing_in_local_ps:
                    p_cur.execute(f"""
                        SELECT id, timestamp, invoice_no, po_no, line, part_scan, 
                               part_no, part_name, box_scan, box_expected, is_box_valid, 
                               rack_no, qty, remark 
                        FROM pack_scans WHERE id = ANY(%s)
                    """, (list(missing_in_local_ps),))
                    rows_to_pull = p_cur.fetchall()
                    if rows_to_pull:
                        s_cur.executemany("""
                            INSERT OR REPLACE INTO pack_scans 
                            (id, timestamp, invoice_no, po_no, line, part_scan, 
                             part_no, part_name, box_scan, box_expected, is_box_valid, 
                             rack_no, qty, remark)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, rows_to_pull)
                        s_conn.commit()
                        pulled_scans = len(rows_to_pull)

            # Fix sequences in postgres
            p_cur.execute("""
                SELECT setval('receive_batches_id_seq', COALESCE((SELECT MAX(id) FROM receive_batches), 1), true);
                SELECT setval('pack_scans_id_seq', COALESCE((SELECT MAX(id) FROM pack_scans), 1), true);
            """)

            p_conn.commit()
            p_conn.close()
            s_conn.close()

            _last_sync_time = datetime.datetime.now()
            _last_sync_status = f"สำเร็จ (ส่งขึ้นคลาวด์: {pushed_scans + pushed_batches}, ดึงลงมา: {pulled_scans + pulled_batches})"
            return {
                "success": True,
                "pushed_batches": pushed_batches,
                "pushed_scans": pushed_scans,
                "pulled_batches": pulled_batches,
                "pulled_scans": pulled_scans,
                "message": _last_sync_status
            }

        except Exception as e:
            _last_sync_status = f"ผิดพลาด: {str(e)}"
            return {
                "success": False,
                "error": str(e),
                "message": _last_sync_status
            }


def async_push_pack_scan(scan_dict):
    """Asynchronously pushes a newly recorded pack scan to Supabase without blocking caller."""
    def _worker():
        try:
            conn = get_supabase_conn(timeout=4)
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO pack_scans (
                    id, timestamp, invoice_no, po_no, line, part_scan, part_no, 
                    part_name, box_scan, box_expected, is_box_valid, rack_no, qty, remark
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (id) DO UPDATE SET
                    timestamp = EXCLUDED.timestamp,
                    invoice_no = EXCLUDED.invoice_no,
                    po_no = EXCLUDED.po_no,
                    line = EXCLUDED.line,
                    part_scan = EXCLUDED.part_scan,
                    part_no = EXCLUDED.part_no,
                    part_name = EXCLUDED.part_name,
                    box_scan = EXCLUDED.box_scan,
                    box_expected = EXCLUDED.box_expected,
                    is_box_valid = EXCLUDED.is_box_valid,
                    rack_no = EXCLUDED.rack_no,
                    qty = EXCLUDED.qty,
                    remark = EXCLUDED.remark;
            """, (
                scan_dict.get("id"),
                scan_dict.get("timestamp"),
                scan_dict.get("invoice_no"),
                scan_dict.get("po_no"),
                scan_dict.get("line"),
                scan_dict.get("part_scan"),
                scan_dict.get("part_no"),
                scan_dict.get("part_name"),
                scan_dict.get("box_scan"),
                scan_dict.get("box_expected"),
                1 if scan_dict.get("is_box_valid") else 0,
                scan_dict.get("rack_no"),
                scan_dict.get("qty", 1),
                scan_dict.get("remark")
            ))
            conn.commit()
            conn.close()
        except Exception as e:
            # Fallback will be caught by sync_all
            pass
    threading.Thread(target=_worker, daemon=True).start()


def async_push_receive_batch(batch_dict):
    """Asynchronously pushes a newly recorded receive batch to Supabase."""
    def _worker():
        try:
            conn = get_supabase_conn(timeout=4)
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO receive_batches (
                    id, timestamp, grn, receive_date, invoice_no, po_no, line, 
                    part_no, part_name, qty_received, kanban, lot, rack_no, raw_scan
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (id) DO UPDATE SET
                    timestamp = EXCLUDED.timestamp,
                    grn = EXCLUDED.grn,
                    receive_date = EXCLUDED.receive_date,
                    invoice_no = EXCLUDED.invoice_no,
                    po_no = EXCLUDED.po_no,
                    line = EXCLUDED.line,
                    part_no = EXCLUDED.part_no,
                    part_name = EXCLUDED.part_name,
                    qty_received = EXCLUDED.qty_received,
                    kanban = EXCLUDED.kanban,
                    lot = EXCLUDED.lot,
                    rack_no = EXCLUDED.rack_no,
                    raw_scan = EXCLUDED.raw_scan;
            """, (
                batch_dict.get("id"),
                batch_dict.get("timestamp"),
                batch_dict.get("grn"),
                batch_dict.get("receive_date"),
                batch_dict.get("invoice_no"),
                batch_dict.get("po_no"),
                batch_dict.get("line"),
                batch_dict.get("part_no"),
                batch_dict.get("part_name"),
                batch_dict.get("qty_received"),
                batch_dict.get("kanban"),
                batch_dict.get("lot"),
                batch_dict.get("rack_no"),
                batch_dict.get("raw_scan")
            ))
            conn.commit()
            conn.close()
        except Exception as e:
            pass
    threading.Thread(target=_worker, daemon=True).start()


def async_delete_pack_scan(scan_id):
    """Asynchronously deletes a scan in Supabase."""
    def _worker():
        try:
            conn = get_supabase_conn(timeout=4)
            cur = conn.cursor()
            cur.execute("DELETE FROM pack_scans WHERE id = %s;", (scan_id,))
            conn.commit()
            conn.close()
        except Exception:
            pass
    threading.Thread(target=_worker, daemon=True).start()


def async_delete_receive_batch(batch_id):
    """Asynchronously deletes a receive batch in Supabase."""
    def _worker():
        try:
            conn = get_supabase_conn(timeout=4)
            cur = conn.cursor()
            cur.execute("DELETE FROM receive_batches WHERE id = %s;", (batch_id,))
            conn.commit()
            conn.close()
        except Exception:
            pass
    threading.Thread(target=_worker, daemon=True).start()


def start_background_sync(interval_sec=60):
    """Starts a background periodic auto-sync worker thread."""
    def _worker():
        while True:
            time.sleep(interval_sec)
            try:
                sync_all(direction="bidirectional")
            except Exception:
                pass
    t = threading.Thread(target=_worker, daemon=True, name="SupabaseAutoSyncWorker")
    t.start()
    return t


def trigger_sync():
    """Triggers an immediate background synchronization."""
    t = threading.Thread(target=sync_all, kwargs={"direction": "bidirectional"}, daemon=True)
    t.start()
    return t

