"""
Flask Web Application for LAMP Packing & Delivery Verification System
Runs locally on Windows: http://127.0.0.1:5000
"""

import sys
import os
import datetime
import threading
import subprocess
import re
import atexit

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass
from flask import Flask, render_template, request, jsonify, send_file, make_response
from lamp_core import LampDataManager, BASE_DIR, is_same_package
import supabase_client

app = Flask(__name__, template_folder=os.path.join(BASE_DIR, "templates"))
app.config['TEMPLATES_AUTO_RELOAD'] = True
mgr = LampDataManager(workspace_dir=BASE_DIR)

# Initialize Supabase auto-sync in background (every 45 seconds)
try:
    supabase_client.start_background_sync(interval_sec=45)
except Exception:
    pass

# Cloudflare HTTPS Tunnel controller for secure camera access
CLOUDFLARE_URL = None
CLOUDFLARE_PROC = None

def start_cloudflare_tunnel(port):
    global CLOUDFLARE_URL, CLOUDFLARE_PROC
    cf_exe = os.path.join(BASE_DIR, "cloudflared.exe")
    if not os.path.exists(cf_exe):
        return

    def _tunnel_worker():
        global CLOUDFLARE_URL, CLOUDFLARE_PROC
        url_regex = re.compile(r'https://[a-zA-Z0-9-]+\.trycloudflare\.com')
        creationflags = 0
        if sys.platform == "win32":
            creationflags = subprocess.CREATE_NO_WINDOW

        import time
        while True:
            try:
                CLOUDFLARE_PROC = subprocess.Popen(
                    [cf_exe, "tunnel", "--url", f"http://127.0.0.1:{port}"],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    bufsize=1,
                    creationflags=creationflags
                )
                for line in CLOUDFLARE_PROC.stderr:
                    m = url_regex.search(line)
                    if m:
                        CLOUDFLARE_URL = m.group(0)
                        print(f"\n   [SSL / HTTPS Live Camera] 🔒 {CLOUDFLARE_URL}\n")
                        break
                # Keep process monitored
                if CLOUDFLARE_PROC:
                    CLOUDFLARE_PROC.wait()
            except Exception as ex:
                print("Cloudflare tunnel notice:", ex)
            time.sleep(5)

    t = threading.Thread(target=_tunnel_worker, daemon=True)
    t.start()

def cleanup_tunnel():
    global CLOUDFLARE_PROC
    if CLOUDFLARE_PROC:
        try:
            CLOUDFLARE_PROC.terminate()
        except Exception:
            pass

atexit.register(cleanup_tunnel)



@app.route("/")
def index():
    resp = make_response(render_template("index.html"))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


@app.route("/wi")
def work_instruction():
    return render_template("wi.html")


@app.route("/api/balance", methods=["GET"])
def get_balance():
    po_param = request.args.get("po_no")
    inv_param = request.args.get("invoice_no")

    target_po = None
    target_inv = None

    # Handle combined parameter "po_no|invoice_no" or "WAIT_PO"
    if str(po_param or "").upper() == "WAIT_PO" or str(inv_param or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_inv = "WAIT_PO"
    elif po_param:
        if "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    # Default to PO 229716 / Invoice 11726100219 (received 08/10/2026) if neither provided
    if not po_param and not inv_param:
        target_po = "229716"
        target_inv = "11726100219"

    lines = mgr.get_po_line_balance(invoice_no=target_inv, po_no=target_po)

    tot_rec = sum(L["received_qty"] for L in lines)
    tot_pack = sum(L["packed_qty"] for L in lines)
    tot_rem = sum(L["remaining_qty"] for L in lines)
    pct = round((tot_pack / tot_rec * 100), 1) if tot_rec > 0 else 0

    # Box summary
    actual_boxes = mgr.get_packing_box_usage(invoice_no=target_inv, po_no=target_po)
    tot_actual_boxes = sum(b["qty"] for b in actual_boxes if "ไม่มีกล่อง" not in b["box_no"])

    # BOM planned boxes per part
    box_calc = {}
    for L in lines:
        b_no = L["box_no"]
        b_name = L["box_name"]
        if b_no and b_no != "-":
            if b_no not in box_calc:
                box_calc[b_no] = {"box_no": b_no, "box_name": b_name, "qty_needed": 0, "qty_packed": 0}
            box_calc[b_no]["qty_needed"] += L["received_qty"]
            box_calc[b_no]["qty_packed"] += L["packed_qty"]

    box_summary_list = sorted(box_calc.values(), key=lambda x: -x["qty_needed"])

    return jsonify({
        "status": "SUCCESS",
        "lines": lines,
        "summary": {
            "total_received": tot_rec,
            "total_packed": tot_pack,
            "total_remaining": tot_rem,
            "percent_complete": pct,
            "total_boxes": tot_actual_boxes or sum(b["qty_packed"] for b in box_summary_list),
        },
        "pending_receives": mgr.get_pending_receives_summary(),
        "box_summary": box_summary_list,
        "actual_box_usage": actual_boxes,
        "pos": mgr.get_pos_list(),
        "invoices": mgr.get_invoices_list(),
        "current_po": target_po or "ALL",
        "current_inv": target_inv or "ALL",
    })


@app.route("/api/pos", methods=["GET"])
def get_pos():
    status = request.args.get("status")
    exclude_comp_raw = str(request.args.get("exclude_completed") or "false").lower()
    exclude_completed = exclude_comp_raw in ("true", "1", "yes")

    pos_data = mgr.get_pos_list(status_filter=status, exclude_completed=exclude_completed)
    all_pos = mgr.get_pos_list(status_filter=None, exclude_completed=False)
    real_pos = [p for p in all_pos if not p.get("is_wait_po")]

    completed_cnt = sum(1 for p in real_pos if p.get("is_complete"))
    pending_cnt = sum(1 for p in real_pos if not p.get("is_complete"))
    in_progress_cnt = sum(1 for p in real_pos if p.get("status") == "IN_PROGRESS")

    return jsonify({
        "status": "SUCCESS",
        "pos": pos_data,
        "total_count": len(real_pos),
        "completed_count": completed_cnt,
        "pending_count": pending_cnt,
        "in_progress_count": in_progress_cnt,
        "exclude_completed": exclude_completed,
        "status_filter": status or "all"
    })


@app.route("/api/check_po_status", methods=["GET", "POST"])
def check_po_status():
    if request.method == "POST":
        data = request.get_json() or {}
        po_no = data.get("po_no")
        invoice_no = data.get("invoice_no")
    else:
        po_no = request.args.get("po_no")
        invoice_no = request.args.get("invoice_no")

    res = mgr.check_po_completion(po_no=po_no, invoice_no=invoice_no)
    return jsonify(res)


@app.route("/api/pos_status_summary", methods=["GET"])
def pos_status_summary():
    status = request.args.get("status")
    exclude_comp_raw = str(request.args.get("exclude_completed") or "false").lower()
    exclude_completed = exclude_comp_raw in ("true", "1", "yes")
    res = mgr.get_all_pos_status_summary(exclude_completed=exclude_completed, status_filter=status)
    return jsonify(res)


def trigger_background_plan_sync(target_date=None):
    """Triggers background synchronization to Plan Production (port 5000) non-blockingly."""
    try:
        import threading
        d = target_date or datetime.date.today().strftime("%Y-%m-%d")
        t = threading.Thread(target=mgr.sync_to_plan_production, kwargs={"target_date": d}, daemon=True)
        t.start()
    except Exception:
        pass


def start_wms_periodic_sync():
    """Starts background periodic auto-sync from WMS (every 10 minutes)."""
    def _worker():
        import time
        time.sleep(5)  # Initial delay on start
        while True:
            try:
                mgr.sync_wms_receives()
            except Exception as e:
                print("WMS Auto-sync notice:", e)
            time.sleep(600)  # Every 10 minutes
    t = threading.Thread(target=_worker, daemon=True)
    t.start()


@app.route("/api/scan_pack", methods=["POST"])
def scan_pack():
    data = request.get_json() or {}
    part_scan = data.get("part_scan", "").strip()
    box_scan = data.get("box_scan", "").strip()
    rack_no = data.get("rack_no", "").strip()
    target_invoice = data.get("target_invoice")
    target_po = data.get("target_po")

    # Support combined "po|inv" or "WAIT_PO"
    if str(target_po or "").upper() == "WAIT_PO" or str(target_invoice or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_invoice = "WAIT_PO"
    elif target_po and "|" in str(target_po):
        pts = str(target_po).split("|")
        target_po = pts[0] if pts[0] != "ALL" else None
        if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-":
            target_invoice = pts[1]
    elif target_invoice and "|" in str(target_invoice):
        pts = str(target_invoice).split("|")
        target_po = pts[0] if pts[0] != "ALL" else None
        if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-":
            target_invoice = pts[1]

    if not part_scan:
        return jsonify({"status": "ERROR", "message": "กรุณาสแกนบาร์โค้ด Part No."}), 400

    result = mgr.process_pack_scan(
        part_scan=part_scan,
        box_scan=box_scan,
        rack_no=rack_no,
        target_invoice=target_invoice,
        target_po=target_po,
    )
    if result.get("status") in ("SUCCESS", "PASS"):
        trigger_background_plan_sync()
    return jsonify(result)


@app.route("/api/scan_receive", methods=["POST"])
def scan_receive():
    data = request.get_json() or {}
    raw_scan = data.get("raw_scan", "").strip()
    qty_box = int(data.get("qty_box", 2))
    kanban = data.get("kanban", "").strip()
    rack_no = data.get("rack_no", "").strip()
    invoice_no = data.get("invoice_no")
    po_no = data.get("po_no")
    line = data.get("line")
    link_po_later = data.get("link_po_later", True)

    if str(invoice_no or "").upper() == "WAIT_PO" or str(po_no or "").upper() == "WAIT_PO":
        invoice_no = "WAIT_PO"
        po_no = "WAIT_PO"

    if not raw_scan:
        return jsonify({"status": "ERROR", "message": "กรุณาระบุข้อมูลที่สแกน"}), 400

    parsed = mgr.parse_receive_tag(raw_scan)
    parsed["qty_box"] = qty_box
    if kanban:
        parsed["kanban"] = kanban

    result = mgr.record_receive(
        parsed_tag=parsed,
        invoice_no=invoice_no,
        po_no=po_no,
        line=line,
        link_po_later=link_po_later,
        rack_no=rack_no,
    )
    return jsonify(result)


@app.route("/api/rack_summary", methods=["GET"])
def get_rack_summary_api():
    rack_no = request.args.get("rack_no", "").strip()
    po_param = request.args.get("po_no")
    inv_param = request.args.get("invoice_no")

    target_po = None
    target_inv = None
    if str(po_param or "").upper() == "WAIT_PO" or str(inv_param or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_inv = "WAIT_PO"
    elif po_param:
        if "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    summary = mgr.get_rack_summary(rack_no=rack_no, po_no=target_po, invoice_no=target_inv)
    return jsonify(summary)


@app.route("/api/rack_scans", methods=["GET"])
def get_rack_scans_api():
    rack_no = request.args.get("rack_no", "").strip()
    limit = int(request.args.get("limit", 100))
    scans = mgr.get_rack_detail_scans(rack_no=rack_no, limit=limit)
    return jsonify({"status": "SUCCESS", "rack_no": rack_no, "scans": scans})


@app.route("/api/racks_parts_summary", methods=["GET"])
def get_racks_parts_summary_api():
    po_param = request.args.get("po_no", "").strip()
    inv_param = request.args.get("invoice_no", "").strip()

    target_po = None
    target_inv = None

    if po_param:
        if po_param.upper() == "WAIT_PO":
            target_po = "WAIT_PO"
        elif "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    summary = mgr.get_racks_parts_summary(po_no=target_po, invoice_no=target_inv)
    return jsonify(summary)


@app.route("/api/rack_package_label", methods=["GET"])
def get_rack_package_label_api():
    rack_no = request.args.get("rack_no", "").strip()
    po_param = request.args.get("po_no")
    inv_param = request.args.get("invoice_no")

    target_po = None
    target_inv = None
    if str(po_param or "").upper() == "WAIT_PO" or str(inv_param or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_inv = "WAIT_PO"
    elif po_param:
        if "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    data = mgr.get_rack_package_label_data(rack_no=rack_no, po_no=target_po, invoice_no=target_inv)
    return jsonify(data)


@app.route("/print_rack_label", methods=["GET"])
def print_rack_label_view():
    rack_no = request.args.get("rack_no", "").strip()
    po_param = request.args.get("po_no")
    inv_param = request.args.get("invoice_no")

    target_po = None
    target_inv = None
    if str(po_param or "").upper() == "WAIT_PO" or str(inv_param or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_inv = "WAIT_PO"
    elif po_param:
        if "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    data = mgr.get_rack_package_label_data(rack_no=rack_no, po_no=target_po, invoice_no=target_inv)
    all_data = mgr.get_rack_package_label_data(rack_no="ALL", po_no=target_po, invoice_no=target_inv)
    all_racks = [l["rack_no"] for l in all_data.get("labels", [])]

    resp = make_response(render_template(
        "package_label_print.html",
        shipper=data.get("shipper", "KOCH"),
        ship_to=data.get("ship_to", "Mazda (MST)"),
        supplier_code=data.get("supplier_code", "KT076"),
        supplier_name=data.get("supplier_name", "THAI KOITO CO.,LTD."),
        labels=data.get("labels", []),
        selected_rack=rack_no if rack_no and rack_no != "ALL" else None,
        target_po=target_po,
        target_inv=target_inv,
        all_racks_list=all_racks
    ))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp


@app.route("/export_rack_label_excel", methods=["GET"])
def export_rack_label_excel_view():
    rack_no = request.args.get("rack_no", "").strip()
    po_param = request.args.get("po_no")
    inv_param = request.args.get("invoice_no")

    target_po = None
    target_inv = None
    if str(po_param or "").upper() == "WAIT_PO" or str(inv_param or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_inv = "WAIT_PO"
    elif po_param:
        if "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    filepath = mgr.export_rack_package_labels_excel(
        rack_no=rack_no,
        po_no=target_po,
        invoice_no=target_inv
    )
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))
    return "Export failed", 500


@app.route("/api/delivery_trip_sheet", methods=["GET"])
def get_delivery_trip_sheet_api():
    po_param = request.args.get("po_no")
    inv_param = request.args.get("invoice_no")

    target_po = None
    target_inv = None
    if str(po_param or "").upper() == "WAIT_PO" or str(inv_param or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_inv = "WAIT_PO"
    elif po_param:
        if "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    data = mgr.get_delivery_trip_sheet_data(po_no=target_po, invoice_no=target_inv)
    return jsonify(data)


@app.route("/print_delivery_trip_sheet", methods=["GET"])
def print_delivery_trip_sheet_view():
    po_param = request.args.get("po_no")
    inv_param = request.args.get("invoice_no")

    target_po = None
    target_inv = None
    if str(po_param or "").upper() == "WAIT_PO" or str(inv_param or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_inv = "WAIT_PO"
    elif po_param:
        if "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    data = mgr.get_delivery_trip_sheet_data(po_no=target_po, invoice_no=target_inv)
    invoices = mgr.get_invoices_list()
    try:
        copies = int(request.args.get("copies", 2))
    except (ValueError, TypeError):
        copies = 2

    resp = make_response(render_template(
        "delivery_trip_sheet_print.html",
        shipper=data.get("shipper", "KOCH"),
        ship_date=data.get("ship_date", datetime.date.today().strftime("%d-%m-%y")),
        invoice_no=data.get("invoice_no", "11726091882"),
        ship_to=data.get("ship_to", "Mazda (MST)"),
        trip_number=data.get("trip_number", 1),
        supplier_name=data.get("supplier_name", "THAI KOITO CO.,LTD."),
        trip_sheet_number=data.get("trip_sheet_number", f"MST_{datetime.date.today().strftime('%Y%m%d')}-001"),
        delivery_note_no=data.get("delivery_note_no", f"DN{datetime.date.today().strftime('%y%m%d0001')}"),
        items=data.get("items", []),
        total_qty=data.get("total_qty", 0),
        target_po=target_po,
        target_inv=target_inv,
        available_invoices=invoices,
        copies=copies
    ))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp


@app.route("/export_delivery_trip_sheet_excel", methods=["GET"])
def export_delivery_trip_sheet_excel_view():
    po_param = request.args.get("po_no")
    inv_param = request.args.get("invoice_no")

    target_po = None
    target_inv = None
    if str(po_param or "").upper() == "WAIT_PO" or str(inv_param or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_inv = "WAIT_PO"
    elif po_param:
        if "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    try:
        copies = int(request.args.get("copies", 2))
    except (ValueError, TypeError):
        copies = 2

    filepath = mgr.export_delivery_trip_sheet_single_excel(
        po_no=target_po,
        invoice_no=target_inv,
        copies=copies
    )
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))
    return "Export failed", 500


@app.route("/api/dn_delivery_note", methods=["GET"])
def get_dn_delivery_note_api():
    po_param = request.args.get("po_no")
    inv_param = request.args.get("invoice_no")

    target_po = None
    target_inv = None
    if str(po_param or "").upper() == "WAIT_PO" or str(inv_param or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_inv = "WAIT_PO"
    elif po_param:
        if "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    data = mgr.get_dn_delivery_note_data(po_no=target_po, invoice_no=target_inv)
    return jsonify(data)


@app.route("/print_dn_delivery_note", methods=["GET"])
def print_dn_delivery_note_view():
    po_param = request.args.get("po_no")
    inv_param = request.args.get("invoice_no")

    target_po = None
    target_inv = None
    if str(po_param or "").upper() == "WAIT_PO" or str(inv_param or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_inv = "WAIT_PO"
    elif po_param:
        if "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    data = mgr.get_dn_delivery_note_data(po_no=target_po, invoice_no=target_inv)
    invoices = mgr.get_invoices_list()
    try:
        copies = int(request.args.get("copies", 1))
    except (ValueError, TypeError):
        copies = 1

    resp = make_response(render_template(
        "dn_delivery_note_print.html",
        company_name=data.get("company_name", "KOCH PACKAGING AND PACKING SERVICES"),
        supplier_code=data.get("supplier_code", "KT076"),
        supplier_name=data.get("supplier_name", "THAI KOITO CO.,LTD."),
        invoice_no=data.get("invoice_no", "11726091882"),
        dn_no=data.get("dn_no", "DN2610070001"),
        delivery_date=data.get("delivery_date", datetime.date.today().strftime("%d-%m-%y")),
        items=data.get("items", []),
        total_invoice_qty=data.get("total_invoice_qty", 0),
        total_dom_qty=data.get("total_dom_qty", 0),
        sign_issuer=data.get("sign_issuer", "KOCH"),
        sign_receiver=data.get("sign_receiver", "MLYA"),
        target_po=target_po,
        target_inv=target_inv,
        available_invoices=invoices,
        copies=copies
    ))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp


@app.route("/export_dn_delivery_note_excel", methods=["GET"])
def export_dn_delivery_note_excel_view():
    po_param = request.args.get("po_no")
    inv_param = request.args.get("invoice_no")

    target_po = None
    target_inv = None
    if str(po_param or "").upper() == "WAIT_PO" or str(inv_param or "").upper() == "WAIT_PO":
        target_po = "WAIT_PO"
        target_inv = "WAIT_PO"
    elif po_param:
        if "|" in po_param:
            pts = po_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif po_param != "ALL":
            target_po = po_param

    if target_po != "WAIT_PO" and inv_param:
        if "|" in inv_param:
            pts = inv_param.split("|")
            target_po = pts[0] if pts[0] != "ALL" else None
            target_inv = pts[1] if len(pts) > 1 and pts[1] != "ALL" and pts[1] != "-" else None
        elif inv_param != "ALL":
            target_inv = inv_param

    try:
        copies = int(request.args.get("copies", 1))
    except (ValueError, TypeError):
        copies = 1

    filepath = mgr.export_dn_delivery_note_single_excel(
        po_no=target_po,
        invoice_no=target_inv,
        copies=copies
    )
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))
    return "Export failed", 500



@app.route("/api/sync_wms", methods=["POST", "GET"])
def sync_wms():
    res = mgr.sync_wms_receives()
    if res.get("status") == "SUCCESS":
        res["pos"] = mgr.get_pos_list()
    return jsonify(res)


@app.route("/api/receive_batches", methods=["GET"])
def get_receive_batches():
    status_filter = request.args.get("status")
    limit = int(request.args.get("limit", 200))
    batches = mgr.get_receive_batches(status_filter=status_filter, limit=limit)
    summary = mgr.get_pending_receives_summary()
    return jsonify({
        "status": "SUCCESS",
        "batches": batches,
        "pending_summary": summary,
    })


@app.route("/api/link_receive_po", methods=["POST"])
def link_receive_po():
    data = request.get_json() or {}
    batch_ids = data.get("batch_ids", [])
    po_no = data.get("po_no")
    line = data.get("line", 0)
    invoice_no = data.get("invoice_no", "")

    if not batch_ids:
        return jsonify({"status": "ERROR", "message": "กรุณาเลือกรายการรับเข้าที่ต้องการเชื่อมต่อ PO"}), 400
    if not po_no or int(po_no) <= 0:
        return jsonify({"status": "ERROR", "message": "กรุณาระบุเลข PO ที่ถูกต้อง"}), 400

    result = mgr.link_receive_batches(
        batch_ids=batch_ids,
        po_no=int(po_no),
        line=int(line or 0),
        invoice_no=str(invoice_no or "").strip()
    )
    return jsonify(result)


@app.route("/api/auto_link_receive_fifo", methods=["POST"])
def auto_link_receive_fifo():
    data = request.get_json() or {}
    batch_ids = data.get("batch_ids")
    result = mgr.auto_link_receive_fifo(batch_ids=batch_ids)
    return jsonify(result)


@app.route("/api/available_pos", methods=["GET"])
def get_available_pos():
    part_no = request.args.get("part_no", "").strip()
    pos = mgr.get_available_pos_for_part(part_no)
    return jsonify({
        "status": "SUCCESS",
        "part_no": part_no,
        "pos": pos
    })


@app.route("/api/scan_qr_image", methods=["POST"])
def scan_qr_image():
    """Decodes QR code and barcodes from an uploaded or captured image file."""
    if "file" not in request.files:
        return jsonify({"status": "ERROR", "message": "No file uploaded"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"status": "ERROR", "message": "Empty filename"}), 400

    try:
        from PIL import Image, ImageOps, ImageEnhance
        import pyzbar.pyzbar as pyzbar
        import io

        raw_bytes = file.read()
        img = Image.open(io.BytesIO(raw_bytes))

        # 1. EXIF orientation correction (critical for photos taken vertically on smartphones)
        try:
            img = ImageOps.exif_transpose(img)
        except Exception:
            pass

        # 2. Downscale if very large (e.g. 12-48MP smartphone camera) to ensure fast and accurate scan
        max_dim = 1800
        if max(img.width, img.height) > max_dim:
            ratio = max_dim / float(max(img.width, img.height))
            new_size = (int(img.width * ratio), int(img.height * ratio))
            img = img.resize(new_size, Image.LANCZOS)

        # 3. Direct decode on original orientation
        results = pyzbar.decode(img)

        # 4. If not found, try grayscale and multi-contrast / multi-scale passes
        if not results:
            gray = ImageOps.grayscale(img)
            # Try contrast enhanced
            for contrast in [1.5, 2.0, 0.8]:
                enh = ImageEnhance.Contrast(gray).enhance(contrast)
                results = pyzbar.decode(enh)
                if results:
                    break
                for s in [1.5, 0.7]:
                    rescaled = enh.resize((int(enh.width * s), int(enh.height * s)), Image.LANCZOS)
                    results = pyzbar.decode(rescaled)
                    if results:
                        break
                if results:
                    break

        # 5. If still not found, try 90 and 270 degree rotation
        if not results:
            for rot in [Image.ROTATE_90, Image.ROTATE_270]:
                rotated = img.transpose(rot)
                results = pyzbar.decode(rotated)
                if results:
                    break

        decoded_items = []
        for d in results:
            text = d.data.decode("utf-8", "ignore")
            decoded_items.append({"type": d.type, "text": text})

        if not decoded_items:
            return jsonify({"status": "WARN", "message": "ไม่พบ QR Code หรือ Barcode ในรูปภาพที่ถ่าย กรุณาจัดระยะให้ชัดเจนแล้วลองอีกครั้ง"}), 200

        # Primary decoded string
        primary = decoded_items[0]["text"]
        # If there are multiple barcodes (e.g. Tag 2 has both Part and Box), extract both
        part_candidate = None
        box_candidate = None
        for it in decoded_items:
            t = it["text"].upper().replace("-", "").strip()
            if "C302515" in t or "C453029" in t or "C522520" in t or "C814130" in t or "C211311" in t:
                box_candidate = it["text"]
            else:
                part_candidate = it["text"]

        return jsonify({
            "status": "SUCCESS",
            "count": len(decoded_items),
            "items": decoded_items,
            "primary_text": primary,
            "part_candidate": part_candidate or primary,
            "box_candidate": box_candidate,
        })
    except Exception as e:
        return jsonify({"status": "ERROR", "message": str(e)}), 500



@app.route("/api/recent_scans", methods=["GET"])
def get_recent_scans():
    limit = int(request.args.get("limit", 25))
    scans = mgr.get_recent_scans(limit=limit)
    return jsonify({"status": "SUCCESS", "scans": scans})


@app.route("/api/delete_pack_scan", methods=["POST"])
def delete_pack_scan():
    data = request.get_json() or {}
    scan_id = data.get("scan_id")
    if not scan_id:
        return jsonify({"status": "ERROR", "message": "ไม่ได้ระบุ ID รายการที่ต้องการลบ"}), 400
    res = mgr.delete_pack_scan(int(scan_id))
    if res.get("status") == "SUCCESS":
        trigger_background_plan_sync()
    return jsonify(res)


@app.route("/api/undo_pack_scan", methods=["POST"])
def undo_pack_scan():
    data = request.get_json() or {}
    target_invoice = data.get("target_invoice")
    res = mgr.delete_last_pack_scan(target_invoice=target_invoice)
    if res.get("status") == "SUCCESS":
        trigger_background_plan_sync()
    return jsonify(res)


@app.route("/api/edit_pack_scan", methods=["POST"])
def edit_pack_scan():
    data = request.get_json() or {}
    scan_id = data.get("scan_id")
    if not scan_id:
        return jsonify({"status": "ERROR", "message": "ไม่ได้ระบุ ID รายการที่ต้องการแก้ไข"}), 400
    res = mgr.edit_pack_scan(
        scan_id=int(scan_id),
        box_scan=data.get("box_scan"),
        rack_no=data.get("rack_no"),
        qty=data.get("qty"),
        po_no=data.get("po_no"),
        line=data.get("line"),
        invoice_no=data.get("invoice_no"),
    )
    if res.get("status") == "SUCCESS":
        trigger_background_plan_sync()
    return jsonify(res)


@app.route("/api/rename_rack", methods=["POST"])
def rename_rack_api():
    data = request.get_json() or {}
    old_rack = data.get("old_rack_no")
    new_rack = data.get("new_rack_no")
    po_no = data.get("po_no")
    invoice_no = data.get("invoice_no")
    res = mgr.rename_rack(old_rack_no=old_rack, new_rack_no=new_rack, po_no=po_no, invoice_no=invoice_no)
    if res.get("status") == "SUCCESS":
        trigger_background_plan_sync()
    return jsonify(res)


@app.route("/api/delete_rack", methods=["POST"])
def delete_rack_api():
    data = request.get_json() or {}
    rack_no = data.get("rack_no")
    po_no = data.get("po_no")
    invoice_no = data.get("invoice_no")
    res = mgr.delete_rack_scans(rack_no=rack_no, po_no=po_no, invoice_no=invoice_no)
    if res.get("status") == "SUCCESS":
        trigger_background_plan_sync()
    return jsonify(res)


@app.route("/api/edit_part_in_rack", methods=["POST"])
def edit_part_in_rack_api():
    data = request.get_json() or {}
    rack_no = data.get("rack_no")
    part_no = data.get("part_no")
    new_rack = data.get("new_rack_no")
    new_qty = data.get("new_qty")
    po_no = data.get("po_no")
    invoice_no = data.get("invoice_no")
    res = mgr.edit_part_in_rack(
        rack_no=rack_no,
        part_no=part_no,
        new_rack_no=new_rack,
        new_qty=new_qty,
        po_no=po_no,
        invoice_no=invoice_no
    )
    if res.get("status") == "SUCCESS":
        trigger_background_plan_sync()
    return jsonify(res)


@app.route("/api/delete_part_from_rack", methods=["POST"])
def delete_part_from_rack_api():
    data = request.get_json() or {}
    rack_no = data.get("rack_no")
    part_no = data.get("part_no")
    po_no = data.get("po_no")
    invoice_no = data.get("invoice_no")
    res = mgr.delete_part_from_rack(rack_no=rack_no, part_no=part_no, po_no=po_no, invoice_no=invoice_no)
    if res.get("status") == "SUCCESS":
        trigger_background_plan_sync()
    return jsonify(res)


@app.route("/api/delete_receive_batch", methods=["POST"])
def delete_receive_batch():
    data = request.get_json() or {}
    batch_ids = data.get("batch_ids") or []
    if not batch_ids and data.get("batch_id"):
        batch_ids = [data.get("batch_id")]
    if not batch_ids:
        return jsonify({"status": "ERROR", "message": "ไม่ได้ระบุรายการที่ต้องการลบ"}), 400
    res = mgr.delete_receive_batches(batch_ids)
    return jsonify(res)


@app.route("/api/edit_receive_batch", methods=["POST"])
def edit_receive_batch():
    data = request.get_json() or {}
    batch_id = data.get("batch_id")
    if not batch_id:
        return jsonify({"status": "ERROR", "message": "ไม่ได้ระบุ ID รายการที่ต้องการแก้ไข"}), 400
    res = mgr.edit_receive_batch(
        batch_id=int(batch_id),
        part_no=data.get("part_no"),
        part_name=data.get("part_name"),
        qty_received=data.get("qty_received"),
        kanban=data.get("kanban"),
        po_no=data.get("po_no"),
        line=data.get("line"),
        invoice_no=data.get("invoice_no"),
        rack_no=data.get("rack_no"),
        lot=data.get("lot"),
    )
    return jsonify(res)


@app.route("/api/invoices", methods=["GET"])
def get_invoices():
    invoices = mgr.get_invoices_list()
    return jsonify({"status": "SUCCESS", "invoices": invoices})



@app.route("/api/export_excel", methods=["POST"])
def export_excel():
    data = request.get_json() or {}
    invoice_no = data.get("invoice_no", "11726091882")

    try:
        out_path = mgr.export_delivery_trip_sheet(invoice_no=invoice_no)
        filename = os.path.basename(out_path)
        return jsonify({
            "status": "SUCCESS",
            "filepath": out_path,
            "filename": filename,
        })
    except Exception as e:
        return jsonify({"status": "ERROR", "message": str(e)}), 500


@app.route("/api/download_export", methods=["GET"])
def download_export():
    filename = request.args.get("filename")
    if not filename:
        return "Filename required", 400
    path = os.path.join(BASE_DIR, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return "File not found", 404


@app.route("/api/seed_historical_scans", methods=["POST"])
def seed_historical_scans():
    """Seeds the 1,502 scans from 5. Delivery trip sheet for lamp.xlsx for full verification testing."""
    import openpyxl
    ts_file = mgr.get_file_path("5. Delivery trip sheet for lamp.xlsx")
    if not ts_file:
        return jsonify({"status": "ERROR", "message": "File not found"}), 404

    wb = openpyxl.load_workbook(ts_file, data_only=True)
    ws = wb[" BP Delivery 280926"]

    import sqlite3
    conn = sqlite3.connect(mgr.db_path)
    cur = conn.cursor()
    cur.execute("DELETE FROM pack_scans WHERE invoice_no = '11726091882'")

    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    records = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        p_scan = str(r[2] or "").strip()
        if not p_scan or p_scan == "Part Scan":
            continue
        rack = str(r[1] or "")
        p_inv = str(r[3] or "").strip()
        p_name = str(r[4] or "").strip()
        qty = int(r[5] or 1)
        po = int(r[6] or 0)
        line = int(r[7] or 0)
        inv = str(r[8] or "11726091882")

        # Correct typographical mismatch: Line 30 for UL4H510K0B in Trip Sheet
        if po == 229530 and line == 3 and p_inv == "UL4H510K0B":
            line = 30

        # Get BOM box
        bom_spec = mgr.get_bom_spec(p_inv)
        carton = bom_spec.get("carton_box")
        box_no = carton["pkg_no"] if carton else "-"

        records.append((
            now_iso, inv, po, line, p_scan, p_inv, p_name,
            box_no, box_no, 1, rack, qty, "OK"
        ))

    cur.executemany("""
        INSERT INTO pack_scans
        (timestamp, invoice_no, po_no, line, part_scan, part_no, part_name, box_scan, box_expected, is_box_valid, rack_no, qty, remark)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, records)
    conn.commit()
    conn.close()

    trigger_background_plan_sync()

    return jsonify({"status": "SUCCESS", "count": len(records), "message": f"ซิงค์ข้อมูลสำเร็จ {len(records)} รายการ"})


@app.route("/api/clear_pack_scans", methods=["POST"])
def clear_pack_scans():
    """Clears pack scans for resetting tests."""
    data = request.get_json(silent=True) or {}
    inv = data.get("invoice_no", "11726091882")
    import sqlite3
    conn = sqlite3.connect(mgr.db_path)
    cur = conn.cursor()
    if inv == "ALL":
        cur.execute("DELETE FROM pack_scans")
    else:
        cur.execute("DELETE FROM pack_scans WHERE invoice_no = ?", (inv,))
    deleted = cur.rowcount
    conn.commit()
    conn.close()

    trigger_background_plan_sync()

    return jsonify({"status": "SUCCESS", "deleted_count": deleted, "message": f"ล้างประวัติการสแกนสำเร็จ ({deleted} รายการ)"})



@app.route("/manifest.json")
def serve_manifest():
    return send_file(os.path.join(BASE_DIR, "static", "manifest.json"))


@app.route("/sw.js")
def serve_sw():
    resp = make_response(send_file(os.path.join(BASE_DIR, "static", "sw.js"), mimetype="application/javascript"))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


DEFAULT_PORT = int(os.environ.get("PORT", 5050))


@app.route("/api/mobile_info")
def get_mobile_info():
    import socket, io, base64, qrcode
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()

    port_str = request.host.split(":")[1] if ":" in request.host else str(DEFAULT_PORT)
    lan_url = f"http://{ip}:{port_str}"

    def make_qr(target_url):
        if not target_url:
            return ""
        qr = qrcode.QRCode(box_size=8, border=2)
        qr.add_data(target_url)
        qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white')
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode('utf-8')

    lan_qr = make_qr(lan_url)
    https_url = CLOUDFLARE_URL
    https_qr = make_qr(https_url) if https_url else ""

    # Primary URL for mobile: Prefer HTTPS so camera stream works without browser security restrictions
    primary_url = https_url if https_url else lan_url
    primary_qr = https_qr if https_qr else lan_qr

    # Ensure static/mobile_qr.png is updated with primary URL
    try:
        qr = qrcode.QRCode(box_size=8, border=2)
        qr.add_data(primary_url)
        qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white')
        img.save(os.path.join(BASE_DIR, "static", "mobile_qr.png"))
    except Exception:
        pass

    return jsonify({
        "ip": ip,
        "port": port_str,
        "url": primary_url,
        "qr_base64": primary_qr,
        "lan_url": lan_url,
        "qr_lan_base64": lan_qr,
        "https_url": https_url,
        "qr_https_base64": https_qr,
        "has_https": bool(https_url),
        "qr_url": "/static/mobile_qr.png"
    })


# ─────────────────────────────────────────────
# Integration with Plan Production (port 5000 - Process: Packing)
# ─────────────────────────────────────────────
@app.route("/api/plan_production/status", methods=["GET"])
def plan_production_status():
    """Checks connection to Plan Production on port 5000 and returns status."""
    import urllib.request
    import json
    plan_url = "http://localhost:5000"
    is_online = False
    record_count = 0
    try:
        req = urllib.request.urlopen(f"{plan_url}/api/production-records?process=Packing", timeout=2)
        if req.status == 200:
            is_online = True
            data = json.loads(req.read().decode("utf-8"))
            record_count = len(data.get("records", []))
    except Exception:
        is_online = False

    target_date = request.args.get("date") or datetime.date.today().strftime("%Y-%m-%d")
    target_po = request.args.get("po_no")
    preview = mgr.get_plan_production_preview(target_date=target_date, target_po=target_po)

    return jsonify({
        "status": "ONLINE" if is_online else "OFFLINE",
        "endpoint": plan_url,
        "plan_url": plan_url,
        "is_online": is_online,
        "existing_packing_records": record_count,
        "preview_items": preview,
        "target_date": target_date
    })


@app.route("/api/sync_to_plan_production", methods=["GET", "POST"])
def sync_to_plan_production_route():
    """Preview (GET) or send (POST) daily packed quantities for target_date to Plan Production (port 5000)."""
    import json
    if request.method == "GET" or request.args.get("preview"):
        target_date = request.args.get("date") or datetime.date.today().strftime("%Y-%m-%d")
        preview_items = mgr.get_plan_production_preview(target_date=target_date)

        # Check existing records in Plan Production to identify UPDATE vs CREATE
        existing_records = []
        try:
            import urllib.request
            req = urllib.request.urlopen("http://localhost:5000/api/production-records?process=Packing&include_unconfirmed=true", timeout=2)
            if req.status == 200:
                rec_data = json.loads(req.read().decode("utf-8"))
                for r in rec_data.get("records", []):
                    r_date = r.get("date", "")
                    if target_date and r_date != target_date:
                        continue
                    existing_records.append(r)
        except Exception:
            pass

        total_qty = 0
        import re
        for item in preview_items:
            pkg = (item.get("package_no") or item.get("box") or "").strip().upper()
            total_qty += item.get("total_qty", 0)
            item["lamp_qty"] = item.get("total_qty", 0)

            # Find matching record in Plan MST (exact, suffix alias, or assembly_items)
            matches = [r for r in existing_records if is_same_package(r.get("package_no"), pkg, r)]
            if matches:
                manuals = [r for r in matches if not ("จากระบบ LAMP Packing" in (r.get("notes") or "") or "LAMP FG Pack" in (r.get("notes") or ""))]
                m_rec = manuals[0] if manuals else matches[0]
                rec_id = m_rec.get("id")
                clean_notes = (m_rec.get("notes") or "").strip()
                tag = re.search(r'\[LAMP Packing:\s*([\d\.]+)\s*ชิ้น\]', clean_notes)
                is_pure_lamp = "จากระบบ LAMP Packing" in clean_notes or "LAMP FG Pack" in clean_notes

                if is_pure_lamp and not tag:
                    base_qty = 0.0
                elif tag:
                    base_qty = max(0.0, float(m_rec.get("fg_qty") or 0) - float(tag.group(1)))
                else:
                    base_qty = float(m_rec.get("fg_qty") or 0)

                combined_qty = base_qty + item.get("total_qty", 0)
                item["fg_qty"] = combined_qty
                item["base_qty"] = base_qty
                item["action"] = f"UPDATE (รวมยอด: {int(combined_qty)} ชิ้น)" if base_qty > 0 else "UPDATE"
                item["record_id"] = rec_id
                item["existing_package_no"] = m_rec.get("package_no")
            else:
                item["fg_qty"] = item.get("total_qty", 0)
                item["base_qty"] = 0.0
                item["action"] = "CREATE"
                item["record_id"] = None

        return jsonify({
            "status": "SUCCESS",
            "preview_items": preview_items,
            "total_qty": total_qty,
            "target_date": target_date
        })

    # POST method: Execute sync
    data = request.get_json() or {}
    target_date = data.get("date") or datetime.date.today().strftime("%Y-%m-%d")
    plan_url = data.get("plan_url") or "http://localhost:5000"
    res = mgr.sync_to_plan_production(target_date=target_date, plan_url=plan_url)
    created = len([x for x in res.get("results", []) if "CREATE" in str(x.get("action"))])
    updated = len([x for x in res.get("results", []) if "UPDATE" in str(x.get("action"))])
    total_qty = sum(x.get("qty", 0) for x in res.get("results", []))
    is_file = any("File Fallback" in str(x.get("action")) for x in res.get("results", []))

    return jsonify({
        "status": "SUCCESS" if res.get("success") else "ERROR",
        "date": target_date,
        "total_qty": total_qty,
        "created_count": created,
        "updated_count": updated,
        "sync_method": "Direct File (production_records.json)" if is_file else "HTTP REST API (Port 5000)",
        "results": res.get("results", [])
    })


@app.route("/api/supabase/status", methods=["GET"])
def get_supabase_status_route():
    """Returns the current Supabase PostgreSQL connection and sync status."""
    try:
        status = supabase_client.get_sync_status()
        return jsonify(status)
    except Exception as e:
        return jsonify({"connected": False, "error": str(e)}), 500


@app.route("/api/supabase/sync", methods=["POST"])
def post_supabase_sync_route():
    """Triggers bidirectional or directional synchronization with Supabase."""
    try:
        req_data = request.get_json(silent=True) or {}
        direction = req_data.get("direction", "bidirectional")
        res = supabase_client.sync_all(direction=direction)
        return jsonify(res)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


if __name__ == "__main__":
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()

    # Detect if port is free or fallback
    def is_port_free(p):
        test_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            test_sock.bind(('0.0.0.0', p))
            test_sock.close()
            return True
        except Exception:
            return False

    # Detect active port (Cloud PORT env or local check)
    env_port = os.environ.get("PORT")
    if env_port:
        active_port = int(env_port)
    else:
        active_port = DEFAULT_PORT
        if not is_port_free(active_port):
            for candidate in [5051, 8080, 8000, 5001]:
                if is_port_free(candidate):
                    active_port = candidate
                    break

    # Start Cloudflare Tunnel for secure camera access (HTTPS)
    start_cloudflare_tunnel(active_port)

    # Start background periodic auto-sync from WMS (every 10 minutes)
    start_wms_periodic_sync()

    print("=" * 72)
    print("   LAMP Verification & Packing Mobile & Web System")
    print(f"   💻 บนคอมพิวเตอร์เครื่องนี้: http://127.0.0.1:{active_port}")
    print(f"   📱 เปิดบนมือถือผ่าน Wi-Fi:   http://{ip}:{active_port}")
    print("   🔒 กำลังเตรียมลิงก์ HTTPS สำหรับเปิดกล้องมือถือสด...")
    print("=" * 72)

    try:
        from waitress import serve
        print(f"   🚀 ทำงานด้วย Waitress Production WSGI Server (รองรับ 24 ชม. และหลายเครื่องพร้อมกัน)")
        serve(app, host="0.0.0.0", port=active_port, threads=8)
    except Exception as e:
        print(f"   ℹ️ รันด้วย Flask Development Server: {e}")
        app.run(host="0.0.0.0", port=active_port, debug=False)



