# -*- coding: utf-8 -*-
"""
DSI 料號生成工具 - 網頁版
"""
import os
import sys
import json
import queue
import threading
import uuid
import logging
from pathlib import Path
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

logging.basicConfig(
    filename=str(Path(__file__).parent / "web_app.log"),
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8",
)

from flask import Flask, render_template, request, jsonify, Response, send_file
from pdf_analyzer import (
    connect_sheets, load_rules, pdf_page_to_bytes, analyze_drawing,
    make_fingerprint, find_existing_by_fingerprint, get_category,
    create_part_number, part_number_exists, get_spreadsheet
)
from google import genai

BASE_DIR      = Path(__file__).parent
CONFIG        = json.loads((BASE_DIR / "parts_config.json").read_text(encoding="utf-8"))
_gemini_key   = os.environ.get("GEMINI_API_KEY") or CONFIG.get("gemini_api_key", "")
UPLOAD_DIR    = BASE_DIR / "uploads"
PROCESSED_DIR = BASE_DIR / "processed"
UPLOAD_DIR.mkdir(exist_ok=True)
PROCESSED_DIR.mkdir(exist_ok=True)

app = Flask(__name__)

_queues: dict[str, queue.Queue] = {}


def process_task(task_id: str, pdf_paths: list[Path]):
    q = _queues[task_id]

    def log(msg: str, status: str = "info"):
        q.put({"type": "log", "msg": msg, "status": status})

    def result(data: dict):
        q.put({"type": "result", "data": data})

    try:
        log("連線 Google Sheets...")
        logging.info(f"Task {task_id} 開始，{len(pdf_paths)} 個 PDF")
        ws_erp, ws_fp, erp_rows, fp_rows, rule_rows = connect_sheets()
        rules  = load_rules(rule_rows)
        client = genai.Client(api_key=_gemini_key)
        log(f"載入分類規則 {len(rules)} 條，開始處理 {len(pdf_paths)} 個 PDF")

        for pdf_path in pdf_paths:
            log(f"處理：{pdf_path.name}")
            try:
                image_bytes = pdf_page_to_bytes(pdf_path)
                data        = analyze_drawing(client, image_bytes)

                project  = data.get("project", "")
                drawing  = data.get("drawing", "")
                material = data.get("material", "")
                revision = data.get("revision", "")

                log(f"  OCR → project={project} | drawing={drawing} | material={material} | rev={revision}")

                fingerprint = make_fingerprint(project, drawing, material, revision)
                existing    = find_existing_by_fingerprint(fingerprint, fp_rows)

                if existing:
                    log(f"  重複：圖面已有料號 {existing}", "warning")
                    result({"file": pdf_path.name, "part_number": existing,
                            "status": "duplicate", "category": "-",
                            "drawing": drawing, "material": material, "revision": revision})
                    continue

                category    = get_category(drawing, material, rules)
                part_number = create_part_number(category, project, drawing, revision, erp_rows)

                if part_number_exists(part_number, erp_rows):
                    log(f"  衝突：料號已存在 {part_number}", "warning")
                    result({"file": pdf_path.name, "part_number": part_number,
                            "status": "conflict", "category": category,
                            "drawing": drawing, "material": material, "revision": revision})
                    continue

                now      = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                file_str = str(pdf_path.resolve())

                erp_row = [now, part_number, category, project, drawing,
                           material, revision, file_str, data.get("designer", "")]
                ws_erp.append_row(erp_row)
                erp_rows.append(erp_row)

                fp_row = [part_number, drawing, material, file_str,
                          fingerprint, revision, now, ""]
                ws_fp.append_row(fp_row)
                fp_rows.append(fp_row)

                renamed = PROCESSED_DIR / f"{part_number}.pdf"
                pdf_path.rename(renamed)
                pdf_path = renamed

                log(f"  建立料號：{part_number}  分類：{category}", "success")
                result({"file": pdf_path.name, "part_number": part_number,
                        "status": "created", "category": category,
                        "drawing": drawing, "material": material, "revision": revision,
                        "download": part_number})

            except Exception as e:
                log(f"  錯誤：{e}", "error")
                result({"file": pdf_path.name, "part_number": "-",
                        "status": "error", "category": "-",
                        "drawing": "-", "material": "-", "revision": "-"})

        log("全部完成！", "success")

    except Exception as e:
        log(f"初始化失敗：{e}", "error")
        logging.error(f"Task {task_id} 失敗：{e}", exc_info=True)
    finally:
        q.put({"type": "done"})
        for p in pdf_paths:
            try:
                p.unlink()
            except Exception:
                pass


# ── 主頁 & 上傳 ────────────────────────────────────────

@app.route("/")
def index():
    sheet_url = "https://docs.google.com/spreadsheets/d/" + CONFIG.get("sheet_id", "")
    return render_template("index.html", sheet_url=sheet_url)


@app.route("/erp")
def erp():
    return render_template("erp.html")


@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("pdfs")
    if not files:
        return jsonify({"error": "未選擇檔案"}), 400

    task_id   = str(uuid.uuid4())
    pdf_paths = []
    for f in files:
        if f.filename.lower().endswith(".pdf"):
            dest = UPLOAD_DIR / f"{task_id}_{f.filename}"
            f.save(str(dest))
            pdf_paths.append(dest)

    if not pdf_paths:
        return jsonify({"error": "沒有有效的 PDF 檔案"}), 400

    _queues[task_id] = queue.Queue()
    threading.Thread(target=process_task, args=(task_id, pdf_paths), daemon=True).start()
    return jsonify({"task_id": task_id})


@app.route("/stream/<task_id>")
def stream(task_id: str):
    def generate():
        q = _queues.get(task_id)
        if not q:
            yield "data: {\"type\":\"done\"}\n\n"
            return
        while True:
            item = q.get()
            yield f"data: {json.dumps(item, ensure_ascii=False)}\n\n"
            if item.get("type") == "done":
                _queues.pop(task_id, None)
                break

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/download/<part_number>")
def download(part_number: str):
    path = PROCESSED_DIR / f"{part_number}.pdf"
    if not path.exists():
        return "檔案不存在", 404
    return send_file(str(path), as_attachment=True, download_name=f"{part_number}.pdf")


# ── 料號 & 圖面 API ────────────────────────────────────

@app.route("/api/parts")
def api_parts():
    try:
        ws_erp, _ws_fp, erp_rows, _fp_rows, _rule_rows = connect_sheets()
        keys = ["created", "part_number", "category", "project",
                "drawing", "material", "revision", "designer", "drawing_no"]
        data = []
        for row in erp_rows:
            padded = row + [""] * max(0, len(keys) - len(row))
            data.append(dict(zip(keys, padded[:len(keys)])))
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/fps")
def api_fps():
    try:
        _ws_erp, _ws_fp, _erp_rows, fp_rows, _rule_rows = connect_sheets()
        keys = ["part_number", "drawing", "material", "fingerprint",
                "revision", "created", "drawing_no"]
        data = []
        for row in fp_rows:
            padded = row + [""] * max(0, len(keys) - len(row))
            data.append(dict(zip(keys, padded[:len(keys)])))
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ── 庫存 API ───────────────────────────────────────────

@app.route("/api/inventory")
def api_inventory():
    try:
        sh   = get_spreadsheet()
        rows = sh.worksheet("庫存管理").get_all_values()
        keys = ["part_number", "name", "qty", "location", "updated"]
        data = [dict(zip(keys, (r + [""] * len(keys))[:len(keys)])) for r in rows[1:] if any(r)]
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/inventory/in", methods=["POST"])
def api_stock_in():
    try:
        b    = request.get_json(force=True)
        pn   = b.get("part_number", "").strip()
        name = b.get("name", "").strip()
        qty  = int(b.get("qty", 0))
        loc  = b.get("location", "").strip()
        if not pn or qty <= 0:
            return jsonify({"ok": False, "error": "料號或數量不正確"}), 400
        now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sh   = get_spreadsheet()
        ws   = sh.worksheet("庫存管理")
        rows = ws.get_all_values()
        for i, row in enumerate(rows[1:], start=2):
            if row and str(row[0]).strip().upper() == pn.upper():
                old = int(row[2]) if len(row) > 2 and row[2] else 0
                ws.update_cell(i, 3, old + qty)
                if loc: ws.update_cell(i, 4, loc)
                ws.update_cell(i, 5, now)
                return jsonify({"ok": True, "new_qty": old + qty})
        ws.append_row([pn, name, qty, loc, now])
        return jsonify({"ok": True, "new_qty": qty})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/inventory/out", methods=["POST"])
def api_stock_out():
    try:
        b   = request.get_json(force=True)
        pn  = b.get("part_number", "").strip()
        qty = int(b.get("qty", 0))
        if not pn or qty <= 0:
            return jsonify({"ok": False, "error": "料號或數量不正確"}), 400
        now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sh   = get_spreadsheet()
        ws   = sh.worksheet("庫存管理")
        rows = ws.get_all_values()
        for i, row in enumerate(rows[1:], start=2):
            if row and str(row[0]).strip().upper() == pn.upper():
                old = int(row[2]) if len(row) > 2 and row[2] else 0
                if old < qty:
                    return jsonify({"ok": False, "error": f"庫存不足（現有 {old}）"}), 400
                ws.update_cell(i, 3, old - qty)
                ws.update_cell(i, 5, now)
                return jsonify({"ok": True, "new_qty": old - qty})
        return jsonify({"ok": False, "error": "找不到該料號庫存"}), 404
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ── 採購單 API ─────────────────────────────────────────

@app.route("/api/purchase")
def api_purchase():
    try:
        sh   = get_spreadsheet()
        rows = sh.worksheet("採購單").get_all_values()
        keys = ["po_no", "part_number", "vendor", "created", "qty", "due_date", "status", "note"]
        data = [dict(zip(keys, (r + [""] * len(keys))[:len(keys)])) for r in rows[1:] if any(r)]
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/purchase", methods=["POST"])
def api_purchase_create():
    try:
        b        = request.get_json(force=True)
        pn       = b.get("part_number", "").strip()
        vendor   = b.get("vendor", "").strip()
        qty      = b.get("qty", "")
        due_date = b.get("due_date", "").strip()
        note     = b.get("note", "").strip()
        if not pn or not vendor:
            return jsonify({"ok": False, "error": "料號與廠商為必填"}), 400
        now    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        today  = datetime.now().strftime("%Y%m%d")
        sh     = get_spreadsheet()
        ws     = sh.worksheet("採購單")
        rows   = ws.get_all_values()[1:]
        prefix = f"PO-{today}-"
        max_seq = max(
            (int(str(r[0])[-3:]) for r in rows if r and str(r[0]).startswith(prefix) and str(r[0])[-3:].isdigit()),
            default=0
        )
        po_no = f"{prefix}{str(max_seq + 1).zfill(3)}"
        ws.append_row([po_no, pn, vendor, now[:10], qty, due_date, "進行中", note])
        return jsonify({"ok": True, "po_no": po_no})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/purchase/<po_no>/complete", methods=["POST"])
def api_purchase_complete(po_no):
    try:
        sh   = get_spreadsheet()
        ws   = sh.worksheet("採購單")
        rows = ws.get_all_values()
        for i, row in enumerate(rows[1:], start=2):
            if row and str(row[0]).strip().upper() == po_no.upper():
                ws.update_cell(i, 7, "完成")
                return jsonify({"ok": True})
        return jsonify({"ok": False, "error": "找不到採購單"}), 404
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ── 承認狀況 API ───────────────────────────────────────

@app.route("/api/approval")
def api_approval():
    try:
        sh   = get_spreadsheet()
        rows = sh.worksheet("承認狀況").get_all_values()
        keys = ["part_number", "vendor", "status", "date", "note", "created"]
        data = [dict(zip(keys, (r + [""] * len(keys))[:len(keys)])) for r in rows[1:] if any(r)]
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/approval", methods=["POST"])
def api_approval_add():
    try:
        b      = request.get_json(force=True)
        pn     = b.get("part_number", "").strip()
        vendor = b.get("vendor", "").strip()
        status = b.get("status", "").strip()
        date   = b.get("date", "").strip()
        note   = b.get("note", "").strip()
        if not pn or not vendor or not status:
            return jsonify({"ok": False, "error": "料號、廠商、狀態為必填"}), 400
        now = datetime.now().strftime("%Y-%m-%d")
        if not date:
            date = now
        sh  = get_spreadsheet()
        sh.worksheet("承認狀況").append_row([pn, vendor, status, date, note, now])
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    print("DSI 料號生成工具 網頁版啟動")
    print("請用瀏覽器開啟 http://localhost:5002")
    app.run(host="0.0.0.0", port=5002, debug=False, threaded=True)
