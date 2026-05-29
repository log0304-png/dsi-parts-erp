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
    create_part_number, part_number_exists
)
from google import genai

BASE_DIR      = Path(__file__).parent
CONFIG        = json.loads((BASE_DIR / "parts_config.json").read_text(encoding="utf-8"))
# 環境變數優先（Render 雲端部署用）
_gemini_key = os.environ.get("GEMINI_API_KEY") or CONFIG.get("gemini_api_key", "")
UPLOAD_DIR    = BASE_DIR / "uploads"
PROCESSED_DIR = BASE_DIR / "processed"
UPLOAD_DIR.mkdir(exist_ok=True)
PROCESSED_DIR.mkdir(exist_ok=True)

app = Flask(__name__)

# 每個任務的進度 queue
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

                # 將 PDF 改名為料號並移至 processed 資料夾
                renamed = PROCESSED_DIR / f"{part_number}.pdf"
                pdf_path.rename(renamed)
                pdf_path = renamed  # 更新路徑，finally 不重複刪除

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

        log(f"全部完成！", "success")

    except Exception as e:
        log(f"初始化失敗：{e}", "error")
        logging.error(f"Task {task_id} 失敗：{e}", exc_info=True)
    finally:
        q.put({"type": "done"})
        # 清除暫存檔
        for p in pdf_paths:
            try:
                p.unlink()
            except Exception:
                pass


@app.route("/")
def index():
    sheet_url = (
        "https://docs.google.com/spreadsheets/d/"
        + CONFIG.get("sheet_id", "")
    )
    return render_template("index.html", sheet_url=sheet_url)


@app.route("/erp")
def erp():
    return render_template("erp.html")


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


if __name__ == "__main__":
    print("DSI 料號生成工具 網頁版啟動")
    print("請用瀏覽器開啟 http://localhost:5002")
    app.run(host="0.0.0.0", port=5002, debug=False, threaded=True)
