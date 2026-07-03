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
import re
import traceback
import requests
from datetime import datetime, timezone, timedelta

if sys.stdout:
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

# ── LINE ME助理 BOT 設定 ──────────────────────────
_TW        = timezone(timedelta(hours=8))
ERP_TOKEN  = os.environ.get("ERP_TOKEN", "") or CONFIG.get("line_erp_token", "")
_erp_states: dict = {}  # { user_id: {"action": "入庫"/"下單"} }

# ── 請款 Sheet 設定 ───────────────────────────────────
EXPENSE_SHEET_ID = "1yf62_kTCfEPt0hYg5IoGsW7_EDddGG2Ft5yiisqLPhM"
EXPENSE_HEADERS  = ["摘要", "項目", "發票號碼", "請款人", "日期",
                    "研發相關", "加油費", "交通費", "房租", "行銷",
                    "郵寄費", "旅費", "餐費", "工程", "辦公室補給", "備註", "發票圖片"]
EXPENSE_COLS     = ["研發相關", "加油費", "交通費", "房租", "行銷",
                    "郵寄費", "旅費", "餐費", "工程", "辦公室補給"]
_expense_sh      = None


def _get_expense_sheet():
    global _expense_sh
    if _expense_sh is None:
        from pdf_analyzer import _get_creds
        import gspread as _gs
        _expense_sh = _gs.authorize(_get_creds()).open_by_key(EXPENSE_SHEET_ID)
        existing = [ws.title for ws in _expense_sh.worksheets()]
        if "請款" not in existing:
            ws = _expense_sh.add_worksheet(title="請款", rows=1000, cols=17)
            ws.update("A1:Q1", [EXPENSE_HEADERS])
    return _expense_sh


def _upload_to_drive(image_bytes: bytes, filename: str) -> str:
    """上傳圖片到 Google Drive，設為公開連結，回傳 =IMAGE() 公式字串"""
    import io
    from pdf_analyzer import _get_creds
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload

    creds   = _get_creds()
    service = build("drive", "v3", credentials=creds)

    file_metadata = {"name": filename}
    media    = MediaIoBaseUpload(io.BytesIO(image_bytes), mimetype="image/jpeg")
    uploaded = service.files().create(
        body=file_metadata, media_body=media, fields="id"
    ).execute()
    file_id = uploaded.get("id")

    service.permissions().create(
        fileId=file_id,
        body={"type": "anyone", "role": "reader"},
    ).execute()

    url = f"https://drive.google.com/uc?id={file_id}"
    return f'=IMAGE("{url}")'


def _get_line_display_name(user_id: str) -> str:
    try:
        resp = requests.get(
            f"https://api.line.me/v2/bot/profile/{user_id}",
            headers={"Authorization": f"Bearer {ERP_TOKEN}"},
            timeout=5,
        )
        if resp.status_code == 200:
            return resp.json().get("displayName", user_id)
    except Exception:
        pass
    return user_id


def _analyze_invoice(image_bytes: bytes) -> dict:
    client = genai.Client(api_key=_gemini_key)
    from google.genai import types as _gt
    col_list = "、".join(EXPENSE_COLS)
    prompt = (
        "你是台灣公司請款AI。分析這張發票或收據圖片，只回傳JSON，不要解釋。\n"
        f"expense_col 必須從以下選一個：{col_list}\n"
        "{\n"
        '  "date": "YYYY-MM-DD（發票日期，若無則今天）",\n'
        '  "invoice_number": "發票號碼（若無則空字串）",\n'
        '  "amount": 金額數字（整數，台幣）,\n'
        '  "items": "品項描述（簡短）",\n'
        '  "expense_col": "費用欄位",\n'
        '  "summary": "摘要（一句話）",\n'
        '  "notes": "備註（若有特殊說明）"\n'
        "}"
    )
    image_part = _gt.Part.from_bytes(data=image_bytes, mime_type="image/jpeg")
    resp = client.models.generate_content(
        model="gemini-3.1-flash-lite",
        contents=[prompt, image_part],
        config=_gt.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.1,
        ),
    )
    return json.loads(resp.text)


def handle_invoice_image(user_id: str, message_id: str, reply_token: str):
    try:
        img_resp = requests.get(
            f"https://api-data.line.me/v2/bot/message/{message_id}/content",
            headers={"Authorization": f"Bearer {ERP_TOKEN}"},
            timeout=15,
        )
        if img_resp.status_code != 200:
            _line_reply(reply_token, "⚠️ 無法取得圖片，請重試。")
            return

        image_bytes = img_resp.content
        data        = _analyze_invoice(image_bytes)
        requester   = _get_line_display_name(user_id)
        expense_col = data.get("expense_col", "")
        if expense_col not in EXPENSE_COLS:
            expense_col = EXPENSE_COLS[0]

        filename    = f"invoice_{datetime.now(_TW).strftime('%Y%m%d_%H%M%S')}_{user_id[:6]}.jpg"
        image_formula = _upload_to_drive(image_bytes, filename)

        row = [""] * 17
        row[0]  = data.get("summary", "")
        row[1]  = data.get("items", "")
        row[2]  = data.get("invoice_number", "")
        row[3]  = requester
        row[4]  = data.get("date", datetime.now(_TW).strftime("%Y-%m-%d"))
        col_idx = EXPENSE_HEADERS.index(expense_col)
        row[col_idx] = data.get("amount", "")
        row[16] = image_formula  # 發票圖片 =IMAGE(url)

        sh = _get_expense_sheet()
        sh.worksheet("請款").append_row(row, value_input_option="USER_ENTERED")

        _line_reply(reply_token, (
            f"✅ 發票已記錄\n"
            f"━━━━━━━━━━━━━━\n"
            f"請款人：{requester}\n"
            f"日　期：{row[4]}\n"
            f"發票號：{row[2] or '─'}\n"
            f"品　項：{row[1]}\n"
            f"金　額：NT$ {data.get('amount', '─')}\n"
            f"類　別：{expense_col}\n"
            f"圖　片：已同步至 Sheet"
        ))

    except Exception as e:
        print(traceback.format_exc(), flush=True)
        _line_reply(reply_token, "⚠️ 發票辨識失敗，請確認圖片清晰後重試。")


def _line_reply(reply_token, text):
    requests.post(
        "https://api.line.me/v2/bot/message/reply",
        headers={"Authorization": f"Bearer {ERP_TOKEN}", "Content-Type": "application/json"},
        json={"replyToken": reply_token, "messages": [{"type": "text", "text": text}]},
        timeout=10,
    )


def handle_erp(text, reply_token, user_id=""):
    try:
        sh  = get_spreadsheet()
        raw = text.strip()
        now = datetime.now(_TW).strftime("%Y-%m-%d %H:%M:%S")

        def reply(msg):
            _line_reply(reply_token, msg)

        def find_in_erp(pn):
            rows = sh.worksheet("ERP資料庫").get_all_values()[1:]
            pn_upper = pn.strip().upper()
            for row in rows:
                if len(row) > 1 and str(row[1]).strip().upper() == pn_upper:
                    return row
            return None

        def find_rows(ws_name, col_idx, value):
            rows = sh.worksheet(ws_name).get_all_values()[1:]
            val_upper = value.strip().upper()
            return [r for r in rows if len(r) > col_idx and str(r[col_idx]).strip().upper() == val_upper]

        # ── 使用說明 ──────────────────────────────────
        if raw in ["使用說明", "說明", "help", "Help", "?", "？"]:
            reply("\n".join([
                "📖 料號ERP BOT 使用說明",
                "─" * 22,
                "【查詢】",
                "查 料號",
                "查庫存 料號",
                "查承認 料號",
                "查下單 料號",
                "查交期 料號",
                "",
                "【庫存操作】",
                "入庫 料號 數量",
                "出庫 料號 數量",
                "我要增加庫存 → 對話式入庫",
                "",
                "【採購】",
                "下單 料號 廠商 數量 交期",
                "例：下單 PN001 台灣製造 100 2026-07-01",
                "到貨 PO單號或料號 數量",
                "我要下單 → 對話式下單",
                "─" * 22,
                "傳「使用說明」可再次查看",
            ]))
            return

        # ── 對話式流程 ────────────────────────────────
        if raw == "我要增加庫存":
            _erp_states[user_id] = {"action": "入庫"}
            reply("請填寫入庫資料：\n\n料號：\n數量：\n儲位：")
            return

        if user_id in _erp_states and _erp_states[user_id]["action"] == "入庫":
            data = {}
            for line in raw.splitlines():
                for key in ["料號", "數量", "儲位"]:
                    if line.startswith(key):
                        parts_kv = re.split(r"[：:]", line, 1)
                        if len(parts_kv) > 1 and parts_kv[1].strip():
                            data[key] = parts_kv[1].strip()
            if all(k in data for k in ["料號", "數量"]):
                del _erp_states[user_id]
                pn  = data["料號"]
                qty = int(data["數量"])
                loc = data.get("儲位", "")
                ws       = sh.worksheet("庫存管理")
                all_rows = ws.get_all_values()
                for i, row in enumerate(all_rows[1:], start=2):
                    if row and str(row[0]).strip().upper() == pn.upper():
                        old = int(row[2]) if row[2] else 0
                        ws.update_cell(i, 3, old + qty)
                        if loc:
                            ws.update_cell(i, 4, loc)
                        ws.update_cell(i, 5, now)
                        reply(f"入庫完成\n料號：{pn}\n入庫：{qty}\n庫存：{old} → {old + qty}\n儲位：{loc or row[3]}")
                        return
                erp_row = find_in_erp(pn)
                ws.append_row([pn, erp_row[4] if erp_row else "", qty, loc, now])
                reply(f"入庫完成（新建庫存）\n料號：{pn}\n庫存：{qty}\n儲位：{loc}")
            else:
                reply("資料不完整，請重新填寫：\n\n料號：\n數量：\n儲位：")
            return

        if raw == "我要下單":
            _erp_states[user_id] = {"action": "下單"}
            reply("請填寫下單資料：\n\n料號：\n廠商：\n數量：\n交期：YYYY-MM-DD")
            return

        if user_id in _erp_states and _erp_states[user_id]["action"] == "下單":
            data = {}
            for line in raw.splitlines():
                for key in ["料號", "廠商", "數量", "交期"]:
                    if line.startswith(key):
                        parts_kv = re.split(r"[：:]", line, 1)
                        if len(parts_kv) > 1 and parts_kv[1].strip():
                            data[key] = parts_kv[1].strip()
            if all(k in data for k in ["料號", "廠商", "數量", "交期"]):
                del _erp_states[user_id]
                pn, vendor, qty, delivery = data["料號"], data["廠商"], int(data["數量"]), data["交期"]
                today   = datetime.now(_TW).strftime("%Y%m%d")
                po_rows = sh.worksheet("採購單").get_all_values()[1:]
                prefix  = f"PO-{today}-"
                max_seq = max((int(str(r[0])[-3:]) for r in po_rows if r and str(r[0]).startswith(prefix)), default=0)
                po_no   = f"{prefix}{str(max_seq + 1).zfill(3)}"
                sh.worksheet("採購單").append_row([po_no, pn, vendor, now[:10], qty, delivery, "待交貨", ""])
                reply(f"採購單建立成功\n單號：{po_no}\n料號：{pn}\n廠商：{vendor}\n數量：{qty}\n交期：{delivery}")
            else:
                reply("資料不完整，請重新填寫：\n\n料號：\n廠商：\n數量：\n交期：YYYY-MM-DD")
            return

        m = re.match(r"^查\s+(\S+)$", raw)
        if m:
            pn  = m.group(1)
            row = find_in_erp(pn)
            if not row:
                reply(f"找不到料號：{pn}")
                return
            reply(f"【料號資料】\n料號：{row[1]}\n分類：{row[2]}\n專案：{row[3]}\n圖面：{row[4]}\n材質：{row[5]}\n版次：{row[6]}")
            return

        m = re.match(r"^查庫存\s+(\S+)$", raw)
        if m:
            pn   = m.group(1)
            rows = find_rows("庫存管理", 0, pn)
            if not rows:
                reply(f"{pn}\n尚無庫存紀錄")
                return
            r = rows[-1]
            reply(f"【庫存查詢】\n料號：{r[0]}\n品名：{r[1]}\n庫存量：{r[2]}\n儲位：{r[3]}\n更新：{r[4]}")
            return

        m = re.match(r"^查承認\s+(\S+)$", raw)
        if m:
            pn   = m.group(1)
            rows = find_rows("承認狀況", 0, pn)
            if not rows:
                reply(f"{pn}\n尚無承認紀錄")
                return
            lines = [f"  {r[1]} | {r[2]} | {r[3]}" for r in rows]
            reply(f"【承認狀況】\n料號：{pn}\n\n廠商 | 狀態 | 日期\n" + "\n".join(lines))
            return

        m = re.match(r"^查下單\s+(\S+)$", raw)
        if m:
            pn   = m.group(1)
            rows = find_rows("採購單", 1, pn)
            if not rows:
                reply(f"{pn}\n尚無採購紀錄")
                return
            lines = [f"  {r[0]} | {r[2]} | {r[4]}個 | {r[6]}" for r in rows[-3:]]
            reply(f"【採購紀錄（最近3筆）】\n料號：{pn}\n\n" + "\n".join(lines))
            return

        m = re.match(r"^查交期\s+(\S+)$", raw)
        if m:
            pn      = m.group(1)
            rows    = find_rows("採購單", 1, pn)
            pending = [r for r in rows if len(r) > 6 and r[6] != "完成"]
            if not pending:
                reply(f"{pn}\n目前無待交貨採購單")
                return
            lines = [f"  {r[0]} | 交期：{r[5]} | {r[4]}個" for r in pending]
            reply(f"【待交貨】\n料號：{pn}\n\n" + "\n".join(lines))
            return

        m = re.match(r"^入庫\s+(\S+)\s+(\d+)$", raw)
        if m:
            pn, qty  = m.group(1), int(m.group(2))
            ws       = sh.worksheet("庫存管理")
            all_rows = ws.get_all_values()
            for i, row in enumerate(all_rows[1:], start=2):
                if row and str(row[0]).strip().upper() == pn.upper():
                    old = int(row[2]) if row[2] else 0
                    ws.update_cell(i, 3, old + qty)
                    ws.update_cell(i, 5, now)
                    reply(f"入庫完成\n料號：{pn}\n入庫：{qty}\n庫存：{old} → {old + qty}")
                    return
            erp_row = find_in_erp(pn)
            ws.append_row([pn, erp_row[4] if erp_row else "", qty, "", now])
            reply(f"入庫完成（新建庫存）\n料號：{pn}\n庫存：{qty}")
            return

        m = re.match(r"^出庫\s+(\S+)\s+(\d+)$", raw)
        if m:
            pn, qty  = m.group(1), int(m.group(2))
            ws       = sh.worksheet("庫存管理")
            all_rows = ws.get_all_values()
            for i, row in enumerate(all_rows[1:], start=2):
                if row and str(row[0]).strip().upper() == pn.upper():
                    old = int(row[2]) if row[2] else 0
                    if old < qty:
                        reply(f"庫存不足\n料號：{pn}\n現有：{old}，出庫：{qty}")
                        return
                    ws.update_cell(i, 3, old - qty)
                    ws.update_cell(i, 5, now)
                    reply(f"出庫完成\n料號：{pn}\n出庫：{qty}\n庫存：{old} → {old - qty}")
                    return
            reply(f"找不到庫存紀錄：{pn}")
            return

        m = re.match(r"^下單\s+(\S+)\s+(\S+)\s+(\d+)\s+(\S+)$", raw)
        if m:
            pn, vendor, qty, delivery = m.group(1), m.group(2), int(m.group(3)), m.group(4)
            today   = datetime.now(_TW).strftime("%Y%m%d")
            po_rows = sh.worksheet("採購單").get_all_values()[1:]
            prefix  = f"PO-{today}-"
            max_seq = max((int(str(r[0])[-3:]) for r in po_rows if r and str(r[0]).startswith(prefix)), default=0)
            po_no   = f"{prefix}{str(max_seq + 1).zfill(3)}"
            sh.worksheet("採購單").append_row([po_no, pn, vendor, now[:10], qty, delivery, "待交貨", ""])
            reply(f"採購單建立\n單號：{po_no}\n料號：{pn}\n廠商：{vendor}\n數量：{qty}\n交期：{delivery}")
            return

        m = re.match(r"^到貨\s+(\S+)\s+(\d+)$", raw)
        if m:
            po_or_pn, qty = m.group(1), int(m.group(2))
            ws_po   = sh.worksheet("採購單")
            po_rows = ws_po.get_all_values()
            pn      = po_or_pn
            for i, row in enumerate(po_rows[1:], start=2):
                if row and str(row[0]).strip().upper() == po_or_pn.upper():
                    pn = row[1]
                    ws_po.update_cell(i, 7, "完成")
                    break
            sh.worksheet("交貨紀錄").append_row([po_or_pn, pn, now[:10], qty, ""])
            ws_inv   = sh.worksheet("庫存管理")
            inv_rows = ws_inv.get_all_values()
            updated  = False
            for i, row in enumerate(inv_rows[1:], start=2):
                if row and str(row[0]).strip().upper() == pn.upper():
                    old = int(row[2]) if row[2] else 0
                    ws_inv.update_cell(i, 3, old + qty)
                    ws_inv.update_cell(i, 5, now)
                    updated = True
                    break
            if not updated:
                erp_row = find_in_erp(pn)
                ws_inv.append_row([pn, erp_row[4] if erp_row else "", qty, "", now])
            reply(f"到貨完成\n採購單：{po_or_pn}\n料號：{pn}\n到貨：{qty}件\n庫存已更新")
            return

        m = re.match(r"^承認\s+(\S+)\s+(\S+)\s+(\S+)$", raw)
        if m:
            pn, vendor, status = m.group(1), m.group(2), m.group(3)
            sh.worksheet("承認狀況").append_row([pn, vendor, status, now[:10], ""])
            reply(f"承認狀況更新\n料號：{pn}\n廠商：{vendor}\n狀態：{status}")
            return

        if raw == "我的ID":
            reply(f"你的 LINE ID：\n{user_id}")
            return

        reply("指令格式：\n查 料號\n查庫存 料號\n查承認 料號\n查下單 料號\n查交期 料號\n入庫 料號 數量\n出庫 料號 數量\n下單 料號 廠商 數量 交期\n到貨 採購單號 數量\n承認 料號 廠商 狀態")

    except Exception as e:
        print(traceback.format_exc(), flush=True)
        _line_reply(reply_token, "處理失敗，請稍後再試")


@app.route("/erp", methods=["POST"])
def webhook_erp():
    events = request.json.get("events", [])
    for event in events:
        if event.get("type") != "message":
            continue
        msg_type    = event["message"].get("type")
        user_id     = event.get("source", {}).get("userId", "")
        reply_token = event["replyToken"]

        if msg_type == "image":
            handle_invoice_image(user_id, event["message"]["id"], reply_token)
        elif msg_type == "text":
            handle_erp(event["message"]["text"], reply_token, user_id)
    return "OK"


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
