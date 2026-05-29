# -*- coding: utf-8 -*-
"""
PDF 工程圖料號分析器
將 OneDrive 資料夾內的 PDF 逐一分析，
用 Gemini OCR 擷取 Title Block，
依照分類規則產生料號，並寫入 Google Sheets AI_PARTS_ERP
"""

import os
import sys
import json
import re
import time
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import fitz
import requests
from google import genai
from google.genai import types
import gspread
from google.oauth2.service_account import Credentials

# ── 設定 ──────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
CONFIG     = json.loads((BASE_DIR / "parts_config.json").read_text(encoding="utf-8"))
PDF_FOLDER = Path(os.path.expandvars(CONFIG["pdf_folder"]))
SHEET_NAME = CONFIG["google"]["spreadsheet_name"]
CREDS_FILE = BASE_DIR / CONFIG["google"]["credentials_file"]
SCOPES     = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

GEMINI_MODEL   = CONFIG.get("gemini_model", "gemini-2.5-flash")
GEMINI_API_KEY = CONFIG.get("gemini_api_key") or os.environ.get("GEMINI_API_KEY", "")

# ── OCR Prompt ────────────────────────────────────────
PROMPT_MAIN = (
    "你是專業機構工程ERP AI。\n"
    "請只分析工程圖右下角的Title Block區域。\n"
    "不要分析尺寸。\n"
    "不要分析圖形。\n"
    "不要分析註解。\n"
    "不要分析表格內容。\n"
    "只分析Title Block。\n"
    "只回傳JSON。\n"
    "不要解釋。\n"
    "project: 只能抓 PROJECT 欄位的完整文字，包含所有單詞，不得截斷或省略任何字元。\n"
    "drawing: 只能抓 DRAWING NAME 欄位。\n"
    "designer: 只能抓 DESIGNER 或 DRAWN BY 欄位。\n"
    "material: 只能抓 Title Block 中材料欄位的值，"
    "欄位標題可能是 MATL、MAT'L、MAT.L、MATERIAL、MATERIALS、MATL. 等任何變形（含撇號）。\n"
    "如果找不到材料欄位，請回傳 UNKNOWN。\n"
    "禁止抓取 FINISH。\n"
    "禁止抓取 SPEC。\n"
    "禁止抓取 NOTE。\n"
    "禁止抓取 DESCRIPTION。\n"
    "revision: 只能抓 REV、REVISION 或 VER 欄位，讀取該欄位標題下方或旁邊的值。如果是空白請回傳 0。\n"
    "drawing_no: 只能抓 DRAWING NO、DWG NO、DWG. NO 或 NO. 欄位的值。如果找不到請回傳空字串。\n"
    "finish: 抓工程圖中是否出現 FINISH 文字。\n"
)

PROMPT_MATERIAL_FALLBACK = (
    "你是專業機構工程ERP AI。\n"
    "請只分析工程圖右下角的Title Block中的 MATL 或 MATERIAL 欄位。\n"
    "只回傳 JSON。\n"
    "不要解釋。\n"
    "material: 只能抓 Title Block 中材料欄位的值，欄位標題可能是 MATL、MAT'L、MAT.L、MATERIAL 等任何變形（含撇號）。\n"
    "如果找不到請回傳 UNKNOWN。\n"
)


# ═══════════════════════════════════════════════════
# Google Sheets 連線
# ═══════════════════════════════════════════════════

def connect_sheets():
    """連線 Google Sheets，回傳三個工作表物件和預載資料"""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if creds_json:
        creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(str(CREDS_FILE), scopes=SCOPES)
    sh    = gspread.authorize(creds).open(SHEET_NAME)

    ws_erp  = sh.worksheet("ERP資料庫")
    ws_fp   = sh.worksheet("圖面特徵庫")
    ws_rule = sh.worksheet("分類規則")

    # 一次讀取所有資料到記憶體，減少 API 呼叫
    erp_rows  = ws_erp.get_all_values()[1:]   # 跳過標題列
    fp_rows   = ws_fp.get_all_values()[1:]
    rule_rows = ws_rule.get_all_values()[1:]

    return ws_erp, ws_fp, erp_rows, fp_rows, rule_rows


# ═══════════════════════════════════════════════════
# Gemini OCR
# ═══════════════════════════════════════════════════

def call_gemini(client, prompt: str, image_bytes: bytes) -> dict:
    image_part = types.Part.from_bytes(data=image_bytes, mime_type="image/png")
    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        temperature=0.1
    )
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model=GEMINI_MODEL,
                contents=[prompt, image_part],
                config=config
            )
            return json.loads(response.text)
        except Exception as err:
            err_str = str(err)
            if "429" in err_str and "per_day" in err_str.lower():
                raise RuntimeError("今日 Gemini 免費額度已用完（每日上限 20 次），請明天再試") from err
            if ("429" in err_str or "503" in err_str) and attempt < 2:
                wait = 45 if "429" in err_str else 15
                print(f"    API 忙碌，等待 {wait} 秒後重試（第 {attempt+1} 次）...")
                time.sleep(wait)
                continue
            if attempt == 2:
                raise RuntimeError(f"Gemini API 失敗：{err}") from err


def analyze_drawing(client, image_bytes: bytes) -> dict:
    data     = call_gemini(client, PROMPT_MAIN, image_bytes)
    material = (data.get("material") or "").strip().upper()
    if not material or material == "UNKNOWN":
        try:
            fallback = call_gemini(client, PROMPT_MATERIAL_FALLBACK, image_bytes)
            mat = (fallback.get("material") or "UNKNOWN").strip()
            data["material"] = mat if mat else "UNKNOWN"
        except Exception:
            data["material"] = "UNKNOWN"
    return data


# ═══════════════════════════════════════════════════
# PDF 工具
# ═══════════════════════════════════════════════════

def pdf_page_to_bytes(pdf_path: Path, page_index: int = 0) -> bytes:
    doc  = fitz.open(str(pdf_path))
    page = doc[page_index if page_index < len(doc) else 0]
    pix  = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0))
    return pix.tobytes("png")


# ═══════════════════════════════════════════════════
# Rule Engine
# ═══════════════════════════════════════════════════

def load_rules(rule_rows: list) -> list:
    rules = []
    for row in rule_rows:
        if not row or not row[0]:
            continue
        rules.append({
            "keyword":  str(row[0]).strip(),
            "category": str(row[1]).strip(),
            "field":    str(row[2]).strip().upper(),
            "priority": int(row[3]) if len(row) > 3 and row[3] else 0,
        })
    return rules


def _normalize(text: str) -> str:
    return re.sub(r"[\s\-_]", "", text).upper()


def get_category(drawing: str, material: str, rules: list) -> str:
    norm_drawing  = _normalize(drawing)
    norm_material = _normalize(material)
    best_cat, best_pri = "UNKNOWN", -1
    for rule in rules:
        norm_kw = _normalize(rule["keyword"])
        matched = (
            (rule["field"] == "DRAWING NAME" and norm_kw in norm_drawing) or
            (rule["field"] == "MATL"          and norm_kw in norm_material)
        )
        if matched and rule["priority"] > best_pri:
            best_pri = rule["priority"]
            best_cat = rule["category"]
    return best_cat


# ═══════════════════════════════════════════════════
# 料號產生
# ═══════════════════════════════════════════════════

def _push_line(message: str):
    token  = CONFIG.get("line_erp_token", "")
    target = CONFIG.get("line_erp_target", "")
    if not token or not target:
        return
    try:
        resp = requests.post(
            "https://api.line.me/v2/bot/message/push",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"to": target, "messages": [{"type": "text", "text": message}]},
            timeout=10,
        )
        if resp.status_code != 200:
            print(f"    LINE 通知失敗 {resp.status_code}：{resp.text}")
        else:
            print(f"    LINE 通知已送出")
    except Exception as e:
        print(f"    LINE 通知失敗：{e}")


def _fp_normalize(text: str) -> str:
    t = (text or "").upper()
    t = re.sub(r"T\s*=\s*", "", t)        # 移除厚度前綴 T= / T =
    t = re.sub(r"[\s\-_.,;/\\]", "", t)   # 移除空格與常見符號
    return t


def make_fingerprint(project: str, drawing: str, material: str,
                     revision: str, drawing_no: str = "") -> str:
    return "|".join([
        _fp_normalize(project)    or "UNK",
        _fp_normalize(drawing)    or "UNKNOWN",
        _fp_normalize(revision)   or "X",
        _fp_normalize(drawing_no),
    ])


def _sig_words(text: str) -> list:
    """從文字中提取有意義的純英文單詞（長度 >= 2，去掉數字、符號）。"""
    clean = re.sub(r"[^A-Za-z]", "_", text).upper()
    return [w for w in clean.split("_") if len(w) >= 2]


def _words_to_abbrev(words: list, length: int) -> str:
    """
    將單詞列表壓縮成固定長度的全英文縮寫。
    盡量平均分配字元，尾端的詞多分一點（更具識別性）。
    """
    if not words:
        return "X" * length
    n = min(len(words), length)
    words = words[:n]
    base, extra = divmod(length, n)
    # 後面的詞多分一個字元（尾端更有辨識度）
    alloc = [base + (1 if i >= n - extra else 0) for i in range(n)]
    result = "".join(w[:c] for w, c in zip(words, alloc))
    return (result + "X" * length)[:length]


def _make_6char(project: str, drawing: str) -> str:
    """
    產生 7-12 碼（6字元）的純英文縮寫：前3碼來自專案，後3碼來自圖名特徵詞。
    - 過濾掉與 project 重複的單詞，讓 drawing 的後3碼更有辨識度。
    """
    proj_words = _sig_words(project)
    draw_words = _sig_words(drawing)

    # 前3碼：project 第一個有意義單詞
    proj_part = (proj_words[0] + "XXX")[:3] if proj_words else "UNK"

    # 後3碼：drawing 去掉與 project 重疊的詞後，取特徵詞縮寫
    proj_top = proj_words[0] if proj_words else ""
    draw_uniq = [w for w in draw_words if w != proj_top] or draw_words

    draw_part = _words_to_abbrev(draw_uniq, 3)
    return proj_part + draw_part


def _check_prefix(candidate_6: str, category: str, this_drawing: str,
                  erp_rows: list) -> tuple:
    """
    檢查候選 6 碼是否與「不同圖面」發生碰撞。
    回傳 (collision: bool, max_seq: int)
    - collision=True  → 這個 6 碼已被別張圖佔用
    - max_seq         → 同張圖已有的最大流水號
    """
    this_norm = _fp_normalize(this_drawing)
    prefix    = f"M-{category}-{candidate_6}-"
    collision = False
    max_seq   = 0
    for row in erp_rows:
        pn = str(row[1]) if len(row) > 1 else ""
        if not pn.startswith(prefix):
            continue
        existing_draw = str(row[4]) if len(row) > 4 else ""  # ERP 欄位[4] = drawing
        if _fp_normalize(existing_draw) != this_norm:
            collision = True          # 不同圖面佔用了同前綴 → 碰撞
            break
        # 同圖面：記錄最大流水號（同圖面變體、版本）
        try:
            max_seq = max(max_seq, int(pn[-2:]))
        except ValueError:
            pass
    return collision, max_seq


def create_part_number(category: str, project: str, drawing: str,
                       revision: str, erp_rows: list) -> str:
    """
    產生料號：M-{category}-{PPPDDD}-{rev}{seq:02d}
    ─────────────────────────────────────────────
    7-12碼（PPPDDD）：純 A-Z，智慧多詞縮寫 + 碰撞自動迴避
      • 前3碼(PPP)：project 特徵詞
      • 後3碼(DDD)：drawing 特徵詞（排除與 PPP 重複的詞）
      • 碰撞時：保持 PPP 不變，後3碼末字元依序換成 A-Z
    流水號：僅對「同一張圖」遞增；不同圖面一定要有不同前綴
    特徵值(RVE)比對：由 find_existing_by_fingerprint 在呼叫前完成，此處不重複
    """
    rev = (revision or "0").strip().upper()[:1] or "0"

    proj_words = _sig_words(project)
    draw_words = _sig_words(drawing)
    proj_top   = proj_words[0] if proj_words else ""
    draw_uniq  = [w for w in draw_words if w != proj_top] or draw_words

    proj_part  = (proj_top + "XXX")[:3] if proj_top else "UNK"

    # ── 嘗試最多 27 種後3碼：原始縮寫 → 末碼換 A → B → … → Z ──
    for attempt in range(27):
        if attempt == 0:
            draw_part = _words_to_abbrev(draw_uniq, 3)
        else:
            # 保留前2碼，第3碼換字母 A-Z
            base2     = _words_to_abbrev(draw_uniq, 2)
            draw_part = base2 + chr(ord("A") + attempt - 1)

        code_6 = proj_part + draw_part
        collision, max_seq = _check_prefix(code_6, category, drawing, erp_rows)
        if not collision:
            prefix = f"M-{category}-{code_6}-{rev}"
            return f"{prefix}{str(max_seq + 1).zfill(2)}"

    # ── 極端 fallback（理論上不應觸發）──
    code_6 = _make_6char(project, drawing)
    prefix = f"M-{category}-{code_6}-{rev}"
    return f"{prefix}01"


# ═══════════════════════════════════════════════════
# 去重檢查（使用記憶體資料）
# ═══════════════════════════════════════════════════

def find_existing_by_fingerprint(fingerprint: str, fp_rows: list):
    """
    比對特徵值，回傳已存在的料號；找不到回傳 None。
    相容兩種欄位格式：
      新格式（無 file_path / 無空白H欄）：特徵值在索引 [3]
      舊格式（含 file_path）              ：特徵值在索引 [4]
    判斷依據：特徵值一定包含 '|'，用此識別欄位是否正確。
    """
    fp_parts = fingerprint.split("|")
    fp_old   = "|".join(fp_parts[:4])   # 舊格式（4欄）相容
    for row in fp_rows:
        # 依序嘗試新格式索引 [3] 和舊格式索引 [4]
        for idx in (3, 4):
            if len(row) <= idx or not row[idx]:
                continue
            existing = str(row[idx]).strip().upper()
            if "|" not in existing:      # 不是特徵值欄位（可能是 file_path），跳過
                continue
            if existing == fingerprint:
                return str(row[0])
            existing_old = "|".join(existing.split("|")[:4])
            if existing_old == fp_old:
                return str(row[0])
    return None


def part_number_exists(part_number: str, erp_rows: list) -> bool:
    pn_upper = part_number.upper()
    return any(
        len(row) > 1 and str(row[1]).strip().upper() == pn_upper
        for row in erp_rows
    )


# ═══════════════════════════════════════════════════
# 主流程
# ═══════════════════════════════════════════════════

def process_pdf(pdf_path: Path, client, ws_erp, ws_fp,
                erp_rows: list, fp_rows: list, rules: list) -> dict:
    print(f"\n  處理：{pdf_path.name}")

    image_bytes = pdf_page_to_bytes(pdf_path)
    data        = analyze_drawing(client, image_bytes)

    print(
        f"    project={data.get('project')} | "
        f"drawing={data.get('drawing')} | "
        f"material={data.get('material')} | "
        f"rev={data.get('revision')}"
    )

    fingerprint = make_fingerprint(
        data.get("project", ""), data.get("drawing", ""),
        data.get("material", ""), data.get("revision", ""),
        data.get("drawing_no", "")
    )
    existing = find_existing_by_fingerprint(fingerprint, fp_rows)
    if existing:
        print(f"    [重複] 圖面已有料號：{existing}")
        _push_line(
            f"[重複圖面通知]\n"
            f"檔案：{pdf_path.name}\n"
            f"Drawing Name：{data.get('drawing', '')}\n"
            f"Drawing NO：{data.get('drawing_no', '')}\n"
            f"已有料號：{existing}"
        )
        return {"status": "duplicate", "part_number": existing, "file": pdf_path.name}

    category    = get_category(data.get("drawing", ""), data.get("material", ""), rules)
    part_number = create_part_number(
        category, data.get("project", ""), data.get("drawing", ""),
        data.get("revision", ""), erp_rows
    )

    if part_number_exists(part_number, erp_rows):
        print(f"    [衝突] 料號已存在：{part_number}")
        return {"status": "conflict", "part_number": part_number, "file": pdf_path.name}

    now       = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_path = str(pdf_path.resolve())

    # 寫入 Google Sheets（明確指定列號，避免 append_row 欄位偏移）
    # file_path 已移除：圖紙位置會變動，不適合永久儲存
    erp_row = [now, part_number, category,
               data.get("project", ""), data.get("drawing", ""),
               data.get("material", ""), data.get("revision", ""),
               data.get("designer", ""), data.get("drawing_no", "")]
    erp_next = len(ws_erp.get_all_values()) + 1   # 取得下一個空白列（1-indexed）
    ws_erp.update(f"A{erp_next}", [erp_row])
    erp_rows.append(erp_row)  # 同步更新記憶體，避免同批次料號衝突

    # 新格式：part_number | drawing | material | fingerprint | revision | timestamp | drawing_no
    # （移除 file_path 欄與空白 H 欄）
    fp_row = [part_number, data.get("drawing", ""), data.get("material", ""),
              fingerprint, data.get("revision", ""), now,
              data.get("drawing_no", "")]
    fp_next = len(ws_fp.get_all_values()) + 1     # 取得下一個空白列（1-indexed）
    ws_fp.update(f"A{fp_next}", [fp_row])
    fp_rows.append(fp_row)

    # 將原始 PDF 改名為料號
    renamed = pdf_path.parent / f"{part_number}.pdf"
    try:
        pdf_path.rename(renamed)
        print(f"    [OK] 建立料號：{part_number}  分類：{category}  檔名已更新")
    except Exception as e:
        print(f"    [OK] 建立料號：{part_number}  分類：{category}（改名失敗：{e}）")
    return {"status": "created", "part_number": part_number, "file": renamed.name}


def main():
    if not GEMINI_API_KEY:
        print("[ERROR] 請先設定 GEMINI_API_KEY 環境變數")
        print("  PowerShell: $env:GEMINI_API_KEY = '你的金鑰'")
        sys.exit(1)

    if not PDF_FOLDER.exists():
        print(f"[ERROR] 找不到 PDF 資料夾：{PDF_FOLDER}")
        sys.exit(1)

    pdf_files = sorted(PDF_FOLDER.glob("*.pdf"))
    if not pdf_files:
        print(f"[ERROR] 資料夾內沒有 PDF 檔案：{PDF_FOLDER}")
        sys.exit(1)

    print(f"找到 {len(pdf_files)} 個 PDF 檔案，開始處理...")
    print(f"PDF 來源：{PDF_FOLDER}")
    print("-" * 50)

    print("連線 Google Sheets...")
    ws_erp, ws_fp, erp_rows, fp_rows, rule_rows = connect_sheets()
    rules = load_rules(rule_rows)
    print(f"載入分類規則：{len(rules)} 條")

    client  = genai.Client(api_key=GEMINI_API_KEY)
    results = []

    for pdf_path in pdf_files:
        try:
            result = process_pdf(pdf_path, client, ws_erp, ws_fp,
                                 erp_rows, fp_rows, rules)
        except Exception as err:
            print(f"    [ERROR] {err}")
            result = {"status": "error", "file": pdf_path.name, "error": str(err)}
        results.append(result)

    created   = sum(1 for r in results if r["status"] == "created")
    duplicate = sum(1 for r in results if r["status"] == "duplicate")
    conflict  = sum(1 for r in results if r["status"] == "conflict")
    errors    = sum(1 for r in results if r["status"] == "error")

    print("\n" + "-" * 50)
    print(f"完成！結果已寫入 Google Sheets：{SHEET_NAME}")
    print(f"建立 {created}  重複 {duplicate}  衝突 {conflict}  錯誤 {errors}")


if __name__ == "__main__":
    main()
