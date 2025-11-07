# app.py (Part 1/2)
# ------------------------------------------------------------
# 故障メール → 正規表現抽出 → 既存テンプレ(.xlsx)へ書込み → ダウンロード
# 3ステップUI / パスコード認証 / 編集不可 / 折りたたみ表示（時系列）
# 仕様反映：
#   - 曜日：日本語（例：月）
#   - 複数行：最大5行。超過は「…」付与
#   - 通報者：原文そのまま（様/電話番号含む）
#   - ファイル名：管理番号_物件名_日付（yyyymmdd）
#   - 時刻セル分割：年・月・日・曜・時・分（個別セル）
#   - rerun は st.rerun() を使用
# ------------------------------------------------------------
import io
import re
import unicodedata
from datetime import datetime
from typing import Dict, Optional, Tuple, List
from datetime import datetime, timedelta, timezone
JST = timezone(timedelta(hours=9))


import streamlit as st
from openpyxl import load_workbook

APP_TITLE = "故障報告メール → Excel自動生成"
PASSCODE_DEFAULT = "1357"  # 公開運用時は .streamlit/secrets.toml の APP_PASSCODE を推奨
PASSCODE = st.secrets.get("APP_PASSCODE", PASSCODE_DEFAULT)

# テンプレートのシート名（ユーザー共有仕様に準拠）
SHEET_NAME = "緊急出動報告書（リンク付き）"

# ====== ユーティリティ ======
WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]

def normalize_text(text: str) -> str:
    if not text:
        return ""
    t = unicodedata.normalize("NFKC", text)
    t = t.replace("：", ":")
    t = t.replace("\t", " ").replace("\r\n", "\n").replace("\r", "\n")
    return t

def _search_one(pattern: str, text: str, flags=0) -> Optional[str]:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else None

def _search_span_between(labels: Dict[str, str], key: str, text: str) -> Optional[str]:
    """
    ラベル key の位置から、次のいずれかのラベル直前までを抽出（複数行対応）
    """
    lab = labels[key]
    others = [v for k, v in labels.items() if k != key]
    boundary = "|".join([f"(?:{v})" for v in others]) if others else r"$"
    pattern = rf"{lab}\s*(.+?)(?=\n(?:{boundary})|\Z)"
    m = re.search(pattern, text, flags=re.DOTALL | re.IGNORECASE)
    return m.group(1).strip() if m else None

def _try_parse_datetime(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    cand = s.strip()
    cand = cand.replace("年", "/").replace("月", "/").replace("日", "")
    cand = cand.replace("-", "/")
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d"):
        try:
            return datetime.strptime(cand, fmt)
        except Exception:
            pass
    return None

def _split_dt_components(dt: Optional[datetime]) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[str], Optional[int], Optional[int]]:
    if not dt:
        return None, None, None, None, None, None
    y = dt.year
    m = dt.month
    d = dt.day
    wd = WEEKDAYS_JA[dt.weekday()]
    hh = dt.hour
    mm = dt.minute
    return y, m, d, wd, hh, mm

def _first_date_yyyymmdd(*vals) -> str:
    for v in vals:
        dt = _try_parse_datetime(v)
        if dt:
            return dt.strftime("%Y%m%d")
    return datetime.now().strftime("%Y%m%d")

def minutes_between(a: Optional[str], b: Optional[str]) -> Optional[int]:
    s = _try_parse_datetime(a)
    e = _try_parse_datetime(b)
    if s and e:
        return int((e - s).total_seconds() // 60)
    return None

def _split_lines(text: Optional[str], max_lines: int = 5) -> List[str]:
    if not text:
        return []
    lines = [ln.strip() for ln in text.splitlines() if ln.strip() != ""]
    if len(lines) <= max_lines:
        return lines
    kept = lines[: max_lines - 1] + [lines[max_lines - 1] + "…"]
    return kept

# ====== 正規表現 抽出 ======
def extract_fields(raw_text: str) -> Dict[str, Optional[str]]:
    """
    共有サンプルフォーマットに準拠して抽出
    """
    t = normalize_text(raw_text)

    subject_case = _search_one(r"件名:\s*【\s*([^】]+)\s*】", t, flags=re.IGNORECASE)
    subject_manageno = _search_one(r"件名:.*?【[^】]+】\s*([A-Z0-9\-]+)", t, flags=re.IGNORECASE)

    single_line = {
        "管理番号": r"管理番号\s*:\s*([A-Za-z0-9\-]+)",
        "物件名": r"物件名\s*:\s*(.+)",
        "住所": r"住所\s*:\s*(.+)",
        "窓口会社": r"窓口\s*:\s*(.+)",
        "メーカー": r"メーカー\s*:\s*(.+)",
        "制御方式": r"制御方式\s*:\s*(.+)",
        "契約種別": r"契約種別\s*:\s*(.+)",
        "受信時刻": r"受信時刻\s*:\s*([0-9/\-:\s]+)",
        "通報者": r"通報者\s*:\s*(.+)",
        "現着時刻": r"現着時刻\s*:\s*([0-9/\-:\s]+)",
        "完了時刻": r"完了時刻\s*:\s*([0-9/\-:\s]+)",
        "対応者": r"対応者\s*:\s*(.+)",
        "送信者": r"送信者\s*:\s*(.+)",
        "受付番号": r"受付番号\s*:\s*([0-9]+)",
        "受付URL": r"詳細はこちら\s*:\s*.*?(https?://\S+)",
        "現着完了登録URL": r"現着・完了登録はこちら\s*:\s*(https?://\S+)",
    }

    multiline_labels = {
        "受信内容": r"受信内容\s*:",
        "現着状況": r"現着状況\s*:",
        "原因": r"原因\s*:",
        "処置内容": r"処置内容\s*:",
        "通報者": r"通報者\s*:",
        "対応者": r"対応者\s*:",
        "送信者": r"送信者\s*:",
        "現着時刻": r"現着時刻\s*:",
        "完了時刻": r"完了時刻\s*:",
    }

    out = {
        "案件種別(件名)": subject_case,
        "管理番号": None,
        "物件名": None,
        "住所": None,
        "窓口会社": None,
        "メーカー": None,
        "制御方式": None,
        "契約種別": None,
        "受信時刻": None,
        "受信内容": None,
        "通報者": None,
        "現着時刻": None,
        "現着状況": None,
        "完了時刻": None,
        "原因": None,
        "処置内容": None,
        "対応者": None,
        "送信者": None,
        "受付番号": None,
        "受付URL": None,
        "現着完了登録URL": None,
    }

    for k, pat in single_line.items():
        out[k] = _search_one(pat, t, flags=re.IGNORECASE | re.MULTILINE)

    if not out["管理番号"] and subject_manageno:
        out["管理番号"] = subject_manageno

    for k in multiline_labels:
        out[k] = _search_span_between(multiline_labels, k, t)

    dur = minutes_between(out["現着時刻"], out["完了時刻"])
    out["作業時間_分"] = str(dur) if dur is not None and dur >= 0 else None
    return out
# ====== テンプレ書き込み ======
# --- 中略（あなたの Part1/2 のまま残してください） ---


def fill_template_xlsx(template_bytes: bytes, data: Dict[str, Optional[str]]) -> bytes:
    """
    .xlsx テンプレ（.xlsb を Excel で保存し直したもの）に値を書き込んで返す
    """
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    # --- 単項目 ---
    if data.get("管理番号"): ws["C12"] = data["管理番号"]
    if data.get("メーカー"): ws["J12"] = data["メーカー"]
    if data.get("制御方式"): ws["M12"] = data["制御方式"]
    if data.get("受信内容"): ws["C15"] = data["受信内容"]
    if data.get("通報者"): ws["C14"] = data["通報者"]
    if data.get("対応者"): ws["L37"] = data["対応者"]

    # --- 処理修理後（C35） ---
    pa = st.session_state.get("processing_after")
    if pa:
        ws["C35"] = pa
    if data.get("所属"): ws["C37"] = data["所属"]

    # --- 現在日付をB5/D5/F5へ書き込み ---
    now = datetime.now(JST)
    ws["B5"] = now.year
    ws["D5"] = now.month
    ws["F5"] = now.day

    # --- 日付時刻分割 ---
    def write_dt_block(base_row: int, src_key: str):
        dt = _try_parse_datetime(data.get(src_key))
        y, m, d, wd, hh, mm = _split_dt_components(dt)
        if y is not None: ws[f"C{base_row}"] = y
        if m is not None: ws[f"F{base_row}"] = m
        if d is not None: ws[f"H{base_row}"] = d
        if wd is not None: ws[f"J{base_row}"] = wd
        if hh is not None: ws[f"M{base_row}"] = f"{hh:02d}"
        if mm is not None: ws[f"O{base_row}"] = f"{mm:02d}"

    write_dt_block(13, "受信時刻")
    write_dt_block(19, "現着時刻")
    write_dt_block(36, "完了時刻")

    # --- 複数行入力 ---
    def fill_multiline(col, start_row, text, max_lines=5):
        lines = _split_lines(text, max_lines=max_lines)
        for i in range(max_lines):
            ws[f"{col}{start_row + i}"] = ""
        for i, line in enumerate(lines):
            ws[f"{col}{start_row + i}"] = line

    fill_multiline("C", 20, data.get("現着状況"))
    fill_multiline("C", 25, data.get("原因"))
    fill_multiline("C", 30, data.get("処置内容"))

    # --- ✅ チェックボックス画像を貼り付け (I10:P11範囲) ---
    try:
        from openpyxl.drawing.image import Image as XLImage
        import base64

        checkbox_base64 = """
iVBORw0KGgoAAAANSUhEUgAAAgAAAAAoCAYAAAAJr9Y/AAAACXBIWXMAAAsTAAALEwEAmpwYAAABhUlEQVR4nO3dMUoDQRCG4f8tAciIp7A3SA3o
HZABGoRkD2kQ2w1QtwQ0bBgoDRqAYWBJSw8BkoHRmHgty6F1u33/9Z1oXnqE0m6Zn+7W9dNbtpYBAAAAAAAAAAAAAAAAAPyhyZzuWZhvf2ezuzxkzq
VmZqZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZm
ZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZmZm
AAAAAAAAAAAAAAAAAAAAAPgN/wECqE4hx9I5pQAAAABJRU5ErkJggg==
        """.replace("\n", "")

        img_data = base64.b64decode(checkbox_base64)
        img = XLImage(io.BytesIO(img_data))
        img.width = 480  # I10:P11範囲に合うよう調整
        img.height = 60
        ws.add_image(img, "I10")

    except Exception as e:
        print(f"チェックボックス貼付エラー: {e}")

    # --- 出力 ---
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
