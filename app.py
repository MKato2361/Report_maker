# app.py
# ------------------------------------------------------------
# 故障メール → 正規表現抽出 → 既存テンプレ(.xlsm)へ書込み → ダウンロード
# 3ステップUI / 編集UI対応（通報者・受信内容・現着状況・原因・処置内容・処理修理後）
# ------------------------------------------------------------
import io
import re
import unicodedata
from datetime import datetime, timedelta, timezone
from typing import Dict, Optional, Tuple, List
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
import streamlit as st

JST = timezone(timedelta(hours=9))

APP_TITLE = "故障報告メール → Excel自動生成（マクロ対応）"
PASSCODE_DEFAULT = "1357"  # 公開運用時は .streamlit/secrets.toml の APP_PASSCODE を推奨
PASSCODE = st.secrets.get("APP_PASSCODE", PASSCODE_DEFAULT)

SHEET_NAME = "緊急出動報告書（リンク付き）"
WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]

# ====== テキスト整形・抽出ユーティリティ ======
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
        return ["" for _ in range(max_lines)]
    lines = [ln.strip() for ln in text.splitlines() if ln.strip() != ""]
    if len(lines) < max_lines:
        lines += [""] * (max_lines - len(lines))
    else:
        lines = lines[:max_lines]
    return lines

# ====== 正規表現 抽出 ======
def extract_fields(raw_text: str) -> Dict[str, Optional[str]]:
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
    }

    out = {k: None for k in single_line.keys() | multiline_labels.keys()}
    out.update({
        "案件種別(件名)": subject_case,
        "受付URL": None,
        "現着完了登録URL": None,
    })

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
def fill_template_xlsx(template_bytes: bytes, data: Dict[str, Optional[str]]) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    # --- 基本情報 ---
    if data.get("管理番号"): ws["C12"] = data["管理番号"]
    if data.get("メーカー"): ws["J12"] = data["メーカー"]
    if data.get("制御方式"): ws["M12"] = data["制御方式"]
    if data.get("受信内容"): ws["C15"] = data["受信内容"]
    if data.get("通報者"): ws["C14"] = data["通報者"]
    if data.get("対応者"): ws["L37"] = data["対応者"]
    pa = st.session_state.get("processing_after")
    if pa:
        ws["C35"] = pa
    if data.get("所属"): ws["C37"] = data["所属"]
    now = datetime.now(JST)
    ws["B5"], ws["D5"], ws["F5"] = now.year, now.month, now.day

    # --- 日付・時刻ブロック ---
    def write_dt_block(base_row: int, src_key: str):
        dt = _try_parse_datetime(data.get(src_key))
        y, m, d, wd, hh, mm = _split_dt_components(dt)
        cellmap = {"Y": f"C{base_row}", "Mo": f"F{base_row}", "D": f"H{base_row}",
                   "W": f"J{base_row}", "H": f"M{base_row}", "Min": f"O{base_row}"}
        if y is not None: ws[cellmap["Y"]] = y
        if m is not None: ws[cellmap["Mo"]] = m
        if d is not None: ws[cellmap["D"]] = d
        if wd is not None: ws[cellmap["W"]] = wd
        if hh is not None: ws[cellmap["H"]] = f"{hh:02d}"
        if mm is not None: ws[cellmap["Min"]] = f"{mm:02d}"

    write_dt_block(13, "受信時刻")
    write_dt_block(19, "現着時刻")
    write_dt_block(36, "完了時刻")

    # --- 複数行ブロック ---
    def fill_multiline(col_letter: str, start_row: int, text: Optional[str], max_lines: int = 5):
        lines = _split_lines(text, max_lines=max_lines)
        for i in range(max_lines):
            ws[f"{col_letter}{start_row + i}"] = ""
        for idx, line in enumerate(lines[:max_lines]):
            ws[f"{col_letter}{start_row + idx}"] = line

    fill_multiline("C", 20, data.get("現着状況"))
    fill_multiline("C", 25, data.get("原因"))
    fill_multiline("C", 30, data.get("処置内容"))

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ====== Step3: 編集UI追加 ======
def editable_field(label, key, max_lines=1):
    data = st.session_state.extracted
    edit_key = f"edit_{key}"
    if edit_key not in st.session_state:
        st.session_state[edit_key] = False

    if not st.session_state[edit_key]:
        value = data.get(key) or ""
        lines = _split_lines(value, max_lines) if max_lines > 1 else [value]
        st.markdown(f"**{label}：**<br>{'<br>'.join(lines)}", unsafe_allow_html=True)
        if st.button("✏️ 編集", key=f"btn_{key}"):
            st.session_state[edit_key] = True
            st.rerun()
    else:
        st.markdown(f"✏️ **{label} 編集中**")
        value = data.get(key) or ""
        if max_lines == 1:
            new_val = st.text_input("内容を入力", value=value, key=f"in_{key}")
        else:
            new_val = st.text_area("内容を入力", value=value, height=max_lines * 25, key=f"ta_{key}")
        if st.button("💾 保存", key=f"save_{key}"):
            st.session_state.extracted[key] = new_val
            st.session_state[edit_key] = False
            st.rerun()

# ====== Streamlit UI ======
st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)

if "step" not in st.session_state:
    st.session_state.step = 1
if "authed" not in st.session_state:
    st.session_state.authed = False
if "extracted" not in st.session_state:
    st.session_state.extracted = None
if "affiliation" not in st.session_state:
    st.session_state.affiliation = ""

# Step1～Step2 は元のまま省略（変更なし）

# Step3: 抽出結果確認
elif st.session_state.step == 3 and st.session_state.authed:
    st.subheader("Step 3. 抽出結果の確認・編集 → Excel生成")

    data = st.session_state.extracted or {}
    with st.expander("編集可能項目", expanded=True):
        editable_field("通報者", "通報者", 1)
        editable_field("受信内容", "受信内容", 4)
        editable_field("現着状況", "現着状況", 5)
        editable_field("原因", "原因", 5)
        editable_field("処置内容", "処置内容", 5)
        editable_field("処理修理後（Step2入力値）", "処理修理後", 1)
