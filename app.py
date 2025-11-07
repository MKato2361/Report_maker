# app.py
# ------------------------------------------------------------
# 故障メール → 正規表現抽出 → 既存テンプレ(.xlsm)へ書込み → ダウンロード
# 3ステップUI / パスコード認証 / 編集対応（複数行入力保持）
# ------------------------------------------------------------
import io
import re
import unicodedata
from datetime import datetime, timedelta, timezone
from typing import Dict, Optional, Tuple, List
import os
from openpyxl import load_workbook
import streamlit as st

JST = timezone(timedelta(hours=9))

APP_TITLE = "故障報告メール → Excel自動生成（マクロ対応・編集可）"
PASSCODE_DEFAULT = "1357"
PASSCODE = st.secrets.get("APP_PASSCODE", PASSCODE_DEFAULT)

SHEET_NAME = "緊急出動報告書（リンク付き）"
WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]

# ====== ユーティリティ ======
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
    cand = s.strip().replace("年", "/").replace("月", "/").replace("日", "").replace("-", "/")
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d"):
        try:
            return datetime.strptime(cand, fmt)
        except Exception:
            pass
    return None

def _split_dt_components(dt: Optional[datetime]):
    if not dt:
        return None, None, None, None, None, None
    return dt.year, dt.month, dt.day, WEEKDAYS_JA[dt.weekday()], dt.hour, dt.minute

def _first_date_yyyymmdd(*vals) -> str:
    for v in vals:
        dt = _try_parse_datetime(v)
        if dt:
            return dt.strftime("%Y%m%d")
    return datetime.now().strftime("%Y%m%d")

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

    multiline_labels = {"受信内容": r"受信内容\s*:", "現着状況": r"現着状況\s*:", "原因": r"原因\s*:", "処置内容": r"処置内容\s*:"}

    out = {k: None for k in single_line.keys() | multiline_labels.keys()}
    out.update({"案件種別(件名)": subject_case, "受付URL": None, "現着完了登録URL": None})

    for k, pat in single_line.items():
        out[k] = _search_one(pat, t, flags=re.IGNORECASE | re.MULTILINE)
    if not out["管理番号"] and subject_manageno:
        out["管理番号"] = subject_manageno
    for k in multiline_labels:
        out[k] = _search_span_between(multiline_labels, k, t)

    return out

# ====== テンプレ書き込み ======
def fill_template_xlsx(template_bytes: bytes, data: Dict[str, Optional[str]]) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    # --- 複数行入力保持（セル内改行対応） ---
    def write_multiline(cell_ref: str, text: Optional[str]):
        if text:
            ws[cell_ref] = text.replace("\r\n", "\n").replace("\r", "\n")
            ws[cell_ref].alignment = ws[cell_ref].alignment.copy(wrapText=True)
        else:
            ws[cell_ref] = ""

    if data.get("管理番号"): ws["C12"] = data["管理番号"]
    if data.get("メーカー"): ws["J12"] = data["メーカー"]
    if data.get("制御方式"): ws["M12"] = data["制御方式"]
    if data.get("通報者"): ws["C14"] = data["通報者"]
    if data.get("対応者"): ws["L37"] = data["対応者"]
    if data.get("所属"): ws["C37"] = data["所属"]

    now = datetime.now(JST)
    ws["B5"], ws["D5"], ws["F5"] = now.year, now.month, now.day

    write_multiline("C15", data.get("受信内容"))
    write_multiline("C20", data.get("現着状況"))
    write_multiline("C25", data.get("原因"))
    write_multiline("C30", data.get("処置内容"))
    write_multiline("C35", data.get("処理修理後"))

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def build_filename(data):
    base_day = _first_date_yyyymmdd(data.get("現着時刻"), data.get("完了時刻"), data.get("受信時刻"))
    mn = (data.get("管理番号") or "UNKNOWN").replace("/", "_")
    nm = (data.get("物件名") or "").replace("/", "_")
    return f"緊急出動報告書_{mn}_{nm}_{base_day}.xlsm"

# ====== UI ======
st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)

if "step" not in st.session_state: st.session_state.step = 1
if "authed" not in st.session_state: st.session_state.authed = False
if "extracted" not in st.session_state: st.session_state.extracted = None

# Step 1
if st.session_state.step == 1:
    st.subheader("Step 1. パスコード認証")
    pw = st.text_input("パスコードを入力してください", type="password")
    if st.button("次へ"):
        if pw == PASSCODE:
            st.session_state.authed = True
            st.session_state.step = 2
            st.rerun()
        else:
            st.error("パスコードが違います。")

# Step 2
elif st.session_state.step == 2 and st.session_state.authed:
    st.subheader("Step 2. メール本文の貼り付け / 所属")
    path = "template.xlsm"
    if os.path.exists(path):
        with open(path, "rb") as f:
            st.session_state.template_xlsx_bytes = f.read()
        st.success("テンプレートを読み込みました。")
    else:
        st.error(f"テンプレートが見つかりません: {path}")
        st.stop()

    aff = st.text_input("所属", value=st.session_state.get("affiliation", ""))
    st.session_state.affiliation = aff
    text = st.text_area("故障完了メール（本文）を貼り付け", height=240)

    if st.button("抽出する"):
        st.session_state.extracted = extract_fields(text)
        st.session_state.extracted["所属"] = aff
        st.session_state.step = 3
        st.rerun()

# Step 3
elif st.session_state.step == 3 and st.session_state.authed:
    st.subheader("Step 3. 抽出結果の確認・編集 → Excel生成")
    data = st.session_state.extracted or {}

    # --- 編集フィールド ---
    data["通報者"] = st.text_area("通報者", value=data.get("通報者", ""), height=40)
    data["受信内容"] = st.text_area("受信内容", value=data.get("受信内容", ""), height=100)
    data["現着状況"] = st.text_area("現着状況", value=data.get("現着状況", ""), height=100)
    data["原因"] = st.text_area("原因", value=data.get("原因", ""), height=100)
    data["処置内容"] = st.text_area("処置内容", value=data.get("処置内容", ""), height=100)
    data["処理修理後"] = st.text_area("処理修理後", value=data.get("処理修理後", ""), height=80)

    st.divider()
    try:
        xlsx = fill_template_xlsx(st.session_state.template_xlsx_bytes, data)
        st.download_button("Excelを生成（.xlsm）", data=xlsx,
                           file_name=build_filename(data),
                           mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                           use_container_width=True)
    except Exception as e:
        st.error(f"Excel生成中にエラー: {e}")

    if st.button("Step2に戻る"): 
        st.session_state.step = 2
        st.rerun()

else:
    st.warning("認証が必要です。")
    st.session_state.step = 1
