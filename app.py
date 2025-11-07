# app.py
# ------------------------------------------------------------
# 故障メール → 正規表現抽出 → 既存テンプレ(.xlsx)へ書込み → ダウンロード
# チェックボックス保持対応版（I10〜P11に一切書き込まない）
# ------------------------------------------------------------
import io
import re
import unicodedata
import tempfile
import shutil
from datetime import datetime, timedelta, timezone
from typing import Dict, Optional, Tuple, List
import streamlit as st
from openpyxl import load_workbook

JST = timezone(timedelta(hours=9))

APP_TITLE = "故障報告メール → Excel自動生成"
PASSCODE_DEFAULT = "1357"
PASSCODE = st.secrets.get("APP_PASSCODE", PASSCODE_DEFAULT)
SHEET_NAME = "緊急出動報告書（リンク付き）"

# ====== 共通ユーティリティ ======
WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]

def normalize_text(text: str) -> str:
    if not text:
        return ""
    t = unicodedata.normalize("NFKC", text)
    t = t.replace("：", ":").replace("\t", " ").replace("\r\n", "\n").replace("\r", "\n")
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

def _split_dt_components(dt: Optional[datetime]) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[str], Optional[int], Optional[int]]:
    if not dt:
        return None, None, None, None, None, None
    y, m, d = dt.year, dt.month, dt.day
    wd = WEEKDAYS_JA[dt.weekday()]
    hh, mm = dt.hour, dt.minute
    return y, m, d, wd, hh, mm

def _first_date_yyyymmdd(*vals) -> str:
    for v in vals:
        dt = _try_parse_datetime(v)
        if dt:
            return dt.strftime("%Y%m%d")
    return datetime.now().strftime("%Y%m%d")

def minutes_between(a: Optional[str], b: Optional[str]) -> Optional[int]:
    s, e = _try_parse_datetime(a), _try_parse_datetime(b)
    if s and e:
        return int((e - s).total_seconds() // 60)
    return None

def _split_lines(text: Optional[str], max_lines: int = 5) -> List[str]:
    if not text:
        return []
    lines = [ln.strip() for ln in text.splitlines() if ln.strip() != ""]
    if len(lines) <= max_lines:
        return lines
    return lines[: max_lines - 1] + [lines[max_lines - 1] + "…"]

# ====== 正規表現 抽出 ======
def extract_fields(raw_text: str) -> Dict[str, Optional[str]]:
    t = normalize_text(raw_text)
    subject_case = _search_one(r"件名:\s*【\s*([^】]+)\s*】", t, re.IGNORECASE)
    subject_manageno = _search_one(r"件名:.*?【[^】]+】\s*([A-Z0-9\-]+)", t, re.IGNORECASE)
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
    out = {k: None for k in [
        "案件種別(件名)","管理番号","物件名","住所","窓口会社","メーカー","制御方式","契約種別",
        "受信時刻","受信内容","通報者","現着時刻","現着状況","完了時刻","原因","処置内容",
        "対応者","送信者","受付番号","受付URL","現着完了登録URL"
    ]}
    out["案件種別(件名)"] = subject_case
    for k, pat in single_line.items():
        out[k] = _search_one(pat, t, re.IGNORECASE | re.MULTILINE)
    if not out["管理番号"] and subject_manageno:
        out["管理番号"] = subject_manageno
    for k in multiline_labels:
        out[k] = _search_span_between(multiline_labels, k, t)
    dur = minutes_between(out["現着時刻"], out["完了時刻"])
    out["作業時間_分"] = str(dur) if dur is not None and dur >= 0 else None
    return out

# ====== テンプレ書込み（チェックボックス保持版） ======
def fill_template_xlsx_safe(template_bytes: bytes, data: Dict[str, Optional[str]]) -> bytes:
    """I10〜P11のチェックボックス領域を保護したまま値を書き込む"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tmpl:
        tmp_tmpl.write(template_bytes)
        tmp_tmpl.flush()
        template_path = tmp_tmpl.name

    tmp_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp_output.close()
    shutil.copyfile(template_path, tmp_output.name)

    wb = load_workbook(tmp_output.name)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    # --- チェックボックス範囲 I10:P11 は書き換え禁止 ---
    protected_range = set([f"{col}{row}" for col in "IJKLMNOP" for row in (10, 11)])

    def safe_write(cell, value):
        if cell not in protected_range:
            ws[cell] = value

    # --- 単項目 ---
    safe_write("C12", data.get("管理番号") or "")
    safe_write("J12", data.get("メーカー") or "")
    safe_write("M12", data.get("制御方式") or "")
    safe_write("C15", data.get("受信内容") or "")
    safe_write("C14", data.get("通報者") or "")
    safe_write("L37", data.get("対応者") or "")

    pa = st.session_state.get("processing_after")
    if pa:
        safe_write("C35", pa)
    if data.get("所属"):
        safe_write("C37", data["所属"])

    now = datetime.now(JST)
    safe_write("B5", now.year)
    safe_write("D5", now.month)
    safe_write("F5", now.day)

    # --- 日付・時刻（分割入力） ---
    def write_dt_block(base_row: int, src_key: str):
        dt = _try_parse_datetime(data.get(src_key))
        y, m, d, wd, hh, mm = _split_dt_components(dt)
        for key, val in {
            f"C{base_row}": y, f"F{base_row}": m, f"H{base_row}": d,
            f"J{base_row}": wd, f"M{base_row}": f"{hh:02d}" if hh else None,
            f"O{base_row}": f"{mm:02d}" if mm else None
        }.items():
            if val is not None:
                safe_write(key, val)

    write_dt_block(13, "受信時刻")
    write_dt_block(19, "現着時刻")
    write_dt_block(36, "完了時刻")

    # --- 複数行 ---
    def fill_multiline(col: str, start: int, text: Optional[str]):
        lines = _split_lines(text)
        for i in range(5):
            safe_write(f"{col}{start + i}", "")
        for i, line in enumerate(lines[:5]):
            safe_write(f"{col}{start + i}", line)

    fill_multiline("C", 20, data.get("現着状況"))
    fill_multiline("C", 25, data.get("原因"))
    fill_multiline("C", 30, data.get("処置内容"))

    wb.save(tmp_output.name)
    with open(tmp_output.name, "rb") as f:
        return f.read()

# ====== ファイル名生成 ======
def build_filename(data: Dict[str, Optional[str]]) -> str:
    base_day = _first_date_yyyymmdd(data.get("現着時刻"), data.get("完了時刻"), data.get("受信時刻"))
    manageno = (data.get("管理番号") or "UNKNOWN").replace("/", "_")
    bname = (data.get("物件名") or "").strip().replace("/", "_")
    if bname:
        return f"緊急出動報告書_{manageno}_{bname}_{base_day}.xlsx"
    return f"緊急出動報告書_{manageno}_{base_day}.xlsx"

# ====== Streamlit UI ======
st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)
if "step" not in st.session_state: st.session_state.step = 1
if "authed" not in st.session_state: st.session_state.authed = False
if "extracted" not in st.session_state: st.session_state.extracted = None
if "template_xlsx_bytes" not in st.session_state: st.session_state.template_xlsx_bytes = None
if "affiliation" not in st.session_state: st.session_state.affiliation = ""

# Step 1: 認証
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

# Step 2: メール貼付 & テンプレ選択
elif st.session_state.step == 2 and st.session_state.authed:
    st.subheader("Step 2. メール本文の貼り付け / テンプレの指定 / 所属")
    tmpl = st.file_uploader("テンプレート（.xlsx）をアップロードしてください", type=["xlsx"])
    if tmpl:
        st.session_state.template_xlsx_bytes = tmpl.read()
        st.success("テンプレートを読み込みました。")
    aff = st.text_input("所属", value=st.session_state.affiliation)
    st.session_state.affiliation = aff
    processing_after = st.text_input("処理修理後（任意）")
    if processing_after: st.session_state["processing_after"] = processing_after
    text = st.text_area("故障完了メール本文", height=240)
    if st.button("抽出する", use_container_width=True):
        if not text.strip():
            st.warning("本文を入力してください。")
        elif not st.session_state.template_xlsx_bytes:
            st.warning("テンプレートを指定してください。")
        else:
            st.session_state.extracted = extract_fields(text)
            st.session_state.extracted["所属"] = st.session_state.affiliation
            st.session_state.step = 3
            st.rerun()

# Step 3: 抽出確認 & Excel生成
elif st.session_state.step == 3 and st.session_state.authed:
    st.subheader("Step 3. 抽出結果の確認 → Excel生成")
    data = st.session_state.extracted or {}
    if not data:
        st.warning("抽出データがありません。Step 2へ戻ってください。")
    else:
        with st.expander("基本情報", expanded=True):
            st.markdown(f"- 管理番号：{data.get('管理番号')}")
            st.markdown(f"- 物件名：{data.get('物件名')}")
        try:
            xlsx_bytes = fill_template_xlsx_safe(st.session_state.template_xlsx_bytes, data)
            fname = build_filename(data)
            st.download_button("Excelを生成（.xlsx）", data=xlsx_bytes,
                file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        except Exception as e:
            st.error(f"テンプレートへの書き込みでエラーが発生しました: {e}")

        if st.button("Step2に戻る", use_container_width=True):
            st.session_state.step = 2
            st.rerun()
        if st.button("最初に戻る", use_container_width=True):
            st.session_state.step = 1
            st.session_state.extracted = None
            st.session_state.template_xlsx_bytes = None
            st.session_state.affiliation = ""
            st.rerun()

else:
    st.warning("認証が必要です。Step 1に戻ります。")
    st.session_state.step = 1
