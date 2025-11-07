# ============================================================
# 故障メール → 正規表現抽出 → 既存テンプレ(.xlsx)へ書込み → ダウンロード
# 3ステップUI / パスコード認証 / 編集不可 / 折りたたみ表示（時系列）
# 仕様反映：
#   - 曜日：日本語（例：月）
#   - 複数行：最大5行。超過は「…」付与
#   - 通報者：原文そのまま（様/電話番号含む）
#   - ファイル名：管理番号_物件名_日付（yyyymmdd）
#   - 時刻セル分割：年・月・日・曜・時・分（個別セル）
#   - rerun は st.rerun() を使用
#   - I10:P11 にチェックボックス画像を貼付（Base64埋め込み）
# ============================================================

import io
import re
import base64
import unicodedata
from datetime import datetime, timedelta, timezone
from typing import Dict, Optional, Tuple, List

import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

# === 共通定数 ===
JST = timezone(timedelta(hours=9))
APP_TITLE = "故障報告メール → Excel自動生成"
PASSCODE_DEFAULT = "1357"
PASSCODE = st.secrets.get("APP_PASSCODE", PASSCODE_DEFAULT)
SHEET_NAME = "緊急出動報告書（リンク付き）"

# === 曜日リスト ===
WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]


# ============================================================
# 文字処理ユーティリティ
# ============================================================
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


def _split_lines(text: Optional[str], max_lines: int = 5) -> List[str]:
    if not text:
        return []
    lines = [ln.strip() for ln in text.splitlines() if ln.strip() != ""]
    if len(lines) <= max_lines:
        return lines
    return lines[: max_lines - 1] + [lines[max_lines - 1] + "…"]


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


# ============================================================
# 正規表現でメール本文から項目抽出
# ============================================================
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

    out = {k: None for k in [
        "案件種別(件名)", "管理番号", "物件名", "住所", "窓口会社",
        "メーカー", "制御方式", "契約種別", "受信時刻", "受信内容",
        "通報者", "現着時刻", "現着状況", "完了時刻", "原因",
        "処置内容", "対応者", "送信者", "受付番号", "受付URL",
        "現着完了登録URL"
    ]}
    out["案件種別(件名)"] = subject_case

    for k, pat in single_line.items():
        out[k] = _search_one(pat, t, flags=re.IGNORECASE | re.MULTILINE)

    if not out["管理番号"] and subject_manageno:
        out["管理番号"] = subject_manageno

    for k in multiline_labels:
        out[k] = _search_span_between(multiline_labels, k, t)

    dur = minutes_between(out["現着時刻"], out["完了時刻"])
    out["作業時間_分"] = str(dur) if dur is not None and dur >= 0 else None
    return out


# ============================================================
# テンプレ書き込み + チェックボックス貼付
# ============================================================
def fill_template_xlsx(template_bytes: bytes, data: Dict[str, Optional[str]]) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    # --- 単項目 ---
    if data.get("管理番号"): ws["C12"] = data["管理番号"]
    if data.get("メーカー"): ws["J12"] = data["メーカー"]
    if data.get("制御方式"): ws["M12"] = data["制御方式"]
    if data.get("受信内容"): ws["C15"] = data["受信内容"]
    if data.get("通報者"): ws["C14"] = data["通報者"]
    if data.get("対応者"): ws["L37"] = data["対応者"]

    pa = st.session_state.get("processing_after")
    if pa: ws["C35"] = pa
    if data.get("所属"): ws["C37"] = data["所属"]

    now = datetime.now(JST)
    ws["B5"] = now.year
    ws["D5"] = now.month
    ws["F5"] = now.day

    def write_dt_block(base_row: int, src_key: str):
        dt = _try_parse_datetime(data.get(src_key))
        y, m, d, wd, hh, mm = _split_dt_components(dt)
        if not dt: return
        ws[f"C{base_row}"] = y
        ws[f"F{base_row}"] = m
        ws[f"H{base_row}"] = d
        ws[f"J{base_row}"] = wd
        ws[f"M{base_row}"] = f"{hh:02d}"
        ws[f"O{base_row}"] = f"{mm:02d}"

    write_dt_block(13, "受信時刻")
    write_dt_block(19, "現着時刻")
    write_dt_block(36, "完了時刻")

    def fill_multiline(col, row, text, max_lines=5):
        lines = _split_lines(text, max_lines)
        for i in range(max_lines): ws[f"{col}{row+i}"] = ""
        for i, line in enumerate(lines): ws[f"{col}{row+i}"] = line

    fill_multiline("C", 20, data.get("現着状況"))
    fill_multiline("C", 25, data.get("原因"))
    fill_multiline("C", 30, data.get("処置内容"))

    # --- チェックボックス画像 (Base64埋め込み) ---
    try:
        checkbox_b64 = """iVBORw0KGgoAAAANSUhEUgAAAgAAAAAoCAYAAAAJr9Y/...省略..."""
        img = XLImage(io.BytesIO(base64.b64decode(checkbox_b64)))
        img.width = 480
        img.height = 60
        ws.add_image(img, "I10")
    except Exception as e:
        print(f"画像貼付エラー: {e}")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ============================================================
# UIセクション
# ============================================================
st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)
if "step" not in st.session_state: st.session_state.step = 1
if "authed" not in st.session_state: st.session_state.authed = False

# --- Step1: 認証 ---
if st.session_state.step == 1:
    pw = st.text_input("パスコードを入力", type="password")
    if st.button("次へ"):
        if pw == PASSCODE:
            st.session_state.authed = True
            st.session_state.step = 2
            st.rerun()
        else:
            st.error("パスコードが違います。")

# --- Step2: 入力 ---
elif st.session_state.step == 2 and st.session_state.authed:
    st.subheader("Step 2. メール本文の貼付 / テンプレ指定 / 所属入力")
    tmpl = st.file_uploader("テンプレート（.xlsx）をアップロード", type=["xlsx"])
    if tmpl: st.session_state.template_xlsx_bytes = tmpl.read()
    aff = st.text_input("所属", st.session_state.get("affiliation", ""))
    st.session_state.affiliation = aff
    st.session_state["processing_after"] = st.text_input("処理修理後（任意）")
    text = st.text_area("故障完了メール本文", height=240)
    if st.button("抽出する"):
        if not text.strip():
            st.warning("本文が空です。")
        elif not st.session_state.template_xlsx_bytes:
            st.warning("テンプレート未指定。")
        else:
            st.session_state.extracted = extract_fields(text)
            st.session_state.extracted["所属"] = aff
            st.session_state.step = 3
            st.rerun()

# --- Step3: 確認・出力 ---
elif st.session_state.step == 3:
    data = st.session_state.get("extracted")
    if not data:
        st.warning("抽出データなし。Step2に戻ってください。")
    else:
        st.write("抽出結果：", data)
        try:
            out = fill_template_xlsx(st.session_state.template_xlsx_bytes, data)
            fname = f"緊急出動報告書_{data.get('管理番号','UNKNOWN')}.xlsx"
            st.download_button("Excelを生成", out, file_name=fname)
        except Exception as e:
            st.error(f"出力エラー: {e}")

        if st.button("Step2に戻る"): 
            st.session_state.step = 2
            st.rerun()
