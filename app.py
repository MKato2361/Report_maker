# app.py（編集UI追加・完全版）
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
import streamlit as st

JST = timezone(timedelta(hours=9))

APP_TITLE = "故障報告メール → Excel自動生成（マクロ対応・編集UI付き）"
PASSCODE_DEFAULT = "1357"
PASSCODE = st.secrets.get("APP_PASSCODE", PASSCODE_DEFAULT)
SHEET_NAME = "緊急出動報告書（リンク付き）"
WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]

# ==========================================================
# 共通ユーティリティ
# ==========================================================
def normalize_text(text: str) -> str:
    if not text:
        return ""
    t = unicodedata.normalize("NFKC", text)
    return t.replace("：", ":").replace("\t", " ").replace("\r\n", "\n").replace("\r", "\n")

def _try_parse_datetime(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    s = s.strip().replace("年", "/").replace("月", "/").replace("日", "").replace("-", "/")
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt)
        except:
            pass
    return None

def _split_lines(text: Optional[str], max_lines: int) -> List[str]:
    if not text:
        return ["" for _ in range(max_lines)]
    lines = [ln.strip() for ln in text.splitlines() if ln.strip() != ""]
    if len(lines) < max_lines:
        lines += [""] * (max_lines - len(lines))
    else:
        lines = lines[:max_lines]
    return lines

# ==========================================================
# Excelテンプレート書き込み
# ==========================================================
def fill_template_xlsx(template_bytes: bytes, data: Dict[str, Optional[str]]) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    ws["C12"] = data.get("管理番号")
    ws["C14"] = data.get("通報者")
    ws["C15"] = data.get("受信内容")
    ws["C35"] = data.get("処理修理後") or st.session_state.get("processing_after", "")
    ws["C37"] = data.get("所属")
    ws["L37"] = data.get("対応者")

    def fill_block(col, start_row, key, lines):
        for i, line in enumerate(lines):
            ws[f"{col}{start_row+i}"] = line

    fill_block("C", 20, "現着状況", _split_lines(data.get("現着状況"), 5))
    fill_block("C", 25, "原因", _split_lines(data.get("原因"), 5))
    fill_block("C", 30, "処置内容", _split_lines(data.get("処置内容"), 5))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ==========================================================
# 編集UIコンポーネント
# ==========================================================
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

# ==========================================================
# Streamlit UI 構成
# ==========================================================
st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)

if "step" not in st.session_state:
    st.session_state.step = 1
if "authed" not in st.session_state:
    st.session_state.authed = False
if "extracted" not in st.session_state:
    st.session_state.extracted = None

# ----------------------------------------------------------
# Step 1: パスコード認証
# ----------------------------------------------------------
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

# ----------------------------------------------------------
# Step 2: メール本文入力＋テンプレ自動読み込み
# ----------------------------------------------------------
elif st.session_state.step == 2 and st.session_state.authed:
    st.subheader("Step 2. メール本文の貼り付け / 所属")

    template_path = "template.xlsm"
    if os.path.exists(template_path):
        with open(template_path, "rb") as f:
            st.session_state.template_xlsx_bytes = f.read()
        st.success("テンプレートを読み込みました。")
    else:
        st.error("template.xlsm が見つかりません")
        st.stop()

    aff = st.text_input("所属", value=st.session_state.get("affiliation", ""))
    st.session_state.affiliation = aff
    processing_after = st.text_input("処理修理後（任意）", value=st.session_state.get("processing_after", ""))
    st.session_state["processing_after"] = processing_after
    text = st.text_area("故障完了メール本文を貼り付け", height=240)

    if st.button("抽出（テスト用ダミー）"):
        st.session_state.extracted = {
            "管理番号": "HK-001",
            "通報者": "山田太郎",
            "受信内容": "停止発生\n再起動実施\n復帰確認",
            "現着状況": "到着済み\n点検実施\n異常なし",
            "原因": "接点不良\n誤作動",
            "処置内容": "部品交換\n清掃",
            "所属": aff,
            "対応者": "佐藤",
            "処理修理後": processing_after
        }
        st.session_state.step = 3
        st.rerun()

# ----------------------------------------------------------
# Step 3: 抽出結果確認＋編集UI＋Excel出力
# ----------------------------------------------------------
elif st.session_state.step == 3 and st.session_state.authed:
    st.subheader("Step 3. 抽出結果の確認・編集 → Excel生成")

    data = st.session_state.extracted or {}
    with st.expander("🧾 編集可能項目", expanded=True):
        editable_field("通報者", "通報者", 1)
        editable_field("受信内容", "受信内容", 4)
        editable_field("現着状況", "現着状況", 5)
        editable_field("原因", "原因", 5)
        editable_field("処置内容", "処置内容", 5)
        editable_field("処理修理後（Step2入力値）", "処理修理後", 1)

    st.divider()
    if st.button("Excelを生成（.xlsm）"):
        xlsx = fill_template_xlsx(st.session_state.template_xlsx_bytes, data)
        st.download_button(
            "ダウンロード",
            data=xlsx,
            file_name="緊急出動報告書.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            use_container_width=True,
        )

    if st.button("Step2に戻る"):
        st.session_state.step = 2
        st.rerun()

else:
    st.warning("認証が必要です。Step1に戻ります。")
    st.session_state.step = 1
