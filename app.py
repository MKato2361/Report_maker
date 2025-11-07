# app.py
# ------------------------------------------------------------
# 故障メール → 正規表現抽出 → 既存テンプレ(.xlsm)へ書込み → ダウンロード
# 3ステップUI / パスコード認証 / 編集不可 / 折りたたみ表示（時系列）
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
APP_TITLE = "故障報告メール → Excel自動生成（マクロ対応）"
PASSCODE_DEFAULT = "1357"
PASSCODE = st.secrets.get("APP_PASSCODE", PASSCODE_DEFAULT)
SHEET_NAME = "緊急出動報告書（リンク付き）"
WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]

# ========== ユーティリティ関数 ==========
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

def minutes_between(a: Optional[str], b: Optional[str]) -> Optional[int]:
    s = _try_parse_datetime(a)
    e = _try_parse_datetime(b)
    if s and e:
        return int((e - s).total_seconds() // 60)
    return None

# ========== 抽出処理 ==========
def extract_fields(raw_text: str) -> Dict[str, Optional[str]]:
    t = normalize_text(raw_text)
    subject_case = _search_one(r"件名:\s*【\s*([^】]+)\s*】", t)
    subject_manageno = _search_one(r"件名:.*?【[^】]+】\s*([A-Z0-9\-]+)", t)

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
        "受付URL": r"詳細はこちら\s*:\s*(https?://\S+)",
        "現着完了登録URL": r"現着・完了登録はこちら\s*:\s*(https?://\S+)",
    }

    multiline_labels = {
        "受信内容": r"受信内容\s*:",
        "現着状況": r"現着状況\s*:",
        "原因": r"原因\s*:",
        "処置内容": r"処置内容\s*:",
    }

    out = {k: None for k in ["案件種別(件名)", "管理番号", "物件名", "住所", "窓口会社", "メーカー", "制御方式",
                              "契約種別", "受信時刻", "受信内容", "通報者", "現着時刻", "現着状況", "完了時刻",
                              "原因", "処置内容", "対応者", "送信者", "受付番号", "受付URL", "現着完了登録URL"]}
    out["案件種別(件名)"] = subject_case

    for k, pat in single_line.items():
        out[k] = _search_one(pat, t, re.IGNORECASE)
    if not out["管理番号"] and subject_manageno:
        out["管理番号"] = subject_manageno
    for k in multiline_labels:
        out[k] = _search_span_between(multiline_labels, k, t)
    dur = minutes_between(out["現着時刻"], out["完了時刻"])
    out["作業時間_分"] = str(dur) if dur else None
    return out

# ========== Excel出力 ==========
def fill_template_xlsx(template_bytes: bytes, data: Dict[str, Optional[str]]) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    now = datetime.now(JST)
    ws["B5"], ws["D5"], ws["F5"] = now.year, now.month, now.day
    ws["C12"], ws["J12"], ws["M12"] = data.get("管理番号"), data.get("メーカー"), data.get("制御方式")
    ws["C14"], ws["L37"], ws["C37"] = data.get("通報者"), data.get("対応者"), data.get("所属")

    def write_multiline(ref, val):
        if not val:
            ws[ref] = ""
        else:
            ws[ref] = val.replace("\r\n", "\n")
            ws[ref].alignment = ws[ref].alignment.copy(wrapText=True)

    write_multiline("C15", data.get("受信内容"))
    write_multiline("C20", data.get("現着状況"))
    write_multiline("C25", data.get("原因"))
    write_multiline("C30", data.get("処置内容"))
    write_multiline("C35", data.get("処理修理後") or st.session_state.get("processing_after", ""))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ========== UI ==========
st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)
if "step" not in st.session_state:
    st.session_state.step = 1
if "authed" not in st.session_state:
    st.session_state.authed = False

# --- Step1: 認証 ---
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

# --- Step2: 本文入力 ---
elif st.session_state.step == 2 and st.session_state.authed:
    st.subheader("Step 2. 本文貼付 / 所属入力")

    template_path = "template.xlsm"
    if os.path.exists(template_path):
        with open(template_path, "rb") as f:
            st.session_state.template_xlsx_bytes = f.read()
        st.success(f"テンプレートを読み込みました: {template_path}")
    else:
        st.error("template.xlsm が見つかりません。")
        st.stop()

    aff = st.text_input("所属", value=st.session_state.get("affiliation", ""))
    st.session_state["affiliation"] = aff
    proc = st.text_input("処理修理後（任意）", value=st.session_state.get("processing_after", ""))
    st.session_state["processing_after"] = proc
    text = st.text_area("故障完了メール本文を貼り付け", height=240)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("抽出する", use_container_width=True):
            if not text.strip():
                st.warning("本文が空です。")
            else:
                st.session_state.extracted = extract_fields(text)
                st.session_state.extracted["所属"] = aff
                st.session_state.step = 3
                st.rerun()
    with col2:
        if st.button("クリア", use_container_width=True):
            st.session_state.extracted = None
            st.session_state.affiliation = ""

# --- Step3: 結果確認・編集 ---
elif st.session_state.step == 3 and st.session_state.authed:
    st.subheader("Step 3. 抽出結果の確認 → Excel生成")
    data = st.session_state.extracted or {}

    def _edit_key(key): return f"__edit_mode__{key}"

    def render_editable_row(label, key, multiline=False, help_text=""):
        if _edit_key(key) not in st.session_state:
            st.session_state[_edit_key(key)] = False
        cols = st.columns([6, 1])
        with cols[0]:
            val = data.get(key) or ""
            st.markdown(f"・**{label}**：{val}" if not multiline else f"・**{label}**：\n\n{val}")
        with cols[1]:
            if not st.session_state[_edit_key(key)]:
                st.button("✏️", key=f"btn_{key}", on_click=lambda: st.session_state.update({_edit_key(key): True}))
            else:
                st.button("❌", key=f"btn_c_{key}", on_click=lambda: st.session_state.update({_edit_key(key): False}))
        if st.session_state[_edit_key(key)]:
            new_val = st.text_area(f"{label}（編集）", value=val, height=120) if multiline else st.text_input(f"{label}（編集）", value=val)
            c1, c2 = st.columns(2)
            with c1:
                def _save(): data[key] = new_val; st.session_state[_edit_key(key)] = False
                st.button("✅ 保存", key=f"save_{key}", on_click=_save, use_container_width=True)
            with c2:
                st.button("キャンセル", key=f"cancel_{key}", on_click=lambda: st.session_state.update({_edit_key(key): False}), use_container_width=True)
            st.markdown("---")
    
    
    
    with st.expander("通報・受付情報", expanded=True):
        st.markdown(f"- 受信時刻：{data.get('受信時刻') or ''}")
        render_editable_row("通報者", "通報者")
        render_editable_row("受信内容", "受信内容", multiline=True)
    with st.expander("現着・作業・完了情報", expanded=True):
        render_editable_row("現着状況", "現着状況", multiline=True)
        render_editable_row("原因", "原因", multiline=True)
        render_editable_row("処置内容", "処置内容", multiline=True)
        render_editable_row("処理修理後", "処理修理後", multiline=True)
        st.markdown(f"- 現着時刻：{data.get('現着時刻') or ''}")
        st.markdown(f"- 完了時刻：{data.get('完了時刻') or ''}")
        render_editable_row("現着状況", "現着状況", multiline=True)
        dur = data.get("作業時間_分")
        if dur:
            st.info(f"作業時間（概算）：{dur} 分")

    with st.expander("基本情報", expanded=True):
        st.markdown(f"- 管理番号：{data.get('管理番号') or ''}")
        st.markdown(f"- 物件名：{data.get('物件名') or ''}")
        st.markdown(f"- 住所：{data.get('住所') or ''}")
        st.markdown(f"- 窓口会社：{data.get('窓口会社') or ''}")
    with st.expander("技術情報", expanded=False):
        render_editable_row("原因", "原因", multiline=True)
        render_editable_row("処置内容", "処置内容", multiline=True)
        st.markdown(f"- 制御方式：{data.get('制御方式') or ''}")
        st.markdown(f"- 契約種別：{data.get('契約種別') or ''}")
        st.markdown(f"- メーカー：{data.get('メーカー') or ''}")


    with st.expander("その他", expanded=False):
        st.markdown(f"- 所属：{data.get('所属') or ''}")
        st.markdown(f"- 対応者：{data.get('対応者') or ''}")
        st.markdown(f"- 送信者：{data.get('送信者') or ''}")
        st.markdown(f"- 受付番号：{data.get('受付番号') or ''}")
        st.markdown(f"- 受付URL：{data.get('受付URL') or ''}")
        st.markdown(f"- 現着・完了登録URL：{data.get('現着完了登録URL') or ''}")
        st.markdown(f"- 案件種別(件名)：{data.get('案件種別(件名)') or ''}")

    st.divider()
    try:
        out_bytes = fill_template_xlsx(st.session_state.template_xlsx_bytes, data)
        fname = f"{data.get('管理番号') or 'UNKNOWN'}_{datetime.now().strftime('%Y%m%d')}.xlsm"
        st.download_button("Excelを生成（.xlsm）", data=out_bytes, file_name=fname, mime="application/vnd.ms-excel.sheet.macroEnabled.12", use_container_width=True)
    except Exception as e:
        st.error(f"テンプレート書き込み中にエラー: {e}")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Step2に戻る", use_container_width=True):
            st.session_state.step = 2; st.rerun()
    with col2:
        if st.button("最初に戻る", use_container_width=True):
            st.session_state.step = 1; st.session_state.extracted = None; st.rerun()

else:
    st.warning("認証が必要です。")
    st.session_state.step = 1
