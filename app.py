# app.py
# ------------------------------------------------------------
# 故障メール → 正規表現抽出 → 既存テンプレ(.xlsm)へ書込み → ダウンロード
# 3ステップUI / パスコード認証 / 編集不可 / 折りたたみ表示（時系列）
# 仕様反映：
#   - 曜日：日本語（例：月）
#   - 複数行：最大5行。超過は「…」付与
#   - 通報者：原文そのまま（様/電話番号含む）
#   - ファイル名：管理番号_物件名_日付（yyyymmdd）
#   - マクロ保持対応（keep_vba=True）
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
def fill_template_xlsx(template_bytes: bytes, data: Dict[str, Optional[str]]) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    def format_lines(text: Optional[str], max_lines: int) -> str:
        """指定行数に合わせて空行を補完"""
        if not text:
            return "\n" * (max_lines - 1)
        lines = [ln.strip() for ln in text.splitlines()]
        while len(lines) < max_lines:
            lines.append("")
        return "\n".join(lines[:max_lines])

    def write_multiline(cell_ref: str, text: Optional[str], max_lines: int):
        s = format_lines(text, max_lines)
        ws[cell_ref] = s
        try:
            ws[cell_ref].alignment = ws[cell_ref].alignment.copy(wrapText=True)
        except Exception:
            pass

    # --- 通報者／受信内容／現着状況／原因／処置内容／処理修理後 ---
    ws["C14"] = format_lines(data.get("通報者"), 1)
    write_multiline("C15", data.get("受信内容"), 4)
    write_multiline("C20", data.get("現着状況"), 5)
    write_multiline("C25", data.get("原因"), 5)
    write_multiline("C30", data.get("処置内容"), 5)

    proc_after = data.get("処理修理後") or st.session_state.get("processing_after")
    write_multiline("C35", proc_after, 1)

    now = datetime.now(JST)
    ws["B5"], ws["D5"], ws["F5"] = now.year, now.month, now.day
    ws["C12"], ws["J12"], ws["M12"] = data.get("管理番号"), data.get("メーカー"), data.get("制御方式")
    ws["L37"], ws["C37"] = data.get("対応者"), data.get("所属")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def build_filename(data: Dict[str, Optional[str]]) -> str:
    base_day = _first_date_yyyymmdd(data.get("現着時刻"), data.get("完了時刻"), data.get("受信時刻"))
    manageno = (data.get("管理番号") or "UNKNOWN").replace("/", "_")
    bname = (data.get("物件名") or "").strip().replace("/", "_")
    if bname:
        return f"緊急出動報告書_{manageno}_{bname}_{base_day}.xlsm"
    return f"緊急出動報告書_{manageno}_{base_day}.xlsm"

# ====== UI（Step1〜2 は元のまま） ======
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

# Step 2: 本文貼付・テンプレ自動読み込み（元のまま）
elif st.session_state.step == 2 and st.session_state.authed:
    st.subheader("Step 2. メール本文の貼り付け / 所属")

    template_path = "template.xlsm"
    if os.path.exists(template_path):
        with open(template_path, "rb") as f:
            st.session_state.template_xlsx_bytes = f.read()
        st.success(f"テンプレートを読み込みました: {template_path}")
    else:
        st.error(f"テンプレートファイルが見つかりません: {template_path}")
        st.stop()

    aff = st.text_input("所属（例：札幌支店 / 本社 / 道央サービスなど）", value=st.session_state.affiliation)
    st.session_state.affiliation = aff
    processing_after = st.text_input("処理修理後（任意）")
    if processing_after:
        st.session_state["processing_after"] = processing_after

    text = st.text_area("故障完了メール（本文）を貼り付け", height=240)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("抽出する", use_container_width=True):
            if not text.strip():
                st.warning("本文が空です。")
            else:
                st.session_state.extracted = extract_fields(text)
                st.session_state.extracted["所属"] = st.session_state.affiliation
                st.session_state.step = 3
                st.rerun()
    with c2:
        if st.button("クリア", use_container_width=True):
            st.session_state.extracted = None
            st.session_state.affiliation = ""
# ===== Step 3：表示は元のまま＋右側に ✏️ 編集アイコンで項目別編集 =====
elif st.session_state.step == 3 and st.session_state.authed:
    st.subheader("Step 3. 抽出結果の確認 → Excel生成")

    data = st.session_state.extracted or {}
    if not data:
        st.warning("抽出データがありません。")
    else:
        # --- 編集モード管理（各項目ごとにトグル） ---
        def _edit_key(key: str) -> str:
            return f"__edit_mode__{key}"
            
        def render_editable_row(label: str, key: str, multiline: bool = False, help_text: str = ""):
    """表示→編集切替（・形式、編集時は置き換え）"""
    　　　　if _edit_key(key) not in st.session_state:
                st.session_state[_edit_key(key)] = False

            if not st.session_state[_edit_key(key)]:
                cols = st.columns([6, 1])
        with cols[0]:
            val = data.get(key) or ""
            if multiline:
                st.markdown(f"・**{label}**：\n\n{val}")
            else:
                st.markdown(f"・**{label}**：{val}")
        with cols[1]:
            st.button("✏️", key=f"btn_edit_{key}", help=f"{label}を編集", on_click=lambda: st.session_state.update({_edit_key(key): True}))
    else:
        if multiline:
            new_val = st.text_area(f"{label}（編集）", value=data.get(key) or "", height=120, help=help_text, key=f"input_{key}")
        else:
            new_val = st.text_input(f"{label}（編集）", value=data.get(key) or "", help=help_text, key=f"input_{key}")

        c1, c2 = st.columns(2)
        with c1:
            def _save():
                data[key] = new_val
                st.session_state[_edit_key(key)] = False
            st.button("✅ 保存", key=f"save_{key}", on_click=_save, use_container_width=True)
        with c2:
            st.button("キャンセル", key=f"cancel_{key}", on_click=lambda: st.session_state.update({_edit_key(key): False}), use_container_width=True)
        st.markdown("---")


        # ====== 元の「表示UI」はそのまま維持 ======
        with st.expander("基本情報", expanded=True):
            st.markdown(f"- 管理番号：{data.get('管理番号') or ''}")
            st.markdown(f"- 物件名：{data.get('物件名') or ''}")
            st.markdown(f"- 住所：{data.get('住所') or ''}")
            st.markdown(f"- 窓口会社：{data.get('窓口会社') or ''}")

        with st.expander("通報・受付情報", expanded=True):
            st.markdown(f"- 受信時刻：{data.get('受信時刻') or ''}")

            # ← ここに編集行を追加（表示は残す）
            render_editable_row("通報者", "通報者", multiline=False)
            render_editable_row("受信内容", "受信内容", multiline=True)

        with st.expander("現着・作業・完了情報", expanded=True):
            st.markdown(f"- 現着時刻：{data.get('現着時刻') or ''}")
            st.markdown(f"- 完了時刻：{data.get('完了時刻') or ''}")
            render_editable_row("現着状況", "現着状況", multiline=True)
            dur = data.get("作業時間_分")
            if dur:
                st.info(f"作業時間（概算）：{dur} 分")



        with st.expander("技術情報", expanded=False):
            render_editable_row("原因", "原因", multiline=True)
            render_editable_row("処置内容", "処置内容", multiline=True)
            st.markdown(f"- 制御方式：{data.get('制御方式') or ''}")
            st.markdown(f"- 契約種別：{data.get('契約種別') or ''}")
            st.markdown(f"- メーカー：{data.get('メーカー') or ''}")


        with st.expander("その他", expanded=False):
            st.markdown(f"- 所属：{data.get('所属') or ''}")
            st.markdown(f"- 対応者：{data.get('対応者') or ''}")
            render_editable_row("処理修理後", "処理修理後", multiline=True, help_text="未入力の場合はStep2の値（処理修理後）を出力時に使用します。")
            st.markdown(f"- 送信者：{data.get('送信者') or ''}")
            st.markdown(f"- 受付番号：{data.get('受付番号') or ''}")
            st.markdown(f"- 受付URL：{data.get('受付URL') or ''}")
            st.markdown(f"- 現着・完了登録URL：{data.get('現着完了登録URL') or ''}")
            st.markdown(f"- 案件種別(件名)：{data.get('案件種別(件名)') or ''}")



        st.divider()

        # 生成・ダウンロード（元のまま）
        try:
            xlsx_bytes = fill_template_xlsx(st.session_state.template_xlsx_bytes, data)
            fname = build_filename(data)
            st.download_button(
                "Excelを生成（.xlsm）",
                data=xlsx_bytes,
                file_name=fname,
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"テンプレート書き込み中にエラー: {e}")

        # ▼ Step2に戻る／最初に戻る（元のまま）
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
