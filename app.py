# ============================================================
# 故障メール → 正規表現抽出 → 既存テンプレ(.xlsx)へ書込み → ダウンロード
# 3ステップUI / パスコード認証 / 編集不可 / 折りたたみ表示（時系列）
# 仕様反映：
#   - 曜日：日本語（例：月）
#   - 複数行：最大5行。超過は「…」付与
#   - 通報者：原文そのまま（様/電話番号含む）
#   - ファイル名：管理番号_物件名_日付（yyyymmdd）
#   - 時刻セル分割：年・月・日・曜・時・分（個別セル）
#   - rerun は# app.py (Part 1/2)
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
        # --- 処理修理後の書込み（C35） ---
    pa = st.session_state.get("processing_after")
    if pa:
        ws["C35"] = pa
    if data.get("所属"): ws["C37"] = data["所属"]  # ★所属 追加（C37）
        # --- 現在日付をB5/D5/F5へ書き込み（JST） ---
    now = datetime.now(JST)
    ws["B5"] = now.year
    ws["D5"] = now.month
    ws["F5"] = now.day


    # --- 日付・時刻（分割入力） ---
    def write_dt_block(base_row: int, src_key: str):
        """
        base_row: 13(受信), 19(現着), 36(完了)
        C(年) F(月) H(日) J(曜) M(時) O(分)
        """
        dt = _try_parse_datetime(data.get(src_key))
        y, m, d, wd, hh, mm = _split_dt_components(dt)
        cellmap = {
            "Y": f"C{base_row}",
            "Mo": f"F{base_row}",
            "D": f"H{base_row}",
            "W": f"J{base_row}",
            "H": f"M{base_row}",
            "Min": f"O{base_row}",
        }
        if y is not None: ws[cellmap["Y"]] = y
        if m is not None: ws[cellmap["Mo"]] = m
        if d is not None: ws[cellmap["D"]] = d
        if wd is not None: ws[cellmap["W"]] = wd
        if hh is not None: ws[cellmap["H"]] = f"{hh:02d}"
        if mm is not None: ws[cellmap["Min"]] = f"{mm:02d}"

    write_dt_block(13, "受信時刻")   # 通報時刻（受信）
    write_dt_block(19, "現着時刻")   # 現着
    write_dt_block(36, "完了時刻")   # 完了

    # --- 複数行（最大5行、超過は「…」） ---
    def fill_multiline(col_letter: str, start_row: int, text: Optional[str], max_lines: int = 5):
        lines = _split_lines(text, max_lines=max_lines)
        for i in range(max_lines):  # 先にクリア
            ws[f"{col_letter}{start_row + i}"] = ""
        for idx, line in enumerate(lines[:max_lines]):
            ws[f"{col_letter}{start_row + idx}"] = line

    fill_multiline("C", 20, data.get("現着状況"), max_lines=5)  # C20~C24
    fill_multiline("C", 25, data.get("原因"), max_lines=5)      # C25~C29
    fill_multiline("C", 30, data.get("処置内容"), max_lines=5)  # C30~C34

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def build_filename(data: Dict[str, Optional[str]]) -> str:
    base_day = _first_date_yyyymmdd(data.get("現着時刻"), data.get("完了時刻"), data.get("受信時刻"))
    manageno = (data.get("管理番号") or "UNKNOWN").replace("/", "_")
    bname = (data.get("物件名") or "").strip().replace("/", "_")
    if bname:
        return f"緊急出動報告書_{manageno}_{bname}_{base_day}.xlsx"
    return f"緊急出動報告書_{manageno}_{base_day}.xlsx"

# ====== UI ======
st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)

if "step" not in st.session_state:
    st.session_state.step = 1
if "authed" not in st.session_state:
    st.session_state.authed = False
if "extracted" not in st.session_state:
    st.session_state.extracted = None
if "template_xlsx_bytes" not in st.session_state:
    st.session_state.template_xlsx_bytes = None
if "affiliation" not in st.session_state:      # ★所属 初期化
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

# Step 2: 本文貼付 + テンプレアップロード + 所属（NEW）
elif st.session_state.step == 2 and st.session_state.authed:
    st.subheader("Step 2. メール本文の貼り付け / テンプレの指定 / 所属")

    tmpl = st.file_uploader(
        "テンプレート（.xlsx）をアップロードしてください（※ .xlsb をExcelで一度 .xlsx 保存したもの）",
        type=["xlsx"],
        accept_multiple_files=False,
    )
    if tmpl is not None:
        st.session_state.template_xlsx_bytes = tmpl.read()
        st.success("テンプレートを読み込みました。")

    # ★ 所属入力欄
    aff = st.text_input("所属", value=st.session_state.affiliation)
    st.session_state.affiliation = aff
        # ▼ 処理修理後（1行入力）
    processing_after = st.text_input("処理修理後（任意）")
    if processing_after:
        st.session_state["processing_after"] = processing_after


    text = st.text_area(
        "故障完了メール（本文）を貼り付け",
        height=240,
        placeholder="例：\n件名: 【故障完了】 HK1-234・ABCビル\n管理番号：HK1-234\n物件名：ABCビル\n住所：北海道札幌市中央区\n…"
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("抽出する", use_container_width=True):
            if not text.strip():
                st.warning("本文が空です。貼り付けてから『抽出する』を押してください。")
            elif not st.session_state.template_xlsx_bytes:
                st.warning("テンプレート（.xlsx）が未指定です。まずはアップロードしてください。")
            else:
                st.session_state.extracted = extract_fields(text)
                # ★ Step3 へ渡す前に所属もデータとして保持する
                st.session_state.extracted["所属"] = st.session_state.affiliation
                st.session_state.step = 3
                st.rerun()

    with c2:
        if st.button("クリア", use_container_width=True):
            st.session_state.extracted = None
            st.session_state.template_xlsx_bytes = None
            st.session_state.affiliation = ""

# Step 3: 抽出確認 → 生成
elif st.session_state.step == 3 and st.session_state.authed:
    st.subheader("Step 3. 抽出結果の確認 → Excel生成")

    data = st.session_state.extracted or {}
    if not data:
        st.warning("抽出データがありません。Step 2に戻って本文を貼り付けてください。")
    else:
        with st.expander("基本情報", expanded=True):
            st.markdown(f"- 管理番号：{data.get('管理番号') or ''}")
            st.markdown(f"- 物件名：{data.get('物件名') or ''}")
            st.markdown(f"- 住所：{data.get('住所') or ''}")
            st.markdown(f"- 窓口会社：{data.get('窓口会社') or ''}")

        with st.expander("通報・受付情報", expanded=True):
            st.markdown(f"- 受信時刻：{data.get('受信時刻') or ''}")
            st.markdown(f"- 通報者：{data.get('通報者') or ''}")
            st.markdown(f"- 通報内容：\n\n{data.get('受信内容') or ''}")

        with st.expander("現着・作業・完了情報", expanded=True):
            st.markdown(f"- 現着時刻：{data.get('現着時刻') or ''}")
            st.markdown(f"- 完了時刻：{data.get('完了時刻') or ''}")
            st.markdown(f"- 現着状況：\n\n{data.get('現着状況') or ''}")
            dur = data.get("作業時間_分")
            if dur:
                st.info(f"作業時間（概算）：{dur} 分")

        with st.expander("技術情報", expanded=False):
            st.markdown(f"- 原因：\n\n{data.get('原因') or ''}")
            st.markdown(f"- 処置内容：\n\n{data.get('処置内容') or ''}")
            st.markdown(f"- 制御方式：{data.get('制御方式') or ''}")
            st.markdown(f"- 契約種別：{data.get('契約種別') or ''}")
            st.markdown(f"- メーカー：{data.get('メーカー') or ''}")

        with st.expander("その他", expanded=False):
            st.markdown(f"- 所属：{data.get('所属') or ''}")  # ★所属 表示
            st.markdown(f"- 対応者：{data.get('対応者') or ''}")
            st.markdown(f"- 処理修理後：{st.session_state.get('processing_after', '')}")
            st.markdown(f"- 送信者：{data.get('送信者') or ''}")
            st.markdown(f"- 受付番号：{data.get('受付番号') or ''}")
            st.markdown(f"- 受付URL：{data.get('受付URL') or ''}")
            st.markdown(f"- 現着・完了登録URL：{data.get('現着完了登録URL') or ''}")
            st.markdown(f"- 案件種別(件名)：{data.get('案件種別(件名)') or ''}")

        st.divider()

        try:
            xlsx_bytes = fill_template_xlsx(st.session_state.template_xlsx_bytes, data)
            fname = build_filename(data)
            st.download_button(
                "Excelを生成（.xlsx）",
                data=xlsx_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"テンプレートへの書き込みでエラーが発生しました: {e}")

                # ▼ Step2に戻るボタン
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
    st.session_state.step = 1 st.rerun() を使用
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
