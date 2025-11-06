# app.py
# ------------------------------------------------------------
# 故障メール → 正規表現抽出 → 既存テンプレ(.xlsx)へ書込み → ダウンロード
# 3ステップUI / パスコード認証 / 編集不可 / 折りたたみ表示（時系列）
# 仕様はユーザー確定内容を反映
#   - 曜日：日本語（例：月）
#   - 複数行：最大5行。超過は「…」付与
#   - 通報者：原文そのまま（様/電話番号含む）
#   - ファイル名：管理番号_物件名_日付（yyyymmdd）
#   - 時刻セル分割：年・月・日・曜・時・分（個別セル）
# ------------------------------------------------------------
import io
import re
import unicodedata
from datetime import datetime
from typing import Dict, Optional, Tuple, List

import streamlit as st
from openpyxl import load_workbook

APP_TITLE = "故障報告メール → Excel自動生成"
PASSCODE_DEFAULT = "1357"  # st.secrets["APP_PASSCODE"] で上書き推奨
PASSCODE = st.secrets.get("APP_PASSCODE", PASSCODE_DEFAULT)

SHEET_NAME = "緊急出動報告書（リンク付き）"  # テンプレのシート名（RTFに準拠）

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
    wd = WEEKDAYS_JA[(dt.weekday() + 1) % 7]  # Python: Mon=0。日本の並びに合わせてもOKだが/月火水…に対応
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
    # 超過時は末尾に「…」
    kept = lines[: max_lines - 1] + [lines[max_lines - 1] + "…"]
    return kept

# ====== 正規表現 抽出（サンプル仕様に準拠） ======
def extract_fields(raw_text: str) -> Dict[str, Optional[str]]:
    t = normalize_text(raw_text)

    # 件名（冗長抽出）
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
        "通報者": r"通報者\s*:\s*(.+)",  # 原文そのまま（様/電話番号含む）
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

    # 参考：作業時間（分）
    dur = minutes_between(out["現着時刻"], out["完了時刻"])
    out["作業時間_分"] = str(dur) if dur is not None and dur >= 0 else None
    return out

# ====== テンプレ書き込み ======
def fill_template_xlsx(template_bytes: bytes, data: Dict[str, Optional[str]]) -> bytes:
    """
    .xlsx テンプレ（.xlsbを人手で保存し直したもの）に対し、
    指定セルへ値を書き込んで新規 .xlsx を返す
    """
    wb = load_workbook(io.BytesIO(template_bytes))
    if SHEET_NAME not in wb.sheetnames:
        # シート名違いを許容するため、最初のシートを使う（保険）
        ws = wb.active
    else:
        ws = wb[SHEET_NAME]

    # --- 簡単項目 ---
    # C12: 管理番号
    if data.get("管理番号"):
        ws["C12"] = data["管理番号"]

    # 物件名（RTFではセル未指定。通常は書くが、未指定のためスキップ）
    # 住所（未指定のためスキップ）
    # 窓口会社（未指定のためスキップ）

    # J12: メーカー
    if data.get("メーカー"):
        ws["J12"] = data["メーカー"]

    # M12: 制御方式
    if data.get("制御方式"):
        ws["M12"] = data["制御方式"]

    # 契約種別（RTFではセル未指定。必要なら後で割付）

    # C15: 通報内容
    if data.get("受信内容"):
        ws["C15"] = data["受信内容"]

    # C14: 通報者（原文そのまま）
    if data.get("通報者"):
        ws["C14"] = data["通報者"]

    # L37: 対応者（作業者） フルネーム
    if data.get("対応者"):
        ws["L37"] = data["対応者"]

    # 受付担当者（送信者）・URLなど（RTFにセル指定なし→必要なら追記可能）

    # --- 日付・時刻（分割入力） ---
    def write_dt_block(base: str, src_key: str):
        """
        base: "13" / "19" / "36" など行番号の基点（通報=13, 現着=19, 完了=36）
        C(年) F(月) H(日) J(曜) M(時) O(分)
        """
        dt = _try_parse_datetime(data.get(src_key))
        y, m, d, wd, hh, mm = _split_dt_components(dt)
        cellmap = {
            "Y": f"C{base}",
            "Mo": f"F{base}",
            "D": f"H{base}",
            "W": f"J{base}",
            "H": f"M{base}",
            "Min": f"O{base}",
        }
        if y is not None: ws[cellmap["Y"]] = y
        if m is not None: ws[cellmap["Mo"]] = m
        if d is not None: ws[cellmap["D"]] = d
        if wd is not None: ws[cellmap["W"]] = wd
        if hh is not None: ws[cellmap["H"]] = f"{hh:02d}"
        if mm is not None: ws[cellmap["Min"]] = f"{mm:02d}"

    # 通報（受信時刻）：13行
    write_dt_block("13", "受信時刻")
    # 現着：19行
    write_dt_block("19", "現着時刻")
    # 完了：36行
    write_dt_block("36", "完了時刻")

    # --- 複数行（最大5行、超過は「…」） ---
    def fill_multiline(col_letter: str, start_row: int, text: Optional[str], max_lines: int = 5):
        lines = _split_lines(text, max_lines=max_lines)
        # まず空クリア
        for i in range(max_lines):
            ws[f"{col_letter}{start_row + i}"] = ""
        for idx, line in enumerate(lines[:max_lines]):
            ws[f"{col_letter}{start_row + idx}"] = line

    # 現着状況：C20~C24
    fill_multiline("C", 20, data.get("現着状況"), max_lines=5)
    # 原因：C25~C29
    fill_multiline("C", 25, data.get("原因"), max_lines=5)
    # 処置内容：C30~C34
    fill_multiline("C", 30, data.get("処置内容"), max_lines=5)

    # 完成バイトへ
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

# Step 1: 認証
if st.session_state.step == 1:
    st.subheader("Step 1. パスコード認証")
    pw = st.text_input("パスコードを入力してください", type="password")
    if st.button("次へ"):
        if pw == PASSCODE:
            st.session_state.authed = True
            st.session_state.step = 2
            st.experimental_rerun()
        else:
            st.error("パスコードが違います。")

# Step 2: 本文貼付 + テンプレアップロード
elif st.session_state.step == 2 and st.session_state.authed:
    st.subheader("Step 2. メール本文の貼り付け / テンプレの指定")

    tmpl = st.file_uploader(
        "テンプレート（.xlsx）をアップロードしてください（※ .xlsb をExcelで一度 .xlsx 保存したもの）",
        type=["xlsx"],
        accept_multiple_files=False,
    )
    if tmpl is not None:
        st.session_state.template_xlsx_bytes = tmpl.read()
        st.success("テンプレートを読み込みました。")

    text = st.text_area(
        "SoftBankメール（件名〜本文）を貼り付け",
        height=240,
        placeholder="例：\n件名: 【故障完了】 HK1-234・ABCビル\n管理番号：HK1-234\n物件名：ABCビル\n住所：北海道札幌市中央区南10条西\n…"
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
                st.session_state.step = 3
                st.experimental_rerun()
    with c2:
        if st.button("クリア", use_container_width=True):
            st.session_state.extracted = None
            st.session_state.template_xlsx_bytes = None

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
            st.markdown(f"- 対応者：{data.get('対応者') or ''}")
            st.markdown(f"- 送信者：{data.get('送信者') or ''}")
            st.markdown(f"- 受付番号：{data.get('受付番号') or ''}")
            st.markdown(f"- 受付URL：{data.get('受付URL') or ''}")
            st.markdown(f"- 現着・完了登録URL：{data.get('現着完了登録URL') or ''}")
            st.markdown(f"- 案件種別(件名)：{data.get('案件種別(件名)') or ''}")

        st.divider()
        # 生成
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

        if st.button("最初に戻る", use_container_width=True):
            st.session_state.step = 1
            st.session_state.extracted = None
            st.session_state.template_xlsx_bytes = None
            st.experimental_rerun()
else:
    st.warning("認証が必要です。Step 1に戻ります。")
    st.session_state.step = 1
