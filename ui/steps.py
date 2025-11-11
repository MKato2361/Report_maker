# report_maker/ui/steps.py
import os, sys, traceback
import streamlit as st
from core.settings import REQUIRED_KEYS
from core.state import (
    get_passcode, ensure_extracted, enter_edit_mode, cancel_edit, save_edit,
    get_working_dict
)
from core.parsing import extract_fields, minutes_between
from core.excel_writer import fill_template_xlsx, build_filename
from ui.components import render_field  # ← ここはモジュール先頭でインポート

def _init_session():
    if "step" not in st.session_state: st.session_state.step = 1
    if "authed" not in st.session_state: st.session_state.authed = False
    if "extracted" not in st.session_state: st.session_state.extracted = None
    if "affiliation" not in st.session_state: st.session_state.affiliation = ""
    if "template_xlsx_bytes" not in st.session_state: st.session_state.template_xlsx_bytes = None
    if "edit_mode" not in st.session_state: st.session_state.edit_mode = False
    if "edit_buffer" not in st.session_state: st.session_state.edit_buffer = {}
    ensure_extracted()

def _fmt_minutes(v):
    # Noneや負値はハイフン表記
    if v is None or v < 0:
        return "—"
    # 60分以上は「X時間YY分」
    if v >= 60:
        h = v // 60
        m = v % 60
        return f"{h}時間{m:02d}分"
    # 60分未満はそのまま「N分」
    return f"{v}分"

def render_app():
    _init_session()
    PASSCODE = get_passcode()

    # Step 1
    if st.session_state.step == 1:
        st.subheader("Step 1. パスコード認証")
        if not PASSCODE:
            st.info("（注意）現在、PASSCODEがSecrets/環境変数に未設定です。開発モード想定で空文字として扱います。")
        pw = st.text_input("パスコードを入力してください", type="password")
        if st.button("次へ", use_container_width=True):
            if pw == PASSCODE:
                st.session_state.authed = True
                st.session_state.step = 2
                st.rerun()
            else:
                st.error("パスコードが違います。")
        return

    # Step 2
    if st.session_state.step == 2 and st.session_state.authed:
        st.subheader("Step 2. メール本文の貼り付け / 所属 / テンプレ選択")

        template_path = "template.xlsm"
        tpl_col1, tpl_col2 = st.columns([0.55, 0.45])
        with tpl_col1:
            st.caption("① 既定：template.xlsm を探します")
            if os.path.exists(template_path) and not st.session_state.template_xlsx_bytes:
                try:
                    with open(template_path, "rb") as f:
                        st.session_state.template_xlsx_bytes = f.read()
                    st.success(f"テンプレートを読み込みました: {template_path}")
                except Exception as e:
                    st.error(f"テンプレートの読み込みに失敗: {e}")
            elif st.session_state.template_xlsx_bytes:
                st.success("テンプレートは読み込み済みです。")
            else:
                st.warning("既定テンプレートが見つかりません。②のアップロードをご利用ください。")

        with tpl_col2:
            st.caption("② またはテンプレ.xlsmをアップロード")
            up = st.file_uploader("テンプレート（.xlsm）", type=["xlsm"], accept_multiple_files=False)
            if up is not None:
                st.session_state.template_xlsx_bytes = up.read()
                st.success(f"アップロード済み: {up.name}")

        if not st.session_state.template_xlsx_bytes:
            st.error("テンプレートが未準備です。template.xlsm を配置するか、上でアップロードしてください。")
            st.stop()

        aff = st.text_input("所属", value=st.session_state.affiliation)
        st.session_state.affiliation = aff

        processing_after = st.text_input("処理修理後（任意）", value=st.session_state.get("processing_after", ""))
        st.session_state["processing_after"] = processing_after

        text = st.text_area("故障完了メール（本文）を貼り付け", height=240, placeholder="ここにメール本文を貼り付け...")

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
                st.session_state.processing_after = ""
                st.rerun()
        return

    # Step 3
    if st.session_state.step == 3 and st.session_state.authed:
        st.subheader("Step 3. 抽出結果の確認・編集 → Excel生成")

        if "processing_after" in st.session_state and st.session_state.extracted is not None:
            if not st.session_state.extracted.get("_processing_after_initialized"):
                st.session_state.extracted["処理修理後"] = st.session_state.get("processing_after", "")
                st.session_state.extracted["_processing_after_initialized"] = True

        # ① 編集対象：枠内に薄めボタンを配置（編集モードの制御）
        with st.expander("① 編集対象（まとめて編集・すべて必須）", expanded=True):
            c_left, c_mid, c_right = st.columns([1, 1, 1])
            with c_right:
                if not st.session_state.get("edit_mode"):
                    if st.button("✏️ 編集モードに入る", key="enter_edit_inline"):
                        enter_edit_mode()
                        st.rerun()
                else:
                    c1, c2 = st.columns([1, 1])
                    with c1:
                        if st.button("✅ すべて保存", key="save_edit_inline"):
                            save_edit()
                            st.success("保存しました")
                            st.rerun()
                    with c2:
                        if st.button("↩️ 変更を破棄", key="cancel_edit_inline"):
                            cancel_edit()
                            st.info("変更を破棄しました")
                            st.rerun()

            # 入力フィールド群
            render_field("通報者", "通報者", 1, editable_in_bulk=True)
            render_field("受信内容", "受信内容", 4, editable_in_bulk=True)
            render_field("現着状況", "現着状況", 5, editable_in_bulk=True)
            render_field("原因", "原因", 5, editable_in_bulk=True)
            render_field("処置内容", "処置内容", 5, editable_in_bulk=True)
            render_field("処理修理後", "処理修理後", 1, editable_in_bulk=True)
            render_field("所属", "所属", 1, editable_in_bulk=True)

        # ② 基本情報（表示）
        with st.expander("② 基本情報（表示）", expanded=True):
            render_field("管理番号", "管理番号", 1)
            render_field("物件名", "物件名", 1)
            render_field("住所", "住所", 2)
            render_field("窓口会社", "窓口会社", 1)
            render_field("制御方式", "制御方式", 1)
            render_field("契約種別", "契約種別", 1)
            render_field("メーカー", "メーカー", 1)

        # ③ 受付・現着・完了（表示）
        data = get_working_dict()
        with st.expander("③ 受付・現着・完了（表示）", expanded=True):
            render_field("受信時刻", "受信時刻", 1)
            render_field("現着時刻", "現着時刻", 1)
            render_field("完了時刻", "完了時刻", 1)

            t_recv_to_arrive = minutes_between(data.get("受信時刻"), data.get("現着時刻"))
            t_work = minutes_between(data.get("現着時刻"), data.get("完了時刻"))
            t_recv_to_done = minutes_between(data.get("受信時刻"), data.get("完了時刻"))

            c1, c2, c3 = st.columns(3)
            with c1: st.info(f"受付〜現着: { _fmt_minutes(t_recv_to_arrive) }")
            with c2: st.info(f"作業時間: { _fmt_minutes(t_work) }")
            with c3: st.info(f"受付〜完了: { _fmt_minutes(t_recv_to_done) }")

        # ④ その他情報（表示）
        with st.expander("④ その他情報（表示）", expanded=False):
            render_field("対応者", "対応者", 1)
            render_field("送信者", "送信者", 1)
            render_field("受付番号", "受付番号", 1)
            render_field("受付URL", "受付URL", 1)
            render_field("現着完了登録URL", "現着完了登録URL", 1)

        st.divider()

        try:
            is_editing = st.session_state.get("edit_mode", False)
            gen_data = get_working_dict()
            missing_now = [k for k in REQUIRED_KEYS if not (gen_data.get(k) or "").strip()]
            can_generate = (not is_editing) and (not missing_now)

            if can_generate:
                xlsx_bytes = fill_template_xlsx(st.session_state.template_xlsx_bytes, gen_data)
                fname = build_filename(gen_data)
                st.download_button(
                    "Excelを生成（.xlsm）",
                    data=xlsx_bytes,
                    file_name=fname,
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    use_container_width=True,
                    disabled=False,
                    help="一括編集モードはオフ、かつ必須項目がすべて入力されている場合に生成できます",
                )
            else:
                st.download_button(
                    "Excelを生成（.xlsm）",
                    data=b"",
                    file_name="未生成.xlsm",
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    use_container_width=True,
                    disabled=True,
                    help="一括編集モード中は保存後に生成できます。必須未入力がある場合も生成できません。",
                )
                if is_editing:
                    st.warning("一括編集中は生成できません。「✅ すべて保存」を押して編集を確定してください。")
                if missing_now:
                    st.error("未入力の必須項目があります： " + "・".join(missing_now))

        except Exception as e:
            st.error(f"テンプレート書き込み中にエラーが発生しました: {e}")
            with st.expander("詳細（開発者向け）"):
                st.code("".join(traceback.format_exception(*sys.exc_info())), language="python")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("Step2に戻る", use_container_width=True):
                st.session_state.step = 2; st.rerun()
        with c2:
            if st.button("最初に戻る", use_container_width=True):
                st.session_state.step = 1
                st.session_state.extracted = None
                st.session_state.affiliation = ""
                st.session_state.processing_after = ""
                st.session_state.edit_mode = False
                st.session_state.edit_buffer = {}
                st.rerun()
        return

    # 認証未完了・その他
    st.warning("認証が必要です。Step1に戻ります。")
    st.session_state.step = 1
