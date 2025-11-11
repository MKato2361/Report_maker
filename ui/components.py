# report_maker/ui/components.py
import streamlit as st
from core.state import get_working_dict, set_working_value
from core.settings import REQUIRED_KEYS
from core.textutil import split_lines

def is_required_missing(data: dict, key: str) -> bool:
    return key in REQUIRED_KEYS and not (data.get(key) or "").strip()

def display_text(value: str, max_lines: int):
    if not value:
        return ""
    if max_lines and max_lines > 1:
        lines = split_lines(value, max_lines=max_lines)
        return "<br>".join(lines)
    return (value or "").replace("\n", "<br>")

def render_field(label: str, key: str, max_lines: int = 1, placeholder: str = "", editable_in_bulk: bool = False):
    data = get_working_dict()
    val = data.get(key) or ""
    missing = is_required_missing(data, key)

    cols = st.columns([0.22, 0.78])
    with cols[0]:
        st.markdown(("⚠️　**" if missing else "**") + f"{label}**")

    with cols[1]:
        if st.session_state.get("edit_mode") and editable_in_bulk:
            if max_lines == 1:
                new_val = st.text_input("", value=val, placeholder=placeholder, key=f"in_{key}")
            else:
                new_val = st.text_area("", value=val, placeholder=placeholder, height=max(80, max_lines * 24), key=f"ta_{key}")
            set_working_value(key, new_val)
        else:
            st.markdown("<span class='missing'>未入力</span>" if missing else display_text(val, max_lines=max_lines),
                        unsafe_allow_html=True)
