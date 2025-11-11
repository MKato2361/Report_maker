# report_maker/core/state.py
import copy
import os
import streamlit as st

def get_passcode() -> str:
    try:
        val = st.secrets.get("APP_PASSCODE")
        if val:
            return str(val)
    except Exception:
        pass
    env_val = os.getenv("APP_PASSCODE")
    return str(env_val) if env_val else ""

def ensure_extracted():
    if "extracted" not in st.session_state or st.session_state.extracted is None:
        st.session_state.extracted = {}

def enter_edit_mode():
    ensure_extracted()
    st.session_state.edit_mode = True
    st.session_state.edit_buffer = copy.deepcopy(st.session_state.extracted)

def cancel_edit():
    st.session_state.edit_mode = False
    st.session_state.edit_buffer = {}

def save_edit():
    st.session_state.extracted = copy.deepcopy(st.session_state.edit_buffer)
    st.session_state.edit_mode = False
    st.session_state.edit_buffer = {}

def get_working_dict() -> dict:
    if st.session_state.get("edit_mode"):
        return st.session_state.edit_buffer
    return st.session_state.extracted or {}

def set_working_value(key: str, value: str):
    if st.session_state.get("edit_mode"):
        st.session_state.edit_buffer[key] = value
    else:
        ensure_extracted()
        st.session_state.extracted[key] = value
