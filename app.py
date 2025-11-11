# report_maker/app.py
import streamlit as st
from core.settings import APP_TITLE
from ui.styles import inject_styles
from ui.steps import render_app

st.set_page_config(page_title=APP_TITLE, layout="centered")
inject_styles()
render_app()
