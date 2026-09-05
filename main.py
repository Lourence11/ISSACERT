import importlib

import streamlit as st

st.set_page_config(
    page_title="Automation Hub",
    page_icon="⚙️",
    layout="wide",
)

page = st.sidebar.selectbox(
    "Choose Automation",
    ["BPI Cert Request", "MAD PL PYTS", "CC Auto"],
    key="main_page_selector",
)

page_modules = {
    "BPI Cert Request": "BpiCertRequest",
    "MAD PL PYTS": "MadPlPyts",
    "CC Auto": "CCauto",
}


def load_selected_module(module_name):
    # Some page modules still call st.set_page_config() at import time.
    # The hub already configured the page, so suppress duplicate calls here.
    original_set_page_config = st.set_page_config
    st.set_page_config = lambda *args, **kwargs: None
    try:
        return importlib.import_module(module_name)
    finally:
        st.set_page_config = original_set_page_config


selected_module = load_selected_module(page_modules[page])
selected_module.run()
