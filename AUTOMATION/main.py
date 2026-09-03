import importlib
import importlib.machinery
import importlib.util
from pathlib import Path

import streamlit as st

st.set_page_config(
    page_title="Automation Hub",
    page_icon="⚙️",
    layout="wide",
)

page = st.sidebar.selectbox(
    "Choose Automation",
    ["BPI Cert Request", "MAD PL PYTS", "CC Auto", "ATU"],
    key="main_page_selector",
)

page_modules = {
    "BPI Cert Request": "BpiCertRequest",
    "MAD PL PYTS": "MadPlPyts",
    "CC Auto": "CCauto",
    "ATU": "ATU",
}


def load_local_module_from_path(module_name):
    base_dir = Path(__file__).resolve().parent
    candidate_paths = [
        base_dir / module_name,
        base_dir / f"{module_name}.py",
    ]

    for candidate_path in candidate_paths:
        if not candidate_path.exists():
            continue

        loader = importlib.machinery.SourceFileLoader(module_name, str(candidate_path))
        spec = importlib.util.spec_from_loader(loader.name, loader)
        if spec is None:
            continue

        module = importlib.util.module_from_spec(spec)
        loader.exec_module(module)
        return module

    raise ModuleNotFoundError(f"Could not find a local module file for '{module_name}'.")


def load_selected_module(module_name):
    # Some page modules still call st.set_page_config() at import time.
    # The hub already configured the page, so suppress duplicate calls here.
    original_set_page_config = st.set_page_config
    st.set_page_config = lambda *args, **kwargs: None
    try:
        try:
            return importlib.import_module(module_name)
        except ModuleNotFoundError:
            return load_local_module_from_path(module_name)
    finally:
        st.set_page_config = original_set_page_config


selected_module = load_selected_module(page_modules[page])
selected_module.run()
