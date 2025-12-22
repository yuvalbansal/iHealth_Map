import streamlit as st
from data.data_loader import load_and_prepare_data
from styles.theme import apply_theme

from views.overview import render as overview_page
from views.clinical import render as clinical_page
from views.lifestyle import render as lifestyle_page
from views.community import render as community_page
from views.socioeconomic import render as socioeconomic_page
from views.downloads import render as downloads_page

PAGES = {
    "Overview": overview_page,
    "Clinical": clinical_page,
    "Lifestyle": lifestyle_page,
    "Community": community_page,
    "Socioeconomic_Profession": socioeconomic_page,
    "Downloads": downloads_page,
}

st.set_page_config(
    page_title="Bharat_iHealthMap",
    layout="wide",
    initial_sidebar_state="collapsed",
)

apply_theme()

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx"],
    label_visibility="collapsed",
)

if uploaded_file:
    st.session_state["uploaded_file"] = uploaded_file

if "uploaded_file" not in st.session_state:
    st.stop()

df, view, cols, meta = load_and_prepare_data(
    st.session_state["uploaded_file"]
)

page = st.radio(
    "Main navigation",
    list(PAGES.keys()),
    format_func=lambda x: x.replace("_", " & "),
    horizontal=True,
    label_visibility="collapsed",
)

PAGES[page](df=df, view=view, cols=cols, meta=meta)
