import streamlit as st
import plotly.io as pio

def apply_theme():
    pio.templates.default = "plotly_dark"
    st.markdown(
        """
        <style>
        body { background-color: #02010c; color: #E8E8E8; }
        </style>
        """,
        unsafe_allow_html=True,
    )
