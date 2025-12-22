import streamlit as st
import pandas as pd
import numpy as np
from typing import List

from data.data_loader import load_and_prepare_data
from utils.formatting import parse_diet, flag_yes
from utils.health_rules import apply_health_assessment
import plotly.io as pio

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

pio.templates.default = "plotly_dark"

st.set_page_config(
    page_title="Bharat iHealthMap",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------
#  Custom CSS – dark background, big fonts
# ---------------------------------------------------------------------
st.markdown(
    """
    <style>
        /* DARK BACKGROUND */
    [data-testid="stAppViewContainer"] {
    background: radial-gradient(circle at top left,
        #161b3f 0, #050816 40%, #02010c 100%);
    font-size: 15px;
    }

    
   /* force all visible content text to gold */
    [data-testid="stAppViewContainer"] p, 
    li {
    color: #E8E8E8; font-size: 20px !important;}

    [data-testid="stAppViewContainer"] h1 { color: #00FF00 !important; }   /* gold */
    [data-testid="stAppViewContainer"] h2 { color: #EFBF04 !important; }   /* cyan */
    [data-testid="stAppViewContainer"] h3 { color: #00FF00 !important; }   /* white */
    [data-testid="stAppViewContainer"] h4 { color: #cccccc !important; }   /* light grey */





    [data-testid="stHeader"] {
        background: rgba(5, 8, 22, 0.1);
    }

    [data-testid="stSidebar"] {
        background: #050816;
    }

    h1, h2, h3, h4 {
        font-weight: 700 !important;
    }

    .metric-card {
        background: #0b1024;
        border-radius: 14px;
        padding: 14px 18px;
        box-shadow: 0 0 16px rgba(15, 23, 42, 0.8);
    }

    .metric-value {
        font-size: 32px;
        font-weight: 800;
        color: #FFD700;
    }

    .metric-label {
        font-size: 14px;
        color: #FFD700;
    }

    .stDataFrame, .stTable {
        font-size: 15px;
    }

    .small-caption {
        font-size: 13px;
        color: #FFD700;
    }

    /* Sidebar text colors */
    [data-testid="stSidebar"] * {
        color: #FFD700 !important;
    }

    /* Sidebar section headers */
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] h4,
    [data-testid="stSidebar"] h5,
    [data-testid="stSidebar"] h6 {
        color: #FFD700 !important;
        font-weight: 700 !important;
    }

    /* Top nav – we'll underline the area where the radio sits */
    .top-nav-wrapper {
        border-bottom: 3px solid #FF9933; /* saffron */
        padding-bottom: 4px;
        margin-top: 10px;
        margin-bottom: 12px;
    }

    /* Radio styling used as top menu */
    .top-nav-wrapper .stRadio > label {
        font-size: 0; /* hide text label */
    }
    .top-nav-wrapper .stRadio div[role="radiogroup"] {
        display: flex;
        justify-content: center;
        gap: 40px;
    }
    .top-nav-wrapper .stRadio div[role="radiogroup"] > div {
        margin-right: 0 !important;
    }
    .top-nav-wrapper .stRadio div[role="radiogroup"] label {
        font-size: 40px !important;
        font-weight: 600 !important;
        color: #FFD700 !important;
        cursor: pointer;
        padding-bottom: 4px;
    }
    .top-nav-wrapper .stRadio div[role="radiogroup"] input:checked + div > label {
        color: #FFD700 !important;
        border-bottom: 3px solid #FFFFFF;
    }
    
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------
#  Page keys + labels for navbar
# ---------------------------------------------------------------------
PAGE_KEYS = [
    "Overview",
    "Clinical",
    "Lifestyle",
    "Community",
    "Socioeconomic_Profession",
    "Downloads",
]

PAGE_LABELS = {
    "Overview": "Overview",
    "Clinical": "Clinical",
    "Lifestyle": "Lifestyle",
    "Community": "Community",
    "Socioeconomic_Profession": "Socioeconomic & Profession",
    "Downloads": "Downloads",
}

if "page" not in st.session_state:
    st.session_state.page = "Overview"

# ---------------------------------------------------------------------
#  TITLE SECTION (TRICOLOUR UNDERLINE LOOK)
# ---------------------------------------------------------------------
st.markdown(
    """
    <div style="text-align:center; margin-top:10px;">
        <span style="display:block; font-size:44px; font-weight:900; color:#EFBF04;">
            Bharat_iHealthMap
        </span>
        <span style="display:block; font-size:28px; font-weight:600; color:#E8E8E8;">
            Community Health Dashboard
        </span>
        <div style="height:6px; margin-top:6px;
            background: linear-gradient(to right,
                #FF9933 0%,  /* saffron */
                #FFFFFF 50%, /* white */
                #138808 100% /* green */
            );
            border-radius:4px;">
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------
#  TOP BAR DATA UPLOAD
# ---------------------------------------------------------------------
st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

upload_col1, upload_col2, upload_col3 = st.columns([2, 4, 2])

with upload_col2:
    st.markdown(
        """
        <div style="
            text-align:center;
            padding:8px;
            border-radius:10px;
            background:rgba(255,255,255,0.08);
            border:1px solid rgba(255,255,255,0.15);
            backdrop-filter: blur(4px);
        ">
            <span style="font-size:18px; color:#FFFFFF; font-weight:200;">
                📁 Upload Data (.xlsx)
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx"],
        label_visibility="collapsed",
        help="Upload an Excel (.xlsx) file for population analytics.",
    )

if uploaded_file:
    st.session_state["uploaded_file"] = uploaded_file

if "uploaded_file" not in st.session_state:
    st.stop()

# 1. Load basic cleaned data (no assessment yet)
df, cols, meta = load_and_prepare_data(
    st.session_state["uploaded_file"]
)

# ---------------------------------------------------------------------
#  Sidebar filters
# ---------------------------------------------------------------------
st.sidebar.header("🔍 Filters")

# Age slider (single definition – uses __AGE__)
age_vals = df["__AGE__"].dropna()
if len(age_vals) > 0:
    amin = int(age_vals.min())
    amax = int(age_vals.max())
    if amin == amax:
        slider_min = max(0, amin - 1)
        slider_max = amin + 1
    else:
        slider_min = max(0, amin)
        slider_max = max(amax, amin + 1)
    age_min, age_max = st.sidebar.slider(
        "Age range (years)",
        min_value=slider_min,
        max_value=slider_max,
        value=(slider_min, slider_max),
        help="All analytics below will respect this filtered age range.",
    )
else:
    age_min, age_max = 0, 120

# Gender filter
if cols.get("gender"):
    gender_vals = sorted(df[cols["gender"]].dropna().astype(str).unique().tolist())
    gender_sel = st.sidebar.multiselect("Gender", options=gender_vals, default=gender_vals)
else:
    gender_sel = []

# Diet
if cols.get("diet"):
    all_diet = df[cols["diet"]].map(parse_diet)
    diet_vals = sorted(all_diet.dropna().unique().tolist())
    diet_sel = st.sidebar.multiselect("Diet type", options=diet_vals, default=diet_vals)
else:
    diet_sel = []

# Tobacco + Alcohol
tob_sel: List[str] = []
alc_sel: List[str] = []
if cols.get("tobacco"):
    tob_vals = sorted(df[cols["tobacco"]].dropna().map(flag_yes).unique().tolist())
    tob_sel = st.sidebar.multiselect("Tobacco / smoking", options=tob_vals, default=tob_vals)
if cols.get("alcohol"):
    alc_vals = sorted(df[cols["alcohol"]].dropna().map(flag_yes).unique().tolist())
    alc_sel = st.sidebar.multiselect("Alcohol / drugs", options=alc_vals, default=alc_vals)

# ---------------------------------------------------------------------
#  Apply Filters
# ---------------------------------------------------------------------
mask = pd.Series(True, index=df.index)

# Age
mask &= df["__AGE__"].between(age_min, age_max) | df["__AGE__"].isna()

# Gender
if cols.get("gender") and gender_sel:
    mask &= df[cols["gender"]].astype(str).isin(gender_sel)

# Diet
if cols.get("diet") and diet_sel:
    mask &= df[cols["diet"]].map(parse_diet).isin(diet_sel)

# Tobacco
if cols.get("tobacco") and tob_sel:
    mask &= df[cols["tobacco"]].map(flag_yes).isin(tob_sel)

# Alcohol
if cols.get("alcohol") and alc_sel:
    mask &= df[cols["alcohol"]].map(flag_yes).isin(alc_sel)

filtered = df[mask].copy()
st.sidebar.success(f"Filtered records: {len(filtered):,} / {len(df):,}")

# ---------------------------------------------------------------------
#  Health Assessment (on filtered data)
# ---------------------------------------------------------------------
if len(filtered) > 0:
    view = apply_health_assessment(filtered, cols)
    meta["filtered_rows"] = len(view)
else:
    view = filtered.copy()
    meta["filtered_rows"] = 0

# ---------------------------------------------------------------------
#  Navigation & Page Rendering
# ---------------------------------------------------------------------
with st.container():
    st.markdown('<div class="top-nav-wrapper">', unsafe_allow_html=True)
    selected_page = st.radio(
        "Main navigation",
        options=PAGE_KEYS,
        index=PAGE_KEYS.index(st.session_state.page)
        if st.session_state.page in PAGE_KEYS
        else 0,
        format_func=lambda key: PAGE_LABELS[key],
        horizontal=True,
        key="topnav_radio",
    )
    st.markdown("</div>", unsafe_allow_html=True)

@st.cache_data(show_spinner=True)
def load_excel(uploaded_file):
    """Safely read Excel using openpyxl if possible."""
    try:
        return pd.read_excel(uploaded_file, engine="openpyxl")
    except TypeError:
        # Fallback for older Pandas/Excel engines
        return pd.read_excel(uploaded_file)


df_raw = load_excel(uploaded_file)
source_name = uploaded_file.name

st.caption(
    f"📄 Data source: **{source_name}** — rows: {len(df_raw):,}, "
    f"columns: {len(df_raw.columns)}"
)

st.session_state.page = selected_page
page = selected_page

PAGES[page](df=df, view=view, cols=cols, meta=meta)
