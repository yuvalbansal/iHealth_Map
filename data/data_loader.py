import streamlit as st
import pandas as pd
import numpy as np
from utils.arrow_safe import make_arrow_safe
from utils.column_detection import detect_columns
from utils.formatting import as_num


@st.cache_data(show_spinner="Processing data...", ttl="2h")
def load_and_prepare_data(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, engine="openpyxl")
    df = make_arrow_safe(df_raw.copy())
    cols = detect_columns(df)

    # ---------------------------------------------------------------------
    #  Basic cleaning & derived columns (Logic from Initial_Code.py)
    # ---------------------------------------------------------------------

    # 1. Age -> numeric helper column __AGE__
    age_col = cols.get("age")
    if age_col:
        # Extract first number found in string
        df["__AGE__"] = df[age_col].astype(str).str.extract(r"(\d+\.?\d*)")[0]
        df["__AGE__"] = pd.to_numeric(df["__AGE__"], errors="coerce")
    else:
        df["__AGE__"] = np.nan

    # 2. BMI compute if missing
    if not cols.get("bmi") and cols.get("height") and cols.get("weight"):
        h_cm = as_num(df[cols["height"]])
        w_kg = as_num(df[cols["weight"]])
        with np.errstate(divide="ignore", invalid="ignore"):
            bmi_calc = w_kg / (h_cm / 100.0) ** 2
        df["__BMI__"] = bmi_calc.round(1)
        cols["bmi"] = "__BMI__"

    # 3. Coerce numeric for lab/vital columns
    numeric_keys = [
        "bmi", "bp_sys", "bp_dia", "glucose", "chol",
        "creatinine", "alt", "protein_total", "albumin",
        "globulin", "waist", "height", "weight"
    ]
    for key in numeric_keys:
        c = cols.get(key)
        if c:
            df[c] = as_num(df[c])

    # 4. ID column (fallback)
    # If no ID, create synthetic
    if not cols.get("lab_id") and not cols.get("name"):
        df["__index__"] = np.arange(1, len(df) + 1)
        cols["lab_id"] = "__index__"

    # Return the cleaned dataframe AND columns. 
    # Assessment will happen in app.py after filtering.
    meta = {
        "source_name": uploaded_file.name,
        "total_rows": len(df_raw),
    }

    return df, cols, meta

