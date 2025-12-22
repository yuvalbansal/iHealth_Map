import pandas as pd
import numpy as np
from utils.arrow_safe import make_arrow_safe
from utils.column_detection import detect_columns
from utils.health_rules import apply_health_assessment

def load_and_prepare_data(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, engine="openpyxl")
    df = make_arrow_safe(df_raw.copy())

    cols = detect_columns(df)
    view = apply_health_assessment(df, cols)

    meta = {
        "source_name": uploaded_file.name,
        "total_rows": len(df_raw),
        "filtered_rows": len(view),
    }

    return df, view, cols, meta
