import re
from typing import List, Optional
import pandas as pd


def _norm(s: str) -> str:
    """Normalize string for fuzzy column matching."""
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())


def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """
    Find the first matching column in df among candidate names.
    Matching is case-insensitive and punctuation-insensitive.
    """
    if df is None or df.empty:
        return None

    lookup = {_norm(c): c for c in df.columns}

    # Exact normalized match
    for cand in candidates:
        key = _norm(cand)
        if key in lookup:
            return lookup[key]

    # Partial / contains-style match
    for col in df.columns:
        col_norm = _norm(col)
        for cand in candidates:
            if _norm(cand) in col_norm:
                return col

    return None


def detect_columns(df: pd.DataFrame) -> dict:
    """
    Centralized column mapping used by all pages.
    """
    return {
        "reg_date": find_col(df, ["RegistrationDate", "Reg Date", "Date"]),
        "lab_id": find_col(df, ["Lab ID", "Registration Number"]),
        "name": find_col(df, ["Name", "Patient Name", "Student Name"]),
        "gender": find_col(df, ["Gender", "Sex"]),
        "age": find_col(df, ["Age", "Age Years", "Age (Years)"]),

        "locality": find_col(df, ["Locality", "Address", "Area", "School", "College"]),
        "education": find_col(df, ["Education", "Qualification"]),
        "occupation": find_col(df, ["Occupation", "Profession"]),
        "income": find_col(df, ["Income", "Income Group"]),

        "glucose": find_col(df, ["Fasting Glucose", "GLUCOSE FASTING"]),
        "chol": find_col(df, ["Cholesterol", "Total Cholesterol"]),
        "creatinine": find_col(df, ["Creatinine"]),
        "alt": find_col(df, ["ALT", "SGPT"]),
        "protein_total": find_col(df, ["Total Protein", "PROTEIN, TOTAL"]),
        "albumin": find_col(df, ["Albumin"]),
        "globulin": find_col(df, ["Globulin"]),

        "weight": find_col(df, ["Weight", "WEIGHT IN KG"]),
        "height": find_col(df, ["Height", "HEIGHT"]),
        "bmi": find_col(df, ["BMI"]),

        "bp_sys": find_col(df, ["Systolic BP", "BP Systolic"]),
        "bp_dia": find_col(df, ["Diastolic BP", "BP Diastolic"]),

        "tobacco": find_col(df, ["Tobacco", "Smoking"]),
        "alcohol": find_col(df, ["Alcohol"]),
        "diet": find_col(df, ["Diet", "Dietary Recall"]),
        "sleep": find_col(df, ["Sleep Pattern"]),
    }
