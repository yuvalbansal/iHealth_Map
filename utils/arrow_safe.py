import pandas as pd


def make_arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert object columns that contain numeric-looking values into numeric types.
    Leaves pure text columns untouched.
    This prevents Streamlit / PyArrow serialization errors.
    """

    def safe_numeric(series: pd.Series) -> pd.Series:
        try:
            converted = pd.to_numeric(series, errors="coerce")
            # Keep conversion only if at least one valid numeric value exists
            if converted.notna().any():
                return converted
            return series
        except Exception:
            return series

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = safe_numeric(df[col])

    return df
