import pandas as pd
import plotly.express as px
from typing import Optional, Tuple


def as_num(series: pd.Series) -> pd.Series:
    """Safely convert a series to numeric, coercing errors to NaN."""
    return pd.to_numeric(series, errors="coerce")


def parse_diet(val) -> str:
    s = str(val).strip().lower()
    if not s or s == "nan":
        return "Unknown"
    s = s.replace("-", " ").replace("/", " ")
    if "non" in s and "veg" in s:
        return "Non-Vegetarian"
    if "veg" in s:
        return "Vegetarian"
    return "Unknown"


def flag_yes(val) -> str:
    s = str(val).strip().lower()
    if not s or s == "nan":
        return "Unknown"
    if any(tok in s for tok in ["yes", "y", "positive", "present"]):
        return "Yes"
    if any(tok in s for tok in ["no", "nil", "none", "absent"]):
        return "No"
    return "Unknown"


def donut_normal_abnormal(
    series: pd.Series,
    title: str,
    thresholds: Tuple[Optional[float], Optional[float]],
) -> px.pie:
    """
    Build a donut chart for Normal vs Abnormal.
    thresholds = (low_ok, high_ok)
    """
    s = pd.to_numeric(series, errors="coerce").dropna()

    if s.empty:
        df_plot = pd.DataFrame({"Category": ["No data"], "Count": [1]})
    else:
        low, high = thresholds
        if low is not None and high is not None:
            normal = s.between(low, high)
        elif low is None and high is not None:
            normal = s <= high
        elif low is not None and high is None:
            normal = s >= low
        else:
            normal = pd.Series(False, index=s.index)

        df_plot = pd.DataFrame(
            {
                "Category": ["Normal", "Abnormal"],
                "Count": [int(normal.sum()), int((~normal).sum())],
            }
        )

    fig = px.pie(
        df_plot,
        values="Count",
        names="Category",
        hole=0.55,
        title=title,
    )
    fig.update_traces(textinfo="percent+label")
    fig.update_layout(showlegend=False)

    return fig
