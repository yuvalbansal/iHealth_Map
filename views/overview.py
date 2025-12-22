import streamlit as st
import plotly.express as px
import pandas as pd
import numpy as np
from utils.formatting import donut_normal_abnormal, as_num

def render(df, view, cols, meta):
    st.header("📊 Population Overview")

    if len(view) == 0:
        st.info("No records match the current filters.")
        return

    total_all = meta["total_rows"]
    total_filt = len(view)

    status_counts = (
        view["Health Status"]
        .value_counts()
        .reindex(["Healthy", "At Risk", "Needs Attention"])
        .fillna(0)
        .astype(int)
    )

    c1, c2, c3, c4 = st.columns(4)
    for col, label, value in [
        (c1, "Total in dataset", f"{total_all:,}"),
        (c2, "Total (filtered)", f"{total_filt:,}"),
        (c3, "Healthy", f"{status_counts.get('Healthy', 0):,}"),
        (c4, "Needs attention", f"{status_counts.get('Needs Attention', 0):,}"),
    ]:
        with col:
            st.markdown(
                    f"""
<div class="metric-card">
  <div class="metric-label">{label}</div>
  <div class="metric-value">{value}</div>
</div>
""",
                    unsafe_allow_html=True,
                )

    st.subheader("Health status distribution")

    colA, colB = st.columns(2)

    with colA:
        fig = px.pie(
            names=status_counts.index,
            values=status_counts.values,
            hole=0.55,
            title="Health status distribution",
        )
        fig.update_traces(textinfo="percent+label")
        st.plotly_chart(fig, width="stretch")

    if cols.get("gender"):
        with colB:
            g = (
                view[cols["gender"]]
                .astype(str)
                .value_counts()
                .reset_index()
            )
            g.columns = ["Gender", "Count"]
            fig_g = px.bar(
                g,
                x="Gender",
                y="Count",
                title="Gender distribution",
                text="Count",
            )
            fig_g.update_traces(textposition="outside")
            st.plotly_chart(fig_g, width="stretch")

    st.subheader("Key biochemical risk flags (approximate)")
    risk_cards = []
    if cols["glucose"]:
        g = as_num(view[cols["glucose"]])
        risk_cards.append(
            ("High fasting glucose (≥126)", f"{100 * (g >= 126).mean():.1f}%")
        )
    if cols["chol"]:
        cchol = as_num(view[cols["chol"]])
        risk_cards.append(
            ("High cholesterol (≥240)", f"{100 * (cchol >= 240).mean():.1f}%")
        )
    if cols["bmi"]:
        b = as_num(view[cols["bmi"]])
        risk_cards.append(("Obesity (BMI ≥30)", f"{100 * (b >= 30).mean():.1f}%"))

    if risk_cards:
        rc1, rc2, rc3 = st.columns(3)
        cols_rc = [rc1, rc2, rc3]
        for idx, (label, val) in enumerate(risk_cards):
            if idx >= len(cols_rc):
                break
            with cols_rc[idx]:
                st.markdown(
                    f"""
<div class="metric-card">
<div class="metric-label">{label}</div>
<div class="metric-value">{val}</div>
</div>
""",
                    unsafe_allow_html=True,
                )


    st.subheader("Normal vs abnormal indicators")

    charts = []

    # Glucose
    if cols.get("glucose"):
        fig = donut_normal_abnormal(
            view[cols["glucose"]],
            "Fasting glucose",
            (70, 99),
        )
        charts.append(fig)

    # Cholesterol
    if cols.get("chol"):
        fig = donut_normal_abnormal(
            view[cols["chol"]],
            "Total cholesterol",
            (0, 199),
        )
        charts.append(fig)

    # BMI
    if cols.get("bmi"):
        b = pd.to_numeric(view[cols["bmi"]], errors="coerce")
        series = pd.Series(np.where((b >= 18.5) & (b < 25), b, np.nan))
        fig = donut_normal_abnormal(
            series,
            "BMI (18.5–24.9 normal)",
            (0, np.inf),
        )
        charts.append(fig)

    # -------------------------------------------------
    # Render charts: 2 per row (half-page each)
    # -------------------------------------------------
    for i in range(0, len(charts), 2):
        row = charts[i : i + 2]
        cols_row = st.columns(2)

        for col, fig in zip(cols_row, row):
            with col:
                st.plotly_chart(fig, width="stretch")

    st.subheader("Sample records")
    st.dataframe(view.head(100), width="stretch")
    st.markdown(
            '<div class="small-caption">Only first 100 rows shown for display; all '
            "calculations are done on the full filtered dataset.</div>",
            unsafe_allow_html=True,
        )
