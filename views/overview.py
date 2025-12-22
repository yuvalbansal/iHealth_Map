import streamlit as st
import plotly.express as px
import pandas as pd
import numpy as np
from utils.formatting import donut_normal_abnormal

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
            st.metric(label, value)

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

    st.subheader("Normal vs abnormal indicators")

    d1, d2, d3 = st.columns(3)

    if cols.get("glucose"):
        with d1:
            fig = donut_normal_abnormal(
                view[cols["glucose"]],
                "Fasting glucose",
                (70, 99),
            )
            st.plotly_chart(fig, width="stretch")

    if cols.get("chol"):
        with d2:
            fig = donut_normal_abnormal(
                view[cols["chol"]],
                "Total cholesterol",
                (0, 199),
            )
            st.plotly_chart(fig, width="stretch")

    if cols.get("bmi"):
        with d3:
            b = pd.to_numeric(view[cols["bmi"]], errors="coerce")
            series = pd.Series(np.where((b >= 18.5) & (b < 25), b, np.nan))
            fig = donut_normal_abnormal(
                series,
                "BMI (18.5–24.9 normal)",
                (0, np.inf),
            )
            st.plotly_chart(fig, width="stretch")

    st.subheader("Sample records")
    st.dataframe(view.head(100), width="stretch")
