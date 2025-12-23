import streamlit as st
import plotly.express as px
import pandas as pd


def render(df, view, cols, meta):
    st.header("🩺 Clinical indicators & risk stratification")

    if len(view) == 0:
        st.info("No records match the filters.")
        return

    col1, col2 = st.columns(2)

    if cols.get("glucose"):
        with col1:
            fig = px.histogram(
                view,
                x=cols["glucose"],
                nbins=40,
                title="Fasting glucose distribution",
            )
            st.plotly_chart(fig, width="stretch")

    if cols.get("chol"):
        with col2:
            fig = px.histogram(
                view,
                x=cols["chol"],
                nbins=40,
                title="Cholesterol distribution",
            )
            st.plotly_chart(fig, width="stretch")

    if cols.get("bp_sys") and cols.get("bp_dia"):
        st.subheader("Blood pressure (Systolic × Diastolic)")

        bp_df = (
            view[[cols["bp_sys"], cols["bp_dia"]]]
            .dropna()
            .astype(float)
        )

        bp_counts = (
            bp_df.groupby([cols["bp_sys"], cols["bp_dia"]], observed=False)
            .size()
            .reset_index(name="Count")
        )

        fig = px.scatter_3d(
            bp_counts,
            x=cols["bp_sys"],
            y=cols["bp_dia"],
            z="Count",
            color="Count",
            title="3D Blood pressure density",
        )

        st.plotly_chart(fig, width="stretch")

    if cols.get("bmi"):
        st.subheader("BMI distribution")
        fig = px.histogram(
            view,
            x=cols["bmi"],
            nbins=40,
            title="BMI distribution",
        )
        st.plotly_chart(fig, width="stretch")
