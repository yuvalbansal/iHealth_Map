import streamlit as st
import plotly.express as px
import pandas as pd


def render(df, view, cols, meta):
    st.header("🏘️ Community / locality health overview")

    loc_col = cols.get("locality")

    if not loc_col:
        st.info("No locality column detected.")
        return

    if len(view) == 0:
        st.info("No records match the filters.")
        return

    agg = view.copy()
    agg["Unhealthy"] = (agg["Health Status"] != "Healthy").astype(int)

    grp = (
        agg.groupby(loc_col)
        .agg(
            Population=("Health Status", "count"),
            Unhealthy_rate=("Unhealthy", "mean"),
        )
        .reset_index()
    )

    grp["Unhealthy_rate"] *= 100

    fig = px.bar(
        grp.sort_values("Population", ascending=False).head(25),
        x=loc_col,
        y="Population",
        title="Top localities by population",
        text="Population",
    )
    fig.update_traces(textposition="outside")
    st.plotly_chart(fig, width="stretch")

    fig2 = px.bar(
        grp.sort_values("Population", ascending=False).head(25),
        x=loc_col,
        y="Unhealthy_rate",
        title="Unhealthy rate by locality (%)",
        text="Unhealthy_rate",
    )
    fig2.update_traces(texttemplate="%{text:.1f}", textposition="outside")
    st.plotly_chart(fig2, width="stretch")

    st.subheader("Locality table")
    st.dataframe(grp.sort_values("Population", ascending=False), width="stretch")
