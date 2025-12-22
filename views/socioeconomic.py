import streamlit as st
import plotly.express as px


def render(df, view, cols, meta):
    st.header("👥 Socioeconomic & profession-wise patterns")

    if len(view) == 0:
        st.info("No records match the filters.")
        return

    if cols.get("income"):
        inc = view[cols["income"]].astype(str).value_counts().head(10)
        fig = px.bar(
            x=inc.index,
            y=inc.values,
            title="Income groups",
            text=inc.values,
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, width="stretch")

    if cols.get("occupation"):
        occ = view[cols["occupation"]].astype(str).value_counts().head(15)
        fig = px.bar(
            x=occ.index,
            y=occ.values,
            title="Top professions",
            text=occ.values,
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, width="stretch")
