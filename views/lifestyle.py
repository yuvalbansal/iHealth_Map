import streamlit as st
import plotly.express as px
from utils.formatting import parse_diet, flag_yes


def render(df, view, cols, meta):
    st.header("🍽️ Lifestyle & behaviour patterns")

    if len(view) == 0:
        st.info("No records match the filters.")
        return

    col1, col2 = st.columns(2)

    if cols.get("diet"):
        with col1:
            d = view[cols["diet"]].map(parse_diet).value_counts()
            fig = px.pie(
                values=d.values,
                names=d.index,
                hole=0.55,
                title="Diet pattern",
            )
            st.plotly_chart(fig, width="stretch")

    if cols.get("sleep"):
        with col2:
            s = view[cols["sleep"]].astype(str).value_counts().head(8)
            fig = px.bar(
                x=s.index,
                y=s.values,
                title="Sleep pattern (top responses)",
                text=s.values,
            )
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, width="stretch")

    col3, col4 = st.columns(2)

    if cols.get("tobacco"):
        with col3:
            t = view[cols["tobacco"]].map(flag_yes).value_counts()
            fig = px.bar(
                x=t.index,
                y=t.values,
                title="Tobacco use",
                text=t.values,
            )
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, width="stretch")

    if cols.get("alcohol"):
        with col4:
            a = view[cols["alcohol"]].map(flag_yes).value_counts()
            fig = px.bar(
                x=a.index,
                y=a.values,
                title="Alcohol / substance use",
                text=a.values,
            )
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, width="stretch")
