import streamlit as st
import pandas as pd
import io

from utils.ppt_builder import build_population_ppt


def render(df, view, cols, meta):
    st.header("📥 Downloads")

    if len(view) == 0:
        st.info("No records match the current filters.")
        return

    # -------------------------------------------------
    # Excel download
    # -------------------------------------------------

    def to_excel_bytes(df_):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_.to_excel(writer, index=False, sheet_name="iHealthMap")
        buf.seek(0)
        return buf.getvalue()

    st.download_button(
        "💾 Download filtered records (.xlsx)",
        data=to_excel_bytes(view),
        file_name="ihealthmap_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # -------------------------------------------------
    # Population PPT
    # -------------------------------------------------

    st.markdown("---")
    st.subheader("📊 Population PPT summary")

    col1, col2 = st.columns(2)
    with col1:
        location = st.text_input("Location for title slide")
    with col2:
        year = st.text_input("Year for title slide")

    if "ppt_buffer" not in st.session_state:
        st.session_state.ppt_buffer = None

    if st.button("Generate PPT summary"):
        with st.spinner("Generating PPT…"):
            st.session_state.ppt_buffer = build_population_ppt(
                view_df=view,
                df_full=df,
                cols_map=cols,
                location_title=location,
                year_title=year,
            )

    if st.session_state.ppt_buffer:
        st.download_button(
            "⬇️ Download PPT summary",
            data=st.session_state.ppt_buffer,
            file_name="bharat_ihealthmap_summary.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
