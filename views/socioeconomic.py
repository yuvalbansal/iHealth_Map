import streamlit as st
import plotly.express as px


def render(df, view, cols, meta):
    st.header("👥 Socioeconomic & profession-wise patterns")

    if len(view) == 0:
        st.info("No records match the filters.")
        return

    if cols.get("income"):
        st.subheader("Income group distribution & health status")
        inc_counts = view[cols["income"]].astype(str).value_counts().head(10)
        fig = px.bar(
            x=inc_counts.index,
            y=inc_counts.values,
            title="Top income categories (filtered)",
            labels={"x": "Income group", "y": "Count"},
            text=inc_counts.values,
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(fig, width="stretch")

        if "Health Status" in view:
            inc_status = (
                view[[cols["income"], "Health Status"]]
                .dropna()
                .groupby([cols["income"], "Health Status"], observed=False)
                .size()
                .reset_index(name="Count")
            )
            # Filter to keep only the top categories found above to avoid clutter
            top_incs = inc_counts.index.tolist()
            inc_status = inc_status[inc_status[cols["income"]].isin(top_incs)]
            
            fig_inc_h = px.bar(
                inc_status,
                x=cols["income"],
                y="Count",
                color="Health Status",
                barmode="group",
                title="Health status by income group (top categories)",
                text="Count",
            )
            fig_inc_h.update_traces(textposition="outside")
            fig_inc_h.update_layout(xaxis_tickangle=-30)
            st.plotly_chart(fig_inc_h, width="stretch")

    if cols.get("occupation"):
        st.subheader("Occupation / profession distribution & risk")
        occ_counts = view[cols["occupation"]].astype(str).value_counts().head(15)
        fig = px.bar(
            x=occ_counts.index,
            y=occ_counts.values,
            title="Top professions (filtered)",
            labels={"x": "Profession", "y": "Count"},
            text=occ_counts.values,
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(fig, width="stretch")
        
        if "Health Status" in view:
            occ_status = (
                view[[cols["occupation"], "Health Status"]]
                .dropna()
                .groupby([cols["occupation"], "Health Status"], observed=False)
                .size()
                .reset_index(name="Count")
            )

            # Keep only top 15 professions by total count
            rank = (
                occ_status.groupby(cols["occupation"], observed=False)["Count"]
                .sum()
                .sort_values(ascending=False)
                .head(15)
                .index
            )
            occ_status_top = occ_status[occ_status[cols["occupation"]].isin(rank)]

            fig_occ2 = px.bar(
                occ_status_top,
                x=cols["occupation"],
                y="Count",
                color="Health Status",
                barmode="group",
                title="Health status distribution by profession (Top 15)",
                text="Count",
                color_discrete_sequence=["#138808", "#FF9933", "#00A3E0"],
            )
            fig_occ2.update_layout(
                xaxis_title="Profession",
                yaxis_title="Number of persons",
                xaxis_tickangle=-65,
            )
            fig_occ2.update_traces(textposition="outside")

            st.plotly_chart(fig_occ2, width="stretch")
