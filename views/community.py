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

    # -------------------------------------------------------------------------
    #  Prepare aggregated metrics
    # -------------------------------------------------------------------------
    agg = view.copy()
    
    # 1. Unhealthy status
    #    (If "Health Status" is missing, we can't compute this reliably, 
    #     but we assume it exists if the pipeline ran.)
    if "Health Status" in agg.columns:
        agg["is_unhealthy"] = (agg["Health Status"] != "Healthy").astype(int)
    else:
        agg["is_unhealthy"] = 0

    # 2. High glucose (Diabetes proxy)
    if cols.get("glucose"):
        # Ensure numeric
        g_series = pd.to_numeric(agg[cols["glucose"]], errors="coerce")
        agg["high_glu"] = (g_series >= 126).astype(int)
    else:
        agg["high_glu"] = 0

    # 3. Obesity (BMI >= 30)
    if cols.get("bmi"):
        b_series = pd.to_numeric(agg[cols["bmi"]], errors="coerce")
        agg["obese"] = (b_series >= 30).astype(int)
    else:
        agg["obese"] = 0

    # Group by locality
    # We need a count column. We can use any column that is never null, 
    # or just size(). Let's use "is_unhealthy" count as a proxy for row count.
    grp = (
        agg.groupby(loc_col)
        .agg(
            Population=("is_unhealthy", "count"),
            Unhealthy_rate=("is_unhealthy", "mean"),
            Diabetes_rate=("high_glu", "mean"),
            Obesity_rate=("obese", "mean"),
        )
        .reset_index()
    )

    # Convert fractions to percentages
    grp["Unhealthy_rate"] *= 100.0
    grp["Diabetes_rate"] *= 100.0
    grp["Obesity_rate"] *= 100.0

    # -------------------------------------------------------------------------
    #  Controls
    # -------------------------------------------------------------------------
    max_locs = min(50, len(grp))
    if max_locs < 5:
        default_val = max_locs
    else:
        default_val = min(25, max_locs)
        
    n_loc = st.slider(
        "Number of localities to display",
        min_value=min(5, max_locs) if max_locs > 0 else 0,
        max_value=max_locs,
        value=default_val,
        step=5,
        help="Affects both the population and risk charts below.",
    )

    if n_loc == 0:
        st.warning("Not enough localities to display.")
        return

    grp_top = grp.sort_values("Population", ascending=False).head(n_loc)

    # -------------------------------------------------------------------------
    #  Charts
    # -------------------------------------------------------------------------
    st.subheader(f"Top {n_loc} localities by population (filtered)")
    
    fig_pop = px.bar(
        grp_top,
        x=loc_col,
        y="Population",
        title="Top localities by number of records",
        text="Population",
        color="Population",
        color_continuous_scale=["#138808", "#FFFFFF", "#FF9933"],
    )
    fig_pop.update_traces(textposition="outside")
    fig_pop.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig_pop, width="stretch")

    st.subheader("Locality-wise risk indicators")
    fig_loc = px.bar(
        grp_top,
        x=loc_col,
        y=["Unhealthy_rate", "Diabetes_rate", "Obesity_rate"],
        barmode="group",
        title="Community risk comparison (Unhealthy / Diabetes / Obesity)",
        labels={"value": "Rate (%)", "variable": "Indicator"},
    )
    fig_loc.update_layout(
        xaxis_title="Locality",
        yaxis_title="Rate (%)",
        xaxis_tickangle=-65,
    )
    fig_loc.update_traces(texttemplate="%{y:.1f}")
    st.plotly_chart(fig_loc, width="stretch")

    st.subheader("Locality table (all)")
    st.dataframe(grp.sort_values("Population", ascending=False), width="stretch")
