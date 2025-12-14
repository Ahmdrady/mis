from __future__ import annotations

import pandas as pd
import plotly.express as px
import streamlit as st

from app import DashboardState
from src.components.export import render_export_button

GROUP_OPTIONS = {
    "Region": "region",
    "Year": "year",
    "Country": "REF_AREA_LABEL",
}

AGG_FUNCS = {
    "Average": "mean",
    "Median": "median",
    "Max": "max",
    "Min": "min",
}


def render(state: DashboardState) -> None:
    st.title("Data Lab")
    df = state.data.copy()

    if df.empty:
        st.warning("No data loaded with the current filters.")
        return

    st.caption("Ad-hoc slicing and multi-metric downloads for analysts.")
    column = st.selectbox("Group by", options=list(GROUP_OPTIONS.keys()))
    agg_label = st.selectbox("Aggregation", options=list(AGG_FUNCS.keys()))

    grouped = (
        df.groupby(GROUP_OPTIONS[column])["OBS_VALUE"]
        .agg(AGG_FUNCS[agg_label])
        .reset_index()
        .rename(columns={GROUP_OPTIONS[column]: column, "OBS_VALUE": f"{agg_label} Inflation"})
    )
    fig = px.bar(
        grouped.sort_values(f"{agg_label} Inflation", ascending=False),
        x=column,
        y=f"{agg_label} Inflation",
        template="plotly_white",
    )
    st.plotly_chart(fig, use_container_width=True)

    st.success(
        f"Insight: {grouped.sort_values(f'{agg_label} Inflation', ascending=False).iloc[0][column]} "
        f"leads on {agg_label.lower()} inflation."
    )

    st.markdown("#### Raw Data Preview")
    st.dataframe(df, hide_index=True, use_container_width=True)

    render_export_button(
        df,
        metadata={
            "Grouping": column,
            "Aggregation": agg_label,
            "Records": str(len(df)),
        },
    )
