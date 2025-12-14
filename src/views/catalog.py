from __future__ import annotations

import pandas as pd
import streamlit as st

from app import DashboardState
from src.components.report_export import generate_inflation_report


def render(state: DashboardState) -> None:
    st.title("Data Catalog")
    df = state.data.copy()

    if df.empty:
        st.warning("No data to display.")
        return

    st.caption("Complete history is available below with lightweight filters for quick lookup.")

    region_options = ["All Regions"] + sorted(df["region"].dropna().unique())
    region = st.selectbox("Region", options=region_options, index=0)
    search_country = st.text_input("Country contains", "")

    filtered = df
    if region != "All Regions":
        filtered = filtered[filtered["region"] == region]
    if search_country:
        mask = filtered["REF_AREA_LABEL"].str.contains(search_country, case=False, na=False)
        filtered = filtered[mask]

    st.metric("Rows Returned", f"{len(filtered):,}")
    st.metric("Date Span", f"{filtered['TIME_PERIOD'].min():%Y-%m} → {filtered['TIME_PERIOD'].max():%Y-%m}")

    st.dataframe(
        filtered[
            [
                "TIME_PERIOD",
                "REF_AREA",
                "REF_AREA_LABEL",
                "region",
                "OBS_VALUE",
                "era",
                "decade",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

    try:
        payload = generate_inflation_report(
            filtered,
            {
                "report_title": "Data Catalog Export",
                "scope": f"Region: {region} | Search: {search_country or 'All'}",
            },
        )
        st.download_button(
            "Download Excel Workbook",
            data=payload,
            file_name="inflation_catalog.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except ModuleNotFoundError as exc:
        st.error(str(exc))

    st.caption("Every data point is accessible without touching code—ideal for analysts who need the raw feed.")
