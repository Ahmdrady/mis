from __future__ import annotations

from typing import Dict

import pandas as pd
import plotly.express as px
import streamlit as st

from app import DashboardState
from src.components.report_export import generate_inflation_report


def render(state: DashboardState) -> None:
    st.title("Global Storyboard")
    df = state.data.copy()

    if df.empty:
        st.warning("Dataset is empty. Run the pipeline first.")
        return

    render_global_kpis(df)
    render_era_table(df)
    render_global_timeline(df)
    render_record_board(df)
    render_export(df, state.filters)


def render_global_kpis(df: pd.DataFrame) -> None:
    first_date = df["TIME_PERIOD"].min()
    last_date = df["TIME_PERIOD"].max()
    span_years = last_date.year - first_date.year + 1

    metrics = [
        {
            "label": "Observations",
            "value": f"{len(df):,}",
            "tooltip": "Total monthly country data points ingested.",
        },
        {
            "label": "Countries",
            "value": df["REF_AREA"].nunique(),
            "tooltip": "Unique REF_AREA codes tracked.",
        },
        {
            "label": "Coverage",
            "value": f"{first_date:%Y} – {last_date:%Y}",
            "tooltip": f"{span_years} years of history.",
        },
        {
            "label": "Global Mean",
            "value": f"{df['OBS_VALUE'].mean():.2f}%",
            "tooltip": "Average inflation across the entire dataset.",
        },
        {
            "label": "Global σ",
            "value": f"{df['OBS_VALUE'].std():.2f} pts",
            "tooltip": "Standard deviation across all records.",
        },
    ]

    cols = st.columns(len(metrics))
    for col, metric in zip(cols, metrics):
        col.metric(metric["label"], metric["value"], help=metric["tooltip"])


def render_era_table(df: pd.DataFrame) -> None:
    st.markdown("### Era Benchmarks")
    era_rows = []
    for era, group in df.groupby("era"):
        top_idx = group["OBS_VALUE"].idxmax()
        bottom_idx = group["OBS_VALUE"].idxmin()
        era_rows.append(
            {
                "Era": era,
                "Span": f"{group['TIME_PERIOD'].min():%Y} – {group['TIME_PERIOD'].max():%Y}",
                "Avg (%)": round(group["OBS_VALUE"].mean(), 2),
                "Median (%)": round(group["OBS_VALUE"].median(), 2),
                "Volatility (σ)": round(group["OBS_VALUE"].std(), 2),
                "Peak Country": df.loc[top_idx, "REF_AREA_LABEL"] if pd.notna(top_idx) else "n/a",
                "Peak (%)": round(df.loc[top_idx, "OBS_VALUE"], 2) if pd.notna(top_idx) else "n/a",
                "Trough Country": df.loc[bottom_idx, "REF_AREA_LABEL"] if pd.notna(bottom_idx) else "n/a",
                "Trough (%)": round(df.loc[bottom_idx, "OBS_VALUE"], 2) if pd.notna(bottom_idx) else "n/a",
            }
        )

    era_df = pd.DataFrame(era_rows).sort_values("Avg (%)", ascending=False)
    st.dataframe(era_df, use_container_width=True, hide_index=True)

    hottest = era_df.iloc[0]
    st.info(
        f"Insight: {hottest['Era']} recorded the highest average inflation "
        f"({hottest['Avg (%)']:.2f}%) and sets the ceiling for stress testing."
    )


def render_global_timeline(df: pd.DataFrame) -> None:
    st.markdown("### Global Average Timeline")
    monthly = (
        df.groupby("TIME_PERIOD")["OBS_VALUE"]
        .mean()
        .reset_index()
        .rename(columns={"OBS_VALUE": "Global Avg"})
    )
    monthly["Rolling 12M"] = monthly["Global Avg"].rolling(window=12, min_periods=6).mean()

    fig = px.line(
        monthly,
        x="TIME_PERIOD",
        y=["Global Avg", "Rolling 12M"],
        labels={"value": "Inflation (%)", "TIME_PERIOD": "Month", "variable": "Series"},
        template="plotly_white",
    )
    fig.update_layout(height=450, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    st.plotly_chart(fig, use_container_width=True)

    latest = monthly.iloc[-1]
    st.success(
        f"Benefit: The latest rolling 12-month average is {latest['Rolling 12M']:.2f}% "
        f"compared with the long-run mean of {monthly['Global Avg'].mean():.2f}%."
    )


def render_record_board(df: pd.DataFrame) -> None:
    st.markdown("### Historical Records")
    top_spikes = df.sort_values("OBS_VALUE", ascending=False).head(10)[
        ["REF_AREA_LABEL", "TIME_PERIOD", "OBS_VALUE", "region"]
    ]
    deep_cool = df.sort_values("OBS_VALUE", ascending=True).head(10)[
        ["REF_AREA_LABEL", "TIME_PERIOD", "OBS_VALUE", "region"]
    ]

    col_spike, col_cool = st.columns(2)
    with col_spike:
        st.caption("Top 10 Monthly Spikes")
        st.dataframe(
            top_spikes.rename(columns={"REF_AREA_LABEL": "Country", "TIME_PERIOD": "Month", "OBS_VALUE": "Inflation (%)"}),
            hide_index=True,
            use_container_width=True,
        )
    with col_cool:
        st.caption("Deepest Deflation Prints")
        st.dataframe(
            deep_cool.rename(columns={"REF_AREA_LABEL": "Country", "TIME_PERIOD": "Month", "OBS_VALUE": "Inflation (%)"}),
            hide_index=True,
            use_container_width=True,
        )

    st.caption(
        "Insight: These record boards surface every extreme observation in the dataset, offering ready-to-use case studies."
    )


def render_export(df: pd.DataFrame, filters: Dict) -> None:
    st.markdown("### Export Full Historical Workbook")
    meta = {
        "report_title": "Global Inflation Intelligence Hub",
        "scope": filters.get("scope_label", "Global Storyboard"),
    }
    try:
        payload = generate_inflation_report(df, meta)
        st.download_button(
            "Download Excel Workbook",
            data=payload,
            file_name="global_inflation_intel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except ModuleNotFoundError as exc:
        st.error(str(exc))
