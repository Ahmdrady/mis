from __future__ import annotations

from typing import List

import pandas as pd
import plotly.express as px
import streamlit as st

from app import DashboardState
from src.components.analytics import compute_breaks
from src.views.trends import GLOBAL_EVENT_LIBRARY


def render(state: DashboardState) -> None:
    st.title("Shock Analyzer")
    df = state.data.copy()

    if df.empty:
        st.warning("No data available for analysis.")
        return

    latest_period = df["TIME_PERIOD"].max()
    lookback_window = st.select_slider(
        "Structural break window (months)",
        options=[3, 6, 9, 12],
        value=6,
    )

    render_event_panel(df, latest_period)
    render_break_table(df, latest_period, lookback_window)
    render_heatmap(df, latest_period)


def render_event_panel(df: pd.DataFrame, latest_period: pd.Timestamp) -> None:
    st.markdown("#### Event Timeline")
    countries = st.multiselect(
        "Countries to overlay",
        options=sorted(df["REF_AREA_LABEL"].unique()),
        default=sorted(df["REF_AREA_LABEL"].unique())[:3],
    )
    if not countries:
        st.info("Select at least one country to display timeline.")
        return

    subset = df[df["REF_AREA_LABEL"].isin(countries)]
    fig = px.line(
        subset,
        x="TIME_PERIOD",
        y="OBS_VALUE",
        color="REF_AREA_LABEL",
        template="plotly_white",
        labels={"OBS_VALUE": "Inflation (%)", "TIME_PERIOD": "Month"},
    )
    for event_date, label in GLOBAL_EVENT_LIBRARY.items():
        event_ts = pd.Timestamp(event_date)
        if subset["TIME_PERIOD"].min() <= event_ts <= subset["TIME_PERIOD"].max():
            fig.add_vline(
                x=event_ts,
                line_dash="dash",
                line_color="orange",
            )
            fig.add_annotation(
                x=event_ts,
                yref="paper",
                y=1.05,
                showarrow=False,
                text=label,
                bgcolor="rgba(255,255,255,0.8)",
                font=dict(size=10),
            )
    fig.update_layout(height=450)
    st.plotly_chart(fig, use_container_width=True)

    st.info(
        "Benefit: Overlaying macro events highlights which markets react sharply to shocks versus those that remain insulated."
    )


def render_break_table(df: pd.DataFrame, latest_period: pd.Timestamp, window: int) -> None:
    st.markdown("#### Structural Break Detector")
    breaks = compute_breaks(df, latest_period, window=window)
    if not breaks:
        st.info("Not enough trailing data to compute breakpoints.")
        return

    data = [
        {
            "Country": bp.country,
            "Latest Avg": bp.latest_avg,
            "Prev Avg": bp.prev_avg,
            "Shift": bp.shift,
        }
        for bp in breaks[:15]
    ]
    table = pd.DataFrame(data).sort_values("Shift", ascending=False)
    st.dataframe(table, hide_index=True, use_container_width=True)

    top = table.iloc[0]
    st.success(
        f"Insight: {top['Country']} shows the largest {window}‑month jump of {top['Shift']:.2f} pts, "
        "flagging potential supply pressure."
    )


def render_heatmap(df: pd.DataFrame, latest_period: pd.Timestamp) -> None:
    st.markdown("#### Month-on-Month Shock Heatmap")
    periods = st.slider(
        "Months to display",
        min_value=6,
        max_value=24,
        value=12,
    )
    recent = df[df["TIME_PERIOD"] >= latest_period - pd.DateOffset(months=periods)]
    pivot = recent.pivot_table(
        index="TIME_PERIOD",
        columns="region",
        values="OBS_VALUE",
        aggfunc="mean",
    ).sort_index()

    if pivot.empty:
        st.info("No data in the selected window.")
        return

    fig = px.imshow(
        pivot.T,
        aspect="auto",
        labels=dict(color="Inflation (%)", x="Region", y="Month"),
        color_continuous_scale="Turbo",
    )
    st.plotly_chart(fig, use_container_width=True)

    volatile_month = pivot.mean(axis=1).diff().abs().idxmax()
    st.warning(
        f"Insight: {volatile_month:%Y-%m} delivered the sharpest cross-region swing, "
        "worth investigating for root causes."
    )
