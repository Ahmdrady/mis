from __future__ import annotations

from typing import List

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from app import DashboardState
from src.components.export import render_export_button

GLOBAL_EVENT_LIBRARY = {
    "2008-09-01": "Global financial crisis peak",
    "2011-08-01": "Commodity supercycle crest",
    "2014-12-01": "Oil price collapse",
    "2020-04-01": "COVID-19 supply chain shock",
    "2022-03-01": "Ukraine conflict escalation",
}


def render(state: DashboardState) -> None:
    st.title("Trend Explorer")
    df = state.data.copy()

    if df.empty:
        st.warning("Adjust the filters — no records to analyze.")
        return

    st.caption(
        "Explore cross-country trajectories, percentile spreads, and annotate trends for contextual storytelling."
    )

    focus_countries = build_country_selector(df)
    indexed_view = st.toggle("Show indexed trend (baseline = 100)", value=False)
    smoothing_window = st.select_slider(
        "Smoothing window (months)",
        options=[1, 3, 6, 12],
        value=3,
    )

    analysis_df = df[df["REF_AREA_LABEL"].isin(focus_countries)]
    if analysis_df.empty:
        st.warning("Selection returned no rows. Choose at least one country.")
        return

    chart_df = prepare_chart_frame(analysis_df, indexed=indexed_view, window=smoothing_window)
    render_multi_country_chart(chart_df, indexed=indexed_view)
    render_spread_band(df, smoothing_window)
    render_change_table(analysis_df)

    metadata = {
        "Countries": ", ".join(focus_countries),
        "Records": str(len(analysis_df)),
        "Date Range": f"{analysis_df['TIME_PERIOD'].min():%Y-%m} – {analysis_df['TIME_PERIOD'].max():%Y-%m}",
    }
    render_export_button(analysis_df, metadata=metadata)


def build_country_selector(df: pd.DataFrame) -> List[str]:
    latest_period = df["TIME_PERIOD"].max()
    latest_slice = (
        df[df["TIME_PERIOD"] == latest_period]
        .groupby("REF_AREA_LABEL")["OBS_VALUE"]
        .mean()
        .sort_values(ascending=False)
    )
    available = sorted(df["REF_AREA_LABEL"].unique())
    default_selection = list(latest_slice.head(min(5, len(latest_slice))).index)

    return st.multiselect(
        "Focus Countries",
        options=available,
        default=default_selection or available[:5],
    )


def prepare_chart_frame(df: pd.DataFrame, *, indexed: bool, window: int) -> pd.DataFrame:
    working = df.sort_values(["REF_AREA_LABEL", "TIME_PERIOD"]).copy()

    if window > 1:
        working["OBS_VALUE"] = (
            working.groupby("REF_AREA_LABEL")["OBS_VALUE"]
            .transform(lambda s: s.rolling(window=window, min_periods=1).mean())
        )

    if indexed:
        working["Indexed"] = (
            working.groupby("REF_AREA_LABEL")["OBS_VALUE"]
            .transform(lambda s: (s / s.iloc[0]) * 100 if s.iloc[0] else 100)
        )
    else:
        working["Indexed"] = working["OBS_VALUE"]

    return working


def render_multi_country_chart(df: pd.DataFrame, *, indexed: bool) -> None:
    y_field = "Indexed" if indexed else "OBS_VALUE"
    y_label = "Indexed (baseline=100)" if indexed else "Inflation (%)"

    fig = px.line(
        df,
        x="TIME_PERIOD",
        y=y_field,
        color="REF_AREA_LABEL",
        markers=False,
        template="plotly_white",
        labels={"TIME_PERIOD": "Month", y_field: y_label, "REF_AREA_LABEL": "Country"},
    )

    for event_date, label in GLOBAL_EVENT_LIBRARY.items():
        event_ts = pd.Timestamp(event_date)
        if df["TIME_PERIOD"].min() <= event_ts <= df["TIME_PERIOD"].max():
            fig.add_vline(
                x=event_ts,
                line_dash="dot",
                line_color="rgba(200,30,30,0.6)",
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

    fig.update_layout(
        height=500,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig, use_container_width=True)
    if not df.empty:
        latest_vals = df[df["TIME_PERIOD"] == df["TIME_PERIOD"].max()]
        if not latest_vals.empty:
            leader = latest_vals.sort_values(y_field, ascending=False).iloc[0]
            st.info(
                f"Benefit: {leader['REF_AREA_LABEL']} currently leads the cohort at "
                f"{leader[y_field]:.2f} {y_label.lower()}, signalling a primary outlier."
            )


def render_spread_band(df: pd.DataFrame, window: int) -> None:
    st.markdown("#### Percentile Spread")
    grouped = df.groupby("TIME_PERIOD")["OBS_VALUE"]
    spread = grouped.quantile([0.1, 0.25, 0.5, 0.75, 0.9]).unstack()
    spread.columns = ["p10", "p25", "median", "p75", "p90"]

    if window > 1:
        spread = spread.rolling(window=window, min_periods=1).mean()

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=spread.index,
            y=spread["p90"],
            mode="lines",
            line=dict(width=0),
            showlegend=False,
        )
    )
    fig.add_trace(
        go.Scatter(
            x=spread.index,
            y=spread["p10"],
            mode="lines",
            fill="tonexty",
            fillcolor="rgba(31,119,180,0.2)",
            line=dict(width=0),
            name="P10–P90 band",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=spread.index,
            y=spread["median"],
            mode="lines",
            line=dict(color="#d62728", width=2),
            name="Median",
        )
    )
    fig.update_layout(
        template="plotly_white",
        height=350,
        yaxis_title="Inflation (%)",
        xaxis_title="Month",
    )
    st.plotly_chart(fig, use_container_width=True)
    band_width = (spread["p90"] - spread["p10"]).mean()
    st.success(
        f"Insight: The average gap between p90 and p10 across the window is {band_width:.2f} pts, "
        "quantifying cross-country dispersion."
    )


def render_change_table(df: pd.DataFrame) -> None:
    st.markdown("#### Month-over-Month Movers (Selection)")

    latest_period = df["TIME_PERIOD"].max()
    prev_period = latest_period - pd.DateOffset(months=1)

    pivot = (
        df[df["TIME_PERIOD"].isin([latest_period, prev_period])]
        .groupby(["REF_AREA_LABEL", "TIME_PERIOD"])["OBS_VALUE"]
        .mean()
        .unstack()
    )

    if prev_period not in pivot.columns:
        st.info("Need at least one previous month to compute movers.")
        return

    pivot["Δ MoM"] = pivot[latest_period] - pivot[prev_period]
    pivot = pivot.reset_index().rename(
        columns={
            "REF_AREA_LABEL": "Country",
            latest_period: f"{latest_period:%Y-%m}",
            prev_period: f"{prev_period:%Y-%m}",
        }
    )

    st.dataframe(
        pivot.sort_values("Δ MoM", ascending=False),
        use_container_width=True,
        hide_index=True,
    )
    if not pivot.empty:
        surge = pivot.sort_values("Δ MoM", ascending=False).iloc[0]
        cool = pivot.sort_values("Δ MoM").iloc[0]
        st.caption(
            f"Insight: {surge['Country']} posted the sharpest monthly climb (+{surge['Δ MoM']:.2f} pts) "
            f"while {cool['Country']} cooled the most ({cool['Δ MoM']:.2f} pts)."
        )
