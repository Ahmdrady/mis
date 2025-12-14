from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from app import DashboardState
from src.components.analytics import country_pivot, summarize_distribution


def render(state: DashboardState) -> None:
    st.title("Market Depth")
    df = state.data.copy()

    if df.empty:
        st.warning("No records under current filters.")
        return

    latest_period = df["TIME_PERIOD"].max()
    st.caption("Correlation, clustering, and distribution diagnostics for the selected slice.")

    focus_countries = st.multiselect(
        "Correlation focus (max 25)",
        options=sorted(df["REF_AREA_LABEL"].unique()),
        default=sorted(df["REF_AREA_LABEL"].unique())[:12],
    )
    if not focus_countries:
        st.info("Select at least one country.")
        return

    pivot = country_pivot(df[df["REF_AREA_LABEL"].isin(focus_countries)])
    corr = pivot.corr().round(2).fillna(0)
    render_corr_heatmap(corr)

    render_profile_scatter(df, latest_period)

    render_distribution_panel(df, latest_period)


def render_corr_heatmap(corr: pd.DataFrame) -> None:
    st.markdown("#### Correlation Grid")
    fig = px.imshow(
        corr,
        color_continuous_scale="RdBu",
        zmin=-1,
        zmax=1,
        labels=dict(color="Correlation"),
        aspect="auto",
    )
    fig.update_layout(height=600)
    st.plotly_chart(fig, use_container_width=True)

    strongest = corr.replace(1.0, np.nan).stack().sort_values(ascending=False)
    if not strongest.empty:
        pair = strongest.index[0]
        value = strongest.iloc[0]
        st.success(
            f"Insight: {pair[0]} and {pair[1]} move together with a correlation of {value:.2f}, "
            "suggesting shared inflation drivers."
        )


def render_profile_scatter(df: pd.DataFrame, latest_period: pd.Timestamp) -> None:
    st.markdown("#### Level vs Volatility Map")
    window = st.select_slider("Volatility window (months)", options=[3, 6, 12, 24], value=12)
    rolling = (
        df.sort_values(["REF_AREA_LABEL", "TIME_PERIOD"])
        .groupby("REF_AREA_LABEL")["OBS_VALUE"]
        .rolling(window=window, min_periods=3)
        .std()
        .reset_index()
        .rename(columns={"OBS_VALUE": "volatility"})
    )
    latest_vol = rolling.groupby("REF_AREA_LABEL").tail(1).set_index("REF_AREA_LABEL")
    latest_slice = (
        df[df["TIME_PERIOD"] == latest_period][["REF_AREA_LABEL", "OBS_VALUE", "region"]]
        .set_index("REF_AREA_LABEL")
    )
    merged = latest_slice.join(latest_vol, how="inner").dropna()

    working = merged.reset_index()
    working["bubble_size"] = working["OBS_VALUE"].abs() + 1

    fig = px.scatter(
        working,
        x="volatility",
        y="OBS_VALUE",
        color="region",
        hover_name="REF_AREA_LABEL",
        size="bubble_size",
        labels={"OBS_VALUE": "Latest Inflation (%)", "volatility": f"{window}M σ"},
        template="plotly_white",
    )
    fig.update_layout(height=450)
    st.plotly_chart(fig, use_container_width=True)

    if not merged.empty:
        hottest = merged["OBS_VALUE"].idxmax()
        st.info(
            f"Benefit: {hottest} leads in inflation level while volatility at "
            f"{merged.loc[hottest, 'volatility']:.2f}, highlighting monitoring priority."
        )


def render_distribution_panel(df: pd.DataFrame, latest_period: pd.Timestamp) -> None:
    st.markdown("#### Distribution Inspector")
    stats = summarize_distribution(df, latest_period)
    if not stats:
        st.info("No stats available for the latest period.")
        return

    fig = go.Figure()
    hist_data = df[df["TIME_PERIOD"] == latest_period]["OBS_VALUE"]
    fig.add_trace(
        go.Histogram(x=hist_data, nbinsx=40, marker_color="#1f77b4", opacity=0.75)
    )
    for label in ["median", "p25", "p75"]:
        fig.add_vline(
            x=stats[label],
            line_dash="dot",
            annotation_text=label.upper(),
            annotation_position="top left",
        )
    fig.update_layout(
        template="plotly_white",
        xaxis_title="Inflation (%)",
        yaxis_title="Countries",
        height=400,
    )
    st.plotly_chart(fig, use_container_width=True)

    st.success(
        f"Insight: Spread runs from {stats['p10']:.1f}% (p10) to {stats['p90']:.1f}% (p90). "
        f"The median sits at {stats['median']:.1f}%, signalling the typical peer level."
    )
