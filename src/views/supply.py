from __future__ import annotations

import pandas as pd
import plotly.express as px
import streamlit as st

from app import DashboardState

THRESHOLDS = [5, 10, 15, 20]


def render(state: DashboardState) -> None:
    st.title("Supply Watch")
    df = state.data.copy()

    if df.empty:
        st.warning("No records loaded.")
        return

    st.caption("Historical perspective on how often regions breach critical inflation thresholds.")

    render_threshold_tracker(df)
    render_persistence_panel(df)
    render_concentration_table(df)


def render_threshold_tracker(df: pd.DataFrame) -> None:
    threshold = st.select_slider("Threshold (%)", options=THRESHOLDS, value=10)
    share = (
        df.assign(flag=df["OBS_VALUE"] >= threshold)
        .groupby("region")["flag"]
        .mean()
        .reset_index()
        .rename(columns={"flag": "Share"})
        .sort_values("Share", ascending=False)
    )

    fig = px.bar(
        share,
        x="region",
        y="Share",
        labels={"region": "Region", "Share": f"Share of months ≥ {threshold}%"},
        text=share["Share"].apply(lambda v: f"{v*100:.1f}%"),
        template="plotly_white",
    )
    fig.update_traces(textposition="outside")
    fig.update_layout(height=400, yaxis_tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

    leader = share.iloc[0]
    st.info(
        f"Insight: {leader['region']} spent {leader['Share']*100:.1f}% of all months above {threshold}%."
    )


def render_persistence_panel(df: pd.DataFrame) -> None:
    st.markdown("#### Persistence Map")
    duration = (
        df.assign(flag=df["OBS_VALUE"] >= 10)
        .groupby("REF_AREA_LABEL")["flag"]
        .sum()
        .reset_index()
        .rename(columns={"flag": "Months ≥10%"})
        .sort_values("Months ≥10%", ascending=False)
        .head(30)
    )
    fig = px.bar(
        duration,
        x="Months ≥10%",
        y="REF_AREA_LABEL",
        orientation="h",
        labels={"REF_AREA_LABEL": "Country"},
        template="plotly_white",
        height=600,
    )
    st.plotly_chart(fig, use_container_width=True)

    if not duration.empty:
        st.caption(
            f"Benefit: {duration.iloc[0]['REF_AREA_LABEL']} logged the longest streak with "
            f"{duration.iloc[0]['Months ≥10%']} months above 10%."
        )


def render_concentration_table(df: pd.DataFrame) -> None:
    st.markdown("#### Concentration Table")
    summary = (
        df.groupby("REF_AREA_LABEL")
        .agg(
            region=("region", "max"),
            avg=("OBS_VALUE", "mean"),
            median=("OBS_VALUE", "median"),
            volatility=("OBS_VALUE", "std"),
            share_gt10=("OBS_VALUE", lambda s: (s >= 10).mean()),
        )
        .reset_index()
        .rename(columns={"REF_AREA_LABEL": "Country"})
    )
    summary["share_gt10"] = summary["share_gt10"] * 100
    st.dataframe(
        summary.sort_values("share_gt10", ascending=False).head(50),
        hide_index=True,
        use_container_width=True,
    )
    st.caption("Insight: Countries with high share_gt10 and volatility deserve supply-chain contingency plans.")
