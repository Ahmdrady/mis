from __future__ import annotations

import pandas as pd
import plotly.express as px
import streamlit as st

from app import DashboardState


METRIC_OPTIONS = {
    "Average Inflation": ("mean", "%"),
    "Median Inflation": ("median", "%"),
    "Volatility (σ)": ("std", " pts"),
    "Max Spike": ("max", "%"),
}


def render(state: DashboardState) -> None:
    st.title("Era Explorer")
    df = state.data.copy()

    if df.empty:
        st.warning("Dataset empty. Re-run ETL.")
        return

    selected_metric = st.selectbox("Metric", options=list(METRIC_OPTIONS.keys()), index=0)
    agg_func, suffix = METRIC_OPTIONS[selected_metric]

    era_stats = (
        df.groupby("era")["OBS_VALUE"]
        .agg(
            mean="mean",
            median="median",
            std="std",
            max="max",
            min="min",
        )
        .reset_index()
    )
    value_column = agg_func
    era_stats["value"] = era_stats[value_column]

    render_metric_chart(era_stats, selected_metric, suffix)
    render_era_detail(df)
    render_country_rank(df)


def render_metric_chart(era_stats: pd.DataFrame, metric_name: str, suffix: str) -> None:
    fig = px.bar(
        era_stats.sort_values("value", ascending=False),
        x="era",
        y="value",
        text=era_stats["value"].apply(lambda v: f"{v:.2f}{suffix}"),
        labels={"era": "Era", "value": metric_name},
        template="plotly_white",
    )
    fig.update_traces(textposition="outside")
    fig.update_layout(height=420)
    st.plotly_chart(fig, use_container_width=True)

    leader = era_stats.sort_values("value", ascending=False).iloc[0]
    st.info(
        f"Insight: {leader['era']} dominates on {metric_name.lower()} "
        f"with {leader['value']:.2f}{suffix}."
    )


def render_era_detail(df: pd.DataFrame) -> None:
    st.markdown("#### Era Trajectories")
    era_trend = (
        df.groupby(["era", "TIME_PERIOD"])["OBS_VALUE"]
        .mean()
        .reset_index()
        .rename(columns={"OBS_VALUE": "Avg"})
    )
    fig = px.line(
        era_trend,
        x="TIME_PERIOD",
        y="Avg",
        color="era",
        template="plotly_white",
        labels={"Avg": "Inflation (%)", "TIME_PERIOD": "Month", "era": "Era"},
    )
    fig.update_layout(height=450)
    st.plotly_chart(fig, use_container_width=True)

    st.caption("Benefit: Compare how each structural period evolved over time without touching any filters.")


def render_country_rank(df: pd.DataFrame) -> None:
    st.markdown("#### Era Leaderboard")
    ranking = (
        df.groupby(["era", "REF_AREA_LABEL"])["OBS_VALUE"]
        .mean()
        .reset_index()
        .rename(columns={"OBS_VALUE": "Avg"})
    )
    top = (
        ranking.sort_values(["era", "Avg"], ascending=[True, False])
        .groupby("era")
        .head(5)
    )
    st.dataframe(
        top.rename(columns={"REF_AREA_LABEL": "Country", "Avg": "Average (%)"}),
        hide_index=True,
        use_container_width=True,
    )
    st.caption("Insight: Every era lists its top five persistently high-inflation countries for quick reference.")
