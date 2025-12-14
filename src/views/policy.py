from __future__ import annotations

import pandas as pd
import plotly.express as px
import streamlit as st

from app import DashboardState

STABILITY_THRESHOLD = 5.0


def render(state: DashboardState) -> None:
    st.title("Policy Ledger")
    df = state.data.copy()

    if df.empty:
        st.warning("Dataset empty.")
        return

    render_policy_kpis(df)
    render_decade_panel(df)
    render_change_feed(df)


def render_policy_kpis(df: pd.DataFrame) -> None:
    stable_share = (df["OBS_VALUE"] <= STABILITY_THRESHOLD).mean() * 100
    severe_share = (df["OBS_VALUE"] >= 15).mean() * 100
    avg_trend = df.groupby("year")["OBS_VALUE"].mean().diff().mean()

    cols = st.columns(3)
    cols[0].metric("Stable Months (≤5%)", f"{stable_share:.1f}%")
    cols[1].metric("Severe Months (≥15%)", f"{severe_share:.1f}%")
    cols[2].metric("Mean Annual Drift", f"{avg_trend:+.2f} pts/year")

    st.caption(
        "Insight: Stability is scarce—only a fraction of observations fall under 5%, while high-pressure months remain common."
    )


def render_decade_panel(df: pd.DataFrame) -> None:
    st.markdown("#### Decade Comparison")
    decade_summary = (
        df.groupby("decade")["OBS_VALUE"]
        .agg(avg="mean", median="median", max="max", min="min", volatility="std")
        .reset_index()
    )
    fig = px.line(
        decade_summary,
        x="decade",
        y=["avg", "median"],
        markers=True,
        labels={"value": "Inflation (%)", "decade": "Decade", "variable": "Statistic"},
        template="plotly_white",
    )
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)

    st.info(
        f"Benefit: {decade_summary.sort_values('avg', ascending=False).iloc[0]['decade']}s "
        "represent the toughest policy environment on record."
    )


def render_change_feed(df: pd.DataFrame) -> None:
    st.markdown("#### Historical Movers")
    first_year = df["year"].min()
    last_year = df["year"].max()

    start = (
        df[df["year"] == first_year]
        .groupby("REF_AREA_LABEL")["OBS_VALUE"]
        .mean()
        .reset_index()
        .rename(columns={"OBS_VALUE": "Start"})
    )
    end = (
        df[df["year"] == last_year]
        .groupby("REF_AREA_LABEL")["OBS_VALUE"]
        .mean()
        .reset_index()
        .rename(columns={"OBS_VALUE": "End"})
    )
    merged = start.merge(end, on="REF_AREA_LABEL", how="inner")
    merged["Change"] = merged["End"] - merged["Start"]

    improvements = merged.sort_values("Change").head(10)
    regressions = merged.sort_values("Change", ascending=False).head(10)

    col_left, col_right = st.columns(2)
    with col_left:
        st.caption(f"Most Improved ({first_year}→{last_year})")
        st.dataframe(improvements.rename(columns={"REF_AREA_LABEL": "Country"}), hide_index=True, use_container_width=True)
    with col_right:
        st.caption(f"Biggest Deteriorations ({first_year}→{last_year})")
        st.dataframe(regressions.rename(columns={"REF_AREA_LABEL": "Country"}), hide_index=True, use_container_width=True)

    st.caption(
        "These comparisons use the earliest vs latest years to surface structural policy successes and failures."
    )
