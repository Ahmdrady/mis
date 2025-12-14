from __future__ import annotations

import pandas as pd
import plotly.express as px
import streamlit as st

from app import DashboardState


def render(state: DashboardState) -> None:
    st.title("Risk Monitor")
    df = state.data.copy()

    if df.empty:
        st.warning("Dataset empty.")
        return

    scores = compute_risk_scores(df)
    render_watchlist(scores)
    render_scatter(scores)
    render_resilience(scores)


def compute_risk_scores(df: pd.DataFrame) -> pd.DataFrame:
    grouped = (
        df.groupby("REF_AREA_LABEL")["OBS_VALUE"]
        .agg(
            mean="mean",
            volatility="std",
            p95=lambda s: s.quantile(0.95),
            p05=lambda s: s.quantile(0.05),
        )
        .reset_index()
        .rename(columns={"REF_AREA_LABEL": "Country"})
    )
    grouped["range"] = grouped["p95"] - grouped["p05"]
    grouped["risk_score"] = normalize(grouped["mean"]) * 0.4 + normalize(grouped["volatility"]) * 0.35 + normalize(
        grouped["range"]
    ) * 0.25
    return grouped.sort_values("risk_score", ascending=False)


def normalize(series: pd.Series) -> pd.Series:
    if series.max() == series.min():
        return pd.Series(0, index=series.index)
    return (series - series.min()) / (series.max() - series.min())


def render_watchlist(scores: pd.DataFrame) -> None:
    st.markdown("#### High-Risk Countries")
    st.dataframe(
        scores.head(15)[["Country", "mean", "volatility", "range", "risk_score"]].rename(
            columns={
                "mean": "Mean (%)",
                "volatility": "σ",
                "range": "P95-P05",
                "risk_score": "Score",
            }
        ),
        hide_index=True,
        use_container_width=True,
    )
    leader = scores.iloc[0]
    st.info(
        f"Insight: {leader['Country']} faces the highest structural risk with "
        f"mean {leader['mean']:.2f}% and volatility {leader['volatility']:.2f}."
    )


def render_scatter(scores: pd.DataFrame) -> None:
    st.markdown("#### Risk Landscape")
    fig = px.scatter(
        scores,
        x="volatility",
        y="mean",
        size="risk_score",
        hover_name="Country",
        labels={"volatility": "Volatility (σ)", "mean": "Mean Inflation (%)", "risk_score": "Risk Score"},
        template="plotly_white",
    )
    fig.update_layout(height=450)
    st.plotly_chart(fig, use_container_width=True)
    st.caption("Benefit: Large bubbles in the upper-right quadrant highlight chronic, high-amplitude markets.")


def render_resilience(scores: pd.DataFrame) -> None:
    st.markdown("#### Resilience Board")
    resilient = scores.sort_values("risk_score").head(15)
    st.dataframe(
        resilient[["Country", "mean", "volatility", "range", "risk_score"]].rename(
            columns={
                "mean": "Mean (%)",
                "volatility": "σ",
                "range": "P95-P05",
                "risk_score": "Score",
            }
        ),
        hide_index=True,
        use_container_width=True,
    )
    st.caption("Insight: These countries combine low mean levels and tight distributions—ideal policy exemplars.")
