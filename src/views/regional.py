from __future__ import annotations

from typing import List

import pandas as pd
import plotly.express as px
import streamlit as st

from app import DashboardState
from src.components.metrics import render_metric_row
from src.data.region_map import REGION_DISPLAY_ORDER, REGION_FALLBACK

HIGH_INFLATION_THRESHOLD = 10.0


def render(state: DashboardState) -> None:
    st.title("Regional Insights")
    df = state.data.copy()

    if "region" not in df.columns:
        st.error("Region metadata missing. Re-run the pipeline to enrich the dataset.")
        return

    region_choice = region_selector(df)
    region_df = df if region_choice == "All Regions" else df[df["region"] == region_choice]

    if region_df.empty:
        st.warning("No records for the selected region/date filters.")
        return

    latest_period = region_df["TIME_PERIOD"].max()
    prev_period = latest_period - pd.DateOffset(months=1)
    trailing_window_start = latest_period - pd.DateOffset(months=12)

    render_metric_row(
        build_region_kpis(
            region_df=region_df,
            global_df=df,
            latest_period=latest_period,
            prev_period=prev_period,
            trailing_start=trailing_window_start,
            region_label=region_choice,
        )
    )

    tab_charts, tab_anomalies, tab_table = st.tabs(
        ["Stacked View", "Anomaly Radar", "Regional Table"]
    )

    with tab_charts:
        render_region_stack(region_df, df, region_choice, latest_period)

    with tab_anomalies:
        render_anomaly_panels(region_df, latest_period, trailing_window_start)

    with tab_table:
        render_region_table(region_df, latest_period, prev_period)


def region_selector(df: pd.DataFrame) -> str:
    options: List[str] = ["All Regions"]
    options += [region for region in REGION_DISPLAY_ORDER if region in df["region"].unique()]

    leftovers = sorted(
        {region for region in df["region"].unique() if region not in options}
    )
    options += leftovers

    return st.selectbox("Focus Region", options=options, index=0 if "All Regions" in options else 0)


def build_region_kpis(
    *,
    region_df: pd.DataFrame,
    global_df: pd.DataFrame,
    latest_period: pd.Timestamp,
    prev_period: pd.Timestamp,
    trailing_start: pd.Timestamp,
    region_label: str,
):
    display_label = region_label if region_label != "All Regions" else "Global"

    region_latest = region_df[region_df["TIME_PERIOD"] == latest_period]
    global_latest = global_df[global_df["TIME_PERIOD"] == latest_period]

    region_avg = region_latest["OBS_VALUE"].mean()
    global_avg = global_latest["OBS_VALUE"].mean()
    prev_avg = region_df.loc[region_df["TIME_PERIOD"] == prev_period, "OBS_VALUE"].mean()

    hotspot_share = share_hotspots(region_latest)
    hotspot_prev = share_hotspots(
        region_df[region_df["TIME_PERIOD"] == prev_period]
    )

    trailing = region_df[region_df["TIME_PERIOD"] >= trailing_start]
    vol = trailing.groupby("TIME_PERIOD")["OBS_VALUE"].mean().std()

    kpis = [
        {
            "label": f"{display_label} Avg ({latest_period:%b %Y})",
            "value": format_value(region_avg),
            "delta": format_delta(region_avg - global_avg, tail="vs global"),
            "delta_color": delta_color(region_avg - global_avg),
        },
        {
            "label": "Momentum",
            "value": format_value(region_avg - prev_avg, suffix=" pp"),
            "delta": format_delta(region_avg - prev_avg, tail="MoM change"),
            "delta_color": delta_color(region_avg - prev_avg),
        },
        {
            "label": f"Hotspots ≥ {HIGH_INFLATION_THRESHOLD:.0f}%",
            "value": format_value(hotspot_share, suffix="%", precision=1),
            "delta": format_delta(hotspot_share - hotspot_prev, suffix=" pts"),
            "delta_color": delta_color(hotspot_share - hotspot_prev),
        },
        {
            "label": "12M Volatility",
            "value": format_value(vol, suffix=" pp"),
            "delta": "std-dev of 12M rolling mean",
            "delta_color": "normal",
        },
    ]
    return kpis


def render_region_stack(
    region_df: pd.DataFrame,
    global_df: pd.DataFrame,
    region_choice: str,
    latest_period: pd.Timestamp,
) -> None:
    st.markdown("#### Regional Stack")
    if region_choice == "All Regions":
        plot_df = (
            global_df[global_df["TIME_PERIOD"] == latest_period]
            .groupby("region")["OBS_VALUE"]
            .mean()
            .reset_index()
            .sort_values("OBS_VALUE", ascending=False)
        )
        fig = px.bar(
            plot_df,
            x="region",
            y="OBS_VALUE",
            text="OBS_VALUE",
            labels={"OBS_VALUE": "Inflation (%)", "region": "Region"},
            title=f"Regional Averages – {latest_period:%b %Y}",
        )
    else:
        plot_df = (
            region_df[region_df["TIME_PERIOD"] == latest_period]
            .groupby("REF_AREA_LABEL")["OBS_VALUE"]
            .mean()
            .reset_index()
            .sort_values("OBS_VALUE", ascending=False)
            .head(15)
        )
        fig = px.bar(
            plot_df,
            x="OBS_VALUE",
            y="REF_AREA_LABEL",
            orientation="h",
            labels={"OBS_VALUE": "Inflation (%)", "REF_AREA_LABEL": "Country"},
            title=f"Top Contributors – {region_choice}",
        )

    fig.update_layout(height=400, template="plotly_white")
    st.plotly_chart(fig, use_container_width=True)

    if not plot_df.empty:
        if region_choice == "All Regions":
            top_row = plot_df.iloc[0]
            st.caption(
                f"Insight: {top_row['region']} leads the world snapshot at {top_row['OBS_VALUE']:.2f}%."
            )
        else:
            top_row = plot_df.iloc[0]
            st.caption(
                f"Benefit: Within {region_choice}, {top_row['REF_AREA_LABEL']} contributes the highest pressure "
                f"({top_row['OBS_VALUE']:.2f}%)."
            )


def render_anomaly_panels(
    region_df: pd.DataFrame,
    latest_period: pd.Timestamp,
    trailing_start: pd.Timestamp,
) -> None:
    st.markdown("#### Anomaly Radar")
    trailing = region_df[region_df["TIME_PERIOD"] >= trailing_start].copy()
    if trailing.empty:
        st.info("Not enough history to compute anomalies.")
        return

    stats = (
        trailing.groupby("REF_AREA_LABEL")["OBS_VALUE"]
        .agg(["mean", "std"])
        .reset_index()
    )
    latest = region_df[region_df["TIME_PERIOD"] == latest_period][
        ["REF_AREA_LABEL", "OBS_VALUE", "region"]
    ]
    merged = latest.merge(stats, on="REF_AREA_LABEL", how="left")
    merged["z_score"] = (merged["OBS_VALUE"] - merged["mean"]) / merged["std"]
    merged = merged.replace([pd.NA, pd.NaT], None).dropna(subset=["z_score"])

    hotspots = merged[merged["z_score"] >= 2].sort_values("z_score", ascending=False)
    cooling = merged[merged["z_score"] <= -2].sort_values("z_score", ascending=True)

    col_hot, col_cool = st.columns(2)
    with col_hot:
        st.caption("Hotspots (z ≥ 2)")
        if hotspots.empty:
            st.write("No significant spikes.")
        else:
            st.dataframe(
                hotspots[
                    [
                        "REF_AREA_LABEL",
                        "region",
                        "OBS_VALUE",
                        "z_score",
                    ]
                ],
                hide_index=True,
                use_container_width=True,
            )

    with col_cool:
        st.caption("Cooling Signals (z ≤ -2)")
        if cooling.empty:
            st.write("No significant cool-downs.")
        else:
            st.dataframe(
                cooling[
                    [
                        "REF_AREA_LABEL",
                        "region",
                        "OBS_VALUE",
                        "z_score",
                    ]
                ],
                hide_index=True,
                use_container_width=True,
            )
    total_hot = len(hotspots)
    total_cool = len(cooling)
    st.info(
        f"Insight: Detector flagged {total_hot} hot spikes and {total_cool} cool-downs over the last 12 months."
    )


def render_region_table(
    region_df: pd.DataFrame,
    latest_period: pd.Timestamp,
    prev_period: pd.Timestamp,
) -> None:
    st.markdown("#### Region Drill Table")
    yoy_period = latest_period - pd.DateOffset(years=1)

    latest = region_df[region_df["TIME_PERIOD"] == latest_period][
        ["REF_AREA_LABEL", "region", "OBS_VALUE"]
    ].rename(columns={"OBS_VALUE": f"{latest_period:%Y-%m}"})
    prev = region_df[region_df["TIME_PERIOD"] == prev_period][
        ["REF_AREA_LABEL", "OBS_VALUE"]
    ].rename(columns={"OBS_VALUE": "Prev"})
    yoy = region_df[region_df["TIME_PERIOD"] == yoy_period][
        ["REF_AREA_LABEL", "OBS_VALUE"]
    ].rename(columns={"OBS_VALUE": "YoY Base"})

    table = latest.merge(prev, on="REF_AREA_LABEL", how="left").merge(
        yoy, on="REF_AREA_LABEL", how="left"
    )
    table["MoM Δ"] = table[f"{latest_period:%Y-%m}"] - table["Prev"]
    table["YoY Δ"] = table[f"{latest_period:%Y-%m}"] - table["YoY Base"]
    table = table.sort_values(f"{latest_period:%Y-%m}", ascending=False).fillna(0)

    st.dataframe(
        table[
            [
                "REF_AREA_LABEL",
                "region",
                f"{latest_period:%Y-%m}",
                "Prev",
                "MoM Δ",
                "YoY Δ",
            ]
        ],
        hide_index=True,
        use_container_width=True,
    )
    if not table.empty:
        top = table.iloc[0]
        st.caption(
            f"Benefit: {top['REF_AREA_LABEL']} leads regional inflation at {top[f'{latest_period:%Y-%m}']:.2f}% "
            f"with MoM Δ {top['MoM Δ']:+.2f} pts."
        )


def share_hotspots(df: pd.DataFrame) -> float:
    if df.empty:
        return float("nan")
    total = df["REF_AREA"].nunique()
    if total == 0:
        return float("nan")
    hotspots = df[df["OBS_VALUE"] >= HIGH_INFLATION_THRESHOLD]["REF_AREA"].nunique()
    return round((hotspots / total) * 100, 2)


def format_value(value: float, suffix: str = "%", precision: int = 2) -> str:
    if pd.isna(value):
        return "—"
    return f"{value:.{precision}f}{suffix}"


def format_delta(value: float, suffix: str = "%", tail: str = "") -> str:
    if pd.isna(value):
        return "n/a"
    base = f"{value:+.2f}{suffix}"
    return f"{base} {tail}".strip()


def delta_color(value: float) -> str:
    if pd.isna(value) or value == 0:
        return "off"
    return "inverse" if value < 0 else "normal"
