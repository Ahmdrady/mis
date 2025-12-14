from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st


@dataclass
class BreakPoint:
    country: str
    latest_avg: float
    prev_avg: float
    shift: float
    period_label: str


@st.cache_data(show_spinner=False)
def country_pivot(df: pd.DataFrame) -> pd.DataFrame:
    pivot = df.pivot_table(
        index="TIME_PERIOD",
        columns="REF_AREA_LABEL",
        values="OBS_VALUE",
        aggfunc="mean",
    ).sort_index()
    return pivot


def compute_volatility(df: pd.DataFrame, window: int = 12) -> pd.DataFrame:
    rolling = (
        df.sort_values(["REF_AREA_LABEL", "TIME_PERIOD"])
        .groupby("REF_AREA_LABEL")["OBS_VALUE"]
        .rolling(window=window, min_periods=3)
        .std()
        .reset_index()
    )
    rolling = rolling.rename(columns={"OBS_VALUE": "volatility"})
    return rolling


def compute_breaks(
    df: pd.DataFrame,
    latest_period: pd.Timestamp,
    window: int = 3,
) -> List[BreakPoint]:
    current_window = df[df["TIME_PERIOD"] <= latest_period]
    current_window = current_window[current_window["TIME_PERIOD"] > latest_period - pd.DateOffset(months=window)]
    prev_window = df[
        (df["TIME_PERIOD"] <= latest_period - pd.DateOffset(months=window))
        & (df["TIME_PERIOD"] > latest_period - pd.DateOffset(months=window * 2))
    ]

    results: List[BreakPoint] = []
    grouped_current = current_window.groupby("REF_AREA_LABEL")["OBS_VALUE"].mean()
    grouped_prev = prev_window.groupby("REF_AREA_LABEL")["OBS_VALUE"].mean()

    for country, current_avg in grouped_current.items():
        prev_avg = grouped_prev.get(country, np.nan)
        if pd.isna(prev_avg):
            continue
        shift = current_avg - prev_avg
        results.append(
            BreakPoint(
                country=country,
                latest_avg=current_avg,
                prev_avg=prev_avg,
                shift=shift,
                period_label=f"Δ{window}m",
            )
        )
    results.sort(key=lambda item: abs(item.shift), reverse=True)
    return results


def summarize_distribution(df: pd.DataFrame, latest_period: pd.Timestamp) -> Dict[str, float]:
    slice_df = df[df["TIME_PERIOD"] == latest_period]["OBS_VALUE"]
    if slice_df.empty:
        return {}
    return {
        "p10": slice_df.quantile(0.1),
        "p25": slice_df.quantile(0.25),
        "median": slice_df.quantile(0.5),
        "p75": slice_df.quantile(0.75),
        "p90": slice_df.quantile(0.9),
        "mean": slice_df.mean(),
        "std": slice_df.std(),
    }


def regional_share(df: pd.DataFrame, latest_period: pd.Timestamp, threshold: float) -> pd.DataFrame:
    slice_df = df[df["TIME_PERIOD"] == latest_period]
    if slice_df.empty:
        return pd.DataFrame(columns=["region", "share"])

    summary = (
        slice_df.assign(flag=slice_df["OBS_VALUE"] >= threshold)
        .groupby("region")
        .agg(
            countries=("REF_AREA", "nunique"),
            high=("flag", "sum"),
        )
        .reset_index()
    )
    summary["share"] = (summary["high"] / summary["countries"]).fillna(0) * 100
    return summary
