from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Tuple

import pandas as pd
import streamlit as st

try:
    import tomllib  # Python 3.11+
except ModuleNotFoundError:  # pragma: no cover
    import tomli as tomllib  # type: ignore


CONFIG_PATH = Path("config") / "settings.toml"


@dataclass
class DashboardState:
    data: pd.DataFrame
    config: Dict
    filters: Dict


@st.cache_data(show_spinner=False)
def load_config(path: Path = CONFIG_PATH) -> Dict:
    if not path.exists():
        raise FileNotFoundError("Missing config/settings.toml. Run pipeline setup.")
    with path.open("rb") as fh:
        return tomllib.load(fh)


@st.cache_data(show_spinner=True)
def load_dataset(processed_path: Path) -> pd.DataFrame:
    from src.components.data import load_processed_data

    return load_processed_data(processed_path)


def load_views() -> Dict[str, Callable[[DashboardState], None]]:
    from src.views import (
        catalog,
        datalab,
        era,
        executive,
        market,
        policy,
        regional,
        risk,
        shock,
        supply,
        trends,
    )

    return {
        "Global Story": executive.render,
        "Era Explorer": era.render,
        "Trend Explorer": trends.render,
        "Regional Atlas": regional.render,
        "Market Depth": market.render,
        "Shock Timeline": shock.render,
        "Supply Watch": supply.render,
        "Policy Ledger": policy.render,
        "Risk Monitor": risk.render,
        "Data Lab": datalab.render,
        "Data Catalog": catalog.render,
    }


def get_active_view(view_registry: Dict[str, Callable]) -> Callable:
    st.sidebar.title("Navigation")
    return view_registry[
        st.sidebar.radio(
            "",
            options=list(view_registry.keys()),
            index=0,
        )
    ]


def apply_global_controls(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict]:
    st.sidebar.title("Global Controls")

    min_period = df["TIME_PERIOD"].min()
    max_period = df["TIME_PERIOD"].max()
    default_range = (min_period.to_pydatetime(), max_period.to_pydatetime())
    date_selection = st.sidebar.date_input(
        "Historical Window",
        value=default_range,
        min_value=min_period.to_pydatetime(),
        max_value=max_period.to_pydatetime(),
    )
    if isinstance(date_selection, tuple) and len(date_selection) == 2:
        start_date, end_date = date_selection
    else:
        start_date = end_date = date_selection
    if start_date > end_date:
        start_date, end_date = end_date, start_date
    start_ts = pd.Timestamp(start_date)
    end_ts = pd.Timestamp(end_date) + pd.offsets.MonthEnd(0)

    filtered = df[(df["TIME_PERIOD"] >= start_ts) & (df["TIME_PERIOD"] <= end_ts)].copy()

    st.sidebar.markdown("### Dimension Filters")
    region_options = sorted(filtered["region"].dropna().unique().tolist())
    selected_regions = st.sidebar.multiselect(
        "Regions",
        options=region_options,
        default=region_options,
    )
    if selected_regions:
        filtered = filtered[filtered["region"].isin(selected_regions)]

    era_options = sorted(filtered["era"].dropna().unique().tolist())
    selected_eras = st.sidebar.multiselect(
        "Eras",
        options=era_options,
        default=era_options,
    )
    if selected_eras:
        filtered = filtered[filtered["era"].isin(selected_eras)]

    st.sidebar.markdown("### Sample Layer")
    max_countries = st.sidebar.slider(
        "Max countries in analysis",
        min_value=10,
        max_value=200,
        value=60,
        step=10,
        help="Limit insights to the top N countries by average inflation.",
    )
    if filtered["REF_AREA"].nunique() > max_countries:
        top_codes = (
            filtered.groupby("REF_AREA")["OBS_VALUE"]
            .mean()
            .sort_values(ascending=False)
            .head(max_countries)
            .index
        )
        filtered = filtered[filtered["REF_AREA"].isin(top_codes)]

    filters = {
        "start_date": start_ts,
        "end_date": end_ts,
        "regions": selected_regions if selected_regions else ["All"],
        "eras": selected_eras if selected_eras else ["All"],
        "max_countries": max_countries,
        "scope_label": _format_scope_label(start_ts, end_ts, selected_regions, selected_eras, max_countries),
    }

    return filtered.reset_index(drop=True), filters


def _format_scope_label(
    start: pd.Timestamp,
    end: pd.Timestamp,
    regions: list,
    eras: list,
    max_countries: int,
) -> str:
    region_label = ", ".join(regions) if regions else "All Regions"
    era_label = ", ".join(eras) if eras else "All Eras"
    return (
        f"{start:%Y-%m} → {end:%Y-%m} | Regions: {region_label} | "
        f"Eras: {era_label} | Sample ≤{max_countries} countries"
    )


def main() -> None:
    st.set_page_config(
        page_title="Global Inflation Intelligence Hub",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    config = load_config()
    processed_dir = Path(config["data"]["processed_dir"])
    processed_file = processed_dir / config["data"]["processed_file"]

    df = load_dataset(processed_file)
    filtered_df, filters = apply_global_controls(df)
    if filtered_df.empty:
        st.error("No data remains for the selected scope. Adjust the filters on the left.")
        return

    view_registry = load_views()
    active_view = get_active_view(view_registry)

    state = DashboardState(
        data=filtered_df,
        config=config,
        filters=filters,
    )

    active_view(state)


if __name__ == "__main__":
    main()
