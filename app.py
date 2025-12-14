from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict

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
    view_registry = load_views()
    active_view = get_active_view(view_registry)

    state = DashboardState(
        data=df,
        config=config,
    )

    active_view(state)


if __name__ == "__main__":
    main()
