from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st


@st.cache_data(show_spinner=False)
def load_processed_data(processed_path: Path) -> pd.DataFrame:
    if not processed_path.exists():
        raise FileNotFoundError(
            f"Processed parquet not found at {processed_path}. Run src.pipeline first."
        )

    df = pd.read_parquet(processed_path)
    if "TIME_PERIOD" in df.columns:
        df["TIME_PERIOD"] = pd.to_datetime(df["TIME_PERIOD"])
    return df
