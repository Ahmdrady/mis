from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict

import pandas as pd

try:  # Python 3.11+
    import tomllib
except ModuleNotFoundError:  # pragma: no cover
    import tomli as tomllib  # type: ignore

from src.data.region_map import ISO_REGION_MAP, REGION_FALLBACK


CONFIG_PATH = Path("config") / "settings.toml"


@dataclass
class PipelineReport:
    rows_processed: int
    countries: int
    periods: int
    time_start: str
    time_end: str
    parquet_path: str

    def to_json(self) -> str:
        return json.dumps(asdict(self), indent=2, default=str)


def load_config(config_path: Path = CONFIG_PATH) -> Dict[str, Any]:
    if not config_path.exists():
        raise FileNotFoundError(f"Config file missing at {config_path}")
    with config_path.open("rb") as fh:
        config = tomllib.load(fh)
    if "data" not in config or "pipeline" not in config:
        raise KeyError("settings.toml must include [data] and [pipeline]")
    return config


def ensure_directories(raw_dir: Path, processed_dir: Path) -> None:
    if not raw_dir.exists():
        raise FileNotFoundError(
            f"Raw directory {raw_dir} not found. Place the CSV there before running the pipeline."
        )
    processed_dir.mkdir(parents=True, exist_ok=True)


def ingest_csv(raw_path: Path) -> pd.DataFrame:
    if not raw_path.exists():
        raise FileNotFoundError(f"Raw CSV not found at {raw_path}")
    return pd.read_csv(
        raw_path,
        dtype={
            "REF_AREA": "string",
            "REF_AREA_LABEL": "string",
            "TIME_PERIOD": "string",
        },
    )


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned["REF_AREA"] = (
        cleaned["REF_AREA"].astype("string").str.strip().str.upper().fillna("UNK")
    )
    cleaned["REF_AREA_LABEL"] = (
        cleaned["REF_AREA_LABEL"].astype("string").str.strip().fillna("Unknown Region")
    )
    cleaned["TIME_PERIOD"] = pd.to_datetime(
        cleaned["TIME_PERIOD"], errors="coerce", format="%Y-%m-%d"
    )
    cleaned = cleaned.dropna(subset=["TIME_PERIOD"])
    cleaned["OBS_VALUE"] = pd.to_numeric(cleaned["OBS_VALUE"], errors="coerce")
    cleaned = cleaned.dropna(subset=["OBS_VALUE"])

    cleaned["region"] = (
        cleaned["REF_AREA"]
        .map(ISO_REGION_MAP)
        .fillna(REGION_FALLBACK)
        .astype("string")
    )

    cleaned["year"] = cleaned["TIME_PERIOD"].dt.year.astype("Int16")
    cleaned["month"] = cleaned["TIME_PERIOD"].dt.month.astype("Int8")
    cleaned["quarter"] = cleaned["TIME_PERIOD"].dt.quarter.astype("Int8")
    cleaned["year_month"] = cleaned["TIME_PERIOD"].dt.to_period("M").astype(str)
    cleaned["decade"] = (cleaned["year"] // 10 * 10).astype("Int16")
    cleaned["era"] = cleaned["TIME_PERIOD"].apply(classify_era).astype("string")

    cleaned = cleaned.sort_values(["REF_AREA", "TIME_PERIOD"]).drop_duplicates(
        subset=["REF_AREA", "TIME_PERIOD"], keep="last"
    )
    return cleaned.reset_index(drop=True)


def classify_era(timestamp: pd.Timestamp) -> str:
    if timestamp < pd.Timestamp("2007-07-01"):
        return "Pre-GFC"
    if timestamp < pd.Timestamp("2014-01-01"):
        return "Post-GFC Recovery"
    if timestamp < pd.Timestamp("2020-01-01"):
        return "Commodity Reset"
    if timestamp < pd.Timestamp("2022-02-01"):
        return "Pandemic Shock"
    return "Geopolitical Reordering"


def write_parquet(
    df: pd.DataFrame,
    processed_path: Path,
    *,
    engine: str,
    compression: str,
) -> None:
    processed_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_parquet(processed_path, engine=engine, compression=compression, index=False)


def run_pipeline() -> PipelineReport:
    config = load_config()
    data_cfg = config["data"]
    pipeline_cfg = config["pipeline"]

    raw_dir = Path(data_cfg["raw_dir"])
    processed_dir = Path(data_cfg["processed_dir"])
    raw_path = raw_dir / data_cfg["raw_file"]
    processed_path = processed_dir / data_cfg["processed_file"]

    ensure_directories(raw_dir, processed_dir)

    df_raw = ingest_csv(raw_path)
    df_clean = clean_dataframe(df_raw)
    write_parquet(
        df_clean,
        processed_path,
        engine=pipeline_cfg.get("parquet_engine", "pyarrow"),
        compression=pipeline_cfg.get("compression", "snappy"),
    )

    time_start = df_clean["TIME_PERIOD"].min()
    time_end = df_clean["TIME_PERIOD"].max()

    return PipelineReport(
        rows_processed=len(df_clean),
        countries=df_clean["REF_AREA"].nunique(),
        periods=df_clean["TIME_PERIOD"].dt.to_period("M").nunique(),
        time_start=time_start.isoformat(),
        time_end=time_end.isoformat(),
        parquet_path=str(processed_path),
    )


if __name__ == "__main__":
    report = run_pipeline()
    print(report.to_json())
