from __future__ import annotations

from datetime import datetime
from typing import Dict, List

import pandas as pd

from src.components.excel_exporter import ExcelExporter, format_currency, format_number, format_percentage


def generate_inflation_report(df: pd.DataFrame, metadata_overrides: Dict[str, str]) -> bytes:
    exporter = ExcelExporter()

    start = df["TIME_PERIOD"].min()
    end = df["TIME_PERIOD"].max()
    cover_meta = {
        "report_title": metadata_overrides.get("report_title", "Global Inflation Intelligence Hub"),
        "report_subtitle": metadata_overrides.get("report_subtitle", "Comprehensive Historical Workbook"),
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "start_date": f"{start:%Y-%m}",
        "end_date": f"{end:%Y-%m}",
        "scope": metadata_overrides.get("scope", "Full Dataset"),
        "key_metrics": {
            "Observations": f"{len(df):,}",
            "Countries": df["REF_AREA"].nunique(),
            "Coverage": f"{start:%Y-%m} – {end:%Y-%m}",
            "Global Mean": f"{df['OBS_VALUE'].mean():.2f}%",
        },
    }
    exporter.create_cover_sheet(cover_meta)

    exporter.add_data_sheet("Global Story", _build_global_sections(df))
    exporter.add_data_sheet("Era Explorer", _build_era_sections(df))
    exporter.add_data_sheet("Trend Explorer", _build_trend_sections(df))
    exporter.add_data_sheet("Regional Atlas", _build_regional_sections(df))
    exporter.add_data_sheet("Market Depth", _build_market_sections(df))
    exporter.add_data_sheet("Supply Watch", _build_supply_sections(df))
    exporter.add_data_sheet("Policy Ledger", _build_policy_sections(df))
    exporter.add_data_sheet("Risk Monitor", _build_risk_sections(df))
    exporter.add_data_dictionary_sheet()
    return exporter.save_to_bytes()


def _build_global_sections(df: pd.DataFrame) -> List[Dict]:
    latest_period = df["TIME_PERIOD"].max()
    latest_slice = df[df["TIME_PERIOD"] == latest_period]
    top_countries = (
        latest_slice.groupby("REF_AREA_LABEL")["OBS_VALUE"]
        .mean()
        .reset_index()
        .sort_values("OBS_VALUE", ascending=False)
        .head(25)
    )
    top_countries_table = top_countries.rename(columns={"REF_AREA_LABEL": "Country", "OBS_VALUE": "Inflation (%)"})

    timeline = (
        df.groupby("TIME_PERIOD")["OBS_VALUE"]
        .mean()
        .reset_index()
        .rename(columns={"OBS_VALUE": "Global Avg"})
    )
    timeline["Rolling 12M"] = timeline["Global Avg"].rolling(window=12, min_periods=6).mean()

    summary = pd.DataFrame(
        [
            ("Observations", f"{len(df):,}"),
            ("Countries", df["REF_AREA"].nunique()),
            ("Date Span", f"{df['TIME_PERIOD'].min():%Y-%m} → {latest_period:%Y-%m}"),
            ("Global Mean", f"{df['OBS_VALUE'].mean():.2f}%"),
            ("Global σ", f"{df['OBS_VALUE'].std():.2f} pts"),
        ],
        columns=["Metric", "Value"],
    )

    return [
        {
            "title": "Global KPIs",
            "data": summary,
            "chart_type": "table",
            "description": "High-level coverage and volatility statistics.",
        },
        {
            "title": f"Top 25 Countries – {latest_period:%B %Y}",
            "data": top_countries_table.round(2),
            "chart_type": "bar",
            "data_raw": top_countries.rename(columns={"REF_AREA_LABEL": "Country"}),
            "chart_config": {"title": "Latest Inflation Leaders", "data_col": "OBS_VALUE", "category_col": "Country"},
        },
        {
            "title": "Global Average Timeline",
            "data": timeline.tail(600).round(3),
            "chart_type": "line",
            "data_raw": timeline[["TIME_PERIOD", "Global Avg"]].tail(600),
            "chart_config": {"title": "Global Average Timeline", "data_col": "Global Avg", "category_col": "TIME_PERIOD"},
        },
    ]


def _build_era_sections(df: pd.DataFrame) -> List[Dict]:
    stats = (
        df.groupby("era")["OBS_VALUE"]
        .agg(avg="mean", median="median", volatility="std", max="max", min="min")
        .reset_index()
    )
    sections: List[Dict] = []
    for col, label in [
        ("avg", "Average Inflation by Era"),
        ("median", "Median Inflation by Era"),
        ("volatility", "Volatility (σ) by Era"),
        ("max", "Peak Inflation by Era"),
        ("min", "Lowest Inflation by Era"),
    ]:
        sections.append(
            {
                "title": label,
                "data": stats[["era", col]].rename(columns={"era": "Era", col: label}).round(2),
                "chart_type": "bar",
                "data_raw": stats[["era", col]],
                "chart_config": {"title": label, "data_col": col, "category_col": "era"},
            }
        )

    ranking = (
        df.groupby(["era", "REF_AREA_LABEL"])["OBS_VALUE"]
        .mean()
        .reset_index()
        .rename(columns={"REF_AREA_LABEL": "Country", "OBS_VALUE": "Average (%)"})
        .sort_values(["era", "Average (%)"], ascending=[True, False])
    )
    top_per_era = ranking.groupby("era").head(5)
    sections.append(
        {
            "title": "Top 5 Countries per Era",
            "data": top_per_era,
            "chart_type": "table",
            "description": "Country leaders within each structural era.",
        }
    )
    return sections


def _build_trend_sections(df: pd.DataFrame) -> List[Dict]:
    latest = df["TIME_PERIOD"].max()
    prev = latest - pd.DateOffset(months=1)
    movers = (
        df[df["TIME_PERIOD"].isin([latest, prev])]
        .groupby(["REF_AREA_LABEL", "TIME_PERIOD"])["OBS_VALUE"]
        .mean()
        .unstack()
        .dropna()
    )
    if latest in movers and prev in movers:
        movers["Δ MoM"] = movers[latest] - movers[prev]
        risers = (
            movers.sort_values("Δ MoM", ascending=False)
            .head(10)
            .reset_index()
            .rename(columns={latest: f"{latest:%Y-%m}", prev: f"{prev:%Y-%m}"})
        )
        decliners = (
            movers.sort_values("Δ MoM")
            .head(10)
            .reset_index()
            .rename(columns={latest: f"{latest:%Y-%m}", prev: f"{prev:%Y-%m}"})
        )
    else:
        risers = pd.DataFrame(columns=["REF_AREA_LABEL", "Δ MoM"])
        decliners = pd.DataFrame(columns=["REF_AREA_LABEL", "Δ MoM"])

    yearly = (
        df.groupby("year")["OBS_VALUE"]
        .mean()
        .reset_index()
        .rename(columns={"OBS_VALUE": "Average (%)"})
    )

    sections = [
        {
            "title": "Yearly Global Averages",
            "data": yearly,
            "chart_type": "line",
            "data_raw": yearly,
            "chart_config": {"title": "Yearly Average Inflation", "data_col": "Average (%)", "category_col": "year"},
        }
    ]
    sections.append(
        {
            "title": "Top 10 Monthly Risers",
            "data": risers.rename(columns={"REF_AREA_LABEL": "Country"}),
            "chart_type": "table",
        }
    )
    sections.append(
        {
            "title": "Top 10 Monthly Decliners",
            "data": decliners.rename(columns={"REF_AREA_LABEL": "Country"}),
            "chart_type": "table",
        }
    )
    return sections


def _build_regional_sections(df: pd.DataFrame) -> List[Dict]:
    share = (
        df.groupby("region")["OBS_VALUE"]
        .agg(avg="mean", volatility="std", count="size")
        .reset_index()
        .sort_values("avg", ascending=False)
    )
    share["Coverage"] = share["count"].apply(format_number)

    latest = df["TIME_PERIOD"].max()
    hotspots = (
        df[df["TIME_PERIOD"] == latest]
        .assign(flag=lambda x: x["OBS_VALUE"] >= 10)
        .groupby("region")["flag"]
        .mean()
        .reset_index()
        .rename(columns={"flag": "Share ≥10%"})
        .sort_values("Share ≥10%", ascending=False)
    )
    hotspots["Share ≥10%"] = hotspots["Share ≥10%"].apply(lambda x: format_percentage(x * 100))

    return [
        {
            "title": "Regional Inflation Profile",
            "data": share[["region", "avg", "volatility", "Coverage"]].rename(
                columns={"region": "Region", "avg": "Average (%)", "volatility": "σ"}
            ),
            "chart_type": "bar",
            "data_raw": share[["region", "avg"]].rename(columns={"region": "Region"}),
            "chart_config": {"title": "Average Inflation by Region", "data_col": "avg", "category_col": "Region"},
        },
        {
            "title": f"Hotspot Share ({latest:%b %Y})",
            "data": hotspots.rename(columns={"region": "Region"}),
            "chart_type": "table",
            "description": "Portion of countries per region with ≥10% inflation in the latest month.",
        },
    ]


def _build_market_sections(df: pd.DataFrame) -> List[Dict]:
    latest = df["TIME_PERIOD"].max()
    window = 12
    rolling = (
        df.sort_values(["REF_AREA_LABEL", "TIME_PERIOD"])
        .groupby("REF_AREA_LABEL")["OBS_VALUE"]
        .rolling(window=window, min_periods=3)
        .std()
        .reset_index()
        .rename(columns={"OBS_VALUE": "volatility"})
    )
    latest_vol = rolling.groupby("REF_AREA_LABEL").tail(1).set_index("REF_AREA_LABEL")
    latest_slice = (
        df[df["TIME_PERIOD"] == latest][["REF_AREA_LABEL", "OBS_VALUE", "region"]]
        .set_index("REF_AREA_LABEL")
    )
    merged = latest_slice.join(latest_vol, how="inner").dropna().reset_index()
    merged = merged.rename(columns={"REF_AREA_LABEL": "Country", "OBS_VALUE": "Latest (%)", "volatility": "σ"})
    merged = merged.sort_values("Latest (%)", ascending=False).head(30)

    distribution = df[df["TIME_PERIOD"] == latest]["OBS_VALUE"]
    hist = pd.cut(distribution, bins=15).value_counts().sort_index().reset_index()
    hist.columns = ["Bin", "Countries"]

    return [
        {
            "title": f"Level vs Volatility (Top 30) – {latest:%b %Y}",
            "data": merged,
            "chart_type": "bar",
            "data_raw": merged[["Country", "Latest (%)"]],
            "chart_config": {"title": "Latest Inflation Levels", "data_col": "Latest (%)", "category_col": "Country"},
        },
        {
            "title": "Latest Month Distribution",
            "data": hist,
            "chart_type": "bar",
            "data_raw": hist,
            "chart_config": {"title": "Distribution Histogram", "data_col": "Countries", "category_col": "Bin"},
        },
    ]


def _build_supply_sections(df: pd.DataFrame) -> List[Dict]:
    threshold = 10
    share_raw = (
        df.assign(flag=df["OBS_VALUE"] >= threshold)
        .groupby("region")["flag"]
        .mean()
        .reset_index()
        .rename(columns={"flag": "share"})
        .sort_values("share", ascending=False)
    )
    share_table = share_raw.copy()
    share_table["Share (%)"] = share_table["share"].apply(lambda x: format_percentage(x * 100))
    share_table = share_table.drop(columns=["share"]).rename(columns={"region": "Region"})

    duration = (
        df.assign(flag=df["OBS_VALUE"] >= threshold)
        .groupby("REF_AREA_LABEL")["flag"]
        .sum()
        .reset_index()
        .rename(columns={"REF_AREA_LABEL": "Country", "flag": "Months ≥10%"})
        .sort_values("Months ≥10%", ascending=False)
        .head(25)
    )

    return [
        {
            "title": "High-Inflation Share by Region",
            "data": share_table,
            "chart_type": "bar",
            "data_raw": share_raw.rename(columns={"region": "Region", "share": "Share"}),
            "chart_config": {"title": "Share of Months ≥10%", "data_col": "Share", "category_col": "Region"},
        },
        {
            "title": "Persistence Leaders (Months ≥10%)",
            "data": duration,
            "chart_type": "table",
            "description": "Countries with the longest history above the high-inflation threshold.",
        },
    ]


def _build_policy_sections(df: pd.DataFrame) -> List[Dict]:
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

    improvements = merged.sort_values("Change").head(15)
    regressions = merged.sort_values("Change", ascending=False).head(15)

    stable_share = (df["OBS_VALUE"] <= 5).mean() * 100
    severe_share = (df["OBS_VALUE"] >= 15).mean() * 100
    policy_summary = pd.DataFrame(
        [
            ("Stable Months (≤5%)", format_percentage(stable_share)),
            ("Severe Months (≥15%)", format_percentage(severe_share)),
            ("Mean Annual Drift", f"{df.groupby('year')['OBS_VALUE'].mean().diff().mean():+.2f} pts/year"),
        ],
        columns=["Metric", "Value"],
    )

    return [
        {"title": "Policy Summary", "data": policy_summary, "chart_type": "table"},
        {
            "title": f"Most Improved ({first_year}→{last_year})",
            "data": improvements.rename(columns={"REF_AREA_LABEL": "Country"}),
            "chart_type": "table",
        },
        {
            "title": f"Biggest Deteriorations ({first_year}→{last_year})",
            "data": regressions.rename(columns={"REF_AREA_LABEL": "Country"}),
            "chart_type": "table",
        },
    ]


def _build_risk_sections(df: pd.DataFrame) -> List[Dict]:
    scores = (
        df.groupby("REF_AREA_LABEL")["OBS_VALUE"]
        .agg(mean="mean", volatility="std", p95=lambda s: s.quantile(0.95), p05=lambda s: s.quantile(0.05))
        .reset_index()
        .rename(columns={"REF_AREA_LABEL": "Country"})
    )
    scores["range"] = scores["p95"] - scores["p05"]
    scores["Risk Score"] = _normalize(scores["mean"]) * 0.4 + _normalize(scores["volatility"]) * 0.35 + _normalize(
        scores["range"]
    ) * 0.25
    scores = scores.sort_values("Risk Score", ascending=False)

    high_risk = scores.head(20)
    resilient = scores.tail(20).sort_values("Risk Score")

    return [
        {
            "title": "High-Risk Countries",
            "data": high_risk[["Country", "mean", "volatility", "range", "Risk Score"]].round(3),
            "chart_type": "bar",
            "data_raw": high_risk[["Country", "Risk Score"]],
            "chart_config": {"title": "Risk Scores (Top 20)", "data_col": "Risk Score", "category_col": "Country"},
        },
        {
            "title": "Resilience Board",
            "data": resilient[["Country", "mean", "volatility", "range", "Risk Score"]].round(3),
            "chart_type": "table",
        },
    ]


def _normalize(series: pd.Series) -> pd.Series:
    if series.max() == series.min():
        return pd.Series(0, index=series.index)
    return (series - series.min()) / (series.max() - series.min())
