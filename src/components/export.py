from __future__ import annotations

from io import BytesIO
from typing import Dict, Iterable

import numpy as np
import pandas as pd
import streamlit as st


def build_excel_report(df: pd.DataFrame, metadata: Dict[str, str]) -> bytes:
    try:
        import xlsxwriter  # noqa: F401
    except ModuleNotFoundError as exc:  # pragma: no cover
        raise ModuleNotFoundError(
            "Excel export requires the 'xlsxwriter' package. Install it via 'pip install xlsxwriter'."
        ) from exc

    if df.empty:
        raise ValueError("Cannot build export with empty dataframe.")

    buf = BytesIO()
    latest_period = df["TIME_PERIOD"].max()
    latest_slice = df[df["TIME_PERIOD"] == latest_period]

    summary = (
        latest_slice.groupby("REF_AREA_LABEL")["OBS_VALUE"]
        .mean()
        .sort_values(ascending=False)
        .reset_index()
        .rename(columns={"REF_AREA_LABEL": "Country", "OBS_VALUE": "Inflation (%)"})
        .head(50)
    )

    trend_pivot = (
        df.pivot_table(
            index="TIME_PERIOD",
            columns="REF_AREA_LABEL",
            values="OBS_VALUE",
            aggfunc="mean",
        )
        .sort_index()
    )

    timeline = (
        df.groupby("TIME_PERIOD")["OBS_VALUE"]
        .mean()
        .reset_index()
        .rename(columns={"OBS_VALUE": "Global Avg"})
    )
    timeline["Rolling 12M"] = timeline["Global Avg"].rolling(window=12, min_periods=6).mean()

    hist = _build_histogram(latest_slice["OBS_VALUE"])

    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        summary.to_excel(writer, sheet_name="Latest Snapshot", index=False)
        trend_pivot.to_excel(writer, sheet_name="Trend Pivot")
        timeline.to_excel(writer, sheet_name="Global Timeline", index=False)
        hist.to_excel(writer, sheet_name="Distribution", index=False)
        df.to_excel(writer, sheet_name="Filtered Data", index=False)

        meta_sheet = pd.DataFrame(
            [{"Key": key, "Value": value} for key, value in metadata.items()]
        )
        meta_sheet.to_excel(writer, sheet_name="Metadata", index=False)

        workbook = writer.book
        header_format = workbook.add_format(
            {"bold": True, "bg_color": "#1f4e78", "font_color": "white"}
        )

        _style_sheet(writer.sheets["Latest Snapshot"], summary.columns, header_format)
        _style_sheet(writer.sheets["Trend Pivot"], trend_pivot.columns, header_format)
        _style_sheet(writer.sheets["Global Timeline"], timeline.columns, header_format)
        _style_sheet(writer.sheets["Distribution"], hist.columns, header_format)
        _style_sheet(writer.sheets["Filtered Data"], df.columns, header_format)

        _add_timeline_chart(workbook, writer.sheets["Global Timeline"], timeline)
        _add_snapshot_chart(workbook, writer.sheets["Latest Snapshot"], summary)
        _add_histogram_chart(workbook, writer.sheets["Distribution"], hist)

    buf.seek(0)
    return buf.getvalue()


def render_export_button(df: pd.DataFrame, *, metadata: Dict[str, str]) -> None:
    if df.empty:
        st.caption("Add data filters to enable export.")
        st.button("Download Excel", disabled=True)
        return
    try:
        payload = build_excel_report(df, metadata)
    except ModuleNotFoundError as exc:
        st.error(str(exc))
        return
    except ValueError as exc:
        st.error(str(exc))
        return

    st.download_button(
        label="Download Excel",
        data=payload,
        file_name="global_inflation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _style_sheet(worksheet, columns: Iterable, header_format) -> None:
    worksheet.freeze_panes(1, 0)
    cols = list(columns)
    worksheet.autofilter(0, 0, 0, max(len(cols) - 1, 0))
    for col_num, _ in enumerate(cols):
        worksheet.set_column(col_num, col_num, 18)
    worksheet.set_row(0, None, header_format)


def _add_timeline_chart(workbook, worksheet, timeline: pd.DataFrame) -> None:
    rows = len(timeline)
    chart = workbook.add_chart({"type": "line"})
    chart.add_series(
        {
            "name": "Global Avg",
            "categories": ["Global Timeline", 1, 0, rows, 0],
            "values": ["Global Timeline", 1, 1, rows, 1],
        }
    )
    chart.add_series(
        {
            "name": "Rolling 12M",
            "categories": ["Global Timeline", 1, 0, rows, 0],
            "values": ["Global Timeline", 1, 2, rows, 2],
        }
    )
    chart.set_title({"name": "Global Average Timeline"})
    chart.set_y_axis({"name": "Inflation (%)"})
    worksheet.insert_chart("E2", chart, {"x_scale": 1.4, "y_scale": 1.2})


def _add_snapshot_chart(workbook, worksheet, summary: pd.DataFrame) -> None:
    rows = len(summary)
    chart = workbook.add_chart({"type": "column"})
    chart.add_series(
        {
            "name": "Latest Inflation",
            "categories": ["Latest Snapshot", 1, 0, rows, 0],
            "values": ["Latest Snapshot", 1, 1, rows, 1],
        }
    )
    chart.set_title({"name": "Top Latest Inflation"})
    chart.set_y_axis({"name": "Inflation (%)"})
    worksheet.insert_chart("D2", chart, {"x_scale": 1.3, "y_scale": 1.2})


def _add_histogram_chart(workbook, worksheet, hist: pd.DataFrame) -> None:
    rows = len(hist)
    chart = workbook.add_chart({"type": "column"})
    chart.add_series(
        {
            "name": "Countries",
            "categories": ["Distribution", 1, 0, rows, 0],
            "values": ["Distribution", 1, 2, rows, 2],
        }
    )
    chart.set_title({"name": "Distribution of Latest Month"})
    chart.set_x_axis({"name": "Inflation Bin"})
    chart.set_y_axis({"name": "Countries"})
    worksheet.insert_chart("E2", chart, {"x_scale": 1.3, "y_scale": 1.2})


def _build_histogram(values: pd.Series) -> pd.DataFrame:
    cleaned = values.dropna()
    if cleaned.empty:
        return pd.DataFrame(
            {"Bin Start": [0], "Bin End": [0], "Countries": [0]}
        )

    bins = min(20, max(6, int(len(cleaned) / 10)))
    counts, edges = np.histogram(cleaned, bins=bins)
    data = []
    for start, end, count in zip(edges[:-1], edges[1:], counts):
        data.append(
            {
                "Bin Start": round(float(start), 2),
                "Bin End": round(float(end), 2),
                "Countries": int(count),
            }
        )
    return pd.DataFrame(data)
