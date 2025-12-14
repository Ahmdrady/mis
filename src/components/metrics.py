from __future__ import annotations

from typing import Iterable, Mapping

import streamlit as st


def render_metric_row(metrics: Iterable[Mapping[str, object]]) -> None:
    metrics = list(metrics)
    if not metrics:
        st.warning("No metrics to display.")
        return

    cols = st.columns(len(metrics))
    for col, metric in zip(cols, metrics):
        with col:
            st.metric(
                label=str(metric.get("label", "")),
                value=str(metric.get("value", "—")),
                delta=str(metric.get("delta", "")),
                delta_color=str(metric.get("delta_color", "normal")),
                help=metric.get("help"),
            )
