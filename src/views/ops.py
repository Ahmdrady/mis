from __future__ import annotations

import streamlit as st

from app import DashboardState


def render(state: DashboardState) -> None:
    st.title("Ops Console")
    df = state.data.copy()

    if df.empty:
        st.warning("Dataframe is empty under the chosen filters.")
        return

    st.caption("Configure alert thresholds and preview who would be notified.")
    threshold = st.slider("Alert threshold (%)", min_value=-5.0, max_value=50.0, value=10.0, step=0.5)
    sensitivity = st.slider("Sensitivity (MoM change)", min_value=0.5, max_value=10.0, value=2.0, step=0.5)

    latest_period = df["TIME_PERIOD"].max()
    prev_period = latest_period - pd.DateOffset(months=1)

    latest = df[df["TIME_PERIOD"] == latest_period][["REF_AREA_LABEL", "region", "OBS_VALUE"]]
    prev = df[df["TIME_PERIOD"] == prev_period][["REF_AREA_LABEL", "OBS_VALUE"]]
    merged = latest.merge(prev, on="REF_AREA_LABEL", suffixes=("", "_prev"))
    merged["delta"] = merged["OBS_VALUE"] - merged["OBS_VALUE_prev"]

    breaches = merged[
        (merged["OBS_VALUE"] >= threshold) | (merged["delta"].abs() >= sensitivity)
    ].sort_values("OBS_VALUE", ascending=False)

    st.subheader("Pending Alerts")
    st.dataframe(
        breaches.rename(
            columns={
                "REF_AREA_LABEL": "Country",
                "region": "Region",
                "OBS_VALUE": f"{latest_period:%Y-%m}",
                "delta": "MoM Δ",
            }
        ),
        hide_index=True,
        use_container_width=True,
    )

    st.info(
        f"Benefit: {len(breaches)} countries would trigger notifications at {threshold}% "
        f"or ±{sensitivity} pp moves."
    )

    with st.form("alert_form"):
        channel = st.selectbox("Notification channel", ["Email", "Teams", "Slack"])
        message = st.text_area(
            "Message template",
            value="Alert: {country} inflation at {value:.1f}% with MoM {delta:+.2f} pp.",
        )
        submitted = st.form_submit_button("Simulate Dispatch")
        if submitted:
            previews = [
                message.format(
                    country=row["REF_AREA_LABEL"],
                    value=row["OBS_VALUE"],
                    delta=row["delta"],
                )
                for _, row in breaches.head(5).iterrows()
            ]
            st.success(
                f"Dispatch Preview ({channel}):\n" + "\n".join(previews) if previews else "No breaches to notify."
            )
