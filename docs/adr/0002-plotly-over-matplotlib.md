# ADR-0002: Plotly over matplotlib/bokeh

**Status:** Accepted — 2026-04-18
**Deciders:** Stanley Feng

## Context
SPC charts in this tool require interactive behavior — hover tooltips over individual points, zoom into time ranges, and click-to-highlight lot IDs — while still supporting server-side PNG export for batch reports.

## Decision
Use Plotly (`plotly.graph_objects`) as the charting library for all visualizations, rendered in Streamlit via `st.plotly_chart` and exported to PNG via `fig.to_image()`.

## Alternatives considered
- **matplotlib**: static PNG only, no hover/zoom/click — fails the interactivity requirement even though export would be trivial.
- **bokeh**: weaker and less idiomatic Streamlit integration; static export story is less mature than Plotly + kaleido.
- **altair / vega-lite**: declarative model is elegant but less flexible for the custom SPC overlays (control limits, out-of-spec markers, annotations) we need.

## Consequences
- ✅ Native interactivity (hover, zoom, pan, legend toggle) with zero extra JS code.
- ✅ Deterministic server-side PNG export via the kaleido backend — same figure object for UI and export.
- ✅ First-class `st.plotly_chart` integration with theming and responsive sizing.
- ⚠️ Large JS bundle inflates initial page weight compared to matplotlib PNGs.
- ⚠️ Occasional JS rendering quirks (legend overflow, tick-label clipping) need figure-level workarounds.
