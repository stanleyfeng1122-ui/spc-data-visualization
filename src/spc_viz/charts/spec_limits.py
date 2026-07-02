"""Shared spec-limit rendering — the one place USL / LSL / Nominal are drawn.

Every chart builder used to hand-roll its own spec lines, tolerance band and
labels; fixes (like stepping specs) landed in one builder and never reached
the others. This module puts all of it behind one small interface:

* :class:`SpecSpan` — a spec over an x-range. ``x0/x1 = None`` means the spec
  is uniform (full axis width); a list with several spans is a stepping spec.
* :class:`SpecStyle` — the visual identity of one chart's spec drawing. The
  four presets below reproduce each builder's historical look exactly; they
  are intentionally separate so consolidation changed zero pixels. Converge
  them deliberately, not accidentally.
* :func:`render_spec_limits` — draws lines / bands / segments onto a figure.
* :func:`spec_axis_annotations` — returns the red ``USL-0.4`` label dicts so
  each builder can splice them into its own annotation ordering unchanged.

Deviation-from-nominal shifting happens in here; callers pass raw specs.
"""

from __future__ import annotations

from dataclasses import dataclass

from plotly.graph_objects import Figure

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class SpecSpan:
    """USL/LSL/Nominal over one x-range; ``x0/x1 = None`` = full width."""

    usl: float | None
    lsl: float | None
    nominal: float | None
    x0: float | None = None
    x1: float | None = None

    @property
    def full_width(self) -> bool:
        return self.x0 is None or self.x1 is None


@dataclass(frozen=True)
class SpecStyle:
    """One chart's spec-drawing identity. Presets preserve historical looks."""

    line_alpha: float = 0.5  # USL/LSL line alpha when uniform
    stepping_line_alpha: float = 0.7  # ... when drawn as stepping segments
    line_width: float = 1.2
    band_alpha: float | None = 0.15  # tolerance-band fill alpha; None = no band
    stepping_band: bool = True  # draw per-segment band rects in stepping mode
    # (annotation_text, annotation_position) per line; None = plain line
    usl_label: tuple[str, str] | None = None
    lsl_label: tuple[str, str] | None = None
    nominal_line: dict | None = None  # plotly line style; None = skip nominal
    nominal_label: tuple[str, str] | None = None
    # Red axis-edge value labels: "rep" (first span only), "unique" (all
    # distinct raw values, USL desc then LSL asc), "encounter" (span order,
    # de-duplicated, deviation-shifted), "none".
    axis_label_mode: str = "rep"
    axis_label_anchor: str = "left"  # "left" adds xshift=5; "right" none


# The four historical looks. Differences between them are accidental drift
# from the hand-rolled era, kept so consolidation was pixel-identical.
PROFILE_STYLE = SpecStyle(band_alpha=0.15, axis_label_mode="unique")
BOX_STYLE = SpecStyle(
    band_alpha=0.10,
    usl_label=("USL", "top right"),
    lsl_label=("LSL", "bottom right"),
    nominal_line=dict(color="rgba(34,197,94,0.5)", dash="dot", width=1),
    nominal_label=("Nominal", "top right"),
    axis_label_mode="rep",
    axis_label_anchor="right",
)
HISTOGRAM_STYLE = SpecStyle(
    line_alpha=0.7,
    line_width=1.5,
    band_alpha=None,
    usl_label=("USL", "top right"),
    lsl_label=("LSL", "top left"),
    nominal_line=dict(color="rgba(34,197,94,0.7)", dash="dot", width=1.2),
    nominal_label=("Nom", "top"),
    axis_label_mode="none",
)
ENVELOPE_STYLE = SpecStyle(
    line_alpha=0.55,
    stepping_line_alpha=0.55,
    band_alpha=0.10,
    stepping_band=False,
    axis_label_mode="encounter",
)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _shift(value: float | None, nominal: float | None, deviation_mode: bool) -> float | None:
    """Deviation-mode shift: value relative to nominal when both known."""
    if value is None:
        return None
    if deviation_mode and nominal is not None:
        return value - nominal
    return value


def _is_stepping(spans: list[SpecSpan]) -> bool:
    unique_usls = {s.usl for s in spans if s.usl is not None}
    unique_lsls = {s.lsl for s in spans if s.lsl is not None}
    return len(unique_usls) > 1 or len(unique_lsls) > 1


# ---------------------------------------------------------------------------
# Public interface
# ---------------------------------------------------------------------------


def render_spec_limits(
    fig: Figure,
    spans: list[SpecSpan],
    *,
    style: SpecStyle,
    orientation: str = "h",
    deviation_mode: bool = False,
    rows: list[dict] | None = None,
) -> None:
    """Draw USL/LSL/Nominal lines, tolerance band, and stepping segments.

    ``rows`` is a list of plotly row/col kwarg dicts to repeat the drawing
    into (facets); ``[{}]`` (the default) draws once on the base axes.
    Full-width spans become ``add_hline``/``add_vline``/``add_hrect``;
    ranged spans become per-segment ``add_shape`` lines and rects.
    """
    if not spans:
        return
    row_kwargs_list = rows if rows else [{}]
    dash = dict(dash="dash", width=style.line_width)
    line_uniform = dict(color=f"rgba(220,38,38,{style.line_alpha})", **dash)
    line_step = dict(color=f"rgba(220,38,38,{style.stepping_line_alpha})", **dash)

    add_line = fig.add_hline if orientation == "h" else fig.add_vline

    def _label_kwargs(label: tuple[str, str] | None) -> dict:
        if label is None:
            return {}
        return dict(annotation_text=label[0], annotation_position=label[1])

    for rk in row_kwargs_list:
        for span in spans:
            usl = _shift(span.usl, span.nominal, deviation_mode)
            lsl = _shift(span.lsl, span.nominal, deviation_mode)

            if span.full_width:
                # Band only makes sense on the value axis; all current charts
                # with a band are horizontal.
                if (
                    style.band_alpha is not None
                    and orientation == "h"
                    and usl is not None
                    and lsl is not None
                ):
                    fig.add_hrect(
                        y0=lsl,
                        y1=usl,
                        fillcolor=f"rgba(34, 197, 94, {style.band_alpha})",
                        line_width=0,
                        layer="below",
                        **rk,
                    )
                if usl is not None:
                    add_line(usl, line=line_uniform, **_label_kwargs(style.usl_label), **rk)
                if lsl is not None:
                    add_line(lsl, line=line_uniform, **_label_kwargs(style.lsl_label), **rk)
                if style.nominal_line is not None and span.nominal is not None:
                    nom = 0.0 if deviation_mode else span.nominal
                    add_line(
                        nom,
                        line=style.nominal_line,
                        **_label_kwargs(style.nominal_label),
                        **rk,
                    )
            else:
                if (
                    style.band_alpha is not None
                    and style.stepping_band
                    and usl is not None
                    and lsl is not None
                ):
                    fig.add_shape(
                        type="rect",
                        x0=span.x0,
                        x1=span.x1,
                        y0=lsl,
                        y1=usl,
                        fillcolor=f"rgba(34, 197, 94, {style.band_alpha})",
                        line_width=0,
                        layer="below",
                        **rk,
                    )
                if usl is not None:
                    fig.add_shape(
                        type="line",
                        x0=span.x0,
                        x1=span.x1,
                        y0=usl,
                        y1=usl,
                        line=line_step,
                        **rk,
                    )
                if lsl is not None:
                    fig.add_shape(
                        type="line",
                        x0=span.x0,
                        x1=span.x1,
                        y0=lsl,
                        y1=lsl,
                        line=line_step,
                        **rk,
                    )


def spec_axis_annotations(
    spans: list[SpecSpan],
    *,
    style: SpecStyle,
    deviation_mode: bool = False,
) -> list[dict]:
    """Red ``USL-0.4`` / ``LSL--0.1`` value labels at the value-axis edge.

    Returned as plain annotation dicts so each builder can append them into
    its own annotation list at the position it historically did.
    """
    if not spans or style.axis_label_mode == "none":
        return []

    values: list[tuple[float, str]] = []
    mode = style.axis_label_mode
    stepping = _is_stepping(spans)

    if mode == "unique" and stepping:
        # All distinct raw values: USL descending, then LSL ascending.
        # (Raw, not deviation-shifted — the engineer reads spec levels.)
        for u in sorted({s.usl for s in spans if s.usl is not None}, reverse=True):
            values.append((u, f"USL-{u:.4g}"))
        for lo in sorted({s.lsl for s in spans if s.lsl is not None}):
            values.append((lo, f"LSL-{lo:.4g}"))
    elif mode == "encounter" and stepping:
        seen: set[tuple[float, str]] = set()
        for s in spans:
            usl = _shift(s.usl, s.nominal, deviation_mode)
            lsl = _shift(s.lsl, s.nominal, deviation_mode)
            for v, prefix in ((usl, "USL"), (lsl, "LSL")):
                if v is None:
                    continue
                key = (round(v, 8), f"{prefix}-{v:.4g}")
                if key in seen:
                    continue
                seen.add(key)
                values.append((v, f"{prefix}-{v:.4g}"))
    else:
        # "rep" (and non-stepping "unique"/"encounter"): first span's values.
        s = spans[0]
        usl = _shift(s.usl, s.nominal, deviation_mode)
        lsl = _shift(s.lsl, s.nominal, deviation_mode)
        if usl is not None:
            values.append((usl, f"USL-{usl:.4g}"))
        if lsl is not None:
            values.append((lsl, f"LSL-{lsl:.4g}"))

    anchor_kwargs: dict = (
        dict(xanchor="left", xshift=5)
        if style.axis_label_anchor == "left"
        else dict(xanchor="right")
    )
    return [
        dict(
            x=0.0,
            y=val,
            xref="paper",
            yref="y",
            text=f"<b>{label}</b>",
            showarrow=False,
            font=dict(size=10, color="rgba(220,38,38,0.9)", family="Arial Black"),
            bgcolor="rgba(255,255,255,0.7)",
            **anchor_kwargs,
        )
        for val, label in values
    ]
