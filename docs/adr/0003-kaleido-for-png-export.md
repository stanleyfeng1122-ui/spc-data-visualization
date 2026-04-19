# ADR-0003: kaleido for PNG export

**Status:** Accepted — 2026-04-18
**Deciders:** Stanley Feng

## Context
The batch-export feature needs to convert many Plotly figures into PNG files server-side without opening a browser, so that users can generate report bundles in one click.

## Decision
Use `kaleido >= 1.2` as the static-image export backend, invoked through `fig.to_image(format="png")`.

## Alternatives considered
- **orca**: Plotly's previous export backend — deprecated and no longer recommended by Plotly maintainers.
- **Selenium + headless Chrome**: works but is heavy (browser driver, profile management, flaky startup) and overkill for static export.
- **matplotlib re-draw**: would require maintaining a parallel chart implementation, breaking feature parity with the interactive Plotly charts.

## Consequences
- ✅ Pure-Python API — `fig.to_image(format="png")` returns bytes with no external process management.
- ✅ Single source of truth: the same Plotly figure powers both the interactive UI and the exported PNG.
- ✅ Stable enough for batch loops over dozens of charts.
- ⚠️ kaleido 1.2+ removed the `__version__` attribute, so any version-sniffing tests need a minor adjustment.
- ⚠️ Installs a bundled Chromium runtime (~80 MB), inflating the environment footprint.
