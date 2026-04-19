# SPC Data Visualization

Local Streamlit app for manufacturing quality engineers to turn vendor CPK Excel files into interactive SPC charts with batch PNG export.

<!-- TODO: add screenshot of main page -->

## Features

- Multi-file XLSX upload with auto-detection of dimensions, metadata, and header rows
- Three chart types: Combined Profile, Box Plot, Histogram
- USL/LSL/Nominal spec limits drawn on every chart
- Group/color/section by any metadata column (Build, CFG, Color, etc.)
- Batch export all selected charts as ZIP of PNGs
- 3 pages: main (analysis), Quick Test (scratch), Sheet Manager (compare)

See [User Stories](./docs/01-user-stories.md) for the full list.

## Quick Start

```bash
# Clone
git clone <repo>
cd "Data Visualiztion"

# Set up environment (Python 3.10)
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Run locally
streamlit run app.py --server.port 8503
```

Note: The launchd service (port 8504) is author-specific and not included. See [ADR-0004](./docs/adr/0004-openpyxl-monkey-patch.md) and [US-018](./docs/01-user-stories.md).

## Project Structure

```
├── app.py                  # Main Streamlit entry point
├── shared_ui.py            # Shared widgets + chart orchestration
├── chart_utils.py          # Chart builders (Profile, Box Plot, Histogram)
├── spc_parser.py           # Excel parsing (with openpyxl monkey-patch)
├── ui_theme.py             # CSS + Plotly theme
├── requirements.txt        # Python dependencies
├── pages/
│   ├── 1_Quick_Test.py    # Scratch analysis page
│   └── 2_Sheet_Manager.py # Sheet comparison tool
├── tests/
│   └── test_spc_charts.py # 24 unit tests
└── docs/
    ├── 00-product-brief.md
    ├── 01-user-stories.md
    ├── 02-architecture.md
    ├── 03-design-doc.md
    ├── adr/               # Architecture Decision Records
    └── plans/             # Implementation phases
```

See [Architecture](./docs/02-architecture.md) for diagrams.

## Documentation

- [Product Brief](./docs/00-product-brief.md) — what this app is for
- [User Stories](./docs/01-user-stories.md) — every feature
- [Architecture](./docs/02-architecture.md) — module + data flow diagrams
- [Design Doc](./docs/03-design-doc.md) — tech decisions & debt
- [ADRs](./docs/adr/) — architecture decision records
- [Implementation Plans](./docs/plans/) — Phase 0 & Phase 1 plans

## Testing

```bash
pytest tests/ -v
```

24 tests covering parser, chart builders, single-point dimension fix, and batch export.

## Contributing

Currently a one-person project. See `docs/` for context before making changes. Use the test suite and spc-test-runner agent.

## License

MIT
