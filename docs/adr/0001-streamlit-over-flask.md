# ADR-0001: Streamlit over Flask/FastAPI

**Status:** Accepted — 2026-04-18
**Deciders:** Stanley Feng

## Context
The SPC visualization tool needs a local web UI for multi-file xlsx analysis by a single user on a single machine. The priority is rapid iteration on data exploration features, not multi-tenant web serving or polished public UX.

## Decision
Adopt Streamlit as the sole web framework, with the entire UI expressed as Python scripts driven by `st.*` widgets and `st.session_state`.

## Alternatives considered
- **Flask + Jinja templates**: would require hand-rolling a frontend framework (forms, state, live updates) for charts and file uploads — too much scaffolding for a single-user tool.
- **FastAPI + React**: forces two codebases (Python backend + JS frontend) and a build pipeline; disproportionate for local analysis.
- **Dash**: heavier callback model, tightly coupled to Plotly-only rendering, and more verbose for the widget-heavy layouts this tool needs.

## Consequences
- ✅ One Python file can be a working web app — minimal ceremony per feature.
- ✅ Native widgets (file uploader, multiselect, tabs) and `st.session_state` remove the need for custom state plumbing.
- ✅ Rapid reload on save enables tight data-exploration feedback loops.
- ⚠️ Styling is limited to Streamlit's theme system plus narrow CSS hooks — custom design is hard.
- ⚠️ UI testing is awkward; most coverage must come from unit-testing the underlying parsers and chart builders.
- ⚠️ `st.session_state` is per-browser-tab, so multi-tab workflows can diverge unexpectedly.
