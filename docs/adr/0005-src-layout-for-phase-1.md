# ADR-0005: src/ layout for Phase 1 refactor

**Status:** Proposed — 2026-04-18 (Phase 1, not yet executed)
**Deciders:** Stanley Feng

## Context
The current codebase uses a flat layout with four files over 600 lines, no package structure, and no type hints. This makes layering unclear, testing harder, and onboarding slower as the tool grows.

## Decision
Restructure into a `src/spc_viz/` package with subpackages `parsers/`, `charts/`, `ui/`, `theme/`, and `config/`, each containing focused modules of roughly 150 lines.

## Alternatives considered
- **Keep the flat layout**: cheapest short-term, but leaves the current code-quality issues (file size, implicit layering, no enforceable boundaries) unresolved.
- **Top-level package (no `src/`)**: functionally equivalent, but `src/` prevents accidental imports of the editable working tree when the package is installed, which is the modern Python convention.

## Consequences
- ✅ Enforceable layer boundaries: `ui/` depends on `charts/` and `parsers/`, not vice versa.
- ✅ Files shrink to ~150 lines each, improving readability and unit-test scoping.
- ✅ Aligns with modern Python project structure (`pyproject.toml` + `src/` layout), easing packaging later.
- ⚠️ Requires thin `pages/` wrappers at the repo root to satisfy Streamlit's multi-page app convention.
- ⚠️ Produces a large Phase 1 diff (moves + renames) that must be reviewed carefully to avoid regressions.
