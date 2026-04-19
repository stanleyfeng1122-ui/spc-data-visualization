# ADR-0004: openpyxl monkey-patch for strict-OOXML

**Status:** Accepted (with technical debt) — 2026-04-18
**Deciders:** Stanley Feng

## Context
Some vendor-supplied xlsx files use the strict-OOXML namespace variant or contain malformed `ExternalReference` entries that cause openpyxl 3.1.x to crash during load. Users cannot easily pre-convert these files, and the tool must read them to be useful.

## Decision
Apply a local monkey-patch to `openpyxl.packaging.workbook.ExternalReference` and rewrite strict-OOXML namespaces to transitional OOXML in memory before handing the stream to openpyxl.

## Alternatives considered
- **Wait for upstream fix**: openpyxl release cadence is slow and this would block real users today.
- **pandas.read_excel with xlrd**: xlrd dropped xlsx support in 2.0 — not an option.
- **Require users to pre-convert files**: shifts the burden onto users and defeats the one-click analysis UX.

## Consequences
- ✅ Problematic vendor files "just work" without user intervention.
- ✅ Patch is isolated to the parser entry point, so the rest of the codebase stays clean.
- ⚠️ The monkey-patch is brittle to openpyxl version upgrades — any internal refactor upstream can silently break it.
- ⚠️ Coverage is still incomplete: LK-formatted files surface a `'NoneType' object has no attribute 'Target'` error that the current patch does not handle.
- ⚠️ Represents acknowledged technical debt; should be revisited when upstream ships proper strict-OOXML support.
