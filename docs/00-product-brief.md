# Product Brief: SPC Data Visualization App

## The User

Manufacturing quality engineers (QEs) who analyze dimension measurement data from inspection labs. Built initially for a single user — the author — with the intent to stay focused and opinionated rather than generic.

## The Problem

Every new product launch generates a stack of vendor inspection reports: five or more Excel workbooks, each with multiple sheets, each holding dozens of measured dimensions across CORR, POR, PP, and AP CPK stages (**CPK** = Process Capability Index; a score of how well a process stays within spec).

> [!warning] Before this tool
> - Open 5+ `.xlsx` files side by side
> - Manually build a chart for each dimension in Excel
> - Redraw spec limits by hand
> - Screenshot every chart, paste into a report
> - Reconcile inconsistent sheet names, header rows, and column layouts across vendors
>
> Result: hours of tedious, repeatable work per product — before any actual analysis begins.

## The Solution

A local Streamlit app that turns raw vendor Excel files into a publishable chart pack in seconds.

- **Ingest** multi-sheet Excel workbooks with mixed vendor formats
- **Auto-detect** dimensions, metadata columns, and header rows — no hand-mapping
- **Render** interactive Plotly charts with Statistical Process Control (**SPC**) reference lines: **USL** (Upper Spec Limit), **LSL** (Lower Spec Limit), and Nominal (target)
- **Chart types**: Combined Profile (one line per unit across all dimensions), Box Plot, and Histogram
- **Group and slice** with color-by, section-by, and row-by controls; exclude outlier points with a click
- **Batch export** every chart in a selection as a single ZIP of PNGs, ready to drop into a report

Three focused pages:

- **Home** — the main multi-sheet charting workflow
- **Quick Test** — a scratchpad for one-off plots from a single file
- **Sheet Manager** — compares dimension coverage across sheets so nothing gets silently dropped

## Why It Exists

Roughly 80% of a QE's charting work is mechanical: same chart types, same spec lines, same export format, different data. This app replaces that mechanical layer. What took an afternoon of opening files, dragging columns, and pasting screenshots now takes the time it takes to upload the workbooks and click Export. Hours become seconds. The engineer's attention moves from producing charts to reading them.

It handles the formats that actually arrive in the inbox — 5x Corrugated flatness, FX K116 PP, LK X3744 CORR/POR, TY K116, TRM X3083, FXJS X3744 — without requiring the vendor to change anything.

## What It Is Not

> [!info] Explicit non-goals
> - **Not multi-user.** No accounts, no sharing, no permissions. One engineer, one laptop.
> - **Not cloud-hosted.** Runs locally on Python 3.10. Data never leaves the machine.
> - **Not a statistical analysis suite.** No ANOVA, no regression, no hypothesis testing. Visualization only.
> - **Not a CPK calculator.** CPK computation remains owned by the QE team's existing process and tooling. This app visualizes measurements against limits — it does not certify capability.

The scope is deliberately narrow: do the repetitive charting job well, and stay out of the way of everything else.
