# Loop plan: C1 (SpcDataset seam) + golden regeneration

Working checklist for an autonomous `/loop`. Each iteration: read this file,
do the **first unchecked step**, run the verification gate, commit, tick the
box, end the iteration. One step per iteration — never batch.

## Verification gate (every iteration)

1. `.venv/bin/python -m pytest tests/test_spc_charts.py -q` → must be green.
2. If the step touched `src/spc_viz/charts/` or `parsers/`:
   `.venv/bin/python tests/snapshot_test.py` → all 15 **actual hashes** must
   equal the previous iteration's hashes (fidelity is hash-vs-hash; the
   goldens are stale until step G2 and must NOT be used as the reference).
3. ruff + mypy on changed files (tools in
   `/Users/zhefeng/Desktop/Vibe Coding/Data_Visualiztion_refactor/.venv/bin`)
   → no NEW error codes vs HEAD.
4. Commit with a `refactor(c1):`-prefixed message. Never push.

**On any red:** do NOT force it green by weakening tests. Stop the loop,
report the failure with output, leave the working tree for inspection.

## Golden review (human-gated — the loop may prepare, never approve)

- [x] G1. (2026-07-03) Render all 15 snapshot actuals; build a side-by-side review page
      (golden vs actual per scenario, no-space path) and surface its path in
      chat. Do NOT run `--update-goldens`.
- [ ] G2. **Only after the user has explicitly approved in chat** (their
      message, not inferred): run the golden update, re-run the snapshot
      suite (15/15 must pass), commit as `test: bless goldens after approved
      visual changes`. If no approval message exists yet, skip and move on —
      re-check next iteration.

## C1 increments (in order; each is one iteration)

- [x] C1-1. (2026-07-03) Deduplicate level/condition detection: `charts/base.py`
      `_source_level`/`_source_condition` become imports of
      `parsers.pairing._detect_source_level`/`_detect_source_condition`
      (re-export from pairing under the public names if needed). Pure
      dedup — zero behavior change.
- [ ] C1-2. Introduce the deep module: `parsers/dataset.py` with
      `SpcDataset` and `load_dataset(sources, sheet_names) -> SpcDataset`.
      It owns: per-sheet parsing (`parse_excel_multi`), the parsed-file
      records (dicts stay INTERNAL), `build_paired_dimension_map`, source
      metadata, and exposes: `.dimensions` (OrderedDict), `.display_labels`
      (via ui-independent label logic — move `build_display_map`'s pure part
      here), `.meta_columns`, `.combined(dim_nos)` →
      `prepare_combined_data` result. Unit tests against
      `examples/*.xlsx` fixtures. No caller migrates yet.
- [ ] C1-3. Migrate `app.py`: the dict-building block + sheet parse loop
      call `load_dataset`; `st.cache_data` wraps the loader. Behavior
      identical (snapshot hash gate applies).
- [ ] C1-4. Migrate `pages/1_Quick_Test.py` the same way.
- [ ] C1-5. Migrate `ui/batch_export.py` onto the dataset's
      `.combined(...)` path so export parity with the main chart is by
      construction (absorbs review candidate C5).
- [ ] C1-6. Close the seam: grep gate — no `pf["` / `pf.get(` outside
      `src/spc_viz/parsers/`; delete now-dead dict plumbing; add the
      `SpcDataset` term to CONTEXT.md; restart the launchd app
      (`launchctl kickstart -k gui/$UID/com.spc.data-visualization`, poll
      `:8503/_stcore/health` = 200, confirm new PID).

## Loop hygiene

- Branch: work directly on `codex/chart-visual-fixes-2026-05-19`.
- If context is running low mid-step, finish the verification gate and
  commit before ending the iteration — never leave a red tree between
  iterations.
- Loop END condition: every box above ticked (G2 may legitimately remain
  unticked if the user never approves — after C1-6, if G2 is still blocked,
  report that and end the loop rather than idling forever).
