# Phase 0 Implementation Plan — Design Documentation

> **For Claude:** REQUIRED SUB-SKILL: Use `superpowers:executing-plans` to implement this plan task-by-task.

**Goal:** Produce Phase 0 planning documents (Product Brief, User Stories, Architecture, Design Doc, 5 ADRs, READMEs) for the SPC Data Visualization app — without changing any source code.

**Architecture:** Multi-agent with lead/supervisor pattern. Lead (Claude main thread) dispatches specialist sub-agents (planner, architect, doc-updater), reviews outputs via 3 quality gates (structural, content, integration), presents audit questions **one at a time** to user, commits approved docs to `docs/phase-0-design` branch.

**Tech Stack:** Git 2.x, Git worktrees (optional for docs), Markdown with Mermaid diagrams, MADR-lite ADR format, Claude agents (`planner`, `architect`, `doc-updater`, `spc-test-runner`).

**Execution model:** Strict waterfall. All Phase 0 work done and merged before Phase 1 starts.

**Audit protocol:** Lead asks one question per message. User replies short (e.g., "A" or "yes, accurate"). Lead moves to next question.

**Source of truth:** [`docs/plans/2026-04-18-spc-restructure-design.md`](./2026-04-18-spc-restructure-design.md) — the design this plan implements.

---

## Timeline

| Phase | Tasks | Est. time |
|---|---|---|
| Prep | P1–P3 | ~15 min |
| Phase 0 Wave 1 | T1–T4 (planner + architect parallel) | ~45 min |
| Phase 0 Wave 2 | T5–T7 (architect continues) | ~30 min |
| Phase 0 Finalize | T8–T10 (doc-updater, test-runner, merge) | ~20 min |
| **Total** | **10 tasks** | **~2 hours** |

---

## Prep: Pre-Phase-0 Setup

### Task P1: Tag the current working state as baseline

**Files:** none (git tag only)

**Step 1:** Verify current branch and clean status
```bash
cd "/Users/zhefeng/Desktop/Vibe Coding/Data Visualiztion"
git status
git log -1 --oneline
```
Expected: on `feature/code-quality-refactor`, latest commit is `418d161 docs: add restructure design doc`. No uncommitted source changes (untracked files in .superset/, .venv/, logs/ are fine).

**Step 2:** Create annotated tag
```bash
git tag -a v0.1-baseline-stable -m "Baseline: working SPC Data Viz app before restructure. Batch export, parser fix, shared_ui refactor all merged."
```

**Step 3:** Verify tag
```bash
git tag | grep v0.1
git show v0.1-baseline-stable --stat | head -20
```
Expected: tag exists, points to commit `418d161`.

**Step 4:** (Optional) Push tag to remote
```bash
git remote -v
# If remote exists:
git push origin v0.1-baseline-stable
# If no remote: skip, note that tag is local-only
```

**Step 5: Audit checkpoint with user**
> Lead asks: "Baseline tag `v0.1-baseline-stable` created. Do you have a git remote I should push to, or skip this step?"

---

### Task P2: Merge `feature/code-quality-refactor` → `main`

**Files:** none (git merge only)

**Step 1:** Check if `main` branch exists locally
```bash
git branch | grep -E "^[* ]+main$" || echo "main branch missing"
```

**Step 2:** Switch to main (create if missing)
```bash
# If main exists:
git checkout main
git pull --ff-only 2>/dev/null || true

# If main missing (local-only repo):
git checkout -b main
```

**Step 3:** Merge the feature branch
```bash
git merge --no-ff feature/code-quality-refactor -m "Merge feature/code-quality-refactor: batch export + parser fix + shared_ui + design doc"
```
Expected: clean merge (no conflicts), tests still pass.

**Step 4:** Re-point baseline tag to merge commit (optional — use `v0.1-baseline-stable-merged` to avoid rewriting history)
```bash
# Skip unless user wants a cleaner baseline pointing to main.
```

**Step 5:** Run full test suite to confirm nothing broken post-merge
```bash
cd "/Users/zhefeng/Desktop/Vibe Coding/Data Visualiztion"
source .venv/bin/activate
pytest tests/ -v --tb=short
```
Expected: 24/24 tests passing.

**Step 6: Audit checkpoint with user**
> Lead asks: "Feature branch merged to main. All 24 tests pass. Ready to create `docs/phase-0-design` branch?"

---

### Task P3: Create `docs/phase-0-design` branch

**Files:** none (git branch only)

**Step 1:** Create and switch to new branch from main
```bash
git checkout -b docs/phase-0-design
git branch --show-current
```
Expected: output `docs/phase-0-design`.

**Step 2:** Verify starting state is same as v0.1-baseline-stable + design doc
```bash
git log --oneline -3
```
Expected: shows the merge commit and `418d161 docs: add restructure design doc`.

**Step 3:** (Informational only — no commit) Confirm branch has clean code state

**Decision point — audit question to user:**

> Lead asks: "On `docs/phase-0-design` branch, ready to start Phase 0. First deliverable will be the Product Brief via the `planner` agent. Dispatch now?"
> - A) Yes, dispatch
> - B) Pause — something to adjust first

---

## Phase 0 Wave 1 — Parallel agents (Product Brief + User Stories + Architecture)

### Task T1: Scaffold `docs/` folder structure

**Files:**
- Create: `docs/README.md` (index, stub)
- Create: `docs/adr/` (empty folder with `.gitkeep`)

**Step 1:** Create docs/ tree
```bash
cd "/Users/zhefeng/Desktop/Vibe Coding/Data Visualiztion"
mkdir -p docs/adr
touch docs/adr/.gitkeep
```

**Step 2:** Write placeholder `docs/README.md`
```markdown
# SPC Data Viz — Documentation

Index of design documents.

- [Product Brief](./00-product-brief.md) — who this is for, what it does, why
- [User Stories](./01-user-stories.md) — features as agile stories
- [Architecture](./02-architecture.md) — module + data flow diagrams
- [Design Doc](./03-design-doc.md) — tech decisions & why
- [ADRs](./adr/) — architecture decision records
- [Plans](./plans/) — phase implementation plans
```

**Step 3:** Commit scaffolding
```bash
git add docs/README.md docs/adr/.gitkeep
git commit -m "docs: scaffold docs/ folder structure"
```

**Step 4:** Verify
```bash
tree docs/ 2>/dev/null || find docs/ -print
```

**No audit needed** — this is pure scaffolding.

---

### Task T2: Dispatch `planner` agent for Product Brief

**Files:**
- Create: `docs/00-product-brief.md`

**Step 1: Lead prepares the dispatch brief**

Brief contents (used in Agent tool prompt):
```
TASK: Write Product Brief

GOAL
Produce a 1-page "working backwards" style product brief for the SPC Data
Visualization app.

SCOPE (must cover)
- The user (who)
- The problem (before this tool existed)
- The solution (what this tool does)
- Non-goals (what this tool is NOT)

ACCEPTANCE CRITERIA
- Under 500 words
- Plain language (no SPC jargon without definition)
- Uses all 4 sections above
- Written as if pitching to a new user seeing the app for the first time
- Saved to: docs/00-product-brief.md

CONTEXT
- Current code: /Users/zhefeng/Desktop/Vibe Coding/Data Visualiztion/
- User description (verbatim): "helps with home-consuming, tedious, repeatable
  works" (user is a quality engineer analyzing manufacturing dimension data)
- Related docs: docs/plans/2026-04-18-spc-restructure-design.md (design doc)

FORMAT
Markdown. Use Obsidian-style callouts if helpful (> [!info]).

DO NOT
- Invent features not in the current code
- Use >500 words
- Add roadmap / future features (out of scope for Phase 0)
```

**Step 2: Dispatch via Agent tool**
```
Agent tool call:
  subagent_type: planner
  description: Write Product Brief
  prompt: <the brief above>
```

**Step 3: Lead reviews output against 3 quality gates**

- Structural gate: file exists at `docs/00-product-brief.md`? Has 4 required sections? Under 500 words?
- Content gate: any hallucinated features? Jargon defined?
- Integration gate: consistent with design doc?

**Step 4: If all gates pass — prepare audit question**

> Lead asks (one question only):
> **"Phase 0 Wave 1 — Audit (Q1)**: On the Product Brief, I described your user as `[quote from generated brief]`. Right?
> - A) Yes, move on
> - B) No, actual user description is: ___"

**Step 5: Handle response**
- If "A" → commit the file
- If "B" → revise file with corrected user description → commit

**Step 6: Commit**
```bash
git add docs/00-product-brief.md
git commit -m "docs: add Product Brief (via planner agent)"
```

---

### Task T3: Dispatch `planner` agent for User Stories

**Files:**
- Create: `docs/01-user-stories.md`

**Step 1: Lead prepares dispatch brief**

```
TASK: Write User Stories

GOAL
Enumerate every user-facing feature of the SPC Data Viz app as an agile user
story with acceptance criteria.

SCOPE
Read these source files to discover features:
- app.py
- shared_ui.py
- chart_utils.py
- pages/1_Quick_Test.py
- pages/2_Sheet_Manager.py

For each distinct feature, write:
- Story ID (US-001, US-002, ...)
- "As a ___, I want to ___, so that ___"
- Acceptance criteria (checklist of verifiable behaviors)
- Note: "✅ implemented" or "⚠️ ambiguous — see notes" or "❌ planned, not implemented"

ACCEPTANCE CRITERIA
- 10-15 user stories (one per distinct feature)
- Each has all 4 fields above
- Ambiguous cases flagged for user clarification
- Saved to: docs/01-user-stories.md

CONTEXT (same as T2 plus)
- User's existing features include: multi-file xlsx upload, dimension
  auto-detect, combined profile / box plot / histogram chart types,
  color-by / section-by / row-by grouping, spec limits (USL/LSL/Nominal),
  batch chart export as ZIP, point exclusion, Quick Test page, Sheet Manager page

DO NOT
- Invent features not present in code
- Merge multiple features into one story
```

**Step 2: Dispatch**
```
Agent tool: subagent_type=planner, description="Write User Stories"
```

**Step 3: Lead reviews against gates**
- Structural: ≥10 stories? Each has all 4 fields?
- Content: any story references a feature that isn't in the code? Any ambiguities clearly flagged?
- Integration: story language consistent with Product Brief?

**Step 4: Lead extracts ambiguities → becomes audit questions**

Lead generates one audit question per ambiguous story. Examples:

> Q2 of N: "US-008 'Color-by grouping' — I see it in the code but can't tell
> if it's from a dropdown or auto-detected. Which?
> - A) User picks from dropdown
> - B) Auto-detected from data
> - C) Both"

**Step 5: Handle each user response** (one question at a time)
- Update the story with the chosen interpretation
- Move to next question

**Step 6: Commit**
```bash
git add docs/01-user-stories.md
git commit -m "docs: add User Stories (via planner agent) with N clarifications"
```

---

### Task T4: Dispatch `architect` agent for Architecture doc

**Files:**
- Create: `docs/02-architecture.md`

**Step 1: Lead prepares dispatch brief**

```
TASK: Write Architecture Doc

GOAL
Produce a visual-first architecture doc (Mermaid diagrams + short prose)
documenting the CURRENT state of the code (not the target Phase 1 state).

DELIVERABLE
- Module diagram (Mermaid graph TB) showing how files import each other
- Data flow diagram (Mermaid sequenceDiagram) showing what happens when
  a user uploads a file and views a chart
- Short prose (~200 words) explaining the diagrams

ACCEPTANCE CRITERIA
- Both diagrams render correctly as Mermaid
- Module diagram shows: app.py, shared_ui.py, chart_utils.py, spc_parser.py,
  ui_theme.py, pages/*.py, external deps (Streamlit, Plotly, kaleido, openpyxl)
- Data flow covers: upload → parse → render chart
- Under 2 pages when rendered
- Saved to: docs/02-architecture.md

DO NOT
- Describe the TARGET Phase 1 architecture (that's in the design doc)
- Include code samples (doc is visual + prose only)
```

**Step 2: Dispatch** (can run in parallel with T2 and T3 — lead dispatches all three at once)
```
Agent tool: subagent_type=architect, description="Write Architecture Doc"
```

**Step 3: Lead reviews gates**
- Structural: 2 Mermaid diagrams? Prose section?
- Content: diagrams reflect actual imports in the code?
- Integration: consistent with design doc's current-state description?

**Step 4: Audit question** (only one expected)

> "Q_N: The architecture diagram shows `ui → (charts + parsers) → config`. Does this match your mental model?
> - A) Yes
> - B) No, actually ___"

**Step 5: Commit**
```bash
git add docs/02-architecture.md
git commit -m "docs: add Architecture diagrams (via architect agent)"
```

---

## Phase 0 Wave 2 — Architect continues (Design Doc + ADRs)

### Task T5: Dispatch `architect` agent for Design Doc

**Files:**
- Create: `docs/03-design-doc.md`

**Step 1: Lead prepares dispatch brief**

```
TASK: Write Tech Design Doc (Google/Meta style)

GOAL
Explain the CURRENT technical design of the SPC Data Viz app: why each
major technology was chosen, what tradeoffs were made, what the constraints
were.

SECTIONS (all required)
- Context — why this app exists, what it replaces
- Goals — primary success criteria
- Non-Goals — explicit exclusions
- Current Technical Design — tech stack, data model, key algorithms
- Alternatives Considered — briefly, for each major choice
- Open Technical Debt — issues documented for future resolution

ACCEPTANCE CRITERIA
- ~2 pages rendered (not a novel)
- Every major tech choice has an alternative considered
- Open technical debt explicitly listed (e.g., the openpyxl monkey-patch)
- Saved to: docs/03-design-doc.md
```

**Step 2: Dispatch**
**Step 3: Lead reviews gates**
**Step 4: Audit questions** (likely 1-2 based on debt items flagged)
**Step 5: Commit**

```bash
git add docs/03-design-doc.md
git commit -m "docs: add Design Doc (via architect agent)"
```

---

### Task T6: Dispatch `architect` agent for 5 ADRs

**Files:**
- Create: `docs/adr/0001-streamlit-over-flask.md`
- Create: `docs/adr/0002-plotly-over-matplotlib.md`
- Create: `docs/adr/0003-kaleido-for-png-export.md`
- Create: `docs/adr/0004-openpyxl-monkey-patch.md`
- Create: `docs/adr/0005-src-layout-for-phase-1.md`

**Step 1: Lead prepares dispatch brief**

```
TASK: Write 5 ADRs (MADR-lite format)

FORMAT PER ADR (each ~1 page)
# ADR NNNN: Short Title

## Status
Accepted — YYYY-MM-DD (or Proposed for future ADRs)

## Context
1-3 sentences: what problem this addresses.

## Decision
1-2 sentences: what we chose.

## Alternatives considered
- Option A: why not chosen
- Option B: why not chosen

## Consequences
- ✅ Benefits
- ⚠️ Tradeoffs / limitations

DELIVERABLES
1. 0001-streamlit-over-flask.md
2. 0002-plotly-over-matplotlib.md
3. 0003-kaleido-for-png-export.md
4. 0004-openpyxl-monkey-patch.md (flag as Accepted but with technical debt)
5. 0005-src-layout-for-phase-1.md (flag as Proposed — not yet executed)

SAVE TO: docs/adr/
```

**Step 2: Dispatch**
**Step 3: Lead reviews gates** (each ADR has all 5 sections?)
**Step 4: Audit questions** — 0-1 expected (ADRs are mostly factual)
**Step 5: Commit**

```bash
git add docs/adr/
git commit -m "docs: add 5 ADRs (via architect agent)"
```

---

### Task T7: Dispatch `doc-updater` agent for READMEs

**Files:**
- Create: `README.md` (root-level, user-facing)
- Update: `docs/README.md` (fill in index now that all docs exist)

**Step 1: Lead prepares dispatch brief**

```
TASK: Write root README.md and finalize docs/README.md index

GOAL
Produce a friendly root README a new user would see first.

ROOT README SECTIONS
- One-line project pitch
- Screenshot or GIF placeholder
- Quick start (how to run locally)
- Features (link to user stories for detail)
- Architecture (link to docs/02-architecture.md)
- Development (link to docs/plans/)

docs/README.md (update existing)
- Replace placeholder bullets with real links to all docs that now exist

DO NOT
- Duplicate content from Product Brief (link instead)
- Include "roadmap" or future features
```

**Step 2: Dispatch**
**Step 3: Lead reviews gates**
**Step 4: Audit question** — 0-1 expected
**Step 5: Commit**

```bash
git add README.md docs/README.md
git commit -m "docs: add root README and finalize docs index"
```

---

## Phase 0 Finalize — Verification & Merge

### Task T8: Run `spc-test-runner` agent

**Files:** none (verification only — Phase 0 changed no code)

**Step 1: Dispatch test runner**
```
Agent tool: subagent_type=spc-test-runner
description: "Verify no code regressions after Phase 0 docs work"
prompt: "Run the full test suite and verify all 24 tests still pass.
Report any that fail. Phase 0 changed no source code — any test regression
would mean someone accidentally modified code, which is a bug to flag."
```

**Step 2: Lead reviews output**
- Expected: 24/24 pass
- If any fail: escalate to user immediately

**Step 3:** No audit question needed unless failures occur.

---

### Task T9: Final audit summary with user

**Step 1: Lead prepares final Phase 0 audit card**

> **Phase 0 Complete — Final audit (1 of 1)**
>
> Phase 0 is done. Summary:
> - 8 new docs created in `docs/`
> - 6 commits on `docs/phase-0-design` branch
> - 0 code changes (verified by test runner, 24/24 tests pass)
>
> Ready to merge `docs/phase-0-design` → `main` and tag `v0.2-designed`?
>
> - A) Yes, merge and move to Phase 1
> - B) Pause — I want to read the docs first
> - C) Changes needed to: ___

**Step 2:** Handle user response
- A → proceed to T10
- B → user reads; lead waits for second approval
- C → revise docs → re-audit → back to A

---

### Task T10: Merge `docs/phase-0-design` → `main` and tag

**Files:** none (git operations)

**Step 1:** Switch to main and merge
```bash
git checkout main
git merge --no-ff docs/phase-0-design -m "Merge Phase 0: design documentation complete"
```

**Step 2:** Tag the merged state
```bash
git tag -a v0.2-designed -m "Phase 0 complete: all design docs produced and approved"
```

**Step 3:** Push tag (if remote exists)
```bash
git push origin main v0.2-designed 2>/dev/null || echo "Local-only repo, skipping push"
```

**Step 4:** Confirm state
```bash
git log --oneline -10
git tag
```
Expected: v0.1-baseline-stable and v0.2-designed both present.

**Step 5: Final handoff to user**

> Lead reports: "Phase 0 complete and merged to main. Next step: Phase 1 implementation plan. Shall I generate it now, or pause here?"

---

## Success Criteria for Phase 0

- [ ] `v0.1-baseline-stable` tag exists
- [ ] `v0.2-designed` tag exists
- [ ] `docs/` contains: README, Product Brief, User Stories, Architecture, Design Doc, 5 ADRs
- [ ] Root `README.md` exists
- [ ] 0 code files modified (verified by `git diff v0.1-baseline-stable v0.2-designed -- '*.py'` = empty)
- [ ] All 24 tests still pass
- [ ] User has approved every doc via one-question-at-a-time audit

---

## Rollback

If anything goes wrong during Phase 0:
```bash
git checkout v0.1-baseline-stable
# You're back to exactly the state before Phase 0 started.
```
