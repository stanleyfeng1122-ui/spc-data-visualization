# SPC App Improvement Workflow

This app should improve through small, visible work cards. Do not let the chat thread become the product spec, bug tracker, test plan, and release log at the same time.

## Operating Principle

One issue becomes one work card. One work card becomes one focused implementation. One implementation gets one test report.

```text
Find issue -> Capture card -> Rank priority -> Diagnose -> Implement -> Test -> Review -> Commit
```

## Improvement Lanes

Use these lanes to organize bugs and improvements.

| Lane | What Belongs Here | Example |
|---|---|---|
| Chart Correctness | Anything that makes the visual output wrong or misleading | Wrong spec line, blank chart, wrong grouping |
| Parser Robustness | File/header/vendor-format detection problems | Metadata columns not detected, vendor typo handling |
| Workflow Speed | Anything that removes repetitive manual work | Batch export, saved defaults, fewer clicks |
| Visual QA | Testing how charts actually look | Golden screenshots, not-blank chart checks |
| Reliability | Server, environment, dependency, restart issues | App not running after reboot |
| Code Structure | Internal cleanup that makes future changes safer | Moving chart logic behind smaller interfaces |

## Priority Rules

| Priority | Meaning | Fix Timing |
|---|---|---|
| P0 | App cannot be used | Stop and fix first |
| P1 | Output can be wrong | Fix before trusting results |
| P2 | Work is slower than it should be | Batch into workflow improvements |
| P3 | Cosmetic or nice-to-have | Do only when nearby |

Most chart and parser issues are P1 because they affect trust.

## Work Card Template

Use this every time a new bug or limitation appears.

```markdown
## Title

## Type
Bug / Limitation / UX improvement / Parser issue / Chart rendering issue / Export issue

## Priority
P0 / P1 / P2 / P3

## What I was doing
Page:
File:
Sheet:
Dimension:
Chart type:
Controls selected:

## What happened

## What I expected

## Evidence
Screenshot:
File path:
Notes:

## Pass condition
How we know the issue is fixed.

## Out of scope
What should not be changed while fixing this.
```

## Agent Handling Rule

When a work card is submitted, the lead agent should:

1. Restate the card in one short paragraph.
2. Identify the likely lane and priority.
3. Inspect the relevant code and files before editing.
4. Propose the smallest fix boundary.
5. Add or update tests when possible.
6. Implement only this card.
7. Run the SPC test agent.
8. Produce a short visual evidence report.

## Test Evidence Required

Each completed card needs evidence that matches the risk.

| Issue Type | Required Evidence |
|---|---|
| Chart correctness | Unit test plus live chart screenshot or golden snapshot |
| Parser robustness | Test file fixture or real-file parsing check |
| Batch export | ZIP generation test plus one exported PNG check |
| UI workflow | Live UI smoke test |
| Server/reliability | Process/port check plus app HTTP response |
| Code refactor | Existing tests pass, no user-facing behavior changed |

## Current Recommended Roadmap

### Phase 1: Trust the Chart

Goal: charts must never be visually misleading.

- Different USL/LSL values across selected dimensions render clearly.
- Single-point dimensions render visible marks.
- Box Plot and Histogram handle multi-dimension spec limits correctly.
- Labels stay readable when multiple dimensions share a spec value.
- Visual test snapshots cover core chart modes.

### Phase 2: Trust the File Parser

Goal: vendor file variations should be detected without manual mapping.

- Header detection handles common vendor variations.
- Metadata/factor columns are detected reliably.
- Vendor typos are surfaced as suggestions, not silently ignored.
- The app shows what it detected and what it could not detect.

### Phase 3: Reduce Repetition

Goal: fewer manual operations during real QE work.

- Batch chart export stays stable.
- Last-used controls can be reused safely.
- Common dimension sets can be saved as presets.
- Export naming is predictable and report-ready.

### Phase 4: Make Testing Routine

Goal: every feature gets the same pass/fail treatment.

- Dedicated SPC test agent runs after changes.
- HTML test reports summarize unit, UI, and visual chart evidence.
- Golden screenshots cover important chart states.
- Failures explain what is wrong, not just that something failed.

### Phase 5: Keep Code Navigable

Goal: future changes stay easy for agents and humans.

- Chart, parser, UI, export, and state logic stay separated.
- New behavior is added behind small interfaces.
- Refactors are done only when they reduce real friction.
- Product decisions are recorded in docs, not buried in chat.

## Daily Usage Pattern

When using the app, capture issues as soon as they appear:

```text
Title:
Priority:
File:
Sheet:
Dimension:
Chart type:
What happened:
Expected:
Screenshot:
```

At the start of a development session, pick only one P0/P1 card or up to three related P2 cards.

## What Not To Do

- Do not bundle unrelated bugs into one implementation.
- Do not redesign the UI unless the current workflow is actually blocked.
- Do not add analysis/calculator features in this visualization-first version.
- Do not accept tests without visual chart evidence for chart changes.
- Do not commit large cleanup mixed with behavior fixes.

