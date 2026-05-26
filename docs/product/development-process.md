# Development Process for the SPC App

This project should use a lightweight control system, not a heavy software process.

## Role Split

Zhefeng's role:
- Product owner
- Quality engineer
- Final judge

Agent role:
- Planner
- Implementer
- Tester
- Explainer

## Core Loop

```text
Observe -> Capture -> Select -> Build -> Review
```

1. **Observe** — use the app normally and send rough issues, screenshots, or files.
2. **Capture** — agent turns rough issues into work cards.
3. **Select** — pick one P1/P2 card to fix.
4. **Build** — agent implements, tests, and commits.
5. **Review** — user checks the app visually and confirms whether it matches expectation.

## Session Start Checklist

Every development session should begin with:

```text
Current branch:
Last commit:
Uncommitted changes:
Today’s target:
Work card:
Test plan:
Commit plan:
```

## Session End Checklist

Every completed fix should end with:

```text
Changed files:
Tests run:
Known risk:
Commit hash:
Next recommended card:
```

## Change Rule

Every meaningful change needs:

```text
Work card -> code -> test -> commit
```

For bigger features:

```text
Mini plan -> work cards -> implementation slices -> test report -> commit
```

## Do Not Do

- Do not fix ten things in one change.
- Do not keep editing without commits.
- Do not mix planning, refactoring, testing, and feature work without a checkpoint.
- Do not write heavy specs for every small issue.

## Current Next Card

```text
WC-005: Multi-factor Section-by should render as nested JMP-style headers.
```

