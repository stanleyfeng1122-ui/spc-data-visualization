# HTML Report Rules

Use HTML reports when the answer is dense enough that Markdown becomes tiring to read. Keep durable project knowledge in Markdown, and use HTML as the visual reading layer.

Source context:
- `/Users/zhefeng/Desktop/Vibe Coding/MD Vault/AI Notes/concept/2026-05-11-html-vision-preferred-ai-output-karpathy.md`
- `/Users/zhefeng/Desktop/Vibe Coding/MD Vault/wiki/claude-code-workflows/claude-code-power-patterns.md`

## Default Rule

Create an HTML artifact for long or complicated outputs where layout improves understanding.

Examples:
- Architecture reviews
- Long test reports
- Codebase maps
- Multi-option planning reports
- Bug triage boards
- Feature specs with dependencies
- Agent-team result summaries
- Visual QA reports with screenshots
- Before/after comparisons

Do not use HTML for:
- Short answers
- Simple command results
- Small bug fixes
- Permanent Obsidian notes
- Anything where clean git diffs matter more than readability

## Required Structure

Every HTML report should include:

1. **Executive summary** — the answer in 5-8 bullets.
2. **Status strip** — scope, date, confidence, and decision needed.
3. **Navigation** — sticky sidebar or top anchors for long reports.
4. **Visual hierarchy** — sections, cards, callouts, tables, and diagrams.
5. **Evidence section** — files, screenshots, test results, source notes, or observed behavior.
6. **Decision section** — what needs user review or approval.
7. **Action plan** — next steps split by priority.
8. **Export block** — copyable Markdown or JSON summary when useful.

## Design Rules

- Make the report scannable before it is detailed.
- Put the most important conclusion at the top.
- Use tables for comparisons and priority lists.
- Use cards only for repeated items, not every section.
- Use color as signal: red for blockers, amber for risks, green for verified, blue for information.
- Avoid decorative visuals that do not improve comprehension.
- Keep body text narrow enough to read comfortably.
- Use sticky navigation for reports longer than roughly four sections.
- Include a print-friendly layout.

## Content Rules

- Lead with findings, not process.
- Separate observed facts from recommendations.
- Mark uncertainty explicitly.
- Use concrete file paths and test names.
- Do not paste huge logs. Summarize and link to the artifact/source.
- Keep recommendations actionable.
- Keep the final action list short enough to execute.

## Report Types

### Codebase / Architecture Report

Use when explaining how the app works or where to refactor.

Must include:
- Module map
- Main data flow
- Current pain points
- Recommended seams/interfaces
- Testability impact
- Risk level

### Bug / Debug Report

Use when a bug has multiple possible causes.

Must include:
- Reproduction path
- Expected vs actual behavior
- Evidence observed
- Root cause
- Fix plan
- Regression tests needed

### Test / Visual QA Report

Use after app changes.

Must include:
- Test environment
- Unit test status
- Live UI status
- Visual chart evidence
- Screenshots or screenshot paths
- Pass/fail conclusion

### Planning / Feature Report

Use before implementing non-trivial features.

Must include:
- User problem
- Proposed workflow
- Out-of-scope list
- Implementation slices
- Test plan
- Open decisions

### Agent-Team Report

Use when multiple agents worked in parallel.

Must include:
- Agent roster
- Assignment per agent
- Outputs received
- Lead-agent review
- Conflicts or disagreements
- Final recommendation

## Trigger Heuristic

Use HTML if two or more are true:

- More than 6 sections are needed.
- The output compares 3+ options.
- The output references 5+ files.
- The output includes screenshots or charts.
- The user needs to review evidence.
- The report will be reused later.
- A table, diagram, or dashboard would explain faster than prose.

## Markdown Companion

For knowledge-worthy work, also create or update a Markdown companion file. The Markdown file stores durable decisions. The HTML file is for reading and review.

Recommended pairing:

```text
docs/communication/<topic>.md       # durable rules / decisions
artifacts/<topic>.html              # visual report
```

## User Review Rule

For Zhefeng's workflow, HTML reports should reduce reading load, not increase it.

The first screen should answer:

1. What changed?
2. Why does it matter?
3. What should I review?
4. What is the next action?

