# SP Search Launch-Readiness Audit Production Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce the SP Search launch-readiness audit document specified in `docs/superpowers/specs/2026-05-02-launch-readiness-audit-design.md` in a single execution session, archive the prior audit, and meet every acceptance criterion in the spec without modifying runtime source code.

**Architecture:** The audit is a single document at `docs/sp-search-launch-readiness-audit.md` with six Parts and five Appendices (A–E). The plan executes the spec's eight methodology passes in a controlled sequence: scaffold → reconciliation → toolkit comparison → PnP v4 parity → journey simulations → differentiator tracks → foundations sweep → roadmap matrix → sprint sequencing → self-review → final commit. Each task produces a committed artifact; no task is allowed to mutate runtime source code.

**Tech Stack:** Markdown only. Read-only tools against the SP Search repo (`/Users/hemantmane/Development/sp-search`) and the spfx-toolkit repo (`/Users/hemantmane/Development/spfx-toolkit`). External fetches against PnP Modern Search v4 docs. Verification commands listed in spec §5.2 (`git status --short`, `git rev-parse --short HEAD`, `npm run type-check`, `npm test`, `npm run package`) are run for evidence and their results are recorded in Appendix E — they do not produce runtime source changes.

**Spec reference:** `docs/superpowers/specs/2026-05-02-launch-readiness-audit-design.md` (single source of truth — when this plan and the spec disagree, the spec wins).

---

## File Structure

| Path | Action | Responsibility |
|------|--------|----------------|
| `docs/sp-search-launch-readiness-audit.md` | Create | The audit document — six Parts + five Appendices |
| `docs/archive/sp-search-comprehensive-audit-2026-03-22.md` | Create (move from `docs/`) | Archived prior audit with redirect header |
| `docs/sp-search-comprehensive-audit.md` | Delete | Replaced by archived copy + new audit |
| `docs/superpowers/plans/2026-05-02-launch-readiness-audit-production.md` | Create | This plan |

No source code, scripts, configs, or tests are modified by this plan. Verification commands in Phase 0 may regenerate `lib/`, `dist/`, or `temp/` artifacts; those are recorded in Appendix E and excluded from any commits via `.gitignore` (already in place — verify in Task 0.4).

---

## Phase 0 — Preparation

### Task 0.1: Capture audit inputs (repo snapshot)

**Files:**
- Read: `package.json`, `/Users/hemantmane/Development/spfx-toolkit/package.json`
- Create: working scratch notes (held in conversation context until Appendix E is written)

- [ ] **Step 1: Capture repo snapshot**

Run:
```bash
git rev-parse --abbrev-ref HEAD
git rev-parse --short HEAD
git status --short
git log --oneline -5
```

Record output. These become rows in Appendix E "Repo snapshot."

- [ ] **Step 2: Capture sp-search package versions**

Run:
```bash
node -e "const p = require('./package.json'); console.log('sp-search', p.version); Object.entries(p.dependencies).filter(([k]) => k.startsWith('@microsoft/sp-') || k === 'react' || k === 'zustand' || k.startsWith('devextreme') || k.startsWith('@fluentui') || k === 'spfx-toolkit').forEach(([k,v]) => console.log(k, v));"
```

Record output. Goes into Appendix E.

- [ ] **Step 3: Capture spfx-toolkit version + recent commits**

Run:
```bash
cd /Users/hemantmane/Development/spfx-toolkit && node -e "console.log(require('./package.json').version)" && git log --oneline -20
```

Record output. Goes into Appendix B (toolkit integration map source) and Appendix E.

- [ ] **Step 4: Hold scratch notes in context for use in Phase 9**

Do NOT write Appendix E yet. Carry the captured strings in conversation memory or a local scratch file outside the repo (for example `/tmp/sp-search-launch-audit-notes.md`) so Task 9.3 can render them into Appendix E with the verification command results from Task 0.2. Do not create tracked scratch files inside the repository.

### Task 0.2: Run verification commands and capture results

**Files:**
- No file changes; outputs captured to scratch notes for Appendix E.

- [ ] **Step 1: Run type-check and capture result**

Run:
```bash
npm run type-check 2>&1 | tail -50
```

Record: pass/fail, exit code, last 50 lines. If the script does not exist, run `npx tsc --noEmit 2>&1 | tail -50` instead and record which command was used.

- [ ] **Step 2: Run tests and capture result**

Run:
```bash
npm test 2>&1 | tail -50
```

Record: pass/fail/skip, exit code, summary. If the Jest harness fails before running tests, record the failure as evidence and link it to the Foundations track entry; do not rely on uncommitted memory files as evidence.

- [ ] **Step 3: Run package and capture result**

Run:
```bash
npm run package 2>&1 | tail -30
```

Record: pass/fail, generated `.sppkg` path under `sharepoint/solution/`, file size. If the script does not exist, list available scripts (`npm run`) and record what was actually executed.

- [ ] **Step 4: Snapshot working tree after verification**

Run:
```bash
git status --short
```

Record any generated/modified files. They will be enumerated in Appendix E "Generated artifacts during verification" and explicitly NOT committed with the audit.

### Task 0.3: Create audit document scaffold

**Files:**
- Create: `docs/sp-search-launch-readiness-audit.md`

- [ ] **Step 1: Write the scaffold with all section headers**

Create `docs/sp-search-launch-readiness-audit.md` with this exact skeleton (all sections present, prose left as `_(populated in Phase X — see plan)_` markers — these markers are the ONLY allowed placeholders and must all be removed by Task 9.1):

```markdown
# SP Search Launch-Readiness Audit

**Date:** 2026-05-02
**Scope:** Pre-launch audit covering 6 web parts + 1 library component (SP Search)
**Audience profile:** Any SPFx-capable tenant, self-serve, no author hand-holding
**Spec:** `docs/superpowers/specs/2026-05-02-launch-readiness-audit-design.md`

## Front Matter

### Repo Snapshot
_(populated in Phase 9 — see plan Task 9.3)_

### Verification Snapshot
_(populated in Phase 9 — see plan Task 9.3)_

### Differentiator Priorities
1. Modern UI Quality
2. End-User Productivity
3. Multi-Instance / Multi-Context
4. Admin Experience
5. Observable & Diagnosable

### Reconciliation Summary (March 22 → Today)
_(populated in Phase 1 — see plan Task 1.3)_

### Reading Guide
- **Effort tiers:** S (≤4h) · M (½–1d) · L (1–3d) · XL (>3d)
- **Priority tiers:** P0 (must ship v1.0) · P1 (should ship v1.0) · P2 (v1.1+) · Defer
- **P0 admission rule:** A finding may only be P0 if it ties to (a) a stated differentiator T1–T5, (b) security, (c) data integrity, (d) a "would prevent install" issue, or (e) a journey Blocker with no documented workaround.
- **Roadmap Matrix (Part 4)** is the executable artifact — open it to pick the next thing to do.

---

## Part 1 — The Two Journeys

### Journey A: Day 1 Admin Install
_(populated in Phase 4 — see plan Tasks 4.1–4.3)_

### Journey B: Day 1 End-User Search
_(populated in Phase 5 — see plan Tasks 5.1–5.3)_

---

## Part 2 — Differentiator Tracks

### T1. Modern UI Quality
_(populated in Phase 6 — see plan Task 6.1)_

### T2. End-User Productivity
_(populated in Phase 6 — see plan Task 6.2)_

### T3. Multi-Instance / Multi-Context
_(populated in Phase 6 — see plan Task 6.3)_

### T4. Admin Experience
_(populated in Phase 6 — see plan Task 6.4)_

### T5. Observable & Diagnosable
_(populated in Phase 6 — see plan Task 6.5)_

---

## Part 3 — Foundations Track
_(populated in Phase 7 — see plan Task 7.3)_

---

## Part 4 — Roadmap Matrix
_(populated in Phase 8 — see plan Task 8.1)_

---

## Part 5 — Recommended Sprint Sequencing
_(populated in Phase 8 — see plan Task 8.2)_

---

## Part 6 — Appendices

### Appendix A — March 22 Audit Reconciliation
_(populated in Phase 1 — see plan Task 1.2)_

### Appendix B — spfx-toolkit Integration Map
_(populated in Phase 2 — see plan Task 2.2)_

### Appendix C — PnP Modern Search v4 Parity Scorecard
_(populated in Phase 3 — see plan Task 3.2)_

### Appendix D — Rejected Ideas
_(populated in Phase 8 — see plan Task 8.3)_

### Appendix E — Evidence and Command Log
_(populated in Phase 9 — see plan Task 9.3)_
```

- [ ] **Step 2: Verify scaffold**

Run:
```bash
grep -c '_(populated in' docs/sp-search-launch-readiness-audit.md
```

Expected: `15` (15 placeholder markers, one per section that gets populated later).

- [ ] **Step 3: Commit the scaffold**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): scaffold launch-readiness audit document

Scaffold-only commit with section headers and population markers.
Each marker references the plan task that fills it.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 0.4: Archive the March 22 audit

**Files:**
- Create: `docs/archive/sp-search-comprehensive-audit-2026-03-22.md`
- Delete: `docs/sp-search-comprehensive-audit.md`

- [ ] **Step 1: Verify .gitignore covers verification artifacts**

Run:
```bash
grep -E '^(lib|dist|temp|sharepoint/solution|node_modules|\.cache)/?' .gitignore || echo "MISSING"
```

Expected: lines for `lib/`, `dist/`, `temp/`, `node_modules/`, plus `.cache/` or `release/`. If any is MISSING, do NOT add it in this plan (out of scope — record in Appendix E "Pre-existing condition: .gitignore missing X" and proceed).

- [ ] **Step 2: Create archive directory if missing**

Run:
```bash
mkdir -p docs/archive
```

- [ ] **Step 3: Move file via git mv (preserves history)**

Run:
```bash
git mv docs/sp-search-comprehensive-audit.md docs/archive/sp-search-comprehensive-audit-2026-03-22.md
```

- [ ] **Step 4: Prepend redirect header**

Use Edit on `docs/archive/sp-search-comprehensive-audit-2026-03-22.md` to prepend:

```markdown
> **ARCHIVED — 2026-05-02.** This audit has been superseded by `docs/sp-search-launch-readiness-audit.md`. See Appendix A of the new audit for per-finding reconciliation status.

---

```

(Insert before the existing `# SP Search Comprehensive Audit Report` line. Do not delete or modify any other content.)

- [ ] **Step 5: Verify archive**

Run:
```bash
head -5 docs/archive/sp-search-comprehensive-audit-2026-03-22.md
ls docs/sp-search-comprehensive-audit.md 2>&1 || echo "correctly removed"
```

Expected: redirect header visible at top, original file no longer exists.

- [ ] **Step 6: Commit archival**

```bash
git add docs/archive/sp-search-comprehensive-audit-2026-03-22.md
git commit -m "$(cat <<'EOF'
docs(audit): archive 2026-03-22 comprehensive audit

Moves prior audit to docs/archive/ with a redirect header pointing at
the new launch-readiness audit. Per-finding reconciliation lives in
Appendix A of the new audit.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 1 — March 22 Audit Reconciliation (Appendix A)

Spec §5 step 2 requires this to come before any new findings.

### Task 1.1: Read and enumerate all March 22 findings

**Files:**
- Read: `docs/archive/sp-search-comprehensive-audit-2026-03-22.md` (full document)

- [ ] **Step 1: Read the entire archived audit**

Use Read tool on `docs/archive/sp-search-comprehensive-audit-2026-03-22.md` (no offset, full file). The document has ~53 findings spread across sections 2–10 with IDs like `BUG-001`, plus per-section issue lists.

- [ ] **Step 2: Enumerate findings into a working table (held in scratch notes)**

For each finding, capture:
- ID (use existing `BUG-NNN` ID if present; otherwise assign `LEGACY-NN` numbered in document order)
- Short title (one line)
- Original severity (Critical / High / Medium / Low)
- Original section (e.g., "2. Critical Bugs", "5. Security")
- File(s) referenced

Hold the enumeration in scratch notes; do not write to the audit file yet.

- [ ] **Step 3: Verify count**

Do not assume the prior audit's count is internally consistent. The executive summary says "12 critical/high issues, 23 medium issues, and 18 low-priority items" while the category table totals 7 critical + 10 high + 26 medium + 20 low. Enumerate the actual findings from headings/tables in document order, record the final enumerated count, and add a short "Prior audit count note" in Appendix A if it differs from either summary number.

### Task 1.2: Reconcile each finding against current code

**Files:**
- Read: source files referenced by each finding (varies)
- Read: `git log --all --oneline` for commits since `2026-03-22`

- [ ] **Step 1: Get the list of commits after the prior audit**

Run:
```bash
git log --since='2026-03-22' --oneline --all
```

Record commit count and commit subjects. Used to substantiate "Closed" claims.

- [ ] **Step 2: Reconcile each finding**

For each finding from Task 1.1, classify into exactly one of:

- **Closed** — the issue is fixed in current code. Cite commit SHA and current file:line.
- **Still-Open** — the issue is unchanged or only partially addressed. Cite current file:line proving it.
- **Obsolete** — the affected code no longer exists or the design changed such that the finding is moot. Explain why.
- **Changed-Form** — the underlying issue persists but with a different shape. Describe the new shape and cite current file:line.

For each finding, run targeted `grep`/`Read` to verify state. Do NOT reconcile from memory.

Hold the reconciliation in scratch notes structured as:
```
| ID | Title | Status | Evidence | New audit reference |
```

`New audit reference` is left as `TBD-trackX` for now and filled in during Task 8.1 once Roadmap IDs exist.

- [ ] **Step 3: Sanity check**

Confirm every finding has exactly one status. Confirm all Closed claims cite a commit SHA and all Still-Open / Changed-Form claims cite a current file:line.

### Task 1.3: Write Appendix A and Reconciliation Summary

**Files:**
- Modify: `docs/sp-search-launch-readiness-audit.md` (replace marker in Appendix A and Front Matter)

- [ ] **Step 1: Write Appendix A**

Replace the Appendix A placeholder with a single table containing every reconciled finding. Columns: `ID | Title | Original Severity | Status | Evidence | Audit Cross-Ref`. Group rows by status (Closed first, then Still-Open, then Changed-Form, then Obsolete) for readability. The `Audit Cross-Ref` column initially holds track-level pointers (e.g., "T1", "Foundations") — Task 8.1 will tighten these to specific Roadmap IDs.

- [ ] **Step 2: Write the Reconciliation Summary in Front Matter**

Replace the Front Matter "Reconciliation Summary" placeholder with a 3–5 row summary table: counts of Closed / Still-Open / Changed-Form / Obsolete, plus a one-paragraph narrative ("Of 53 findings, X were closed by … the largest open concentration is in …"). No vague language — every adjective backed by a number.

- [ ] **Step 3: Verify**

Run:
```bash
grep -A2 '## Appendix A' docs/sp-search-launch-readiness-audit.md | head -10
grep '_(populated in Phase 1' docs/sp-search-launch-readiness-audit.md && echo "FAIL: Phase 1 placeholder still present" || echo "OK"
```

Expected: Appendix A header followed by table; Phase 1 placeholders gone.

- [ ] **Step 4: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): reconcile March 22 audit findings (Appendix A)

Every prior finding classified Closed / Still-Open / Changed-Form / Obsolete
with code-level evidence. Cross-references to differentiator tracks are
track-level for now; Task 8.1 will tighten to Roadmap IDs.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 2 — spfx-toolkit Integration Map (Appendix B)

### Task 2.1: Inventory toolkit capabilities

**Files:**
- Read: `/Users/hemantmane/Development/spfx-toolkit/src/components/index.ts` and component subdirectories
- Read: recent commits via `git log --since='2026-01-01' --oneline`
- Read: relevant changelog/docs if present

- [ ] **Step 1: Enumerate components and recent additions**

Run:
```bash
cd /Users/hemantmane/Development/spfx-toolkit && ls src/components/ && git log --since='2026-01-01' --oneline | head -50
```

Capture component list and commit list. Use committed repo evidence (`CLAUDE.md` if present, package exports, component folders, and commit log) to identify NEW capabilities since SP Search last integrated: Comments, ManageAccess (improvements), browser storage utilities, HTML sanitization, FormContext fixes, CssLoader compat aliases, plus any others surfaced.

- [ ] **Step 2: Read each new capability's exported surface**

For each new capability, Read its `index.ts` and (if small) its main component file. Capture:
- Capability name
- One-line description
- Public exports relevant to consumers
- Any prerequisites (Graph permissions, providers, etc.)

### Task 2.2: Match capabilities to SP Search consumption

**Files:**
- Read: SP Search source as needed to confirm current consumption
- Modify: `docs/sp-search-launch-readiness-audit.md` (Appendix B)

- [ ] **Step 1: For each capability, decide Adopt / Consider / No Fit**

Build a working table:
```
| Capability | Status | Where it would land in SP Search | Effort | Differentiator | Notes |
```

Status taxonomy:
- **Adopt** — clear win, replaces existing custom code or unlocks a stated differentiator.
- **Consider** — plausible fit but needs design work; defer to per-track plan.
- **No Fit** — not applicable to SP Search v1.0.

`Where it would land` is a file path or feature reference (e.g., "Comments → searchManager/components/ResultAnnotations.tsx", "HTML sanitization → SearchResults snippet rendering"). `Differentiator` is one of T1–T5 or "Foundations".

- [ ] **Step 2: Write Appendix B**

Replace the Appendix B placeholder with a header paragraph (one sentence describing the appendix purpose), the working table, and a "Toolkit version inspected" line citing the version captured in Task 0.1 step 3.

- [ ] **Step 3: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): map spfx-toolkit capabilities to SP Search (Appendix B)

Adopt / Consider / No Fit classification with target file references and
differentiator alignment. Toolkit version inspected recorded for repeatability.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 3 — PnP Modern Search v4 Parity Scorecard (Appendix C)

### Task 3.1: Map PnP v4 feature surface

**Files:**
- Read: `docs/pnp-modern-search-alignment.md` (existing alignment notes)
- WebFetch: `https://microsoft-search.github.io/pnp-modern-search/` (PnP v4 docs)
- WebFetch: `https://github.com/microsoft-search/pnp-modern-search` (README)

- [ ] **Step 1: Read existing alignment notes**

Use Read on `docs/pnp-modern-search-alignment.md` to capture what mapping the project already documents. This is the starting point — confirm and extend, do not duplicate.

- [ ] **Step 2: Fetch current PnP v4 feature list**

Use WebFetch on the PnP docs site root. Extract the navigation tree / feature index. Capture access date and exact URL.

- [ ] **Step 3: Build the feature surface list**

Produce a list of PnP v4 features: Search Box features, Search Results features, Search Filters features, Search Verticals features, layout list, customizations / extensibility, etc. Use the PnP nav as the canonical structure to avoid omissions.

### Task 3.2: Grade each feature

**Files:**
- Read: SP Search source as needed to verify each grade
- Modify: `docs/sp-search-launch-readiness-audit.md` (Appendix C)

- [ ] **Step 1: Grade each feature**

For each PnP v4 feature, assign exactly one grade:
- **Better** — SP Search exceeds (cite the SP Search file/feature)
- **Parity** — equivalent (cite both)
- **Worse** — SP Search has it but inferior (cite gap)
- **Missing** — not in SP Search

Build a table:
```
| Area | PnP v4 Feature | Grade | SP Search Equivalent | Notes |
```

- [ ] **Step 2: Write Appendix C**

Replace the Appendix C placeholder with:
- One-paragraph header noting purpose ("informs positioning, not forced parity")
- The grading table
- A short "Positioning takeaways" subsection (3–5 bullets) extracting the highest-leverage gaps and strengths
- "Sources consulted" lines with URLs and access date

- [ ] **Step 3: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): grade PnP Modern Search v4 parity (Appendix C)

Per-feature grading (Better/Parity/Worse/Missing) with positioning
takeaways. Source URLs and access date recorded per spec evidence
standard.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 4 — Journey A: Day 1 Admin Install

### Task 4.1: Walk Journey A steps 1–6 (install → add web parts)

**Files:**
- Read: `scripts/Setup-SPSearchSite.ps1`, `scripts/Search-ScenarioPresets.ps1`, `scripts/Provision-SPSearchLists.ps1`, `scripts/Provision-SPSearchPage.ps1`, `scripts/Map-CrawledProperties.ps1`
- Read: `docs/deployment-guide.md`, `docs/provisioning-guide.md`, `docs/admin-guide.md`
- Read: `config/package-solution.json`
- Read: each web part `*.manifest.json`

- [ ] **Step 1: Walk steps 1–3 (acquire → upload → add app)**

For steps 1, 2, 3 (download `.sppkg`, upload to app catalog, add app to site), inspect:
- What `.sppkg` artifact name is produced (check `config/package-solution.json` `solution.name`)
- Documented install path in `docs/deployment-guide.md`
- API permissions or admin consent required (check `package-solution.json` `webApiPermissionRequests`)

For each, record one of: Blocker / Confusion / Polish / OK, with file:line evidence.

- [ ] **Step 2: Walk steps 4–5 (provisioning scripts)**

Read both provisioning scripts. For each:
- Required parameters
- Failure modes (auth, permissions, throttling, missing prereqs)
- Idempotency (safe to re-run?)
- Quality of error messages
- Whether the docs explain prerequisites (PnP.PowerShell version, app registration, etc.)

Log friction with severity. Confirm whether the scripts handle PnP.PowerShell version differences, deprecated cmdlet parameters, authentication mode changes, and idempotent reruns; cite script/docs evidence rather than uncommitted memory files.

- [ ] **Step 3: Walk step 6 (add web parts in edit mode)**

Read each web part's manifest (`src/webparts/*/SP*WebPart.manifest.json`). Note:
- Display name and description (do they make sense without docs?)
- Default property values (sane for unknown tenants?)
- Icon presence
- Whether the web parts surface helpful empty states in edit mode

- [ ] **Step 4: Hold findings for Task 4.3**

Carry friction logs in scratch notes. Do not write to audit file yet.

### Task 4.2: Walk Journey A steps 7–12 (configure → publish → handoff)

**Files:**
- Read: each web part's web part class (`src/webparts/*/SP*WebPart.ts`) — focus on `getPropertyPaneConfiguration` and default property values
- Read: `src/propertyPaneControls/*.ts`
- Read: `src/webparts/spSearchManager/components/AdminDashboard.tsx`

- [ ] **Step 1: Walk step 7 (configure searchContextId across web parts)**

Read property panes for the searchContextId field. Note:
- Is the field discoverable (top of property pane vs buried)?
- Does it explain itself (callout/help text)?
- Is there a default that makes the simple "single instance per page" case work without admin intervention?
- What happens if two web parts have different IDs by accident?

Log friction.

- [ ] **Step 2: Walk step 8 (configure scope/filters/columns/layout)**

For each of Box, Results, Filters, Verticals, Manager: skim `getPropertyPaneConfiguration` output. Capture:
- Number of property pane fields
- Use of grouping / collapsible sections
- Validation feedback (does the property pane warn on bad config?)
- Discoverability of the scenario presets entry point (T4 ties in here)

- [ ] **Step 3: Walk step 9 (run a test query)**

Inspect what happens when an admin types a query in edit mode:
- Does the search execute?
- Is there a "no results" state that makes sense?
- Does the Debug FAB surface relevant info?

- [ ] **Step 4: Walk steps 10–12 (saved searches, publish, handoff)**

Read `SearchManager` configuration paths. Inspect:
- Provisioned-list dependency story (`SearchSavedQueries`, `SearchHistory` lists from `Provision-SPSearchLists.ps1`)
- Permission/role inheritance and any item-level permission behavior documented in committed source/docs
- What an admin sees on the page after publish (admin dashboard?)

- [ ] **Step 5: Hold findings for Task 4.3**

### Task 4.3: Write Journey A section

**Files:**
- Modify: `docs/sp-search-launch-readiness-audit.md` (Part 1 → Journey A)

- [ ] **Step 1: Write Journey A narrative**

Replace the Journey A placeholder with one numbered subsection per step (12 subsections). Each subsection:
- Heading: `#### Step N — <step name>`
- One-paragraph narrative of what the admin does
- A "Friction:" callout listing each finding with severity tag (`[Blocker]`, `[Confusion]`, `[Polish]`)
- Each finding cites file:line and proposes the owning track (e.g., "→ T4.D3" or "→ Foundations")

If a step has no friction, write `Friction: none observed.`

- [ ] **Step 2: Verify**

Run:
```bash
grep -c '#### Step ' docs/sp-search-launch-readiness-audit.md
```

Expected: at least `12` (Journey A steps) — Journey B will add 12 more in Phase 5, so the post-Phase-5 expected count is 24. After Phase 4, expect exactly 12.

- [ ] **Step 3: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): write Journey A — Day 1 admin install (Part 1)

12-step walkthrough from .sppkg in hand to handoff. Friction logged
inline with severity and owning-track pointers per spec §4.2.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 5 — Journey B: Day 1 End-User Search

### Task 5.1: Walk Journey B steps 1–6 (land → query → filter)

**Files:**
- Read: `src/webparts/spSearchBox/components/*.tsx`
- Read: `src/webparts/spSearchVerticals/components/*.tsx`
- Read: `src/webparts/spSearchFilters/components/*.tsx`
- Read: `src/webparts/spSearchResults/components/*.tsx` (focus on empty state + first paint)
- Read: `src/libraries/spSearchStore/orchestrator/SearchOrchestrator.ts` (search trigger flow)

- [ ] **Step 1: Walk step 1 (lands on search page)**

Inspect the first paint experience. What does the user see before they type? Empty state? Marketing copy? Suggestions? Read the relevant React components and the orchestrator's initial state behaviour.

- [ ] **Step 2: Walk step 2 (types a query)**

Read the SearchBox component. Note:
- Debounce interval
- Visible affordances (clear button, submit button, voice/etc.)
- Behaviour on Enter vs blur vs auto-search
- Accessibility (labels, ARIA)

- [ ] **Step 3: Walk step 3 (sees suggestions or doesn't)**

Read the suggestion provider registry and SearchBox suggestion rendering. Note:
- Default suggestion providers active out-of-the-box
- Latency / loading state
- Empty suggestion handling
- Keyboard navigation through suggestions

- [ ] **Step 4: Walk step 4 (empty state if no results)**

Read SearchResults empty-state rendering. Note:
- Quality of the empty-state message
- Whether it offers next steps ("try removing filters", "search all verticals")
- Whether it ties into the "why no results" panel mentioned in T5

- [ ] **Step 5: Walk step 5 (switches verticals)**

Read SearchVerticals. Note:
- Visual affordance (tabs vs other)
- Badge count timing (when do counts update?)
- Behaviour when switching mid-query
- Per-vertical `dataProviderId` routing visible to user (people vs sharepoint)

- [ ] **Step 6: Walk step 6 (applies filters)**

Read SearchFilters and a representative filter type (CheckboxFilter, DateRangeFilter). Note:
- Discoverability of available filters
- Behaviour with no refiners returned
- Stability under rapid re-querying (refiner stability mode)
- Active filter pill bar (add/remove from results page)

- [ ] **Step 7: Hold findings for Task 5.3**

### Task 5.2: Walk Journey B steps 7–12 (layout → detail → save → share → return → mobile)

**Files:**
- Read: `src/webparts/spSearchResults/components/*.tsx` (layouts, detail panel, bulk actions)
- Read: `src/webparts/spSearchManager/components/*.tsx` (saved searches, share)
- Read: `src/libraries/spSearchStore/store/middleware/urlSyncMiddleware.ts` and `src/libraries/spSearchStore/utils/filterUrlAliases.ts` (deep link + URL alias handling)
- Read: relevant SCSS modules for mobile/responsive behavior

- [ ] **Step 1: Walk step 7 (switches layouts)**

Read the layout switcher and one or two layouts (CardLayout, ListLayout). Note:
- Discoverability of switcher
- Persistence of layout choice (per user? per page?)
- Loading shimmer parity across layouts
- Lazy loading behaviour (does the user see a flash?)

- [ ] **Step 2: Walk step 8 (opens detail panel)**

Read the detail panel and one or two cell renderers. Note:
- Open animation / focus management
- Metadata completeness
- Version history behaviour
- Close behaviour (Esc, click-outside, focus return)

- [ ] **Step 3: Walk step 9 (saves a search)**

Read SavedSearchList + SearchManagerService save path. Note:
- Affordance to save (where is the button?)
- What gets saved (query + filters + layout + sort?)
- Naming UX (auto-name vs prompt)
- Confirmation feedback

- [ ] **Step 4: Walk step 10 (shares with a colleague)**

Read ShareSearchDialog + permission paths. Note:
- People picker UX
- Permission level granted (read vs edit)
- Shared search visibility for the recipient
- Notification (does the recipient know? — likely no, that's a finding)

- [ ] **Step 5: Walk step 11 (returns via deep link)**

Read urlSyncMiddleware and filter URL alias handling. Note:
- Which state survives the URL round-trip
- BUG-003 status (pending URL filters timeout) — confirm fixed or still open via Phase 1 reconciliation
- Multi-context URL namespacing under real navigation
- Restoration timing (does the page render then jump?)

- [ ] **Step 6: Walk step 12 (mobile)**

Read mobile-relevant SCSS + responsive logic. Note:
- Layout adaptation per viewport (derive breakpoints from SCSS/source, not memory notes)
- Touch target sizing
- Filter access on mobile (panel? collapsed?)
- DataGrid behaviour on small viewports (iOS momentum scroll)

- [ ] **Step 7: Hold findings for Task 5.3**

### Task 5.3: Write Journey B section

**Files:**
- Modify: `docs/sp-search-launch-readiness-audit.md` (Part 1 → Journey B)

- [ ] **Step 1: Write Journey B narrative**

Replace the Journey B placeholder with 12 numbered subsections, same shape as Journey A: heading, narrative paragraph, Friction callout with severity tags and track pointers.

- [ ] **Step 2: Verify**

Run:
```bash
grep -c '#### Step ' docs/sp-search-launch-readiness-audit.md
```

Expected: `24` (12 from Journey A + 12 from Journey B).

- [ ] **Step 3: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): write Journey B — Day 1 end-user search (Part 1)

12-step walkthrough from landing on the page through mobile flow.
Friction logged inline with severity and owning-track pointers per
spec §4.2.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 6 — Differentiator Tracks (T1–T5)

Each track follows the spec §4.3 sub-structure: Current State / Gap to "Amazing" / Deliverables / Out of Scope for v1.0. Each deliverable carries: short description, why it matters (tied to differentiator), effort tier, priority tier, dependencies, source friction (journey step or audit finding), and acceptance signal (per spec §4.5 + §5.1).

**Cap per track: target 10 deliverables, max 15** (per spec §8 risk mitigation).

### Task 6.1: T1 Modern UI Quality

**Files:**
- Read: `src/webparts/spSearchResults/components/*Layout.tsx`
- Read: `src/webparts/spSearchResults/components/*.tsx`
- Read: `src/styles/`, all `*.module.scss` in webparts
- Read: any theme tokens / Fluent theme integration
- Modify: `docs/sp-search-launch-readiness-audit.md` (Part 2 → T1)

- [ ] **Step 1: Read the relevant subsystem**

Walk: every layout, the empty state component(s), the loading shimmer/skeleton implementations, the detail panel transitions, mobile breakpoints, theme integration, error states. Capture file:line evidence per finding.

- [ ] **Step 2: Write T1 section**

Replace the T1 placeholder with the four-part structure:
- **Current state** (paragraph + bullet list, code-grounded)
- **Gap to "amazing"** (paragraph framed against the audience profile: "any tenant, self-serve")
- **Deliverables** (numbered list, each: description, why-it-matters, effort, priority, depends-on, source, acceptance signal)
- **Out of scope for v1.0** (bullet list, one-line rationale per item)

- [ ] **Step 3: Verify deliverable count and shape**

Each deliverable must include all seven fields (description, why, effort, priority, depends-on, source, acceptance signal). Count check: 5–15 deliverables.

- [ ] **Step 4: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): T1 Modern UI Quality track

Current state, gap analysis, and sized/prioritized deliverables for
the visual quality bar. Each deliverable carries effort, priority,
dependencies, source friction, and acceptance signal per spec §4.5.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 6.2: T2 End-User Productivity

**Files:**
- Read: `src/webparts/spSearchManager/components/*.tsx` (saved, collections, history, share, annotations)
- Read: `src/libraries/spSearchStore/services/SearchManagerService.ts`
- Read: `src/libraries/spSearchStore/providers/*SuggestionProvider.ts`, `RecentSearchProvider.ts`, and `TrendingQueryProvider.ts`
- Read: keyboard handlers in SearchBox + Results
- Read: bulk actions toolbar + export paths
- Read: spfx-toolkit Comments component (`/Users/hemantmane/Development/spfx-toolkit/src/components/Comments/`)
- Modify: `docs/sp-search-launch-readiness-audit.md` (Part 2 → T2)

- [ ] **Step 1: Read the relevant subsystem**

Walk: saved searches (CRUD + persistence + sharing), collections/pinboards, history, annotations, query templates, recent/trending suggestions, keyboard shortcuts, multi-select bulk actions, CSV export, the new Comments component (would land where?). Cite file:line per finding.

- [ ] **Step 2: Write T2 section**

Same four-part structure as T1.

- [ ] **Step 3: Verify shape and count**

Same checks as T1.

- [ ] **Step 4: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): T2 End-User Productivity track

Sized/prioritized deliverables for saved searches, collections, history,
sharing, suggestions, bulk actions, export, and Comments integration.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 6.3: T3 Multi-Instance / Multi-Context

**Files:**
- Read: `src/libraries/spSearchStore/store/storeRegistry.ts` (window-backed singleton)
- Read: `src/libraries/spSearchStore/store/middleware/urlSyncMiddleware.ts`
- Read: `src/libraries/spSearchStore/utils/filterUrlAliases.ts`
- Read: each web part class for `searchContextId` property handling
- Read: per-vertical `dataProviderId` routing in vertical slice + orchestrator
- Modify: `docs/sp-search-launch-readiness-audit.md` (Part 2 → T3)

- [ ] **Step 1: Read the relevant subsystem with collision/isolation lens**

Walk: how `searchContextId` flows from property → store registry → URL namespacing. Trace every code path that reads or writes `__sp_search_context_map__`.

- [ ] **Step 2: Verify required scenarios are addressed**

Per spec §4.3 (T3), this track MUST include findings/deliverables covering:
- Two independent search experiences on the same page with different context IDs
- Two web parts accidentally sharing a context ID when isolation was expected
- URL deep-link parameters for two contexts on the same page
- Navigation away/back and store cleanup expectations

If any scenario lacks an explicit finding or deliverable, add one (or document why it is fully covered today and no work is needed — with proof).

- [ ] **Step 3: Write T3 section**

Same four-part structure as T1.

- [ ] **Step 4: Verify shape, count, and scenario coverage**

In addition to the standard checks, confirm all four required scenarios from Step 2 are addressed in either Current State or Deliverables.

- [ ] **Step 5: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): T3 Multi-Instance / Multi-Context track

Sized/prioritized deliverables for namespace correctness, URL collision,
per-vertical provider routing, and singleton-backing patterns. Includes
explicit collision and isolation scenarios per spec §4.3.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 6.4: T4 Admin Experience

**Files:**
- Read: each web part `getPropertyPaneConfiguration` method
- Read: `src/propertyPaneControls/*.ts`
- Read: `src/webparts/spSearchResults/presets/searchPresets.ts`
- Read: `src/webparts/spSearchManager/components/AdminDashboard.tsx`, `CoverageStatsSection.tsx`, `QualityMetricsSection.tsx`, `ZeroResultsPanel.tsx`, `SearchInsightsPanel.tsx`
- Read: scenario provisioning script `scripts/Search-ScenarioPresets.ps1`
- Modify: `docs/sp-search-launch-readiness-audit.md` (Part 2 → T4)

- [ ] **Step 1: Read the relevant subsystem**

Walk: property pane discoverability, scenario preset coverage (general/documents/news/people/media/custom shipped; KB/Hub/Policy pending), schema picker UX, edit-mode validation, provisioning script robustness, admin dashboard depth across all four panels (Coverage, Quality, Health/Zero-Results, Insights), default value sanity for unknown tenants.

- [ ] **Step 2: Write T4 section**

Same four-part structure as T1.

- [ ] **Step 3: Verify shape and count**

- [ ] **Step 4: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): T4 Admin Experience track

Sized/prioritized deliverables for property pane UX, scenario presets,
schema picker, edit-mode validation, provisioning robustness, and admin
dashboard depth.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 6.5: T5 Observable & Diagnosable

**Files:**
- Read: `src/libraries/spSearchStore/debug/*.ts` (DebugCollector, IDebugTypes)
- Read: `src/webparts/spSearchResults/components/DebugFab.tsx`, `DebugPanel.tsx` (or wherever they live)
- Read: error surfacing patterns across web parts (toasts, panels, inline)
- Read: `src/libraries/spSearchStore/services/CoverageStatsService.ts`
- Modify: `docs/sp-search-launch-readiness-audit.md` (Part 2 → T5)

- [ ] **Step 1: Read the relevant subsystem with the local-vs-telemetry lens**

Walk: Debug FAB tabs (Query / Network / State / Logs / Errors per spec) — feature complete or partial? Error surfacing patterns. "Why no results" panel — exists, missing, or stub? Coverage stats. Support bundle export — does any path produce one today?

- [ ] **Step 2: Verify required local-vs-telemetry split is addressed**

Per spec §4.3 (T5), this track MUST distinguish:
- **Local diagnostics** — admin/user-visible debug state, support bundle, recent network/search calls, config snapshot. No external send.
- **Telemetry** — optional aggregate signals. Opt-in only, never captures query text, user identity, result titles, URLs, tenant names, or list item content.

If either category lacks coverage, add findings/deliverables.

- [ ] **Step 3: Write T5 section**

Same four-part structure as T1, with the local-vs-telemetry distinction visible in headings or callouts.

- [ ] **Step 4: Verify shape, count, and split coverage**

- [ ] **Step 5: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): T5 Observable & Diagnosable track

Sized/prioritized deliverables for Debug FAB completeness, error
surfacing, why-no-results panel, support bundle export, and telemetry
strategy. Local diagnostics and telemetry kept explicitly distinct.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 7 — Foundations Track (Part 3)

### Task 7.1: Read foundations subject areas

**Files:**
- Read: SP Search source for `window.location.href`, `dangerouslySetInnerHTML`, `eval`, `new Function`, untrusted-string sinks (security)
- Read: `gulpfile.js`, `tsconfig.json`, `config/config.json`, `package.json` scripts (build/CI)
- Read: any existing accessibility code (focus management, ARIA roles, motion-reduction respect)
- Read: README.md (if present), `docs/admin-guide.md`, `docs/deployment-guide.md`, `docs/provisioning-guide.md`, `docs/extensibility-guide.md`
- Read: bundle size evidence (`npm run stats` or `npm run stats:json` output if available, otherwise document as gap)
- Read: SPFx 1.22 / Heft migration evidence (current branch, commits since migration started)

- [ ] **Step 1: Security sweep**

Run:
```bash
grep -rn "window\.location\.href" src/
grep -rn "dangerouslySetInnerHTML" src/
grep -rn "innerHTML" src/
grep -rn "eval(" src/
```

For each match, classify: safe / risky / unverified. BUG-004 (`newPageUrl` XSS) per the prior audit must appear in this sweep — confirm whether reconciled as Closed or Still-Open in Phase 1.

- [ ] **Step 2: 1.22 / Heft migration completion check**

Run:
```bash
git log --oneline feat/spfx-1.22-heft-migration ^main | head -30
git diff main...feat/spfx-1.22-heft-migration --stat | tail -20
```

Capture: scope of unmerged changes and any obvious gaps from committed evidence or command output (test failures, lint warnings, package output, etc.).

- [ ] **Step 3: Accessibility baseline scan**

Search for ARIA usage, focus management patterns, motion-reduction respect (`prefers-reduced-motion`), keyboard handlers. Cite file:line. Identify the gaps; do not attempt a full WCAG audit (that's out of scope for a doc audit — flag as a deliverable).

- [ ] **Step 4: CI / release engineering check**

Look for: a CI workflow file, a versioning policy (CHANGELOG, semver tags), `.sppkg` build automation, smoke tests, release checklist. Cite presence or absence.

- [ ] **Step 5: Documentation check**

Inventory existing docs in `docs/`. Note what exists, what's missing for "any SPFx tenant can install" (top-level README, per-web-part config reference, scenario gallery, troubleshooting/FAQ, contributing guide if open-sourced).

- [ ] **Step 6: Telemetry plumbing check**

Search for any existing telemetry hook, analytics call, fetch to external endpoint. There likely isn't one — confirm and capture as a deliverable.

- [ ] **Step 7: Performance budgets check**

Look for documented or enforced bundle size budgets. Capture current bundle sizes if available from build output (or flag as a deliverable to document them).

### Task 7.2: Hold findings, decide deliverables

- [ ] **Step 1: Aggregate findings into Foundations deliverables**

For each subject area in Task 7.1, decide deliverables (same shape as differentiator deliverables: description, why, effort, priority, depends-on, source, acceptance signal). Cap target ~10 deliverables, max 15.

- [ ] **Step 2: Apply the P0 admission rule**

For each Foundations deliverable, if marked P0, confirm it ties to (a)–(e) per spec §6. Security and "would prevent install" issues qualify directly.

### Task 7.3: Write Foundations Track

**Files:**
- Modify: `docs/sp-search-launch-readiness-audit.md` (Part 3)

- [ ] **Step 1: Write Foundations section**

Replace the Part 3 placeholder. Structure: brief intro paragraph, then one subsection per subject area (Security, 1.22 migration, Accessibility, CI/release, Documentation, Telemetry, Performance budgets). Each subsection: short current-state paragraph + numbered deliverables. Each deliverable must include all seven fields from the spec evidence standard: description, why-it-matters, effort, priority, depends-on, source, and acceptance signal.

- [ ] **Step 2: Verify**

Run:
```bash
grep -c '_(populated in Phase 7' docs/sp-search-launch-readiness-audit.md
```

Expected: `0` (Phase 7 placeholder removed).

- [ ] **Step 3: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): Foundations track (Part 3)

Cross-cutting deliverables: security, SPFx 1.22 migration completion,
accessibility baseline, CI/release engineering, docs, telemetry plumbing,
performance budgets.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 8 — Roadmap Matrix, Sprint Sequencing, Rejected Ideas

### Task 8.1: Compile the Roadmap Matrix

**Files:**
- Modify: `docs/sp-search-launch-readiness-audit.md` (Part 4 + Appendix A back-references)

- [ ] **Step 1: Assign Roadmap IDs to every deliverable**

Walk every deliverable from T1, T2, T3, T4, T5, and Foundations. Assign IDs in the form `<track>.D<n>` (e.g., `T1.D1`, `T4.D7`, `Found.D3`). Also assign cross-cutting deliverables a single primary track (the one most responsible) plus secondary tracks via the matrix `Track(s)` column.

- [ ] **Step 2: Write the Roadmap Matrix**

Replace the Part 4 placeholder with a single sortable table:

```
| ID | Deliverable | Track(s) | Effort | Priority | Depends on | Source | Acceptance Signal |
```

Order: P0 first (within P0 by track in T1→T5→Foundations order), then P1, then P2, then Defer. Include every deliverable defined in Phases 6–7. Every cell must have content — no `TBD`.

- [ ] **Step 3: Update Appendix A cross-references**

Walk Appendix A (written in Phase 1). Replace each `TBD-trackX` cross-reference with the actual Roadmap ID(s) where the finding is now addressed. If a Still-Open finding has no matching deliverable, that's a plan failure — go back to the relevant track and add a deliverable, then re-do this step.

- [ ] **Step 4: Verify**

Run:
```bash
grep -c '| T[1-5]\.D\|| Found\.D' docs/sp-search-launch-readiness-audit.md
grep -c 'TBD-track' docs/sp-search-launch-readiness-audit.md
```

Expected: first count > 0 (every Roadmap row + every Appendix A cross-ref). Second count = 0 (all placeholder cross-refs replaced).

- [ ] **Step 5: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): Roadmap Matrix (Part 4) + Appendix A cross-refs

Single sortable table with every deliverable assigned an ID, effort,
priority, dependencies, source, and acceptance signal. Appendix A
cross-references updated to point at concrete Roadmap IDs.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 8.2: Write Sprint Sequencing (Part 5)

**Files:**
- Modify: `docs/sp-search-launch-readiness-audit.md` (Part 5)

- [ ] **Step 1: Slot deliverables into sprints**

Three sprints (per spec §4.6):
- **Sprint 4 — Foundations + Critical UX**: security, 1.22 merge, top journey blockers, accessibility quick wins
- **Sprint 5 — Differentiator Depth**: bulk of T1–T5 P0/P1 deliverables
- **Sprint 6 — Polish + Docs**: remaining P1, docs site, sample gallery, release engineering

Each sprint lists 8–12 deliverables drawn from the matrix BY ID ONLY (do not invent new work — spec §5.3 self-review rule).

- [ ] **Step 2: Write Part 5**

Replace the Part 5 placeholder with three subsections (one per sprint), each containing:
- One-sentence sprint theme
- Deliverable list as a table: `ID | Deliverable | Effort | Priority | Rationale-for-this-sprint`
- Rough total effort for the sprint (sum of S/M/L/XL — use S=0.5d, M=1d, L=2d, XL=4d as a heuristic stated in the section preamble)

After the three sprints, list `P2 / Defer` items in a single bullet list of IDs (no bodies — they live in the Roadmap Matrix).

- [ ] **Step 3: Verify**

Run:
```bash
grep -c '#### Sprint ' docs/sp-search-launch-readiness-audit.md
```

Expected: `3`.

- [ ] **Step 4: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): Recommended Sprint Sequencing (Part 5)

Three 2-week sprints with deliverables referenced by Roadmap ID only.
P2/Defer items listed but not slotted.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 8.3: Write Rejected Ideas (Appendix D)

**Files:**
- Modify: `docs/sp-search-launch-readiness-audit.md` (Appendix D)

- [ ] **Step 1: List rejected ideas**

Capture ideas considered during the audit (from journeys, tracks, foundations) and consciously dropped. Each gets a one-line rationale. Examples likely to appear: "third-party search integrations (e.g., M365 Search APIs not currently used) — out of scope for any-tenant install", "redesign of Zustand store — not required for launch", "Excel export beyond CSV — defer unless the current source proves it is already supported and launch-critical".

If no rejected ideas surfaced (unlikely), write a single line: "No ideas were formally rejected during this audit." But this is a smell — most audits surface at least 3–5 deferrals.

- [ ] **Step 2: Write Appendix D**

Replace the placeholder with a bulleted list, one line per idea: `- **<idea>** — <one-line rationale>`.

- [ ] **Step 3: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): Rejected Ideas (Appendix D)

Ideas considered and consciously dropped during the audit, with
one-line rationale each.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Phase 9 — Self-Review and Finalization

### Task 9.1: Run the spec self-review checklist

Spec §5.3 defines the checklist. Run each item against the audit document.

- [ ] **Step 1: Every P0 must cite the P0 admission rule category**

Run:
```bash
grep -nE '\| P0 \|' docs/sp-search-launch-readiness-audit.md
```

For each P0 row, confirm the deliverable text or a reference column names which P0 admission rule category (a/b/c/d/e) it satisfies. If any P0 lacks justification, either add it or downgrade the priority.

- [ ] **Step 2: Every roadmap row must have a source and acceptance signal**

Run:
```bash
grep -nE '\| T[1-5]\.D|\| Found\.D' docs/sp-search-launch-readiness-audit.md | grep -E '\| {0,2}\|' && echo "FAIL: empty cells found" || echo "OK"
```

Expected: `OK`. If FAIL, fix the empty cells.

- [ ] **Step 3: Every journey Blocker maps to at least one P0/P1 row**

Walk Journey A and Journey B sections. For each `[Blocker]` tag, confirm the referenced deliverable ID resolves in the Roadmap Matrix and is P0 or P1 (or has an explicit documented workaround spelled out inline). If any Blocker maps only to a P2/Defer row with no workaround, escalate the priority.

- [ ] **Step 4: Every March 22 finding reconciled exactly once**

Run:
```bash
grep -cE '^\| (BUG|LEGACY)-' docs/sp-search-launch-readiness-audit.md
```

Expected: matches the count from Task 1.1 step 3 (~53). If counts don't match, fix Appendix A.

- [ ] **Step 5: Every toolkit capability marked Adopt / Consider / No Fit**

Run:
```bash
grep -cE '\| (Adopt|Consider|No Fit) \|' docs/sp-search-launch-readiness-audit.md
```

Expected: matches the Appendix B row count from Task 2.2 step 1.

- [ ] **Step 6: Sprint sequencing contains only Roadmap IDs**

Walk Part 5. Confirm every entry in every sprint is a Roadmap ID (no descriptive-only bullets). If any entry lacks an ID, replace it with the matching Roadmap ID or remove it.

- [ ] **Step 7: No placeholder markers remain**

Run:
```bash
grep -nE '_\(populated in Phase|TBD|TODO|FIXME' docs/sp-search-launch-readiness-audit.md
```

Expected: zero matches. If any, fix them.

### Task 9.2: Verify all spec acceptance criteria

Spec §9 lists 7 acceptance criteria. Verify each.

- [ ] **Step 1: Verify file exists and has all sections**

Run:
```bash
test -f docs/sp-search-launch-readiness-audit.md && echo "OK: file exists"
grep -cE '^## (Front Matter|Part [1-6]|Part 1|Part 2|Part 3|Part 4|Part 5|Part 6)' docs/sp-search-launch-readiness-audit.md
grep -cE '^### Appendix [A-E]' docs/sp-search-launch-readiness-audit.md
```

Expected: file exists; Parts heading count ≥ 6; Appendix count = 5.

- [ ] **Step 2: Verify every March 22 finding appears in Appendix A**

Already covered by Task 9.1 step 4.

- [ ] **Step 3: Verify Roadmap Matrix completeness**

Already covered by Task 9.1 step 2.

- [ ] **Step 4: Verify P0 justification rule**

Already covered by Task 9.1 step 1.

- [ ] **Step 5: Verify journey Blocker mapping**

Already covered by Task 9.1 step 3.

- [ ] **Step 6: Verify Appendix E exists with required content (after Task 9.3 writes it)**

Deferred until after Task 9.3 — this step revisits after Appendix E lands.

- [ ] **Step 7: Verify old audit moved with redirect header**

Run:
```bash
test -f docs/archive/sp-search-comprehensive-audit-2026-03-22.md && echo "OK: archived"
test -f docs/sp-search-comprehensive-audit.md && echo "FAIL: original still exists" || echo "OK: original removed"
head -3 docs/archive/sp-search-comprehensive-audit-2026-03-22.md | grep -q "ARCHIVED" && echo "OK: redirect header present" || echo "FAIL: redirect header missing"
```

Expected: archived file exists; original removed; redirect header present.

### Task 9.3: Write Appendix E (Evidence and Command Log)

**Files:**
- Modify: `docs/sp-search-launch-readiness-audit.md` (Appendix E)

- [ ] **Step 1: Write Appendix E**

Replace the Appendix E placeholder with subsections:
1. **Repo snapshot** — branch, commit SHA, date/time, captured during Task 0.1
2. **Package versions** — sp-search version, SPFx package versions, key dependency versions, spfx-toolkit version (from Task 0.1 steps 2–3)
3. **Verification commands** — for each of the 5 commands in spec §5.2, list: command, exit code, brief result, link to truncated log if needed (from Task 0.2)
4. **Generated artifacts during verification** — files modified by verification commands and explicitly NOT committed with the audit (from Task 0.2 step 4)
5. **External sources consulted** — each PnP v4 doc URL with access date (from Task 3.1)
6. **Skipped checks** — any check skipped, with reason (e.g., "npm test skipped — Jest harness blocker, see Found.D-X")

- [ ] **Step 2: Re-run Task 9.2 step 6**

Run:
```bash
grep -A2 '## Appendix E' docs/sp-search-launch-readiness-audit.md | head -10
```

Confirm Appendix E has content.

- [ ] **Step 3: Commit**

```bash
git add docs/sp-search-launch-readiness-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): Appendix E — evidence and command log

Repo snapshot, package versions, verification command results, generated
artifacts, external sources with access dates, and skipped checks. Closes
the spec evidence standard (§5.1).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 9.4: Final sanity sweep and present

- [ ] **Step 1: Final placeholder sweep**

Run:
```bash
grep -nE '_\(populated|TBD|TODO|FIXME|XXX' docs/sp-search-launch-readiness-audit.md
```

Expected: zero matches. If any appear, fix and amend the relevant prior commit (use `git commit --amend` only on commits NOT yet pushed; otherwise add a `docs(audit): fix stray placeholder` commit).

- [ ] **Step 2: Word count and section count summary**

Run:
```bash
wc -w docs/sp-search-launch-readiness-audit.md
grep -cE '^### ' docs/sp-search-launch-readiness-audit.md
grep -cE '^#### ' docs/sp-search-launch-readiness-audit.md
```

Capture for the user-facing summary.

- [ ] **Step 3: Confirm working tree is clean**

Run:
```bash
git status --short
```

Expected: clean (no uncommitted changes). If any appear and they belong to the audit, commit them. If they belong to verification artifacts, confirm `.gitignore` covers them; if not, do NOT commit them and note in Appendix E.

- [ ] **Step 4: Surface results to the user**

Output to user:
- Path to the audit document
- Total deliverable count broken down by track (T1: N · T2: N · T3: N · T4: N · T5: N · Foundations: N)
- Total P0 / P1 / P2 / Defer counts
- Top 3 P0 launch blockers (one line each)
- Suggested next step: "Audit complete. Per spec §7, the next step is per-track implementation plans via the writing-plans skill — Foundations first, then T1–T5 in any order. Want to start on Foundations?"

---

## Self-Review (writing-plans skill output)

After writing this plan, fresh-eyes pass against the spec:

**Spec coverage check.** Walked spec sections 1 through 10:
- §1 Context — covered by plan header.
- §2 Goals — Goals 1–6 each map to specific tasks: G1 → Task 0.3 (scaffold) + all later writes; G2 → Tasks 1.1–1.3; G3 → Tasks 6.1–6.5 + 7.3; G4 → Tasks 4.1–5.3; G5 → Task 8.1 (matrix has all required columns including acceptance signal); G6 → handoff offered after Task 9.4.
- §3 + §3.1 + §3.2 — Non-goals respected (no runtime source changes); Launch-Ready Bar applied via P0 admission rule (§6) referenced in Task 9.1 step 1; Audit Inputs captured in Tasks 0.1, 0.2, 3.1, 9.3.
- §4 Document structure — Front Matter (Task 0.3 + 1.3 + 9.3); Part 1 journeys (Tasks 4–5); Part 2 tracks (Tasks 6.1–6.5); Part 3 Foundations (Task 7.3); Part 4 Roadmap Matrix (Task 8.1); Part 5 Sprint Sequencing (Task 8.2); Part 6 Appendices A–E (Tasks 1.3, 2.2, 3.2, 8.3, 9.3).
- §5 Methodology — 8 passes: code-grounded enforced via "cite file:line" instructions; reconciliation first (Phase 1); toolkit comparison (Phase 2); journey simulation (Phases 4–5); PnP v4 (Phase 3); track passes (Phase 6); foundations sweep (Phase 7); no fixes (called out in Architecture + every commit subject is `docs(audit):`).
- §5.1 Evidence Standard — referenced in track tasks; "no weasel words" enforced by Task 9.1.
- §5.2 Verification commands — Task 0.2 runs all five.
- §5.3 Self-review pass — Task 9.1 walks all 6 self-review items explicitly.
- §6 Prioritization Framework — Effort and Priority tiers used consistently in track tasks; P0 admission rule enforced in Task 9.1 step 1; admission rule expanded form (a–e) reflected in Task 9.1.
- §7 Outputs and Follow-up — audit produced; per-track plans flagged as next step in Task 9.4 step 4.
- §8 Risks — cap on deliverable count per track (target 10, max 15) called out in Phase 6 preamble; P0 inflation guard in Task 9.1; verification artifacts kept out of audit commits per Task 0.2 step 4 + Task 9.4 step 3.
- §9 Acceptance Criteria — Task 9.2 walks all 7.
- §10 Out of Scope — respected (no runtime code changes anywhere; spec confirms doc moves are allowed).

**Placeholder scan.** Searched for the red flags from the skill: "TBD" appears only as the placeholder text the plan instructs to fix (Tasks 1.2, 8.1) and the explicit grep checks in Tasks 9.1/9.4 that fail on it. No "implement later" / "fill in details" / "similar to Task N" / "appropriate error handling" patterns. The `_(populated in Phase X — see plan Task X.Y)_` markers in Task 0.3 are intentional inline references that explicitly point at the populating task; they are removed by Tasks 1.3, 2.2, 3.2, 4.3, 5.3, 6.1–6.5, 7.3, 8.1, 8.2, 8.3, 9.3, and a final sweep in 9.4 step 1 verifies zero remain.

**Type / consistency.** Track IDs used consistently (T1–T5 + Foundations); deliverable IDs use `T<n>.D<n>` and `Found.D<n>`; Roadmap Matrix columns match across Tasks 8.1, 9.1, and 9.2; severity tags (`Blocker`/`Confusion`/`Polish`) match spec §4.2; verification commands match spec §5.2.

No issues found. Plan is internally consistent with the spec.

---

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-05-02-launch-readiness-audit-production.md`.

Two execution options for producing the audit:

**1. Subagent-Driven (recommended)** — I dispatch a fresh subagent per task with explicit context, review the output between tasks, faster iteration on quality. Best for an audit because each phase produces a discrete artifact that benefits from a fresh perspective.

**2. Inline Execution** — I execute tasks in this session using executing-plans, batched with checkpoints for your review at phase boundaries. Lower coordination overhead but heavier on context.

Which approach?
