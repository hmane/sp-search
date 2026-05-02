---
name: SP Search Launch-Readiness Audit
description: Spec for the comprehensive pre-launch audit of SP Search, organized as journeys + differentiator tracks + foundations, producing a single audit document and per-track follow-up implementation plans.
status: approved
date: 2026-05-02
owner: Hemant Mane
---

# SP Search Launch-Readiness Audit — Specification

## 1. Context

SP Search is a SharePoint search solution built as 6 SPFx 1.22 web parts plus 1 library component. It is intended to be a modern alternative to PnP Modern Search v4. The solution is **pre-launch** with a target audience of **"any SPFx tenant can install it"** — generic, self-serve, docs-driven, no hand-holding. There is no fixed launch date; the gating criterion is quality.

Two prior audits and a flurry of recent feature work have left the project with overlapping context:

- `docs/sp-search-comprehensive-audit.md` (2026-03-22, 775 lines, 53 findings)
- `docs/sp-search-requirements.md` (2026-02-06, 1,868 lines, v1.4)
- Since 2026-03-22: SPFx 1.22 / Heft migration, Admin Manager web part, Admin Dashboard, Debug FAB + DebugPanel, CoverageStatsService, lazy-loading sweep, lint cleanup
- spfx-toolkit (the in-house dependency) has shipped material updates: Comments component, ManageAccess, browser storage utilities, HTML sanitization, FormContext fixes, CssLoader compat aliases — none yet integrated

The current state is therefore not faithfully described by any single existing artifact, and a fresh, authoritative pre-launch audit is required.

## 2. Goals

This audit will:

1. Produce a single authoritative document (`docs/sp-search-launch-readiness-audit.md`) that supersedes `docs/sp-search-comprehensive-audit.md` and reflects the **actual** current state of the codebase.
2. Reconcile every finding from the 2026-03-22 audit (Closed / Still-Open / Obsolete / Changed-Form) with code-level evidence.
3. Frame all findings against five user-selected differentiators: **Modern UI Quality, End-User Productivity, Multi-Instance / Multi-Context, Admin Experience, Observable & Diagnosable**.
4. Surface the friction an unaccompanied admin or end-user would hit on Day 1, via two end-to-end journey walkthroughs.
5. Produce an executable roadmap matrix: every deliverable tagged with effort (S/M/L/XL), priority (P0/P1/P2/Defer), differentiator(s), dependencies, source, and acceptance signal.
6. Lay the foundation for **per-track implementation plans** (one Foundations + one per differentiator) to follow via the writing-plans skill.

## 3. Non-Goals

- **No code changes during the audit.** Pure analysis. Implementation lives in the per-track plans that follow.
- **Not a re-spec of requirements.** The requirements doc remains authoritative for *what* SP Search is supposed to do. The audit is about gaps between current state and launch-ready state.
- **No exhaustive line-by-line review.** Targeted reads guided by the differentiator and journey lenses. We will not catalogue stylistic nits or dead code unless it affects launch readiness.
- **Not a unilateral PnP v4 feature copy.** PnP v4 parity scorecard informs *positioning*, not a forced parity backlog. Missing PnP features are deliverables only if they tie to a stated differentiator.
- **Not a live tenant certification.** The audit can inspect scripts, manifests, package output, and documentation, but it does not prove behavior in every tenant topology without a later smoke run in a real SharePoint tenant.

## 3.1 Launch-Ready Bar

For this audit, "launch-ready" means an unknown SPFx-capable tenant can install, configure, diagnose, and operate SP Search from the published package and docs without author intervention. A launch blocker is any issue that would likely cause one of these outcomes:

- The package cannot be built, packaged, uploaded, or added to a site using documented steps.
- A default or documented configuration exposes a security, privacy, or data integrity risk.
- A core Day 1 admin or end-user journey dead-ends without a clear recovery path.
- A differentiator is presented as a product promise but is missing, fragile, or too hard to discover.
- A failure mode cannot be diagnosed from UI, docs, logs, or an exportable support artifact.

## 3.2 Audit Inputs

The audit should explicitly record the exact inputs used:

- Current branch and commit SHA
- `package.json` version, SPFx package versions, and `spfx-toolkit` dependency source/version
- Relevant generated package artifacts under `sharepoint/solution/` or `release/`, if present
- Prior docs read: requirements, March 22 audit, deployment/admin/provisioning guides, PnP alignment notes
- Commands run and whether they passed, failed, or were skipped with reason
- External docs consulted for PnP Modern Search v4 parity, with links and access date

## 4. Audit Document Structure

Single document at `docs/sp-search-launch-readiness-audit.md`. The 2026-03-22 audit moves to `docs/archive/sp-search-comprehensive-audit-2026-03-22.md` with a header note pointing at the new document.

### 4.1 Front Matter

- Date, scope, audience profile ("any SPFx tenant, self-serve, no hand-holding")
- Repo snapshot: branch, commit SHA, `package.json` version, SPFx version family, toolkit source/version
- Verification snapshot: type-check/test/package command results, or explicit skip reasons
- Stated differentiator priorities (the 5 the project is investing in)
- Reconciliation summary: count of March 22 findings closed vs. still-open vs. obsolete vs. changed-form (full table in Appendix A)
- Reading guide: how P0/P1/P2/Defer are defined; how to use the Roadmap Matrix

### 4.2 Part 1 — The Two Journeys

Narrative walkthroughs that anchor the audit in lived experience. ~3–4 pages each.

**Journey A: Day 1 Admin Install**

Walk every step a tenant admin takes from .sppkg in hand to working search experience on a published page:

1. Download / receive `.sppkg`
2. Upload to tenant or site app catalog
3. Add app to a site
4. Run provisioning script (`Setup-SPSearchSite.ps1`)
5. Run scenario presets script (`Search-ScenarioPresets.ps1`)
6. Open a page in edit mode, add Search Box / Verticals / Filters / Results / Manager
7. Configure searchContextId across web parts
8. Open property panes, configure scope / filters / columns / layout
9. Run a test query
10. Configure saved searches, sharing, history retention
11. Publish the page
12. Hand off to end users

Friction is logged inline with severity (Blocker / Confusion / Polish) and cross-referenced into the relevant differentiator track. Each friction point names the file or experience that produced it, plus a suggested owner track.

**Journey severity definitions:**

- **Blocker** — the user cannot complete the journey step safely or reliably.
- **Confusion** — the step is technically possible but requires guessing, prior author knowledge, or source-code reading.
- **Polish** — the step works but weakens trust, product quality, accessibility, or perceived maturity.

**Journey B: Day 1 End-User Search**

Walk an end user's first encounter:

1. Lands on a search page (no prior context)
2. Types a query
3. Sees suggestions (or doesn't)
4. Reads the empty state if no results
5. Switches verticals
6. Applies filters
7. Switches layouts
8. Opens detail panel
9. Saves the search
10. Shares the search with a colleague
11. Returns later via deep link
12. Repeats on mobile

Same friction-logging discipline.

### 4.3 Part 2 — Differentiator Tracks

Five tracks, identical sub-structure each:

- **Current state** — code-grounded summary with file references
- **Gap to "amazing"** — what an admin/user would expect from a "self-serve any tenant" launch that isn't there
- **Deliverables** — numbered list. Each deliverable has: short description, why it matters (tied to differentiator), effort tier, priority tier, dependencies, source friction (which journey step or audit finding it resolves)
- **Out of scope for v1.0** — explicit deferrals with one-line rationale

The five tracks:

#### T1. Modern UI Quality

Layout polish, empty states, loading shimmer, mobile responsiveness, dark mode story, theming consistency, micro-interactions, illustration vs icon strategy, error states, animation/transition quality, typography hierarchy, color usage.

#### T2. End-User Productivity

Saved searches, collections / pinboards, sharing, history, annotations, keyboard shortcuts, multi-select bulk actions, export (CSV today, XLSX deferred), Comments component integration (new spfx-toolkit capability), recent + trending suggestions, query templates, personal vs shared library boundaries.

#### T3. Multi-Instance / Multi-Context

`searchContextId` correctness, URL parameter namespacing (`?ctx1.q=...`), per-vertical `dataProviderId` routing, cross-context coordination patterns, sample multi-context pages, isolation guarantees under stress, library-component singleton backing (the `window.__sp_search_context_map__` pattern), regression risks when admins reuse context IDs across pages.

This track must include at least one explicit collision scenario and one explicit isolation scenario:

- Two independent search experiences on the same page with different context IDs.
- Two web parts accidentally sharing a context ID when the admin expected isolation.
- URL deep-link parameters for two contexts on the same page.
- Navigation away/back behavior and store cleanup expectations.

#### T4. Admin Experience

Property pane discoverability, scenario presets (`general`/`documents`/`news`/`people`/`media`/`custom` shipped; `knowledgeBase`/`hubSearch`/`policySearch` pending), schema picker UX, edit-mode validation/lint, provisioning script robustness, Admin Dashboard depth (Coverage Stats / Quality Metrics / Health / Insights), property pane error handling, default value sanity for unknown tenants.

#### T5. Observable & Diagnosable

Debug FAB feature completeness (Query / Network / State / Logs / Errors tabs), error surfacing patterns (toasts vs panels vs inline), "why no results" panel, telemetry hook strategy (opt-in, anonymous, what fields), Admin Dashboard analytics (Health + Insights), exportable support bundle (state snapshot + recent network calls + config), logging discipline (avoid PII).

This track must distinguish local diagnostics from product telemetry:

- **Local diagnostics**: user/admin-visible debug state, support bundle, recent network/search calls, config snapshot. Available without sending data anywhere.
- **Telemetry**: optional aggregate product signals. Opt-in only, disabled by default, never captures query text, user identity, result titles, URLs, tenant names, or list item content.

### 4.4 Part 3 — Foundations Track

Cross-cutting work that doesn't sit cleanly in any single differentiator but blocks launch.

- **Security hardening**: BUG-004 (`newPageUrl` XSS) plus a sweep for similar patterns; HTML sanitization adoption from spfx-toolkit; CSP-friendliness; review of all `window.location.href` and `dangerouslySetInnerHTML` usages
- **SPFx 1.22 / Heft migration completion**: the current branch (`feat/spfx-1.22-heft-migration`) is unmerged — verify nothing regressed, run a smoke checklist, decide merge criteria
- **Accessibility baseline**: target WCAG 2.1 AA; keyboard navigation, screen reader, focus management, contrast, ARIA roles, motion-reduction respect
- **CI / release engineering**: versioning policy (semver?), .sppkg build pipeline, smoke tests, release checklist, changelog convention
- **Documentation**: README → docs site decision; minimum viable docs: top-level README, per-web-part config reference, scenario gallery, troubleshooting / FAQ, contributing guide if open-sourced
- **Telemetry plumbing**: opt-in only; what is captured (query timing, error rates, NEVER queries themselves or PII); how admins enable/disable; storage location
- **Performance budgets**: define and enforce per-web-part bundle size budgets; document current sizes; add CI check for budget breach

### 4.5 Part 4 — Roadmap Matrix

Single sortable table. One row per deliverable. Columns:

| ID | Deliverable | Track(s) | Effort | Priority | Depends on | Source | Acceptance Signal |
|----|-------------|----------|--------|----------|------------|--------|-------------------|

`Source` references either a journey step (e.g., "Journey A step 4") or an audit finding (e.g., "T2.D7", "Foundations.S2"). `Acceptance Signal` is the smallest observable proof that the deliverable is done, such as "unit test covers URL namespace collision" or "admin guide has runnable tenant app catalog install path." This page is the executable artifact — what someone opens to pick the next thing to do.

### 4.6 Part 5 — Recommended Sprint Sequencing

Three suggested 2-week sprints (solo developer assumed):

- **Sprint 4 — Foundations + Critical UX**: security, 1.22 merge, top journey blockers, accessibility quick wins
- **Sprint 5 — Differentiator Depth**: bulk of T1–T5 P0/P1 deliverables
- **Sprint 6 — Polish + Docs**: remaining P1, docs site, sample gallery, release engineering

Each sprint lists 8–12 deliverables drawn from the matrix. P2 / Defer items listed at the end but not slotted.

### 4.7 Part 6 — Appendices

- **Appendix A — March 22 Audit Reconciliation**: every finding from the prior audit, status (Closed with commit ref / Still-Open / Obsolete / Changed-Form), and where in this audit it now appears
- **Appendix B — spfx-toolkit Integration Map**: each new toolkit capability matched to a deliverable in this audit (or marked "no fit")
- **Appendix C — PnP Modern Search v4 Parity Scorecard**: feature-by-feature grading (Better / Parity / Worse / Missing); informs positioning, not forced parity
- **Appendix D — Rejected Ideas**: ideas considered and dropped, one-line rationale each, so they don't keep coming back
- **Appendix E — Evidence and Command Log**: repo snapshot, commands run, command results, skipped checks, external sources consulted

## 5. Methodology

The audit will be produced in **a single sitting** at the user's request, with the following discipline:

1. **Code-grounded.** Every finding cites a file path and line range. Memory and prior docs are consulted as starting points but not treated as truth.
2. **March 22 reconciliation first.** Every prior finding is categorised before new findings are added. Avoids duplicate discovery and keeps Appendix A trustworthy.
3. **spfx-toolkit comparison pass.** Read the toolkit's recent commits and exports; for each new capability, identify whether it should replace existing custom code or unlock a new feature. Output: Appendix B.
4. **Journey simulation.** Walk Journeys A and B file-by-file (manifest → web part class → React tree → store). Friction logged inline with severity.
5. **PnP v4 parity scorecard.** Read PnP v4 docs (web-fetch as needed); grade each feature. Output: Appendix C.
6. **Differentiator track passes.** For each of T1–T5, dedicated read of the relevant subsystem under that lens.
7. **Foundations sweep.** Security, accessibility, build, CI, docs, telemetry, performance.
8. **No fixes.** This is purely diagnostic. Fixes belong to the per-track plans.

To keep the document navigable despite size, prose is kept tight; code references use `[file.ts:42](file.ts#L42)` markdown links; tables are used wherever a list of similar items repeats.

### 5.1 Evidence Standard

Each finding and deliverable should include:

- **Observed evidence**: file path + line number, command output summary, generated artifact path, or documentation URL.
- **Impact**: who is affected and how the issue appears in the Day 1 journey.
- **Recommendation**: one concrete next action, not a broad theme.
- **Acceptance signal**: what future implementation can prove without re-litigating the finding.

Avoid unsupported wording such as "probably", "seems", or "may be" unless the uncertainty itself is recorded as the finding and the next action is to verify it.

### 5.2 Recommended Verification Commands

Run these when feasible and record results in Appendix E. If any command is skipped, record the reason.

```bash
git status --short
git rev-parse --short HEAD
npm run type-check
npm test
npm run package
```

The audit remains "no code changes"; generated output from verification commands is evidence, not implementation. If generated files change during verification, note that in Appendix E and do not mix those generated changes with audit document edits unless the user explicitly approves.

### 5.3 Self-Review Pass

Before the audit is considered complete:

- Every P0 must cite the P0 admission rule category it satisfies.
- Every roadmap row must have a source and acceptance signal.
- Every journey blocker must map to at least one roadmap row.
- Every March 22 finding must be reconciled exactly once.
- Every toolkit capability in Appendix B must be marked Adopt / Consider / No Fit, with rationale.
- The final sprint sequencing must contain only roadmap IDs, not newly invented work.

## 6. Prioritization Framework

Each deliverable is tagged with two attributes.

**Effort tiers**

- **S** — ≤ 4 hours
- **M** — ½ to 1 day
- **L** — 1 to 3 days
- **XL** — more than 3 days

**Priority tiers**

- **P0 — Must ship in v1.0.** Without this, the launch is embarrassing or unsafe.
- **P1 — Should ship in v1.0.** Strongly elevates launch quality. Defer only under cost pressure.
- **P2 — v1.1+ candidate.** Tracked but not slotted into launch sprints.
- **Defer / Reject.** Considered and dropped, with one-line reason in Appendix D.

**P0 admission rule.** A finding may only be P0 if it ties to one of:

- (a) A stated differentiator (T1–T5)
- (b) Security
- (c) Data integrity
- (d) A "would prevent install" issue
- (e) A journey Blocker with no documented workaround

This rule keeps the P0 list honest and forces clear justification for anything blocking launch.

## 7. Outputs and Follow-up

**Audit deliverables** (this spec's scope):

1. `docs/sp-search-launch-readiness-audit.md` — the full document above
2. `docs/archive/sp-search-comprehensive-audit-2026-03-22.md` — moved with header redirect note
3. No runtime source code changes

**Follow-up plan deliverables** (separate session via writing-plans skill, after audit acceptance):

One implementation plan per track, **six plans total**:

- `docs/superpowers/plans/2026-MM-DD-foundations-plan.md`
- `docs/superpowers/plans/2026-MM-DD-modern-ui-quality-plan.md`
- `docs/superpowers/plans/2026-MM-DD-end-user-productivity-plan.md`
- `docs/superpowers/plans/2026-MM-DD-multi-context-plan.md`
- `docs/superpowers/plans/2026-MM-DD-admin-experience-plan.md`
- `docs/superpowers/plans/2026-MM-DD-observable-diagnosable-plan.md`

Per-track plans (rather than one master plan) are chosen because they are independently sized, parallelizable, mergeable in any order, and easier to retire or rescope without disturbing the rest. Foundations is sequenced first because security and 1.22 merge gate other work.

Each plan, when produced, will:

- Pull the corresponding track's deliverables from the Roadmap Matrix
- Sequence them with dependencies respected
- Define test criteria per deliverable
- Identify per-deliverable risk and mitigation
- Match the writing-plans skill's expected format

## 8. Risks and Mitigations

| Risk | Mitigation |
|------|------------|
| Audit becomes too long to be useful | Hard cap on deliverable count per track (target ~10, max 15); P2/Defer items captured as one-liners not paragraphs |
| P0 list inflation | P0 admission rule (Section 6) enforced in self-review pass |
| Findings drift from current code due to stale reading | All findings cite file + line; spot-check during self-review |
| Per-track plans duplicate work across tracks | Roadmap Matrix is single source of truth; per-track plans reference matrix IDs, not invent new ones |
| Audit blocks on perfectionism | Single-sitting production by user request; explicit "publish then iterate" stance — Appendix A and the matrix can be amended after launch |
| "No code changes" conflicts with moving the old audit | Treat documentation movement as an audit artifact; no source/runtime implementation changes |
| Generated artifacts change during verification | Record generated diffs in Appendix E and keep them out of the audit commit unless separately approved |
| External PnP docs drift after audit | Record source links and access date in Appendix E |

## 9. Acceptance Criteria

- [ ] `docs/sp-search-launch-readiness-audit.md` exists and contains all six Parts and Appendices A-E
- [ ] Every March 22 audit finding appears in Appendix A with a status
- [ ] Roadmap Matrix has at least one entry per track and every entry is sized, prioritized, sourced, and assigned an acceptance signal
- [ ] No P0 deliverable lacks a justification under the admission rule
- [ ] Every journey Blocker maps to at least one P0/P1 roadmap row or an explicit documented workaround
- [ ] Appendix E records repo snapshot, commands run, command results, skipped checks, and external sources
- [ ] Old audit moved to `docs/archive/` with header redirect
- [ ] User has reviewed and approved the audit before per-track plans are produced

## 10. Out of Scope (Explicit)

- Producing the per-track implementation plans (separate skill invocation)
- Any runtime source code changes
- Re-writing the requirements document
- A full PnP v4 feature parity push not tied to differentiators
- Production rollout planning, marketing, or naming/branding decisions
