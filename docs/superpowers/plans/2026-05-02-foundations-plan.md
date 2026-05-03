# Foundations Track Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Land all 13 Foundations track deliverables (Found.D1–Found.D13) so SP Search can ship a self-serve `.sppkg` against a tagged main commit, with budgets enforced, tests running, accessibility/security baselines in place, and the SPFx 1.22 / Heft migration merged.

**Architecture:** Two execution tranches inside one plan. **Tranche 1 (P0 ship-blockers, ~Sprint 4)**: Found.D1 → Found.D13 → Found.D7 (+ Task 3.5 amendment) → Found.D2 (Phase A only). **Tranche 2 (P1/P2 polish + spillover, ~Sprint 4 tail → Sprint 6)**: Found.D3 → Found.D5 → Found.D8 → Found.D6 → Found.D4 → Found.D10 → Found.D9 → Found.D11. Cross-track bidirectional dependencies declared in the audit Roadmap Matrix (T1.D9, T3.D9, T4.D6/D9/D10, T5.D8/D9) are preserved verbatim.

**Found.D12 reclassified to Defer (post Task 3.5).** The original Foundations plan placed D12 (Bundle reduction sweep, XL P0) as the final Tranche 1 task. Task 3.5 (Found.D7 amendment, commit `63b4cc1`) discovered that the audit's Phase 7 bundle inventory captured non-production sizes (~7.96–14.75 MB unhashed), but actual `heft build --production` minified output is 470K–862K — already at 22-43% of SPFx informal 1–2 MB guidance. The "install-blocking, 4–7× over guidance" framing that motivated D12 P0 was based on stale numbers. With Task 3.5's hash-aware gate enforcing the production budget, D12's active reduction work is no longer needed at v1.0. Reclassified to **Defer** (re-evaluate at v1.1+ if a sub-1 MB ceiling becomes a goal). The original Task 5 body is preserved below as a SUPERSEDED stub for audit-trail integrity.

**Tech Stack:** SPFx 1.22.2 + Heft 1.2.7 + React 17.0.1 + TypeScript 5.3.3 + Jest 29.7 + ts-jest 29.4.6 + GitHub Actions + axe-core + webpack-bundle-analyzer + DevExtreme 22.2.x + Fluent UI v8.106 + spfx-toolkit (file:../spfx-toolkit).

**Source authority:** `docs/sp-search-launch-readiness-audit.md` (head `3e74521`, Part 3 Foundations + Part 4 Roadmap Matrix Found.* rows + Part 5 Sprint 4 sequencing + Appendix E §4 verification commands). Every deliverable in this plan corresponds 1:1 with a Found.D* row in the Roadmap Matrix; no new work is invented here.

**P0 admission rule (per memory `feedback_p0_admission_rule_phrasing.md`):** Each P0 deliverable's Why field below uses the explicit `P0 rule: (X) <name>` tag form. Categories from spec §6: (a) differentiator, (b) security, (c) data integrity, (d) prevents install/build, (e) journey blocker.

**Note on tranche split vs audit Sprint 4:** The audit's Sprint 4 (Part 5) bundles Found.D1, D3, D5, D7, D8, D2, D12, D13, D6 + four cross-track items into ~20.5 dev-days. This plan keeps **all 13 Foundations deliverables in one document** (per user direction) and orders them as two tranches: Tranche 1 = D1, D13, D7, D2, D12 (the user-specified P0 execution tranche); Tranche 2 = D3, D5, D8, D6, D4, D10, D9, D11 (P1/P2 plus D5 P0 deferred per user direction). The Tranche 1 path treats D2's release tag as a **manual first cut** (per the audit body: "this deliverable ships the manual first run; D8 ships the automation"). Tranche 2 D8 retroactively wires the GitHub Actions automation. Tranche 2 D5 then closes the broader docs sweep — Tranche 1 D2 self-contains its CLAUDE.md gulp/1.21 sub-edits.

---

## File Structure

### Created

| Path | Purpose | Owner deliverable |
|------|---------|---------------------|
| `tests/styles/pnpPropertyControlsFix.test.ts` | Guard test for D1 ordering bug | Found.D1 |
| `tests/store/lifecycle.test.ts` | T3.D9 trail-marker smoke test exercising Heft Jest harness | Found.D13 |
| `tests/utils/safeNavigate.test.ts` | 5-case unit test for navigation policy | Found.D4 |
| `tests/a11y/smokeAxe.test.tsx` | axe-core jest-axe smoke test (4 render shapes) | Found.D6 |
| `src/libraries/spSearchStore/utils/safeNavigate.ts` | Centralised `safeNavigate(target)` helper | Found.D4 |
| `src/styles/motion.scss` | Shared `prefers-reduced-motion` mixin | Found.D6 |
| `src/libraries/spSearchStore/telemetry/TelemetryTransport.ts` | HTTPS POST transport + batch scheduler + config loader | Found.D9 |
| `src/libraries/spSearchStore/telemetry/ITelemetryConfig.ts` | Config interface (mirrors SP list shape) | Found.D9 |
| `config/bundle-budgets.json` | Per-web-part byte budgets | Found.D7 |
| `scripts/check-bundle-sizes.js` | Node script that exits non-zero on budget breach | Found.D7 |
| `.github/workflows/build.yml` | PR + main CI: install, lint, type-check, test, package, bundle-gate, artifact upload | Found.D8 |
| `.github/workflows/release.yml` | Tag-push (`v*`) CI: rebuild + GitHub Releases + auto release notes | Found.D8 |
| `.github/PULL_REQUEST_TEMPLATE.md` | PR template with affected web-part / repro / screenshot / env sections | Found.D8 |
| `.github/ISSUE_TEMPLATE/bug_report.md` | Bug template | Found.D8 |
| `.github/ISSUE_TEMPLATE/feature_request.md` | Feature template | Found.D8 |
| `.github/dependabot.yml` | Weekly devDeps PRs, auto-merge disabled | Found.D8 |
| `CHANGELOG.md` | Keep-a-Changelog history (pre-populated from BUG-001..BUG-012 closures) | Found.D8 |
| `CONTRIBUTING.md` | Short contributor guide → CLAUDE.md + smoke checklist + semver | Found.D8 |
| `README.md` | Top-level orientation → docs/* + CHANGELOG + Releases | Found.D5 |
| `docs/release-policy.md` | SemVer 2.0 policy + lockstep convention | Found.D8 |
| `docs/release-smoke-checklist.md` | 7-step pre-merge gate enumeration | Found.D2 |
| `docs/release-runs/v1.0.0-rc.1.md` | First smoke-checklist run log | Found.D2 |
| `docs/performance-budgets.md` | Current/budget/aspirational per-web-part table | Found.D7 |
| `docs/accessibility.md` | WCAG 2.1 AA scoped conformance statement | Found.D6 |
| `docs/privacy-notice.md` | Telemetry "what we collect / never collect" | Found.D9 |
| `docs/performance/bundle-sizes-baseline.json` | Pre-Sprint-4 baseline snapshot of `release/analysis-logs/bundle-sizes.json` (the latter is gitignored, regenerated each build) | Found.D7 |

### Modified

| Path | Change | Owner deliverable |
|------|--------|---------------------|
| `src/styles/pnpPropertyControlsFix.ts` | Hoist `PNP_COLLECTION_DATA_CSS` const above `ensurePnpPropertyControlStyles()` | Found.D1 |
| `package.json` | `type-check` script → `tsc --noEmit -p tsconfig.json`; align `version` to `1.0.0` (D11 lockstep with package-solution) | Found.D3, Found.D11 |
| `jest.config.js` | Compatible with Heft `heft test` invocation; ensure `tests/store/lifecycle.test.ts` resolves | Found.D13 |
| `config/package-solution.json` | Add `webApiPermissionRequests` for the Microsoft Graph scope verified in Task 11 Step 1; clear `developer.mpnId`; populate `developer.{websiteUrl,privacyUrl,termsOfUseUrl}`; bump `solution.version` lockstep to `1.0.0.0` | Found.D10, Found.D11 |
| `src/webparts/spSearchResults/components/documentTitleUtils.ts` | Delete local `sanitizeSummaryHtml` (line 158-end); replace with toolkit re-export | Found.D4 |
| `src/webparts/spSearchResults/components/ListLayout.tsx` | Import sanitizer from `spfx-toolkit/lib/utilities/htmlUtils/sanitizeHtml` | Found.D4 |
| `src/libraries/spSearchStore/providers/QuickResultsSuggestionProvider.ts:80` | Route through `safeNavigate` | Found.D4 |
| `src/webparts/spSearchManager/components/SpSearchManager.tsx:646` | Route through `safeNavigate` (lines 369/548/644 are `read`-only `new URL(window.location.href)`, exempt) | Found.D4 |
| `src/webparts/spSearchResults/components/SpSearchResults.tsx:570` | Route through `safeNavigate` (line 568 is `read`-only `new URL(...)`, exempt) | Found.D4 |
| `src/webparts/spSearchBox/components/SpSearchBox.tsx:358` | Convert existing inline allowlist to call `safeNavigate` | Found.D4 |
| `src/webparts/spSearchBox/components/SpSearchBox.tsx:729` | Add `aria-describedby` linking to a hidden description span (A11Y-005) | Found.D6 |
| `src/webparts/spSearchBox/components/SpSearchBox.tsx:736` | Replace `<div role="radiogroup">` with `<fieldset>` + `<legend>` (A11Y-004) | Found.D6 |
| `src/webparts/spSearchBox/components/SpSearchBox.module.scss` (+5 others) | Wrap `transition`/`animation`/`@keyframes` rules in motion mixin | Found.D6 |
| `scripts/Provision-SPSearchLists.ps1` | Add provisioning for `SearchTelemetryConfig` (single-row, hidden) + `SearchTelemetryOptIn` (per-user) lists | Found.D9 |
| `docs/admin-guide.md:221-226` | Manager defaults table → match `SpSearchManagerWebPart.manifest.json:29-32` (all `false`) | Found.D5 |
| `docs/provisioning-guide.md:131-132` | Replace "no automatic background cleanup" with the actual 24h sweep + 90-day retention text | Found.D5 |
| `docs/deployment-guide.md:18-22, 164` | Replace `gulp bundle --ship` / `gulp package-solution --ship` with `npm run package`; replace `npm run type-check` with `tsc --noEmit -p tsconfig.json` reference if D3 picked option (a) | Found.D3, Found.D5 |
| `CLAUDE.md` lines 7, 48, 115, 305-309, 433-439, 458, 464, 492-493 | Drop `SPFx 1.21.1`, `gulp serve`, `gulp bundle --ship`, `gulp clean-cache`, `Toast / ToastProvider` claims, shipped-as-backlog presets/XLSX | Found.D2 (lines 7/48/115/492-493) + Found.D5 (lines 305-309, 433-439, 458, 464) |
| `~/.claude/projects/-Users-hemantmane-Development-sp-search/memory/MEMORY.md` | Drop the same shipped-but-listed-as-backlog drift on the "Sprint 4: BACKLOG" line per audit D5 (e) | Found.D5 |

### Files explicitly NOT touched in this plan

- `src/styles/*.module.scss` style-loader pipeline beyond motion mixin (out-of-scope per audit)
- Source-mapped production builds (out-of-scope per audit)
- spfx-toolkit `file:../` link → npm version migration (deferred to v1.1 per audit Out-of-scope item 5)
- Window-key namespace consolidation (deferred to v1.1 per T3 OOS item 7)

---

# Tranche 1 — P0 execution (Sprint 4 ship-blockers)

Order rationale: D1 unlocks the build (else nothing else can package). D13 unlocks the test gate (else D2 pre-merge regression cannot run). D7 publishes the budget that D12 reduces toward (else D12 has no acceptance criterion). D2 ships the manual `v1.0.0-rc.1` release tag that D11/D8 later automate. D12 closes the bundle reduction sweep before any Tranche-2 P1 work lands additional code into the not-yet-budget-enforced web parts.

## Task 1: Found.D1 — Fix `pnpPropertyControlsFix.ts` lint blocker so `npm run package` produces `.sppkg`

**Why (P0 rule: (d) prevents install/build):** `npm run package` halts at ESLint stage before TypeScript or Webpack runs because `src/styles/pnpPropertyControlsFix.ts:33` references `PNP_COLLECTION_DATA_CSS` before its declaration on `:42`. SARIF evidence: `release/analysis-logs/lint.sarif` shows `"level": "error"`, `"text": "'PNP_COLLECTION_DATA_CSS' was used before it was defined."`. No `.sppkg` is produced through documented commands. Audit Found.D1.

**Files:**
- Modify: `src/styles/pnpPropertyControlsFix.ts:15-175` (hoist const)
- Create: `tests/styles/pnpPropertyControlsFix.test.ts`

- [ ] **Step 1: Read the broken file** — open `src/styles/pnpPropertyControlsFix.ts` and confirm `PNP_COLLECTION_DATA_CSS` is declared on `:42` and used on `:33`.

- [ ] **Step 2: Write the failing guard test**

Create `tests/styles/pnpPropertyControlsFix.test.ts`. The module holds an `injected` boolean at module scope, so the test resets the module registry between specs to ensure each `import` returns a fresh instance with `injected = false`:

```typescript
describe('pnpPropertyControlsFix', () => {
  beforeEach(() => {
    document.head.querySelectorAll('#sp-search-pnp-property-controls-fix').forEach(n => n.remove());
    jest.resetModules();
  });

  it('injects the style tag with non-empty content', async () => {
    const { ensurePnpPropertyControlStyles } = await import('../../src/styles/pnpPropertyControlsFix');
    ensurePnpPropertyControlStyles();
    const tag = document.getElementById('sp-search-pnp-property-controls-fix');
    expect(tag).not.toBeNull();
    expect(tag!.textContent).toContain('.collectionData_f8375039');
    expect(tag!.textContent).toContain('.tableRow_f8375039');
  });

  it('is idempotent across multiple calls within one module instance', async () => {
    const { ensurePnpPropertyControlStyles } = await import('../../src/styles/pnpPropertyControlsFix');
    ensurePnpPropertyControlStyles();
    ensurePnpPropertyControlStyles();
    expect(document.querySelectorAll('#sp-search-pnp-property-controls-fix')).toHaveLength(1);
  });

  it('re-injects after a fresh module load (proves module-level injected flag is what guards repeats)', async () => {
    const first = await import('../../src/styles/pnpPropertyControlsFix');
    first.ensurePnpPropertyControlStyles();
    document.head.querySelectorAll('#sp-search-pnp-property-controls-fix').forEach(n => n.remove());
    jest.resetModules();
    const second = await import('../../src/styles/pnpPropertyControlsFix');
    second.ensurePnpPropertyControlStyles();
    expect(document.getElementById('sp-search-pnp-property-controls-fix')).not.toBeNull();
  });
});
```

The `jest.resetModules()` call discards the cached module, so each `await import(...)` re-evaluates the module body and resets `injected` to `false`. Without this, the second test's `ensurePnpPropertyControlStyles()` call would early-return (because `injected === true` from the first test) and the assertion would fail even though the module is correct.

- [ ] **Step 3: Run lint to verify the existing failure**

Run: `npm test 2>&1 | head -30`
Expected: `'PNP_COLLECTION_DATA_CSS' was used before it was defined.` at `pnpPropertyControlsFix.ts:33:23`. Build halts before Jest invocation.

- [ ] **Step 4: Hoist the constant** — edit `src/styles/pnpPropertyControlsFix.ts` so the file order is: file-level JSDoc → `STYLE_TAG_ID` → `injected` → `PNP_COLLECTION_DATA_CSS` (the entire block previously on `:42-175`) → `ensurePnpPropertyControlStyles()`. Keep the inline JSDoc on the const explaining the baked `_f8375039` hash. Do **not** add an `eslint-disable` comment — prefer the mechanical reorder per audit body shape (a).

- [ ] **Step 5: Verify lint now passes**

Run: `npm test 2>&1 | tail -30`
Expected: ESLint reports zero `@typescript-eslint/no-use-before-define` violations against `pnpPropertyControlsFix.ts`. Build proceeds past lint to TypeScript/Jest. (If Found.D13 has not yet landed, Jest may still fail — that is Task 2's domain.)

- [ ] **Step 6: Run end-to-end package**

Run: `npm run package 2>&1 | tail -20`
Expected: Heft completes both `build --clean --production` and `package-solution --production`. `sharepoint/solution/sp-search.sppkg` exists.

Verify: `ls -la sharepoint/solution/sp-search.sppkg`
Expected: file exists with current timestamp.

- [ ] **Step 7: Regenerate SARIF and confirm zero error-level entries**

Run: `npm run package 2>&1 | tee /tmp/d1-build.log; grep -c '"level": "error"' release/analysis-logs/lint.sarif || echo 0`
Expected: `0`.

Note: `release/analysis-logs/lint.sarif` lives under the gitignored `release/` tree (per `.gitignore`) and is regenerated on every build. We do **not** commit the SARIF file. The SARIF before/after evidence is captured in the commit message text and in the smoke-checklist run log (Task 4) instead. If long-term SARIF retention is desired, copy the post-fix SARIF to `docs/release-runs/v1.0.0-rc.1.sarif` (under tracked `docs/`) as part of the Task 4 run log.

- [ ] **Step 8: Commit**

```bash
git add src/styles/pnpPropertyControlsFix.ts tests/styles/pnpPropertyControlsFix.test.ts
git commit -m "$(cat <<'EOF'
fix(styles): hoist PNP_COLLECTION_DATA_CSS above ensurePnpPropertyControlStyles (Found.D1)

ESLint @typescript-eslint/no-use-before-define halted Heft lint at
src/styles/pnpPropertyControlsFix.ts:33, blocking npm run package /
npm test. Mechanical reorder hoists the const above its first use; no
rule suppression. Adds tests/styles/pnpPropertyControlsFix.test.ts
guard (with jest.resetModules() between specs to defeat the
module-level injected flag) so a future re-introduction of the same
ordering trips Jest.

Pre-fix SARIF (release/analysis-logs/lint.sarif, gitignored):
  "level": "error", "text": "'PNP_COLLECTION_DATA_CSS' was used before it was defined."
Post-fix SARIF: 0 error-level entries.
.sppkg lands at sharepoint/solution/sp-search.sppkg per
config/package-solution.json:38.

Closes Found.D1 P0 (audit Part 3 + Part 4 + Appendix E §4 row 5).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 2: Found.D13 — Fix Jest harness `ts-jest`/`jest-util` resolution failure for `npm test`

**Why (P0 rule: (d) prevents install/build):** Without a working test harness, every track's "unit test passes" acceptance signal cannot be verified by CI. T3.D9 (lifecycle integration tests) and T5.D2/D3/D6/D7 unit-test acceptance signals all silently invalidate. The "self-serve any tenant" launch profile cannot ship a release artifact whose test gate does not run. Audit Found.D13. Independent of D1 in cause but D1 must precede so `npm test` reaches the Jest stage.

**Files:**
- Modify: `jest.config.js` (verify SPFx-compatible Heft test alignment)
- Modify: `package.json` (devDeps pin verification only — do not regress `ts-jest@^29.4.6` / `jest@^29.7.0` / `jest-util@^29.7.0`)
- Create: `tests/store/lifecycle.test.ts` (placeholder smoke test; T3.D9 adopts later)

- [ ] **Step 1: Diagnose the current failure** — run `npm test` (post-D1) and capture the exact resolution error.

Run: `npm test 2>&1 | tail -60 > /tmp/d13-diag.log; cat /tmp/d13-diag.log`
Expected: one of (a) Heft `heft test` succeeds and reports "no tests found" (the documented "broken" state is stale and we add the smoke test only), (b) `Cannot find module 'jest-util'` / `ts-jest` resolution failure with exact missing dep path, or (c) some other failure.

- [ ] **Step 1b: Discover the SPFx Heft Jest config shape — DO NOT delete `jest.config.js` until this completes**

Before committing to any fix shape, inspect what the local SPFx Heft installation actually expects:

```bash
# Is heft-jest-plugin installed?
find node_modules -maxdepth 4 -name "heft-jest-plugin" -type d 2>/dev/null
ls node_modules/@rushstack/heft-jest-plugin/ 2>/dev/null | head
ls node_modules/@microsoft/spfx-heft-plugins/lib/heftPlugins/ 2>/dev/null | head

# What does the rig package expect?
grep -rn "jest" node_modules/@microsoft/spfx-web-build-rig/heft.json node_modules/@microsoft/spfx-web-build-rig/lib/ 2>/dev/null | head

# Is there a documented shared-config schema we should follow?
find node_modules/@rushstack/heft-jest-plugin -name "*.json" -path "*shared*" 2>/dev/null
find node_modules/@rushstack/heft-jest-plugin -name "*.json" -path "*default*" 2>/dev/null

# What does Heft's own test schema look like?
cat node_modules/@rushstack/heft-jest-plugin/lib/exports/jest-shared.config.json 2>/dev/null | head -40
```

Record findings in `/tmp/d13-discovery.md`. The output of these commands drives whether the fix shape is (b.i) Heft-shared-config, (b.ii) keep standalone, or (c) something else (e.g. SPFx 1.22 ships its own Jest config under `@microsoft/spfx-heft-plugins/heftPlugins/JestPlugin/` that we should adopt verbatim).

- [ ] **Step 2: Pick the resolution shape based on discovery**

If Step 1 shape (a) — `heft test` already runs cleanly: proceed directly to Step 4 (add smoke test only). Commit message notes the documented "broken" state was stale post `a5f28c1` Heft migration.

If Step 1 shape (b) AND Step 1b confirms `@rushstack/heft-jest-plugin/lib/exports/jest-shared.config.json` exists with the documented schema: proceed with (b.i) — adopt the Heft shared config.

If Step 1 shape (b) AND Step 1b shows the SPFx rig has its own Jest plugin shape (different schema, different config path): adapt the Step 3 config file to match what the rig actually expects — DO NOT use the (b.i) template verbatim if the schema differs.

If Step 1 shape (b) AND Step 1b shows the rig does not bundle a Jest plugin at all: fall back to (b.ii) — keep `jest.config.js`, pin compatible `ts-jest` / `jest-util` / `jest` versions, run `npm rebuild jest-util ts-jest`. Verify with `npx jest --listTests`.

Document the chosen path in `CLAUDE.md` Common Commands section + `docs/release-smoke-checklist.md` (D2).

- [ ] **Step 3: Apply the chosen fix**

If (b.i) AND Step 1b confirmed the schema matches:

Create `config/jest-shared-config.json` (verify the `extends` path is exactly what `find node_modules/@rushstack/heft-jest-plugin -name "jest-shared.config.json"` reported in Step 1b — do not copy this template verbatim if the path differs):

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/heft-jest-plugin/v0/jest-schema.json",
  "extends": "@rushstack/heft-jest-plugin/lib/exports/jest-shared.config.json",
  "rootDir": "../",
  "testMatch": ["<rootDir>/tests/**/*.test.ts", "<rootDir>/tests/**/*.test.tsx"],
  "moduleNameMapper": {
    "\\.(css|scss)$": "<rootDir>/tests/__mocks__/styleMock.js",
    "^@pnp/(.*)$": "<rootDir>/tests/__mocks__/pnpMock.js",
    "^@microsoft/(.*)$": "<rootDir>/tests/__mocks__/pnpMock.js",
    "^spfx-toolkit/lib/utilities/context(.*)$": "<rootDir>/tests/__mocks__/spfxContextMock.js",
    "^spfx-toolkit/(.*)$": "<rootDir>/node_modules/spfx-toolkit/$1",
    "^@store/(.*)$": "<rootDir>/src/libraries/spSearchStore/$1",
    "^@interfaces/(.*)$": "<rootDir>/src/libraries/spSearchStore/interfaces/$1",
    "^@services/(.*)$": "<rootDir>/src/libraries/spSearchStore/services/$1",
    "^@providers/(.*)$": "<rootDir>/src/libraries/spSearchStore/providers/$1",
    "^@registries/(.*)$": "<rootDir>/src/libraries/spSearchStore/registries/$1",
    "^@orchestrator/(.*)$": "<rootDir>/src/libraries/spSearchStore/orchestrator/$1",
    "^@webparts/(.*)$": "<rootDir>/src/webparts/$1"
  },
  "transformIgnorePatterns": ["node_modules/(?!(@pnp|spfx-toolkit)/)"],
  "collectCoverageFrom": [
    "src/libraries/**/*.ts",
    "!src/libraries/**/index.ts",
    "!src/**/*.d.ts"
  ]
}
```

Then delete `jest.config.js`.

If (b.ii): pin devDep versions in `package.json` and add `npm rebuild` step.

- [ ] **Step 4: Add the trail-marker smoke test**

Create `tests/store/lifecycle.test.ts`:

```typescript
/**
 * Trail-marker smoke test for the Jest harness post-Found.D13.
 * Real lifecycle assertions land in T3.D9 (Sprint 6 dep on this fix).
 * Until then this test exists only to prove npm test runs at least one
 * spec end-to-end through the SPFx-Heft Jest pipeline.
 */
describe('Foundations Found.D13 — Jest harness smoke', () => {
  it('runs at least one spec to completion', () => {
    expect(1 + 1).toBe(2);
  });

  it('resolves ts-jest TypeScript transform', () => {
    const value: number = 42;
    expect(value).toBe(42);
  });
});
```

- [ ] **Step 5: Verify the harness runs end-to-end**

Run: `npm test 2>&1 | tail -20`
Expected: Jest discovers and runs `tests/store/lifecycle.test.ts` (and any other existing test under `tests/`). Output reports `2 passed` minimum (the two specs in the new file). Process exits 0.

Run: `npm test -- --testPathPattern lifecycle.test.ts 2>&1 | tail -10`
Expected: same — `Tests: 2 passed`.

- [ ] **Step 6: Update CLAUDE.md Common Commands**

Edit `CLAUDE.md` lines 488-501 — replace the bare `npx jest` block with:

```markdown
## Common Commands

```bash
# Development
npm start                                     # heft start --clean (local workbench)
npm run package                               # heft build --clean --production && heft package-solution --production

# Testing
npm test                                      # heft test (Heft-managed Jest invocation)
npm test -- --watch                           # watch mode
npm test -- --testPathPattern <pattern>       # filtered run

# spfx-toolkit (in toolkit directory)
cd /Users/hemantmane/Development/spfx-toolkit && npm run build
```
```

(Note: this overlaps with Tranche 2 D5; D5 owns the `gulp` → npm sweep across the rest of CLAUDE.md. Here we only touch the Common Commands block to align with D13.)

- [ ] **Step 7: Commit**

```bash
git add config/jest-shared-config.json tests/store/lifecycle.test.ts CLAUDE.md
git rm jest.config.js  # if (b.i) chosen
git commit -m "$(cat <<'EOF'
test(harness): align Jest harness with Heft test pipeline (Found.D13)

Switches to @rushstack/heft-jest-plugin shared config (config/jest-shared-config.json)
and removes the standalone jest.config.js that was tripping
ts-jest / jest-util resolution under heft test. Adds
tests/store/lifecycle.test.ts as a trail-marker smoke spec so
npm test runs at least one spec to completion through the migrated
Heft test pipeline. T3.D9 (Sprint 6) adopts the same file for full
disposeStore lifecycle assertions.

CLAUDE.md Common Commands block updated to npm test (Heft) — broader
gulp/1.21 sweep deferred to Found.D5 in Tranche 2.

Closes Found.D13 P0 (audit Part 3 + Part 4 + CLAUDE.md Sprint 4 backlog).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: Found.D7 — Per-web-part bundle size budget + CI breach gate + attribution dashboard

**Why (P0 rule: (d) prevents install/build):** Per-web-part bundles range 7.96–14.75 MB before lazy chunks; a typical four-web-part page is ~46 MB unminified. SPFx informal guidance is 1–2 MB. Without per-web-part attribution, every Adopt-class deliverable (T2.D4, T2.D5, T2.D9) ships blind into a budget no one has set; one careless adoption flips bundles from "very large" to "site-blocking." Audit Found.D7. CI workflow consumer is Tranche 2 D8.

**Files:**
- Create: `docs/performance-budgets.md`
- Create: `config/bundle-budgets.json`
- Create: `scripts/check-bundle-sizes.js`
- Create: `docs/performance/bundle-sizes-baseline.json` (a snapshot copy of the gitignored `release/analysis-logs/bundle-sizes.json` — `release/` is in `.gitignore`, so the per-PR script output cannot be committed without `-f`. The baseline lives under tracked `docs/performance/` instead.)

- [ ] **Step 1: Write the per-web-part current/budget/aspirational table**

Create `docs/performance-budgets.md`:

```markdown
# SP Search — Performance Budgets

> Owned by Foundations track (Found.D7). CI gate at `scripts/check-bundle-sizes.js`. Adoption deliverables that breach a budget block at the PR review surface; coordinate with track lead before raising a budget.

## Per-web-part bundle size budgets (production)

Sizes are unminified `release/assets/sp-search-*-web-part.js` post `heft build --clean --production`. SPFx informal "good citizen" guidance is 1–2 MB per web part; the smallest SP Search bundle is ~4× over that guidance and the largest is ~7× over.

| Web part | Current (bytes) | Budget (1.5×) | Aspirational (50%) | Notes |
|----------|-----------------|---------------|--------------------|-------|
| sp-search-filters-web-part.js | 14,752,081 | 22,128,121 | 7,376,040 | Tree-shake DevExtreme TreeView/FilterBuilder (Found.D12) |
| sp-search-results-web-part.js | 14,503,213 | 21,754,819 | 7,251,606 | Tree-shake DevExtreme DataGrid; defer ResultDetailPanel |
| sp-search-admin-manager-web-part.js | 11,123,315 | 16,684,972 | 5,561,657 | Defer admin-only Insights chunks |
| sp-search-manager-web-part.js | 11,116,954 | 16,675,431 | 5,558,477 | Defer SearchHistory + SearchCollections panels |
| sp-search-box-web-part.js | 8,488,976 | 12,733,464 | 4,244,488 | Defer KQL completion dropdown |
| sp-search-verticals-web-part.js | 7,956,603 | 11,934,904 | 3,978,301 | Smallest bundle; baseline target for v1.1+ |

## Lazy chunk inventory (consumed on demand)

| Chunk | Size (bytes) | Loaded by |
|-------|--------------|-----------|
| chunk.vendors-fluentui-Dialog | 7,103,488 | SearchManager + AdminManager dialogs |
| chunk.vendors-devextreme-react_core | 3,706,880 | DataGrid Layout (Results) |
| chunk.xlsx_xlsx_mjs | 2,621,440 | DataGrid CSV/XLSX export |
| chunk.vendors-devextreme-react_date-box | 1,124,352 | DateRange filter |
| chunk.spfx-toolkit_PeoplePicker | 2,001,920 | People-picker filter |
| chunk.spfx-toolkit_SearchManager | 545,792 | SearchManager panel |
| chunk.spfx-toolkit_VersionHistory | 450,560 | Detail panel version tab |
| chunk.spfx-toolkit_DataGridContent | 14,336 | DataGrid Layout |

## Enforcement

`scripts/check-bundle-sizes.js` runs after `heft build --clean --production`, reads `release/assets/sp-search-*-web-part.js` via `fs.statSync`, compares against `config/bundle-budgets.json`, exits non-zero on breach. CI wiring lives in `.github/workflows/build.yml` (Found.D8).

Per-PR attribution dashboard at `release/analysis-logs/bundle-sizes.json` enumerates the top-10 contributing modules per web part — consumed by reviewers to identify which adopted dependency drove a budget breach.

## Roadmap link

Active reduction work belongs to Found.D12 (Tranche 1, P0). Aspirational column targets v1.1+ (post-Sprint 6). Out-of-scope items per audit Foundations Out-of-scope §1: source-mapped production builds.
```

- [ ] **Step 2: Write the budgets JSON**

Create `config/bundle-budgets.json`:

```json
{
  "$schema": "./bundle-budgets.schema.json",
  "policy": "per-web-part-byte-ceiling",
  "owner": "Foundations.Found.D7",
  "budgets": {
    "sp-search-filters-web-part.js": 22128121,
    "sp-search-results-web-part.js": 21754819,
    "sp-search-admin-manager-web-part.js": 16684972,
    "sp-search-manager-web-part.js": 16675431,
    "sp-search-box-web-part.js": 12733464,
    "sp-search-verticals-web-part.js": 11934904
  },
  "aspirational": {
    "sp-search-filters-web-part.js": 7376040,
    "sp-search-results-web-part.js": 7251606,
    "sp-search-admin-manager-web-part.js": 5561657,
    "sp-search-manager-web-part.js": 5558477,
    "sp-search-box-web-part.js": 4244488,
    "sp-search-verticals-web-part.js": 3978301
  }
}
```

- [ ] **Step 3: Write the breach-gate script**

Create `scripts/check-bundle-sizes.js`:

```javascript
#!/usr/bin/env node
/**
 * Per-web-part bundle size breach gate (Foundations Found.D7).
 * Reads release/assets/sp-search-*-web-part.js, compares against
 * config/bundle-budgets.json, exits non-zero on breach.
 * Emits release/analysis-logs/bundle-sizes.json for the per-PR attribution dashboard.
 */

const fs = require('fs');
const path = require('path');

const REPO_ROOT = path.resolve(__dirname, '..');
const ASSETS_DIR = path.join(REPO_ROOT, 'release', 'assets');
const BUDGETS_PATH = path.join(REPO_ROOT, 'config', 'bundle-budgets.json');
const REPORT_PATH = path.join(REPO_ROOT, 'release', 'analysis-logs', 'bundle-sizes.json');

if (!fs.existsSync(BUDGETS_PATH)) {
  console.error(`[bundle-gate] missing budgets file: ${BUDGETS_PATH}`);
  process.exit(2);
}

const { budgets } = JSON.parse(fs.readFileSync(BUDGETS_PATH, 'utf8'));
const breaches = [];
const report = { generatedAt: new Date().toISOString(), webParts: {} };

for (const [name, budget] of Object.entries(budgets)) {
  const file = path.join(ASSETS_DIR, name);
  if (!fs.existsSync(file)) {
    console.error(`[bundle-gate] missing asset: ${file} (run heft build --production first)`);
    process.exit(2);
  }
  const actual = fs.statSync(file).size;
  const delta = actual - budget;
  const pct = ((actual / budget) * 100).toFixed(1);
  const status = actual <= budget ? 'PASS' : 'BREACH';
  console.log(`[${status}] ${name}: ${actual.toLocaleString()} bytes (budget ${budget.toLocaleString()}, ${pct}%, delta ${delta >= 0 ? '+' : ''}${delta.toLocaleString()})`);
  report.webParts[name] = { actual, budget, delta, pctOfBudget: Number(pct), status };
  if (status === 'BREACH') breaches.push({ name, actual, budget, delta });
}

fs.mkdirSync(path.dirname(REPORT_PATH), { recursive: true });
fs.writeFileSync(REPORT_PATH, JSON.stringify(report, null, 2));
console.log(`[bundle-gate] report written: ${REPORT_PATH}`);

if (breaches.length > 0) {
  console.error(`\n[bundle-gate] FAILED: ${breaches.length} web part(s) exceed budget`);
  for (const b of breaches) {
    console.error(`  ${b.name}: ${b.actual.toLocaleString()} bytes exceeds ${b.budget.toLocaleString()} budget by ${b.delta.toLocaleString()} bytes`);
  }
  process.exit(1);
}

console.log(`\n[bundle-gate] OK: all ${Object.keys(budgets).length} web parts within budget`);
process.exit(0);
```

Make executable: `chmod +x scripts/check-bundle-sizes.js`

- [ ] **Step 4: Verify the gate passes against current sizes**

Run: `node scripts/check-bundle-sizes.js 2>&1`
Expected: `[bundle-gate] OK: all 6 web parts within budget`. Each web part reports `[PASS] sp-search-X-web-part.js: <bytes> bytes (budget Y, ~67%, delta -<bytes>)`. Exit code 0.

Verify: `cat release/analysis-logs/bundle-sizes.json | head -10`
Expected: well-formed JSON with `generatedAt` + `webParts` keys.

- [ ] **Step 5: Verify the gate fails on a deliberate breach (sanity-check)**

Run: `cp release/assets/sp-search-filters-web-part.js release/assets/sp-search-filters-web-part.js.bak; dd if=/dev/zero bs=1 count=8000000 >> release/assets/sp-search-filters-web-part.js 2>/dev/null; node scripts/check-bundle-sizes.js; echo "exit=$?"; mv release/assets/sp-search-filters-web-part.js.bak release/assets/sp-search-filters-web-part.js`
Expected: `[BREACH] sp-search-filters-web-part.js: ...`, exit code 1. Then file restored.

- [ ] **Step 6: Add npm script for invocation**

Edit `package.json` `scripts` block — add:

```json
"check:bundles": "node scripts/check-bundle-sizes.js"
```

(Insert between `clean:all` and `stats`.)

- [ ] **Step 7: Verify the npm script works**

Run: `npm run check:bundles 2>&1 | tail -10`
Expected: same output as Step 4.

- [ ] **Step 8: Snapshot the baseline under tracked docs/performance/**

`release/analysis-logs/bundle-sizes.json` is gitignored. Copy a baseline snapshot under tracked `docs/performance/` so reviewers can diff future runs without `git add -f`:

```bash
mkdir -p docs/performance
cp release/analysis-logs/bundle-sizes.json docs/performance/bundle-sizes-baseline.json
```

- [ ] **Step 9: Commit**

```bash
git add docs/performance-budgets.md docs/performance/bundle-sizes-baseline.json config/bundle-budgets.json scripts/check-bundle-sizes.js package.json
git commit -m "$(cat <<'EOF'
build(perf): per-web-part bundle size budgets + CI breach gate (Found.D7)

Defines per-web-part byte budgets at 1.5x current sizes (Filters 22MB,
Results 22MB, Manager/Admin 17MB, Box 13MB, Verticals 12MB) plus
aspirational 50% targets for v1.1+. scripts/check-bundle-sizes.js exits
non-zero on breach and emits release/analysis-logs/bundle-sizes.json
(gitignored, regenerated each build). Baseline snapshot committed at
docs/performance/bundle-sizes-baseline.json. CI wiring deferred to
Found.D8 build.yml.

Active reduction work to bring sizes toward the 1.5x ceiling lives in
Found.D12 (Tranche 1).

Closes Found.D7 P0 (audit Part 3 + Part 4 + Phase 7 inventory).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 4: Found.D2 — Merge SPFx 1.22 / Heft migration branch with smoke checklist + release tag

**Why (P0 rule: (d) prevents install/build + (e) Journey A Step 12 [Polish]):** The 91-commit `feat/spfx-1.22-heft-migration` branch is the de facto shipping branch but has never landed on `main`. Admins cloning `main` get the prior SPFx 1.21 + gulp build chain; admins cloning the feature branch get the un-shipped Heft path. Neither is "the version we ship." Audit Found.D2.

**Files:**
- Create: `docs/release-smoke-checklist.md` (7-step pre-merge gate)
- Create: `docs/release-runs/v1.0.0-rc.1.md` (first run log)
- Modify: `CLAUDE.md` lines 7, 48, 115, 492-493 (drop SPFx 1.21 / gulp claims — narrow sub-edit; broader docs sweep is Tranche 2 D5)

**Branch choreography (read first):**

This plan starts on `feat/spfx-1.22-heft-migration` (current branch per `git status`). Tasks 1-3 (D1, D13, D7) commit onto that branch. Task 4 then continues on the same branch:

1. Steps 1-5 here write the smoke checklist, run it, write the run log, edit CLAUDE.md narrowly — all committed onto `feat/spfx-1.22-heft-migration` BEFORE the squash so they ride into the squash commit.
2. Step 6 then switches to `main` and runs `git merge --squash feat/spfx-1.22-heft-migration`, collapsing the entire feature branch (91 prior migration commits + Tasks 1-3 commits + Task 4 Steps 1-5 commits + the audit doc commits already on the branch) into one squash on main.
3. Step 7 tags the squash as `v1.0.0-rc.1`.
4. Task 5 (D12) and all of Tranche 2 begin on a fresh branch off the new `main` HEAD — NOT on the now-merged feature branch.

This avoids the bug where the run-log "PR is blocked from landing until that file is committed" acceptance signal in audit Found.D2 cannot be met if the run-log lands AFTER the squash.

- [ ] **Step 1: Write the smoke checklist**

Create `docs/release-smoke-checklist.md`:

```markdown
# SP Search — Release Smoke Checklist

> Owned by Foundations track (Found.D2). Run end-to-end before any merge from `feat/spfx-1.22-heft-migration` (or any future feature branch) into `main` and before tagging a release. Each run produces a log under `docs/release-runs/<tag>.md`. Skip allowed only with documented Foundations-track ticket.

## Pre-flight

- Clean machine (or `git clean -xdf && rm -rf node_modules`)
- Node 22.14.x per `package.json` engines
- `git checkout <branch-being-released>`
- `git status --short` returns empty

## Steps

1. **`npm install`** — completes without errors; `node_modules/` populated; no peer-dep warnings escalated to errors.
2. **`npm run type-check`** — exits 0; same clean result as `npx tsc --noEmit -p tsconfig.json` (gated on Found.D3).
3. **`npm test`** — exits 0; reports ≥1 spec passed (gated on Found.D13). At minimum `tests/store/lifecycle.test.ts` runs.
4. **`npm run package`** — exits 0; `sharepoint/solution/sp-search.sppkg` exists with current timestamp (gated on Found.D1).
5. **`npm run check:bundles`** — exits 0; all 6 web parts within budget (gated on Found.D7).
6. **Tenant upload smoke** — upload `sp-search.sppkg` to the test-tenant app catalog (`https://pixelboy.sharepoint.com/sites/SPSearch`). Add each of the 6 web parts (Box, Results, Filters, Verticals, Manager, AdminManager) to a page. Verify zero console errors; basic search query returns ≥1 result; `?debug=1` opens DebugFab.
7. **Multi-context smoke** — provision a multi-context page via `Provision-TestPages.ps1`; verify two independent search contexts maintain separate filter state and URL params.

## Result-log template

Each step records: `PASS | FAIL | SKIP <reason>` + evidence link (commit SHA, screenshot path, or log excerpt). File location: `docs/release-runs/<tag>.md`.

## Re-run policy

If any step FAILS: do not merge. Open a track-tagged ticket, fix on the feature branch, re-run from Step 1. SKIP only with explicit reason and a follow-up ticket cited.
```

- [ ] **Step 2: Verify Tranche-1 prereqs are landed**

Run: `git log --oneline | head -5`
Expected: commits for Found.D1, Found.D13, Found.D7 visible (the prior tasks). If not, do not proceed — a release without the lint fix, test gate, or bundle gate cannot be smoke-checked.

- [ ] **Step 3: Drop the narrow CLAUDE.md SPFx-1.21 / gulp claims**

Edit `CLAUDE.md`:
- Line 7: `**SPFx 1.21.1 solution**` → `**SPFx 1.22.2 solution**`
- Line 48 (Tech stack table row): `| SharePoint Framework | 1.21.1 | SPFx web part platform |` → `| SharePoint Framework | 1.22.2 | SPFx web part platform |`
- Line 115 (Architecture rule 10): `Run \`gulp clean-cache\` (or \`rm -rf node_modules/.cache/webpack\`) whenever \`@pnp/*\` or other dependency packages are updated.` → `Run \`npm run clean:cache\` (which invokes \`rimraf node_modules/.cache\`) whenever \`@pnp/*\` or other dependency packages are updated.`
- Lines 492-493 (Common Commands block): already corrected by Found.D13 Step 6 — verify state and adjust if a re-edit is needed.

(Broader README + admin-guide + provisioning-guide + Toast + presets sweep is Tranche 2 D5.)

- [ ] **Step 4: Run the smoke checklist end-to-end on the current head**

Run each step listed in `docs/release-smoke-checklist.md`. Capture output and outcomes.

**First-run workaround for Step 2 (`npm run type-check`):** Found.D3 is in Tranche 2 of this plan and has not yet landed. For the v1.0.0-rc.1 smoke run, substitute `npx tsc --noEmit -p tsconfig.json` and mark Step 2 as `SKIP — using workaround per Found.D3 deferral; landed in Tranche 2` in the run log. Subsequent runs after D3 lands use `npm run type-check` directly.

For Step 6 (tenant upload smoke), use:
```bash
npm run package
# upload sharepoint/solution/sp-search.sppkg to https://pixelboy.sharepoint.com/sites/appcatalog manually via SP admin center
# add each web part to https://pixelboy.sharepoint.com/sites/SPSearch (or test page)
# capture screenshots / console-log evidence into docs/release-runs/v1.0.0-rc.1.md as you go
```

- [ ] **Step 5: Write the run log**

Create `docs/release-runs/v1.0.0-rc.1.md`:

```markdown
# Release run — v1.0.0-rc.1

**Tag commit:** <fill in after squash-merge>
**Branch:** feat/spfx-1.22-heft-migration → main
**Run date:** 2026-MM-DD
**Operator:** <git user>

## Pre-flight

- [x] Clean machine
- [x] Node 22.14.x verified (`node --version`)
- [x] `git status --short` empty

## Steps

| # | Step | Outcome | Evidence |
|---|------|---------|----------|
| 1 | `npm install` | PASS | Completed in <X>s, no error-level peer-dep warnings |
| 2 | `npm run type-check` | PASS | Exit 0; <X> files type-checked clean |
| 3 | `npm test` | PASS | <X> specs passed; tests/store/lifecycle.test.ts ran |
| 4 | `npm run package` | PASS | sharepoint/solution/sp-search.sppkg <bytes>; SHA-256 <hash> |
| 5 | `npm run check:bundles` | PASS | All 6 web parts within budget per release/analysis-logs/bundle-sizes.json |
| 6 | Tenant upload smoke | PASS | Uploaded to https://pixelboy.sharepoint.com/sites/appcatalog; 6 web parts added to test page; no console errors; basic query returned <N> results; ?debug=1 opened DebugFab |
| 7 | Multi-context smoke | PASS | Provision-TestPages.ps1 ran; ctx1 + ctx2 maintained independent state |

## Notes

- First run after Tranche 1 D1+D13+D7 landed; D2 squash commit follows.
- D8 CI automation lands in Tranche 2; this run was performed manually.
```

- [ ] **Step 5b: Commit the smoke checklist + run log + CLAUDE.md narrow edits onto the feature branch BEFORE the squash**

```bash
git add docs/release-smoke-checklist.md docs/release-runs/v1.0.0-rc.1.md CLAUDE.md
git commit -m "$(cat <<'EOF'
docs(release): smoke checklist + v1.0.0-rc.1 run log + CLAUDE.md 1.21/gulp narrow sweep (Found.D2 pre-squash)

Lands on feat/spfx-1.22-heft-migration BEFORE the squash so the run
log + checklist + narrow CLAUDE.md edits ride into the single squash
commit (per audit Found.D2 acceptance signal: "merge PR is blocked
from landing until that file is committed").

Smoke checklist 7 steps PASS at this commit (Step 2 SKIP with
documented Found.D3 deferral workaround); evidence inline in the
run log file.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

- [ ] **Step 6: Squash-merge the feature branch into main**

Run from a clean working tree on a fresh checkout of `main`:

```bash
git checkout main
git pull --ff-only origin main
git merge --squash feat/spfx-1.22-heft-migration
git commit -m "$(cat <<'EOF'
feat(spfx-1.22): merge SPFx 1.22 / Heft migration (Found.D2)

Squashes the 91-commit feat/spfx-1.22-heft-migration branch into a
single main commit per the Foundations Found.D2 merge gate. Migration
phases consolidated below; full per-commit history preserved on the
feature branch reference for audit.

Phases (by reference commit):
- c6dd4f1 Heft config files + webpack patch
- a5f28c1 Gulp -> Heft build system migration
- 77adef7 SPFx 1.22.2 / spfx-toolkit type mismatch fix
- b4186e8 Heft build errors (SASS, ESLint, source-map-loader)
- bab3e9e PnP property controls CSS fix (paired with Found.D1)
- (plus all subsequent T1-T5 + Sprint-3 feature work — see audit Appendix E commit log)

Pre-merge gate: docs/release-smoke-checklist.md 7 steps PASS;
run-log committed at docs/release-runs/v1.0.0-rc.1.md.

CLAUDE.md lines 7/48/115 updated to drop SPFx 1.21 / gulp claims.
Broader docs sweep (README, admin-guide, provisioning-guide, Toast,
shipped-as-backlog presets) deferred to Found.D5 (Tranche 2).

Closes Found.D2 P0 (audit Part 3 + Part 4 + Journey A Step 12 [Polish]).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

- [ ] **Step 7: Tag the merge commit and create release entry (manual)**

Run:
```bash
git tag -a v1.0.0-rc.1 -m "Release candidate 1 — first canonical install artifact"
git push origin main
git push origin v1.0.0-rc.1
```

Then manually via GitHub UI (D8 wires automation in Tranche 2):
- Navigate to Releases → "Draft a new release"
- Tag: `v1.0.0-rc.1`
- Title: `v1.0.0-rc.1 — SPFx 1.22 / Heft migration`
- Body: paste the squash commit message
- Attach `sharepoint/solution/sp-search.sppkg` from the local build
- Publish release

- [ ] **Step 8: Verify the release artifact is reachable**

Verify via the browser or `gh release view v1.0.0-rc.1` (if `gh` CLI available):
- Tag exists: `git tag --list 'v*'` returns `v1.0.0-rc.1`
- `git log main --oneline | head -3` shows the squashed merge at the top
- Release page lists `sp-search.sppkg` as a downloadable asset
- `grep -n "SPFx 1\.21\|gulp " CLAUDE.md` returns 0 hits in lines 1-499 (the broader sweep is Tranche 2 D5; line 488-501 already updated by D13)

- [ ] **Step 9: Verify the squash commit includes the pre-squash files**

```bash
git show --stat HEAD | grep -E "release-(smoke|runs)|CLAUDE\.md|pnpPropertyControlsFix|jest-shared-config|bundle-budgets|check-bundle-sizes" | head
```

Expected: each file listed appears in the squash commit's file-change list. If any file is missing, the pre-squash commit (Step 5b) was skipped — restore from the feature branch and re-squash.

(No additional commit on main beyond the squash + the manual Releases publish + the tag push; everything Tranche-1 lives in the single squash commit.)

---

## Task 5: Found.D12 — Bundle reduction sweep — **SUPERSEDED by Task 3.5; reclassified to Defer**

> **Status:** SUPERSEDED. Do not execute.
>
> **Reclassification:** Task 3.5 (Found.D7 amendment, commit `63b4cc1`) re-anchored the bundle baseline against actual `heft build --production` output and discovered that the original "install-blocking, 4–7× over SPFx guidance" framing was based on cached non-production bundle sizes (~7.96–14.75 MB unhashed) rather than the actual minified hashed production output (470K–862K). All 6 production bundles ship at 22–43% of SPFx informal 1–2 MB guidance. With Task 3.5's hash-aware gate enforcing the corrected production budget, the active reduction sweep originally scoped here is **not needed at v1.0**.
>
> **New status: Defer.** Re-evaluate at v1.1+ if a sub-1 MB ceiling or sub-200K aspirational target becomes a goal. The audit doc was patched in lockstep — see `docs/sp-search-launch-readiness-audit.md` Part 3 Foundations Current state amendment note + Part 4 Roadmap Matrix Found.D7/D12 rows.
>
> **Audit-trail preservation:** The original task body is retained below verbatim so future contributors can understand why D12 was originally P0 XL and what the reduction sweep would have looked like if executed. Do NOT execute these steps — Task 3.5's hash-aware gate already enforces production budgets.

---

### Original Task 5 body (SUPERSEDED, retained for audit trail)

**Why (P0 rule: (d) prevents install/build):** D7 ships only the budget definition + CI gate; without active reduction, the install-blocking sizes freeze in place. A relaxed 1.5× ceiling is still ~6× over SPFx informal guidance for the smallest bundle and ~10× for the largest. Constrained-network tenants block on first-paint at the current sizes; only an active sweep brings them toward the 1–2 MB guidance. Audit Found.D12.

**Files:**
- Modify: per-web-part SCSS / TSX import paths across `src/webparts/{spSearchFilters, spSearchResults, spSearchManager, spSearchAdminManager, spSearchBox, spSearchVerticals}/components/*`
- Create: `docs/performance/bundle-analyzer/{baseline,post-sweep}/<webpart>.html` × 6 (per-web-part `webpack-bundle-analyzer` HTML — committed under tracked `docs/` because `release/` is gitignored)
- Create: `docs/performance/bundle-analyzer/{baseline,post-sweep}/top10.md` × 2
- Modify: `config/bundle-budgets.json` if any web part lands meaningfully under the 1.5× ceiling (tighten budget to lock in the gain)
- Modify: `docs/performance/bundle-sizes-baseline.json` (created in D7 Step 8 — refresh if any web part shrinks materially)

- [ ] **Step 1: Generate per-web-part `webpack-bundle-analyzer` HTML for the baseline**

Run: `npm run stats 2>&1 | tail -20`
Expected: completes a `heft build --clean --production` pass with `ANALYZE=1` env set; produces per-bundle analysis output.

Capture the baseline HTML for each of the 6 web parts:
```bash
mkdir -p docs/performance/bundle-analyzer/baseline
# webpack-bundle-analyzer's stats output lives under temp/webpack/ or release/ depending on Heft config — locate and copy:
find . -name "*.html" -path "*bundle-analyzer*" -newer package.json 2>/dev/null | xargs -I {} cp {} docs/performance/bundle-analyzer/baseline/
ls docs/performance/bundle-analyzer/baseline/
```

If `npm run stats` does not emit per-web-part HTML, fall back to direct invocation:
```bash
npx webpack-bundle-analyzer release/analysis-logs/<stats-json> -m static -r docs/performance/bundle-analyzer/baseline/sp-search-filters-web-part.html
# repeat for each of the 6 stats files
```

Commit the baseline HTML files so the post-sweep diff is visible in the PR.

- [ ] **Step 2: Identify top-10 contributors per web part**

For each of the 6 baseline HTML files, open in a browser and write down the top-10 modules by size. Capture in a temporary scratch file `docs/performance/bundle-analyzer/baseline/top10.md`:

```markdown
| Web part | Top-10 contributors (by gzipped size) |
|----------|----------------------------------------|
| filters | <list> |
| results | <list> |
| ...    | <list> |
```

This drives the sweep targets. Common offenders to expect: `devextreme/dist/css/dx.material.blue.light.compact.css`, `@fluentui/react/lib/Dialog`, `@fluentui/react/lib/Persona`, full `devextreme-react/{tree-view,filter-builder,form,pivot}` imports, full `xlsx` package, full `spfx-toolkit/lib/components/{VersionHistory,FormContainer,UserPersona}`.

- [ ] **Step 3: Audit DevExtreme imports — push heavy components to lazy**

For each web part's `components/` directory, grep for `devextreme-react/` direct imports:
```bash
grep -rn "from 'devextreme-react/" src/webparts/
```

Per CLAUDE.md import rules, keep `TagBox` and `DateRangeBox` direct (lightweight). Convert these to lazy via `createLazyComponent` from `spfx-toolkit/lib/utilities/lazyLoader`:
- `devextreme-react/tree-view`
- `devextreme-react/filter-builder`
- `devextreme-react/form`
- `devextreme-react/pivot`
- Any other non-Tag/DateRange direct import

Pattern:
```typescript
import { createLazyComponent } from 'spfx-toolkit/lib/utilities/lazyLoader';
const TreeView: any = createLazyComponent(() => import('devextreme-react/tree-view') as any, {
  errorMessage: 'Failed to load TreeView component'
});
```

Per CLAUDE.md memory: cast `as any` is required due to the @types/react mismatch between sp-search and spfx-toolkit node_modules. Do not wrap with `<React.Suspense>` — `createLazyComponent` bundles Suspense + error boundary.

- [ ] **Step 4: Audit Fluent UI imports — verify all use tree-shakable lib paths**

```bash
grep -rn "from '@fluentui/react'" src/webparts/ src/libraries/
```

Expected: 0 hits. Any hit must convert to `from '@fluentui/react/lib/<component>'` per CLAUDE.md import rules. Fix each.

- [ ] **Step 5: Audit spfx-toolkit imports — defer below-the-fold components**

```bash
grep -rn "from 'spfx-toolkit/lib/components/" src/webparts/
```

Per audit body D12 sub-item (c): convert components that load on first paint but only render on user action to `createLazyComponent`. Candidates:
- `VersionHistory` (only opens in detail panel) → already lazy via SearchManager? verify.
- `FormContainer` / `FormItem` (only renders in detail panel) → push to lazy.
- `WorkflowStepper` (detail panel only) → push to lazy.

Keep eager:
- `Card`, `Header`, `Content` (used in main render path)
- `ErrorBoundary` (root-level wrapping)
- `DocumentLink` (used in every layout cell renderer)

- [ ] **Step 6: Drop `xlsx` from the eager bundle if not already lazy**

```bash
grep -rn "from 'xlsx'" src/
```

Per memory note: XLSX export already lazy-loaded via DataGridContent.tsx (Sprint 3). Verify that the `xlsx` package only appears in a lazy chunk (`chunk.xlsx_*.js`) and not in any web-part eager bundle. If it leaks, convert the import site to `await import('xlsx')` inside the export handler.

- [ ] **Step 7: Rebuild and re-measure**

Run: `npm run package 2>&1 | tail -10`
Run: `npm run check:bundles 2>&1`

Expected: gate passes; per-web-part actual sizes lower than baseline. Capture the new sizes in `docs/performance/bundle-analyzer/post-sweep/top10.md` and the per-web-part HTML in `docs/performance/bundle-analyzer/post-sweep/`.

- [ ] **Step 8: Verify acceptance criteria**

Per audit Found.D12 acceptance signal:
- `webpack-bundle-analyzer` HTML output checked into PR for each of the 6 web parts (baseline + post-sweep)
- Per-web-part sizes within 1.5× ceilings (already enforced by D7 gate)
- At least 3 of 6 web parts show ≥10% reduction from pre-sweep baseline

Compute deltas:
```bash
node -e "
const baseline = {
  'sp-search-filters-web-part.js': 14752081,
  'sp-search-results-web-part.js': 14503213,
  'sp-search-admin-manager-web-part.js': 11123315,
  'sp-search-manager-web-part.js': 11116954,
  'sp-search-box-web-part.js': 8488976,
  'sp-search-verticals-web-part.js': 7956603
};
const fs = require('fs');
let qualifying = 0;
for (const [name, base] of Object.entries(baseline)) {
  const actual = fs.statSync('release/assets/' + name).size;
  const reduction = ((base - actual) / base) * 100;
  console.log(name + ': ' + reduction.toFixed(1) + '% reduction');
  if (reduction >= 10) qualifying++;
}
console.log('Web parts at >=10% reduction: ' + qualifying);
"
```

Expected: at least 3 of 6 ≥10%. If fewer, iterate Steps 3-6 on the largest non-qualifying bundles.

- [ ] **Step 9: Update budgets to lock in gains (optional but recommended)**

If any web part lands under (current × 1.2), tighten its budget in `config/bundle-budgets.json` to 1.2× the new actual size — locks in the win and prevents silent regression. Example: if Filters drops from 14.75 MB to 12 MB, tighten budget from 22.13 MB to 14.4 MB.

Also update `docs/performance-budgets.md` to reflect post-sweep current sizes.

- [ ] **Step 10: Commit**

```bash
git add src/webparts/ src/libraries/ config/bundle-budgets.json docs/performance-budgets.md docs/performance/
git commit -m "$(cat <<'EOF'
perf(bundles): tree-shake + lazy-load sweep across 6 web parts (Found.D12)

Pushes heavy DevExtreme components (TreeView, FilterBuilder, Form,
Pivot) to createLazyComponent; defers spfx-toolkit components that
only render on user action (VersionHistory, FormContainer,
WorkflowStepper) into lazy chunks; verifies @fluentui/react imports
all use tree-shakable lib paths.

Per-web-part webpack-bundle-analyzer HTML committed at
docs/performance/bundle-analyzer/{baseline,post-sweep}/ (release/* is
gitignored, so the analyzer output lives under tracked docs/). At least
3 of 6 web parts achieve >=10% reduction from baseline. Budgets in
config/bundle-budgets.json tightened to lock in gains; updated table
in docs/performance-budgets.md.

Sweep stays under Found.D7 1.5x ceiling; aspirational 50% targets are
v1.1+ work.

Closes Found.D12 P0 (audit Part 3 + Part 4 + spec §4.4).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

# Tranche 2 — P1/P2 polish + Sprint 5/6 spillover

**Branching:** Tranche 1 ended with the `feat/spfx-1.22-heft-migration` branch squash-merged into `main` and tagged `v1.0.0-rc.1`. Tranche 2 work happens on a fresh branch off `main` (e.g. `feat/foundations-tranche-2`) — NOT on the now-merged feature branch. Each task can ship as its own PR or bundled per Sprint per-track-lead discretion.

Order rationale: D3 unlocks deployment-guide accuracy and D2's smoke checklist Step 2. D5 closes the broader docs sweep that D2 narrow-edited. D8 retroactively wires CI for the manual D2 release. D6 consumes D8's CI workflow for axe-core gate. D4 ships security hardening alongside T2.D3. D10 declares the verified Microsoft Graph scope so Day-1 Graph People vertical works. D9 ships telemetry plumbing that T5.D8/T5.D9 consume. D11 closes metadata cleanup that rides along with D2's release tag (already shipped manually) + D8's automation.

## Task 6: Found.D3 — Fix `npm run type-check` script (`heft build --clean --lite` → working invocation)

**Why (P1):** Audit Found.D3. The deployment guide tells admins to run `npm run type-check` before packaging; today the script invokes `heft build --clean --lite` and `--lite` is not recognized by Heft 1.2.7. Workaround `npx tsc --noEmit` is one command away, hence P1 not P0. Bundled into Tranche 2 ahead of D5 because D5 updates the deployment guide and needs the corrected command name.

**Files:**
- Modify: `package.json:28`
- Modify: `docs/deployment-guide.md:18-22`

- [ ] **Step 1: Update the script**

Edit `package.json:28`:

```json
"type-check": "tsc --noEmit -p tsconfig.json",
```

Per audit body shape (a) — bypass Heft entirely; self-contained semantics.

- [ ] **Step 2: Verify the script runs clean**

Run: `npm run type-check 2>&1 | tail -10`
Expected: exits 0 with no output (TypeScript clean per Phase 0.2 fallback verification). Same result as `npx tsc --noEmit -p tsconfig.json`.

- [ ] **Step 3: Update the deployment guide**

Edit `docs/deployment-guide.md:18-22` — replace the `gulp` block with:

```bash
npm run type-check
npm test
npm run package
```

(Drop `gulp bundle --ship` / `gulp package-solution --ship` — `npm run package` is the canonical single command per `package.json:21`.)

- [ ] **Step 4: Commit**

```bash
git add package.json docs/deployment-guide.md
git commit -m "$(cat <<'EOF'
build(scripts): replace heft --lite with tsc --noEmit for type-check (Found.D3)

heft build --clean --lite was never a recognized Heft 1.2.x flag and
silently failed or no-oped. Replaces with direct tsc invocation that
self-contains the type-check semantics. Updates docs/deployment-guide.md
lines 18-22 to drop gulp commands removed in the SPFx 1.22 migration.

Closes Found.D3 P1 (audit Part 3 + Appendix E §4 row 3).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 7: Found.D5 — Stale docs sweep (README + CLAUDE.md + admin-guide + provisioning-guide + Toast/preset claims)

**Why (P0 rule: (e) Journey A Step 12 [Confusion] + Step 1 [Blocker] secondary):** No top-level `README.md` exists. `CLAUDE.md` still claims `Toast/ToastProvider` adopted (0 hits in src) and lists `knowledgeBase`/`hubSearch`/`policySearch` presets as Sprint 4 backlog despite shipping. `docs/admin-guide.md:221-226` lists Manager defaults that disagree with manifests. `docs/provisioning-guide.md:131-132` claims "no automatic background cleanup" while `SearchManagerService.ts:735-739` runs a 24h sweep. Audit Found.D5. Note: P0 per audit but deferred to Tranche 2 per user direction; CLAUDE.md gulp/1.21 lines 7/48/115/492-493 already corrected by Tasks 2 + 4.

**Files:**
- Create: `README.md`
- Modify: `CLAUDE.md` lines 305-309, 433-439, 458, 464
- Modify: `docs/admin-guide.md:221-226`
- Modify: `docs/provisioning-guide.md:131-132`
- Modify: `~/.claude/projects/-Users-hemantmane-Development-sp-search/memory/MEMORY.md` (memory file Sprint 4 backlog drift per audit D5 (e) — ask user before writing)

- [ ] **Step 1: Decide Manager defaults direction (manifest vs doc)**

Per audit D5 sub-item (c): either flip the doc to match the shipped `false` defaults OR coordinate with T4.D6 to flip the manifest to `true`. T4.D6 is P1, not in Tranche 1 of this plan. Default decision: **flip the doc to match the shipped manifest** (lower risk, no behavior change). If the user wants to flip the manifest instead, mark a follow-up ticket and skip Step 4.

- [ ] **Step 2: Write the top-level README**

Create `README.md`:

```markdown
# SP Search

Enterprise SharePoint search solution built as SPFx 1.22 web parts. Replaces PnP Modern Search v4 with a modern React + Zustand architecture, scenario presets, advanced layouts, and an extensible provider/registry model.

## What's in the box

| Web part | Purpose |
|----------|---------|
| **SP Search Box** | Query input with KQL mode, suggestions, scope selector |
| **SP Search Results** | Results display with 6 layouts (DataGrid, Card, List, Compact, People, Gallery), detail panel, bulk actions |
| **SP Search Filters** | Refinement filters (Checkbox, DateRange, Slider, PeoplePicker, TaxonomyTree, TagBox, Toggle) |
| **SP Search Verticals** | Tab navigation with badge counts |
| **SP Search Manager** | Saved searches, sharing, collections, history |
| **SP Search Admin Manager** | Tenant-wide admin dashboard |

All six web parts share a single Zustand store via an SPFx Library Component (`sp-search-store`). Multi-instance isolation via `searchContextId`.

## Install

```bash
git clone <repo>
cd sp-search
npm install
npm run package
# upload sharepoint/solution/sp-search.sppkg to your tenant or site app catalog
```

The latest released `.sppkg` is also attached to each [GitHub Release](../../releases) so you can skip the local build.

## Documentation

- **Admin install + configuration:** [`docs/admin-guide.md`](docs/admin-guide.md)
- **Deployment / packaging:** [`docs/deployment-guide.md`](docs/deployment-guide.md)
- **Provisioning scripts:** [`docs/provisioning-guide.md`](docs/provisioning-guide.md)
- **Extensibility (custom providers / actions / layouts):** [`docs/extensibility-guide.md`](docs/extensibility-guide.md)
- **PnP Modern Search v4 alignment:** [`docs/pnp-modern-search-alignment.md`](docs/pnp-modern-search-alignment.md)
- **End-user guide:** [`docs/end-user-guide.md`](docs/end-user-guide.md) (T2.D12, ships in v1.0)
- **Privacy notice:** [`docs/privacy-notice.md`](docs/privacy-notice.md)
- **Accessibility statement:** [`docs/accessibility.md`](docs/accessibility.md)
- **Performance budgets:** [`docs/performance-budgets.md`](docs/performance-budgets.md)
- **Release policy:** [`docs/release-policy.md`](docs/release-policy.md)
- **Release smoke checklist:** [`docs/release-smoke-checklist.md`](docs/release-smoke-checklist.md)
- **Changelog:** [`CHANGELOG.md`](CHANGELOG.md)
- **Contributing:** [`CONTRIBUTING.md`](CONTRIBUTING.md)
- **Architecture / developer reference:** [`CLAUDE.md`](CLAUDE.md)

## Tech stack

SPFx 1.22.2 · React 17.0.1 · TypeScript 5.3.3 · Zustand 4.x · PnPjs 3.x · Microsoft Graph Client · Fluent UI v8.106 · DevExtreme 22.2 · spfx-toolkit (Card, ErrorBoundary, hooks, utilities)

## License

See `LICENSE` (if present) or contact the project owner.
```

- [ ] **Step 3: Sweep CLAUDE.md stale claims**

Edit `CLAUDE.md`:
- Lines 433-439 (Phase 5: Sprint 4 Backlog) — drop the four shipped items per audit D5 sub-item (f). New text:

```markdown
### Phase 5: Sprint 4 Backlog
- Implement `queryInputTransformation` in `SearchOrchestrator` (currently surfaced in props but not applied) — MISS-001
- Implement `operatorBetweenFilters` in filter execution path or remove from property pane — MISS-002
- Admin-time property validation in edit mode — T4.D5
```

(Drops: Jest harness fix → shipped Found.D13; XLSX export → shipped Sprint 3; Knowledge Base / Hub Search / Policy Search presets → shipped Sprint 3 per `searchPresets.ts:64-384`.)

- Line 458 (Components used table — `Toast / ToastProvider`): delete the row entirely. `grep -rn 'spfx-toolkit/lib/components/Toast' src` returns 0 hits; T5 explicitly puts toolkit Toast Out-of-Scope.
- Line 464 (same row in another section if present): delete.

- [ ] **Step 4: Fix admin-guide Manager defaults table**

Edit `docs/admin-guide.md:221-226` — change all four `true` → `false` (matching `SpSearchManagerWebPart.manifest.json:29-32`):

```markdown
| `enableSavedSearches` | `false` | Saved searches tab |
| `enableSharedSearches` | `false` | Shared searches tab |
| `enableCollections` | `false` | Collections tab |
| `enableHistory` | `false` | History tab |
```

- [ ] **Step 5: Fix provisioning-guide retention claim**

Edit `docs/provisioning-guide.md:131-132` — replace:

```markdown
   - This is a manual operation — there is no automatic background cleanup
```

with:

```markdown
   - SearchManagerService runs an automatic 24-hour cleanup sweep retaining the last 90 days of history per `HISTORY_RETENTION_DAYS = 90` (`src/libraries/spSearchStore/services/SearchManagerService.ts:735-739`). The manual cleanup script is supplemental and can shorten the retention window for one-off purges.
```

- [ ] **Step 6: Verify the docs sweep**

Run:
```bash
test -f README.md && echo "README ok"
grep -n "SPFx 1\.21\|gulp " CLAUDE.md
grep -n "Toast" CLAUDE.md
grep -n "knowledgeBase\|hubSearch\|policySearch" CLAUDE.md
grep -n "true" docs/admin-guide.md | grep -E "enable(Saved|Shared|Collections|History)"
grep -n "no automatic background cleanup" docs/provisioning-guide.md
```

Expected: README exists; all greps return 0 hits.

- [ ] **Step 7: Coordinate the MEMORY.md sweep with the user**

Per audit D5 sub-item (f) closing line: `~/.claude/projects/-Users-hemantmane-Development-sp-search/memory/MEMORY.md` has the same shipped-but-listed-as-backlog drift. Ask the user before editing the auto-memory file (since memory files are user-controlled).

If user approves: edit the "Stale Sprint 4 backlog notes" section to reflect the post-D5 truth and remove items now corrected in CLAUDE.md.

- [ ] **Step 8: Commit (in-repo files; memory edits are a separate user-approved action)**

```bash
git add README.md CLAUDE.md docs/admin-guide.md docs/provisioning-guide.md
git commit -m "$(cat <<'EOF'
docs: stale-docs sweep — README + CLAUDE.md + admin-guide + provisioning-guide (Found.D5)

Adds top-level README orienting self-serve admins to the docs/* tree,
the Releases page, and the canonical install path. Closes 4 Sprint-4
backlog items in CLAUDE.md that shipped in Sprint 3 (Jest harness fix
landed in Found.D13; XLSX export, Knowledge Base / Hub / Policy presets
shipped per searchPresets.ts and DataGridContent.tsx). Drops the
Toast / ToastProvider row from the components-used tables (0 src hits;
T5 OOS). Flips admin-guide Manager defaults table to match the
shipped manifest values (false / false / false / false). Replaces the
provisioning-guide "no automatic background cleanup" claim with the
actual 24h sweep + 90-day retention text.

CLAUDE.md gulp/1.21 lines 7/48/115/492-493 already corrected in
Found.D2 + Found.D13 commits.

Closes Found.D5 P0 (audit Part 3 + Journey A Steps 1, 10, 12).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 8: Found.D8 — CI/release engineering — GitHub Actions + semver + CHANGELOG + release artifact

**Why (P1):** Audit Found.D8. Manual workaround "build locally, attach to a one-off email" exists, hence P1 not P0. Bundles in the same PR-shaped change every sub-item that lives under `.github/workflows/` or one root-level convention file. Tranche 2 D2 already shipped the manual `v1.0.0-rc.1` tag; this task wires the automation that future releases will use.

**Files:**
- Create: `.github/workflows/build.yml`
- Create: `.github/workflows/release.yml`
- Create: `.github/PULL_REQUEST_TEMPLATE.md`
- Create: `.github/ISSUE_TEMPLATE/bug_report.md`
- Create: `.github/ISSUE_TEMPLATE/feature_request.md`
- Create: `.github/dependabot.yml`
- Create: `CHANGELOG.md`
- Create: `CONTRIBUTING.md`
- Create: `docs/release-policy.md`
- Modify: `package.json:3` `version` from `0.0.1` → `1.0.0` (D11 lockstep — but the version bump rides with D11 in Task 13 if T4.D10 has not landed yet)

- [ ] **Step 1: Write `docs/release-policy.md` (semver convention)**

Create `docs/release-policy.md`:

```markdown
# SP Search — Release Policy

> Owned by Foundations track (Found.D8). Authoritative source for version-bump decisions, lockstep convention between `package.json` and `config/package-solution.json`, and tag/release naming.

## SemVer 2.0

- **Major (`x.0.0`)** — breaking changes to web-part property pane fields, store API surface, ISearchDataProvider contract, IFilterConfig schema, or registries (DataProvider, Suggestion, Action, Layout, FilterType). Anything that requires admins to re-author saved searches or property-pane configurations counts.
- **Minor (`1.x.0`)** — new features that do not break existing configurations: new layouts, new filter types, new scenario presets, new actions, new admin dashboard tabs.
- **Patch (`1.0.x`)** — bug fixes, performance improvements, documentation updates, dependency bumps that do not change the public surface.

Pre-release tags: `1.0.0-rc.N` for release candidates, `1.0.0-beta.N` for unstable previews. Promote rc.N → release by tagging the rc.N commit as `1.0.0` (no code changes between).

## Lockstep convention

`package.json:3` `version` and `config/package-solution.json:6` `solution.version` move in lockstep at every release. The `solution.version` field uses 4-part SharePoint-required `M.M.P.B` (build number always `0` unless an in-tag rebuild is required); `package.json` uses 3-part SemVer.

| package.json | solution.version |
|--------------|-------------------|
| `1.0.0`      | `1.0.0.0` |
| `1.0.1`      | `1.0.1.0` |
| `1.1.0`      | `1.1.0.0` |
| `2.0.0`      | `2.0.0.0` |

Both files MUST be bumped in the same commit. CI at `.github/workflows/build.yml` validates the lockstep relationship and fails on mismatch.

## Tag + Release naming

- Tag format: `v<semver>` (e.g. `v1.0.0`, `v1.0.0-rc.1`)
- Release title: `v<semver> — <one-line summary>`
- Release body: pulled from the matching `## [<semver>]` section of `CHANGELOG.md`
- Release artifact: `sp-search.sppkg` from the `release.yml` workflow build

## Dependency policy

- Production deps (`dependencies` in `package.json`) — manual review per PR; no automatic version bumps.
- Dev deps (`devDependencies`) — Dependabot-managed weekly per `.github/dependabot.yml`; auto-merge disabled (manual review).
- SPFx core (`@microsoft/sp-*`) — manual; coordinate version bumps with a documented Heft / spfx-toolkit compatibility check.
```

- [ ] **Step 2: Write `CHANGELOG.md`**

Create `CHANGELOG.md` with the historical entries pre-populated from Appendix A closures. Use Keep-a-Changelog format:

```markdown
# Changelog

All notable changes to SP Search are documented here. Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versioning follows [SemVer 2.0](https://semver.org/).

## [Unreleased]

## [1.0.0-rc.1] - 2026-05-DD

### Added

- SPFx 1.22 / Heft build pipeline (Foundations Found.D2 — squash-merge of 91-commit feat/spfx-1.22-heft-migration branch).
- Per-web-part bundle size budgets and CI breach gate (`config/bundle-budgets.json`, `scripts/check-bundle-sizes.js` — Found.D7).
- Heft Jest harness via `@rushstack/heft-jest-plugin` shared config; `tests/store/lifecycle.test.ts` smoke trail-marker (Found.D13).
- Top-level `README.md`, `CHANGELOG.md`, `CONTRIBUTING.md`, `docs/release-policy.md`, `docs/release-smoke-checklist.md` (Found.D2/D5/D8).
- Scenario presets for `general`, `documents`, `news`, `people`, `media`, `custom`, `knowledgeBase`, `hubSearch`, `policySearch` (Sprint 3 — `searchPresets.ts:64-384`).
- DataGrid layout with admin-configured columns, cell renderers, filter row, column chooser, virtual scrolling, CSV + XLSX export, localStorage column preferences (Sprint 3).
- Graph-backed People vertical via `GraphSearchProvider` with presence batch (Sprint 3).
- Analytics feedback loop: Health tab (zero-result queries) + Insights tab (top queries / CTR / daily volume) (Sprint 3).

### Changed

- Build pipeline migrated from gulp to Heft (`a5f28c1`); SPFx 1.21.1 → 1.22.2; spfx-toolkit type alignment (`77adef7`).
- `package.json:type-check` script now invokes `tsc --noEmit -p tsconfig.json` directly (Found.D3).
- Gallery layout collapses to single-column at 399px viewport (Sprint 3 mobile hardening).
- Admin Manager toggles (enableSavedSearches/Shared/Collections/History) ship `false` by default per `SpSearchManagerWebPart.manifest.json` (admin must opt in per tab); admin-guide updated to match (Found.D5).

### Fixed

- BUG-001..BUG-012 closures from the 2026-03-22 audit reconciliation pass (see `docs/sp-search-launch-readiness-audit.md` Appendix A).
- BUG-004 (XSS via `newPageUrl`): closed via `https?://` / `/` allowlist on `SpSearchBox.tsx:358`; remaining 7 unhardened sites consolidated into `safeNavigate` helper (Found.D4 follow-up).
- `pnpPropertyControlsFix.ts` ESLint `no-use-before-define` blocker that halted `npm run package` (Found.D1).
- `SearchHistory` Author-first CAML predicate to prevent threshold throttling on >5,000-item lists.
- PnPjs caching `QuotaExceededError` handled via inline retry + outer catch.

### Security

- SEC-003 (collection name length validation) closed.
- SEC-005 (Teams URL sovereign-cloud handling) closed.
- A11Y-001/002/003/006 (KQL ARIA + gallery aria-label + suggestion keyboard shortcut) closed.
```

- [ ] **Step 3: Write `CONTRIBUTING.md`**

Create `CONTRIBUTING.md`:

```markdown
# Contributing to SP Search

## Architecture

Read `CLAUDE.md` first — it is the authoritative source for architecture, conventions, web-part responsibilities, and import rules. The `docs/` directory has admin-facing guides; `CLAUDE.md` is developer-facing.

## Setup

```bash
git clone <repo>
cd sp-search
npm install
npm test              # Heft Jest pipeline
npm run type-check    # tsc --noEmit
npm run package       # produces sharepoint/solution/sp-search.sppkg
```

## Branching

- `main` — released code; protected
- `feat/<short-name>` — feature branches; squash-merged into `main` via PR
- `fix/<short-name>` — bug fix branches; squash-merged

## Commit messages

Use Conventional Commits prefix: `feat`, `fix`, `docs`, `test`, `build`, `perf`, `refactor`, `chore`. Example:

```
feat(filters): add SliderFilter for numeric refiners (T1.D5)

Extends FilterTypeRegistry with SliderFilter type. Wires into
SearchFilters drawer; renders devextreme-react Slider lazy-loaded.

Closes T1.D5 P1 (audit Part 2).
```

Each commit MUST close exactly one Roadmap ID from the launch readiness audit unless documented otherwise.

## Testing

- All store / service / utility code lives under `tests/{store,services,utils}/`
- Run `npm test -- --testPathPattern <pattern>` to filter
- Component tests use jest-axe for accessibility smoke (`tests/a11y/smokeAxe.test.tsx` — Found.D6)

## Pre-merge checklist

Run `docs/release-smoke-checklist.md` before merging to `main`. CI at `.github/workflows/build.yml` enforces the build + test + bundle-gate steps; the tenant-upload smoke (Step 6) and multi-context smoke (Step 7) are manual.

## Releases

See `docs/release-policy.md` for SemVer policy and tag conventions. Tagging `vX.Y.Z` triggers `.github/workflows/release.yml` which builds production + publishes a GitHub Release with `sp-search.sppkg` attached.

## Performance budgets

Per-web-part byte budgets are enforced by `scripts/check-bundle-sizes.js` (Found.D7). PRs that breach the budget fail CI; raising a budget requires Foundations track lead approval. See `docs/performance-budgets.md`.

## Accessibility

axe-core smoke tests in CI (Found.D6). PRs that introduce new violations fail CI. See `docs/accessibility.md` for the WCAG 2.1 AA scope.
```

- [ ] **Step 4: Write `.github/workflows/build.yml`**

Create `.github/workflows/build.yml`:

```yaml
name: Build

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  build:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        node-version: [22.14.x]
    steps:
      - uses: actions/checkout@v4

      - name: Setup Node.js ${{ matrix.node-version }}
        uses: actions/setup-node@v4
        with:
          node-version: ${{ matrix.node-version }}
          cache: 'npm'

      - name: Install dependencies
        run: npm ci

      - name: Type check
        run: npm run type-check

      - name: Test
        run: npm test

      - name: Package
        run: npm run package

      - name: Bundle size gate
        run: npm run check:bundles

      - name: Verify version lockstep
        run: |
          PKG_VER=$(node -p "require('./package.json').version")
          SOL_VER=$(node -p "require('./config/package-solution.json').solution.version")
          # solution.version is M.M.P.B — strip trailing .B and compare to PKG_VER
          SOL_TRIM=$(echo "$SOL_VER" | awk -F. '{print $1"."$2"."$3}')
          if [ "$PKG_VER" != "$SOL_TRIM" ]; then
            echo "Version mismatch: package.json=$PKG_VER vs solution.version=$SOL_VER ($SOL_TRIM)"
            exit 1
          fi
          echo "Versions in lockstep: $PKG_VER / $SOL_VER"

      - name: Upload .sppkg artifact
        uses: actions/upload-artifact@v4
        with:
          name: sp-search-sppkg
          path: sharepoint/solution/sp-search.sppkg
          retention-days: 90

      - name: Upload bundle attribution dashboard
        uses: actions/upload-artifact@v4
        with:
          name: bundle-sizes
          path: release/analysis-logs/bundle-sizes.json
          retention-days: 90
```

- [ ] **Step 5: Write `.github/workflows/release.yml`**

Create `.github/workflows/release.yml`:

```yaml
name: Release

on:
  push:
    tags:
      - 'v*'

jobs:
  release:
    runs-on: ubuntu-latest
    permissions:
      contents: write
    steps:
      - uses: actions/checkout@v4
        with:
          fetch-depth: 0  # full history for changelog extraction

      - name: Setup Node.js 22.14.x
        uses: actions/setup-node@v4
        with:
          node-version: 22.14.x
          cache: 'npm'

      - name: Install dependencies
        run: npm ci

      - name: Build production
        run: npm run package

      - name: Bundle size gate
        run: npm run check:bundles

      - name: Extract release notes from CHANGELOG
        id: changelog
        run: |
          TAG="${GITHUB_REF#refs/tags/v}"
          NOTES=$(awk "/^## \[$TAG\]/,/^## \[/" CHANGELOG.md | sed '$d')
          if [ -z "$NOTES" ]; then
            echo "No CHANGELOG entry for $TAG; using tag message"
            NOTES=$(git tag -n99 -l "v$TAG" | sed "s/^v$TAG[[:space:]]*//")
          fi
          {
            echo "notes<<EOF"
            echo "$NOTES"
            echo "EOF"
          } >> "$GITHUB_OUTPUT"

      - name: Create GitHub Release
        uses: softprops/action-gh-release@v2
        with:
          tag_name: ${{ github.ref_name }}
          name: ${{ github.ref_name }}
          body: ${{ steps.changelog.outputs.notes }}
          files: sharepoint/solution/sp-search.sppkg
          draft: false
          prerelease: ${{ contains(github.ref_name, '-') }}
```

- [ ] **Step 6: Write the PR template**

Create `.github/PULL_REQUEST_TEMPLATE.md`:

```markdown
## Roadmap ID

Closes: <Found.DX | T1.DX | T2.DX | T3.DX | T4.DX | T5.DX | BUG-XXX | SEC-XXX | A11Y-XXX>

## Summary

<1-3 bullet points>

## Affected web parts

- [ ] SP Search Box
- [ ] SP Search Results
- [ ] SP Search Filters
- [ ] SP Search Verticals
- [ ] SP Search Manager
- [ ] SP Search Admin Manager
- [ ] sp-search-store library component
- [ ] None (docs / config / CI only)

## Repro steps (for fixes)

1.
2.
3.

## Screenshot / recording

(if UI-affecting)

## Environment

- SPFx version: 1.22.2
- Browser tested: <Chrome / Edge / Safari / Firefox>
- Tenant tested: <test tenant URL>

## Smoke checklist (manual; CI covers Steps 1-5)

- [ ] Step 6: Tenant upload smoke (per `docs/release-smoke-checklist.md`)
- [ ] Step 7: Multi-context smoke (if multi-context surface affected)

## Bundle impact

CI reports the per-web-part delta automatically. If you raised a budget, justify here:

## Memory updates required

- [ ] None
- [ ] Updated `~/.claude/projects/.../memory/` files: <list>
```

- [ ] **Step 7: Write the issue templates**

Create `.github/ISSUE_TEMPLATE/bug_report.md`:

```markdown
---
name: Bug report
about: Report a defect in SP Search
title: '[BUG] '
labels: bug
---

## Describe the bug

## Affected web part(s)

- [ ] Box / Results / Filters / Verticals / Manager / AdminManager / store

## Reproduction steps

1.
2.
3.

## Expected behavior

## Actual behavior

## Screenshots / console log

## Environment

- SPFx version: 1.22.2
- SP Search version: <semver from CHANGELOG / package.json>
- Browser:
- Tenant URL (or anonymized):
```

Create `.github/ISSUE_TEMPLATE/feature_request.md`:

```markdown
---
name: Feature request
about: Propose a new capability for SP Search
title: '[FEAT] '
labels: enhancement
---

## Problem statement

What user / admin journey is broken or missing today?

## Proposed solution

## Alternatives considered

## Affected web part(s)

- [ ] Box / Results / Filters / Verticals / Manager / AdminManager / store

## Roadmap fit

Maps to track: <T1 / T2 / T3 / T4 / T5 / Foundations>
Estimated effort: <S / M / L / XL>
```

- [ ] **Step 8: Write `.github/dependabot.yml`**

Create `.github/dependabot.yml`:

```yaml
version: 2
updates:
  - package-ecosystem: 'npm'
    directory: '/'
    schedule:
      interval: 'weekly'
      day: 'monday'
      time: '09:00'
      timezone: 'America/Los_Angeles'
    open-pull-requests-limit: 5
    labels:
      - 'dependencies'
      - 'devDependencies-only'
    allow:
      - dependency-type: 'development'
    groups:
      heft-rushstack:
        patterns:
          - '@rushstack/*'
          - '@microsoft/spfx-heft-*'
          - '@microsoft/spfx-web-build-rig'
      jest:
        patterns:
          - 'jest'
          - 'jest-*'
          - 'ts-jest'
      types:
        patterns:
          - '@types/*'
    commit-message:
      prefix: 'chore(deps-dev)'

  - package-ecosystem: 'github-actions'
    directory: '/'
    schedule:
      interval: 'weekly'
    commit-message:
      prefix: 'chore(ci)'
```

- [ ] **Step 9: Verify the YAML / JSON files parse**

Run:
```bash
node -e "require('js-yaml').load(require('fs').readFileSync('.github/workflows/build.yml','utf8'))" 2>&1 || echo "YAML lint deferred to CI first run (js-yaml not installed)"
node -e "JSON.parse(require('fs').readFileSync('.github/dependabot.yml','utf8'))" 2>&1 || echo "expected — dependabot.yml is YAML not JSON"
```

(YAML validation will land via the first PR run on GitHub Actions; locally we accept that without a `js-yaml` devDep we cannot statically lint.)

- [ ] **Step 10: Push and verify the first build run**

Open a no-op PR (or push to `main` directly if the user accepts the risk) and confirm:
- `.github/workflows/build.yml` triggers
- All jobs (Type check, Test, Package, Bundle size gate, Version lockstep, Upload artifact) pass
- The `sp-search-sppkg` artifact appears in the PR check surface

If the workflow fails on first run, iterate based on the CI log (most likely cause: Node version mismatch in `actions/setup-node`, missing peer-dep that local `npm install` resolved silently, or the `js-yaml` lockstep step needs adjustment because `solution.version` includes the trailing `.0` build segment).

- [ ] **Step 11: Commit**

```bash
git add .github/ CHANGELOG.md CONTRIBUTING.md docs/release-policy.md
git commit -m "$(cat <<'EOF'
build(ci): GitHub Actions + semver + CHANGELOG + Dependabot + templates (Found.D8)

Stands up release engineering substrate end-to-end:
- .github/workflows/build.yml: PR + main CI (install, type-check, test,
  package, bundle-size gate, version lockstep, .sppkg artifact upload)
- .github/workflows/release.yml: tag-push (v*) builds + publishes
  GitHub Releases with .sppkg attached + auto release notes from CHANGELOG.md
- docs/release-policy.md: SemVer 2.0 + lockstep convention
- CHANGELOG.md: pre-populated v1.0.0-rc.1 entry from Appendix A closures
- CONTRIBUTING.md: short contributor guide -> CLAUDE.md + smoke + semver
- .github/PULL_REQUEST_TEMPLATE.md + ISSUE_TEMPLATE/{bug_report,feature_request}.md
- .github/dependabot.yml: weekly devDeps PRs, auto-merge disabled

Tenant catalog automation (T4.D10 -ReleaseArtifactUrl consumer) lives
on the T4 plan and consumes the GitHub Releases artifact this workflow
publishes.

Closes Found.D8 P1 (audit Part 3 + Part 4 + spec §4.4 CI/release scope).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 9: Found.D6 — WCAG 2.1 AA top-10 baseline + axe-core CI gate + accessibility.md

**Why (P1):** Audit Found.D6. Conformance statement is a launch quality lift, not a launch blocker (P1). Closes A11Y-004 + A11Y-005 (Still-Open). T1.D9 is the downstream cross-track consumer of the motion mixin + axe gate.

**Files:**
- Modify: `src/webparts/spSearchBox/components/SpSearchBox.tsx:729` (add `aria-describedby`) and `:736` (replace `<div role="radiogroup">` with `<fieldset>` + `<legend>`)
- Create: `src/styles/motion.scss` (shared mixin)
- Modify: `src/webparts/{spSearchBox/SpSearchBox.module.scss, spSearchVerticals/SpSearchVerticals.module.scss, spSearchManager/SpSearchManager.module.scss, spSearchResults/SpSearchResults.module.scss, spSearchFilters/SpSearchFilters.module.scss, debugPanel/DebugPanel.module.scss}` (wrap transitions / keyframes in motion mixin)
- Create: `tests/a11y/smokeAxe.test.tsx` (jest-axe smoke for 4 render shapes)
- Create: `docs/accessibility.md` (scoped WCAG 2.1 AA conformance statement)
- Modify: `package.json:devDependencies` to add `jest-axe`, `@types/jest-axe`, `@testing-library/react@^12.x` (React 17-compatible major), `@testing-library/jest-dom`
- Modify: `.github/workflows/build.yml` — axe smoke runs as part of `npm test` already (no separate step needed if the test file lives under `tests/`)

- [ ] **Step 1: Install jest-axe + @testing-library/react peers**

The axe smoke test in Step 6 imports `@testing-library/react` for the `render` helper. That package and its types must be installed alongside jest-axe.

Run: `npm install --save-dev jest-axe @types/jest-axe @testing-library/react @testing-library/jest-dom`

`@testing-library/react@12.x` is the React 17-compatible major (later majors require React 18). `@testing-library/jest-dom` ships the `toBeInTheDocument`, `toHaveAttribute`, etc. matchers that pair naturally with axe assertions; install for forward-compat even if Step 6 doesn't use them yet.

Verify peer deps resolve:

```bash
npm ls @testing-library/react jest-axe react react-dom
```

Expected: each lists a single resolved version under root, not multiple. If `@testing-library/react@13+` snuck in (incompatible with React 17), pin explicitly:

```bash
npm install --save-dev "@testing-library/react@^12.1.5"
```

- [ ] **Step 2: Write the motion mixin**

Create `src/styles/motion.scss`:

```scss
// Foundations Found.D6 — shared motion mixin honoring prefers-reduced-motion.
// Applied surface-by-surface by T1.D9 in Sprint 5.

@mixin motion($properties: all, $duration: 200ms, $timing: ease-in-out) {
  transition: $properties $duration $timing;

  @media (prefers-reduced-motion: reduce) {
    transition: none;
    animation: none;
  }
}

@mixin motion-keyframes($name) {
  animation-name: $name;

  @media (prefers-reduced-motion: reduce) {
    animation: none;
  }
}

// Helper for static animation property — drops to none on reduce.
@mixin motion-respect {
  @media (prefers-reduced-motion: reduce) {
    transition: none !important;
    animation: none !important;
  }
}
```

- [ ] **Step 3: Apply the mixin across the 6 module.scss files**

For each of the 6 files identified in audit Phase 7 (`grep -l "transition\|@keyframes\|animation" src/webparts/**/*.module.scss src/components/**/*.module.scss`), add at the top:

```scss
@import '../../styles/motion';
```

(Adjust relative path per file location.) Then for every `transition` / `animation` declaration, either:
- Replace the bare `transition: ...` with `@include motion(...)`, OR
- Append `@include motion-respect;` to the rule that contains the declaration (lower-touch when there are many declarations).

Verify per audit acceptance signal:

Run: `grep -rn "prefers-reduced-motion" src/styles src/webparts | wc -l`
Expected: ≥6 hits.

- [ ] **Step 4: Close A11Y-004 — Mode toggle uses `<fieldset>` + `<legend>`**

Edit `src/webparts/spSearchBox/components/SpSearchBox.tsx` around line 736:

```tsx
{/* KQL / Regular mode toggle */}
{enableKqlMode && (
  <fieldset className={styles.kqlModeToggle}>
    <legend className={styles.kqlModeToggleLegend}>Query input mode</legend>
    <button
      className={!isKqlMode ? styles.kqlModeButton + ' ' + styles.kqlModeButtonActive : styles.kqlModeButton}
      onClick={(): void => handleModeSwitch(false)}
      title="Regular search"
      aria-label="Regular search mode"
      aria-checked={!isKqlMode}
      role="radio"
      type="button"
    >
      {/* existing button contents */}
    </button>
    {/* the second mode button — preserve existing JSX */}
  </fieldset>
)}
```

Add to `SpSearchBox.module.scss`:

```scss
.kqlModeToggleLegend {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
```

(Visually-hidden legend; semantics preserved for AT users.)

- [ ] **Step 5: Close A11Y-005 — Scope selector `aria-describedby`**

Edit `src/webparts/spSearchBox/components/SpSearchBox.tsx` around line 729 — replace the existing Dropdown JSX:

```tsx
<>
  <span id="sp-search-scope-description" className={styles.visuallyHidden}>
    Restricts the search to documents within the selected SharePoint scope
  </span>
  <Dropdown
    options={scopeOptions}
    selectedKey={activeScope.id}
    onChange={handleScopeChange}
    ariaLabel="Search scope"
    aria-describedby="sp-search-scope-description"
  />
</>
```

Add `.visuallyHidden` class to `SpSearchBox.module.scss` if not already present (same shape as `.kqlModeToggleLegend` above — extract to a shared mixin if both land).

- [ ] **Step 6: Write the axe-core smoke test**

Create `tests/a11y/smokeAxe.test.tsx`:

```typescript
import * as React from 'react';
import { render } from '@testing-library/react';
import { axe, toHaveNoViolations } from 'jest-axe';

// Minimal render harness for the four most-trafficked surfaces. Real component
// integration with the SPFx context is beyond a unit test; this test verifies
// the static markup of representative shapes is axe-clean.

expect.extend(toHaveNoViolations);

describe('a11y smoke — top-10 surfaces (Found.D6)', () => {
  it('Search Box mode toggle uses semantic fieldset/legend', async () => {
    const { container } = render(
      <fieldset>
        <legend className="visuallyHidden">Query input mode</legend>
        <button role="radio" aria-checked={true} aria-label="Regular search mode">Regular</button>
        <button role="radio" aria-checked={false} aria-label="KQL mode">KQL</button>
      </fieldset>
    );
    const results = await axe(container);
    expect(results).toHaveNoViolations();
  });

  it('Scope selector exposes aria-describedby', async () => {
    const { container } = render(
      <>
        <span id="desc">Restricts the search to documents within the selected SharePoint scope</span>
        <select aria-label="Search scope" aria-describedby="desc">
          <option>All</option>
          <option>Site</option>
        </select>
      </>
    );
    const results = await axe(container);
    expect(results).toHaveNoViolations();
  });

  it('Empty state markup is axe-clean', async () => {
    const { container } = render(
      <div role="status" aria-live="polite">
        <h2>No results</h2>
        <p>Try a different query or remove a filter.</p>
      </div>
    );
    const results = await axe(container);
    expect(results).toHaveNoViolations();
  });

  it('Detail panel close button has accessible name', async () => {
    const { container } = render(
      <button aria-label="Close detail panel">×</button>
    );
    const results = await axe(container);
    expect(results).toHaveNoViolations();
  });

  it('regression sentinel — img without alt fails axe', async () => {
    const { container } = render(<img src="x.png" />);
    const results = await axe(container);
    expect(results.violations.length).toBeGreaterThan(0);
  });
});
```

- [ ] **Step 7: Verify axe runs**

Run: `npm test -- --testPathPattern smokeAxe.test.tsx 2>&1 | tail -10`
Expected: 5 specs run (4 pass + 1 deliberate-failure-as-PASS). The regression sentinel proves the gate actually fails when violations exist.

- [ ] **Step 8: Write the conformance statement**

Create `docs/accessibility.md`:

```markdown
# SP Search — Accessibility Conformance Statement (WCAG 2.1 AA)

> Owned by Foundations track (Found.D6). Scoped — covers axe-core-tested surfaces only. Manual screen-reader pass on the full surface is v1.1+ work.

## Scope

This statement covers the four surfaces exercised by `tests/a11y/smokeAxe.test.tsx` (axe-core CI gate) plus the two A11Y-004 / A11Y-005 closures landed in this release:

1. **Search Box mode toggle** (`SpSearchBox.tsx:736`) — semantic `<fieldset>` + visually-hidden `<legend>`, role="radio" buttons with `aria-checked`. (A11Y-004 closed.)
2. **Search Box scope selector** (`SpSearchBox.tsx:729`) — Fluent UI Dropdown with `aria-describedby` linking to a hidden description span. (A11Y-005 closed.)
3. **Empty state markup** (`SearchResults.tsx`) — `role="status" aria-live="polite"`.
4. **Detail panel close button** (`ResultDetailPanel.tsx`) — explicit `aria-label`.

## Conformance

We claim WCAG 2.1 Level AA conformance for the surfaces enumerated above. All other surfaces inherit a baseline level via the SPFx host and Fluent UI v8 (which itself conforms to WCAG 2.1 AA), but have not been independently audited as of v1.0.

## Testing approach

- **Static analysis (CI gate)** — `axe-core` via `jest-axe` on every PR. New violations on the four enumerated surfaces fail the build.
- **Motion preference** — `prefers-reduced-motion: reduce` honored across all 6 module.scss files via the `src/styles/motion.scss` mixin. Verified via `grep -rn "prefers-reduced-motion" src/styles src/webparts` (≥6 hits).
- **Keyboard navigation** — Tab order verified manually for the Search Box, Filters drawer, Detail panel close button, and Manager tabs. Esc closes any open panel.
- **Focus visible** — relies on Fluent UI v8 default focus rings; no custom focus-ring suppression in `*.module.scss`.

## Known limitations (v1.0)

- No manual screen-reader (NVDA / JAWS / VoiceOver) pass on file. Full conformance verification is v1.1+.
- DataGrid layout (DevExtreme) accessibility relies on DevExtreme's own a11y posture; we do not re-audit it.
- Filters drawer does not yet ship `FocusTrapZone` (T1.D1 dep — Sprint 5).
- ManageAccess panel does not yet exist (T2.D5 — Sprint 6 deferred).

## Out-of-scope per Foundations Out-of-scope §1

- Exhaustive WCAG 2.1 AA audit beyond the top-10 surface gaps. Full audit deferred to v1.1+ once usage data identifies highest-leverage surfaces.

## Reporting an accessibility issue

File a bug via `.github/ISSUE_TEMPLATE/bug_report.md` with the surface URL, AT used, and reproduction steps. Tag `accessibility` label.
```

- [ ] **Step 9: Run smoke build to verify the whole pipeline**

Run: `npm test 2>&1 | tail -10`
Expected: all tests pass including the new `smokeAxe.test.tsx`.

Run: `npm run package 2>&1 | tail -5`
Expected: builds cleanly (no SCSS regressions from the motion mixin imports).

- [ ] **Step 10: Commit**

```bash
git add src/styles/motion.scss src/webparts/spSearchBox/components/SpSearchBox.tsx src/webparts/spSearchBox/components/SpSearchBox.module.scss src/webparts/**/*.module.scss tests/a11y/smokeAxe.test.tsx docs/accessibility.md package.json package-lock.json
git commit -m "$(cat <<'EOF'
a11y: WCAG 2.1 AA top-10 baseline + axe-core CI gate + accessibility.md (Found.D6)

Closes Still-Open A11Y-004 (mode toggle <fieldset>+<legend>) + A11Y-005
(scope selector aria-describedby). Adds shared src/styles/motion.scss
mixin honoring prefers-reduced-motion across 6 module.scss files (T1.D9
applies surface-by-surface in Sprint 5).

tests/a11y/smokeAxe.test.tsx ships 4 axe-clean assertions + 1 regression
sentinel proving the gate fails on a missing img alt. CI runs as part
of npm test; no separate workflow step needed.

docs/accessibility.md publishes the scoped WCAG 2.1 AA conformance
statement; full manual screen-reader pass deferred to v1.1+ per
Foundations Out-of-scope §1.

Closes Found.D6 P1 (audit Part 3 + Appendix A A11Y-004/A11Y-005).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 10: Found.D4 — HTML sanitizer adoption + `window.location.href` security review (`safeNavigate`)

**Why (P1):** Audit Found.D4. No new exposure has been demonstrated in the 7 unhardened sites (admin-configured properties, not user-typed); the hardening is precautionary, hence P1. T2.D3 covers the parallel JSON schema validation at the input side; this task closes the output / navigation side.

**Files:**
- Create: `src/libraries/spSearchStore/utils/safeNavigate.ts`
- Create: `tests/utils/safeNavigate.test.ts`
- Modify: `src/libraries/spSearchStore/providers/QuickResultsSuggestionProvider.ts:80`
- Modify: `src/webparts/spSearchManager/components/SpSearchManager.tsx:646` (line 369/548/644 are `read`-only `new URL(window.location.href)` — exempt)
- Modify: `src/webparts/spSearchResults/components/SpSearchResults.tsx:570` (line 568 read-only — exempt)
- Modify: `src/webparts/spSearchBox/components/SpSearchBox.tsx:358` (convert existing inline allowlist to call `safeNavigate`)
- Modify: `src/webparts/spSearchResults/components/documentTitleUtils.ts:158` (delete local `sanitizeSummaryHtml`)
- Modify: `src/webparts/spSearchResults/components/ListLayout.tsx:69` (import sanitizer from `spfx-toolkit/lib/utilities/htmlUtils/sanitizeHtml`)

- [ ] **Step 1: Verify spfx-toolkit sanitizer surface exists**

Run: `ls /Users/hemantmane/Development/spfx-toolkit/lib/utilities/htmlUtils/sanitizeHtml* 2>&1`
Expected: `sanitizeHtml.js` + `.d.ts` exist. If missing (toolkit out of date locally), run `cd /Users/hemantmane/Development/spfx-toolkit && npm run build` first.

- [ ] **Step 2: Write the failing safeNavigate tests**

Create `tests/utils/safeNavigate.test.ts`:

```typescript
import { safeNavigate } from '../../src/libraries/spSearchStore/utils/safeNavigate';

describe('safeNavigate (Found.D4)', () => {
  let originalAssign: typeof window.location.assign;
  let assignedTo: string | null = null;

  beforeEach(() => {
    assignedTo = null;
    // jsdom's window.location.assign is non-spyable in older versions; redefine
    Object.defineProperty(window, 'location', {
      writable: true,
      value: {
        ...window.location,
        assign: (url: string) => { assignedTo = url; },
      },
    });
  });

  it('allows https:// URLs', () => {
    expect(safeNavigate('https://example.com/doc.pdf')).toBe(true);
    expect(assignedTo).toBe('https://example.com/doc.pdf');
  });

  it('allows http:// URLs', () => {
    expect(safeNavigate('http://example.com/')).toBe(true);
    expect(assignedTo).toBe('http://example.com/');
  });

  it('allows root-relative paths', () => {
    expect(safeNavigate('/sites/SPSearch/Pages/Search.aspx')).toBe(true);
    expect(assignedTo).toBe('/sites/SPSearch/Pages/Search.aspx');
  });

  it('rejects javascript: URLs', () => {
    expect(safeNavigate('javascript:alert(1)')).toBe(false);
    expect(assignedTo).toBeNull();
  });

  it('rejects data: URLs', () => {
    expect(safeNavigate('data:text/html,<script>alert(1)</script>')).toBe(false);
    expect(assignedTo).toBeNull();
  });

  it('rejects empty / null / undefined', () => {
    expect(safeNavigate('')).toBe(false);
    expect(safeNavigate(null as any)).toBe(false);
    expect(safeNavigate(undefined as any)).toBe(false);
    expect(assignedTo).toBeNull();
  });

  it('rejects whitespace-only', () => {
    expect(safeNavigate('   ')).toBe(false);
    expect(assignedTo).toBeNull();
  });
});
```

Run: `npm test -- --testPathPattern safeNavigate.test.ts 2>&1 | tail -10`
Expected: tests fail with `Cannot find module ../../src/libraries/spSearchStore/utils/safeNavigate`.

- [ ] **Step 3: Implement safeNavigate**

Create `src/libraries/spSearchStore/utils/safeNavigate.ts`:

```typescript
/**
 * Centralised navigation policy (Foundations Found.D4).
 * Validates the target URL against an https? / root-relative allowlist,
 * rejecting javascript:, data:, and other dangerous schemes.
 *
 * Returns true if navigation occurred, false otherwise. Never throws.
 *
 * Usage: replace any direct `window.location.href = X` write with
 * `safeNavigate(X)`. ESLint rule (or grep guard) flags new direct writes.
 */
export function safeNavigate(target: string | null | undefined): boolean {
  if (typeof target !== 'string') return false;
  const trimmed = target.trim();
  if (trimmed.length === 0) return false;

  // Allowlist: absolute http/https or root-relative paths only.
  const isAbsoluteHttp = /^https?:\/\//i.test(trimmed);
  const isRootRelative = trimmed.startsWith('/') && !trimmed.startsWith('//');
  if (!isAbsoluteHttp && !isRootRelative) return false;

  // Reject the dangerous schemes explicitly even if they snuck past the allowlist.
  const lower = trimmed.toLowerCase();
  if (lower.startsWith('javascript:') || lower.startsWith('data:') || lower.startsWith('vbscript:')) {
    return false;
  }

  window.location.assign(trimmed);
  return true;
}
```

Run: `npm test -- --testPathPattern safeNavigate.test.ts 2>&1 | tail -10`
Expected: all 7 specs pass.

- [ ] **Step 4: Migrate the navigation sites**

For each navigation site, replace `window.location.href = X` with `safeNavigate(X)`:

- `src/libraries/spSearchStore/providers/QuickResultsSuggestionProvider.ts:80`:
  ```typescript
  // before: window.location.href = item.url;
  // after:
  import { safeNavigate } from '@store/utils/safeNavigate';
  // ...
  safeNavigate(item.url);
  ```

- `src/webparts/spSearchManager/components/SpSearchManager.tsx:646`:
  ```typescript
  // before: window.location.href = url.toString();
  // after:
  safeNavigate(url.toString());
  ```
  (Add the import at the top of the file via `@store/utils/safeNavigate`.)

- `src/webparts/spSearchResults/components/SpSearchResults.tsx:570`:
  ```typescript
  safeNavigate(url.toString());
  ```

- `src/webparts/spSearchBox/components/SpSearchBox.tsx:358`:
  Replace the existing inline `https?://` / `/` allowlist (BUG-004 hardening) with the centralised helper:
  ```typescript
  safeNavigate(targetUrl);
  ```

Lines 369/548/644 of `SpSearchManager.tsx` and line 568 of `SpSearchResults.tsx` are `read`-only `new URL(window.location.href)` calls (capturing the current URL for serialization). Those are exempt from the hardening — `new URL(window.location.href)` cannot inject a navigation. Leave them unchanged but add a comment:

```typescript
// safe: read-only URL capture for serialization (Found.D4 exempt)
const url = new URL(window.location.href);
```

- [ ] **Step 5: Adopt the toolkit sanitizer**

Edit `src/webparts/spSearchResults/components/ListLayout.tsx:8` and `:69`:

```typescript
// before:
// import { ..., sanitizeSummaryHtml, ... } from './documentTitleUtils';
// dangerouslySetInnerHTML={{ __html: sanitizeSummaryHtml(item.summary) }}

// after:
import { sanitizeHtml } from 'spfx-toolkit/lib/utilities/htmlUtils/sanitizeHtml';
// ...
dangerouslySetInnerHTML={{ __html: sanitizeHtml(item.summary) }}
```

Edit `src/webparts/spSearchResults/components/documentTitleUtils.ts` — delete the local `sanitizeSummaryHtml` function (lines 155-end per audit body). Verify other consumers via:

```bash
grep -rn "sanitizeSummaryHtml" src
```

Expected: 0 hits after the deletion + ListLayout migration.

- [ ] **Step 6: Verify the security sweep**

Run:
```bash
grep -rn "window\.location\.href = " src
```

Expected: 0 direct-assignment hits. The remaining `window.location.href` references are read-only `new URL(window.location.href)` capture calls (lines 369/548/644 in SpSearchManager.tsx, line 568 in SpSearchResults.tsx).

```bash
grep -rn "spfx-toolkit/lib/utilities/htmlUtils" src
```

Expected: ≥1 hit (ListLayout.tsx).

- [ ] **Step 7: Run the full test suite**

Run: `npm test 2>&1 | tail -10`
Expected: all tests pass; `safeNavigate.test.ts` reports 7 specs.

- [ ] **Step 8: Re-run BUG-004 closure smoke**

Manually verify the BUG-004 closure is preserved: in the dev workbench, set a Search Box `newPageUrl` property to `javascript:alert(1)` and confirm clicking does not execute. (Can be deferred to D2 smoke checklist Step 6 on next run.)

- [ ] **Step 9: Commit**

```bash
git add src/libraries/spSearchStore/utils/safeNavigate.ts tests/utils/safeNavigate.test.ts src/libraries/spSearchStore/providers/QuickResultsSuggestionProvider.ts src/webparts/spSearchManager/components/SpSearchManager.tsx src/webparts/spSearchResults/components/SpSearchResults.tsx src/webparts/spSearchBox/components/SpSearchBox.tsx src/webparts/spSearchResults/components/ListLayout.tsx src/webparts/spSearchResults/components/documentTitleUtils.ts
git commit -m "$(cat <<'EOF'
sec: safeNavigate helper + spfx-toolkit sanitizeHtml adoption (Found.D4)

Centralises navigation policy in src/libraries/spSearchStore/utils/safeNavigate.ts;
all 5 in-product navigation sites (QuickResultsSuggestionProvider:80,
SpSearchManager:646, SpSearchResults:570, SpSearchBox:358) now route
through it. The 4 read-only `new URL(window.location.href)` capture
sites (Manager:369/548/644, Results:568) are exempt with comments.

Adopts spfx-toolkit/lib/utilities/htmlUtils/sanitizeHtml at
ListLayout.tsx:69; deletes the local sanitizeSummaryHtml from
documentTitleUtils.ts. 7-case unit test in tests/utils/safeNavigate.test.ts
covers https / http / root-relative / javascript / data / empty / whitespace.

T2.D3 (saved-search JSON schema validation) covers the input side
in Sprint 4.

Closes Found.D4 P1 (audit Part 3 + Appendix B HTML sanitization Adopt + Phase 7).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 11: Found.D10 — Declare `webApiPermissionRequests` for `People.Read`

**Why (P1):** Audit Found.D10. Workaround "add the permission manually" is documented (`admin-guide.md:240-242`), hence P1. T4.D9 (Pre-Flight tab) consumes this — the row goes green when the declaration ships.

**Files:**
- Modify: `config/package-solution.json`
- Modify: `docs/deployment-guide.md` (Smoke Test Checklist) and `README.md` (install path note)

- [ ] **Step 1: Verify the actual Graph endpoint and least-privileged scope before locking the declaration**

The audit assumes `People.Read` based on `admin-guide.md:240-242` and the prior Sprint-3 debugging captured in CLAUDE.md memory. But `GraphSearchProvider` calls `/search/query` with `entityTypes: ['person']` (verify at `src/libraries/spSearchStore/providers/GraphSearchProvider.ts:182`), NOT `/me/people`. Microsoft documents `/search/query` and `/me/people` permissions separately:

- People API (`/me/people`, `/users/{id}/people`): https://learn.microsoft.com/en-us/graph/people-insights-overview
- Microsoft Search query API (`/search/query`): https://learn.microsoft.com/en-us/graph/api/search-query

Before declaring the permission, verify against current Graph docs:

```bash
# Confirm the Graph endpoint our code actually hits
grep -n "\.api(" src/libraries/spSearchStore/providers/GraphSearchProvider.ts
grep -n "entityTypes" src/libraries/spSearchStore/providers/GraphSearchProvider.ts
```

Then fetch the current Graph docs for `/search/query` permissions per `entityTypes`. Use the WebFetch tool or the Context7 MCP server (`mcp__plugin_context7_context7__query-docs` with library id `microsoft-graph` if available) — do NOT trust training data alone for the scope name. The endpoint–scope mapping has changed across Graph API versions; what was correct in Sprint-3 debugging may not be the current least-privileged scope.

Document the verified scope in `/tmp/d10-scope-verification.md` with the source URL and date before proceeding to Step 2.

If the verified scope is `People.Read`: proceed with Step 2 as written.
If the verified scope differs (e.g. Graph search per-entityType permissions specify `User.Read.All` or `Sites.Read.All` for the `person` entityType): substitute the correct scope name in Step 2 and update `docs/admin-guide.md:240-242` to match in the same commit.

- [ ] **Step 2: Add the declaration with the verified scope**

Edit `config/package-solution.json` — add to the `solution` block (after `developer`, before `metadata`). The example below uses `People.Read` from the audit body; substitute the scope verified in Step 1 if it differs:

```json
"webApiPermissionRequests": [
  {
    "resource": "Microsoft Graph",
    "scope": "People.Read"
  }
],
```

- [ ] **Step 3: Re-package and verify the manifest carries the request**

Run: `npm run package 2>&1 | tail -5`

Inspect the generated `.sppkg` (it's a zip):
```bash
mkdir -p /tmp/sppkg-inspect && unzip -o sharepoint/solution/sp-search.sppkg -d /tmp/sppkg-inspect 2>&1 | tail -5
grep -rn "People.Read" /tmp/sppkg-inspect/ 2>&1 | head -5
```

Expected: at least one match in the extracted manifests confirming the scope is present.

- [ ] **Step 4: Document the post-install approval step**

Edit `docs/deployment-guide.md` Smoke Test Checklist section — add:

```markdown
- After uploading the `.sppkg`, the SharePoint admin center "API access" page surfaces a pending `People.Read` request for Microsoft Graph. Approve it before deploying the People vertical (Graph-backed search results).
```

Edit `README.md` install section — add a one-line note:

```markdown
> After deploying, approve the `Microsoft Graph: People.Read` API access request in the SharePoint admin center for the People vertical to function.
```

- [ ] **Step 5: Smoke-test on a tenant (manual; deferred to next D2 smoke checklist run)**

Note in the next `docs/release-runs/<tag>.md` Step 6 entry: "API access page shows `People.Read` pending — approved." Cannot be CI-tested.

- [ ] **Step 6: Commit**

```bash
git add config/package-solution.json docs/deployment-guide.md README.md
git commit -m "$(cat <<'EOF'
feat(security): declare webApiPermissionRequests for People.Read (Found.D10)

Adds the People.Read Microsoft Graph permission request to
config/package-solution.json so the SharePoint admin "API access"
page surfaces a pending approval after .sppkg upload. Without this,
the Graph People vertical silently no-ops on Day 1 (admin must add
the permission manually per admin-guide.md:240-242).

T4.D9 Pre-Flight tab row "(a) webApiPermissionRequests declared" now
goes green.

Closes Found.D10 P1 (audit Part 3 + Journey A Step 1 [Polish]).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 12: Found.D9 — Telemetry plumbing (opt-in tenant config + transport service + privacy notice)

**Why (P1):** Audit Found.D9. No shipped feature consumes it without T5.D8 + T5.D9 also landing, hence P1. Foundations ships the wire and the storage; T5 ships the schema and the consumer. The destination is admin-configured per the SearchTelemetryConfig list — same plumbing supports App Insights, Azure Monitor, custom HTTPS, or tenant-internal logger.

**Files:**
- Create: `src/libraries/spSearchStore/telemetry/ITelemetryConfig.ts`
- Create: `src/libraries/spSearchStore/telemetry/ITelemetrySignal.ts` (placeholder; T5.D8 owns the canonical schema later)
- Create: `src/libraries/spSearchStore/telemetry/TelemetryTransport.ts`
- Create: `tests/telemetry/TelemetryTransport.test.ts`
- Create: `docs/privacy-notice.md`
- Modify: `scripts/Provision-SPSearchLists.ps1` — add `SearchTelemetryConfig` (single-row hidden) + `SearchTelemetryOptIn` (per-user) lists
- Modify: `README.md` (link `docs/privacy-notice.md`) — already added in D5 step 2

- [ ] **Step 1: Write the config + signal interfaces**

Create `src/libraries/spSearchStore/telemetry/ITelemetryConfig.ts`:

```typescript
export interface ITelemetryConfig {
  isEnabled: boolean;
  destinationEndpoint: string;
  batchIntervalSeconds: number;
  batchSizeMax: number;
  privacyAcknowledgedBy?: string;
  privacyAcknowledgedAt?: string;
}

export const TELEMETRY_DEFAULTS: ITelemetryConfig = {
  isEnabled: false,
  destinationEndpoint: '',
  batchIntervalSeconds: 300,
  batchSizeMax: 50,
};
```

Create `src/libraries/spSearchStore/telemetry/ITelemetrySignal.ts`:

```typescript
/**
 * Minimal signal interface so Foundations Found.D9 transport can compile
 * before T5.D8 lands the canonical schema. T5.D8 replaces this interface
 * with the full type discriminated union (kind: 'queryTiming' | 'errorRate' | ...).
 *
 * Foundations enforces: NEVER capture queryText, userId, resultTitle, urls,
 * tenantName, or list item content. T5.D8 enforces this at compile time
 * via the ITelemetrySignal discriminated union.
 */
export interface ITelemetrySignal {
  kind: string;
  timestamp: string;
  // Type-erased payload until T5.D8 lands the schema. Transport never
  // inspects payload contents — purely the wire.
  [key: string]: unknown;
}
```

- [ ] **Step 2: Write the failing transport test**

Create `tests/telemetry/TelemetryTransport.test.ts`:

```typescript
import { TelemetryTransport } from '../../src/libraries/spSearchStore/telemetry/TelemetryTransport';
import { ITelemetryConfig } from '../../src/libraries/spSearchStore/telemetry/ITelemetryConfig';

describe('TelemetryTransport (Found.D9)', () => {
  it('flush is a no-op when isEnabled=false', async () => {
    const fetchMock = jest.fn();
    (global as any).fetch = fetchMock;
    const config: ITelemetryConfig = {
      isEnabled: false,
      destinationEndpoint: 'https://example.com/telemetry',
      batchIntervalSeconds: 300,
      batchSizeMax: 50,
    };
    const transport = new TelemetryTransport(() => Promise.resolve(config));
    await transport.flush([{ kind: 'queryTiming', timestamp: new Date().toISOString() }]);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('flush POSTs to destinationEndpoint when isEnabled=true', async () => {
    const fetchMock = jest.fn().mockResolvedValue({ ok: true, status: 200 });
    (global as any).fetch = fetchMock;
    const config: ITelemetryConfig = {
      isEnabled: true,
      destinationEndpoint: 'https://example.com/telemetry',
      batchIntervalSeconds: 300,
      batchSizeMax: 50,
    };
    const transport = new TelemetryTransport(() => Promise.resolve(config));
    await transport.flush([{ kind: 'queryTiming', timestamp: '2026-05-02T00:00:00Z' }]);
    expect(fetchMock).toHaveBeenCalledWith(
      'https://example.com/telemetry',
      expect.objectContaining({ method: 'POST', headers: { 'Content-Type': 'application/json' } })
    );
  });

  it('flush retries with exponential backoff on 5xx, max 3 attempts', async () => {
    let attempt = 0;
    const fetchMock = jest.fn().mockImplementation(() => {
      attempt++;
      if (attempt < 3) return Promise.resolve({ ok: false, status: 503 });
      return Promise.resolve({ ok: true, status: 200 });
    });
    (global as any).fetch = fetchMock;
    const config: ITelemetryConfig = {
      isEnabled: true,
      destinationEndpoint: 'https://example.com/telemetry',
      batchIntervalSeconds: 300,
      batchSizeMax: 50,
    };
    const transport = new TelemetryTransport(() => Promise.resolve(config), { backoffMs: 0 });
    await transport.flush([{ kind: 'queryTiming', timestamp: '2026-05-02T00:00:00Z' }]);
    expect(fetchMock).toHaveBeenCalledTimes(3);
  });

  it('flush does not throw when fetch rejects', async () => {
    (global as any).fetch = jest.fn().mockRejectedValue(new Error('network'));
    const config: ITelemetryConfig = {
      isEnabled: true,
      destinationEndpoint: 'https://example.com/telemetry',
      batchIntervalSeconds: 300,
      batchSizeMax: 50,
    };
    const transport = new TelemetryTransport(() => Promise.resolve(config), { backoffMs: 0, maxAttempts: 1 });
    await expect(transport.flush([{ kind: 'queryTiming', timestamp: '2026-05-02T00:00:00Z' }])).resolves.toBeUndefined();
  });
});
```

Run: `npm test -- --testPathPattern TelemetryTransport.test.ts 2>&1 | tail -10`
Expected: failure — module not found.

- [ ] **Step 3: Implement the transport**

Create `src/libraries/spSearchStore/telemetry/TelemetryTransport.ts`:

```typescript
import { ITelemetryConfig } from './ITelemetryConfig';
import { ITelemetrySignal } from './ITelemetrySignal';

export interface TelemetryTransportOptions {
  backoffMs?: number;       // base backoff; doubles per attempt
  maxAttempts?: number;     // default 3
  configRefreshSeconds?: number;  // default 60
}

/**
 * Foundations Found.D9 — HTTPS POST transport for opt-in telemetry.
 * Never inspects payload contents. T5.D8's ITelemetrySignal discriminated
 * union enforces the never-captured field list at compile time.
 *
 * Usage: instantiated by sp-search-store; consumes the SearchTelemetryConfig
 * SP list via the configLoader callback. Returns immediately when
 * config.isEnabled === false (the no-op default).
 */
export class TelemetryTransport {
  private cachedConfig: ITelemetryConfig | null = null;
  private lastConfigLoad = 0;
  private readonly backoffMs: number;
  private readonly maxAttempts: number;
  private readonly configRefreshMs: number;

  constructor(
    private readonly configLoader: () => Promise<ITelemetryConfig>,
    options: TelemetryTransportOptions = {},
  ) {
    this.backoffMs = options.backoffMs ?? 2000;
    this.maxAttempts = options.maxAttempts ?? 3;
    this.configRefreshMs = (options.configRefreshSeconds ?? 60) * 1000;
  }

  public async flush(batch: ITelemetrySignal[]): Promise<void> {
    if (batch.length === 0) return;
    const config = await this.loadConfig();
    if (!config.isEnabled || !config.destinationEndpoint) return;

    const body = JSON.stringify({ signals: batch });
    let attempt = 0;
    while (attempt < this.maxAttempts) {
      attempt++;
      try {
        const res = await fetch(config.destinationEndpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body,
        });
        if (res.ok) return;
        if (res.status >= 400 && res.status < 500 && res.status !== 408 && res.status !== 429) {
          return;  // permanent client error — drop
        }
      } catch {
        // network error — retry with backoff
      }
      if (attempt < this.maxAttempts) {
        const delay = this.backoffMs * Math.pow(2, attempt - 1);
        if (delay > 0) await new Promise(resolve => setTimeout(resolve, delay));
      }
    }
  }

  private async loadConfig(): Promise<ITelemetryConfig> {
    const now = Date.now();
    if (this.cachedConfig && (now - this.lastConfigLoad) < this.configRefreshMs) {
      return this.cachedConfig;
    }
    this.cachedConfig = await this.configLoader();
    this.lastConfigLoad = now;
    return this.cachedConfig;
  }
}
```

Run: `npm test -- --testPathPattern TelemetryTransport.test.ts 2>&1 | tail -10`
Expected: all 4 specs pass.

- [ ] **Step 4: Write the privacy notice**

Create `docs/privacy-notice.md`:

```markdown
# SP Search — Privacy Notice (Telemetry)

> Owned by Foundations track (Found.D9). Read this before enabling telemetry. T5.D8 ships the schema; T5.D9 ships the aggregate dashboard view.

## What we collect (when telemetry is enabled)

When an admin enables telemetry via the `SearchTelemetryConfig` list and a user opts in via the Admin Manager Telemetry property pane group:

- Query timing — milliseconds end-to-end per search request
- Error rates — count of failed search requests, grouped by error class (no message bodies)
- Refiner usage — count of filter applications, grouped by filter type (Checkbox, DateRange, etc.)
- Layout switches — count of layout changes, grouped by layout key (DataGrid, Card, etc.)
- Vertical switches — count of vertical changes, grouped by vertical key
- Feature adoption — flags indicating whether a user opens the detail panel, opens the Search Manager, exports CSV/XLSX
- Anonymized session ID — SHA-256 hash of `tenantId + userPrincipalName + 'sp-search-telemetry-v1'`, truncated to first 8 hex chars

All counts are aggregated client-side per `BatchIntervalSeconds` (default 300s) before transmission.

## What we NEVER collect

- Query text (the literal string typed by users)
- User identity (email, login name, display name, UPN, or any reversible token)
- Result titles, URLs, or summaries
- Tenant name, site collection name, list name, or item content
- Page URL or referrer
- IP address (relies on transport infrastructure to redact at the destination)
- Browser fingerprint, geolocation, or device identifier

The `ITelemetrySignal` interface (T5.D8) enforces the never-captured field list at compile time. The transport (`TelemetryTransport.ts`) never inspects payload contents — it is purely the wire.

## Where the data goes

The destination is admin-configured in the `SearchTelemetryConfig` list (`DestinationEndpoint` field). Same plumbing supports:
- Application Insights ingestion endpoint
- Azure Monitor custom logs
- Tenant-internal log collector
- A custom HTTPS POST endpoint of the admin's choice

The `.sppkg` ships with telemetry **disabled by default** (`IsEnabled: false`). No data leaves the tenant unless an admin both (a) sets `IsEnabled: true` + a destination URL, and (b) at least one user opts in via the Admin Manager property pane.

## Opt-in / opt-out

- **Opt in** — Admin sets `SearchTelemetryConfig.IsEnabled = true` + a destination URL. End users see the "View what we send" Panel (T5.D8) on the property pane and can opt in per user. Opt-in events recorded in the `SearchTelemetryOptIn` list (per-user, anonymized hash only).
- **Opt out** — End users can clear their opt-in by toggling the Admin Manager property pane setting back to off; immediately stops telemetry transmission for that user. Admin can disable tenant-wide by setting `SearchTelemetryConfig.IsEnabled = false`.

## Data retention

The transport does not retain anything client-side beyond the in-flight batch. Retention at the destination is the admin's policy.

## Compliance

This plumbing is compliant by design with the spec §4.3 T5 "never captured" list. Tenant-specific compliance posture (GDPR, CCPA, HIPAA, FedRAMP) depends on the destination endpoint and the admin's data processing agreement with that endpoint.

## Reporting a privacy concern

File a bug via `.github/ISSUE_TEMPLATE/bug_report.md` tagged `privacy`. Include the SearchTelemetryConfig.DestinationEndpoint value and the specific signal that surfaced the concern.
```

- [ ] **Step 5: Extend the provisioning script with the two SP lists**

Edit `scripts/Provision-SPSearchLists.ps1` — append a new section provisioning the two telemetry lists:

```powershell
# ============================================================
# Foundations Found.D9 — Telemetry lists
# ============================================================

function Add-SearchTelemetryLists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )

    Write-Host "[Found.D9] Provisioning SearchTelemetryConfig + SearchTelemetryOptIn..." -ForegroundColor Cyan

    # SearchTelemetryConfig — single-row, hidden, item-level read for everyone, write for Members
    $configList = Get-PnPList -Identity 'SearchTelemetryConfig' -ErrorAction SilentlyContinue
    if (-not $configList) {
        $configList = New-PnPList -Title 'SearchTelemetryConfig' -Template GenericList -Url 'Lists/SearchTelemetryConfig' -OnQuickLaunch:$false
        Set-PnPList -Identity 'SearchTelemetryConfig' -Hidden:$true
        Add-PnPField -List 'SearchTelemetryConfig' -DisplayName 'IsEnabled' -InternalName 'IsEnabled' -Type Boolean -AddToDefaultView | Out-Null
        Add-PnPField -List 'SearchTelemetryConfig' -DisplayName 'DestinationEndpoint' -InternalName 'DestinationEndpoint' -Type Text -AddToDefaultView | Out-Null
        Add-PnPField -List 'SearchTelemetryConfig' -DisplayName 'BatchIntervalSeconds' -InternalName 'BatchIntervalSeconds' -Type Number -AddToDefaultView | Out-Null
        Add-PnPField -List 'SearchTelemetryConfig' -DisplayName 'BatchSizeMax' -InternalName 'BatchSizeMax' -Type Number -AddToDefaultView | Out-Null
        Add-PnPField -List 'SearchTelemetryConfig' -DisplayName 'PrivacyAcknowledgedBy' -InternalName 'PrivacyAcknowledgedBy' -Type Text -AddToDefaultView | Out-Null
        Add-PnPField -List 'SearchTelemetryConfig' -DisplayName 'PrivacyAcknowledgedAt' -InternalName 'PrivacyAcknowledgedAt' -Type DateTime -AddToDefaultView | Out-Null

        # Single-row default (disabled). Admins toggle IsEnabled to opt in.
        Add-PnPListItem -List 'SearchTelemetryConfig' -Values @{
            'Title' = 'SP Search Telemetry Config (single row)'
            'IsEnabled' = $false
            'DestinationEndpoint' = ''
            'BatchIntervalSeconds' = 300
            'BatchSizeMax' = 50
        } | Out-Null

        Write-Host "[Found.D9] SearchTelemetryConfig provisioned (disabled, default row)." -ForegroundColor Green
    }
    else {
        Write-Host "[Found.D9] SearchTelemetryConfig already exists; skipping." -ForegroundColor Yellow
    }

    # SearchTelemetryOptIn — per-user consent, anonymized hash only
    $optInList = Get-PnPList -Identity 'SearchTelemetryOptIn' -ErrorAction SilentlyContinue
    if (-not $optInList) {
        New-PnPList -Title 'SearchTelemetryOptIn' -Template GenericList -Url 'Lists/SearchTelemetryOptIn' -OnQuickLaunch:$false | Out-Null
        Set-PnPList -Identity 'SearchTelemetryOptIn' -Hidden:$true
        Add-PnPField -List 'SearchTelemetryOptIn' -DisplayName 'ConsentTimestamp' -InternalName 'ConsentTimestamp' -Type DateTime -AddToDefaultView | Out-Null
        Add-PnPField -List 'SearchTelemetryOptIn' -DisplayName 'ConsentVersion' -InternalName 'ConsentVersion' -Type Text -AddToDefaultView | Out-Null
        Add-PnPField -List 'SearchTelemetryOptIn' -DisplayName 'AnonHash' -InternalName 'AnonHash' -Type Text -AddToDefaultView | Out-Null
        Write-Host "[Found.D9] SearchTelemetryOptIn provisioned." -ForegroundColor Green
    }
    else {
        Write-Host "[Found.D9] SearchTelemetryOptIn already exists; skipping." -ForegroundColor Yellow
    }
}

# Invocation: append to the existing main flow
# Add-SearchTelemetryLists -SiteUrl $SiteUrl
```

(Wire the `Add-SearchTelemetryLists` invocation into the existing main flow at the bottom of the script.)

- [ ] **Step 6: Update the README link to the privacy notice**

Verify `README.md` (created in D5 Step 2) already lists `docs/privacy-notice.md`. If not, add it under Documentation.

- [ ] **Step 7: Run the test suite**

Run: `npm test 2>&1 | tail -10`
Expected: all tests pass; `TelemetryTransport.test.ts` reports 4 specs.

- [ ] **Step 8: Commit**

```bash
git add src/libraries/spSearchStore/telemetry/ tests/telemetry/ docs/privacy-notice.md scripts/Provision-SPSearchLists.ps1
git commit -m "$(cat <<'EOF'
feat(telemetry): opt-in plumbing — SP list config + transport + privacy notice (Found.D9)

Foundations-side substrate that T5.D8 (schema + emitter shim) and
T5.D9 (aggregate dashboard) consume. Ships:

- src/libraries/spSearchStore/telemetry/ITelemetryConfig.ts: list shape
- src/libraries/spSearchStore/telemetry/ITelemetrySignal.ts: minimal
  signal interface (T5.D8 lands the canonical schema)
- src/libraries/spSearchStore/telemetry/TelemetryTransport.ts: HTTPS
  POST with exponential backoff (max 3 attempts), no-op when
  isEnabled=false, never inspects payload contents
- tests/telemetry/TelemetryTransport.test.ts: 4 specs covering
  no-op default, POST shape, retry with backoff, swallowed network errors
- docs/privacy-notice.md: "what we collect / never collect" lists +
  destination policy + opt-in/out
- scripts/Provision-SPSearchLists.ps1: adds SearchTelemetryConfig
  (single-row hidden, disabled by default) + SearchTelemetryOptIn
  (per-user, anonymized hash only) lists

Telemetry ships disabled by default; no data leaves the tenant
unless admin opts in (config list IsEnabled=true + destination URL)
AND user opts in (T5.D8 property pane).

Closes Found.D9 P1 (audit Part 3 + spec §4.4 telemetry plumbing).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 13: Found.D11 — `solution.developer.mpnId` cleanup + `package.json` version alignment

**Why (P2):** Audit Found.D11. Cosmetic only at install time; rides along with D2's release tag (already shipped in Tranche 1) + D8's automation. T4.D10 explicitly bundles the same `mpnId` fix — coordinate so the change ships exactly once.

**Files:**
- Modify: `package.json:3` (`version: "0.0.1"` → `"1.0.0"`)
- Modify: `config/package-solution.json` — clear `mpnId`; populate `developer.{websiteUrl, privacyUrl, termsOfUseUrl}`; bump `solution.version` from `1.0.13.0` to `1.0.0.0` (lockstep with package.json per release-policy.md)

- [ ] **Step 1: Decide the MPN ID** — Ask the user: "Does the publishing organisation have a Partner Center MPN ID? If yes, what is it? If no or undecided, ship empty string with a follow-up ticket."

If the user answers with a real MPN ID: use it. If empty or undecided: ship `""`.

- [ ] **Step 2: Pick the canonical project URLs**
- `websiteUrl` — repo URL (e.g. `https://github.com/<org>/sp-search`)
- `privacyUrl` — link to the published `docs/privacy-notice.md` raw URL or a hosted equivalent
- `termsOfUseUrl` — link to the LICENSE file in the repo, or a hosted ToS

If any are not yet decided: ship `""` and open a follow-up ticket per the audit body.

- [ ] **Step 3: Update package.json**

Edit `package.json:3`:

```json
"version": "1.0.0",
```

- [ ] **Step 4: Update package-solution.json**

Edit `config/package-solution.json`:

```json
"version": "1.0.0.0",
...
"developer": {
  "name": "Hemant Mane",
  "websiteUrl": "https://github.com/<org>/sp-search",
  "privacyUrl": "https://raw.githubusercontent.com/<org>/sp-search/main/docs/privacy-notice.md",
  "termsOfUseUrl": "https://github.com/<org>/sp-search/blob/main/LICENSE",
  "mpnId": ""
},
```

(Replace `<org>` with the actual GitHub organisation. If repo is not yet published, ship `""` for the URLs and open a follow-up ticket.)

- [ ] **Step 5: Verify the lockstep gate from D8 build.yml passes**

Run: `npm run package 2>&1 | tail -5`
Expected: clean build.

Locally simulate the lockstep check:
```bash
PKG_VER=$(node -p "require('./package.json').version")
SOL_VER=$(node -p "require('./config/package-solution.json').solution.version")
SOL_TRIM=$(echo "$SOL_VER" | awk -F. '{print $1"."$2"."$3}')
[ "$PKG_VER" = "$SOL_TRIM" ] && echo "lockstep OK ($PKG_VER / $SOL_VER)" || echo "MISMATCH: $PKG_VER vs $SOL_VER"
```

Expected: `lockstep OK (1.0.0 / 1.0.0.0)`.

- [ ] **Step 6: Inspect the .sppkg manifest for the cleaned metadata**

```bash
mkdir -p /tmp/sppkg-d11 && unzip -o sharepoint/solution/sp-search.sppkg -d /tmp/sppkg-d11 2>&1 | tail -3
grep -rn "Undefined-1.21" /tmp/sppkg-d11/ 2>&1 | head -3
```

Expected: 0 hits in the extracted manifest.

- [ ] **Step 7: Update CHANGELOG**

Edit `CHANGELOG.md` `[Unreleased]` section — add:

```markdown
### Changed
- `package.json:version` aligned to `1.0.0` from generator default `0.0.1`; `config/package-solution.json:solution.version` aligned to `1.0.0.0` (lockstep — Found.D11).
- `solution.developer.mpnId` cleared from `Undefined-1.21.1` (SPFx generator default) to empty string [or real MPN ID] (Found.D11).
- `solution.developer.websiteUrl / privacyUrl / termsOfUseUrl` populated with canonical project URLs (Found.D11).
```

- [ ] **Step 8: Commit**

```bash
git add package.json config/package-solution.json CHANGELOG.md
git commit -m "$(cat <<'EOF'
chore(release): mpnId cleanup + package.json/solution.version lockstep at 1.0.0 (Found.D11)

Drops the SPFx generator default solution.developer.mpnId =
"Undefined-1.21.1" that was carrying through to the SharePoint admin
"Apps you can add" pane. Populates developer.websiteUrl /
privacyUrl / termsOfUseUrl with canonical project URLs. Aligns
package.json:version (0.0.1) and solution.version (1.0.13.0) at
1.0.0 / 1.0.0.0 per docs/release-policy.md lockstep convention.

T4.D10 owns the parallel Deploy-SPSearchSolution.ps1 -ReleaseArtifactUrl
automation; this change ships only the metadata + version-alignment
subset Foundations owns at release-tag time.

Closes Found.D11 P2 (audit Part 3 + Journey A Step 2 [Polish]).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

# Tranche checklist — final state

After all 13 tasks complete:

- [ ] `npm run package` exits 0 from a clean checkout; `sharepoint/solution/sp-search.sppkg` exists.
- [ ] `npm run type-check` exits 0.
- [ ] `npm test` exits 0; ≥1 spec passes (`tests/store/lifecycle.test.ts` minimum).
- [ ] `npm run check:bundles` exits 0; `release/analysis-logs/bundle-sizes.json` written.
- [ ] `git tag --list 'v*'` includes `v1.0.0-rc.1`.
- [ ] `main` branch HEAD contains the squashed SPFx 1.22 / Heft migration commit.
- [ ] `.github/workflows/{build.yml,release.yml}` present; first PR runs CI green.
- [ ] `CHANGELOG.md`, `CONTRIBUTING.md`, `README.md`, `docs/release-policy.md`, `docs/release-smoke-checklist.md`, `docs/release-runs/v1.0.0-rc.1.md`, `docs/performance-budgets.md`, `docs/accessibility.md`, `docs/privacy-notice.md` all exist.
- [ ] `grep -n "SPFx 1\.21\|gulp " CLAUDE.md` returns 0 hits in core sections.
- [ ] `grep -n "Toast" CLAUDE.md` returns 0 hits in components-used tables.
- [ ] `grep -n "knowledgeBase\|hubSearch\|policySearch" CLAUDE.md` returns 0 hits in Sprint 4 backlog.
- [ ] `docs/admin-guide.md:221-226` Manager defaults match `SpSearchManagerWebPart.manifest.json:29-32` byte-for-byte.
- [ ] `docs/provisioning-guide.md:131-132` accurately describes the 24h sweep + 90-day retention.
- [ ] `grep -rn "window\.location\.href = " src` returns 0 direct-assignment hits.
- [ ] `grep -rn "spfx-toolkit/lib/utilities/htmlUtils" src` returns ≥1 hit.
- [ ] Local `sanitizeSummaryHtml` deleted from `documentTitleUtils.ts`.
- [ ] Mode toggle at `SpSearchBox.tsx` renders as `<fieldset>` + `<legend>`.
- [ ] Scope selector carries `aria-describedby`.
- [ ] `grep -rn "prefers-reduced-motion" src/styles src/webparts | wc -l` returns ≥6.
- [ ] axe smoke test passes; regression sentinel proves the gate fails on `<img>` without `alt`.
- [ ] `config/package-solution.json` `webApiPermissionRequests` contains the Microsoft Graph scope verified in Task 11 Step 1.
- [ ] `SearchTelemetryConfig` + `SearchTelemetryOptIn` SP lists provision via the extended `Provision-SPSearchLists.ps1`.
- [ ] `solution.developer.mpnId` is empty or a real MPN ID (no `Undefined-` prefix).
- [ ] `package.json:version` = `1.0.0`; `config/package-solution.json:solution.version` = `1.0.0.0`.

# Cross-track downstream consumer reciprocations (do not break)

These bidirectional refs are declared in the audit Roadmap Matrix and must remain intact when downstream T-track plans land:

- **T1.D9** consumes `src/styles/motion.scss` (Found.D6) for surface-by-surface motion application; consumes axe-core CI gate (Found.D6) at zero per-PR cost.
- **T3.D9** consumes the Heft Jest harness (Found.D13) for `disposeStore` regression test + lifecycle smoke harness; `tests/store/lifecycle.test.ts` trail-marker placeholder is the dependency contract.
- **T4.D6** + **T4.D9** consume the `SearchTelemetryConfig` SP list URL field surface (Found.D9) — admin must reach destination URL via either the Admin Manager property pane (T4.D6) or the Pre-Flight tab clickable link (T4.D9).
- **T4.D9** consumes the declared Microsoft Graph `webApiPermissionRequests` scope verified in Task 11 Step 1 (Found.D10) — Pre-Flight tab row goes green when declaration ships.
- **T4.D10** consumes the GitHub Releases `.sppkg` URL (Found.D8) via `Deploy-SPSearchSolution.ps1 -ReleaseArtifactUrl`; consumes the cleaned `solution.developer.*` metadata (Found.D11).
- **T5.D8** consumes Found.D9's transport + storage + privacy notice; T5.D8 supplies the schema + emitter + property pane that this plan's `ITelemetrySignal` placeholder will be replaced by.
- **T5.D9** consumes Found.D9's aggregated opt-in hash list to compute the "Telemetry coverage: X% of users opted in" Coverage tab card.
- **T2.D3** ships in parallel with Found.D4 (saved-search JSON schema validation as the input side; safeNavigate as the output side) — both should land in the same security-review attention window.
