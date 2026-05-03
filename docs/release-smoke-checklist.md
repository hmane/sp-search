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
