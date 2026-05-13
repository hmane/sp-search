# Changelog

All notable changes to SP Search are documented here. Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versioning follows [SemVer 2.0](https://semver.org/).

## [Unreleased]

### Changed

- `package.json:version` aligned to `1.0.0` from generator default `0.0.1`; `config/package-solution.json:solution.version` aligned to `1.0.0.0` (lockstep — Found.D11).
- `solution.developer.mpnId` cleared from `Undefined-1.21.1` (SPFx generator default) to empty string (Found.D11). Populate with a real Partner Center MPN ID once one is registered.
- `solution.developer.websiteUrl / privacyUrl / termsOfUseUrl` populated with canonical project URLs derived from `git remote get-url origin` (Found.D11).
- CI/release tooling standardized on Azure DevOps (the repo is hosted in Azure Repos). `docs/release-policy.md`, `docs/performance-budgets.md`, `docs/accessibility.md`, `CONTRIBUTING.md`, `README.md`, and `docs/release-runs/v1.0.0-rc.1.md` updated to reference the ADO build/release pipeline instead of GitHub Actions.

### Removed

- `.github/` directory removed entirely — the GitHub Actions CI workflow (`build.yml`), release workflow (`release.yml`), Dependabot config (`dependabot.yml`), and GitHub issue/PR templates added in Found.D8. The build/test/bundle gates and `.sppkg` packaging run via the project's Azure DevOps pipeline; dependency review is manual; bugs and PRs use ADO work items.

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
