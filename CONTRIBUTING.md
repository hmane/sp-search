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
