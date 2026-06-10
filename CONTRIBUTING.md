# Contributing to SP Search

## Architecture

Read `CLAUDE.md` first — it is the authoritative source for architecture, conventions, web-part responsibilities, and import rules. The `docs/` directory has admin-facing guides; `CLAUDE.md` is developer-facing.

## Setup

This repo consumes `spfx-toolkit` via `file:../spfx-toolkit` — clone both repos as siblings before installing.

```bash
# In your workspace directory:
git clone https://github.com/dodgeandcox/sp-search.git
git clone https://github.com/dodgeandcox/spfx-toolkit.git

# Build the toolkit first
cd spfx-toolkit && npm install && npm run build

# Then this repo
cd ../sp-search && npm install
npm test              # Heft Jest pipeline
npm run type-check    # tsc --noEmit
npm run package       # produces sharepoint/solution/sp-search.sppkg
```

Requirements: Node 22.14+ (< 23), npm 10+. Heft replaces gulp.

## Branching

- `main` — released code; protected
- `feat/<short-name>` — feature branches; squash-merged into `main` via PR
- `fix/<short-name>` — bug fix branches; squash-merged

## Commit messages

Use Conventional Commits prefix: `feat`, `fix`, `docs`, `test`, `build`, `perf`, `refactor`, `chore`. Example:

```
feat(filters): add SliderFilter for numeric refiners

Extends FilterTypeRegistry with SliderFilter type. Wires into
SearchFilters drawer; renders devextreme-react Slider lazy-loaded.
```

The pre-GA convention of tagging commits with Roadmap IDs (T1.D5, MISS-001, etc.) is retired — that audit is closed. Reference an issue or PR ID where helpful.

## Testing

- All store / service / utility code lives under `tests/{store,services,utils}/`
- Run `npm test -- --testPathPattern <pattern>` to filter
- Component tests use jest-axe for accessibility smoke (`tests/a11y/smokeAxe.test.tsx` — Found.D6)

## Pre-merge checklist

Run `docs/release-smoke-checklist.md` before merging to `main`. The build pipeline runs the build + test + bundle-gate steps (`npm ci` → `npm run type-check` → `npm test` → `npm run package` → `npm run check:bundles`); the tenant-upload smoke (Step 6) and multi-context smoke (Step 7) are manual.

## Releases

See `docs/release-policy.md` for SemVer policy and tag conventions. The release pipeline builds production and publishes `sp-search.sppkg` as a pipeline artifact for the tagged commit.

## Performance budgets

Per-web-part byte budgets are enforced by `scripts/check-bundle-sizes.js`. PRs that breach the budget fail the build pipeline; raising a budget requires explicit approval. See `docs/performance-budgets.md`.

## Accessibility

axe-core smoke tests run in the build pipeline. PRs that introduce new violations fail the build. See `docs/accessibility.md` for the WCAG 2.1 AA scope.

## Working with Claude Code

If you use [Claude Code](https://claude.com/claude-code) in this repo: see the **Working with Claude Code** section of [`README.md`](README.md) for what the repo ships (CLAUDE.md, 7 subagent definitions under `.claude/agents/`, shared permission settings). Personal-machine permissions go in `.claude/settings.local.json` (gitignored) — don't commit them.
