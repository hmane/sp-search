# SP Search — Release Policy

> Owned by Foundations track (Found.D8). Authoritative source for version-bump decisions, lockstep convention between `package.json` and `config/package-solution.json`, and tag/release naming.

## SemVer 2.0

- **Major (`x.0.0`)** — breaking changes to web-part property pane fields, store API surface, ISearchDataProvider contract, IFilterConfig schema, or registries (DataProvider, Suggestion, Action, Layout, FilterType). Anything that requires admins to re-author saved searches or property-pane configurations counts.
- **Minor (`1.x.0`)** — new features that do not break existing configurations: new layouts, new filter types, new scenario presets, new actions, new admin dashboard tabs.
- **Patch (`1.0.x`)** — bug fixes, performance improvements, documentation updates, dependency bumps that do not change the public surface.

Pre-release tags: `1.0.0-rc.N` for release candidates, `1.0.0-beta.N` for unstable previews. Promote rc.N to release by tagging the rc.N commit as `1.0.0` (no code changes between).

## Lockstep convention

`package.json:3` `version` and `config/package-solution.json:6` `solution.version` move in lockstep at every release. The `solution.version` field uses 4-part SharePoint-required `M.M.P.B` (build number always `0` unless an in-tag rebuild is required); `package.json` uses 3-part SemVer.

| package.json | solution.version |
|--------------|-------------------|
| `1.0.0`      | `1.0.0.0` |
| `1.0.1`      | `1.0.1.0` |
| `1.1.0`      | `1.1.0.0` |
| `2.0.0`      | `2.0.0.0` |

Both files MUST be bumped in the same commit. CI at `.github/workflows/build.yml` validates the lockstep relationship and fails on mismatch.

> **Tranche-2 transient state:** between Found.D8 (this doc) landing and Found.D11 closing, `package.json` may still read `0.0.1` while `config/package-solution.json` reads `1.0.13.0`. The lockstep gate will fail until D11 ships, by design — that surfaces the alignment debt rather than hiding it.

## Tag + Release naming

- Tag format: `v<semver>` (e.g. `v1.0.0`, `v1.0.0-rc.1`)
- Release title: `v<semver> — <one-line summary>`
- Release body: pulled from the matching `## [<semver>]` section of `CHANGELOG.md`
- Release artifact: `sp-search.sppkg` from the `release.yml` workflow build

## Dependency policy

- Production deps (`dependencies` in `package.json`) — manual review per PR; no automatic version bumps.
- Dev deps (`devDependencies`) — Dependabot-managed weekly per `.github/dependabot.yml`; auto-merge disabled (manual review).
- SPFx core (`@microsoft/sp-*`) — manual; coordinate version bumps with a documented Heft / spfx-toolkit compatibility check.
