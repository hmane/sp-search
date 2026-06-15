# SP Search

Enterprise SharePoint search solution built as SPFx 1.22 web parts. Replaces PnP Modern Search v4 with a modern React + Zustand architecture, scenario presets, advanced layouts, and an extensible provider/registry model.

## What's in the box

| Web part | Purpose |
|----------|---------|
| **SP Search Box** | Query input with KQL mode, suggestions, scope selector |
| **SP Search Results** | Results display with 6 layouts (DataGrid, Card, List, Compact, People, Gallery), detail panel, row actions |
| **SP Search Filters** | Refinement filters (Checkbox, Date range, Slider, People, Taxonomy Tag Box, Tag Box, Toggle) |
| **SP Search Results + Filters** | Optional full-width wrapper that renders Results and Filters from one web part |
| **SP Search Verticals** | Tab navigation with badge counts |
| **SP Search Manager** | Saved searches, sharing, collections, history |
| **SP Search Admin Manager** | Tenant-wide admin dashboard |

All seven web parts share a single Zustand store via an SPFx Library Component (`spSearchStore`). Multi-instance isolation via `searchContextId`.

## First-time setup

Requirements: **Node 22.14+ (< 23)**, npm 10+, gulp not required (Heft replaces it).

This repo depends on **spfx-toolkit** as a sibling directory (`file:../spfx-toolkit` in `package.json`). Clone both repos as siblings:

```bash
# Pick a workspace directory
mkdir ~/Development && cd ~/Development

# Clone both repos as siblings
git clone https://github.com/dodgeandcox/sp-search.git
git clone https://github.com/dodgeandcox/spfx-toolkit.git

# Build the toolkit first (sp-search consumes its compiled lib/)
cd spfx-toolkit && npm install && npm run build

# Then sp-search
cd ../sp-search && npm install
npm test
npm run type-check      # tsc --noEmit
npm run package         # produces sharepoint/solution/sp-search.sppkg
# upload sharepoint/solution/sp-search.sppkg to your tenant or site app catalog
```

If you put the toolkit at a different relative path, update `"spfx-toolkit": "file:../spfx-toolkit"` in [package.json](package.json).

Tagged releases also publish `sp-search.sppkg` as a build artifact on the project's Azure DevOps pipeline, so you can download it instead of building locally.

> After deploying, approve the Microsoft Graph API access requests shown by the package. `People.Read` powers the People vertical; `User.Read` powers audience targeting through Graph `/me/memberOf`.

## Documentation

- **Admin install + configuration:** [`docs/admin-guide.md`](docs/admin-guide.md)
- **Deployment / packaging:** [`docs/deployment-guide.md`](docs/deployment-guide.md)
- **Provisioning scripts:** [`docs/provisioning-guide.md`](docs/provisioning-guide.md)
- **Extensibility (custom providers / actions / layouts):** [`docs/extensibility-guide.md`](docs/extensibility-guide.md)
- **PnP Modern Search v4 alignment:** [`docs/pnp-modern-search-alignment.md`](docs/pnp-modern-search-alignment.md)
- **Performance budgets:** [`docs/performance-budgets.md`](docs/performance-budgets.md)
- **Release smoke checklist:** [`docs/release-smoke-checklist.md`](docs/release-smoke-checklist.md)
- **Accessibility statement (WCAG 2.1 AA):** [`docs/accessibility.md`](docs/accessibility.md)
- **Privacy notice (telemetry):** [`docs/privacy-notice.md`](docs/privacy-notice.md)
- **Release policy (SemVer + lockstep):** [`docs/release-policy.md`](docs/release-policy.md)
- **Changelog:** [`CHANGELOG.md`](CHANGELOG.md)
- **Contributing:** [`CONTRIBUTING.md`](CONTRIBUTING.md)
- **Architecture / developer reference:** [`CLAUDE.md`](CLAUDE.md)

## Working with Claude Code

This repo is set up to bootstrap [Claude Code](https://claude.com/claude-code) on a fresh clone:

- **[`CLAUDE.md`](CLAUDE.md)** at the repo root is the authoritative architecture + rules reference. It is loaded automatically into every Claude Code session.
- **[`.claude/agents/`](.claude/agents/)** ships 7 specialized subagent definitions Claude can invoke for focused tasks:
  - `webpart-builder` — SPFx web part class + property pane + onInit/render plumbing
  - `store-architect` — Zustand store, orchestrator, URL sync, registries
  - `search-provider` — `ISearchDataProvider`, `SearchService`, `TokenService`
  - `layout-builder` — 6 result layouts, cell renderers, detail panel, per-row ECB
  - `filter-builder` — 7 filter types, formatters, pill bar, special-field handling
  - `search-manager` — saved searches, sharing, history, admin tabs, hidden lists
  - `testing` — jest test patterns, mock fixtures, ignored-test debt
- **[`.claude/settings.json`](.claude/settings.json)** is the org-shared Claude permission set (committed). Personal-machine permissions live in `.claude/settings.local.json` which is gitignored.

When you open this repo in Claude Code, ask it to "summarize the architecture" or "check what the search runtime path looks like" — it has full context immediately. For multi-step audit-style work, ask it to "dispatch parallel agents to audit X, Y, Z."

## Tech stack

SPFx 1.22.2 . React 17.0.1 . TypeScript 5.3.3 . Zustand 4.x . PnPjs 3.x . Microsoft Graph Client . Fluent UI v8.106 . DevExtreme 22.2 . spfx-toolkit (Card, ErrorBoundary, hooks, utilities)

## License

See `LICENSE` (if present) or contact the project owner.
