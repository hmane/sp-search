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

> After deploying, approve the `Microsoft Graph: People.Read` API access request in the SharePoint admin center for the People vertical to function (declared in `webApiPermissionRequests` per Found.D10).

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

## Tech stack

SPFx 1.22.2 . React 17.0.1 . TypeScript 5.3.3 . Zustand 4.x . PnPjs 3.x . Microsoft Graph Client . Fluent UI v8.106 . DevExtreme 22.2 . spfx-toolkit (Card, ErrorBoundary, hooks, utilities)

## License

See `LICENSE` (if present) or contact the project owner.
