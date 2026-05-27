# Design — Result link behaviour (Stream C / #7)

> Status: approved 2026-05-12.
> Scope: new property-pane config on the Results web part + a link-resolution util + a small `selectedProperties` baseline addition + wiring across the five result layouts. No store-schema break; defaults preserve today's behaviour.

## Problem

Clicking a result's title navigates via a plain `<a href={getResultAnchorProps(item).href} target="_blank" rel="noopener noreferrer">` ([documentTitleUtils.ts:335](src/webparts/spSearchResults/components/documentTitleUtils.ts#L335)) — always a new tab, always the raw `item.url`, with **no admin control**. Admins want: open in same tab / new tab / the detail panel; for documents, open the file vs its properties form; for list items, open the display form vs the edit form. (`OpenAction` always `window.open(_blank)`; `PreviewAction` opens the existing `ResultDetailPanel` side-drawer, gated by `enablePreviewPanel`. Verticals already have an `openBehavior: 'currentTab' | 'newTab'` per-vertical precedent — [IVerticalDefinition.ts:32](src/libraries/spSearchStore/interfaces/IVerticalDefinition.ts#L32).)

## Decision

A new **"Result link behaviour"** section on the `SpSearchResults` web-part property pane. Per-content-type "what does a click resolve to", plus a global "how does it open". Out-of-box behaviour is unchanged.

### Config (new web-part properties)

| Property | Control | Values (default first) | Effect |
|----------|---------|------------------------|--------|
| `resultClickTarget` | `PropertyPaneChoiceGroup` | `panel` *(default — preserves today's behaviour)* · `newTab` · `sameTab` · `sidePanel` | How a title click opens. `panel` → keep today's `DocumentTitleHoverCard` centred preview Modal for previewable files (PDF/Office/image/text via `buildPreviewUrl`), navigate (new tab) for non-previewable. `newTab` / `sameTab` → always navigate (sets `<a target>`), including for previewable files (suppresses the Modal). `sidePanel` → intercept the click and `setPreviewItem(item)` to open `ResultDetailPanel` for every result (forces `enablePreviewPanel` true). |
| `documentLinkMode` | `PropertyPaneDropdown` (visible when `resultClickTarget !== 'panel'`) | `file` *(default)* · `propertiesForm` | For results classified **Document**: `file` → `buildBrowserOpenUrl(item)` (already exists — `ServerRedirectedURL` when present so Office docs open in the web app, else the file URL); `propertiesForm` → the item's DispForm. |
| `listItemLinkMode` | `PropertyPaneDropdown` (visible when `resultClickTarget !== 'panel'`) | `displayForm` *(default)* · `editForm` | For results classified **ListItem**: `displayForm` → `item.url` (search already returns a list item's URL as its DispForm); `editForm` → the item's EditForm. |

Results classified **Page**, **Site**, **Folder**, or **Other** (external-connector items, people, …) always navigate to `item.url` — there is no "form" alternative for them, so no per-type control. No "edit the page" mode in scope.

When `resultClickTarget === 'panel'`, `documentLinkMode` / `listItemLinkMode` are irrelevant (the panel just shows the item) — the property-pane controls for them are hidden in that mode.

### Content-type classification

The provider's baseline `selectedProperties` (`DEFAULT_SELECTED_PROPERTIES` in `src/libraries/spSearchStore/services/SearchService.ts`) already requests `contentclass`, `FileExtension`, `ServerRedirectedURL`, `IsDocument`, `NormListID`, and `SPSiteURL`. It gains `ListId` and `ListItemID` — required for `buildFormUrl(item, pageType)` (already exists at [documentTitleUtils.ts:377](src/webparts/spSearchResults/components/documentTitleUtils.ts#L377)) to resolve. (Side benefit: this also fixes the dead "View / Edit item" links inside `DocumentTitleHoverCard` — they call `buildFormUrl` today but return `undefined` because the props aren't requested.) `classifyResult(item)` → `'document' | 'page' | 'listItem' | 'site' | 'folder' | 'other'`:

1. `FileExtension === 'aspx'` **or** `ContentClass === 'STS_ListItem_WebPageLibrary'` → `page`.
2. else non-empty `FileExtension` **or** `ContentClass` starts with `STS_ListItem_DocumentLibrary` **or** `ContentClass === 'STS_Document'` → `document`.
3. else `ContentClass === 'STS_Site'` or `'STS_Web'` → `site`.
4. else `ContentClass === 'STS_List_DocumentLibrary'`/`'STS_List'` (a container, not an item) → `folder`.
5. else `ContentClass` starts with `'STS_ListItem'` → `listItem`.
6. else → `other`.

Anything unclassified, or a form URL that can't be built (missing `ListID`/`ListItemID`), falls back to "navigate to `item.url`" — always safe.

### URL construction

- **Document → file**: `buildBrowserOpenUrl(item)` (existing in `documentTitleUtils.ts`).
- **Document → propertiesForm** / **ListItem → displayForm**: `displayForm` for a list item is just `item.url`; for a document it's `{item.siteUrl}/_layouts/15/listform.aspx?PageType=4&ListId={ListID}&ID={ListItemID}`.
- **ListItem / Document → editForm**: `{item.siteUrl}/_layouts/15/listform.aspx?PageType=6&ListId={ListID}&ID={ListItemID}`. (Exact `PageType` values — 4 = display, 6 = edit — verified against current SharePoint during implementation; if `listform.aspx` proves unreliable, fall back to deriving the list's form URL from `ParentLink`/`Path`.)

All constructed URLs are https or root-relative (built from `item.siteUrl` + `/_layouts/15/...` or from `item.url`) and are used as the `<a href>`. There is no programmatic `window.location` navigation in this path — navigation happens via the anchor, and `panel` mode `preventDefault()`s it — so no `safeNavigate` call is added; the existing `safeNavigate`-guarded paths (reset, etc.) are untouched.

### Wiring

New util `src/webparts/spSearchResults/components/resultLink.ts`:
- `classifyResult(item: ISearchResult): ResultKind` — pure.
- `resolveResultLink(item: ISearchResult, cfg: { clickTarget; documentLinkMode; listItemLinkMode }): { href: string; target?: string; rel?: string; openInPanel: boolean }` — pure; replaces `getResultAnchorProps`. `openInPanel` true iff `clickTarget === 'panel'`. `target`/`rel` set from `clickTarget` (`newTab` → `_blank` + `noopener noreferrer`; `sameTab` → undefined).

All five layouts (`ListLayout`, `CompactLayout`, `CardLayout`, `PeopleLayout`, `GalleryLayout`, plus `DataGridContent`'s title cell) call `resolveResultLink` instead of `getResultAnchorProps`, and pass `clickTarget` down to `DocumentTitleHoverCard` (new prop). `DocumentTitleHoverCard.handleClick` is updated:

- `clickTarget = 'panel'` → unchanged (preserve today's Modal for previewable files; let the `<a>` navigate for non-previewable).
- `clickTarget = 'newTab'` / `'sameTab'` → always let the `<a>` navigate (skip the Modal even for previewable files).
- `clickTarget = 'sidePanel'` → `e.preventDefault()` + `store.getState().setPreviewItem(item)` to open `ResultDetailPanel` (the side-panel `setPreviewItem` action — the component reads from the store, so this works for every layout).

History-click logging (`orchestrator.logClickedItem` via `onItemClick`) fires in all cases. The `<a>` stays a real `<a href>` for keyboard accessibility — middle-click / Ctrl-click still open the resolved URL in a new tab even in `panel` / `sidePanel` mode, which is the expected browser behaviour.

The web part passes the three config values down through `ISpSearchResultsProps` → `SpSearchResults` → the layout components (alongside the existing `titleDisplayMode` etc.).

### Action toolbar — unchanged

`OpenAction` keeps `window.open(_blank)`; `PreviewAction` keeps `setPreviewItem`. They are explicit user actions; the new config governs only the default title-click. (If wanted later, `OpenAction` could read `resultClickTarget` — out of scope now.)

### Defaults preserve today's behaviour

`resultClickTarget = 'panel'`, `documentLinkMode = 'file'`, `listItemLinkMode = 'displayForm'` → preserves today's behaviour byte-for-byte: previewable files open in the existing centred preview Modal (via `DocumentTitleHoverCard`); non-previewable files navigate to `item.url` in a new tab. Manifest `preconfiguredEntries` get the three new properties at these defaults; existing pages pick them up as the same defaults.

## Verification

- `npm run type-check` clean; `npm test` green with **new unit tests**: `tests/.../resultLink.test.ts` covering `classifyResult` (each ContentClass/extension case) and `resolveResultLink` (each `clickTarget` × each per-type mode, plus the missing-`ListID` fallback). Both are pure functions.
- `npm run package` (no Sass/TS warnings) + `npm run check:bundles` green (results bundle delta: a few KB for the util + property-pane controls; well within budget).
- Edit-mode property pane renders the new section; the per-type dropdowns hide when `resultClickTarget === 'panel'`.

## Out of scope / follow-ups

- "Edit the page" mode for page results; making `OpenAction` respect `resultClickTarget`.
- Stream C #8 (image preview) and Stream B (#2/#3/#4) — separate.
- Per-vertical override of result link behaviour (verticals only override `dataProviderId`/layout today) — not now.
