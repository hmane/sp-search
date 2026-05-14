# SP Search — End-User Guide

> A short reference for SharePoint users navigating a page powered by
> SP Search. Hand this URL to users; admins should bookmark
> `docs/admin-guide.md` instead.

If your SharePoint admin has installed SP Search on a page, you'll
see a search box across the top, a results list below, and a filters
sidebar (when configured). This guide covers the user-facing actions
you can take.

## Search basics

Type a query in the search box and press **Enter**. Results appear
below, ordered by SharePoint's relevance ranking. If verticals are
configured, you'll see tabs above the results (e.g. *All*,
*Documents*, *Pages*, *People*) — click a tab to scope your search.

**Tip:** Press `/` anywhere on the page to jump straight to the
search box. Press `?` to see all keyboard shortcuts.

## Open a result

Click the title of any result to open it. Most results open in a
new tab; videos and images open in a side panel for preview.

For Office documents, the **Detail panel** (opens by clicking the
right-arrow icon on a row) gives you:

- A live preview of the document
- Quick **Open** + **Download** buttons
- The author's profile card (click the name to open their Teams chat)
- Last-modified date with hover-tooltip showing the exact time
- A **Previous / Next** arrow pair to step through results without
  closing the panel (or press `Alt+Left` / `Alt+Right`)

## Save a search

Found a useful query you want to re-run later? Click the **Save
search** button in the Search Manager surface. Your saved search
remembers:

- The query text
- Any filters you applied
- The vertical you were on
- The scope (when narrowed past "All")
- The active layout (List / Compact / Grid / etc.)

To restore a saved search, open the **Saved Searches** tab in the
Search Manager and click any row. The search box, filters, and
layout snap back to the saved state.

## Share a search with a coworker

Saved searches can be shared. Open the **Saved Searches** tab in
the Search Manager, find the search you want to share, click the
**Share** icon in the row's action menu, and pick a coworker from
the people picker. They'll see the shared search appear in their
**Saved Searches** tab with a small red badge announcing it.

The badge clears when they dismiss the new-share MessageBar at the
top of the tab. The notification is delivered without email — it
shows up the next time they open the search page (or within 60s if
they're already viewing it).

## Owned vs. Shared with me

The **Saved Searches** and **Collections** tabs filter on ownership:

- **All** — every saved search (or collection) you can see.
- **Owned** — searches/collections you created.
- **Shared with me** — searches/collections someone else created
  and shared with you. Each row in this view shows a
  "Shared by &lt;Name&gt;" badge so you know who the sender was.

The filter is hidden when no one has shared anything with you (the
toggle would only have a single state).

## Collections — pin search results for later

Want to keep specific results from a search rather than the whole
search query? Use **Collections**. Open the Search Manager's
**Collections** tab and create a new collection (give it a name +
optional tags). Then, on a search results page, hover any row's
title and click the **Pin** icon — the result lands in the
selected collection.

Collections persist across sessions and can be shared with
coworkers the same way saved searches can.

## Per-row actions

Hover any row and a **"..."** button appears at the trailing edge
(List view), or click the existing menu indicator in Compact /
DataGrid rows. The menu offers:

- **Open in new tab** — open the result in a fresh browser tab.
- **Download** — download the underlying file (when supported).
- **Copy link** — copy the result's URL to your clipboard.
- *DataGrid only*: **View item**, **Edit item**, and **Delete**
  also appear when you have the appropriate permissions on the
  underlying list item.

To pin a result into a collection, use the **Pin** icon next to
the title (visible on hover).

## Export results

The result toolbar has a **Download** icon that opens an export
menu:

- **Export all to CSV** / **XLSX** — download every result on the
  current page in the layout's column shape.

Both formats include the columns your admin configured for the
active layout. CSV uses UTF-8 with a byte-order mark so Excel on
Windows opens it correctly; XLSX is a real `.xlsx` workbook with
auto-sized columns.

The DataGrid layout also has its own DevExtreme-built export menu
in its toolbar; both paths produce the same shape.

## Keyboard shortcuts

| Keys | Action |
|------|--------|
| `/` | Focus the search box from anywhere on the page |
| `?` | Open this shortcut help (modal) |
| `Esc` | Close the open panel or dialog |
| `Enter` | Open the focused result |
| `Alt+Left` / `Alt+Right` | Previous / next result in the detail panel |

Shortcuts don't fire while you're typing in a text box — `/` types
a slash in the search input, `?` types a question mark, etc.

## What if my search returns nothing?

If you see a "No results" empty state:

1. Check the **Filters** sidebar — you may have filters applied
   that exclude what you're looking for. Click **Clear all** to
   reset.
2. Try a different **Vertical** tab if your admin enabled them —
   the same query may match documents but not people, or vice
   versa.
3. Verify the spelling of your query. The empty state suggests
   broadening if your query is unusual.
4. Some content takes 15-60 minutes to appear in search after
   it's uploaded; very recent uploads may not yet be indexed.

If you're consistently seeing empty results across queries,
contact your SharePoint admin — they have a built-in **Pre-Flight**
diagnostic that surfaces tenant-wide search readiness issues.

## Help and feedback

For installation issues, admin configuration, or feature requests,
contact the admin who deployed SP Search on this page. Bug reports
go to [the GitHub issue tracker](https://github.com/hmane/sp-search/issues).
