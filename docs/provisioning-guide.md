# SP Search — Provisioning Script Documentation

The `Provision-SPSearchLists.ps1` script creates and configures the 3 hidden SharePoint lists required by SP Search.

---

## Prerequisites

| Requirement | Details |
|-------------|---------|
| **PnP.PowerShell** | `Install-Module PnP.PowerShell -Scope CurrentUser` |
| **Permissions** | Site Collection Admin or Full Control on the target site |
| **PowerShell** | 5.1+ (Windows) or 7.x (cross-platform) |

---

## Usage

### Basic

```powershell
.\scripts\Provision-SPSearchLists.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/search"
```

### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `-SiteUrl` | String | **Yes** | — | Full URL of the target SharePoint site collection. |

The script uses `-Interactive` authentication via `Connect-PnPOnline`, which opens a browser window for sign-in.

---

## What Gets Provisioned

### List 1: SearchSavedQueries

Stores saved/shared searches and state snapshots.

| Column | Type | Indexed | Description |
|--------|------|---------|-------------|
| Title | Text (default) | Yes | Search name / label |
| QueryText | Note | No | The search query text |
| SearchState | Note | No | Full Zustand state JSON |
| SearchUrl | URL | No | Deep link URL |
| EntryType | Choice | Yes | `SavedSearch`, `SharedSearch`, or `StateSnapshot` |
| Category | Text | Yes | User-defined category |
| SharedWith | UserMulti | No | Users this search is shared with |
| ResultCount | Number | No | Result count at time of save |
| LastUsed | DateTime | Yes | Last time this search was loaded |
| ExpiresAt | DateTime | Yes | Expiration time for state snapshots |

**Permissions:** Inherits from parent + all authenticated users can add items. Item-level permissions are set programmatically when sharing (breakRoleInheritance per item).

---

### List 2: SearchHistory

Stores per-user search history. **This list will exceed 5,000 items** — all indexes are critical for list view threshold compliance.

| Column | Type | Indexed | Description |
|--------|------|---------|-------------|
| Title | Text (default) | No | Query text summary |
| Author | Person (default) | **Yes** | **CRITICAL** — must be first CAML predicate |
| QueryHash | Text | Yes | SHA-256 hash for deduplication |
| Vertical | Text | Yes | Active vertical key at search time |
| Scope | Text | No | Active scope at search time |
| SearchState | Note | No | Full query state JSON |
| ResultCount | Number | No | Number of results returned |
| ClickedItems | Note | No | JSON array of clicked result tracking |
| SearchTimestamp | DateTime | Yes | **CRITICAL** — for date-range filtering |

**Permissions:** Item-level security enabled (`ReadSecurity=2, WriteSecurity=2`). Users can only read and edit their own items — prevents visibility of others' search history.

**Index Validation:** The script verifies that Author and SearchTimestamp indexes were successfully created. If either fails, the script exits with an error rather than continuing silently, because missing indexes will cause list view threshold failures at scale.

---

### List 3: SearchCollections

Stores pinboard collections of search results.

| Column | Type | Indexed | Description |
|--------|------|---------|-------------|
| Title | Text (default) | Yes | Collection name |
| ItemUrl | URL | No | URL of the pinned result |
| ItemTitle | Text | No | Title of the pinned result |
| ItemMetadata | Note | No | Cached metadata JSON |
| CollectionName | Text | Yes | Parent collection identifier |
| Tags | Note | No | User-applied tags (JSON array) |
| SharedWith | UserMulti | No | Users this collection is shared with |
| SortOrder | Number | No | Manual sort position within collection |

**Permissions:** Inherits from parent + all authenticated users can add items.

---

## Idempotency

The script is **fully safe to re-run**:

| Scenario | Behavior |
|----------|----------|
| List already exists | Skips creation, shows `[EXISTS]` in yellow |
| Column already exists | Skips creation, shows `[EXISTS]` in yellow |
| Index already exists | Checks if field is indexed; skips if yes |
| Permissions already set | `Set-PnPList` is idempotent |

---

## Post-Provisioning Checklist

1. **Verify SearchHistory indexes**
   - The script validates these automatically
   - If you see a yellow warning, check manually via SharePoint list settings

2. **Test basic functionality**
   - Add SP Search web parts to a page
   - Perform a search — history entry should be logged
   - Save a search — entry should appear in SearchSavedQueries

3. **Configure promoted results (optional)**
   - Use SharePoint Query Rules (Search admin) to create promoted results
   - See [admin-guide.md](./admin-guide.md) for guidance

4. **Configure history cleanup (optional)**
   - Call `cleanupHistory(ttlDays)` via the SearchManagerService API to delete entries older than the specified TTL
   - This is a manual operation — there is no automatic background cleanup

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `PnP.PowerShell module not found` | `Install-Module PnP.PowerShell -Scope CurrentUser` |
| `403 Forbidden on list creation` | Ensure you're a Site Collection Admin |
| `Index creation failed` | Check if the site has hit the index limit (20 per list). Remove unused indexes. |
| `SearchHistory threshold errors at runtime` | Verify Author + SearchTimestamp indexes exist. Re-run the script to recreate if missing. |
