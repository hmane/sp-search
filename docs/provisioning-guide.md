# SP Search — Provisioning Script Documentation

The `Provision-SPSearchLists.ps1` script creates and configures the three core hidden SharePoint lists required by SP Search plus the optional telemetry lists used only when telemetry is explicitly enabled.

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
| Author | Person (default) | Yes | Owner filter for per-user queries |
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
| QueryText | Note | No | Full query text shown in Search Manager history and suggestions |
| QueryHash | Text | Yes | SHA-256 hash for deduplication |
| Vertical | Text | Yes | Active vertical key at search time |
| SearchPageUrl | Text | No | Page where the query was executed |
| SearchState | Note | No | Full query state JSON |
| UseCount | Number | No | Number of times the same query state was reused |
| ResultCount | Number | No | Number of results returned |
| IsZeroResult | Boolean | No | True when a search returned zero results. Used by the Health panel. |
| ClickedItems | Note | No | JSON array of clicked result tracking |
| SearchTimestamp | DateTime | Yes | **CRITICAL** — for date-range filtering |

**Permissions:** Item-level security enabled (`ReadSecurity=2, WriteSecurity=2`). Users can only read and edit their own items — prevents visibility of others' search history.

**Index Validation:** The script verifies that Author and SearchTimestamp indexes were successfully created. If either fails, the script exits with an error rather than continuing silently, because missing indexes will cause list view threshold failures at scale.

**Admin analytics note:** User-facing history queries still filter by `Author` first, but the admin Health and Insights views intentionally query cross-user data by `SearchTimestamp` first so they stay index-friendly above 5,000 items.

---

### List 3: SearchCollections

Stores pinboard collections of search results.

| Column | Type | Indexed | Description |
|--------|------|---------|-------------|
| Title | Text (default) | Yes | Collection name |
| Author | Person (default) | Yes | Owner filter for per-user queries |
| ItemUrl | URL | No | URL of the pinned result |
| ItemTitle | Text | No | Title of the pinned result |
| ItemMetadata | Note | No | Cached metadata JSON |
| CollectionName | Text | Yes | Parent collection identifier |
| Tags | Note | No | User-applied tags (JSON array) |
| SharedWith | UserMulti | No | Users this collection is shared with |
| SortOrder | Number | No | Manual sort position within collection |

**Permissions:** Inherits from parent + all authenticated users can add items.

---

### List 4: SearchTelemetryConfig

Stores optional telemetry configuration. Telemetry is disabled by default and dormant unless an admin explicitly enables it.

| Column | Type | Indexed | Description |
|--------|------|---------|-------------|
| Title | Text (default) | No | Config row label |
| IsEnabled | Boolean | No | Tenant-level telemetry enablement flag |
| DestinationEndpoint | Text | No | HTTPS endpoint for telemetry POSTs |
| BatchIntervalSeconds | Number | No | Flush interval |
| BatchSizeMax | Number | No | Maximum signals per batch |
| PrivacyAcknowledgedBy | Text | No | Admin identifier for privacy acknowledgement |
| PrivacyAcknowledgedAt | DateTime | No | Acknowledgement timestamp |

---

### List 5: SearchTelemetryOptIn

Stores optional per-user telemetry consent records. No active runtime code currently writes this list unless telemetry is wired/enabled.

| Column | Type | Indexed | Description |
|--------|------|---------|-------------|
| Title | Text (default) | No | Consent row label |
| ConsentTimestamp | DateTime | No | Consent timestamp |
| ConsentVersion | Text | No | Privacy/consent version |
| AnonHash | Text | No | Non-reversible anonymized user/session hash |

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
   - SearchManagerService runs an automatic 24-hour cleanup sweep retaining the last 90 days of history per `HISTORY_RETENTION_DAYS = 90` (`src/libraries/spSearchStore/services/SearchManagerService.ts:735-739`). The manual `cleanupHistory` call is supplemental and can shorten the retention window for one-off purges.

5. **Provision scenario pages (optional)**
   - Use `scripts/Search-ScenarioPresets.ps1` to create fully-configured starter pages for built-in scenarios such as `documents`, `knowledge-base`, or `policy-search`
   - See [deployment-guide.md](./deployment-guide.md) for example commands

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `PnP.PowerShell module not found` | `Install-Module PnP.PowerShell -Scope CurrentUser` |
| `403 Forbidden on list creation` | Ensure you're a Site Collection Admin |
| `Index creation failed` | Check if the site has hit the index limit (20 per list). Remove unused indexes. |
| `SearchHistory threshold errors at runtime` | Verify Author + SearchTimestamp indexes exist. Re-run the script to recreate if missing. |
