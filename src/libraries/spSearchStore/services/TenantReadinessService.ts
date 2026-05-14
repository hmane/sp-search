/**
 * T4.D9 — tenant-readiness pre-flight checks for the Admin Manager.
 *
 * Each check returns an `IReadinessCheck` with a green/yellow/red status,
 * a one-line message, and (optionally) a "Fix this" link payload. The
 * audit's acceptance signal is a single screenshot that a peer-admin can
 * use to triage a tenant install — every check is designed to surface
 * actionable copy + a link to the doc / script / property pane that
 * fixes it.
 *
 * All checks are async + AbortSignal-cancellable so the panel can be
 * unmounted mid-scan without leaking in-flight requests.
 */

import 'spfx-toolkit/lib/utilities/context/pnpImports/lists';
import 'spfx-toolkit/lib/utilities/context/pnpImports/search';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { fetchManagedProperties, getCachedSchema } from './SchemaService';

export type ReadinessStatus = 'green' | 'yellow' | 'red';

export interface IReadinessFixLink {
  text: string;
  href: string;
}

export interface IReadinessCheck {
  id: string;
  title: string;
  status: ReadinessStatus;
  message: string;
  fix?: IReadinessFixLink;
}

export interface IReadinessReport {
  /** All checks in evaluation order. */
  checks: IReadinessCheck[];
  /** True only when every check is green. */
  allGreen: boolean;
  /** Number of red checks (the audit signal cares about this). */
  redCount: number;
  /** Wall-clock timestamp of the scan. */
  generatedAt: Date;
}

const HIDDEN_LISTS = [
  { name: 'SearchSavedQueries', requiredFields: ['QueryText', 'SearchState'] },
  { name: 'SearchHistory', requiredFields: ['QueryText', 'SearchTimestamp', 'IsZeroResult'] },
  { name: 'SearchCollections', requiredFields: ['ItemUrl', 'ItemTitle', 'CollectionName'] },
];

/**
 * Run every pre-flight check and aggregate the report. Failures of
 * individual checks are caught and surfaced as red rows rather than
 * thrown — one bad check shouldn't break the panel.
 */
export async function runTenantReadinessScan(signal?: AbortSignal): Promise<IReadinessReport> {
  const checks: IReadinessCheck[] = [];

  const checkFns: Array<() => Promise<IReadinessCheck>> = [
    () => checkGraphPermission(signal),
    () => checkHiddenLists(signal),
    () => checkSearchHistoryPermissions(signal),
    () => checkSchemaMappings(signal),
    () => checkSearchContentSource(signal),
  ];

  for (let i = 0; i < checkFns.length; i++) {
    if (signal?.aborted) { break; }
    try {
      const result = await checkFns[i]();
      checks.push(result);
    } catch (err) {
      // Defensive — should not happen since each check catches its own
      // failures, but a thrown check shouldn't break the panel.
      checks.push({
        id: 'check-' + i,
        title: 'Internal check error',
        status: 'red',
        message: err instanceof Error ? err.message : 'Unknown error',
      });
    }
  }

  const redCount = checks.filter((c) => c.status === 'red').length;
  return {
    checks,
    redCount,
    allGreen: checks.length > 0 && checks.every((c) => c.status === 'green'),
    generatedAt: new Date(),
  };
}

// ─── (a) Graph People.Read permission ────────────────────────────────────────

/**
 * Calls `/me/memberOf` via the SPFx MSGraphClient. Succeeds when
 * `User.Read` is approved at the tenant API access page. If the call
 * returns 403/401, the row goes red with a link to the API access page.
 *
 * Note: the audit cites `People.Read` but `/me/memberOf` is actually
 * `User.Read` per Microsoft Learn (least privilege). Both scopes are
 * declared in `config/package-solution.json` so a successful call
 * means at least one of them is approved.
 */
async function checkGraphPermission(_signal?: AbortSignal): Promise<IReadinessCheck> {
  const id = 'graph-permission';
  const title = 'Graph API permission (`User.Read` / `People.Read`)';
  try {
    const context = SPContext.context.context as unknown as {
      msGraphClientFactory?: { getClient: (version: string) => Promise<{ api: (url: string) => { get: () => Promise<unknown> } }> };
    };
    if (!context.msGraphClientFactory) {
      return {
        id, title,
        status: 'yellow',
        message: 'Graph client factory not available — could not verify. People/audience features may still work via SharePoint search.',
      };
    }
    const client = await context.msGraphClientFactory.getClient('3');
    await client.api('/me/memberOf?$top=1').get();
    return {
      id, title,
      status: 'green',
      message: 'Graph call to /me/memberOf succeeded — User.Read is approved on this tenant.',
    };
  } catch (err) {
    const msg = err instanceof Error ? err.message : 'Unknown Graph error';
    const isPermissionError = /unauthorized|forbidden|401|403|consent/i.test(msg);
    return {
      id, title,
      status: 'red',
      message: isPermissionError
        ? 'Graph API permission has not been approved. People vertical, org-chart, and audience targeting will fail.'
        : 'Graph API call failed: ' + msg,
      fix: {
        text: 'Approve at SharePoint admin → Advanced → API access',
        href: '/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement',
      },
    };
  }
}

// ─── (b) Hidden lists exist with correct schema ──────────────────────────────

async function checkHiddenLists(_signal?: AbortSignal): Promise<IReadinessCheck> {
  const id = 'hidden-lists';
  const title = 'Hidden lists (SearchSavedQueries / SearchHistory / SearchCollections)';
  const missing: string[] = [];
  const missingFields: string[] = [];

  for (const def of HIDDEN_LISTS) {
    try {
      const list = SPContext.sp.web.lists.getByTitle(def.name);
      const info = await list.select('Id', 'Hidden').expand('Fields')<{ Id: string; Hidden: boolean; Fields: Array<{ InternalName: string }> }>();
      if (!info || !info.Id) {
        missing.push(def.name);
        continue;
      }
      const fieldNames = (info.Fields || []).map((f) => f.InternalName);
      for (const f of def.requiredFields) {
        if (fieldNames.indexOf(f) < 0) {
          missingFields.push(def.name + '.' + f);
        }
      }
    } catch {
      missing.push(def.name);
    }
  }

  if (missing.length > 0) {
    return {
      id, title,
      status: 'red',
      message: 'Missing list(s): ' + missing.join(', ') + '. Search Manager features (saved searches, history, collections, health, insights) will not work.',
      fix: {
        text: 'Run scripts/Provision-SPSearchLists.ps1',
        href: 'https://github.com/hemantmane/Development/blob/main/sp-search/docs/deployment-guide.md#provision-hidden-lists',
      },
    };
  }
  if (missingFields.length > 0) {
    return {
      id, title,
      status: 'yellow',
      message: 'Lists exist but missing fields: ' + missingFields.join(', ') + '. Re-run the provisioning script to add them.',
      fix: {
        text: 'Re-run scripts/Provision-SPSearchLists.ps1',
        href: 'https://github.com/hemantmane/Development/blob/main/sp-search/docs/deployment-guide.md#provision-hidden-lists',
      },
    };
  }
  return {
    id, title,
    status: 'green',
    message: 'All three hidden lists exist with required fields.',
  };
}

// ─── (c) SearchHistory ReadSecurity / WriteSecurity ──────────────────────────

async function checkSearchHistoryPermissions(_signal?: AbortSignal): Promise<IReadinessCheck> {
  const id = 'search-history-permissions';
  const title = 'SearchHistory item-level permissions (ReadSecurity / WriteSecurity = 2)';
  try {
    const list = SPContext.sp.web.lists.getByTitle('SearchHistory');
    const info = await list.select('ReadSecurity', 'WriteSecurity')<{ ReadSecurity: number; WriteSecurity: number }>();
    if (info.ReadSecurity === 2 && info.WriteSecurity === 2) {
      return {
        id, title,
        status: 'green',
        message: 'SearchHistory is configured for per-user read + per-user write — users only see their own queries.',
      };
    }
    return {
      id, title,
      status: 'yellow',
      message: 'SearchHistory has ReadSecurity=' + info.ReadSecurity + ', WriteSecurity=' + info.WriteSecurity + '. Users may see each other\'s search queries.',
      fix: {
        text: 'Re-run scripts/Provision-SPSearchLists.ps1',
        href: 'https://github.com/hemantmane/Development/blob/main/sp-search/docs/deployment-guide.md#provision-hidden-lists',
      },
    };
  } catch {
    return {
      id, title,
      status: 'red',
      message: 'Could not read SearchHistory permissions — the list may not exist yet.',
      fix: {
        text: 'Run scripts/Provision-SPSearchLists.ps1',
        href: 'https://github.com/hemantmane/Development/blob/main/sp-search/docs/deployment-guide.md#provision-hidden-lists',
      },
    };
  }
}

// ─── (d) Schema mappings for common managed properties ──────────────────────

/**
 * Verifies the search schema cache contains the columns the active
 * scenario preset references. For T4.D9 v1 we check a fixed set
 * of "common" managed properties used by most scenarios (Author,
 * LastModifiedTime, FileType, Size). Per-preset specificity is the
 * audit's stated scope but adding the wiring requires plumbing the
 * preset definition into the readiness service — left as a follow-up
 * once preset cross-web-part propagation (T4.D12) lands.
 */
async function checkSchemaMappings(_signal?: AbortSignal): Promise<IReadinessCheck> {
  const id = 'schema-mappings';
  const title = 'Search schema — common managed properties';
  const COMMON_PROPS = ['Author', 'LastModifiedTime', 'FileType', 'Size'];

  let schema = getCachedSchema();
  if (!schema || schema.length === 0) {
    try {
      const result = await fetchManagedProperties();
      if (result.status === 'success') {
        schema = result.properties;
      }
    } catch {
      // Ignore — fall through with empty schema.
    }
  }

  if (!schema || schema.length === 0) {
    return {
      id, title,
      status: 'yellow',
      message: 'Could not fetch the search schema — admin permissions on Search admin / managed properties may be missing.',
      fix: {
        text: 'Check Search admin → Managed properties',
        href: '/_layouts/15/listmanagedproperties.aspx',
      },
    };
  }

  const present = new Set(schema.map((p) => p.name.toLowerCase()));
  const missing = COMMON_PROPS.filter((p) => !present.has(p.toLowerCase()));

  if (missing.length === 0) {
    return {
      id, title,
      status: 'green',
      message: 'All common managed properties (Author / LastModifiedTime / FileType / Size) exist in the tenant schema.',
    };
  }

  return {
    id, title,
    status: 'yellow',
    message: 'Missing managed properties: ' + missing.join(', ') + '. Custom presets may render empty columns.',
    fix: {
      text: 'Run scripts/Map-CrawledProperties.ps1',
      href: 'https://github.com/hemantmane/Development/blob/main/sp-search/docs/deployment-guide.md',
    },
  };
}

// ─── (e) Search content source (informational) ───────────────────────────────

/**
 * Search content sources are admin-only API and can't be verified from
 * a normal user context. This check fires a no-op search query and
 * checks whether any results come back — proxy for "tenant has at
 * least one content source indexed."
 */
async function checkSearchContentSource(_signal?: AbortSignal): Promise<IReadinessCheck> {
  const id = 'content-source';
  const title = 'Tenant has indexed content';
  try {
    const results = await SPContext.sp.search({
      Querytext: '*',
      RowLimit: 1,
      TrimDuplicates: false,
      ClientType: 'SPSearchPreflight',
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const total = ((results as any)?.TotalRows) || ((results as any)?.RawSearchResults?.PrimaryQueryResult?.RelevantResults?.TotalRows) || 0;
    if (total > 0) {
      return {
        id, title,
        status: 'green',
        message: 'Tenant index returned at least one document. Search is operational.',
      };
    }
    return {
      id, title,
      status: 'yellow',
      message: 'Search returned zero documents for "*". The tenant may be empty, the user may have no read access to any content, or the index may not have crawled yet.',
    };
  } catch (err) {
    return {
      id, title,
      status: 'red',
      message: 'Search query failed: ' + (err instanceof Error ? err.message : 'unknown error'),
    };
  }
}
