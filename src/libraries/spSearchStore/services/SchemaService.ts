import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { IManagedProperty } from '@interfaces/index';

// ─── Constants ──────────────────────────────────────────────

const SCHEMA_CACHE_KEY = 'sp-search-schema-v1';

/** Maps SharePoint ManagedType integers to human-readable type names */
const MANAGED_TYPE_MAP: Record<number, string> = {
  1: 'Text',
  2: 'Integer',
  3: 'Decimal',
  4: 'DateTime',
  5: 'YesNo',
  6: 'Double',
  7: 'Binary',
};

// ─── Types ──────────────────────────────────────────────────

/** Raw API response shape from /_api/search/manage/schema/managedproperties */
interface ISchemaApiResponse {
  value?: ISchemaApiProperty[];
  d?: {
    results?: ISchemaApiProperty[];
  };
}

interface ISchemaApiProperty {
  Name: string;
  ManagedType?: number;
  Type?: string;
  Aliases?: string[] | string;
  Queryable?: boolean;
  Retrievable?: boolean;
  Refinable?: boolean;
  Sortable?: boolean;
}

/**
 * Result of a schema fetch operation.
 * Discriminated by `status` so the UI can differentiate permission issues from errors.
 */
export interface ISchemaResult {
  status: 'success' | 'unauthorized' | 'error';
  properties: IManagedProperty[];
  errorMessage?: string;
}

// ─── Public API ─────────────────────────────────────────────

/**
 * Fetch managed properties from the SharePoint Search Administration API.
 *
 * - Checks sessionStorage cache first (unless forceRefresh is true)
 * - On success: caches result in sessionStorage, returns status='success'
 * - On HTTP 403: returns status='unauthorized' (user lacks Search Admin permissions)
 * - On other errors: returns status='error' with message
 *
 * @param forceRefresh - Bypass sessionStorage cache and re-fetch from server
 */
export async function fetchManagedProperties(forceRefresh?: boolean): Promise<ISchemaResult> {
  // Check sessionStorage cache
  if (!forceRefresh) {
    const cached = readCache();
    if (cached) {
      return { status: 'success', properties: cached };
    }
  }

  try {
    const url = SPContext.webAbsoluteUrl + '/_api/search/manage/schema/managedproperties';
    const response = await SPContext.http.get<ISchemaApiResponse>(url);

    // SPContext.http.get does NOT throw on non-2xx — check response.ok explicitly
    if (!response.ok) {
      if (response.status === 403 || response.status === 401) {
        SPContext.logger.warn('SchemaService: Unauthorized — user lacks Search Admin permissions');
        return { status: 'unauthorized', properties: [] };
      }
      const errorMsg = 'Schema API returned HTTP ' + String(response.status);
      SPContext.logger.error('SchemaService: ' + errorMsg);
      return { status: 'error', properties: [], errorMessage: errorMsg };
    }

    // Extract array from either OData format
    let rawProperties: ISchemaApiProperty[] = [];
    if (response.data) {
      if (response.data.value) {
        rawProperties = response.data.value;
      } else if (response.data.d && response.data.d.results) {
        rawProperties = response.data.d.results;
      }
    }

    // Map to IManagedProperty[]
    const properties: IManagedProperty[] = [];
    for (let i = 0; i < rawProperties.length; i++) {
      const raw = rawProperties[i];
      if (!raw.Name) {
        continue;
      }

      // Resolve type: prefer string Type, fall back to ManagedType integer mapping
      let typeName = 'Text';
      if (raw.Type && typeof raw.Type === 'string') {
        typeName = raw.Type;
      } else if (raw.ManagedType !== undefined) {
        typeName = MANAGED_TYPE_MAP[raw.ManagedType] || 'Text';
      }

      // Resolve alias: handle string or string[] format
      let alias: string | undefined;
      if (raw.Aliases) {
        if (typeof raw.Aliases === 'string' && raw.Aliases.length > 0) {
          alias = raw.Aliases;
        } else if (Array.isArray(raw.Aliases) && raw.Aliases.length > 0) {
          alias = raw.Aliases[0];
        }
      }

      properties.push({
        name: raw.Name,
        type: typeName,
        alias: alias,
        queryable: raw.Queryable === true,
        retrievable: raw.Retrievable === true,
        refinable: raw.Refinable === true,
        sortable: raw.Sortable === true,
      });
    }

    // Sort alphabetically by name
    properties.sort(function (a, b): number {
      return a.name.localeCompare(b.name);
    });

    // Only cache successful responses with actual data
    writeCache(properties);

    SPContext.logger.info('SchemaService: Fetched managed properties', { count: properties.length });
    return { status: 'success', properties: properties };
  } catch (error) {
    // Network errors, timeouts, etc. still throw
    const message = error instanceof Error ? error.message : 'Failed to fetch schema';
    SPContext.logger.error('SchemaService: Failed to fetch managed properties', error);
    return { status: 'error', properties: [], errorMessage: message };
  }
}

/**
 * Synchronous read of cached schema from sessionStorage.
 * Returns undefined if no cached data exists.
 * Use this when you need cached data without triggering a fetch.
 */
export function getCachedSchema(): IManagedProperty[] | undefined {
  return readCache();
}

// ─── Internal helpers ───────────────────────────────────────

function readCache(): IManagedProperty[] | undefined {
  try {
    const raw = sessionStorage.getItem(SCHEMA_CACHE_KEY);
    if (raw) {
      return JSON.parse(raw) as IManagedProperty[];
    }
  } catch {
    // Corrupt cache — ignore
  }
  return undefined;
}

function writeCache(properties: IManagedProperty[]): void {
  try {
    sessionStorage.setItem(SCHEMA_CACHE_KEY, JSON.stringify(properties));
  } catch {
    // sessionStorage full or unavailable — ignore
  }
}
