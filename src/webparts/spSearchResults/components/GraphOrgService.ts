/**
 * GraphOrgService — fetches manager and direct-reports relationships from
 * Microsoft Graph for use in the People layout org-chart panel.
 *
 * Results are cached in-memory per user ID for the lifetime of the web part
 * so repeated expansions don't re-hit the API.
 */

import type { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';

export interface IOrgPerson {
  id: string;
  displayName: string;
  jobTitle?: string;
  department?: string;
  mail?: string;
  officeLocation?: string;
  userPrincipalName?: string;
}

interface IUserOrgCache {
  manager?: IOrgPerson | null;  // null means "confirmed no manager"
  reports?: IOrgPerson[];
}

const PERSON_SELECT = 'id,displayName,jobTitle,department,mail,officeLocation,userPrincipalName';

export class GraphOrgService {
  private _client: MSGraphClientV3;
  private _cache: Map<string, IUserOrgCache> = new Map();

  public constructor(client: MSGraphClientV3) {
    this._client = client;
  }

  /**
   * Fetches the manager for a given user ID (GUID or UPN).
   * Returns null when the user has no manager or Graph returns 404.
   * Returns undefined when the request fails for any other reason.
   */
  public async fetchManager(userId: string): Promise<IOrgPerson | null | undefined> {
    const cached = this._cache.get(userId);
    if (cached && 'manager' in cached) {
      return cached.manager;
    }

    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const response: any = await this._client
        .api('/users/' + encodeURIComponent(userId) + '/manager')
        .select(PERSON_SELECT)
        .get();

      const manager: IOrgPerson | null = response ? this._mapPerson(response) : null;
      this._setCacheManager(userId, manager);
      return manager;
    } catch (err: unknown) {
      // 404 = user has no manager configured; treat as null
      if (this._isNotFound(err)) {
        this._setCacheManager(userId, null);
        return null;
      }
      // Other errors (forbidden, network) — return undefined so callers can show a graceful fallback
      return undefined;
    }
  }

  /**
   * Fetches the direct reports for a given user ID (GUID or UPN).
   * Returns an empty array when the user has no reports.
   * Returns undefined when the request fails.
   */
  public async fetchDirectReports(userId: string): Promise<IOrgPerson[] | undefined> {
    const cached = this._cache.get(userId);
    if (cached && cached.reports !== undefined) {
      return cached.reports;
    }

    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const response: any = await this._client
        .api('/users/' + encodeURIComponent(userId) + '/directReports')
        .select(PERSON_SELECT)
        .get();

      const reports: IOrgPerson[] = Array.isArray(response?.value)
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        ? (response.value as any[]).map((r) => this._mapPerson(r))
        : [];

      this._setCacheReports(userId, reports);
      return reports;
    } catch {
      return undefined;
    }
  }

  private _mapPerson(r: Record<string, unknown>): IOrgPerson {
    return {
      id: typeof r.id === 'string' ? r.id : '',
      displayName: typeof r.displayName === 'string' ? r.displayName : '',
      jobTitle: typeof r.jobTitle === 'string' ? r.jobTitle : undefined,
      department: typeof r.department === 'string' ? r.department : undefined,
      mail: typeof r.mail === 'string' ? r.mail : undefined,
      officeLocation: typeof r.officeLocation === 'string' ? r.officeLocation : undefined,
      userPrincipalName: typeof r.userPrincipalName === 'string' ? r.userPrincipalName : undefined,
    };
  }

  private _setCacheManager(userId: string, manager: IOrgPerson | null): void {
    const entry = this._cache.get(userId) || {};
    entry.manager = manager;
    this._cache.set(userId, entry);
  }

  private _setCacheReports(userId: string, reports: IOrgPerson[]): void {
    const entry = this._cache.get(userId) || {};
    entry.reports = reports;
    this._cache.set(userId, entry);
  }

  private _isNotFound(err: unknown): boolean {
    if (!err || typeof err !== 'object') { return false; }
    const e = err as Record<string, unknown>;
    // Graph SDK surfaces HTTP status in statusCode or status
    const status = e.statusCode ?? e.status;
    return status === 404;
  }
}
