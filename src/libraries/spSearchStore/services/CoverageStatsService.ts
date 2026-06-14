import 'spfx-toolkit/lib/utilities/context/pnpImports/search';
import type { ISearchQuery as IPnPSearchQuery, SearchResults } from '@pnp/sp/search';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import type { ISearchScope } from '@interfaces/index';
import { spLog } from '@store/utils/spLog';

export interface ICoverageConfig {
  queryTemplate: string;
  scope: ISearchScope;
  resultSourceId: string | undefined;
  refinementFilters: string | undefined;
}

export interface ICoverageStatsResult {
  itemCount: number;
  newest: Date | undefined;
  oldest: Date | undefined;
  fileTypes: Array<{ type: string; count: number }>;
  actualSites: Array<{ url: string; count: number }>;
  timestamp: number;
}

export class CoverageStatsService {
  private readonly _config: ICoverageConfig;

  public constructor(config: ICoverageConfig) {
    this._config = config;
  }

  private _baseRequest(): IPnPSearchQuery {
    const req: IPnPSearchQuery = {
      Querytext: '*',
      QueryTemplate: this._config.queryTemplate || '{searchTerms}',
      RowLimit: 1,
      SelectProperties: ['Title', 'LastModifiedTime'],
      TrimDuplicates: false,
      ClientType: 'SPSearchCoverage',
    };

    if (this._config.scope && this._config.scope.kqlPath) {
      req.Querytext = '* ' + this._config.scope.kqlPath;
    }

    const sourceId = this._config.resultSourceId ||
      (this._config.scope && this._config.scope.resultSourceId) || undefined;
    if (sourceId) {
      req.SourceId = sourceId;
    }

    if (this._config.refinementFilters) {
      const filters = this._config.refinementFilters
        .split(',')
        .map(function (f: string): string { return f.trim(); })
        .filter(Boolean);
      if (filters.length > 0) {
        req.RefinementFilters = filters;
      }
    }

    return req;
  }

  public async getItemCount(signal: AbortSignal): Promise<number> {
    if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
    const req = this._baseRequest();
    req.RowLimit = 1;
    req.SelectProperties = ['Title'];
    const results: SearchResults = await SPContext.sp.search(req);
    return results.TotalRows;
  }

  public async getFreshness(signal: AbortSignal): Promise<{ newest: Date | undefined; oldest: Date | undefined }> {
    if (signal.aborted) throw new DOMException('Aborted', 'AbortError');

    const newestReq = this._baseRequest();
    newestReq.RowLimit = 1;
    newestReq.SelectProperties = ['LastModifiedTime'];
    newestReq.SortList = [{ Property: 'LastModifiedTime', Direction: 1 }];

    const oldestReq = this._baseRequest();
    oldestReq.RowLimit = 1;
    oldestReq.SelectProperties = ['LastModifiedTime'];
    oldestReq.SortList = [{ Property: 'LastModifiedTime', Direction: 0 }];

    const [newestResults, oldestResults] = await Promise.all([
      SPContext.sp.search(newestReq).catch(function (): undefined { return undefined; }),
      SPContext.sp.search(oldestReq).catch(function (): undefined { return undefined; }),
    ]);

    let newest: Date | undefined;
    let oldest: Date | undefined;

    if (newestResults && newestResults.PrimarySearchResults.length > 0) {
      const val = (newestResults.PrimarySearchResults[0] as Record<string, unknown>).LastModifiedTime;
      if (val) newest = new Date(val as string);
    }
    if (oldestResults && oldestResults.PrimarySearchResults.length > 0) {
      const val = (oldestResults.PrimarySearchResults[0] as Record<string, unknown>).LastModifiedTime;
      if (val) oldest = new Date(val as string);
    }

    // Fallback: if sorted queries returned no results (LastModifiedTime may not be sortable),
    // run an unsorted query and extract LastModifiedTime from the first result
    if (!newest && !oldest) {
      try {
        const fallbackReq = this._baseRequest();
        fallbackReq.RowLimit = 1;
        fallbackReq.SelectProperties = ['LastModifiedTime'];
        const fallbackResults: SearchResults = await SPContext.sp.search(fallbackReq);
        if (fallbackResults.PrimarySearchResults.length > 0) {
          const val = (fallbackResults.PrimarySearchResults[0] as Record<string, unknown>).LastModifiedTime;
          if (val) {
            newest = new Date(val as string);
          }
        }
        spLog.warn('LastModifiedTime may not be sortable; freshness data is limited');
      } catch {
        // Fallback failed — leave dates undefined
      }
    }

    return { newest, oldest };
  }

  public async getFileTypeBreakdown(signal: AbortSignal): Promise<Array<{ type: string; count: number }>> {
    if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
    const req = this._baseRequest();
    req.RowLimit = 1;
    req.SelectProperties = ['Title'];
    req.Refiners = 'FileType';

    const results: SearchResults = await SPContext.sp.search(req);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const raw = results.RawSearchResults as any;
    const refiners = raw?.PrimaryQueryResult?.RefinementResults?.Refiners;
    if (!refiners || !Array.isArray(refiners) || refiners.length === 0) {
      return [];
    }

    const fileTypeRefiner = refiners[0];
    if (!fileTypeRefiner.Entries) return [];

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return fileTypeRefiner.Entries.map(function (e: any): { type: string; count: number } {
      return { type: e.RefinementName || e.RefinementValue || '', count: e.RefinementCount || 0 };
    }).sort(function (a: { count: number }, b: { count: number }): number { return b.count - a.count; });
  }

  public async getSiteDistribution(signal: AbortSignal): Promise<Array<{ url: string; count: number }>> {
    if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
    const req = this._baseRequest();
    req.RowLimit = 1;
    req.SelectProperties = ['Title'];
    req.Refiners = 'SPWebUrl';

    const results: SearchResults = await SPContext.sp.search(req);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const raw = results.RawSearchResults as any;
    const refiners = raw?.PrimaryQueryResult?.RefinementResults?.Refiners;
    if (!refiners || !Array.isArray(refiners) || refiners.length === 0) {
      return [];
    }

    const siteRefiner = refiners[0];
    if (!siteRefiner.Entries) return [];

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return siteRefiner.Entries.map(function (e: any): { url: string; count: number } {
      return { url: e.RefinementName || e.RefinementValue || '', count: e.RefinementCount || 0 };
    }).sort(function (a: { count: number }, b: { count: number }): number { return b.count - a.count; });
  }

  public async runAll(signal: AbortSignal): Promise<ICoverageStatsResult> {
    const [itemCount, freshness, fileTypes, actualSites] = await Promise.all([
      this.getItemCount(signal),
      this.getFreshness(signal),
      this.getFileTypeBreakdown(signal),
      this.getSiteDistribution(signal),
    ]);

    return {
      itemCount,
      newest: freshness.newest,
      oldest: freshness.oldest,
      fileTypes,
      actualSites,
      timestamp: Date.now(),
    };
  }
}
