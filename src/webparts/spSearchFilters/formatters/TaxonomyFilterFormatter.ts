import 'spfx-toolkit/lib/utilities/context/pnpImports/taxonomy';
import type { IFilterValueFormatter, IFilterConfig } from '@interfaces/index';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';

/**
 * Module-level label cache: maps GP0|#GUID tokens → resolved term labels.
 * Persists for the page session to avoid redundant taxonomy API calls.
 */
const labelCache: Map<string, string> = new Map();

/**
 * Extract a GUID from a GP0|#GUID style taxonomy refiner token.
 */
function extractGuid(token: string): string | undefined {
  const guidRegex: RegExp = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
  const match: RegExpMatchArray | null = token.match(guidRegex);
  return match ? match[0] : undefined;
}

/**
 * Extract an embedded label from a taxonomy token if present.
 * Format: L0|#0GUID|Label Text
 */
function extractEmbeddedLabel(token: string): string | undefined {
  const parts: string[] = token.split('|');
  if (parts.length >= 3) {
    const last: string = parts[parts.length - 1].trim();
    if (last.length > 0 && !last.startsWith('#')) {
      return last;
    }
  }
  return undefined;
}

/**
 * Resolve a taxonomy term GUID to its label via PnP Taxonomy API.
 * Returns the GUID as fallback if resolution fails.
 */
async function resolveTermLabel(guid: string): Promise<string> {
  try {
    // Use any cast since PnP taxonomy typings may not expose getTermById directly
    const termStore = SPContext.sp.termStore as any;
    const termInfo = await termStore.getTermById(guid)();
    if (termInfo && termInfo.labels && termInfo.labels.length > 0) {
      const defaultLabel = termInfo.labels.find((l: { isDefault: boolean }) => l.isDefault);
      return defaultLabel ? defaultLabel.name : termInfo.labels[0].name;
    }
  } catch {
    // Term lookup failed — fall back to GUID
  }
  return guid;
}

/**
 * TaxonomyFilterFormatter — formats taxonomy/managed metadata refinement values.
 *
 * Handles GP0|#GUID format tokens from SharePoint Search refiners.
 * Resolves GUIDs to human-readable term labels via PnP Taxonomy API.
 * Caches resolved labels in a module-level Map for the page session.
 */
export const TaxonomyFilterFormatter: IFilterValueFormatter = {
  id: 'taxonomy',

  formatForDisplay: async function (rawValue: string, _config: IFilterConfig): Promise<string> {
    // Check cache first
    if (labelCache.has(rawValue)) {
      return labelCache.get(rawValue) as string;
    }

    // Try to extract embedded label (L0|#0GUID|Label format)
    const embedded: string | undefined = extractEmbeddedLabel(rawValue);
    if (embedded) {
      labelCache.set(rawValue, embedded);
      return embedded;
    }

    // Extract GUID and resolve via Taxonomy API
    const guid: string | undefined = extractGuid(rawValue);
    if (!guid) {
      return rawValue;
    }

    const label: string = await resolveTermLabel(guid);
    labelCache.set(rawValue, label);
    return label;
  },

  formatForQuery: function (displayValue: unknown, _config: IFilterConfig): string {
    // Taxonomy refinement tokens are used as-is (GP0|#GUID format)
    return String(displayValue);
  },

  formatForUrl: function (rawValue: string): string {
    // Extract just the GUID for compact URLs
    const guid: string | undefined = extractGuid(rawValue);
    return guid || encodeURIComponent(rawValue);
  },

  parseFromUrl: function (urlValue: string): string {
    // If it's a bare GUID, wrap it in GP0|# format
    const guidRegex: RegExp = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (guidRegex.test(urlValue)) {
      return 'GP0|#' + urlValue;
    }
    return decodeURIComponent(urlValue);
  }
};
