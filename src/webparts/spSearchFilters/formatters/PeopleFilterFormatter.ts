import 'spfx-toolkit/lib/utilities/context/pnpImports/lists';
import type { IFilterValueFormatter, IFilterConfig } from '@interfaces/index';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';

/**
 * Module-level display name cache: maps claim strings → display names.
 * Persists for the page session to avoid redundant profile API calls.
 */
const nameCache: Map<string, string> = new Map();

/**
 * Extract the email/login from a claim string.
 * Format: i:0#.f|membership|john@contoso.com → john@contoso.com
 */
function extractLogin(claim: string): string {
  const parts: string[] = claim.split('|');
  if (parts.length >= 3) {
    return parts[parts.length - 1];
  }
  return claim;
}

/**
 * Resolve a claim string to a display name via User Profile service.
 * Falls back to extracting the login portion of the claim.
 */
async function resolveDisplayName(claim: string): Promise<string> {
  try {
    const profile = await SPContext.sp.profiles.getPropertiesFor(claim);
    if (profile && profile.DisplayName) {
      return profile.DisplayName;
    }
  } catch {
    // Profile lookup failed
  }

  // Fallback: extract login name and format it
  const login: string = extractLogin(claim);
  // Convert "john.doe@contoso.com" → "John Doe"
  const atIndex: number = login.indexOf('@');
  if (atIndex > 0) {
    const name: string = login.substring(0, atIndex);
    const parts: string[] = name.split('.');
    const formatted: string[] = [];
    for (let i: number = 0; i < parts.length; i++) {
      if (parts[i].length > 0) {
        formatted.push(parts[i].charAt(0).toUpperCase() + parts[i].substring(1));
      }
    }
    return formatted.join(' ');
  }

  return login;
}

/**
 * PeopleFilterFormatter — formats people/user refinement values.
 *
 * Handles claim string tokens (i:0#.f|membership|...) from SharePoint Search.
 * Resolves to display names via PnP User Profile service.
 * Caches resolved names in a module-level Map for the page session.
 */
export const PeopleFilterFormatter: IFilterValueFormatter = {
  id: 'people',

  formatForDisplay: async function (rawValue: string, _config: IFilterConfig): Promise<string> {
    // Check cache first
    if (nameCache.has(rawValue)) {
      return nameCache.get(rawValue) as string;
    }

    const displayName: string = await resolveDisplayName(rawValue);
    nameCache.set(rawValue, displayName);
    return displayName;
  },

  formatForQuery: function (displayValue: unknown, _config: IFilterConfig): string {
    // People refinement tokens use the raw claim string
    return String(displayValue);
  },

  formatForUrl: function (rawValue: string): string {
    // URL-encode the claim string (contains special characters like #, |)
    return encodeURIComponent(rawValue);
  },

  parseFromUrl: function (urlValue: string): string {
    return decodeURIComponent(urlValue);
  }
};
