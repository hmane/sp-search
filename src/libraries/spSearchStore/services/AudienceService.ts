import { SPContext } from 'spfx-toolkit/lib/utilities/context';

// `/me/memberOf` returns a heterogeneous Collection(directoryObject). `@odata.type`
// is an OData annotation, NOT a selectable property — putting it in `$select` makes
// Graph reject the request with HTTP 400. It is emitted automatically for the
// derived types (group / directoryRole / …), so `$select=id` alone keeps the type
// discriminator below working.
//
// Permission: `User.Read` (delegated, work/school) — the documented least-
// privilege scope for the signed-in user's own direct memberships per
// learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0
// (table "Permissions for the signed-in user's direct memberships"). Declared
// in `config/package-solution.json` → `webApiPermissionRequests`. Tenants
// upgrading from a pre-Stream-D build must approve the added scope at
// SharePoint admin → "API access"; until approved the call returns 401/403
// and `resolveUserGroupIds()` resolves to [] (fail-closed: audience-targeted
// content stays hidden).
const GRAPH_MEMBER_OF_URL = 'https://graph.microsoft.com/v1.0/me/memberOf?$select=id';

/** Cached group IDs — groups don't change mid-session */
let cachedGroupIds: string[] | undefined;

/**
 * AudienceService — resolves the current user's Azure AD security group
 * memberships and provides audience-checking utilities.
 *
 * Uses SPContext.http with AadHttpClientFactory to call Microsoft Graph.
 * Results are cached for the lifetime of the page session.
 */

/**
 * Resolve the current user's Azure AD group IDs via Graph /me/memberOf.
 * Cached after first call — subsequent calls return the cached result.
 */
export async function resolveUserGroupIds(): Promise<string[]> {
  if (cachedGroupIds !== undefined) {
    return cachedGroupIds;
  }

  try {
    const response = await SPContext.http.get<{
      value: Array<{ id: string; '@odata.type': string }>;
    }>(GRAPH_MEMBER_OF_URL, {
      useAuth: true,
      resourceUri: 'https://graph.microsoft.com',
    });

    // SPContext.http.get does NOT throw on non-2xx — check response.ok
    if (!response.ok) {
      SPContext.logger.warn('AudienceService: Graph API returned HTTP ' + String(response.status));
      const empty: string[] = [];
      cachedGroupIds = empty; // eslint-disable-line require-atomic-updates -- intentional last-write-wins cache
      return empty;
    }

    const groupIds: string[] = [];
    if (response.data && response.data.value) {
      for (let i = 0; i < response.data.value.length; i++) {
        const entry = response.data.value[i];
        // Include groups and directory roles (both relevant for audience targeting)
        if (
          entry['@odata.type'] === '#microsoft.graph.group' ||
          entry['@odata.type'] === '#microsoft.graph.directoryRole'
        ) {
          groupIds.push(entry.id);
        }
      }
    }

    const result = groupIds;
    cachedGroupIds = result; // eslint-disable-line require-atomic-updates -- intentional last-write-wins cache
    SPContext.logger.info('AudienceService: Resolved user groups', { count: result.length });
    return result;
  } catch (error) {
    SPContext.logger.warn('AudienceService: Failed to resolve user groups', { error });
    // On failure, return empty — all audience-targeted items will be hidden
    // This is safer than showing everything (fail-closed)
    const fallback: string[] = [];
    cachedGroupIds = fallback; // eslint-disable-line require-atomic-updates -- intentional last-write-wins cache
    return fallback;
  }
}

/**
 * Check if the current user is in at least one of the specified audience groups.
 *
 * @param audienceGroups - Azure AD group IDs that should see this content
 * @param userGroupIds - The current user's group IDs
 * @returns true if audienceGroups is empty (no targeting) or user is in at least one group
 */
export function isInAudience(audienceGroups: string[] | undefined, userGroupIds: string[]): boolean {
  // No audience targeting = visible to everyone
  if (!audienceGroups || audienceGroups.length === 0) {
    return true;
  }

  // Check if user is in any of the target groups
  for (let i = 0; i < audienceGroups.length; i++) {
    if (userGroupIds.indexOf(audienceGroups[i]) >= 0) {
      return true;
    }
  }

  return false;
}
