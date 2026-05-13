/**
 * Context values used to resolve template tokens in query strings.
 * Populated from SPFx PageContext and web part properties.
 */
export interface ITokenContext {
  /** User's search query text */
  queryText: string;
  /** Current site collection GUID */
  siteId: string;
  /** Current site collection URL */
  siteUrl: string;
  /** Current web GUID */
  webId: string;
  /** Current web URL */
  webUrl: string;
  /** Hub site URL; empty string if the site is not part of a hub */
  hubSiteUrl: string;
  /** Current user display name */
  userDisplayName: string;
  /** Current user email / login name */
  userEmail: string;
  /** Current list ID (populated when on a list page, empty otherwise) */
  listId: string;
}

/**
 * Token resolution service that replaces template tokens in query strings.
 *
 * Supported tokens:
 *   {searchTerms}         - user's query text
 *   {Site.ID}             - current site GUID
 *   {Site.URL}            - current site URL
 *   {Web.ID}              - current web GUID
 *   {Web.URL}             - current web URL
 *   {Hub}                 - hub site URL (empty string if not in hub)
 *   {Today}               - current date in ISO format (YYYY-MM-DD)
 *   {Today+N} / {Today-N} - date offset by N days
 *   {User.Name}           - current user display name
 *   {User.Email}          - current user email / login name
 *   {PageContext.listId}  - current list ID (if on a list page)
 *
 * Unknown tokens are left unreplaced.
 */
export class TokenService {

  /**
   * Replace all recognised `{...}` tokens in the template string
   * with their corresponding values from the provided context.
   *
   * This is a pure function with no side effects.
   *
   * @param template - The query template containing `{...}` tokens.
   * @param context  - The context values to substitute.
   * @returns The template string with all recognised tokens replaced.
   */
  public static resolveTokens(template: string, context: ITokenContext): string {
    if (!template) {
      return '';
    }

    // Build a map of simple (non-date) token replacements
    const tokenMap: Record<string, string> = {
      'searchTerms': context.queryText,
      'Site.ID': context.siteId,
      'Site.URL': context.siteUrl,
      'Web.ID': context.webId,
      'Web.URL': context.webUrl,
      'Hub': context.hubSiteUrl,
      'User.Name': context.userDisplayName,
      'User.Email': context.userEmail,
      'PageContext.listId': context.listId,
    };

    // Replace all {token} occurrences using a single regex pass.
    // The regex matches any `{...}` pattern including date offsets like {Today+5}.
    return template.replace(/\{([^}]+)\}/g, (match: string, tokenKey: string): string => {

      // Check simple token map first
      if (Object.prototype.hasOwnProperty.call(tokenMap, tokenKey)) {
        return tokenMap[tokenKey];
      }

      // Handle {Today}, {Today+N}, {Today-N}
      const todayMatch: RegExpExecArray | undefined =
        /^Today([+-]\d+)?$/.exec(tokenKey) || undefined;

      if (todayMatch !== undefined) {
        const today: Date = new Date();
        const offset: string | undefined = todayMatch[1];

        if (offset !== undefined) {
          const days: number = parseInt(offset, 10);
          today.setDate(today.getDate() + days);
        }

        return TokenService._formatDate(today);
      }

      // Unknown token: leave unreplaced
      return match;
    });
  }

  /**
   * MISS-001 — Apply a `queryInputTransformation` template to the user's
   * raw search text. The transformation may reference any token the
   * full {@link resolveTokens} method understands, not just `{searchTerms}`.
   *
   * Used by the orchestrator's `_buildEffectiveQueryText` to expand
   * admin-configured patterns like
   *   `({searchTerms}) AND owner:{User.Email}`
   * before the result is handed to the data provider. Without this,
   * tokens other than `{searchTerms}` flow into the provider unresolved
   * and reach SharePoint as literal text (server-side substitution only
   * handles `{searchTerms}` in the outer `QueryTemplate` field).
   *
   * Empty-query behaviour: if `rawQuery` is empty/whitespace, the
   * transformation is still resolved (so non-`{searchTerms}` parts
   * survive). If the resolved + trimmed result is empty, returns `*`
   * (SharePoint's match-all sentinel) instead of an empty query string.
   *
   * @param transformation - The admin-configured template (defaults to `{searchTerms}` when blank).
   * @param rawQuery       - The user's typed query text.
   * @param context        - Token context (everything except `queryText`, which is overridden by `rawQuery`).
   */
  public static applyQueryInputTransformation(
    transformation: string,
    rawQuery: string,
    context: ITokenContext
  ): string {
    const template = transformation && transformation.trim() ? transformation : '{searchTerms}';
    const tokenContext: ITokenContext = { ...context, queryText: rawQuery || '' };
    const resolved = TokenService.resolveTokens(template, tokenContext).trim();
    return resolved || '*';
  }

  /**
   * Format a Date as an ISO date string (YYYY-MM-DD).
   */
  private static _formatDate(date: Date): string {
    const year: number = date.getFullYear();
    const monthStr: string = String(date.getMonth() + 1);
    const month: string = ('0' + monthStr).slice(-2);
    const dayStr: string = String(date.getDate());
    const day: string = ('0' + dayStr).slice(-2);
    return `${year}-${month}-${day}`;
  }
}
