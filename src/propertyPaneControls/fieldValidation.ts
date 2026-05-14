/**
 * T4.D8 — shared property-pane validators for high-impact free-text fields.
 *
 * These functions each match the `onGetErrorMessage` signature used by
 * SPFx's `PropertyPaneTextField`: return `''` for valid input, return a
 * non-empty error string to surface a red banner under the field.
 *
 * Centralising the validators here means the same field across multiple
 * web parts (e.g. if `expectedSiteUrls` ever appears on both Manager and
 * Admin Manager) shows the same error copy.
 */

// ─── coverageSourcePageUrl ──────────────────────────────────────────────────

/**
 * Accept:
 *   - empty (field is optional)
 *   - absolute https:// URL on `*.sharepoint.com` host
 *   - server-relative `/sites/...` or `/teams/...` path
 * Reject everything else with a single clear copy.
 */
export function validateCoverageSourcePageUrl(value: string): string {
  const trimmed = (value || '').trim();
  if (!trimmed) {
    return '';
  }

  // Server-relative paths.
  if (trimmed.startsWith('/')) {
    if (/^\/(sites|teams)\//i.test(trimmed)) {
      return '';
    }
    return 'Server-relative URL must start with /sites/ or /teams/ (e.g. /sites/search/SitePages/Search.aspx).';
  }

  // Absolute URLs.
  if (/^http:\/\//i.test(trimmed)) {
    return 'SharePoint Online requires https — change http:// to https://.';
  }

  if (!/^https:\/\//i.test(trimmed)) {
    return 'Must be a server-relative path (/sites/...) or an absolute https:// SharePoint URL.';
  }

  try {
    const parsed = new URL(trimmed);
    if (!/\.sharepoint\.com$/i.test(parsed.host)) {
      return 'URL host must end with .sharepoint.com.';
    }
  } catch {
    return 'Not a valid URL.';
  }

  return '';
}

// ─── expectedSiteUrls ───────────────────────────────────────────────────────

/**
 * The `expectedSiteUrls` field is multi-line — one URL per line. Validate
 * per-line and report the line number of the first failure so admins can
 * locate the bad row without staring at the textarea.
 */
export function validateExpectedSiteUrlsField(value: string): string {
  const raw = value || '';
  if (raw.trim() === '') {
    return '';
  }

  const lines = raw.split('\n');
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) {
      continue;
    }

    if (/\s/.test(line)) {
      return 'Line ' + (i + 1) + ': URL must not contain spaces.';
    }

    if (/^http:\/\//i.test(line)) {
      return 'Line ' + (i + 1) + ': SharePoint Online requires https — change http:// to https://.';
    }

    if (!/^https:\/\//i.test(line)) {
      return 'Line ' + (i + 1) + ': must be an absolute https:// URL (e.g. https://contoso.sharepoint.com/sites/Hub).';
    }

    try {
      const parsed = new URL(line);
      if (!parsed.host) {
        return 'Line ' + (i + 1) + ': not a valid URL.';
      }
    } catch {
      return 'Line ' + (i + 1) + ': not a valid URL.';
    }
  }

  return '';
}

// ─── newPageQueryParameter ──────────────────────────────────────────────────

/**
 * URL-query-parameter name: must start with a letter, then letters /
 * digits / dash / underscore. Excludes special URL characters (`?`, `&`,
 * `=`, `#`, space, etc.) that would corrupt the query string.
 */
export function validateNewPageQueryParameter(value: string): string {
  const trimmed = (value || '').trim();
  if (!trimmed) {
    return '';
  }
  if (!/^[A-Za-z][A-Za-z0-9_\-]*$/.test(trimmed)) {
    return 'Use letters, digits, dash, and underscore only. Must start with a letter.';
  }
  return '';
}
