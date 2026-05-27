/**
 * T4.D5 — shared edit-mode validators for Filters / Manager / Admin Manager.
 *
 * Each validator is a pure function returning IConfigValidationIssue[]. The
 * wiring into web part edit-mode MessageBar surfaces lives in the respective
 * components — see `SpSearchResults.tsx:804` for the canonical pattern.
 *
 * Per audit acceptance signal: "Filters typo raises did-you-mean; malformed
 * coverage URL flagged; expectedSiteUrls non-tenant URL warns; 4 validators
 * unit-tested."
 */

import type { IManagedProperty } from '../interfaces/ISearchDataProvider';
import {
  validateManagedProperty,
  type IValidateOptions,
} from '../utils/managedPropertyValidation';

export type ValidationSeverity = 'error' | 'warning' | 'info';

export interface IConfigValidationIssue {
  /** Stable key for React reconciliation. */
  id: string;
  severity: ValidationSeverity;
  message: string;
  /** Optional row index (0-based) for collection-cell validators. */
  rowIndex?: number;
  /** Optional field key — lets the UI focus the offending field. */
  fieldKey?: string;
}

// ─── validateCoverageProfileSourceUrls ──────────────────────────────────────

export interface ICoverageProfileLike {
  title?: string;
  sourceUrls: string;
}

/**
 * Check each profile's `sourceUrls` (comma-separated) for shape:
 *   - parseable absolute URL or server-relative `/sites/...` / `/teams/...`
 *   - if absolute, host must match the current tenant authority
 *
 * Returns one issue per row that has any bad URL. The audit signal only
 * requires per-row granularity, not per-URL.
 */
export function validateCoverageProfileSourceUrls(
  profiles: ICoverageProfileLike[],
  tenantRoot: string
): IConfigValidationIssue[] {
  const issues: IConfigValidationIssue[] = [];
  const tenantAuthority = safeAuthority(tenantRoot);

  for (let i = 0; i < profiles.length; i++) {
    const profile = profiles[i];
    const raw = (profile.sourceUrls || '').trim();
    if (!raw) {
      continue;
    }

    const urls = raw.split(',').map((u) => u.trim()).filter(Boolean);
    if (urls.length === 0) {
      continue;
    }

    let malformedCount = 0;
    let externalCount = 0;
    const externalHosts: string[] = [];

    for (let u = 0; u < urls.length; u++) {
      const url = urls[u];
      if (isServerRelativeSpUrl(url)) {
        continue;
      }
      const parsed = tryParseUrl(url);
      if (!parsed) {
        malformedCount++;
        continue;
      }
      const auth = parsed.hostAuthority;
      if (tenantAuthority && auth !== tenantAuthority) {
        externalCount++;
        if (externalHosts.indexOf(auth) < 0) {
          externalHosts.push(auth);
        }
      }
    }

    const title = profile.title ? '"' + profile.title + '"' : 'Row ' + (i + 1);

    if (malformedCount > 0) {
      issues.push({
        id: 'coverage-source-url-malformed-' + i,
        severity: 'error',
        rowIndex: i,
        fieldKey: 'sourceUrls',
        message: title + ' has ' + malformedCount + ' source URL(s) that are not a valid absolute or `/sites/`/`/teams/` server-relative URL.',
      });
      continue;
    }

    if (externalCount > 0) {
      issues.push({
        id: 'coverage-source-url-external-' + i,
        severity: 'warning',
        rowIndex: i,
        fieldKey: 'sourceUrls',
        message: title + ' points at a different tenant (' + externalHosts.join(', ') + '). Cross-tenant coverage queries require the federated provider — confirm this is intentional.',
      });
    }
  }

  return issues;
}

// ─── validateExpectedSiteUrls ───────────────────────────────────────────────

/**
 * Validate the multi-line `expectedSiteUrls` admin field. Each line is one
 * absolute URL. Best-effort URL parse; reachability check via Graph is out
 * of scope for v1.0 (the audit explicitly notes "no blocking" and we don't
 * want to fire a Graph call from inside the property pane).
 */
export function validateExpectedSiteUrls(
  urls: string[],
  tenantRoot: string
): IConfigValidationIssue[] {
  const issues: IConfigValidationIssue[] = [];
  const tenantAuthority = safeAuthority(tenantRoot);

  for (let i = 0; i < urls.length; i++) {
    const url = (urls[i] || '').trim();
    if (!url) {
      continue;
    }

    if (!/^https?:\/\//i.test(url)) {
      issues.push({
        id: 'expected-site-url-malformed-' + i,
        severity: 'error',
        rowIndex: i,
        fieldKey: 'expectedSiteUrls',
        message: 'Line ' + (i + 1) + ': "' + url + '" is not a valid absolute URL — expected an https:// site URL.',
      });
      continue;
    }

    if (url.toLowerCase().startsWith('http://')) {
      issues.push({
        id: 'expected-site-url-not-https-' + i,
        severity: 'error',
        rowIndex: i,
        fieldKey: 'expectedSiteUrls',
        message: 'Line ' + (i + 1) + ': SharePoint Online requires https — change "' + url + '" to https://.',
      });
      continue;
    }

    const parsed = tryParseUrl(url);
    if (!parsed) {
      issues.push({
        id: 'expected-site-url-unparseable-' + i,
        severity: 'error',
        rowIndex: i,
        fieldKey: 'expectedSiteUrls',
        message: 'Line ' + (i + 1) + ': "' + url + '" is not a valid URL.',
      });
      continue;
    }

    if (tenantAuthority && parsed.hostAuthority !== tenantAuthority) {
      issues.push({
        id: 'expected-site-url-external-' + i,
        severity: 'warning',
        rowIndex: i,
        fieldKey: 'expectedSiteUrls',
        message: 'Line ' + (i + 1) + ': "' + url + '" is on a different tenant (' + parsed.hostAuthority + ') — confirm this is intentional.',
      });
    }
  }

  return issues;
}

// ─── validateManagedPropertyCollection ──────────────────────────────────────

export interface IManagedPropertyRow {
  property: string;
  /** Optional managed-property-like field carrying a managed property — overrides `property` when present. */
  managedProperty?: string;
}

/**
 * Wrap the row-level `validateManagedProperty` (T4.D3) over a collection of
 * rows. Empty cells skip silently; cold-cache schema also skips silently.
 */
export function validateManagedPropertyCollection(
  items: IManagedPropertyRow[],
  schema: IManagedProperty[] | undefined,
  options: IValidateOptions = {}
): IConfigValidationIssue[] {
  const issues: IConfigValidationIssue[] = [];

  for (let i = 0; i < items.length; i++) {
    const row = items[i];
    const value = (row.managedProperty || row.property || '').trim();
    if (!value) {
      continue;
    }

    const result = validateManagedProperty(value, schema, options);
    if (!result.valid) {
      // ts-jest non-strict mode doesn't narrow discriminated unions on `!result.valid`.
      // Explicit cast keeps the impl typed at strict and ts-jest runtime.
      const errorResult = result as { valid: false; message: string };
      issues.push({
        id: 'managed-property-' + i,
        severity: 'error',
        rowIndex: i,
        fieldKey: row.managedProperty !== undefined ? 'managedProperty' : 'property',
        message: 'Row ' + (i + 1) + ': ' + errorResult.message,
      });
    }
  }

  return issues;
}

// ─── validateRefinementFilterCollection ─────────────────────────────────────

export interface IRefinementFilterRow {
  property: string;
  operator: string;
  value: string;
}

const SUPPORTED_OPERATORS = new Set<string>([
  'equals',
  'range',
  'startsWith',
]);

/**
 * Check each refinement filter row for:
 *   - present operator (required)
 *   - operator is in the supported set
 *   - value type matches operator (range must be `min,max`)
 *   - value is non-empty
 *
 * Matches the operator subset handled by `SearchService.buildRefinementFilters`.
 */
export function validateRefinementFilterCollection(
  items: IRefinementFilterRow[]
): IConfigValidationIssue[] {
  const issues: IConfigValidationIssue[] = [];

  for (let i = 0; i < items.length; i++) {
    const row = items[i];
    const property = (row.property || '').trim();
    const operator = (row.operator || '').trim();
    const value = (row.value || '').trim();

    if (!property) {
      continue;
    }

    if (!operator) {
      issues.push({
        id: 'refinement-filter-no-operator-' + i,
        severity: 'error',
        rowIndex: i,
        fieldKey: 'operator',
        message: 'Row ' + (i + 1) + ': missing operator — pick one of equals, range, startsWith.',
      });
      continue;
    }

    if (!SUPPORTED_OPERATORS.has(operator)) {
      issues.push({
        id: 'refinement-filter-bad-operator-' + i,
        severity: 'error',
        rowIndex: i,
        fieldKey: 'operator',
        message: 'Row ' + (i + 1) + ': unsupported operator "' + operator + '". Pick one of equals, range, startsWith.',
      });
      continue;
    }

    if (!value) {
      issues.push({
        id: 'refinement-filter-no-value-' + i,
        severity: 'error',
        rowIndex: i,
        fieldKey: 'value',
        message: 'Row ' + (i + 1) + ': value is required.',
      });
      continue;
    }

    if (operator === 'range') {
      const parts = value.split(',').map((s) => s.trim()).filter(Boolean);
      if (parts.length !== 2) {
        issues.push({
          id: 'refinement-filter-range-shape-' + i,
          severity: 'error',
          rowIndex: i,
          fieldKey: 'value',
          message: 'Row ' + (i + 1) + ': range operator requires two values separated by a comma (e.g. "0,1000000").',
        });
      }
    }
  }

  return issues;
}

// ─── helpers ────────────────────────────────────────────────────────────────

function safeAuthority(url: string): string {
  if (!url) { return ''; }
  const parsed = tryParseUrl(url);
  return parsed ? parsed.hostAuthority : '';
}

function tryParseUrl(url: string): { hostAuthority: string } | undefined {
  try {
    const u = new URL(url);
    // Authority = host (Node's URL doesn't preserve userInfo for SP URLs).
    return { hostAuthority: u.host.toLowerCase() };
  } catch {
    return undefined;
  }
}

function isServerRelativeSpUrl(url: string): boolean {
  const trimmed = (url || '').trim();
  return /^\/(sites|teams)\//i.test(trimmed);
}
