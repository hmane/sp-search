/**
 * T4.D5 — edit-mode validators consumed by Filters / Manager / Admin Manager.
 *
 * Each validator is a pure function. The wiring into web part property panes
 * lives in the respective component files (Filters / Manager / Admin Manager
 * edit-mode MessageBar surfaces).
 *
 * Per audit acceptance signal: "Filters typo raises did-you-mean; malformed
 * coverage URL flagged; expectedSiteUrls non-tenant URL warns; 4 validators
 * unit-tested."
 */

import {
  validateCoverageProfileSourceUrls,
  validateExpectedSiteUrls,
  validateManagedPropertyCollection,
  validateRefinementFilterCollection,
} from '../../src/libraries/spSearchStore/configValidation/sharedValidators';
import type { IManagedProperty } from '../../src/libraries/spSearchStore/interfaces/ISearchDataProvider';

// ─── validateCoverageProfileSourceUrls ──────────────────────────────────────

describe('validateCoverageProfileSourceUrls', () => {
  const tenantRoot = 'https://dodgeandcox.sharepoint.com';

  it('passes when sourceUrls list is empty', () => {
    expect(validateCoverageProfileSourceUrls([], tenantRoot)).toEqual([]);
  });

  it('passes when every URL is on the current tenant', () => {
    const profiles = [
      { title: 'Docs', sourceUrls: 'https://dodgeandcox.sharepoint.com/sites/SPSearch/Shared Documents' },
      { title: 'Hub',  sourceUrls: 'https://dodgeandcox.sharepoint.com/sites/Hub/Docs' },
    ];
    expect(validateCoverageProfileSourceUrls(profiles, tenantRoot)).toEqual([]);
  });

  it('accepts server-relative URLs starting with /sites/ or /teams/', () => {
    const profiles = [
      { title: 'A', sourceUrls: '/sites/SPSearch/Docs, /teams/Marketing/Library' },
    ];
    expect(validateCoverageProfileSourceUrls(profiles, tenantRoot)).toEqual([]);
  });

  it('flags a URL that points at a different tenant', () => {
    const profiles = [
      { title: 'External', sourceUrls: 'https://other-tenant.sharepoint.com/sites/Foo/Docs' },
    ];
    const issues = validateCoverageProfileSourceUrls(profiles, tenantRoot);
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe('warning');
    expect(issues[0].rowIndex).toBe(0);
    expect(issues[0].message).toMatch(/different tenant/i);
  });

  it('flags an entirely malformed URL', () => {
    const profiles = [
      { title: 'Junk', sourceUrls: 'not-a-url' },
    ];
    const issues = validateCoverageProfileSourceUrls(profiles, tenantRoot);
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe('error');
    expect(issues[0].message).toMatch(/not a valid/i);
  });

  it('returns one issue per row even when many rows are bad', () => {
    const profiles = [
      { title: 'A', sourceUrls: 'https://dodgeandcox.sharepoint.com/sites/Ok/Docs' },
      { title: 'B', sourceUrls: 'broken' },
      { title: 'C', sourceUrls: 'https://other-tenant.sharepoint.com/sites/X' },
    ];
    const issues = validateCoverageProfileSourceUrls(profiles, tenantRoot);
    expect(issues).toHaveLength(2);
    expect(issues.map((i) => i.rowIndex).sort()).toEqual([1, 2]);
  });
});

// ─── validateExpectedSiteUrls ───────────────────────────────────────────────

describe('validateExpectedSiteUrls', () => {
  const tenantRoot = 'https://dodgeandcox.sharepoint.com';

  it('passes when list is empty', () => {
    expect(validateExpectedSiteUrls([], tenantRoot)).toEqual([]);
  });

  it('passes when all URLs are on the tenant', () => {
    const urls = [
      'https://dodgeandcox.sharepoint.com/sites/Hub',
      'https://dodgeandcox.sharepoint.com/sites/SPSearch',
    ];
    expect(validateExpectedSiteUrls(urls, tenantRoot)).toEqual([]);
  });

  it('flags a non-https URL', () => {
    const issues = validateExpectedSiteUrls(['http://dodgeandcox.sharepoint.com/sites/X'], tenantRoot);
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe('error');
    expect(issues[0].message).toMatch(/https/i);
  });

  it('flags a URL on a different tenant', () => {
    const issues = validateExpectedSiteUrls(
      ['https://contoso.sharepoint.com/sites/Foo'],
      tenantRoot
    );
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe('warning');
    expect(issues[0].message).toMatch(/different tenant/i);
  });

  it('preserves rowIndex (line number) so admins can locate the bad row', () => {
    const issues = validateExpectedSiteUrls(
      [
        'https://dodgeandcox.sharepoint.com/sites/A',
        'junk',
        'https://contoso.sharepoint.com/sites/B',
      ],
      tenantRoot
    );
    expect(issues).toHaveLength(2);
    expect(issues[0].rowIndex).toBe(1);
    expect(issues[1].rowIndex).toBe(2);
  });
});

// ─── validateManagedPropertyCollection ──────────────────────────────────────

const SCHEMA: IManagedProperty[] = [
  { name: 'Author',           type: 'Text',     refinable: true, retrievable: true, sortable: false, queryable: true },
  { name: 'LastModifiedTime', type: 'DateTime', refinable: true, retrievable: true, sortable: true,  queryable: true },
  { name: 'FileType',         type: 'Text',     refinable: true, retrievable: true, sortable: false, queryable: true },
];

describe('validateManagedPropertyCollection', () => {
  it('passes on an empty collection', () => {
    expect(validateManagedPropertyCollection([], SCHEMA)).toEqual([]);
  });

  it('passes silently when schema is undefined (cold cache)', () => {
    const items = [{ property: 'BogusProp' }];
    expect(validateManagedPropertyCollection(items, undefined)).toEqual([]);
  });

  it('flags a typo with Did-You-Mean', () => {
    const items = [{ property: 'LastModifedTime' }];  // Modifed vs Modified
    const issues = validateManagedPropertyCollection(items, SCHEMA);
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe('error');
    expect(issues[0].message).toMatch(/Did you mean "LastModifiedTime"/);
    expect(issues[0].rowIndex).toBe(0);
  });

  it('flags an unknown property without a near match as a plain error', () => {
    const items = [{ property: 'TotallyBogus' }];
    const issues = validateManagedPropertyCollection(items, SCHEMA);
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe('error');
    expect(issues[0].message).toMatch(/not a known managed property/i);
  });

  it('skips empty property cells (caller decides if required)', () => {
    const items = [{ property: '' }, { property: 'Author' }];
    expect(validateManagedPropertyCollection(items, SCHEMA)).toEqual([]);
  });

  it('flags non-sortable when requireSortable is set', () => {
    const items = [{ property: 'Author' }];
    const issues = validateManagedPropertyCollection(items, SCHEMA, { requireSortable: true });
    expect(issues).toHaveLength(1);
    expect(issues[0].message).toMatch(/not sortable/i);
  });
});

// ─── validateRefinementFilterCollection ─────────────────────────────────────

describe('validateRefinementFilterCollection', () => {
  it('passes on an empty collection', () => {
    expect(validateRefinementFilterCollection([])).toEqual([]);
  });

  it('passes for a well-formed equals filter', () => {
    const items = [{ property: 'FileType', operator: 'equals', value: 'docx' }];
    expect(validateRefinementFilterCollection(items)).toEqual([]);
  });

  it('passes for a well-formed range filter', () => {
    const items = [{ property: 'Size', operator: 'range', value: '0,1000000' }];
    expect(validateRefinementFilterCollection(items)).toEqual([]);
  });

  it('flags missing operator', () => {
    const items = [{ property: 'FileType', operator: '', value: 'docx' }];
    const issues = validateRefinementFilterCollection(items);
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe('error');
    expect(issues[0].message).toMatch(/operator/i);
  });

  it('flags an unsupported operator', () => {
    const items = [{ property: 'FileType', operator: 'fuzzyish', value: 'docx' }];
    const issues = validateRefinementFilterCollection(items);
    expect(issues).toHaveLength(1);
    expect(issues[0].message).toMatch(/unsupported operator/i);
  });

  it('flags a range operator without comma-separated value', () => {
    const items = [{ property: 'Size', operator: 'range', value: '1000000' }];
    const issues = validateRefinementFilterCollection(items);
    expect(issues).toHaveLength(1);
    expect(issues[0].message).toMatch(/range.*two values|comma/i);
  });

  it('flags a blank value as error', () => {
    const items = [{ property: 'FileType', operator: 'equals', value: '' }];
    const issues = validateRefinementFilterCollection(items);
    expect(issues).toHaveLength(1);
    expect(issues[0].message).toMatch(/value/i);
  });

  it('preserves rowIndex across multiple rows', () => {
    const items = [
      { property: 'FileType', operator: 'equals', value: 'docx' },
      { property: 'Size',     operator: '',       value: '1000' },
      { property: 'Author',   operator: 'range',  value: 'singleValue' },
    ];
    const issues = validateRefinementFilterCollection(items);
    expect(issues).toHaveLength(2);
    expect(issues[0].rowIndex).toBe(1);
    expect(issues[1].rowIndex).toBe(2);
  });
});
