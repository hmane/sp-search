import {
  validateManagedProperty,
  suggestCloseManagedProperty,
} from '../../src/libraries/spSearchStore/utils/managedPropertyValidation';
import type { IManagedProperty } from '../../src/libraries/spSearchStore/interfaces/ISearchDataProvider';

/**
 * T4.D3 — validate managed-property values against the cached schema.
 * Catches the `LastModifedTime` typo before save and rejects non-sortable
 * properties on `collapseSpecification`.
 */

function prop(name: string, flags: Partial<IManagedProperty> = {}): IManagedProperty {
  return {
    name,
    type: 'Text',
    queryable: true,
    retrievable: true,
    refinable: false,
    sortable: false,
    ...flags,
  };
}

const SAMPLE_SCHEMA: IManagedProperty[] = [
  prop('LastModifiedTime', { sortable: true, type: 'DateTime' }),
  prop('Created', { sortable: true, type: 'DateTime' }),
  prop('Author'),
  prop('AuthorOWSUSER', { refinable: true }),
  prop('FileType', { refinable: true }),
  prop('Size', { sortable: true, type: 'Integer' }),
  prop('Title', { sortable: true }),
];

describe('suggestCloseManagedProperty', () => {
  it('returns the closest name when input is a one-char typo', () => {
    expect(suggestCloseManagedProperty('LastModifedTime', SAMPLE_SCHEMA)).toBe('LastModifiedTime');
  });

  it('returns the closest name when input is two chars off', () => {
    expect(suggestCloseManagedProperty('LastModfiedTme', SAMPLE_SCHEMA)).toBe('LastModifiedTime');
  });

  it('is case-insensitive on the input (lowercase + typo still finds LastModifiedTime)', () => {
    expect(suggestCloseManagedProperty('lastmodifedtime', SAMPLE_SCHEMA)).toBe('LastModifiedTime');
  });

  it('returns undefined for input with no near match', () => {
    expect(suggestCloseManagedProperty('TotallyUnrelatedField', SAMPLE_SCHEMA)).toBeUndefined();
  });

  it('returns undefined for an empty schema', () => {
    expect(suggestCloseManagedProperty('Anything', [])).toBeUndefined();
  });

  it('returns undefined when the input already exists in the schema (no suggestion needed)', () => {
    expect(suggestCloseManagedProperty('LastModifiedTime', SAMPLE_SCHEMA)).toBeUndefined();
  });
});

describe('validateManagedProperty — basic', () => {
  it('passes when the value matches a property in the schema', () => {
    expect(validateManagedProperty('LastModifiedTime', SAMPLE_SCHEMA)).toEqual({ valid: true });
  });

  it('passes case-insensitively', () => {
    expect(validateManagedProperty('lastmodifiedtime', SAMPLE_SCHEMA)).toEqual({ valid: true });
  });

  it('passes silently with no schema available (cache cold)', () => {
    expect(validateManagedProperty('AnythingGoes', [])).toEqual({ valid: true });
    expect(validateManagedProperty('AnythingGoes', undefined)).toEqual({ valid: true });
  });

  it('passes for empty/whitespace input (caller decides if required)', () => {
    expect(validateManagedProperty('', SAMPLE_SCHEMA)).toEqual({ valid: true });
    expect(validateManagedProperty('   ', SAMPLE_SCHEMA)).toEqual({ valid: true });
  });

  it('returns a did-you-mean suggestion for the typo case from the audit acceptance signal', () => {
    const result = validateManagedProperty('LastModifedTime', SAMPLE_SCHEMA);
    expect(result.valid).toBe(false);
    // ts-jest non-strict mode doesn't narrow the discriminated union on !valid.
    const failed = result as { valid: false; message: string };
    if (!failed.valid) {
      expect(failed.message).toMatch(/did you mean/i);
      expect(failed.message).toContain('LastModifiedTime');
    }
  });

  it('returns "unknown property" when the value has no close match', () => {
    const result = validateManagedProperty('CompletelyMadeUp', SAMPLE_SCHEMA);
    expect(result.valid).toBe(false);
    // ts-jest non-strict mode doesn't narrow the discriminated union on !valid.
    const failed = result as { valid: false; message: string };
    if (!failed.valid) {
      expect(failed.message).toMatch(/not a known managed property/i);
    }
  });
});

describe('validateManagedProperty — requireSortable (T4.D3 collapse spec acceptance)', () => {
  it('passes when the value is sortable', () => {
    expect(validateManagedProperty('LastModifiedTime', SAMPLE_SCHEMA, { requireSortable: true })).toEqual({ valid: true });
  });

  it('rejects when the property exists but is NOT sortable', () => {
    const result = validateManagedProperty('FileType', SAMPLE_SCHEMA, { requireSortable: true });
    expect(result.valid).toBe(false);
    // ts-jest non-strict mode doesn't narrow the discriminated union on !valid.
    const failed = result as { valid: false; message: string };
    if (!failed.valid) {
      expect(failed.message).toMatch(/sortable/i);
    }
  });

  it('still returns did-you-mean for typo + sortable requirement', () => {
    const result = validateManagedProperty('LastModifedTime', SAMPLE_SCHEMA, { requireSortable: true });
    expect(result.valid).toBe(false);
    // ts-jest non-strict mode doesn't narrow the discriminated union on !valid.
    const failed = result as { valid: false; message: string };
    if (!failed.valid) {
      expect(failed.message).toContain('LastModifiedTime');
    }
  });
});

describe('validateManagedProperty — requireRefinable', () => {
  it('rejects when the property exists but is not refinable', () => {
    const result = validateManagedProperty('Author', SAMPLE_SCHEMA, { requireRefinable: true });
    expect(result.valid).toBe(false);
    // ts-jest non-strict mode doesn't narrow the discriminated union on !valid.
    const failed = result as { valid: false; message: string };
    if (!failed.valid) {
      expect(failed.message).toMatch(/refinable/i);
    }
  });

  it('passes when the property is refinable', () => {
    expect(validateManagedProperty('AuthorOWSUSER', SAMPLE_SCHEMA, { requireRefinable: true })).toEqual({ valid: true });
  });
});
