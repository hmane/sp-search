import { validateWebPartConfig, IValidationInput } from '../../src/webparts/spSearchResults/components/configValidation';

const BASE_VALID: IValidationInput = {
  defaultLayout: 'list',
  availableLayouts: ['list', 'compact', 'grid'],
  selectedPropertyColumns: [
    { property: 'Title',            alias: 'Title'    },
    { property: 'Author',           alias: 'Author'   },
    { property: 'LastModifiedTime', alias: 'Modified' },
    { property: 'FileType',         alias: 'Type'     },
  ],
  gridPropertyColumns: [
    { property: 'Author',           alias: 'Author'   },
    { property: 'LastModifiedTime', alias: 'Modified' },
    { property: 'FileType',         alias: 'Type'     },
  ],
  queryTemplate: '{searchTerms}',
};

describe('validateWebPartConfig', () => {
  it('returns no warnings for a valid configuration', () => {
    expect(validateWebPartConfig(BASE_VALID)).toHaveLength(0);
  });

  describe('default layout not in available layouts', () => {
    it('emits an error when defaultLayout is not in availableLayouts', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        defaultLayout: 'card',
        availableLayouts: ['list', 'compact'],
      };
      const warnings = validateWebPartConfig(input);
      expect(warnings.find((w) => w.id === 'layout-default-unavailable')).toBeDefined();
      expect(warnings.find((w) => w.id === 'layout-default-unavailable')?.level).toBe('error');
    });

    it('does not warn when defaultLayout is "list" even if availableLayouts omits it', () => {
      // 'list' is the baseline fallback — not an actionable mismatch
      const input: IValidationInput = {
        ...BASE_VALID,
        defaultLayout: 'list',
        availableLayouts: ['compact', 'grid'],
      };
      const warnings = validateWebPartConfig(input);
      expect(warnings.find((w) => w.id === 'layout-default-unavailable')).toBeUndefined();
    });

    it('does not warn when defaultLayout matches an available layout', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        defaultLayout: 'grid',
        availableLayouts: ['list', 'compact', 'grid'],
      };
      expect(validateWebPartConfig(input).find((w) => w.id === 'layout-default-unavailable')).toBeUndefined();
    });
  });

  describe('DataGrid column validation', () => {
    it('emits an error when grid is in availableLayouts but no columns are configured', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        availableLayouts: ['list', 'grid'],
        gridPropertyColumns: [],
      };
      const warnings = validateWebPartConfig(input);
      expect(warnings.find((w) => w.id === 'grid-no-columns')).toBeDefined();
      expect(warnings.find((w) => w.id === 'grid-no-columns')?.level).toBe('error');
    });

    it('emits an error when defaultLayout is grid but no columns are configured', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        defaultLayout: 'grid',
        availableLayouts: ['list', 'grid'],
        gridPropertyColumns: [],
      };
      const warnings = validateWebPartConfig(input);
      expect(warnings.find((w) => w.id === 'grid-no-columns')).toBeDefined();
    });

    it('emits a warning when grid is enabled but columns contain no meaningful fields', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        availableLayouts: ['list', 'grid'],
        gridPropertyColumns: [{ property: 'CustomField', alias: 'Custom' }],
      };
      const warnings = validateWebPartConfig(input);
      expect(warnings.find((w) => w.id === 'grid-sparse-columns')).toBeDefined();
      expect(warnings.find((w) => w.id === 'grid-sparse-columns')?.level).toBe('warning');
    });

    it('does not warn about sparse columns when grid is not in available layouts', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        availableLayouts: ['list', 'compact'],
        gridPropertyColumns: [{ property: 'CustomField', alias: 'Custom' }],
      };
      expect(validateWebPartConfig(input).find((w) => w.id === 'grid-sparse-columns')).toBeUndefined();
    });
  });

  describe('query template validation', () => {
    it('emits a warning when queryTemplate does not contain {searchTerms}', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        queryTemplate: 'IsDocument:1 FileType:pdf',
      };
      const warnings = validateWebPartConfig(input);
      expect(warnings.find((w) => w.id === 'query-no-searchterms')).toBeDefined();
      expect(warnings.find((w) => w.id === 'query-no-searchterms')?.level).toBe('warning');
    });

    it('does not warn for queryTemplate that contains {searchTerms} (case-insensitive)', () => {
      for (const tmpl of ['{searchTerms}', '{SEARCHTERMS}', '{SearchTerms} IsDocument:1']) {
        const input: IValidationInput = { ...BASE_VALID, queryTemplate: tmpl };
        expect(validateWebPartConfig(input).find((w) => w.id === 'query-no-searchterms')).toBeUndefined();
      }
    });

    it('does not warn when queryTemplate is empty or default', () => {
      for (const tmpl of ['', '{searchTerms}']) {
        const input: IValidationInput = { ...BASE_VALID, queryTemplate: tmpl };
        expect(validateWebPartConfig(input).find((w) => w.id === 'query-no-searchterms')).toBeUndefined();
      }
    });
  });

  describe('Card/Gallery thumbnail validation', () => {
    it('emits a warning when card is enabled but PictureThumbnailURL is missing', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        availableLayouts: ['list', 'card'],
      };
      const warnings = validateWebPartConfig(input);
      expect(warnings.find((w) => w.id === 'card-gallery-no-thumbnail')).toBeDefined();
      expect(warnings.find((w) => w.id === 'card-gallery-no-thumbnail')?.level).toBe('warning');
    });

    it('does not warn when PictureThumbnailURL is in selectedPropertyColumns', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        availableLayouts: ['list', 'card'],
        selectedPropertyColumns: [
          ...BASE_VALID.selectedPropertyColumns,
          { property: 'PictureThumbnailURL', alias: 'Thumbnail' },
        ],
      };
      expect(validateWebPartConfig(input).find((w) => w.id === 'card-gallery-no-thumbnail')).toBeUndefined();
    });

    it('does not warn when selectedPropertyColumns is empty (grid-no-columns takes priority)', () => {
      // When there are no columns at all, the separate grid-no-columns warning is more actionable
      const input: IValidationInput = {
        ...BASE_VALID,
        availableLayouts: ['list', 'card'],
        selectedPropertyColumns: [],
      };
      expect(validateWebPartConfig(input).find((w) => w.id === 'card-gallery-no-thumbnail')).toBeUndefined();
    });
  });

  describe('People layout profile fields validation', () => {
    it('emits a warning when people layout is enabled but no profile fields are configured', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        availableLayouts: ['people'],
        selectedPropertyColumns: [{ property: 'Title', alias: 'Name' }],
      };
      const warnings = validateWebPartConfig(input);
      expect(warnings.find((w) => w.id === 'people-no-profile-fields')).toBeDefined();
    });

    it('does not warn when WorkEmail is present', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        availableLayouts: ['people'],
        selectedPropertyColumns: [
          { property: 'Title',     alias: 'Name'  },
          { property: 'WorkEmail', alias: 'Email' },
        ],
      };
      expect(validateWebPartConfig(input).find((w) => w.id === 'people-no-profile-fields')).toBeUndefined();
    });
  });

  describe('managed property name validation', () => {
    it('emits an error for property names containing spaces', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        selectedPropertyColumns: [
          { property: 'Title',     alias: 'Title' },
          { property: 'My Prop',   alias: 'Bad'  },
        ],
      };
      const warnings = validateWebPartConfig(input);
      expect(warnings.find((w) => w.id === 'invalid-property-names')).toBeDefined();
      expect(warnings.find((w) => w.id === 'invalid-property-names')?.level).toBe('error');
    });

    it('accepts standard managed property name formats', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        selectedPropertyColumns: [
          { property: 'Title',                 alias: 'Title'   },
          { property: 'LastModifiedTime',      alias: 'Date'    },
          { property: 'ows_taxId_Department',  alias: 'Dept'    },
          { property: 'SPS-Skills',            alias: 'Skills'  },
          { property: 'HitHighlightedSummary', alias: 'Summary' },
        ],
      };
      expect(validateWebPartConfig(input).find((w) => w.id === 'invalid-property-names')).toBeUndefined();
    });

    it('emits an error for property names starting with a digit', () => {
      const input: IValidationInput = {
        ...BASE_VALID,
        selectedPropertyColumns: [{ property: '1BadProp', alias: 'Bad' }],
      };
      expect(validateWebPartConfig(input).find((w) => w.id === 'invalid-property-names')).toBeDefined();
    });
  });
});
