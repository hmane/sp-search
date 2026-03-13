/**
 * Admin-time (edit-mode) property validation for the Search Results web part.
 *
 * All checks are purely structural — no API calls, no async operations.
 * Call validateWebPartConfig() and render the returned warnings in a
 * MessageBar block when isEditMode is true.
 */

export type ConfigWarningLevel = 'error' | 'warning' | 'info';

export interface IConfigWarning {
  /** Stable key for React reconciliation */
  id: string;
  level: ConfigWarningLevel;
  message: string;
}

export interface IValidationInput {
  /** Layout that will be active when the page loads */
  defaultLayout: string;
  /** Layouts currently enabled in the toolbar */
  availableLayouts: string[];
  /** Admin-configured display columns */
  selectedPropertyColumns: Array<{ property: string; alias: string }>;
  /** Data Grid metadata columns. Title is fixed and excluded from this list. */
  gridPropertyColumns?: Array<{ property: string; alias: string }>;
  /** KQL query template from property pane / store */
  queryTemplate: string;
}

/** Managed property names must be alphanumeric + underscore only (no spaces, hyphens, dots). */
const VALID_MANAGED_PROPERTY = /^[A-Za-z][A-Za-z0-9_\-.]*$/;

/**
 * Properties that the Card and Gallery layouts use for thumbnails.
 * Warn if these layouts are available but neither property is configured.
 */
const THUMBNAIL_PROPERTIES = ['PictureThumbnailURL', 'OWSTaxIdMetadataAllTagsInfo'];

/**
 * Properties expected for a useful People layout.
 * At least one should be in selectedPropertyColumns when People is enabled.
 */
const PEOPLE_PROPERTIES = ['WorkEmail', 'JobTitle', 'Department', 'WorkPhone'];

/**
 * Properties expected for DataGrid to render more than a bare title column.
 * If Grid is active and none of these are configured, the grid is nearly empty.
 */
const GRID_MEANINGFUL_PROPERTIES = [
  'Author', 'LastModifiedTime', 'FileType', 'FileSize', 'Size', 'SiteName', 'Created'
];

export function validateWebPartConfig(input: IValidationInput): IConfigWarning[] {
  const warnings: IConfigWarning[] = [];
  const { defaultLayout, availableLayouts, selectedPropertyColumns, gridPropertyColumns = [], queryTemplate } = input;

  const propNames = selectedPropertyColumns.map((c) => c.property);
  const gridEnabled = availableLayouts.indexOf('grid') >= 0 || defaultLayout === 'grid';
  const cardEnabled = availableLayouts.indexOf('card') >= 0 || defaultLayout === 'card';
  const galleryEnabled = availableLayouts.indexOf('gallery') >= 0 || defaultLayout === 'gallery';
  const peopleEnabled = availableLayouts.indexOf('people') >= 0 || defaultLayout === 'people';

  // ── 1. Default layout not available ──────────────────────────────────────────
  if (defaultLayout !== 'list' && availableLayouts.indexOf(defaultLayout) < 0) {
    warnings.push({
      id: 'layout-default-unavailable',
      level: 'error',
      message:
        'Default layout "' + defaultLayout + '" is not in the available layout set. ' +
        'The page will fall back to List. Enable "' + defaultLayout + '" in the layout toggles, ' +
        'or change the default layout.',
    });
  }

  // ── 2. Grid enabled with no columns ──────────────────────────────────────────
  if (gridEnabled && gridPropertyColumns.length === 0) {
    warnings.push({
      id: 'grid-no-columns',
      level: 'error',
      message:
        'DataGrid is enabled but no grid metadata columns are configured. ' +
        'Choose fields in the Data Grid Columns section, or apply a preset.',
    });
  }

  // ── 3. Grid enabled but only one meaningful column ────────────────────────────
  if (
    gridEnabled &&
    gridPropertyColumns.length > 0 &&
    !GRID_MEANINGFUL_PROPERTIES.some((p) =>
      gridPropertyColumns.some((c) => c.property === p)
    )
  ) {
    warnings.push({
      id: 'grid-sparse-columns',
      level: 'warning',
      message:
        'DataGrid columns do not include common fields like Author, Modified, or File Type. ' +
        'Consider adding these for a more useful grid view.',
    });
  }

  // ── 4. Query template missing {searchTerms} ───────────────────────────────────
  if (
    queryTemplate &&
    queryTemplate.trim() !== '' &&
    queryTemplate.toLowerCase().indexOf('{searchterms}') < 0
  ) {
    warnings.push({
      id: 'query-no-searchterms',
      level: 'warning',
      message:
        'Query template "' + queryTemplate + '" does not include {searchTerms}. ' +
        'User queries will be ignored and the same fixed results will always be returned. ' +
        'This is intentional for browse scenarios; otherwise add {searchTerms}.',
    });
  }

  // ── 5. Card or Gallery enabled without a thumbnail property ──────────────────
  if (
    (cardEnabled || galleryEnabled) &&
    selectedPropertyColumns.length > 0 &&
    !THUMBNAIL_PROPERTIES.some((p) => propNames.indexOf(p) >= 0)
  ) {
    const layoutNames = [
      cardEnabled   ? 'Card'    : '',
      galleryEnabled ? 'Gallery' : '',
    ].filter(Boolean).join('/');
    warnings.push({
      id: 'card-gallery-no-thumbnail',
      level: 'warning',
      message:
        layoutNames + ' layout is enabled but PictureThumbnailURL is not in selected properties. ' +
        'Cards will render without thumbnails. Add PictureThumbnailURL to enable thumbnail images.',
    });
  }

  // ── 6. People layout enabled without people-relevant properties ───────────────
  if (
    peopleEnabled &&
    selectedPropertyColumns.length > 0 &&
    !PEOPLE_PROPERTIES.some((p) => propNames.indexOf(p) >= 0)
  ) {
    warnings.push({
      id: 'people-no-profile-fields',
      level: 'warning',
      message:
        'People layout is enabled but no people profile properties (WorkEmail, JobTitle, ' +
        'Department, WorkPhone) are configured. Apply the People preset for the correct ' +
        'property set, or add profile properties manually.',
    });
  }

  // ── 7. Malformed managed property names ──────────────────────────────────────
  const invalidProps = selectedPropertyColumns
    .map((c) => c.property)
    .filter((name) => name && !VALID_MANAGED_PROPERTY.test(name));

  if (invalidProps.length > 0) {
    warnings.push({
      id: 'invalid-property-names',
      level: 'error',
      message:
        'The following selected properties have names that are not valid managed property ' +
        'identifiers and will return no results: ' +
        invalidProps.map((p) => '"' + p + '"').join(', ') + '. ' +
        'Managed property names must start with a letter and contain only letters, numbers, ' +
        'underscores, hyphens, or dots.',
    });
  }

  return warnings;
}
