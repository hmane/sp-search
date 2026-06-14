/**
 * Sprint 5 / Filter editor UX redesign — pure logic for the master-detail
 * Configure Refiners panel. Tells the React form which sections + fields
 * are relevant for a given filterType so the form only surfaces inputs
 * the user can actually use.
 *
 * Kept in its own module + 100% pure so the visibility matrix can be
 * unit-tested without spinning up React.
 */

export type FilterEditorSection =
  | 'basic'
  | 'display'
  | 'behavior'
  | 'toggleLabels'
  | 'dataFormat'
  | 'conditional'
  | 'audience';

export type FilterEditorField =
  | 'managedProperty'
  | 'displayName'
  | 'urlAlias'
  | 'filterType'
  | 'operator'
  | 'maxValues'
  | 'showCount'
  | 'defaultExpanded'
  | 'sortBy'
  | 'sortDirection'
  | 'multiValues'
  | 'hideZeroCountValues'
  | 'dependsOn'
  | 'showWhenParentHasValue'
  | 'resetWhenParentChanges'
  | 'trueLabel'
  | 'falseLabel'
  | 'invertBoolean'
  | 'defaultValue'
  | 'dataType'
  | 'valueSplitDelimiter'
  | 'audience';

const LIST_LIKE_TYPES: ReadonlyArray<string> = [
  'checkbox',
  'dropdown',
  'people',
  'taxonomy',
  'tagbox',
];

/**
 * Filter types where the underlying managed-property data shape varies enough
 * that admins may need to override the auto-detect heuristic and/or split
 * raw values on a delimiter. Date / slider / toggle / people / taxonomy refiners
 * have fixed value shapes and don't need this knob.
 */
const DATA_FORMAT_TYPES: ReadonlyArray<string> = [
  'checkbox',
  'tagbox',
  'dropdown',
  'text',
];

function isListLike(filterType: string): boolean {
  return LIST_LIKE_TYPES.indexOf(filterType) !== -1;
}

function hasDataFormat(filterType: string): boolean {
  return DATA_FORMAT_TYPES.indexOf(filterType) !== -1;
}

export function getRelevantSections(filterType: string): FilterEditorSection[] {
  const sections: FilterEditorSection[] = ['basic'];
  if (isListLike(filterType)) {
    sections.push('display');
    sections.push('behavior');
  }
  if (filterType === 'toggle') {
    sections.push('toggleLabels');
  }
  if (hasDataFormat(filterType)) {
    sections.push('dataFormat');
  }
  sections.push('conditional');
  sections.push('audience');
  return sections;
}

export function isFieldRelevant(field: FilterEditorField, filterType: string): boolean {
  switch (field) {
    // ── Always relevant ─────────────────────────────────────
    case 'managedProperty':
    case 'displayName':
    case 'urlAlias':
    case 'filterType':
    case 'defaultExpanded':
    case 'dependsOn':
    case 'showWhenParentHasValue':
    case 'resetWhenParentChanges':
    case 'audience':
      return true;

    // ── Toggle-only ────────────────────────────────────────
    case 'trueLabel':
    case 'falseLabel':
    case 'invertBoolean':
    case 'defaultValue':
      return filterType === 'toggle';

    // ── Data-format only (checkbox / tagbox / dropdown / text) ──
    case 'dataType':
    case 'valueSplitDelimiter':
      return hasDataFormat(filterType);

    // ── List-like only ─────────────────────────────────────
    case 'operator':
    case 'maxValues':
    case 'showCount':
    case 'hideZeroCountValues':
      return isListLike(filterType);

    // ── List-like, but people refiners are person-ranked, not sortable ──
    case 'sortBy':
    case 'sortDirection':
      return isListLike(filterType) && filterType !== 'people';

    // ── Dropdown is single-select by definition; multiValues isn't meaningful there ──
    case 'multiValues':
      return isListLike(filterType) && filterType !== 'dropdown';

    default: {
      const _exhaustive: never = field;
      return Boolean(_exhaustive);
    }
  }
}
