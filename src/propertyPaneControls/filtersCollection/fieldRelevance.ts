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
  | 'audience';

const LIST_LIKE_TYPES: ReadonlyArray<string> = [
  'checkbox',
  'dropdown',
  'people',
  'taxonomy',
  'tagbox',
];

function isListLike(filterType: string): boolean {
  return LIST_LIKE_TYPES.indexOf(filterType) !== -1;
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
