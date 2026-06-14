/**
 * Sprint 5 / Filter editor UX redesign — pure logic for the master-detail
 * Configure Refiners panel. Decides which sections + sub-fields are
 * relevant for a given filterType so the form only surfaces inputs the
 * user can actually use.
 */

import {
  getRelevantSections,
  isFieldRelevant,
} from '../../src/propertyPaneControls/filtersCollection/fieldRelevance';
import type {
  FilterEditorSection,
  FilterEditorField,
} from '../../src/propertyPaneControls/filtersCollection/fieldRelevance';

describe('getRelevantSections', () => {
  it('returns Basic + Conditional + Audience for every filter type', () => {
    const types = ['checkbox', 'dropdown', 'daterange', 'text', 'slider', 'people', 'taxonomy', 'tagbox', 'toggle'];
    for (const t of types) {
      const sections = getRelevantSections(t);
      expect(sections).toContain('basic' as FilterEditorSection);
      expect(sections).toContain('conditional' as FilterEditorSection);
      expect(sections).toContain('audience' as FilterEditorSection);
    }
  });

  it('includes Display + Behavior for list-style refiners', () => {
    const listLike = ['checkbox', 'dropdown', 'people', 'taxonomy', 'tagbox'];
    for (const t of listLike) {
      const sections = getRelevantSections(t);
      expect(sections).toContain('display' as FilterEditorSection);
      expect(sections).toContain('behavior' as FilterEditorSection);
    }
  });

  it('omits Display + Behavior for non-list refiners', () => {
    const nonList = ['daterange', 'text', 'slider'];
    for (const t of nonList) {
      const sections = getRelevantSections(t);
      expect(sections).not.toContain('display' as FilterEditorSection);
      expect(sections).not.toContain('behavior' as FilterEditorSection);
    }
  });

  it('includes the Toggle Labels section only for toggle', () => {
    expect(getRelevantSections('toggle')).toContain('toggleLabels' as FilterEditorSection);
    expect(getRelevantSections('checkbox')).not.toContain('toggleLabels' as FilterEditorSection);
    expect(getRelevantSections('daterange')).not.toContain('toggleLabels' as FilterEditorSection);
  });

  it('falls back to Basic + Conditional + Audience for an unknown filterType', () => {
    const sections = getRelevantSections('mystery-type-not-registered');
    expect(sections).toEqual(['basic', 'conditional', 'audience']);
  });
});

describe('isFieldRelevant', () => {
  it('treats the four always-on fields as relevant for every type', () => {
    const alwaysOn: FilterEditorField[] = ['managedProperty', 'displayName', 'urlAlias', 'filterType'];
    const types = ['checkbox', 'dropdown', 'daterange', 'text', 'slider', 'people', 'taxonomy', 'tagbox', 'toggle'];
    for (const t of types) {
      for (const f of alwaysOn) {
        expect(isFieldRelevant(f, t)).toBe(true);
      }
    }
  });

  it('hides toggle-only fields from non-toggle types', () => {
    const toggleOnly: FilterEditorField[] = ['trueLabel', 'falseLabel', 'invertBoolean'];
    for (const f of toggleOnly) {
      expect(isFieldRelevant(f, 'checkbox')).toBe(false);
      expect(isFieldRelevant(f, 'daterange')).toBe(false);
      expect(isFieldRelevant(f, 'toggle')).toBe(true);
    }
  });

  it('hides list-only fields from non-list types', () => {
    const listOnly: FilterEditorField[] = ['operator', 'maxValues', 'showCount', 'hideZeroCountValues'];
    for (const f of listOnly) {
      expect(isFieldRelevant(f, 'daterange')).toBe(false);
      expect(isFieldRelevant(f, 'text')).toBe(false);
      expect(isFieldRelevant(f, 'slider')).toBe(false);
      expect(isFieldRelevant(f, 'toggle')).toBe(false);
      expect(isFieldRelevant(f, 'checkbox')).toBe(true);
    }
  });

  it('hides sortBy/sortDirection from people (people refiners are person-ranked, not sortable)', () => {
    expect(isFieldRelevant('sortBy', 'people')).toBe(false);
    expect(isFieldRelevant('sortDirection', 'people')).toBe(false);
    expect(isFieldRelevant('sortBy', 'checkbox')).toBe(true);
    expect(isFieldRelevant('sortDirection', 'checkbox')).toBe(true);
  });

  it('hides multiValues from dropdown (dropdown is single-select by definition)', () => {
    expect(isFieldRelevant('multiValues', 'dropdown')).toBe(false);
    expect(isFieldRelevant('multiValues', 'checkbox')).toBe(true);
    expect(isFieldRelevant('multiValues', 'tagbox')).toBe(true);
  });

  it('always exposes defaultExpanded — every section header is collapsible', () => {
    const types = ['checkbox', 'daterange', 'text', 'slider', 'toggle'];
    for (const t of types) {
      expect(isFieldRelevant('defaultExpanded', t)).toBe(true);
    }
  });

  it('exposes dependsOn for every type', () => {
    expect(isFieldRelevant('dependsOn', 'checkbox')).toBe(true);
    expect(isFieldRelevant('dependsOn', 'toggle')).toBe(true);
    expect(isFieldRelevant('dependsOn', 'daterange')).toBe(true);
  });

  it('exposes audience for every type', () => {
    expect(isFieldRelevant('audience', 'checkbox')).toBe(true);
    expect(isFieldRelevant('audience', 'toggle')).toBe(true);
    expect(isFieldRelevant('audience', 'daterange')).toBe(true);
  });

  it('exposes dataType + valueSplitDelimiter only for checkbox / tagbox / dropdown / text', () => {
    const fields: FilterEditorField[] = ['dataType', 'valueSplitDelimiter'];
    const shown = ['checkbox', 'tagbox', 'dropdown', 'text'];
    const hidden = ['daterange', 'slider', 'people', 'taxonomy', 'toggle'];
    for (const f of fields) {
      for (const t of shown) {
        expect(isFieldRelevant(f, t)).toBe(true);
      }
      for (const t of hidden) {
        expect(isFieldRelevant(f, t)).toBe(false);
      }
    }
  });

  it('includes the Data format section only for data-format-relevant types', () => {
    const shown = ['checkbox', 'tagbox', 'dropdown', 'text'];
    const hidden = ['daterange', 'slider', 'people', 'taxonomy', 'toggle'];
    for (const t of shown) {
      expect(getRelevantSections(t)).toContain('dataFormat' as FilterEditorSection);
    }
    for (const t of hidden) {
      expect(getRelevantSections(t)).not.toContain('dataFormat' as FilterEditorSection);
    }
  });
});
