import {
  normalizeColumnConfigItem,
  applyColumnPropertySelection,
  IColumnConfigItem,
  ColumnRenderer,
  ColumnVisibility,
  MultiValueSeparator,
} from '../../../src/webparts/spSearchResults/components/ColumnConfigField/columnConfig';

/**
 * TDD tests for the Stream B / Phase 1 column-config normalizer.
 *
 * The normalizer wraps a legacy `{ uniqueId, property }` item or a
 * partially-populated new-shape item with defaults; `renderer: ''` is the
 * migration sentinel that routes through today's auto-detect path.
 */
describe('normalizeColumnConfigItem', () => {
  describe('legacy { uniqueId, property } migration', () => {
    it('preserves uniqueId and property', () => {
      const result = normalizeColumnConfigItem({ uniqueId: 'starter-grid-0', property: 'Author' });
      expect(result.property).toBe('Author');
      expect(result.uniqueId).toBe('starter-grid-0');
    });

    it('defaults alias to property when alias is missing', () => {
      const result = normalizeColumnConfigItem({ uniqueId: 'x', property: 'Author' });
      expect(result.alias).toBe('Author');
    });

    it('defaults visibility to "defaultOn"', () => {
      const result = normalizeColumnConfigItem({ uniqueId: 'x', property: 'Author' });
      expect(result.visibility).toBe('defaultOn');
    });

    it('defaults renderer to "" (auto-detect sentinel)', () => {
      const result = normalizeColumnConfigItem({ uniqueId: 'x', property: 'Author' });
      expect(result.renderer).toBe('');
    });

    it('leaves width / maxLength / seeMoreLink / multiValueSeparator undefined when absent', () => {
      const result = normalizeColumnConfigItem({ uniqueId: 'x', property: 'Author' });
      expect(result.width).toBeUndefined();
      expect(result.maxLength).toBeUndefined();
      expect(result.seeMoreLink).toBeUndefined();
      expect(result.multiValueSeparator).toBeUndefined();
    });

    it('generates a uniqueId when none is supplied', () => {
      const result = normalizeColumnConfigItem({ property: 'Author' });
      expect(result.uniqueId).toMatch(/^col-/);
    });

    it('trims whitespace around the property name', () => {
      const result = normalizeColumnConfigItem({ uniqueId: 'x', property: '  Author  ' });
      expect(result.property).toBe('Author');
    });
  });

  describe('new-shape pass-through', () => {
    it('preserves every explicitly provided field unchanged', () => {
      const raw: IColumnConfigItem = {
        uniqueId: 'col-1',
        property: 'CustomField',
        alias: 'My Field',
        width: 200,
        visibility: 'defaultOff',
        renderer: 'date',
        maxLength: 100,
        seeMoreLink: true,
        multiValueSeparator: 'pill',
      };
      const result = normalizeColumnConfigItem(raw);
      expect(result).toEqual(raw);
    });
  });

  describe('renderer values', () => {
    const valid: ColumnRenderer[] = [
      '',
      'text',
      'richText',
      'date',
      'number',
      'fileSize',
      'persona',
      'tags',
      'boolean',
      'url',
      'fileType',
    ];
    valid.forEach((renderer) => {
      it('accepts renderer="' + renderer + '"', () => {
        const result = normalizeColumnConfigItem({ uniqueId: 'x', property: 'p', renderer });
        expect(result.renderer).toBe(renderer);
      });
    });

    it('falls back to "" for an unknown renderer string', () => {
      const result = normalizeColumnConfigItem({
        uniqueId: 'x',
        property: 'p',
        renderer: 'bogus' as ColumnRenderer,
      });
      expect(result.renderer).toBe('');
    });
  });

  describe('visibility values', () => {
    const valid: ColumnVisibility[] = ['always', 'defaultOn', 'defaultOff'];
    valid.forEach((visibility) => {
      it('accepts visibility="' + visibility + '"', () => {
        const result = normalizeColumnConfigItem({ uniqueId: 'x', property: 'p', visibility });
        expect(result.visibility).toBe(visibility);
      });
    });

    it('falls back to "defaultOn" for an unknown visibility string', () => {
      const result = normalizeColumnConfigItem({
        uniqueId: 'x',
        property: 'p',
        visibility: 'bogus' as ColumnVisibility,
      });
      expect(result.visibility).toBe('defaultOn');
    });
  });

  describe('multiValueSeparator values', () => {
    const valid: MultiValueSeparator[] = ['comma', 'newline', 'semicolon', 'pill'];
    valid.forEach((sep) => {
      it('accepts multiValueSeparator="' + sep + '"', () => {
        const result = normalizeColumnConfigItem({
          uniqueId: 'x',
          property: 'p',
          multiValueSeparator: sep,
        });
        expect(result.multiValueSeparator).toBe(sep);
      });
    });

    it('falls back to undefined for an unknown separator', () => {
      const result = normalizeColumnConfigItem({
        uniqueId: 'x',
        property: 'p',
        multiValueSeparator: 'bogus' as MultiValueSeparator,
      });
      expect(result.multiValueSeparator).toBeUndefined();
    });
  });

  describe('numeric coercion', () => {
    it('treats width=0 as auto (undefined)', () => {
      const result = normalizeColumnConfigItem({ uniqueId: 'x', property: 'p', width: 0 });
      expect(result.width).toBeUndefined();
    });

    it('treats maxLength=0 as no-truncation (undefined)', () => {
      const result = normalizeColumnConfigItem({ uniqueId: 'x', property: 'p', maxLength: 0 });
      expect(result.maxLength).toBeUndefined();
    });

    it('rejects negative width', () => {
      const result = normalizeColumnConfigItem({ uniqueId: 'x', property: 'p', width: -10 });
      expect(result.width).toBeUndefined();
    });

    it('preserves positive width', () => {
      const result = normalizeColumnConfigItem({ uniqueId: 'x', property: 'p', width: 220 });
      expect(result.width).toBe(220);
    });
  });
});

describe('applyColumnPropertySelection', () => {
  it('uses the selected managed property alias when the column alias is still blank', () => {
    const current = normalizeColumnConfigItem({ uniqueId: 'x', property: '' });
    const result = applyColumnPropertySelection(current, {
      key: 'LastModifiedTime',
      text: 'Modified (LastModifiedTime)',
      alias: 'Modified',
    });

    expect(result.property).toBe('LastModifiedTime');
    expect(result.alias).toBe('Modified');
  });

  it('falls back to the managed property name when no alias is available', () => {
    const current = normalizeColumnConfigItem({ uniqueId: 'x', property: '' });
    const result = applyColumnPropertySelection(current, {
      key: 'RefinableString100',
      text: 'RefinableString100',
    });

    expect(result.property).toBe('RefinableString100');
    expect(result.alias).toBe('RefinableString100');
  });

  it('derives the alias from option text when dropdown metadata is not preserved', () => {
    const current = normalizeColumnConfigItem({ uniqueId: 'x', property: '' });
    const result = applyColumnPropertySelection(current, {
      key: 'LastModifiedTime',
      text: 'Modified (LastModifiedTime)',
    });

    expect(result.property).toBe('LastModifiedTime');
    expect(result.alias).toBe('Modified');
  });

  it('prefers the displayed alias when structured alias was defaulted to the property name', () => {
    const current = normalizeColumnConfigItem({ uniqueId: 'x', property: '' });
    const result = applyColumnPropertySelection(current, {
      key: 'RefinableString100',
      text: 'Document Type (RefinableString100)',
      alias: 'RefinableString100',
    });

    expect(result.property).toBe('RefinableString100');
    expect(result.alias).toBe('Document Type');
  });

  it('preserves a custom alias when the selected property changes', () => {
    const current = normalizeColumnConfigItem({
      uniqueId: 'x',
      property: 'Author',
      alias: 'Owner',
    });
    const result = applyColumnPropertySelection(current, {
      key: 'LastModifiedTime',
      text: 'Modified (LastModifiedTime)',
      alias: 'Modified',
    });

    expect(result.property).toBe('LastModifiedTime');
    expect(result.alias).toBe('Owner');
  });
});
