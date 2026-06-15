import * as React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import {
  renderText,
  renderRichText,
  renderNumber,
  renderFileSize,
  renderBoolean,
  renderTags,
  renderUrl,
  renderFileType,
  cleanSearchResultDisplayText,
} from '../../../src/webparts/spSearchResults/components/renderCell';
import type { IColumnConfigItem } from '../../../src/webparts/spSearchResults/components/ColumnConfigField/columnConfig';

/**
 * Stream B / Phase 2 — snapshot-light tests for the new cell renderers.
 *
 * The renderers are pure React-element factories. These tests render each to
 * static markup and assert a few high-signal substrings. Full DOM behaviour
 * tests are integration-level and out of scope per the spec.
 */

function col(overrides: Partial<IColumnConfigItem> = {}): IColumnConfigItem {
  return {
    uniqueId: 't-' + (overrides.property || 'p'),
    property: overrides.property || 'TestField',
    alias: overrides.alias || 'Test',
    visibility: 'defaultOn',
    renderer: '',
    ...overrides,
  };
}

function html(element: React.ReactElement): string {
  return renderToStaticMarkup(element);
}

describe('renderCell — Stream B / Phase 2', () => {
  describe('renderText', () => {
    it('emits the string value', () => {
      expect(html(renderText('hello world', col()))).toContain('hello world');
    });

    it('strips SharePoint calculated value type prefixes', () => {
      const out = html(renderText('string;#Electronic Bank Statements', col()));
      expect(out).toContain('Electronic Bank Statements');
      expect(out).not.toContain('string;#');
    });

    it('does NOT strip values that merely contain a non-type prefix before ;#', () => {
      // "Approved" is not a calculated-column output type, so this is a real
      // value (e.g. a multi-value field), not a type prefix — keep it intact.
      const out = html(renderText('Approved;#Rejected', col()));
      expect(out).toContain('Approved;#Rejected');
    });

    it('emits the muted dash placeholder for empty values', () => {
      expect(html(renderText('', col()))).toContain('--');
    });

    it('truncates when maxLength is set + appends ellipsis', () => {
      const out = html(renderText('a very long text indeed', col({ renderer: 'text', maxLength: 8 })));
      expect(out).toContain('a very l…');
    });

    it('does NOT truncate when maxLength is undefined', () => {
      const out = html(renderText('a very long text indeed', col({ renderer: 'text' })));
      expect(out).toContain('a very long text indeed');
    });
  });

  describe('renderRichText', () => {
    it('preserves safe inline HTML tags', () => {
      const out = html(renderRichText('Hello <strong>world</strong>', col({ renderer: 'richText' })));
      expect(out).toContain('<strong>world</strong>');
    });

    it('strips disallowed tags like <script>', () => {
      const out = html(renderRichText('safe <script>alert(1)</script>', col({ renderer: 'richText' })));
      expect(out).not.toContain('<script>');
      expect(out).toContain('safe');
    });

    it('emits muted dash for empty', () => {
      expect(html(renderRichText('', col()))).toContain('--');
    });
  });

  describe('renderNumber', () => {
    it('formats numbers using locale grouping', () => {
      // Default jsdom locale is en-US; 1,234,567 expected.
      const out = html(renderNumber(1234567, col({ renderer: 'number' })));
      expect(out).toContain('1,234,567');
    });

    it('coerces numeric strings', () => {
      expect(html(renderNumber('42', col({ renderer: 'number' })))).toContain('42');
    });

    it('emits muted dash for non-numeric input', () => {
      expect(html(renderNumber('not-a-number', col({ renderer: 'number' })))).toContain('--');
    });

    it('emits muted dash for null', () => {
      expect(html(renderNumber(null, col({ renderer: 'number' })))).toContain('--');
    });
  });

  describe('renderFileSize', () => {
    it('formats bytes as KB/MB/GB', () => {
      const out = html(renderFileSize(1024 * 1024 * 5, col({ renderer: 'fileSize' })));
      // formatFileSize from documentTitleUtils returns "5 MB" or "5.0 MB" style.
      expect(out).toMatch(/5(\.\d+)?\s*MB/);
    });

    it('emits muted dash for zero / missing', () => {
      expect(html(renderFileSize(0, col({ renderer: 'fileSize' })))).toContain('--');
    });
  });

  describe('renderBoolean', () => {
    it('renders a checkmark icon for true', () => {
      const out = html(renderBoolean(true, col({ renderer: 'boolean' })));
      // Fluent renders Icon as an <i> with iconName-derived data-icon-name
      expect(out).toMatch(/Checkmark|CheckMark/);
    });

    it('renders a cross icon for false', () => {
      const out = html(renderBoolean(false, col({ renderer: 'boolean' })));
      expect(out).toMatch(/Cancel|StatusErrorFull/);
    });

    it('treats "true" / "yes" strings as true', () => {
      expect(html(renderBoolean('true', col({ renderer: 'boolean' })))).toMatch(/Checkmark|CheckMark/);
      expect(html(renderBoolean('Yes', col({ renderer: 'boolean' })))).toMatch(/Checkmark|CheckMark/);
    });

    it('emits muted dash for null', () => {
      expect(html(renderBoolean(null, col({ renderer: 'boolean' })))).toContain('--');
    });
  });

  describe('renderTags', () => {
    it('splits a comma-separated string into pills when separator=pill', () => {
      const out = html(renderTags('one, two, three', col({ renderer: 'tags', multiValueSeparator: 'pill' })));
      expect(out).toContain('one');
      expect(out).toContain('two');
      expect(out).toContain('three');
    });

    it('renders array values joined by a comma when separator=comma', () => {
      const out = html(renderTags(['alpha', 'beta', 'gamma'], col({ renderer: 'tags', multiValueSeparator: 'comma' })));
      expect(out).toContain('alpha, beta, gamma');
    });

    it('strips SharePoint calculated prefixes before splitting tag text', () => {
      const out = html(renderTags('string;#Electronic Bank Statements', col({ renderer: 'tags', multiValueSeparator: 'comma' })));
      expect(out).toContain('Electronic Bank Statements');
      expect(out).not.toContain('string;#');
    });

    it('renders array values one per line when separator=newline', () => {
      const out = html(renderTags(['alpha', 'beta'], col({ renderer: 'tags', multiValueSeparator: 'newline' })));
      // newlines or <br> markers
      expect(out).toMatch(/alpha[\s\S]*beta/);
    });

    it('emits muted dash for empty', () => {
      expect(html(renderTags([], col({ renderer: 'tags', multiValueSeparator: 'comma' })))).toContain('--');
    });
  });

  describe('renderUrl', () => {
    it('renders the value as a clickable anchor', () => {
      const out = html(renderUrl('https://example.com/path', col({ renderer: 'url' })));
      expect(out).toContain('href="https://example.com/path"');
      expect(out).toContain('https://example.com/path');
    });

    it('respects maxLength for the visible text', () => {
      const out = html(renderUrl('https://example.com/very/long/path/here', col({ renderer: 'url', maxLength: 18 })));
      // Truncated visible label, but href still full.
      expect(out).toContain('href="https://example.com/very/long/path/here"');
      expect(out).toContain('…');
    });

    it('emits muted dash for empty', () => {
      expect(html(renderUrl('', col({ renderer: 'url' })))).toContain('--');
    });
  });

  describe('renderFileType', () => {
    it('renders uppercase extension label', () => {
      expect(html(renderFileType('pdf', col({ renderer: 'fileType' })))).toContain('PDF');
    });

    it('emits muted dash for empty', () => {
      expect(html(renderFileType('', col({ renderer: 'fileType' })))).toContain('--');
    });
  });

  describe('cleanSearchResultDisplayText', () => {
    it('strips each documented calculated-column output type', () => {
      expect(cleanSearchResultDisplayText('string;#Electronic Bank Statements')).toBe('Electronic Bank Statements');
      expect(cleanSearchResultDisplayText('float;#42.5')).toBe('42.5');
      expect(cleanSearchResultDisplayText('datetime;#2024-01-01T00:00:00Z')).toBe('2024-01-01T00:00:00Z');
      expect(cleanSearchResultDisplayText('boolean;#1')).toBe('1');
    });

    it('keeps multi-line calculated values whole (does not stop at the newline)', () => {
      expect(cleanSearchResultDisplayText('string;#Line one\nLine two')).toBe('Line one\nLine two');
    });

    it('leaves values intact when the prefix is not a calculated type', () => {
      // Real multi-value / data values that happen to contain ";#" must survive.
      expect(cleanSearchResultDisplayText('Approved;#Rejected')).toBe('Approved;#Rejected');
      expect(cleanSearchResultDisplayText('Engineering;#Finance')).toBe('Engineering;#Finance');
      expect(cleanSearchResultDisplayText('1;#Engineering;#2;#Finance')).toBe('1;#Engineering;#2;#Finance');
    });

    it('leaves plain values and other SharePoint encodings untouched', () => {
      expect(cleanSearchResultDisplayText('Just a title')).toBe('Just a title');
      expect(cleanSearchResultDisplayText('GP0|#guid;Label')).toBe('GP0|#guid;Label');
      expect(cleanSearchResultDisplayText('')).toBe('');
    });
  });
});
