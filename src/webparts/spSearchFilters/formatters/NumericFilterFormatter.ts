import type { IFilterValueFormatter, IFilterConfig } from '@interfaces/index';

/**
 * Format bytes to human-readable file size.
 */
function formatFileSize(bytes: number): string {
  if (bytes < 1024) {
    return bytes + ' B';
  }
  if (bytes < 1024 * 1024) {
    return (bytes / 1024).toFixed(1) + ' KB';
  }
  if (bytes < 1024 * 1024 * 1024) {
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }
  return (bytes / (1024 * 1024 * 1024)).toFixed(1) + ' GB';
}

/**
 * Parse an FQL range token into min/max numbers.
 * Handles formats: range(decimal(N), decimal(N)), range(N, max), range(min, N)
 */
function parseRange(token: string): { min: number | undefined; max: number | undefined } | undefined {
  const regex: RegExp = /^range\((?:decimal\()?([^),]+)\)?,\s*(?:decimal\()?([^),]+)\)?\)$/;
  const match: RegExpMatchArray | null = token.match(regex);
  if (!match) {
    return undefined;
  }
  const minStr: string = match[1].trim();
  const maxStr: string = match[2].trim();
  const minVal: number | undefined = minStr === 'min' ? undefined : parseFloat(minStr);
  const maxVal: number | undefined = maxStr === 'max' ? undefined : parseFloat(maxStr);
  return { min: minVal, max: maxVal };
}

/**
 * NumericFilterFormatter — formats numeric range refinement values.
 *
 * Handles:
 * - File size (bytes → KB/MB/GB) when managed property name contains 'size'
 * - FQL range(decimal(min), decimal(max)) token generation
 * - Human-readable display with appropriate number formatting
 */
export const NumericFilterFormatter: IFilterValueFormatter = {
  id: 'slider',

  formatForDisplay: function (rawValue: string, config: IFilterConfig): string {
    const parsed = parseRange(rawValue);
    if (!parsed) {
      return rawValue;
    }

    const isFileSize: boolean = config.managedProperty.toLowerCase().indexOf('size') >= 0;
    const formatter = isFileSize ? formatFileSize : function (n: number): string {
      return n.toLocaleString();
    };

    if (parsed.min === undefined && parsed.max !== undefined && !isNaN(parsed.max)) {
      return '< ' + formatter(parsed.max);
    }
    if (parsed.min !== undefined && !isNaN(parsed.min) && parsed.max === undefined) {
      return '> ' + formatter(parsed.min);
    }
    if (parsed.min !== undefined && !isNaN(parsed.min) && parsed.max !== undefined && !isNaN(parsed.max)) {
      return formatter(parsed.min) + ' – ' + formatter(parsed.max);
    }

    return rawValue;
  },

  formatForQuery: function (displayValue: unknown, _config: IFilterConfig): string {
    // Expects { min: number, max: number } or a pre-formatted FQL string
    if (typeof displayValue === 'string') {
      return displayValue;
    }
    const range = displayValue as { min?: number; max?: number };
    const minPart: string = range.min !== undefined ? 'decimal(' + range.min + ')' : 'min';
    const maxPart: string = range.max !== undefined ? 'decimal(' + range.max + ')' : 'max';
    return 'range(' + minPart + ', ' + maxPart + ')';
  },

  formatForUrl: function (rawValue: string): string {
    // Convert range(decimal(100), decimal(500)) → "100:500"
    const parsed = parseRange(rawValue);
    if (!parsed) {
      return rawValue;
    }
    const minStr: string = parsed.min !== undefined ? String(parsed.min) : '';
    const maxStr: string = parsed.max !== undefined ? String(parsed.max) : '';
    return minStr + ':' + maxStr;
  },

  parseFromUrl: function (urlValue: string): string {
    // Convert "100:500" → range(decimal(100), decimal(500))
    const parts: string[] = urlValue.split(':');
    if (parts.length !== 2) {
      return urlValue;
    }
    const minVal: number = parts[0] ? parseFloat(parts[0]) : NaN;
    const maxVal: number = parts[1] ? parseFloat(parts[1]) : NaN;
    const minPart: string = !isNaN(minVal) ? 'decimal(' + minVal + ')' : 'min';
    const maxPart: string = !isNaN(maxVal) ? 'decimal(' + maxVal + ')' : 'max';
    return 'range(' + minPart + ', ' + maxPart + ')';
  }
};
