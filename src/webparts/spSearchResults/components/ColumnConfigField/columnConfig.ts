/**
 * Stream B / Phase 1 — column-config schema for the DataGrid layout.
 *
 * Replaces the legacy `{ uniqueId, property }` shape on
 * `gridPropertiesCollection`. The renderer union covers Phases 1-3 so a
 * Phase-2 saved value (`richText`, `tags`, etc.) round-trips through a
 * Phase-1 build without being reset; the Phase-1 DataGrid renderer falls
 * back to auto-detect for renderer values it doesn't yet handle.
 *
 * Spec: docs/superpowers/specs/2026-05-13-stream-b-column-config-design.md
 */

export type ColumnVisibility = 'always' | 'defaultOn' | 'defaultOff';

/** Empty string is the migration sentinel — routes through today's auto-detect path. */
export type ColumnRenderer =
  | ''
  | 'text'
  | 'richText'
  | 'date'
  | 'number'
  | 'fileSize'
  | 'persona'
  | 'tags'
  | 'boolean'
  | 'url'
  | 'fileType';

export type MultiValueSeparator = 'comma' | 'newline' | 'semicolon' | 'pill';

export interface IColumnConfigItem {
  uniqueId: string;
  property: string;
  alias: string;
  width?: number;
  visibility: ColumnVisibility;
  renderer: ColumnRenderer;
  maxLength?: number;
  seeMoreLink?: boolean;
  multiValueSeparator?: MultiValueSeparator;
}

export interface IColumnPropertyOption {
  key: string;
  text: string;
  alias?: string;
}

/** Legacy shape stored by pre-Phase-1 pages. */
export interface ILegacyColumnItem {
  uniqueId?: string;
  property?: string;
}

const VALID_VISIBILITIES: ColumnVisibility[] = ['always', 'defaultOn', 'defaultOff'];

const VALID_RENDERERS: ColumnRenderer[] = [
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

/**
 * Renderers exposed in the side-panel editor's dropdown.
 *
 * Phase 1 shipped with the Phase-1 subset (`'' / text / date / fileType /
 * fileSize / url / persona`). Phase 2 added the remaining types
 * (`richText / number / tags / boolean`) once their cell renderers landed in
 * `renderCell.tsx`. The constant name is kept for back-compat with the
 * Phase-1 import path; the order matches the dropdown's visible order.
 */
export const PHASE_1_RENDERERS: ColumnRenderer[] = [
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

const VALID_SEPARATORS: MultiValueSeparator[] = ['comma', 'newline', 'semicolon', 'pill'];

let _uniqueIdSeq = 0;

export function generateColumnUniqueId(): string {
  _uniqueIdSeq = (_uniqueIdSeq + 1) % 1e9;
  return 'col-' + Date.now().toString(36) + '-' + _uniqueIdSeq.toString(36);
}

function pickEnum<T extends string>(value: unknown, allowed: T[], fallback: T): T {
  return typeof value === 'string' && (allowed as string[]).indexOf(value) >= 0 ? (value as T) : fallback;
}

function pickPositiveNumber(value: unknown): number | undefined {
  return typeof value === 'number' && isFinite(value) && value > 0 ? value : undefined;
}

/**
 * Migration normalizer. Wraps a legacy `{ uniqueId, property }` item or a
 * partially-populated new-shape item with safe defaults. `renderer: ''` is
 * the sentinel that routes through the existing auto-detect — every migrated
 * page renders byte-for-byte identically until an admin opens the editor.
 */
export function normalizeColumnConfigItem(
  raw: Partial<IColumnConfigItem> & ILegacyColumnItem
): IColumnConfigItem {
  const property = String(raw.property || '').trim();
  const aliasRaw = typeof raw.alias === 'string' ? raw.alias.trim() : '';
  const uniqueId = String(raw.uniqueId || '').trim() || generateColumnUniqueId();

  return {
    uniqueId,
    property,
    alias: aliasRaw || property,
    width: pickPositiveNumber(raw.width),
    visibility: pickEnum<ColumnVisibility>(raw.visibility, VALID_VISIBILITIES, 'defaultOn'),
    renderer: pickEnum<ColumnRenderer>(raw.renderer, VALID_RENDERERS, ''),
    maxLength: pickPositiveNumber(raw.maxLength),
    seeMoreLink: typeof raw.seeMoreLink === 'boolean' ? raw.seeMoreLink : undefined,
    multiValueSeparator:
      typeof raw.multiValueSeparator === 'string' && VALID_SEPARATORS.indexOf(raw.multiValueSeparator) >= 0
        ? raw.multiValueSeparator
        : undefined,
  };
}

function getAliasFromOptionText(text: string, property: string): string {
  const cleanText = String(text || '').trim().replace(/\s+\(already in use\)$/, '');
  const suffix = ' (' + property + ')';
  if (cleanText.length > suffix.length && cleanText.slice(-suffix.length) === suffix) {
    return cleanText.slice(0, cleanText.length - suffix.length).trim();
  }
  return '';
}

export function applyColumnPropertySelection(
  current: IColumnConfigItem,
  option: IColumnPropertyOption
): IColumnConfigItem {
  const property = String(option.key || '').trim();
  const currentProperty = String(current.property || '').trim();
  const currentAlias = String(current.alias || '').trim();
  const optionAlias = typeof option.alias === 'string' ? option.alias.trim() : '';
  const textAlias = getAliasFromOptionText(option.text, property);
  const aliasIsDefault = !currentAlias || currentAlias === currentProperty;
  const nextAlias = optionAlias && optionAlias !== property
    ? optionAlias
    : (textAlias || optionAlias || property);

  return {
    ...current,
    property,
    alias: aliasIsDefault ? nextAlias : current.alias,
  };
}
