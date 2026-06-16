import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { UserPersona as _UserPersona } from 'spfx-toolkit/lib/components/UserPersona';
import { sanitizeHtml } from 'spfx-toolkit/lib/utilities/htmlUtils/sanitizeHtml';
import { formatFileSize, formatRelativeDate, formatDateTime } from './documentTitleUtils';
import { isSafeHttpUrl } from '@store/utils/safeNavigate';
import {
  IColumnConfigItem,
  MultiValueSeparator,
  BadgeColor,
  IBadgeColorRule,
  AUTO_COLOR_PALETTE,
} from './ColumnConfigField/columnConfig';
import styles from './SpSearchResults.module.scss';

// spfx-toolkit's UserPersona ships with stricter @types/react than this project
// — cast to any per the project's existing pattern (see ListLayout.tsx).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const UserPersona: any = _UserPersona;

/**
 * Stream B / Phase 2 — pure cell renderers.
 *
 * Each renderer takes a raw cell value and the column config and returns a
 * React element. Renderers that need component state (title, author with its
 * full hover card + actions, etc.) stay in `DataGridContent.tsx` as
 * `React.useCallback`s closing over local state.
 *
 * Spec: docs/superpowers/specs/2026-05-13-stream-b-column-config-design.md
 */

const ELLIPSIS = '…';

function muted(): React.ReactElement {
  return <span className={styles.gridCellMuted}>--</span>;
}

function truncate(text: string, maxLength: number | undefined): string {
  if (!maxLength || maxLength <= 0 || text.length <= maxLength) {
    return text;
  }
  return text.slice(0, maxLength) + ELLIPSIS;
}

// SharePoint surfaces Calculated-column values with a leading field-type
// token, e.g. "string;#Electronic Bank Statements" or "datetime;#2024-01-01".
// Calculated columns only emit one of four output types (string/float/datetime/
// boolean), so we strip ONLY those exact prefixes — a legitimate value that
// merely starts with another word followed by ";#" (e.g. "Approved;#Rejected")
// is left intact. `[\s\S]*` keeps multi-line values whole (`.` stops at \n).
const CALCULATED_TYPE_PREFIX = /^(?:string|float|datetime|boolean);#([\s\S]*)$/;

export function cleanSearchResultDisplayText(value: string): string {
  const raw = String(value || '');
  const match = CALCULATED_TYPE_PREFIX.exec(raw);
  return match ? match[1] : raw;
}

// ─── badge color resolution ──────────────────────────────

export interface IResolvedBadge {
  color: BadgeColor;
  icon?: string;
}

function hashBadgeColorIndex(s: string, len: number): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) {
    h = (h * 31 + s.charCodeAt(i)) | 0;
  }
  return Math.abs(h) % len;
}

/**
 * Resolve a badge token's color: an admin map entry (case-insensitive) wins;
 * otherwise auto-color from a stable hash; otherwise neutral.
 */
export function resolveBadgeColor(
  value: string,
  map: Map<string, IBadgeColorRule> | undefined,
  autoColorUnmapped: boolean
): IResolvedBadge {
  const key = value.trim().toLowerCase();
  const mapped = map ? map.get(key) : undefined;
  if (mapped) {
    return mapped.icon ? { color: mapped.color, icon: mapped.icon } : { color: mapped.color };
  }
  if (autoColorUnmapped) {
    return { color: AUTO_COLOR_PALETTE[hashBadgeColorIndex(key, AUTO_COLOR_PALETTE.length)] };
  }
  return { color: 'neutral' };
}

// Cache the built lookup map per column-config rules array. `renderTags` is the
// DataGrid `cellRender` (called per cell); the rules array is the same stable
// reference for every cell in a column, so this builds the map once per column
// instead of once per cell. Keyed by the array reference → auto-GC'd when the
// column config is replaced.
const BADGE_COLOR_MAP_CACHE = new WeakMap<IBadgeColorRule[], Map<string, IBadgeColorRule>>();

function buildBadgeColorMap(rules: IBadgeColorRule[] | undefined): Map<string, IBadgeColorRule> | undefined {
  if (!rules || rules.length === 0) {
    return undefined;
  }
  const cached = BADGE_COLOR_MAP_CACHE.get(rules);
  if (cached) {
    return cached;
  }
  const map = new Map<string, IBadgeColorRule>();
  for (const rule of rules) {
    // Key by the cleaned display form so an admin who pastes a raw SharePoint
    // value (e.g. "string;#Approved") still matches the cleaned cell token.
    map.set(cleanSearchResultDisplayText(rule.value).trim().toLowerCase(), rule);
  }
  BADGE_COLOR_MAP_CACHE.set(rules, map);
  return map;
}

function toStringValue(value: unknown): string {
  if (value === undefined || value === null) {
    return '';
  }
  if (typeof value === 'string') {
    return cleanSearchResultDisplayText(value);
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  if (Array.isArray(value)) {
    return value.map(toStringValue).filter(Boolean).join(', ');
  }
  if (typeof value === 'object') {
    const obj = value as Record<string, unknown>;
    if (typeof obj.displayText === 'string') {
      return cleanSearchResultDisplayText(obj.displayText);
    }
    return JSON.stringify(value);
  }
  return String(value);
}

// ─── text ─────────────────────────────────────────────────

export function renderText(value: unknown, column: IColumnConfigItem): React.ReactElement {
  const raw = toStringValue(value).trim();
  if (!raw) {
    return muted();
  }
  const truncated = truncate(raw, column.maxLength);
  return <span title={raw}>{truncated}</span>;
}

// ─── richText ─────────────────────────────────────────────

export function renderRichText(value: unknown, column: IColumnConfigItem): React.ReactElement {
  const raw = toStringValue(value).trim();
  if (!raw) {
    return muted();
  }
  const safe = sanitizeHtml(raw);
  // When a max length is set, truncate on the PLAIN text (tags stripped) so we
  // never cut a tag in half and inject broken markup. The truncated form renders
  // as text; the untruncated form keeps its formatting.
  if (column.maxLength && column.maxLength > 0) {
    const plain = safe.replace(/<[^>]*>/g, '').trim();
    if (plain.length > column.maxLength) {
      const clipped = plain.slice(0, column.maxLength) + ELLIPSIS;
      return <span className={styles.gridRichText} title={plain}>{clipped}</span>;
    }
  }
  return <span className={styles.gridRichText} dangerouslySetInnerHTML={{ __html: safe }} />;
}

// ─── number ───────────────────────────────────────────────

export function renderNumber(value: unknown, _column: IColumnConfigItem): React.ReactElement {
  let n: number | undefined;
  if (typeof value === 'number' && isFinite(value)) {
    n = value;
  } else if (typeof value === 'string' && value.trim()) {
    const parsed = Number(value);
    if (isFinite(parsed)) {
      n = parsed;
    }
  }
  if (n === undefined) {
    return muted();
  }
  // Intl.NumberFormat with default locale — picks up the browser/jsdom locale.
  return <span>{new Intl.NumberFormat().format(n)}</span>;
}

// ─── fileSize ─────────────────────────────────────────────

export function renderFileSize(value: unknown, _column: IColumnConfigItem): React.ReactElement {
  const bytes = typeof value === 'number'
    ? value
    : parseInt(String(value || '0'), 10) || 0;
  if (!bytes) {
    return muted();
  }
  return <span title={String(bytes) + ' bytes'}>{formatFileSize(bytes)}</span>;
}

// ─── date ─────────────────────────────────────────────────

export function renderDate(value: unknown, _column: IColumnConfigItem): React.ReactElement {
  const raw = typeof value === 'string' ? value : '';
  if (!raw) {
    return muted();
  }
  const relative = formatRelativeDate(raw);
  if (!relative) {
    // Unparseable date string — show the muted dash, not a blank cell.
    return muted();
  }
  return (
    <span className={styles.gridDateCell} title={formatDateTime(raw)}>
      {relative}
    </span>
  );
}

// ─── boolean ──────────────────────────────────────────────

function isTruthy(value: unknown): boolean | undefined {
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    return value !== 0;
  }
  if (typeof value === 'string') {
    const v = value.trim().toLowerCase();
    if (v === '') {
      return undefined;
    }
    if (v === 'true' || v === 'yes' || v === '1') {
      return true;
    }
    if (v === 'false' || v === 'no' || v === '0') {
      return false;
    }
    return undefined;
  }
  return undefined;
}

export function renderBoolean(value: unknown, _column: IColumnConfigItem): React.ReactElement {
  const truth = isTruthy(value);
  if (truth === undefined) {
    return muted();
  }
  return (
    <Icon
      iconName={truth ? 'CheckMark' : 'Cancel'}
      className={truth ? styles.gridBooleanTrue : styles.gridBooleanFalse}
      ariaLabel={truth ? 'Yes' : 'No'}
    />
  );
}

// ─── tags ─────────────────────────────────────────────────

const SPLIT_REGEX_CACHE = new Map<string, RegExp>();

function resolveSplitRegex(delimiter: string | undefined): RegExp {
  if (!delimiter) {
    return /\s*[,;]\s*/;
  }
  const cached = SPLIT_REGEX_CACHE.get(delimiter);
  if (cached) {
    return cached;
  }
  const normalized = delimiter.replace(/\\n/g, '\n').replace(/\\t/g, '\t');
  const escaped = normalized.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  /* eslint-disable @rushstack/security/no-unsafe-regexp -- input is fully
   * regex-escaped above; pattern is \s*<literal>\s* with no nested
   * quantifiers, so no injection or ReDoS is possible. */
  const re = new RegExp('\\s*' + escaped + '\\s*');
  /* eslint-enable @rushstack/security/no-unsafe-regexp */
  SPLIT_REGEX_CACHE.set(delimiter, re);
  return re;
}

function splitTagValue(value: unknown, delimiter?: string): string[] {
  if (Array.isArray(value)) {
    return value.map((v) => toStringValue(v).trim()).filter(Boolean);
  }
  if (typeof value !== 'string') {
    return [];
  }
  const trimmed = value.trim();
  if (!trimmed) {
    return [];
  }
  const cleaned = cleanSearchResultDisplayText(trimmed);
  return cleaned
    .split(resolveSplitRegex(delimiter))
    .map((part) => cleanSearchResultDisplayText(part.trim()).trim())
    .filter(Boolean);
}

const SEPARATOR_JOIN: Record<MultiValueSeparator, string> = {
  comma: ', ',
  semicolon: '; ',
  newline: '\n',
  pill: '',
  badge: '',
};

export function renderTags(value: unknown, column: IColumnConfigItem): React.ReactElement {
  const parts = splitTagValue(value, column.splitDelimiter);
  if (parts.length === 0) {
    return muted();
  }
  const sep = column.multiValueSeparator || 'comma';
  if (sep === 'badge') {
    const colorMap = buildBadgeColorMap(column.valueColorMap);
    const autoColor = column.autoColorUnmapped !== false;
    const badgeStyles = styles as unknown as Record<string, string>;
    return (
      <span className={styles.gridTagsCell}>
        {parts.map((part, idx) => {
          const resolved = resolveBadgeColor(part, colorMap, autoColor);
          const colorClass = badgeStyles['gridBadge--' + resolved.color] || '';
          return (
            <span
              key={part + '-' + String(idx)}
              className={styles.gridBadge + ' ' + colorClass}
              title={part}
            >
              {resolved.icon ? (
                <Icon iconName={resolved.icon} className={styles.gridBadgeIcon} aria-hidden={true} />
              ) : null}
              {part}
            </span>
          );
        })}
      </span>
    );
  }
  if (sep === 'pill') {
    return (
      <span className={styles.gridTagsCell}>
        {parts.map((part, idx) => (
          <span key={part + '-' + String(idx)} className={styles.gridTagPill}>{part}</span>
        ))}
      </span>
    );
  }
  if (sep === 'newline') {
    return (
      <span className={styles.gridTagsCell}>
        {parts.map((part, idx) => (
          <React.Fragment key={part + '-' + String(idx)}>
            {idx > 0 && <br />}
            {part}
          </React.Fragment>
        ))}
      </span>
    );
  }
  return <span title={parts.join(SEPARATOR_JOIN[sep])}>{parts.join(SEPARATOR_JOIN[sep])}</span>;
}

// ─── persona ──────────────────────────────────────────────

interface IPersonaLike {
  displayName: string;
  email?: string;
}

/** Coerce arbitrary persona-shaped data into a `{ displayName, email }` pair. */
function extractPersona(value: unknown): IPersonaLike | undefined {
  if (value === undefined || value === null || value === '') {
    return undefined;
  }
  if (typeof value === 'object') {
    const obj = value as Record<string, unknown>;
    const displayName = typeof obj.displayText === 'string' ? obj.displayText
      : typeof obj.displayName === 'string' ? obj.displayName
      : typeof obj.name === 'string' ? obj.name
      : undefined;
    const email = typeof obj.email === 'string' ? obj.email
      : typeof obj.mail === 'string' ? obj.mail
      : undefined;
    if (displayName) {
      return { displayName, email };
    }
    return undefined;
  }
  const raw = String(value).trim();
  if (!raw) {
    return undefined;
  }
  // Claim string: `i:0#.f|membership|user@tenant.onmicrosoft.com` — extract the
  // upn portion as the email; display name is the upn until we resolve it.
  const claimMatch = /^i:0#\.f\|membership\|(.+)$/.exec(raw);
  if (claimMatch) {
    const upn = claimMatch[1];
    return { displayName: upn, email: upn };
  }
  // Email-like
  if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(raw)) {
    return { displayName: raw, email: raw };
  }
  return { displayName: raw };
}

/**
 * Split a people value into individual personas. People managed properties can
 * carry several values joined by ';' (co-authors, service/app accounts). Split
 * on ';' only — never comma — since a display name may contain a comma.
 */
function extractPersonaList(value: unknown): IPersonaLike[] {
  if (Array.isArray(value)) {
    return value.map(extractPersona).filter((p): p is IPersonaLike => !!p);
  }
  if (typeof value === 'string' && value.indexOf(';') >= 0) {
    return value.split(/\s*;\s*/).map((part) => extractPersona(part)).filter((p): p is IPersonaLike => !!p);
  }
  const single = extractPersona(value);
  return single ? [single] : [];
}

function personaElement(persona: IPersonaLike, key?: string): React.ReactElement {
  return (
    <UserPersona
      key={key}
      userIdentifier={persona.email || persona.displayName}
      displayName={persona.displayName}
      size={24}
      displayMode="avatarAndName"
    />
  );
}

export function renderPersona(value: unknown, _column: IColumnConfigItem): React.ReactElement {
  const people = extractPersonaList(value);
  if (people.length === 0) {
    return muted();
  }
  if (people.length === 1) {
    return personaElement(people[0]);
  }
  // One persona per line for a multi-value people field.
  return (
    <span className={styles.gridPersonaList}>
      {people.map((persona, idx) => personaElement(persona, (persona.email || persona.displayName) + '-' + String(idx)))}
    </span>
  );
}

// ─── url ──────────────────────────────────────────────────

export function renderUrl(value: unknown, column: IColumnConfigItem): React.ReactElement {
  const raw = toStringValue(value).trim();
  if (!raw) {
    return muted();
  }

  let url = raw;
  let label = raw;
  // SharePoint Hyperlink fields surface as "https://url, Description". If the
  // text before ", " is itself a valid URL, treat the remainder as the label.
  // (isSafeHttpUrl only validates the scheme, so "https://x, Desc" would pass as
  // a whole — we must split first to keep the description out of the href.)
  const commaIdx = raw.indexOf(', ');
  if (commaIdx > 0 && isSafeHttpUrl(raw.slice(0, commaIdx).trim())) {
    url = raw.slice(0, commaIdx).trim();
    label = raw.slice(commaIdx + 2).trim() || url;
  }

  // Only render a clickable link for safe http(s)/root-relative URLs — a raw
  // managed-property value could be a javascript:/data: scheme (stored XSS).
  if (!isSafeHttpUrl(url)) {
    return <span title={raw}>{truncate(label, column.maxLength)}</span>;
  }
  return (
    <a href={url} target="_blank" rel="noopener noreferrer" className={styles.gridTitleLink} title={url}>
      {truncate(label, column.maxLength)}
    </a>
  );
}

// ─── fileType ─────────────────────────────────────────────

export function renderFileType(value: unknown, _column: IColumnConfigItem): React.ReactElement {
  const label = toStringValue(value).trim().toUpperCase();
  if (!label) {
    return muted();
  }
  return <span className={styles.gridTypeBadge}>{label}</span>;
}
