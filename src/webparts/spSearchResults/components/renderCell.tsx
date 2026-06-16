import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { UserPersona as _UserPersona } from 'spfx-toolkit/lib/components/UserPersona';
import { sanitizeHtml } from 'spfx-toolkit/lib/utilities/htmlUtils/sanitizeHtml';
import { formatFileSize, formatRelativeDate, formatDateTime } from './documentTitleUtils';
import type { IColumnConfigItem, MultiValueSeparator } from './ColumnConfigField/columnConfig';
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
  // Sanitize first, then truncate the sanitized HTML's plain text length —
  // safer than truncating raw HTML which could leave a tag half-open.
  const safe = sanitizeHtml(raw);
  const display = column.maxLength && safe.length > column.maxLength
    ? safe.slice(0, column.maxLength) + ELLIPSIS
    : safe;
  return <span className={styles.gridRichText} dangerouslySetInnerHTML={{ __html: display }} />;
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
  return (
    <span className={styles.gridDateCell} title={formatDateTime(raw)}>
      {formatRelativeDate(raw)}
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

function splitTagValue(value: unknown): string[] {
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
  // Heuristic: taxonomy GP0|#GUID;Label format and pipe-separated user claims
  // are accepted as-is for Phase 2 — admin can switch separator as needed.
  // Default: split on commas or semicolons.
  return cleaned.split(/\s*[,;]\s*/).map((part) => cleanSearchResultDisplayText(part.trim()).trim()).filter(Boolean);
}

const SEPARATOR_JOIN: Record<MultiValueSeparator, string> = {
  comma: ', ',
  semicolon: '; ',
  newline: '\n',
  pill: '',
  badge: '',
};

export function renderTags(value: unknown, column: IColumnConfigItem): React.ReactElement {
  const parts = splitTagValue(value);
  if (parts.length === 0) {
    return muted();
  }
  const sep = column.multiValueSeparator || 'comma';
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

export function renderPersona(value: unknown, _column: IColumnConfigItem): React.ReactElement {
  const persona = extractPersona(value);
  if (!persona) {
    return muted();
  }
  return (
    <UserPersona
      userIdentifier={persona.email || persona.displayName}
      displayName={persona.displayName}
      size={24}
      displayMode="avatarAndName"
    />
  );
}

// ─── url ──────────────────────────────────────────────────

export function renderUrl(value: unknown, column: IColumnConfigItem): React.ReactElement {
  const raw = typeof value === 'string' ? value.trim() : '';
  if (!raw) {
    return muted();
  }
  const visible = truncate(raw, column.maxLength);
  return (
    <a href={raw} target="_blank" rel="noopener noreferrer" className={styles.gridTitleLink} title={raw}>
      {visible}
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
