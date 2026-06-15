# DataGrid Token / Badge Cell Rendering Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a per-column configurable split delimiter and a colored "Badges" display style to the DataGrid `tags` renderer, with an admin value→color map and a deterministic auto-color fallback.

**Architecture:** Extend the existing per-column config (`IColumnConfigItem`) with three optional fields and one new `multiValueSeparator` value (`'badge'`). The pure cell renderer (`renderCell.tsx`) gains a configurable split and a color-resolution helper; the column-config side panel (`ColumnConfigPanel.tsx`) gains the editing UI. No new renderer type, no dispatch changes, fully back-compatible.

**Tech Stack:** SPFx 1.22 / React 17 / TypeScript 5.3, Fluent UI v8, Jest 29 (via Heft), SCSS modules.

**Spec:** [docs/superpowers/specs/2026-06-15-datagrid-token-badge-cells-design.md](../specs/2026-06-15-datagrid-token-badge-cells-design.md)

**Toolchain note:** `npm test` runs a full Heft build (TypeScript compile + lint) *before* Jest. So a "verify it fails" step on a test that imports a not-yet-existing symbol fails at the **compile** stage (`Module has no exported member …`) rather than as a Jest assertion. That compile error *is* the expected red. After implementing, the same command compiles and runs the tests green.

**Run a single test file:** `npm test -- --test-path-pattern "<pattern>"`

---

## File Structure

| File | Responsibility | Change |
|---|---|---|
| `src/webparts/spSearchResults/components/ColumnConfigField/columnConfig.ts` | Per-column config schema + normalizer | New types/palette/fields + validation |
| `src/webparts/spSearchResults/components/renderCell.tsx` | Pure cell renderers | Configurable split + color resolution + badge chips |
| `src/webparts/spSearchResults/components/SpSearchResults.module.scss` | Grid cell styles | `.gridBadge` base + 9 color modifiers + icon class |
| `src/webparts/spSearchResults/components/ColumnConfigField/ColumnConfigPanel.tsx` | Column-config side panel | Split field, "Badges" option, value-color editor, auto-color toggle |
| `src/webparts/spSearchResults/components/ColumnConfigField/ColumnConfigField.module.scss` | Panel styles | `.badgeRuleRow` flex row |
| `tests/webparts/spSearchResults/columnConfig.test.ts` | Normalizer tests | New-field round-trip tests |
| `tests/webparts/spSearchResults/renderCell.test.tsx` | Renderer tests | Split + color-resolution + badge tests |

---

## Task 1: Config model — types, palette, fields, normalizer

**Files:**
- Modify: `src/webparts/spSearchResults/components/ColumnConfigField/columnConfig.ts`
- Modify: `src/webparts/spSearchResults/components/renderCell.tsx` (one line — keep the `SEPARATOR_JOIN` record exhaustive after the union grows)
- Test: `tests/webparts/spSearchResults/columnConfig.test.ts`

- [ ] **Step 1: Write the failing tests**

First, extend the **existing** import at the top of `tests/webparts/spSearchResults/columnConfig.test.ts` to add `BADGE_COLORS` and `AUTO_COLOR_PALETTE` (do NOT add a second import statement). The existing import becomes:

```ts
import {
  normalizeColumnConfigItem,
  applyColumnPropertySelection,
  BADGE_COLORS,
  AUTO_COLOR_PALETTE,
  IColumnConfigItem,
  ColumnRenderer,
  ColumnVisibility,
} from '../../../src/webparts/spSearchResults/components/ColumnConfigField/columnConfig';
```

Then append this new describe block at the end of the file:

```ts
describe('normalizeColumnConfigItem — badge/split fields', () => {
  it('keeps a valid split delimiter and drops empty / over-long ones', () => {
    expect(normalizeColumnConfigItem({ uniqueId: 'a', property: 'P', splitDelimiter: '|' }).splitDelimiter).toBe('|');
    expect(normalizeColumnConfigItem({ uniqueId: 'a', property: 'P', splitDelimiter: '' }).splitDelimiter).toBeUndefined();
    expect(normalizeColumnConfigItem({ uniqueId: 'a', property: 'P', splitDelimiter: '123456789' }).splitDelimiter).toBeUndefined();
  });

  it('accepts the badge multi-value separator', () => {
    expect(normalizeColumnConfigItem({ uniqueId: 'a', property: 'P', multiValueSeparator: 'badge' }).multiValueSeparator).toBe('badge');
  });

  it('normalizes the value color map — trims, drops invalid colors, dedupes by lowercased value', () => {
    const out = normalizeColumnConfigItem({
      uniqueId: 'a',
      property: 'P',
      valueColorMap: [
        { value: '  Approved ', color: 'green' },
        { value: 'approved', color: 'red' },        // duplicate (case-insensitive) — dropped
        { value: 'Pending', color: 'not-a-color' as never },  // invalid color — dropped
        { value: '', color: 'blue' },               // empty value — dropped
        { value: 'Overdue', color: 'red', icon: 'Warning' },
      ],
    });
    expect(out.valueColorMap).toEqual([
      { value: 'Approved', color: 'green' },
      { value: 'Overdue', color: 'red', icon: 'Warning' },
    ]);
  });

  it('returns undefined for an absent or fully-invalid color map', () => {
    expect(normalizeColumnConfigItem({ uniqueId: 'a', property: 'P' }).valueColorMap).toBeUndefined();
    expect(normalizeColumnConfigItem({ uniqueId: 'a', property: 'P', valueColorMap: [] }).valueColorMap).toBeUndefined();
  });

  it('preserves autoColorUnmapped only when boolean', () => {
    expect(normalizeColumnConfigItem({ uniqueId: 'a', property: 'P', autoColorUnmapped: false }).autoColorUnmapped).toBe(false);
    expect(normalizeColumnConfigItem({ uniqueId: 'a', property: 'P' }).autoColorUnmapped).toBeUndefined();
  });

  it('exposes the palette constants', () => {
    expect(BADGE_COLORS).toContain('neutral');
    expect(BADGE_COLORS).toContain('magenta');
    expect(AUTO_COLOR_PALETTE).not.toContain('neutral');
    expect(AUTO_COLOR_PALETTE.length).toBe(8);
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npm test -- --test-path-pattern "columnConfig"`
Expected: FAIL — TypeScript compile error `Module … has no exported member 'BADGE_COLORS'` (and `AUTO_COLOR_PALETTE`).

- [ ] **Step 3: Implement the config model**

In `columnConfig.ts`, extend the `MultiValueSeparator` union and the `IColumnConfigItem` interface, add the badge types/palette, extend `VALID_SEPARATORS`, and add the normalizer logic.

Change the separator union:

```ts
export type MultiValueSeparator = 'comma' | 'newline' | 'semicolon' | 'pill' | 'badge';
```

Add, immediately after the `IColumnPropertyOption` interface:

```ts
export type BadgeColor =
  | 'neutral' | 'blue' | 'teal' | 'green'
  | 'amber'   | 'orange' | 'red' | 'purple' | 'magenta';

/** All selectable badge colors, in dropdown order. */
export const BADGE_COLORS: BadgeColor[] = [
  'neutral', 'blue', 'teal', 'green', 'amber', 'orange', 'red', 'purple', 'magenta',
];

/** Colors used to auto-color unmapped values (excludes neutral). */
export const AUTO_COLOR_PALETTE: BadgeColor[] = [
  'blue', 'teal', 'green', 'amber', 'orange', 'red', 'purple', 'magenta',
];

const BADGE_COLOR_SET = new Set<string>(BADGE_COLORS);

/** One admin-defined value→color rule for the badge display style. */
export interface IBadgeColorRule {
  value: string;
  color: BadgeColor;
  icon?: string;
}
```

Add the three fields to `IColumnConfigItem` (after `multiValueSeparator?`):

```ts
  /** Character(s) to split the value on for tags/badges. Unset → default , and ;. */
  splitDelimiter?: string;
  /** Admin value→color rules for the 'badge' display style. */
  valueColorMap?: IBadgeColorRule[];
  /** Auto-color unmapped badge values (default true). */
  autoColorUnmapped?: boolean;
```

Update the separator constant:

```ts
const VALID_SEPARATORS: MultiValueSeparator[] = ['comma', 'newline', 'semicolon', 'pill', 'badge'];
```

Add this helper above `normalizeColumnConfigItem`:

```ts
function normalizeBadgeColorMap(raw: unknown): IBadgeColorRule[] | undefined {
  if (!Array.isArray(raw)) {
    return undefined;
  }
  const result: IBadgeColorRule[] = [];
  const seen = new Set<string>();
  for (const entry of raw) {
    if (!entry || typeof entry !== 'object') {
      continue;
    }
    const e = entry as Record<string, unknown>;
    const value = typeof e.value === 'string' ? e.value.trim() : '';
    const color = typeof e.color === 'string' && BADGE_COLOR_SET.has(e.color) ? (e.color as BadgeColor) : undefined;
    if (!value || !color) {
      continue;
    }
    const key = value.toLowerCase();
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    const icon = typeof e.icon === 'string' && e.icon.trim() ? e.icon.trim() : undefined;
    result.push(icon ? { value, color, icon } : { value, color });
    if (result.length >= 50) {
      break;
    }
  }
  return result.length > 0 ? result : undefined;
}
```

Add these three properties to the object returned by `normalizeColumnConfigItem` (after the existing `multiValueSeparator:` property):

```ts
    splitDelimiter:
      typeof raw.splitDelimiter === 'string' && raw.splitDelimiter.length > 0 && raw.splitDelimiter.length <= 8
        ? raw.splitDelimiter
        : undefined,
    valueColorMap: normalizeBadgeColorMap(raw.valueColorMap),
    autoColorUnmapped: typeof raw.autoColorUnmapped === 'boolean' ? raw.autoColorUnmapped : undefined,
```

- [ ] **Step 3b: Keep `SEPARATOR_JOIN` exhaustive**

Growing the `MultiValueSeparator` union makes the `Record<MultiValueSeparator, string>` literal `SEPARATOR_JOIN` in `renderCell.tsx` incomplete — a compile error the whole tree (and this task's own test run) would hit. Add the `badge` key now (its value is unused because Task 4 renders badges before reaching the join, but the key must exist for the type to be exhaustive):

```ts
const SEPARATOR_JOIN: Record<MultiValueSeparator, string> = {
  comma: ', ',
  semicolon: '; ',
  newline: '\n',
  pill: '',
  badge: '',
};
```

> If a `grep -n "Record<MultiValueSeparator" src/` turns up any other exhaustive map, add a `badge` entry there too. As of this plan, `SEPARATOR_JOIN` is the only one.

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npm test -- --test-path-pattern "columnConfig"`
Expected: PASS — all `columnConfig` tests green (existing 38 + the new block), and the full-tree compile succeeds (no missing-key error on `SEPARATOR_JOIN`).

- [ ] **Step 5: Commit**

```bash
git add src/webparts/spSearchResults/components/ColumnConfigField/columnConfig.ts src/webparts/spSearchResults/components/renderCell.tsx tests/webparts/spSearchResults/columnConfig.test.ts
git commit -m "feat(results): add badge/split fields to column config schema"
```

---

## Task 2: Configurable split delimiter

**Files:**
- Modify: `src/webparts/spSearchResults/components/renderCell.tsx:188-204` (`splitTagValue`) and `renderTags`
- Test: `tests/webparts/spSearchResults/renderCell.test.tsx`

- [ ] **Step 1: Write the failing tests**

Add inside the existing `describe('renderCell — Stream B / Phase 2', ...)` block in `renderCell.test.tsx`, within the `describe('renderTags', ...)` group:

```ts
    it('splits on a custom delimiter when splitDelimiter is set', () => {
      const out = html(renderTags('HR|Finance|Legal', col({ renderer: 'tags', multiValueSeparator: 'newline', splitDelimiter: '|' })));
      expect(out).toMatch(/HR[\s\S]*Finance[\s\S]*Legal/);
    });

    it('treats the \\n token in splitDelimiter as a newline split', () => {
      const out = html(renderTags('alpha\nbeta', col({ renderer: 'tags', multiValueSeparator: 'comma', splitDelimiter: '\\n' })));
      expect(out).toContain('alpha, beta');
    });

    it('still splits on the default , and ; when splitDelimiter is unset', () => {
      const out = html(renderTags('a, b; c', col({ renderer: 'tags', multiValueSeparator: 'comma' })));
      expect(out).toContain('a, b, c');
    });
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npm test -- --test-path-pattern "renderCell"`
Expected: FAIL — the first two assertions fail (custom delimiter ignored; value not split), because `splitTagValue` ignores `column.splitDelimiter`.

- [ ] **Step 3: Implement the configurable split**

In `renderCell.tsx`, add this helper directly above `splitTagValue`:

```ts
function resolveSplitRegex(delimiter: string | undefined): RegExp {
  if (!delimiter) {
    return /\s*[,;]\s*/;
  }
  const normalized = delimiter.replace(/\\n/g, '\n').replace(/\\t/g, '\t');
  const escaped = normalized.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp('\\s*' + escaped + '\\s*');
}
```

Change the `splitTagValue` signature and its split line to accept the delimiter:

```ts
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
```

In `renderTags`, change the first line from `const parts = splitTagValue(value);` to:

```ts
  const parts = splitTagValue(value, column.splitDelimiter);
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npm test -- --test-path-pattern "renderCell"`
Expected: PASS — all `renderCell` tests green (existing 33 + the 3 new split tests).

- [ ] **Step 5: Commit**

```bash
git add src/webparts/spSearchResults/components/renderCell.tsx tests/webparts/spSearchResults/renderCell.test.tsx
git commit -m "feat(results): configurable split delimiter for tags renderer"
```

---

## Task 3: Badge color resolution + hash

**Files:**
- Modify: `src/webparts/spSearchResults/components/renderCell.tsx`
- Test: `tests/webparts/spSearchResults/renderCell.test.tsx`

- [ ] **Step 1: Write the failing tests**

Two import edits at the top of `renderCell.test.tsx`:
1. Add `resolveBadgeColor` to the existing `from '…/renderCell'` import (which already lists `cleanSearchResultDisplayText`).
2. Add `IBadgeColorRule` to the existing `import type { IColumnConfigItem } from '…/ColumnConfigField/columnConfig';` line → `import type { IColumnConfigItem, IBadgeColorRule } from '…/ColumnConfigField/columnConfig';`.

Then add a new `describe` block (sibling to the others, inside the top-level describe):

```ts
  describe('resolveBadgeColor', () => {
    const map = new Map<string, IBadgeColorRule>([
      ['approved', { value: 'Approved', color: 'green' }],
      ['overdue', { value: 'Overdue', color: 'red', icon: 'Warning' }],
    ]);

    it('returns the mapped color (case-insensitive) and icon', () => {
      expect(resolveBadgeColor('APPROVED', map, true)).toEqual({ color: 'green' });
      expect(resolveBadgeColor('overdue', map, true)).toEqual({ color: 'red', icon: 'Warning' });
    });

    it('auto-colors an unmapped value with a stable, non-neutral color', () => {
      const a = resolveBadgeColor('Engineering', map, true);
      const b = resolveBadgeColor('Engineering', map, true);
      expect(a).toEqual(b);                 // stable
      expect(a.color).not.toBe('neutral');  // from the auto palette
    });

    it('falls back to neutral when auto-color is off and value is unmapped', () => {
      expect(resolveBadgeColor('Engineering', map, false)).toEqual({ color: 'neutral' });
    });

    it('works with no map at all', () => {
      expect(resolveBadgeColor('Anything', undefined, false)).toEqual({ color: 'neutral' });
      expect(resolveBadgeColor('Anything', undefined, true).color).not.toBe('neutral');
    });
  });
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npm test -- --test-path-pattern "renderCell"`
Expected: FAIL — compile error `Module … has no exported member 'resolveBadgeColor'`.

- [ ] **Step 3: Implement the resolver**

In `renderCell.tsx`, the existing import is a type-only import:

```ts
import type { IColumnConfigItem, MultiValueSeparator } from './ColumnConfigField/columnConfig';
```

It now needs a runtime value (`AUTO_COLOR_PALETTE`), so replace it with a single regular import that carries both the types and the value (this mirrors `ColumnConfigPanel.tsx`, which already imports mixed types+values from this module in one plain `import`, confirming the repo does not enforce `import type` separation):

```ts
import {
  IColumnConfigItem,
  MultiValueSeparator,
  BadgeColor,
  IBadgeColorRule,
  AUTO_COLOR_PALETTE,
} from './ColumnConfigField/columnConfig';
```

Add this block directly below the `cleanSearchResultDisplayText` function (`buildBadgeColorMap` is intentionally deferred to Task 4, where it is first used, so nothing here is unused):

```ts
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
```

(`IBadgeColorRule` is used in the `map` parameter type; `BadgeColor` in `IResolvedBadge`; `AUTO_COLOR_PALETTE` in the auto branch — so every imported symbol is referenced and lint stays clean.)

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npm test -- --test-path-pattern "renderCell"`
Expected: PASS — the new `resolveBadgeColor` block green; existing tests still green.

- [ ] **Step 5: Commit**

```bash
git add src/webparts/spSearchResults/components/renderCell.tsx tests/webparts/spSearchResults/renderCell.test.tsx
git commit -m "feat(results): badge color resolution with auto-color fallback"
```

---

## Task 4: Badge rendering in renderTags + styles

**Files:**
- Modify: `src/webparts/spSearchResults/components/renderCell.tsx` (`buildBadgeColorMap` + `renderTags` badge branch)
- Modify: `src/webparts/spSearchResults/components/SpSearchResults.module.scss` (after `.gridTagPill`, ~line 2496)
- Test: `tests/webparts/spSearchResults/renderCell.test.tsx`

- [ ] **Step 1: Write the failing tests**

Add inside the `describe('renderTags', ...)` group in `renderCell.test.tsx`:

```ts
    it('renders mapped badge tokens with the color class, icon, and title', () => {
      const out = html(renderTags('Approved', col({
        renderer: 'tags',
        multiValueSeparator: 'badge',
        valueColorMap: [{ value: 'Approved', color: 'green', icon: 'CheckMark' }],
      })));
      expect(out).toContain('Approved');
      expect(out).toContain('gridBadge--green');
      expect(out).toContain('title="Approved"');
    });

    it('auto-colors unmapped badge values with a non-neutral class by default', () => {
      const out = html(renderTags('Engineering;Finance', col({ renderer: 'tags', multiValueSeparator: 'badge' })));
      expect(out).toContain('Engineering');
      expect(out).toContain('Finance');
      expect(out).toContain('gridBadge--');
      expect(out).not.toContain('gridBadge--neutral');
    });

    it('uses neutral badges when auto-color is disabled and value is unmapped', () => {
      const out = html(renderTags('Engineering', col({ renderer: 'tags', multiValueSeparator: 'badge', autoColorUnmapped: false })));
      expect(out).toContain('gridBadge--neutral');
    });

    it('emits the muted dash for an empty badge value', () => {
      expect(html(renderTags('', col({ renderer: 'tags', multiValueSeparator: 'badge' })))).toContain('--');
    });
```

> Note: the tests assert on the *class name string* `gridBadge--green`. In Jest the SCSS module is mapped to an identity proxy (class names pass through as their literal keys), so `styles['gridBadge--green']` resolves to the string `'gridBadge--green'`. This holds in this repo's existing renderer tests (e.g. they assert `gridCellMuted`-style classes via substrings).

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npm test -- --test-path-pattern "renderCell"`
Expected: FAIL — `gridBadge--green` not present (the `'badge'` separator currently falls through to the inline-join branch).

- [ ] **Step 3a: Add the badge styles**

In `SpSearchResults.module.scss`, add directly after the `.gridTagPill { … }` block (around line 2496):

```scss
.gridBadge {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 1px 8px;
  border-radius: 10px;
  font-size: 11px;
  line-height: 16px;
  font-weight: 600;
  white-space: nowrap;
}

.gridBadgeIcon {
  font-size: 11px;
}

// Fixed, AA-contrast color pairs (soft fill + dark text). Intentionally not
// theme accents so the nine options stay distinct and legible in any theme.
.gridBadge--neutral { background: #f3f2f1; color: #323130; }
.gridBadge--blue    { background: #deecf9; color: #004578; }
.gridBadge--teal    { background: #cdeef2; color: #005b70; }
.gridBadge--green   { background: #dff6dd; color: #0b6a0b; }
.gridBadge--amber   { background: #fff4ce; color: #735c00; }
.gridBadge--orange  { background: #fce4d6; color: #8a3707; }
.gridBadge--red     { background: #fde7e9; color: #a4262c; }
.gridBadge--purple  { background: #eddbf5; color: #5c2e91; }
.gridBadge--magenta { background: #fce4f3; color: #9b0062; }
```

- [ ] **Step 3b: Implement the badge branch in `renderTags`**

In `renderCell.tsx`, add the `buildBadgeColorMap` helper directly **after** the `resolveBadgeColor` function added in Task 3:

```ts
function buildBadgeColorMap(rules: IBadgeColorRule[] | undefined): Map<string, IBadgeColorRule> | undefined {
  if (!rules || rules.length === 0) {
    return undefined;
  }
  const map = new Map<string, IBadgeColorRule>();
  for (const rule of rules) {
    map.set(rule.value.trim().toLowerCase(), rule);
  }
  return map;
}
```

(`SEPARATOR_JOIN` already has its `badge: ''` key from Task 1, so no change there.)

Then, in `renderTags`, add the badge branch immediately **before** the existing `if (sep === 'pill')` block:

```ts
  if (sep === 'badge') {
    const colorMap = buildBadgeColorMap(column.valueColorMap);
    const autoColor = column.autoColorUnmapped !== false;
    const badgeStyles = styles as Record<string, string>;
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
```

(`Icon` is already imported at the top of `renderCell.tsx`.)

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npm test -- --test-path-pattern "renderCell"`
Expected: PASS — the four new badge tests green; existing tests still green.

- [ ] **Step 5: Commit**

```bash
git add src/webparts/spSearchResults/components/renderCell.tsx src/webparts/spSearchResults/components/SpSearchResults.module.scss tests/webparts/spSearchResults/renderCell.test.tsx
git commit -m "feat(results): colored badge display style for tags renderer"
```

---

## Task 5: Column-config side-panel UI

**Files:**
- Modify: `src/webparts/spSearchResults/components/ColumnConfigField/ColumnConfigPanel.tsx`
- Modify: `src/webparts/spSearchResults/components/ColumnConfigField/ColumnConfigField.module.scss`

No unit test — the panel is a Fluent property-pane component with no existing unit tests in this repo; it is verified by build + lint + the manual smoke check in Task 6. (All testable logic — split, color resolution, normalization — is already covered by Tasks 1–4.)

- [ ] **Step 1: Add the panel row style**

In `ColumnConfigField.module.scss`, add after the `.panelSection { … }` block:

```scss
.badgeRuleRow {
  display: flex;
  align-items: flex-end;
  gap: 8px;
  margin-bottom: 6px;
}
```

- [ ] **Step 2: Extend the panel imports + add the color-options constant**

In `ColumnConfigPanel.tsx`, change the `columnConfig` import to add the badge symbols:

```ts
import {
  IColumnConfigItem,
  IColumnPropertyOption,
  ColumnRenderer,
  ColumnVisibility,
  MultiValueSeparator,
  BadgeColor,
  BADGE_COLORS,
  PHASE_1_RENDERERS,
  applyColumnPropertySelection,
} from './columnConfig';
```

Add the new separator option to `SEPARATOR_OPTIONS`:

```ts
const SEPARATOR_OPTIONS: IDropdownOption[] = [
  { key: 'comma', text: 'Comma-separated' },
  { key: 'newline', text: 'One per line' },
  { key: 'semicolon', text: 'Semicolon-separated' },
  { key: 'pill', text: 'Pills' },
  { key: 'badge', text: 'Badges (colored)' },
];
```

Add this module-level constant directly below `SEPARATOR_OPTIONS`:

```ts
const BADGE_COLOR_OPTIONS: IDropdownOption[] = BADGE_COLORS.map((c) => ({
  key: c,
  text: c.charAt(0).toUpperCase() + c.slice(1),
}));
```

- [ ] **Step 3: Replace the separator block with the split + badge editor**

In `ColumnConfigPanel.tsx`, replace the entire existing block:

```tsx
        {rendererSupportsSeparator(draft.renderer) && (
          <Dropdown
            label="Multi-value separator"
            selectedKey={draft.multiValueSeparator || 'comma'}
            options={SEPARATOR_OPTIONS}
            onChange={(_e, option): void => {
              if (option) {
                update({ multiValueSeparator: option.key as MultiValueSeparator });
              }
            }}
          />
        )}
```

with:

```tsx
        {rendererSupportsSeparator(draft.renderer) && (
          <>
            <TextField
              label="Split character"
              value={draft.splitDelimiter || ''}
              placeholder=";  ,  |"
              description={'Character to split the value on. Use \\n for a new line. Leave blank for the default ( , and ; ).'}
              onChange={(_e, newValue): void => update({ splitDelimiter: newValue || undefined })}
            />
            <Dropdown
              label="Display style"
              selectedKey={draft.multiValueSeparator || 'comma'}
              options={SEPARATOR_OPTIONS}
              onChange={(_e, option): void => {
                if (option) {
                  update({ multiValueSeparator: option.key as MultiValueSeparator });
                }
              }}
            />
            {draft.multiValueSeparator === 'badge' && (
              <div className={styles.panelSection}>
                <Toggle
                  label="Auto-color other values"
                  checked={draft.autoColorUnmapped !== false}
                  onChange={(_e, checked): void => update({ autoColorUnmapped: !!checked })}
                />
                {(draft.valueColorMap || []).map((rule, idx) => (
                  <div key={'rule-' + String(idx)} className={styles.badgeRuleRow}>
                    <TextField
                      label={idx === 0 ? 'Value' : undefined}
                      value={rule.value}
                      placeholder="Approved"
                      onChange={(_e, v): void => {
                        const next = (draft.valueColorMap || []).map((r, i) => (i === idx ? { ...r, value: v || '' } : r));
                        update({ valueColorMap: next });
                      }}
                    />
                    <Dropdown
                      label={idx === 0 ? 'Color' : undefined}
                      selectedKey={rule.color}
                      options={BADGE_COLOR_OPTIONS}
                      onChange={(_e, option): void => {
                        if (option) {
                          const next = (draft.valueColorMap || []).map((r, i) => (i === idx ? { ...r, color: option.key as BadgeColor } : r));
                          update({ valueColorMap: next });
                        }
                      }}
                    />
                    <TextField
                      label={idx === 0 ? 'Icon (optional)' : undefined}
                      value={rule.icon || ''}
                      placeholder="CheckMark"
                      onChange={(_e, v): void => {
                        const next = (draft.valueColorMap || []).map((r, i) => (i === idx ? { ...r, icon: v || undefined } : r));
                        update({ valueColorMap: next });
                      }}
                    />
                    <DefaultButton
                      text="Remove"
                      onClick={(): void => {
                        update({ valueColorMap: (draft.valueColorMap || []).filter((_r, i) => i !== idx) });
                      }}
                    />
                  </div>
                ))}
                <DefaultButton
                  text="Add value color"
                  iconProps={{ iconName: 'Add' }}
                  onClick={(): void => {
                    update({ valueColorMap: [...(draft.valueColorMap || []), { value: '', color: 'blue' }] });
                  }}
                />
              </div>
            )}
          </>
        )}
```

(`TextField`, `Dropdown`, `Toggle`, `DefaultButton`, and `IDropdownOption` are already imported in this file.)

- [ ] **Step 4: Build + lint to verify it compiles cleanly**

Run: `npm test -- --test-path-pattern "columnConfig"`
(The `npm test` run performs the full TypeScript compile + ESLint over the whole `src/` tree before Jest, so this also validates `ColumnConfigPanel.tsx` and `renderCell.tsx` compile and lint clean.)
Expected: build `---- build finished ----` with no TypeScript/ESLint errors, then Jest PASS.

- [ ] **Step 5: Commit**

```bash
git add src/webparts/spSearchResults/components/ColumnConfigField/ColumnConfigPanel.tsx src/webparts/spSearchResults/components/ColumnConfigField/ColumnConfigField.module.scss
git commit -m "feat(results): badge/split editor in column config panel"
```

---

## Task 6: Full-suite verification + manual smoke

**Files:** none (verification only)

- [ ] **Step 1: Run the full results-renderer + config test suites**

Run: `npm test -- --test-path-pattern "renderCell|columnConfig"`
Expected: PASS — 0 failures across both files (existing + all new tests).

- [ ] **Step 2: Production build check**

Run: `npm run package`
Expected: build + package succeed; no bundle-budget failure for `sp-search-results-web-part.js` or `sp-search-experience-web-part.js` (the badge code is small; budgets are 1.1 MB / 1.6 MB).

- [ ] **Step 3: Manual smoke (local workbench — `npm start`)**

Verify in a DataGrid column configured as "Tags / badges":
- A multi-value field with `splitDelimiter` set to `|` shows separate badges.
- `Display style = One per line` stacks the values; `= Badges (colored)` shows chips.
- A value mapped in the editor (e.g. `Approved → green`, icon `CheckMark`) renders with that color + icon.
- An unmapped value gets a stable auto color; toggling **Auto-color other values** off makes unmapped values neutral grey.
- A single-value status column set to Badges shows one colored chip.
- Switching the site to a dark theme keeps badge text legible.

- [ ] **Step 4: Final commit (only if Step 3 surfaced fixes)**

```bash
git add -A
git commit -m "fix(results): badge rendering smoke-test adjustments"
```

---

## Notes for the implementer

- **DRY:** reuse `cleanSearchResultDisplayText` (already applied inside `splitTagValue`) — do not re-strip prefixes in the badge branch.
- **YAGNI:** no new renderer type, no dispatch changes, no loc files, no theme-token color math — fixed accessible pairs per spec.
- **Back-compat:** every new field is optional and `multiValueSeparator` only adds a value; existing `tags`/`pill` columns are unaffected. Saved configs without these fields normalize to `undefined`/default.
- The Experience web part (`SpSearchExperienceWebPart`) renders the same Results component and column config, so it inherits this behavior with no change.
