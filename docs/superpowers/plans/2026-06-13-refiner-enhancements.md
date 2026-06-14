# Refiner Enhancements Implementation Plan

> **Historical implementation plan:** this file records the task plan used
> during the 2026-06-13 refiner work. The final shipped taxonomy implementation
> pivoted from PnP `TaxonomyPicker` to DevExtreme `TagBox` with term-label
> resolution so refiner counts and cascade narrowing stay intact. Use the design
> spec and admin guide for current behavior.

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix three refiner bugs (people picker crash, multi-value selection clobber, taxonomy selection-doesn't-persist) and add five enhancements (type-aware values, delimited value splits, Fluent taxonomy UI, configurable toggle defaults, native DevExtreme TagBox).

**Architecture:** Two-PR split following [docs/superpowers/specs/2026-06-13-refiner-enhancements-design.md](../specs/2026-06-13-refiner-enhancements-design.md). PR1 (tasks 1-7) fixes the bugs + small enhancements. PR2 (tasks 8-11) adds data-type awareness and replaces the taxonomy UI. The pivot is a new batched callback `onReplaceRefinerValues({ filterName, values })` that multi-value filters use to emit their full intended selection in one call — eliminates the stale-closure clobber that today drops all but the last delta in a multi-value change.

**Tech Stack:** SPFx 1.22.2, React 17, TypeScript 5.3, Zustand 4.x, Fluent UI v8, DevExtreme 22.2.x, `@pnp/spfx-controls-react` 3.x, Heft 0.x test pipeline (Jest), PnPjs 3.x.

---

## File structure overview

**New files:**
- `src/propertyPaneControls/pnpStyleShims/TaxonomyPicker.module.scss.js` — Conditional, only if PR2 pre-flight reports the PnP SCSS crash.

**Modified across both PRs (responsibility per file):**

| File | Responsibility |
|---|---|
| `src/libraries/spSearchStore/interfaces/IFilterTypes.ts` | Add `dataType`, `valueSplitDelimiter`, `defaultValue` to `IFilterConfig` |
| `src/libraries/spSearchStore/providers/SharePointSearchProvider.ts` | `_mapRefiners` gets `filterConfig` arg + preprocessing pass |
| `src/libraries/spSearchStore/store/storeRegistry.ts` | Toggle `defaultValue` seed during `initializeSearchContext` |
| `src/webparts/spSearchFilters/components/SpSearchFilters.tsx` | Add `handleReplaceRefinerValues` + helper `applyReplaceRefinerValues` |
| `src/webparts/spSearchFilters/components/TagBoxFilter.tsx` | Use batched callback |
| `src/webparts/spSearchFilters/components/TaxonomyTreeFilter.tsx` | Use batched callback (PR1) → replaced by Fluent wrapper (PR2) |
| `src/webparts/spSearchFilters/components/DropdownFilter.tsx` | Use batched callback (multi mode only) |
| `src/webparts/spSearchFilters/components/PeoplePickerFilter.tsx` | Replace `@pnp/spfx-controls-react` PeoplePicker with Fluent `NormalPeoplePicker` + batched callback |
| `src/webparts/spSearchFilters/components/ToggleFilter.tsx` | Read `defaultValue` from config |
| `src/webparts/spSearchFilters/components/SpSearchFilters.module.scss` | Remove `.dx-tag` overrides |
| `src/propertyPaneControls/filtersCollection/FiltersCollectionControl.tsx` | New "Data format" section + Toggle default control |
| `src/propertyPaneControls/filtersCollection/fieldRelevance.ts` | Relevance map updates for new fields |
| `gulpfile.js` | Conditional webpack alias (only if PR2 pre-flight fails) |

**Test files (new):**

| Test file | Covers |
|---|---|
| `tests/webparts/spSearchFilters/applyReplaceRefinerValues.test.ts` | Batched callback helper |
| `tests/webparts/spSearchFilters/multiValueCascade.test.ts` | Multi-value filters keep all selections after a batch change |
| `tests/providers/mapRefinersPreprocessing.test.ts` | `_mapRefiners` strip + split |
| `tests/store/toggleDefaultValueSeed.test.ts` | Toggle defaults seeded by `initializeSearchContext` |

---

# PR1 — Bugs + small enhancements

### Task 1: Batched callback foundation (helper + parent)

**Files:**
- Create test: `tests/webparts/spSearchFilters/applyReplaceRefinerValues.test.ts`
- Modify: `src/webparts/spSearchFilters/components/SpSearchFilters.tsx` (add helper at module scope around line 240; add `handleReplaceRefinerValues` callback near line 454; add to the props passed down to filter components in the render loop near line 387)

**Why this is one task:** The helper is pure and trivial; the parent wires it. We don't migrate any component yet, so the existing per-delta callback still works — system is in a green state at end of task.

- [ ] **Step 1: Write the failing test**

Create `tests/webparts/spSearchFilters/applyReplaceRefinerValues.test.ts`:

```ts
import { applyReplaceRefinerValues } from '@webparts/spSearchFilters/components/SpSearchFilters';
import type { IActiveFilter } from '@interfaces/index';

describe('applyReplaceRefinerValues', () => {
  const base: IActiveFilter[] = [
    { filterName: 'FileType', value: 'pdf' },
    { filterName: 'Author', value: 'jdoe' },
  ];

  it('replaces all values for the named filter and preserves the rest', () => {
    const next = applyReplaceRefinerValues(base, 'FileType', [
      { filterName: 'FileType', value: 'docx' },
      { filterName: 'FileType', value: 'xlsx' },
    ]);
    expect(next).toEqual([
      { filterName: 'Author', value: 'jdoe' },
      { filterName: 'FileType', value: 'docx' },
      { filterName: 'FileType', value: 'xlsx' },
    ]);
  });

  it('returns a new array reference even when contents are equivalent', () => {
    const next = applyReplaceRefinerValues(base, 'FileType', [
      { filterName: 'FileType', value: 'pdf' },
    ]);
    expect(next).not.toBe(base);
  });

  it('clears the filter when values is empty', () => {
    const next = applyReplaceRefinerValues(base, 'FileType', []);
    expect(next).toEqual([{ filterName: 'Author', value: 'jdoe' }]);
  });

  it('ignores values whose filterName does not match the target', () => {
    const next = applyReplaceRefinerValues(base, 'FileType', [
      { filterName: 'FileType', value: 'docx' },
      { filterName: 'BogusName', value: 'should-be-ignored' },
    ]);
    expect(next).toEqual([
      { filterName: 'Author', value: 'jdoe' },
      { filterName: 'FileType', value: 'docx' },
    ]);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx heft test --test-path-pattern applyReplaceRefinerValues`
Expected: FAIL with `Cannot find module '@webparts/spSearchFilters/components/SpSearchFilters' has no exported member 'applyReplaceRefinerValues'`.

- [ ] **Step 3: Add helper to SpSearchFilters.tsx**

In `src/webparts/spSearchFilters/components/SpSearchFilters.tsx`, add this exported helper at module scope (suggested location: after `buildNextFilters` near line 240):

```ts
/**
 * Pure helper used by the multi-value batched callback. Removes every active
 * filter matching `filterName`, then appends the supplied values (filtering
 * out any that name a different filter). Always returns a new array reference
 * so the Zustand store + orchestrator subscribe trigger reliably.
 */
export function applyReplaceRefinerValues(
  current: IActiveFilter[],
  filterName: string,
  values: IActiveFilter[]
): IActiveFilter[] {
  const kept = current.filter(function (f: IActiveFilter): boolean {
    return f.filterName !== filterName;
  });
  const accepted = values.filter(function (v: IActiveFilter): boolean {
    return v.filterName === filterName;
  });
  return kept.concat(accepted);
}
```

- [ ] **Step 4: Wire the parent callback**

In the same file, add `handleReplaceRefinerValues` next to `handleToggleRefiner` (around line 454):

```ts
/** Multi-value batched: replace all values for a single filterName in one call. */
function handleReplaceRefinerValues(payload: {
  filterName: string;
  values: IActiveFilter[];
}): void {
  if (!store) {
    return;
  }

  if (applyMode === 'instant') {
    const replaced = applyReplaceRefinerValues(filters, payload.filterName, payload.values);
    const nextFilters = clearDependentFilters(replaced, payload.filterName, configs);
    store.setState({ activeFilters: nextFilters, currentPage: 1 });
  } else {
    const current: IActiveFilter[] = hasPendingChanges ? pendingFilters : filters;
    const replaced = applyReplaceRefinerValues(current, payload.filterName, payload.values);
    const updated = clearDependentFilters(replaced, payload.filterName, configs);
    setPendingFilters(updated);
    setHasPendingChanges(!areFiltersEqual(updated, filters));
  }
}
```

Then add `onReplaceRefinerValues={handleReplaceRefinerValues}` to the props passed to the filter component render loop (search for `onToggleRefiner={handleToggleRefiner}` in the JSX and add the new prop alongside it).

- [ ] **Step 5: Extend `IFilterTypeProps` (or the per-component prop interfaces)**

The exact prop shape lives on each filter component. The minimum: pass `onReplaceRefinerValues` as an optional prop on `ITagBoxFilterProps`, `IPeoplePickerFilterProps`, `ITaxonomyTreeFilterProps`, `IDropdownFilterProps`. Concrete change in each component's prop interface:

```ts
export interface ITagBoxFilterProps {
  // ... existing fields ...
  onReplaceRefinerValues?: (payload: { filterName: string; values: IActiveFilter[] }) => void;
}
```

(Same shape pasted into the other 3 interfaces. Add this to their prop interfaces only — the implementation switch happens in the dedicated tasks below.)

- [ ] **Step 6: Run test to verify it passes**

Run: `npx heft test --test-path-pattern applyReplaceRefinerValues`
Expected: PASS, 4 tests.

- [ ] **Step 7: Run full build + tests**

Run: `npx heft build && npx heft test --silent`
Expected: Build green, all 442 (+ 4 new = 446) tests pass.

- [ ] **Step 8: Commit**

```bash
git add tests/webparts/spSearchFilters/applyReplaceRefinerValues.test.ts \
        src/webparts/spSearchFilters/components/SpSearchFilters.tsx \
        src/webparts/spSearchFilters/components/TagBoxFilter.tsx \
        src/webparts/spSearchFilters/components/PeoplePickerFilter.tsx \
        src/webparts/spSearchFilters/components/TaxonomyTreeFilter.tsx \
        src/webparts/spSearchFilters/components/DropdownFilter.tsx
git commit -m "$(cat <<'EOF'
feat(filters): add batched onReplaceRefinerValues callback

Foundation for the multi-toggle clobber fix. handleReplaceRefinerValues
accepts the full intended selection for one filterName in one call;
applyReplaceRefinerValues helper handles the array transform purely.

No component migrated yet — existing per-delta callbacks still in use.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 2: Migrate TagBoxFilter to batched callback

**Files:**
- Modify: `src/webparts/spSearchFilters/components/TagBoxFilter.tsx` (lines 108-129 — the multi-toggle loop)
- Create test: `tests/webparts/spSearchFilters/multiValueCascade.test.ts` (start with the TagBox case; we'll grow the file across tasks 2-5)

- [ ] **Step 1: Write the failing test**

Create `tests/webparts/spSearchFilters/multiValueCascade.test.ts`:

```ts
import { render, fireEvent } from '@testing-library/react';
import * as React from 'react';
import TagBoxFilter from '@webparts/spSearchFilters/components/TagBoxFilter';
import type { IActiveFilter, IFilterConfig, IRefinerValue } from '@interfaces/index';

describe('TagBoxFilter batched callback', () => {
  it('emits a single onReplaceRefinerValues call with all selected values', () => {
    const onReplaceRefinerValues = jest.fn();
    const onToggleRefiner = jest.fn();
    const config: IFilterConfig = {
      managedProperty: 'FileType',
      displayName: 'File type',
      filterType: 'tagbox',
    };
    const values: IRefinerValue[] = [
      { name: 'pdf', value: '"pdf"', count: 10 },
      { name: 'docx', value: '"docx"', count: 5 },
      { name: 'xlsx', value: '"xlsx"', count: 3 },
    ];
    const activeFilters: IActiveFilter[] = [];

    const { getByRole } = render(
      <TagBoxFilter
        filterName="FileType"
        values={values}
        config={config}
        activeFilters={activeFilters}
        onToggleRefiner={onToggleRefiner}
        onReplaceRefinerValues={onReplaceRefinerValues}
      />
    );

    // Simulate selecting all 3 values at once (paste-multi pattern).
    // The TagBox onValueChange fires once with the full array of selected keys.
    const tagBox = getByRole('combobox') as HTMLInputElement;
    fireEvent.change(tagBox, {
      target: { selectedItems: ['"pdf"', '"docx"', '"xlsx"'] }
    });

    expect(onReplaceRefinerValues).toHaveBeenCalledTimes(1);
    expect(onReplaceRefinerValues.mock.calls[0][0]).toEqual({
      filterName: 'FileType',
      values: [
        expect.objectContaining({ filterName: 'FileType', value: '"pdf"' }),
        expect.objectContaining({ filterName: 'FileType', value: '"docx"' }),
        expect.objectContaining({ filterName: 'FileType', value: '"xlsx"' }),
      ],
    });
    expect(onToggleRefiner).not.toHaveBeenCalled();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx heft test --test-path-pattern multiValueCascade`
Expected: FAIL — TagBoxFilter currently loops `onToggleRefiner` calls and doesn't have `onReplaceRefinerValues` plumbed.

- [ ] **Step 3: Replace the loop in TagBoxFilter.tsx**

Open `src/webparts/spSearchFilters/components/TagBoxFilter.tsx`, locate lines 108-129 (the `handleValueChange` function with the for-loops calling `onToggleRefiner`). Replace with:

```ts
function handleValueChange(e: { value?: string[] }): void {
  const nextSelectedTokens: string[] = Array.isArray(e?.value) ? e.value : [];

  if (onReplaceRefinerValues) {
    const batchedValues: IActiveFilter[] = nextSelectedTokens.map(function (token: string): IActiveFilter {
      return {
        filterName,
        value: token,
        displayValue: labelForToken(token, values),
        operator,
      };
    });
    onReplaceRefinerValues({ filterName, values: batchedValues });
    return;
  }

  // Fallback for callers that haven't migrated (shouldn't happen in our parent,
  // but keeps the component standalone-usable).
  const previousTokens = getSelectedRefinerTokens(filterName, values, activeFilters);
  const added = nextSelectedTokens.filter(function (t: string): boolean {
    return previousTokens.indexOf(t) < 0;
  });
  const removed = previousTokens.filter(function (t: string): boolean {
    return nextSelectedTokens.indexOf(t) < 0;
  });
  for (let i = 0; i < added.length; i++) {
    onToggleRefiner({ filterName, value: added[i], displayValue: labelForToken(added[i], values), operator });
  }
  for (let i = 0; i < removed.length; i++) {
    onToggleRefiner({ filterName, value: removed[i], operator });
  }
}

function labelForToken(token: string, refinerValues: IRefinerValue[]): string | undefined {
  for (let i = 0; i < refinerValues.length; i++) {
    if (refinerValues[i].value === token) {
      return refinerValues[i].name;
    }
  }
  return undefined;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx heft test --test-path-pattern multiValueCascade`
Expected: PASS, 1 test.

- [ ] **Step 5: Run full test suite**

Run: `npx heft test --silent`
Expected: All tests pass (446 → 447).

- [ ] **Step 6: Commit**

```bash
git add tests/webparts/spSearchFilters/multiValueCascade.test.ts \
        src/webparts/spSearchFilters/components/TagBoxFilter.tsx
git commit -m "$(cat <<'EOF'
fix(filters): TagBoxFilter emits one batched onReplaceRefinerValues call

Replaces the per-delta onToggleRefiner loop that captured a stale React
closure and silently dropped all but the last selection delta on bulk
changes (paste-multi, programmatic selection).

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 3: Migrate DropdownFilter (multi mode) to batched callback

**Files:**
- Modify: `src/webparts/spSearchFilters/components/DropdownFilter.tsx` (lines 83-105 — the `syncToSelection` loop)
- Modify test: `tests/webparts/spSearchFilters/multiValueCascade.test.ts` (add DropdownFilter `describe` block)

- [ ] **Step 1: Add the failing test to multiValueCascade.test.ts**

Append to the existing test file:

```ts
import DropdownFilter from '@webparts/spSearchFilters/components/DropdownFilter';

describe('DropdownFilter (multi) batched callback', () => {
  it('emits a single onReplaceRefinerValues call when multi-selection changes', () => {
    const onReplaceRefinerValues = jest.fn();
    const onToggleRefiner = jest.fn();
    const config: IFilterConfig = {
      managedProperty: 'Department',
      displayName: 'Department',
      filterType: 'dropdown',
      multiValues: true,
    };
    const values: IRefinerValue[] = [
      { name: 'Finance', value: '"Finance"', count: 50 },
      { name: 'Legal', value: '"Legal"', count: 12 },
      { name: 'Engineering', value: '"Engineering"', count: 87 },
    ];

    const { getByRole } = render(
      <DropdownFilter
        filterName="Department"
        values={values}
        config={config}
        activeFilters={[]}
        onToggleRefiner={onToggleRefiner}
        onReplaceRefinerValues={onReplaceRefinerValues}
      />
    );

    // Simulate selecting two options at once.
    const dropdown = getByRole('combobox');
    fireEvent.change(dropdown, {
      target: { selectedOptions: ['"Finance"', '"Legal"'] }
    });

    expect(onReplaceRefinerValues).toHaveBeenCalledTimes(1);
    expect(onReplaceRefinerValues.mock.calls[0][0].values).toHaveLength(2);
    expect(onToggleRefiner).not.toHaveBeenCalled();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx heft test --test-path-pattern multiValueCascade`
Expected: FAIL on the new Dropdown test (TagBox test still passes).

- [ ] **Step 3: Replace the syncToSelection loop in DropdownFilter.tsx**

Open `src/webparts/spSearchFilters/components/DropdownFilter.tsx`, locate `syncToSelection` (around line 83). Replace its body:

```ts
function syncToSelection(nextKeys: string[]): void {
  if (config?.multiValues !== false && onReplaceRefinerValues) {
    const batchedValues: IActiveFilter[] = nextKeys.map(function (key: string): IActiveFilter {
      return {
        filterName,
        value: key,
        displayValue: findRefinerLabel(key, values),
        operator,
      };
    });
    onReplaceRefinerValues({ filterName, values: batchedValues });
    return;
  }

  // Single-select mode — preserved per-toggle path.
  const previousKeys = getSelectedRefinerTokens(filterName, values, activeFilters);
  const added = nextKeys.filter(function (k: string): boolean { return previousKeys.indexOf(k) < 0; });
  const removed = previousKeys.filter(function (k: string): boolean { return nextKeys.indexOf(k) < 0; });
  for (let i = 0; i < added.length; i++) {
    onToggleRefiner({ filterName, value: added[i], displayValue: findRefinerLabel(added[i], values), operator });
  }
  for (let i = 0; i < removed.length; i++) {
    onToggleRefiner({ filterName, value: removed[i], operator });
  }
}

function findRefinerLabel(token: string, refinerValues: IRefinerValue[]): string | undefined {
  for (let i = 0; i < refinerValues.length; i++) {
    if (refinerValues[i].value === token) {
      return refinerValues[i].name;
    }
  }
  return undefined;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx heft test --test-path-pattern multiValueCascade`
Expected: PASS, 2 tests.

- [ ] **Step 5: Run full test suite**

Run: `npx heft test --silent`
Expected: All tests pass (447 → 448).

- [ ] **Step 6: Commit**

```bash
git add tests/webparts/spSearchFilters/multiValueCascade.test.ts \
        src/webparts/spSearchFilters/components/DropdownFilter.tsx
git commit -m "$(cat <<'EOF'
fix(filters): DropdownFilter (multi) emits batched onReplaceRefinerValues

Same fix shape as TagBoxFilter. Single-select dropdown unchanged.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 4: Migrate TaxonomyTreeFilter to batched callback

**Files:**
- Modify: `src/webparts/spSearchFilters/components/TaxonomyTreeFilter.tsx` (lines 223-261 — `handleSelectionChanged`; line 205 — useEffect deps)
- Modify test: `tests/webparts/spSearchFilters/multiValueCascade.test.ts` (add Taxonomy case)

Note: This component is going to be **replaced** in PR2 with PnP's `TaxonomyPicker`. Migrating it now is still worth it because (a) PR2 may take a couple of days and we need the bug fix live before then, (b) the migration is small, (c) it also fixes Issue G (selection persistence) by stabilising the `selectedKeys` derivation.

- [ ] **Step 1: Add the failing test**

Append to `tests/webparts/spSearchFilters/multiValueCascade.test.ts`:

```ts
import { buildTaxonomyBatchPayload } from '@webparts/spSearchFilters/components/TaxonomyTreeFilter';

describe('buildTaxonomyBatchPayload', () => {
  it('maps selected node keys to a single batched payload', () => {
    const tokenMap = new Map<string, string>([
      ['aaaa-bbbb-cccc-1', 'GP0|#aaaa-bbbb-cccc-1'],
      ['aaaa-bbbb-cccc-2', 'GP0|#aaaa-bbbb-cccc-2'],
    ]);
    const labelMap = new Map<string, string>([
      ['aaaa-bbbb-cccc-1', 'Electronics'],
      ['aaaa-bbbb-cccc-2', 'Books'],
    ]);

    const payload = buildTaxonomyBatchPayload(
      'owstaxIdProductCategory',
      ['aaaa-bbbb-cccc-1', 'aaaa-bbbb-cccc-2'],
      tokenMap,
      labelMap,
      'eq'
    );

    expect(payload).toEqual({
      filterName: 'owstaxIdProductCategory',
      values: [
        {
          filterName: 'owstaxIdProductCategory',
          value: 'GP0|#aaaa-bbbb-cccc-1',
          displayValue: 'Electronics',
          operator: 'eq',
        },
        {
          filterName: 'owstaxIdProductCategory',
          value: 'GP0|#aaaa-bbbb-cccc-2',
          displayValue: 'Books',
          operator: 'eq',
        },
      ],
    });
  });

  it('falls back to GP0|#<guid> when tokenMap has no entry for a key', () => {
    const payload = buildTaxonomyBatchPayload(
      'Cat',
      ['unmapped-guid'],
      new Map(),
      new Map(),
      'eq'
    );
    expect(payload.values[0].value).toBe('GP0|#unmapped-guid');
    expect(payload.values[0].displayValue).toBeUndefined();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx heft test --test-path-pattern multiValueCascade`
Expected: FAIL — `buildTaxonomyBatchPayload` is not exported.

- [ ] **Step 3: Extract `buildTaxonomyBatchPayload` and call it from `handleSelectionChanged`**

Open `src/webparts/spSearchFilters/components/TaxonomyTreeFilter.tsx`. Add this exported pure helper at module scope (above the component definition):

```ts
/**
 * Pure helper — builds the batched onReplaceRefinerValues payload for the
 * taxonomy filter from the TreeView's currently-selected node keys.
 * Exported so the unit test can drive the conversion logic without
 * mounting a real DevExtreme TreeView.
 */
export function buildTaxonomyBatchPayload(
  filterName: string,
  selectedKeys: string[],
  tokenMap: Map<string, string>,
  labelMap: Map<string, string>,
  operator: string
): { filterName: string; values: IActiveFilter[] } {
  return {
    filterName,
    values: selectedKeys.map(function (key: string): IActiveFilter {
      return {
        filterName,
        value: tokenMap.get(key) || ('GP0|#' + key),
        displayValue: labelMap.get(key),
        operator,
      };
    }),
  };
}
```

Then replace `handleSelectionChanged` (lines 223-261) with:

```ts
function handleSelectionChanged(e: { component?: { getSelectedNodeKeys(): string[] } }): void {
  if (!e || !e.component) {
    return;
  }
  const keys: string[] = e.component.getSelectedNodeKeys();

  if (onReplaceRefinerValues) {
    onReplaceRefinerValues(
      buildTaxonomyBatchPayload(filterName, keys, tokenMap, labelMap, operator)
    );
    return;
  }

  // Fallback path (no batched callback wired): per-delta. Preserved for
  // back-compat but the parent always wires onReplaceRefinerValues.
  const selectedTokens: string[] = [];
  for (let i = 0; i < keys.length; i++) {
    const token = tokenMap.get(keys[i]) || buildToken(keys[i]);
    selectedTokens.push(token);
  }
  const currentValues: string[] = [];
  for (let i = 0; i < activeFilters.length; i++) {
    if (activeFilters[i].filterName === filterName) {
      currentValues.push(activeFilters[i].value);
    }
  }
  for (let i = 0; i < selectedTokens.length; i++) {
    if (currentValues.indexOf(selectedTokens[i]) < 0) {
      onToggleRefiner({
        filterName,
        value: selectedTokens[i],
        displayValue: labelMap.get(keys[i]) || undefined,
        operator,
      });
    }
  }
  for (let i = 0; i < currentValues.length; i++) {
    if (selectedTokens.indexOf(currentValues[i]) < 0) {
      onToggleRefiner({ filterName, value: currentValues[i], operator });
    }
  }
}
```

- [ ] **Step 4: Also fix the useEffect dependency that triggers loadTree on every search response (Issue G)**

In the same file, locate the `useEffect` at line 205 (the one calling `loadTree`). Change the dep array from:

```ts
}, [values, config?.termSetId, showCount, countMap]);
```

to:

```ts
}, [config?.termSetId]);  // Only re-load when term set ID changes
```

And inside `loadTree`, derive count updates from `countMap` directly inside `mapTermTree` without dependency on the effect re-run.

If `loadTree` reads `values` to build the fallback tree (when no termSetId), keep a separate `useEffect` that updates `treeItems` from `values` WITHOUT calling `loadTree`:

```ts
React.useEffect(function (): void {
  if (config?.termSetId) {
    return;  // Tree comes from term set, not from refiner values
  }
  setTreeItems(buildFallbackItems(values, showCount));
}, [values, config?.termSetId, showCount]);
```

This stops the full re-load on every search response, which is what was wiping the visible selection.

- [ ] **Step 5: Run test to verify it passes**

Run: `npx heft test --test-path-pattern multiValueCascade`
Expected: PASS, 3 tests.

- [ ] **Step 6: Run full test suite**

Run: `npx heft test --silent`
Expected: All pass (448 → 449).

- [ ] **Step 7: Commit**

```bash
git add tests/webparts/spSearchFilters/multiValueCascade.test.ts \
        src/webparts/spSearchFilters/components/TaxonomyTreeFilter.tsx
git commit -m "$(cat <<'EOF'
fix(filters): TaxonomyTreeFilter emits batched callback + persists selection

Replaces the per-delta onToggleRefiner loop with one onReplaceRefinerValues
call (Issue A fix). Splits the loadTree useEffect so it only re-runs when
the term set ID changes — the previous deps re-triggered the DevExtreme
TreeView mount on every search response and wiped the visual selection
(Issue G fix).

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 5: Replace PeoplePicker — Fluent NormalPeoplePicker + batched callback (Issues E + A combined)

**Files:**
- Modify: `src/webparts/spSearchFilters/components/PeoplePickerFilter.tsx` (whole file rewrite)
- Modify test: `tests/webparts/spSearchFilters/multiValueCascade.test.ts` (add People case)

This task is bigger than the others because we're swapping the third-party broken control AND migrating to the batched callback in one shot — no point doing them separately when both touch the same file.

- [ ] **Step 1: Add the failing test**

Append to `tests/webparts/spSearchFilters/multiValueCascade.test.ts`:

```ts
import PeoplePickerFilter from '@webparts/spSearchFilters/components/PeoplePickerFilter';

describe('PeoplePickerFilter (Fluent) batched callback', () => {
  it('emits a single onReplaceRefinerValues call when persona selection changes', () => {
    const onReplaceRefinerValues = jest.fn();
    const onToggleRefiner = jest.fn();
    const config: IFilterConfig = {
      managedProperty: 'Author',
      displayName: 'Author',
      filterType: 'people',
    };

    // PeoplePickerFilter exposes __test_onItemsChange so we can drive
    // NormalPeoplePicker's onChange handler without spinning up the picker UI.
    const captureHandler: { fn?: (personas: Array<{ secondaryText: string; text: string }>) => void } = {};
    render(
      <PeoplePickerFilter
        filterName="Author"
        values={[]}
        config={config}
        activeFilters={[]}
        onToggleRefiner={onToggleRefiner}
        onReplaceRefinerValues={onReplaceRefinerValues}
        __test_captureHandler={captureHandler}
      />
    );

    expect(captureHandler.fn).toBeDefined();
    captureHandler.fn!([
      { secondaryText: 'i:0#.f|membership|jdoe@contoso.com', text: 'John Doe' },
      { secondaryText: 'i:0#.f|membership|asmith@contoso.com', text: 'Alice Smith' },
    ]);

    expect(onReplaceRefinerValues).toHaveBeenCalledTimes(1);
    expect(onReplaceRefinerValues.mock.calls[0][0].values).toHaveLength(2);
    expect(onReplaceRefinerValues.mock.calls[0][0].values[0].value).toContain('jdoe');
    expect(onToggleRefiner).not.toHaveBeenCalled();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx heft test --test-path-pattern multiValueCascade`
Expected: FAIL — `__test_captureHandler` not a recognised prop.

- [ ] **Step 3: Rewrite PeoplePickerFilter.tsx**

Replace the entire contents of `src/webparts/spSearchFilters/components/PeoplePickerFilter.tsx` with:

```tsx
import * as React from 'react';
import {
  NormalPeoplePicker,
  IPersonaProps,
  ValidationState,
} from '@fluentui/react/lib/Pickers';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import styles from './SpSearchFilters.module.scss';
import type {
  IActiveFilter,
  IFilterConfig,
  IRefinerValue,
} from '@interfaces/index';

export interface IPeoplePickerFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
  onReplaceRefinerValues?: (payload: { filterName: string; values: IActiveFilter[] }) => void;
  /** Test-only — captures the onChange handler for unit tests. */
  __test_captureHandler?: { fn?: (personas: IPersonaProps[]) => void };
}

interface IClaimSuggestion {
  loginName: string;
  displayName: string;
  email?: string;
}

async function resolvePeople(filter: string): Promise<IClaimSuggestion[]> {
  if (!filter || filter.length < 2) {
    return [];
  }
  const principalType = 1 | 4;  // User + SecurityGroup
  const body = {
    queryParams: {
      QueryString: filter,
      MaximumEntitySuggestions: 25,
      PrincipalSource: 15,         // All
      PrincipalType: principalType,
      AllowEmailAddresses: true,
    },
  };
  const url = SPContext.webAbsoluteUrl + '/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser';
  const response = await SPContext.http.post<{ value: string }>(url, body);
  if (!response.ok || !response.data?.value) {
    return [];
  }
  const parsed = JSON.parse(response.data.value) as Array<{
    Key: string;
    DisplayText: string;
    EntityData: { Email?: string };
  }>;
  return parsed.map(function (p): IClaimSuggestion {
    return {
      loginName: p.Key,
      displayName: p.DisplayText,
      email: p.EntityData?.Email,
    };
  });
}

const PeoplePickerFilter: React.FC<IPeoplePickerFilterProps> = (props) => {
  const {
    filterName,
    config,
    activeFilters,
    onToggleRefiner,
    onReplaceRefinerValues,
    __test_captureHandler,
  } = props;

  const operator = config?.operator || 'eq';

  const initialPersonas = React.useMemo(function (): IPersonaProps[] {
    return activeFilters
      .filter(function (f: IActiveFilter): boolean { return f.filterName === filterName; })
      .map(function (f: IActiveFilter): IPersonaProps {
        return {
          text: f.displayValue || f.value,
          secondaryText: f.value,
        };
      });
  }, [activeFilters, filterName]);

  const [selectedPersonas, setSelectedPersonas] = React.useState<IPersonaProps[]>(initialPersonas);

  React.useEffect(function (): void {
    setSelectedPersonas(initialPersonas);
  }, [initialPersonas]);

  function handleResolveSuggestions(filter: string, currentPersonas: IPersonaProps[] = []): Promise<IPersonaProps[]> {
    return resolvePeople(filter).then(function (claims: IClaimSuggestion[]): IPersonaProps[] {
      const currentLoginNames = currentPersonas
        .map(function (p: IPersonaProps): string { return p.secondaryText || ''; });
      return claims
        .filter(function (c: IClaimSuggestion): boolean { return currentLoginNames.indexOf(c.loginName) < 0; })
        .map(function (c: IClaimSuggestion): IPersonaProps {
          return {
            text: c.displayName,
            secondaryText: c.loginName,
            tertiaryText: c.email,
          };
        });
    });
  }

  function handleItemsChange(items?: IPersonaProps[]): void {
    const next = items || [];
    setSelectedPersonas(next);

    if (onReplaceRefinerValues) {
      const batched: IActiveFilter[] = next.map(function (p: IPersonaProps): IActiveFilter {
        return {
          filterName,
          value: p.secondaryText || '',
          displayValue: p.text,
          operator,
        };
      });
      onReplaceRefinerValues({ filterName, values: batched });
      return;
    }

    // Fallback (back-compat) — same shape as before this change.
    const previous = activeFilters
      .filter(function (f: IActiveFilter): boolean { return f.filterName === filterName; })
      .map(function (f: IActiveFilter): string { return f.value; });
    const nextLogins = next.map(function (p): string { return p.secondaryText || ''; });
    for (let i = 0; i < nextLogins.length; i++) {
      if (previous.indexOf(nextLogins[i]) < 0) {
        onToggleRefiner({ filterName, value: nextLogins[i], displayValue: next[i].text, operator });
      }
    }
    for (let i = 0; i < previous.length; i++) {
      if (nextLogins.indexOf(previous[i]) < 0) {
        onToggleRefiner({ filterName, value: previous[i], operator });
      }
    }
  }

  // Test hook
  React.useEffect(function (): void {
    if (__test_captureHandler) {
      __test_captureHandler.fn = handleItemsChange;
    }
    return function (): void {
      if (__test_captureHandler) {
        __test_captureHandler.fn = undefined;
      }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div className={styles.peoplePickerFilterContainer}>
      <NormalPeoplePicker
        onResolveSuggestions={handleResolveSuggestions}
        selectedItems={selectedPersonas}
        onChange={handleItemsChange}
        onValidateInput={(): ValidationState => ValidationState.invalid}
        pickerSuggestionsProps={{
          suggestionsHeaderText: 'Suggested people',
          noResultsFoundText: 'No matches',
        }}
        resolveDelay={250}
      />
    </div>
  );
};

export default PeoplePickerFilter;
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx heft test --test-path-pattern multiValueCascade`
Expected: PASS, 4 tests.

- [ ] **Step 5: Run full test suite**

Run: `npx heft build && npx heft test --silent`
Expected: Build green; all tests pass (449 → 450).

- [ ] **Step 6: Commit**

```bash
git add tests/webparts/spSearchFilters/multiValueCascade.test.ts \
        src/webparts/spSearchFilters/components/PeoplePickerFilter.tsx
git commit -m "$(cat <<'EOF'
fix(filters): replace PnP PeoplePicker with Fluent NormalPeoplePicker

The @pnp/spfx-controls-react PeoplePicker crashed with "Cannot read
properties of undefined (reading 'defaultClass')" because its
PeoplePickerComponent.module.scss JS import resolves to undefined under
SPFx 1.22's sp-css-loader (same root cause as PropertyFieldCollectionData
which we already replaced).

Switches to Fluent's first-party NormalPeoplePicker, using a direct
clientPeoplePickerSearchUser REST call via SPContext.http for claim
resolution. Same single-batched callback shape as the other multi-value
filters.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 6: Toggle defaultValue end-to-end (Issue D)

**Files:**
- Modify: `src/libraries/spSearchStore/interfaces/IFilterTypes.ts` (add `defaultValue?: boolean` to `IFilterConfig`)
- Modify: `src/webparts/spSearchFilters/components/ToggleFilter.tsx` (already reads from activeFilters — no change needed unless config-aware initial render is wanted; verify)
- Modify: `src/libraries/spSearchStore/store/storeRegistry.ts` (seed activeFilters from toggle defaults during init)
- Modify: `src/propertyPaneControls/filtersCollection/FiltersCollectionControl.tsx` (add a Toggle control to the "Toggle labels" section)
- Modify: `src/propertyPaneControls/filtersCollection/fieldRelevance.ts` (add `defaultValue` to the `'toggle'` relevance entry)
- Create test: `tests/store/toggleDefaultValueSeed.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/store/toggleDefaultValueSeed.test.ts`:

```ts
import { seedToggleDefaults } from '@store/store/storeRegistry';
import type { IFilterConfig, IActiveFilter } from '@interfaces/index';

describe('seedToggleDefaults', () => {
  const toggleConfig: IFilterConfig = {
    managedProperty: 'IsConfidential',
    displayName: 'Confidential only',
    filterType: 'toggle',
    defaultValue: true,
  };

  it('seeds a synthetic active filter when no URL state present', () => {
    const seeded = seedToggleDefaults([], [toggleConfig]);
    expect(seeded).toHaveLength(1);
    expect(seeded[0]).toEqual({
      filterName: 'IsConfidential',
      value: 'true',
      operator: 'eq',
    });
  });

  it('does NOT override an existing URL-restored active filter', () => {
    const urlRestored: IActiveFilter[] = [
      { filterName: 'IsConfidential', value: 'false', operator: 'eq' },
    ];
    const seeded = seedToggleDefaults(urlRestored, [toggleConfig]);
    expect(seeded).toEqual(urlRestored);
  });

  it('ignores configs without defaultValue', () => {
    const noDefault: IFilterConfig = { ...toggleConfig, defaultValue: undefined };
    const seeded = seedToggleDefaults([], [noDefault]);
    expect(seeded).toEqual([]);
  });

  it('ignores non-toggle filter types even with defaultValue set', () => {
    const checkboxLike: IFilterConfig = {
      ...toggleConfig,
      filterType: 'checkbox',
      defaultValue: true,
    };
    const seeded = seedToggleDefaults([], [checkboxLike]);
    expect(seeded).toEqual([]);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx heft test --test-path-pattern toggleDefaultValueSeed`
Expected: FAIL — `seedToggleDefaults` not exported.

- [ ] **Step 3: Add `defaultValue` to IFilterConfig**

In `src/libraries/spSearchStore/interfaces/IFilterTypes.ts`, locate the `IFilterConfig` interface (lines 55-105) and add:

```ts
/**
 * Initial value for a 'toggle' filter when no URL / restored state is present.
 * URL-restore always wins over this default.
 */
defaultValue?: boolean;
```

- [ ] **Step 4: Add `seedToggleDefaults` helper to storeRegistry.ts**

Open `src/libraries/spSearchStore/store/storeRegistry.ts`, add this exported helper at module scope (above `initializeSearchContext`):

```ts
import type { IActiveFilter, IFilterConfig } from '@interfaces/index';

/**
 * Pure helper: for each toggle filter with `defaultValue` set whose property
 * is not already active in the URL-restored state, push a synthetic active
 * filter. Used during initializeSearchContext after URL restore runs.
 */
export function seedToggleDefaults(
  current: IActiveFilter[],
  configs: IFilterConfig[]
): IActiveFilter[] {
  const activeNames = new Set<string>();
  for (let i = 0; i < current.length; i++) {
    activeNames.add(current[i].filterName);
  }
  const additions: IActiveFilter[] = [];
  for (let i = 0; i < configs.length; i++) {
    const c = configs[i];
    if (c.filterType !== 'toggle') {
      continue;
    }
    if (c.defaultValue === undefined) {
      continue;
    }
    if (activeNames.has(c.managedProperty)) {
      continue;
    }
    additions.push({
      filterName: c.managedProperty,
      value: c.defaultValue ? 'true' : 'false',
      operator: 'eq',
    });
  }
  if (additions.length === 0) {
    return current;
  }
  return current.concat(additions);
}
```

- [ ] **Step 5: Wire `seedToggleDefaults` into `initializeSearchContext`**

In the same file, find where URL-restore writes activeFilters to the store (around line 278, comment "T3.D6 — defaults; admins override via initializeSearchContext options."). Right after that line, add:

```ts
const seededFilters = seedToggleDefaults(
  store.getState().activeFilters,
  store.getState().filterConfig || []
);
if (seededFilters !== store.getState().activeFilters) {
  store.setState({ activeFilters: seededFilters });
}
```

- [ ] **Step 6: Add the property pane Toggle control**

In `src/propertyPaneControls/filtersCollection/FiltersCollectionControl.tsx`, locate the "Toggle labels" Section (around line 285-308). Add a Toggle control inside it:

```tsx
<Toggle
  label="Default value"
  checked={!!draft.defaultValue}
  inlineLabel
  onChange={(_e, checked): void => updateDraftField('defaultValue', checked === true)}
  onText="On"
  offText="Off"
/>
```

(`updateDraftField` is the existing helper that updates the draft refiner — verify exact name when implementing.)

- [ ] **Step 7: Add `defaultValue` to fieldRelevance.ts**

In `src/propertyPaneControls/filtersCollection/fieldRelevance.ts`, locate the `'toggle'` entry in the relevance map (around lines 36-38) and add `defaultValue` to the visible fields list.

- [ ] **Step 8: Run test to verify it passes**

Run: `npx heft test --test-path-pattern toggleDefaultValueSeed`
Expected: PASS, 4 tests.

- [ ] **Step 9: Run full test suite**

Run: `npx heft build && npx heft test --silent`
Expected: Build green; all tests pass (450 → 454).

- [ ] **Step 10: Commit**

```bash
git add src/libraries/spSearchStore/interfaces/IFilterTypes.ts \
        src/libraries/spSearchStore/store/storeRegistry.ts \
        src/propertyPaneControls/filtersCollection/FiltersCollectionControl.tsx \
        src/propertyPaneControls/filtersCollection/fieldRelevance.ts \
        tests/store/toggleDefaultValueSeed.test.ts
git commit -m "$(cat <<'EOF'
feat(filters): configurable defaultValue for Toggle filter type

Admin sets a default in the property pane; initializeSearchContext seeds
the synthetic active filter on first load when no URL state is present.
URL-restore wins over defaults so shareable links keep working.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 7: Remove TagBox CSS overrides (Issue H)

**Files:**
- Modify: `src/webparts/spSearchFilters/components/SpSearchFilters.module.scss` (delete lines 622-647 — the `:global` block targeting `.dx-tagbox` / `.dx-tag` / `.dx-tag-remove-button` / `.dx-list-item-selected`)

- [ ] **Step 1: Delete the overrides**

Open `src/webparts/spSearchFilters/components/SpSearchFilters.module.scss`. Locate lines 614-648 (the `/* ── TagBox filter ──────────────────────────────────────── */` section). Keep the `.tagBoxFilterContainer` rule (lines 616-621). Delete the entire `:global { ... }` block (lines 622-647). The result should look like:

```scss
/* ── TagBox filter ──────────────────────────────────────── */

.tagBoxFilterContainer {
  padding: 4px 0;
  display: flex;
  flex-direction: column;
  gap: 4px;
}
```

- [ ] **Step 2: Run the build**

Run: `npx heft build`
Expected: Build green.

- [ ] **Step 3: Manual verification**

Run: `npm start`
Open the workbench, configure a TagBox refiner, select a value. The pill should match DevExtreme's stock light theme (same grey-on-white as the DataGrid's tag cells). This is the intended look.

- [ ] **Step 4: Commit**

```bash
git add src/webparts/spSearchFilters/components/SpSearchFilters.module.scss
git commit -m "$(cat <<'EOF'
style(filters): drop failed Fluent overrides on DevExtreme TagBox

The Fluent-blue pill overrides lost the CSS specificity fight with
dx.light.css, so users saw a grey pill anyway. Letting DevExtreme's
native theme render is consistent with the DataGrid's tag cells.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### PR1 milestone — push and open PR (optional checkpoint)

At this point PR1's scope (E + A + G + D + H) is complete. If working through a PR cadence, push and open the PR now:

```bash
git push origin main  # direct-to-main per project cadence
```

Verify in production / workbench:
- Manage filters dialog (already fixed in earlier commit) — opens without crash ✅
- TagBox/People/Taxonomy multi-select — picks land in `activeFilters` correctly ✅
- Toggle with `defaultValue: true` — shows "On" on first load ✅
- TagBox pill — DevExtreme grey ✅

---

# PR2 — Type-aware refiners + Taxonomy UI

### Task 8: Add `dataType` and `valueSplitDelimiter` to IFilterConfig

**Files:**
- Modify: `src/libraries/spSearchStore/interfaces/IFilterTypes.ts` (add two optional fields to `IFilterConfig`)

This is a tiny task by itself — a type-only change — but it unblocks Tasks 9 + 10.

- [ ] **Step 1: Modify IFilterConfig**

In `src/libraries/spSearchStore/interfaces/IFilterTypes.ts`, inside the `IFilterConfig` interface (after the `defaultValue` field added in Task 6), add:

```ts
/**
 * Underlying SharePoint data type of the managed property. Controls
 * value preprocessing in SharePointSearchProvider._mapRefiners.
 * 'auto' (default) runs a heuristic: strip "type;#" prefix when present.
 */
dataType?: 'auto' | 'text' | 'choiceMulti' | 'lookup' | 'calculated' |
           'datetime' | 'yesno' | 'number';

/**
 * If set, refiner values for this filter are split on this delimiter,
 * trimmed, deduplicated, and counts are aggregated per token. Useful
 * for Text columns that store comma/newline-separated tag-like values.
 */
valueSplitDelimiter?: string;
```

- [ ] **Step 2: Verify the build**

Run: `npx heft build`
Expected: Green — no consumers yet, just an interface addition.

- [ ] **Step 3: Commit**

```bash
git add src/libraries/spSearchStore/interfaces/IFilterTypes.ts
git commit -m "$(cat <<'EOF'
feat(filters): add dataType + valueSplitDelimiter to IFilterConfig

Type-only addition. Consumed by _mapRefiners (next commit) and the
property pane Data format section.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 9: `_mapRefiners` value preprocessing pass (Issues B + F)

**Files:**
- Modify: `src/libraries/spSearchStore/providers/SharePointSearchProvider.ts` (extend `_mapRefiners` signature to accept `filterConfig: IFilterConfig[]`; add preprocessing pass)
- Modify: caller sites in the same file that invoke `_mapRefiners` (pass `this._filterConfig` or equivalent)
- Create test: `tests/providers/mapRefinersPreprocessing.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/providers/mapRefinersPreprocessing.test.ts`:

```ts
import { mapRefinersWithPreprocessing } from '@providers/SharePointSearchProvider';
import type { IFilterConfig } from '@interfaces/index';

describe('mapRefinersWithPreprocessing', () => {
  // The actual SP response shape — { Name, Entries: [{ RefinementName, RefinementValue, RefinementCount }] }.
  function refinerResponse(name: string, entries: Array<{ value: string; count: number }>): unknown {
    return {
      Name: name,
      Entries: entries.map(function (e) {
        return {
          RefinementName: e.value,
          RefinementValue: e.value,
          RefinementCount: e.count,
        };
      }),
    };
  }

  it('passes values through unchanged when no config exists', () => {
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('FileType', [{ value: 'pdf', count: 5 }])],
      []
    );
    expect(result[0].values[0].name).toBe('pdf');
    expect(result[0].values[0].count).toBe(5);
  });

  it('strips "string;#" prefix when dataType=choiceMulti', () => {
    const config: IFilterConfig[] = [{
      managedProperty: 'DocType',
      displayName: 'Doc type',
      filterType: 'checkbox',
      dataType: 'choiceMulti',
    }];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('DocType', [
        { value: 'string;#Articles', count: 10 },
        { value: 'string;#Letters', count: 4 },
      ])],
      config
    );
    expect(result[0].values).toEqual([
      expect.objectContaining({ name: 'Articles', count: 10 }),
      expect.objectContaining({ name: 'Letters', count: 4 }),
    ]);
  });

  it('renders empty-after-strip as "(blank)" while preserving original token for KQL', () => {
    const config: IFilterConfig[] = [{
      managedProperty: 'DocType',
      displayName: 'Doc type',
      filterType: 'checkbox',
      dataType: 'choiceMulti',
    }];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('DocType', [
        { value: 'string;#', count: 100 },
      ])],
      config
    );
    expect(result[0].values[0].name).toBe('(blank)');
    expect(result[0].values[0].value).toBe('string;#');
  });

  it('auto-detects "type;#" prefix when dataType is unspecified (auto default)', () => {
    const config: IFilterConfig[] = [{
      managedProperty: 'Amount',
      displayName: 'Amount',
      filterType: 'checkbox',
    }];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('Amount', [
        { value: 'int;#1500', count: 3 },
        { value: 'int;#2000', count: 7 },
      ])],
      config
    );
    expect(result[0].values[0].name).toBe('1500');
    expect(result[0].values[1].name).toBe('2000');
  });

  it('does NOT strip when dataType=text even if "string;#" prefix is present', () => {
    const config: IFilterConfig[] = [{
      managedProperty: 'RawText',
      displayName: 'Raw',
      filterType: 'checkbox',
      dataType: 'text',
    }];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('RawText', [{ value: 'string;#literal', count: 1 }])],
      config
    );
    expect(result[0].values[0].name).toBe('string;#literal');
  });

  it('splits on valueSplitDelimiter, aggregates counts, dedupes tokens', () => {
    const config: IFilterConfig[] = [{
      managedProperty: 'Tags',
      displayName: 'Tags',
      filterType: 'tagbox',
      valueSplitDelimiter: ',',
    }];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('Tags', [
        { value: 'finance, hot, urgent', count: 5 },
        { value: 'finance, archived', count: 3 },
        { value: 'urgent', count: 2 },
      ])],
      config
    );
    const byName = new Map(result[0].values.map(function (v) { return [v.name, v.count]; }));
    expect(byName.get('finance')).toBe(8);
    expect(byName.get('hot')).toBe(5);
    expect(byName.get('urgent')).toBe(7);
    expect(byName.get('archived')).toBe(3);
  });

  it('splits on newline AND strips prefix when both are configured', () => {
    const config: IFilterConfig[] = [{
      managedProperty: 'MultiTags',
      displayName: 'MultiTags',
      filterType: 'tagbox',
      dataType: 'choiceMulti',
      valueSplitDelimiter: '\n',
    }];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('MultiTags', [
        { value: 'string;#alpha\nstring;#beta', count: 4 },
      ])],
      config
    );
    const byName = new Map(result[0].values.map(function (v) { return [v.name, v.count]; }));
    expect(byName.get('alpha')).toBe(4);
    expect(byName.get('beta')).toBe(4);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx heft test --test-path-pattern mapRefinersPreprocessing`
Expected: FAIL — `mapRefinersWithPreprocessing` not exported.

- [ ] **Step 3: Add the preprocessing function to SharePointSearchProvider.ts**

In `src/libraries/spSearchStore/providers/SharePointSearchProvider.ts`, locate `_mapRefiners` (lines 288-308). Refactor:

1. Rename existing `_mapRefiners` to an inner private method but keep its current pass-through behaviour as a fallback.
2. Add a new exported function `mapRefinersWithPreprocessing` at module scope:

```ts
import type { IFilterConfig, IRefinerValue, IRefiner } from '@interfaces/index';

/**
 * Strip "type;#" prefix when appropriate and split values on a configured
 * delimiter. Used by SharePointSearchProvider._mapRefiners.
 *
 * Returns the original raw token in `value` (preserving the KQL filter
 * clause's match semantics) and the cleaned/split display label in `name`.
 */
export function mapRefinersWithPreprocessing(
  refinerResponses: Array<{
    Name: string;
    Entries: Array<{ RefinementName: string; RefinementValue: string; RefinementCount: number }>;
  }>,
  filterConfig: IFilterConfig[]
): IRefiner[] {
  const configByName = new Map<string, IFilterConfig>();
  for (let i = 0; i < filterConfig.length; i++) {
    configByName.set(filterConfig[i].managedProperty, filterConfig[i]);
  }

  return refinerResponses.map(function (r): IRefiner {
    const config = configByName.get(r.Name);
    const valueAggregator = new Map<string, { count: number; rawValue: string }>();

    for (let i = 0; i < r.Entries.length; i++) {
      const entry = r.Entries[i];
      const tokens = splitEntryValue(entry.RefinementName, config?.valueSplitDelimiter);

      for (let j = 0; j < tokens.length; j++) {
        const cleaned = preprocessValue(tokens[j], config?.dataType);
        const displayName = cleaned.length === 0 ? '(blank)' : cleaned;
        const existing = valueAggregator.get(displayName);
        if (existing) {
          existing.count += entry.RefinementCount;
        } else {
          valueAggregator.set(displayName, {
            count: entry.RefinementCount,
            rawValue: entry.RefinementValue,
          });
        }
      }
    }

    const values: IRefinerValue[] = [];
    valueAggregator.forEach(function (agg, name): void {
      values.push({ name, value: agg.rawValue, count: agg.count });
    });

    return { filterName: r.Name, values };
  });
}

function splitEntryValue(raw: string, delimiter: string | undefined): string[] {
  if (!delimiter) {
    return [raw];
  }
  const parts = raw.split(delimiter);
  const trimmed: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    const t = parts[i].trim();
    if (t.length > 0) {
      trimmed.push(t);
    }
  }
  return trimmed.length > 0 ? trimmed : [raw];
}

function preprocessValue(raw: string, dataType: IFilterConfig['dataType']): string {
  // Decide whether to strip the "type;#" prefix.
  const STRIP_TYPES = new Set(['choiceMulti', 'lookup', 'calculated', 'datetime', 'number', 'yesno']);
  const shouldStrip =
    (dataType !== 'text') &&
    (STRIP_TYPES.has(dataType || '') || (dataType === undefined || dataType === 'auto'));

  if (!shouldStrip) {
    return raw;
  }

  // Heuristic: prefix pattern is ^[A-Za-z]+;#
  const m = /^[A-Za-z]+;#(.*)$/.exec(raw);
  if (m) {
    return m[1];
  }
  return raw;
}
```

- [ ] **Step 4: Wire the new function into `_mapRefiners`**

In the same file, change `_mapRefiners` so the instance method now delegates to `mapRefinersWithPreprocessing` with the active `filterConfig` from state. The provider needs to read `filterConfig` from the store at request time. Concrete change (find the existing `_mapRefiners` method):

```ts
private _mapRefiners(
  rawResponse: Array<{ Name: string; Entries: Array<{ RefinementName: string; RefinementValue: string; RefinementCount: number }> }>
): IRefiner[] {
  const filterConfig: IFilterConfig[] = this._store?.getState().filterConfig || [];
  return mapRefinersWithPreprocessing(rawResponse, filterConfig);
}
```

(`this._store` is the provider's reference to the Zustand store; verify the exact field name when implementing — search for `this._store` or `getState().filterConfig` in the file.)

- [ ] **Step 5: Run test to verify it passes**

Run: `npx heft test --test-path-pattern mapRefinersPreprocessing`
Expected: PASS, 7 tests.

- [ ] **Step 6: Run full test suite**

Run: `npx heft build && npx heft test --silent`
Expected: Build green; all tests pass (454 → 461).

- [ ] **Step 7: Commit**

```bash
git add tests/providers/mapRefinersPreprocessing.test.ts \
        src/libraries/spSearchStore/providers/SharePointSearchProvider.ts
git commit -m "$(cat <<'EOF'
feat(filters): type-aware preprocessing for refiner values

Strips 'string;#' / 'int;#' / 'datetime;#' prefixes when filter dataType
is choiceMulti/lookup/calculated/auto (heuristic), preserves original raw
value for the KQL clause. Splits values on configurable delimiter for
text columns that store comma/newline-separated multi-values, aggregating
counts per token.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 10: Property pane "Data format" section

**Files:**
- Modify: `src/propertyPaneControls/filtersCollection/FiltersCollectionControl.tsx` (add new Section after "Toggle labels")
- Modify: `src/propertyPaneControls/filtersCollection/fieldRelevance.ts` (add `dataType` and `valueSplitDelimiter` to relevance map for checkbox/tagbox/dropdown/text)

- [ ] **Step 1: Add the property pane UI**

In `src/propertyPaneControls/filtersCollection/FiltersCollectionControl.tsx`, add a new Section (placement: after the "Toggle labels" section, before the closing right-pane wrapper). Use the existing Section / control pattern in that file:

```tsx
{isFieldRelevant(draft.filterType, 'dataType') && (
  <Section title="Data format">
    <Dropdown
      label="Underlying data type"
      selectedKey={draft.dataType || 'auto'}
      options={[
        { key: 'auto', text: 'Auto-detect (recommended)' },
        { key: 'text', text: 'Text — no preprocessing' },
        { key: 'choiceMulti', text: 'Choice (multi-value)' },
        { key: 'lookup', text: 'Lookup' },
        { key: 'calculated', text: 'Calculated column' },
        { key: 'datetime', text: 'Date / time' },
        { key: 'yesno', text: 'Yes / no' },
        { key: 'number', text: 'Number' },
      ]}
      onChange={(_e, option): void => updateDraftField('dataType', option?.key)}
    />
    <TextField
      label="Split values on (optional)"
      description="Empty = no split. Examples: comma (,) semicolon (;) newline (\n) pipe (|)"
      value={draft.valueSplitDelimiter || ''}
      onChange={(_e, v): void => updateDraftField('valueSplitDelimiter', v || undefined)}
    />
  </Section>
)}
```

- [ ] **Step 2: Update fieldRelevance.ts**

In `src/propertyPaneControls/filtersCollection/fieldRelevance.ts`, find the relevance map. Add `dataType` and `valueSplitDelimiter` to the field lists for `'checkbox'`, `'tagbox'`, `'dropdown'`, `'text'`:

```ts
// In each of those filterType entries, append 'dataType' and 'valueSplitDelimiter':
checkbox: [..., 'dataType', 'valueSplitDelimiter'],
tagbox:   [..., 'dataType', 'valueSplitDelimiter'],
dropdown: [..., 'dataType', 'valueSplitDelimiter'],
text:     [..., 'dataType', 'valueSplitDelimiter'],
```

(Verify the exact existing structure when implementing — the map may use a different shape.)

- [ ] **Step 3: Run the build**

Run: `npx heft build && npx heft test --silent`
Expected: Green; tests pass.

- [ ] **Step 4: Manual verification**

Run: `npm start`. Open the workbench, edit a Filters web part. In Configure Filters → Add refiner → set filterType to `checkbox`. Verify the new "Data format" section appears with the two controls. Change filterType to `toggle` — section should disappear (toggle isn't in the relevance list).

- [ ] **Step 5: Commit**

```bash
git add src/propertyPaneControls/filtersCollection/FiltersCollectionControl.tsx \
        src/propertyPaneControls/filtersCollection/fieldRelevance.ts
git commit -m "$(cat <<'EOF'
feat(filters): property pane 'Data format' section for dataType + split

Admin-facing UI for the preprocessing fields landed in IFilterConfig +
_mapRefiners. Visible for checkbox/tagbox/dropdown/text filter types.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 11: Replace TaxonomyTreeFilter with PnP `TaxonomyPicker` (Issue C)

**Files:**
- Create: `src/propertyPaneControls/pnpStyleShims/TaxonomyPicker.module.scss.js` (CONDITIONAL — only if step 1 pre-flight reports the SCSS crash)
- Modify: `gulpfile.js` (CONDITIONAL — webpack alias for the shim)
- Modify: `src/webparts/spSearchFilters/components/TaxonomyTreeFilter.tsx` (replace DevExtreme TreeView with PnP TaxonomyPicker)

- [ ] **Step 1: Pre-flight check — does PnP TaxonomyPicker have the same SCSS-import crash?**

Inspect the PnP package on disk:

```bash
cat node_modules/@pnp/spfx-controls-react/lib/controls/taxonomyPicker/TaxonomyPicker.js | head -30
ls node_modules/@pnp/spfx-controls-react/lib/controls/taxonomyPicker/
```

Look for `import styles from './TaxonomyPicker.module.scss'` near the top. Check whether `TaxonomyPicker.module.scss.css` exists alongside it but NO corresponding JS shim — that's the crash signature (same as CollectionDataViewer and PeoplePickerComponent).

**If the JS shim is missing** (high chance based on what we saw with the other two PnP controls), proceed with the alias-shim approach in step 2. **If it's present**, skip to step 3.

- [ ] **Step 2: (Conditional) Create the SCSS shim + webpack alias**

Create `src/propertyPaneControls/pnpStyleShims/TaxonomyPicker.module.scss.js`:

```js
// Shim for @pnp/spfx-controls-react/lib/controls/taxonomyPicker/TaxonomyPicker.module.scss
// SPFx 1.22's sp-css-loader can't process the pre-compiled .module.css PnP ships,
// so the JS import resolves to undefined and the control crashes on `styles.X`.
// This shim exports the original (hashed) class names PnP compiled into its CSS,
// so DOM class names match selectors in TaxonomyPicker.module.scss.css.
module.exports = {
  // NOTE: confirm hashes against the actual TaxonomyPicker.module.scss.css file —
  // the patterns below are placeholders that the implementer must replace with
  // the real hashed names from `grep -E '^\.' TaxonomyPicker.module.scss.css`.
  taxonomyPicker: 'taxonomyPicker_<hash>',
  termSet: 'termSet_<hash>',
  // ... add every class referenced as `styles.X` in TaxonomyPicker.js
};
```

(The implementer reads the compiled .css to extract the hashed names, like we did with `pnpPropertyControlsFix.ts` for the CollectionData control. Pattern is `\.<className>_<hash>`.)

In `gulpfile.js`, add the alias inside `additionalConfiguration`:

```js
generatedConfiguration.resolve = generatedConfiguration.resolve || {};
generatedConfiguration.resolve.alias = generatedConfiguration.resolve.alias || {};
generatedConfiguration.resolve.alias[
  '@pnp/spfx-controls-react/lib/controls/taxonomyPicker/TaxonomyPicker.module.scss'
] = require('path').resolve(
  __dirname,
  'src/propertyPaneControls/pnpStyleShims/TaxonomyPicker.module.scss.js'
);
```

- [ ] **Step 3: Rewrite TaxonomyTreeFilter.tsx to use TaxonomyPicker**

Replace `src/webparts/spSearchFilters/components/TaxonomyTreeFilter.tsx` with:

```tsx
import * as React from 'react';
import { TaxonomyPicker, IPickerTerms } from '@pnp/spfx-controls-react/lib/TaxonomyPicker';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import styles from './SpSearchFilters.module.scss';
import type {
  IActiveFilter,
  IFilterConfig,
  IRefinerValue,
} from '@interfaces/index';

export interface ITaxonomyTreeFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
  onReplaceRefinerValues?: (payload: { filterName: string; values: IActiveFilter[] }) => void;
}

function extractGuid(taxonomyToken: string): string | undefined {
  // GP0|#0a1b2c3d-... → 0a1b2c3d-...
  const m = /^GP0\|#([0-9a-f-]+)/i.exec(taxonomyToken);
  return m ? m[1] : undefined;
}

function buildToken(guid: string): string {
  return 'GP0|#' + guid;
}

const TaxonomyTreeFilter: React.FC<ITaxonomyTreeFilterProps> = (props) => {
  const { filterName, config, activeFilters, onToggleRefiner, onReplaceRefinerValues } = props;
  const operator = config?.operator || 'eq';

  if (!config?.termSetId) {
    return (
      <div className={styles.taxonomyTreeError}>
        Taxonomy filter is not configured — set <strong>Term set ID</strong> in the
        property pane to use this control.
      </div>
    );
  }

  const initialTerms: IPickerTerms = React.useMemo(function (): IPickerTerms {
    return activeFilters
      .filter(function (f: IActiveFilter): boolean { return f.filterName === filterName; })
      .map(function (f: IActiveFilter): IPickerTerms[number] {
        const guid = extractGuid(f.value) || '';
        return {
          key: guid,
          name: f.displayValue || guid,
          path: '',
          termSet: config.termSetId || '',
        } as IPickerTerms[number];
      });
  }, [activeFilters, filterName, config.termSetId]);

  function handleChange(terms: IPickerTerms): void {
    if (onReplaceRefinerValues) {
      const batched: IActiveFilter[] = (terms || []).map(function (t): IActiveFilter {
        return {
          filterName,
          value: buildToken(t.key),
          displayValue: t.name,
          operator,
        };
      });
      onReplaceRefinerValues({ filterName, values: batched });
      return;
    }

    // Fallback: per-delta
    const previousGuids = activeFilters
      .filter(function (f): boolean { return f.filterName === filterName; })
      .map(function (f): string { return extractGuid(f.value) || ''; });
    const nextGuids = (terms || []).map(function (t): string { return t.key; });
    for (let i = 0; i < nextGuids.length; i++) {
      if (previousGuids.indexOf(nextGuids[i]) < 0) {
        onToggleRefiner({ filterName, value: buildToken(nextGuids[i]), displayValue: terms[i].name, operator });
      }
    }
    for (let i = 0; i < previousGuids.length; i++) {
      if (nextGuids.indexOf(previousGuids[i]) < 0) {
        onToggleRefiner({ filterName, value: buildToken(previousGuids[i]), operator });
      }
    }
  }

  return (
    <div className={styles.taxonomyTreeContainer}>
      <TaxonomyPicker
        allowMultipleSelections={true}
        termsetNameOrID={config.termSetId}
        panelTitle={config.displayName || 'Pick terms'}
        label=""
        context={SPContext.spfxContext as unknown as ConstructorParameters<typeof TaxonomyPicker>[0]['context']}
        initialValues={initialTerms}
        onChange={handleChange}
        isTermSetSelectable={false}
      />
    </div>
  );
};

export default TaxonomyTreeFilter;
```

- [ ] **Step 4: Run the build**

Run: `npx heft build`
Expected: Green.

- [ ] **Step 5: Manual verification in workbench**

Run: `npm start`. Configure a taxonomy refiner with a known term set ID. Verify:
- The taxonomy filter renders a Fluent-themed picker (label, input, "Browse terms" link)
- Clicking Browse opens a real tree panel
- Selecting two terms persists the selection in the picker AFTER the search re-fires
- Active filter pills show the term names

If the picker crashes with `styles.<X>` undefined, redo step 2 (alias-shim).

- [ ] **Step 6: Run full test suite**

Run: `npx heft test --silent`
Expected: All tests pass. (Unit tests for the new file are minimal because TaxonomyPicker is third-party; manual verification is the primary signal.)

- [ ] **Step 7: Commit**

```bash
git add src/webparts/spSearchFilters/components/TaxonomyTreeFilter.tsx
# If alias shim was needed:
git add src/propertyPaneControls/pnpStyleShims/TaxonomyPicker.module.scss.js gulpfile.js
git commit -m "$(cat <<'EOF'
feat(filters): replace DevExtreme TreeView with PnP TaxonomyPicker

Fluent-themed real term tree with browse + search, replacing the
DevExtreme TreeView that read as 'search box with flat options below'.
Selection state derives synchronously from activeFilters via
initialValues — no more loadTree race wiping the visual selection.

[Conditional commit body if alias shim was needed:]
PnP TaxonomyPicker's TaxonomyPicker.module.scss JS import resolves to
undefined under SPFx 1.22's sp-css-loader (same root cause as the other
PnP controls we replaced/shimmed). Adds a tiny JS shim mapping the
hashed class names and a webpack alias to make `styles.X` resolve at
runtime.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### PR2 milestone

Push and verify:

```bash
git push origin main
```

In production:
- `string;#` prefix is gone from refiner values for Choice-multi columns ✅
- Text refiner with `valueSplitDelimiter: ','` splits comma-separated values ✅
- Taxonomy renders Fluent-themed real tree, selection persists ✅

---

## Self-review checklist (run before declaring done)

### Spec coverage

| Spec item | Task |
|---|---|
| Issue E — Fluent People picker | Task 5 |
| Issue A — Multi-toggle clobber | Tasks 1-5 (foundation + 4 component migrations) |
| Issue G — Selection persistence | Task 4 (taxonomy useEffect dep fix) + Task 11 (replacement uses initialValues derivation) |
| Issue D — Toggle defaultValue | Task 6 |
| Issue H — DevExtreme native TagBox | Task 7 |
| Issue B — Data type awareness | Tasks 8, 9, 10 |
| Issue F — Delimited value splits | Tasks 8, 9, 10 |
| Issue C — Taxonomy UI | Task 11 |

All 8 spec items covered.

### Test design

Each task ships a fast Jest unit test against pure helpers (the `applyReplaceRefinerValues` / `seedToggleDefaults` / `mapRefinersWithPreprocessing` exports) so component integration tests don't have to spin up heavy controls. Component-level tests use the existing `@testing-library/react` setup with `__test_*` hooks for non-DOM event sources. The manual verification gates each PR milestone.

### Type consistency

- `applyReplaceRefinerValues(current: IActiveFilter[], filterName: string, values: IActiveFilter[]): IActiveFilter[]` — used identically in Task 1 (definition), Task 2-5 (callers reference via `onReplaceRefinerValues` callback shape).
- `onReplaceRefinerValues: (payload: { filterName: string; values: IActiveFilter[] }) => void` — same shape on every component's prop interface.
- `seedToggleDefaults(current: IActiveFilter[], configs: IFilterConfig[]): IActiveFilter[]` — defined and called from one location.
- `mapRefinersWithPreprocessing(refinerResponses, filterConfig: IFilterConfig[]): IRefiner[]` — `filterConfig` is the same name as the existing Zustand store slice key (verified).

### Placeholder scan

- "verify exact existing structure when implementing" appears in Task 10 step 2 — fieldRelevance map shape needs runtime verification. **Resolution:** the engineer can grep one line: `grep -n 'checkbox' src/propertyPaneControls/filtersCollection/fieldRelevance.ts` to see the existing shape. Not a TBD — a small verification step.
- Task 11 step 2 has `taxonomyPicker_<hash>` placeholders — these are intentional: the implementer extracts real values from the compiled `.css` file at implementation time, exactly like `pnpPropertyControlsFix.ts` did before being removed. The pattern + the source-of-truth file are both named.

No other placeholders. No "TODO". No vague "handle edge cases".

### Scope

11 tasks across 2 PRs is appropriately scoped for one design. Each task is independently committable; no task leaves the build broken.
