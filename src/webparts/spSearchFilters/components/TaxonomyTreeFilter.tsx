import * as React from 'react';
import { TagBox } from 'devextreme-react/tag-box';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import 'spfx-toolkit/lib/utilities/context/pnpImports/taxonomy';
import styles from './SpSearchFilters.module.scss';
import type {
  IActiveFilter,
  IFilterConfig,
  IRefinerValue,
  IReplaceRefinerValuesPayload,
} from '@interfaces/index';

export interface ITaxonomyTreeFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
  /**
   * Batched callback — emits the full intended selection in one payload.
   * Replaces per-delta `onToggleRefiner` so the parent doesn't process N
   * toggles against stale closures (Issue A).
   */
  onReplaceRefinerValues?: (payload: IReplaceRefinerValuesPayload) => void;
}

/**
 * PnP taxonomy term store API shape used by this filter.
 *
 * - `getTermById(guid)()` resolves a single term (used when no termSetId is
 *   configured and we need to fan out per-GUID).
 * - `sets.getById(setId).terms()` returns ALL terms in a flat set in a
 *   single call (used when termSetId is provided — flat-taxonomy fast path).
 *
 * Matches the API the original DevExtreme TreeView implementation used
 * (commit 13709f1) before Task 11.
 */
interface IPnPTermStoreApi {
  getTermById(id: string): () => Promise<{
    labels?: Array<{ name: string; isDefault: boolean }>;
  } | undefined>;
  sets: {
    getById(id: string): {
      terms: {
        (): Promise<Array<{
          id: string;
          labels?: Array<{ name: string; isDefault: boolean }>;
        }>>;
      };
    };
  };
}

/**
 * Refinement tokens for taxonomy values are emitted as `GP0|#<guid>`.
 * Exported so the test suite (and sibling consumers) can drive the
 * conversion logic without re-implementing the regex.
 */
export function extractGuid(taxonomyToken: string): string | undefined {
  if (!taxonomyToken) {
    return undefined;
  }
  const match: RegExpExecArray | null = /^GP0\|#([0-9a-fA-F-]+)/i.exec(taxonomyToken);
  return match ? match[1] : undefined;
}

function getDefaultLabel(
  labels: Array<{ name: string; isDefault: boolean }> | undefined
): string | undefined {
  if (!labels || labels.length === 0) {
    return undefined;
  }
  const found = labels.find(function (l: { isDefault: boolean }): boolean { return l.isDefault; });
  return found ? found.name : labels[0].name;
}

function areStringArraysEqual(a: string[], b: string[]): boolean {
  if (a.length !== b.length) {
    return false;
  }
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
}

/**
 * Pure helper — builds the batched `onReplaceRefinerValues` payload for
 * the taxonomy TagBox from the editor's full intended selection.
 *
 * Exported so the unit test can drive the conversion logic without
 * mounting a real DevExtreme TagBox. Falls back to the raw token when the
 * labelMap has no entry (e.g. cascade response introduced a new GUID
 * before label resolution finished). Output depends ONLY on the inputs —
 * never on prior activeFilters state — so the parent's per-delta loop
 * can't capture a stale closure (Issue A).
 */
export function buildTaxonomyTagBoxBatchPayload(input: {
  filterName: string;
  selectedTokens: string[];
  labelMap: Map<string, string>;
  operator: 'AND' | 'OR';
}): IReplaceRefinerValuesPayload {
  const values: IActiveFilter[] = [];
  for (let i = 0; i < input.selectedTokens.length; i++) {
    const token: string = input.selectedTokens[i];
    const guid: string | undefined = extractGuid(token);
    const label: string | undefined = guid
      ? input.labelMap.get(guid.toLowerCase())
      : undefined;
    values.push({
      filterName: input.filterName,
      value: token,
      displayValue: label || token,
      operator: input.operator,
    });
  }
  return { filterName: input.filterName, values: values };
}

/**
 * Resolve labels for a flat term set in a single round-trip.
 *
 * Project owner's organization uses flat taxonomies (no hierarchy needed),
 * so `sets.getById(setId).terms()` returns the full label dictionary in
 * one call — far cheaper than per-GUID lookups against the refiner bucket.
 */
async function resolveLabelsForTermSet(termSetId: string): Promise<Map<string, string>> {
  const result: Map<string, string> = new Map<string, string>();
  const termStoreApi = (SPContext.sp as unknown as { termStore: IPnPTermStoreApi }).termStore;
  const terms = await termStoreApi.sets.getById(termSetId).terms();
  for (let i = 0; i < terms.length; i++) {
    const t = terms[i];
    const label: string | undefined = getDefaultLabel(t.labels);
    if (t.id && label) {
      result.set(t.id.toLowerCase(), label);
    }
  }
  return result;
}

/**
 * Resolve labels for an arbitrary list of GUIDs by fanning out per-term
 * lookups in parallel. Used when no termSetId is configured (fallback) and
 * for incremental cascade updates that introduce GUIDs outside the set.
 *
 * Silently skips lookups that fail — the caller falls back to displaying
 * the raw token.
 */
async function resolveLabelsByGuid(guids: string[]): Promise<Map<string, string>> {
  const result: Map<string, string> = new Map<string, string>();
  const unique: string[] = Array.from(new Set(
    guids.map(function (g: string): string { return g.toLowerCase(); })
  ));
  const termStoreApi = (SPContext.sp as unknown as { termStore: IPnPTermStoreApi }).termStore;
  await Promise.all(unique.map(async function (guid: string): Promise<void> {
    try {
      const termInfo = await termStoreApi.getTermById(guid)();
      const label: string | undefined = getDefaultLabel(termInfo ? termInfo.labels : undefined);
      if (label) {
        result.set(guid, label);
      }
    } catch {
      /* Silent — caller falls back to raw token. */
    }
  }));
  return result;
}

/**
 * TaxonomyTreeFilter — flat-taxonomy TagBox with per-term refiner counts.
 *
 * Replaces the PnP TaxonomyPicker (Task 11) which couldn't show per-term
 * counts or narrow by cascade. Project owner's org uses flat taxonomies,
 * so a DevExtreme TagBox keyed on the refiner bucket gives:
 *   - "Electronics (30)" style per-term counts
 *   - Cascade-narrowed options (only terms present in current results)
 *   - Search-as-you-type, multiline chips, OK/Cancel apply
 *
 * Label resolution happens via the PnP term store on mount:
 *   - If `config.termSetId` is set → one round-trip for the full set
 *   - Otherwise → per-GUID lookups for the GUIDs in `values`
 *
 * Loading state: while the initial labelMap is empty AND we have refiner
 * values to label, render a Fluent Spinner. Cascade updates that
 * introduce new GUIDs DON'T gate the first paint — they resolve lazily
 * in the background and fall back to the GUID in the meantime.
 */
const TaxonomyTreeFilter: React.FC<ITaxonomyTreeFilterProps> = (
  props: ITaxonomyTreeFilterProps
): React.ReactElement => {
  const { filterName, values, config, activeFilters, onToggleRefiner, onReplaceRefinerValues } = props;
  const operator: 'AND' | 'OR' = config && config.operator === 'AND' ? 'AND' : 'OR';
  const allowMultiple: boolean = !config || config.multiValues !== false;
  const showCount: boolean = !config || config.showCount !== false;

  const [labelMap, setLabelMap] = React.useState<Map<string, string>>(new Map());
  const [isInitialLoading, setIsInitialLoading] = React.useState<boolean>(true);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

  // Tracks which scope has been hydrated so re-renders (e.g. cascade
  // value changes) don't re-fetch the entire term set.
  const hydratedScopeRef = React.useRef<string | undefined>(undefined);

  const termSetId: string | undefined = config ? config.termSetId : undefined;

  // Initial label resolution. Fires on mount and whenever the configured
  // termSetId changes. We DON'T depend on `values` here — that would
  // re-fetch on every search response and either spam the API or wipe a
  // partially-resolved map. Cascade updates flow through the secondary
  // effect below.
  React.useEffect(function (): () => void {
    let cancelled = false;
    const scopeKey: string = termSetId || 'fallback';

    if (hydratedScopeRef.current === scopeKey) {
      // Already loaded for this scope — nothing to do.
      setIsInitialLoading(false);
      return function (): void { /* no-op cleanup */ };
    }

    setIsInitialLoading(true);
    setErrorMessage(undefined);

    const guidsFromValues: string[] = values
      .map(function (v: IRefinerValue): string | undefined { return extractGuid(v.value); })
      .filter(function (g: string | undefined): boolean { return !!g; }) as string[];

    const work: Promise<Map<string, string>> = termSetId
      ? resolveLabelsForTermSet(termSetId)
      : resolveLabelsByGuid(guidsFromValues);

    work
      .then(function (resolved: Map<string, string>): void {
        if (cancelled) return;
        setLabelMap(resolved);
        hydratedScopeRef.current = scopeKey;
        setIsInitialLoading(false);
      })
      .catch(function (err: Error): void {
        if (cancelled) return;
        SPContext.logger.warn('TaxonomyTreeFilter: failed to load taxonomy labels', { error: err });
        setErrorMessage('Failed to load taxonomy labels — showing GUIDs.');
        // Mark as hydrated so the spinner clears and the user can still
        // operate the filter (tokens render as raw GUIDs).
        hydratedScopeRef.current = scopeKey;
        setIsInitialLoading(false);
      });

    return function (): void { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [termSetId]);

  // Incremental cascade resolution. When a search response brings in new
  // GUIDs not yet in labelMap, resolve them in the background without
  // gating the visible UI. Silently swallows errors — affected tokens
  // fall back to GUID display.
  React.useEffect(function (): void {
    if (isInitialLoading) {
      return;
    }
    const unresolved: string[] = [];
    for (let i = 0; i < values.length; i++) {
      const guid: string | undefined = extractGuid(values[i].value);
      if (guid && !labelMap.has(guid.toLowerCase())) {
        unresolved.push(guid);
      }
    }
    if (unresolved.length === 0) {
      return;
    }
    resolveLabelsByGuid(unresolved)
      .then(function (resolved: Map<string, string>): void {
        if (resolved.size === 0) {
          return;
        }
        setLabelMap(function (prev: Map<string, string>): Map<string, string> {
          const next: Map<string, string> = new Map(prev);
          resolved.forEach(function (label: string, guid: string): void { next.set(guid, label); });
          return next;
        });
      })
      .catch(function (): void { /* silent — fall back to raw GUID */ });
  }, [values, labelMap, isInitialLoading]);

  // Enriched dataSource for the TagBox: each item carries its raw token
  // (used as value) and a resolved display label with optional count.
  const enrichedItems = React.useMemo(function (): Array<{ value: string; displayName: string }> {
    const nextItems = values.map(function (v: IRefinerValue): { value: string; displayName: string } {
      const guid: string | undefined = extractGuid(v.value);
      const resolved: string | undefined = guid ? labelMap.get(guid.toLowerCase()) : undefined;
      const label: string = resolved || v.name || guid || v.value;
      const displayName: string = showCount
        ? label + ' (' + String(v.count) + ')'
        : label;
      return { value: v.value, displayName: displayName };
    });

    const included = new Set<string>();
    for (let i = 0; i < nextItems.length; i++) {
      included.add(nextItems[i].value);
    }

    for (let i = 0; i < activeFilters.length; i++) {
      const active = activeFilters[i];
      if (active.filterName !== filterName || included.has(active.value)) {
        continue;
      }
      nextItems.push({
        value: active.value,
        displayName: active.displayValue || active.value,
      });
      included.add(active.value);
    }

    return nextItems;
  }, [activeFilters, filterName, values, labelMap, showCount]);

  // Selected tokens are the raw refiner tokens already in activeFilters
  // for this filter. The TagBox value array is keyed on these.
  const selectedTokens: string[] = React.useMemo(function (): string[] {
    const out: string[] = [];
    for (let i = 0; i < activeFilters.length; i++) {
      const f = activeFilters[i];
      if (f.filterName === filterName) {
        out.push(f.value);
      }
    }
    return out;
  }, [activeFilters, filterName]);

  // Guard against re-entrant onValueChanged from programmatic value
  // updates after the store commits.
  const isUpdatingRef = React.useRef<boolean>(false);

  function handleValueChanged(e: { value?: string[] }): void {
    if (isUpdatingRef.current) {
      return;
    }
    isUpdatingRef.current = true;

    let nextTokens: string[] = Array.isArray(e.value) ? e.value : [];
    if (!allowMultiple && nextTokens.length > 1) {
      nextTokens = [nextTokens[nextTokens.length - 1]];
    }
    if (areStringArraysEqual(nextTokens, selectedTokens)) {
      isUpdatingRef.current = false;
      return;
    }

    if (onReplaceRefinerValues) {
      onReplaceRefinerValues(buildTaxonomyTagBoxBatchPayload({
        filterName: filterName,
        selectedTokens: nextTokens,
        labelMap: labelMap,
        operator: operator,
      }));
    } else {
      // Per-delta fallback — kept so the component remains usable
      // standalone, but the in-tree FilterGroup always wires
      // onReplaceRefinerValues.
      const previous: string[] = selectedTokens;
      for (let i = 0; i < nextTokens.length; i++) {
        if (previous.indexOf(nextTokens[i]) < 0) {
          const guid: string | undefined = extractGuid(nextTokens[i]);
          const label: string | undefined = guid ? labelMap.get(guid.toLowerCase()) : undefined;
          onToggleRefiner({
            filterName: filterName,
            value: nextTokens[i],
            displayValue: label || nextTokens[i],
            operator: operator,
          });
        }
      }
      for (let i = 0; i < previous.length; i++) {
        if (nextTokens.indexOf(previous[i]) < 0) {
          onToggleRefiner({
            filterName: filterName,
            value: previous[i],
            operator: operator,
          });
        }
      }
    }

    setTimeout(function (): void { isUpdatingRef.current = false; }, 0);
  }

  if (errorMessage && labelMap.size === 0) {
    // Hard failure with no labels at all — show the error inline. The
    // TagBox still renders below (so the user can operate on raw tokens),
    // but the error explains why labels look like GUIDs.
    return (
      <div className={styles.taxonomyTreeContainer}>
        <div className={styles.taxonomyTreeError}>{errorMessage}</div>
        <TagBox
          dataSource={enrichedItems}
          valueExpr="value"
          displayExpr="displayName"
          value={selectedTokens}
          onValueChanged={handleValueChanged}
          searchEnabled={true}
          searchMode="contains"
          showClearButton={true}
          showSelectionControls={true}
          multiline={true}
          applyValueMode="useButtons"
          placeholder="Select terms..."
          maxDisplayedTags={5}
          showMultiTagOnly={false}
        />
      </div>
    );
  }

  if (isInitialLoading && values.length > 0) {
    return (
      <div className={styles.taxonomyTreeContainer}>
        <Spinner size={SpinnerSize.small} label="Loading terms..." labelPosition="right" />
      </div>
    );
  }

  return (
    <div className={styles.taxonomyTreeContainer}>
      <TagBox
        dataSource={enrichedItems}
        valueExpr="value"
        displayExpr="displayName"
        value={selectedTokens}
        onValueChanged={handleValueChanged}
        searchEnabled={true}
        searchMode="contains"
        showClearButton={true}
        showSelectionControls={true}
        multiline={true}
        applyValueMode="useButtons"
        placeholder="Select terms..."
        maxDisplayedTags={5}
        showMultiTagOnly={false}
      />
    </div>
  );
};

export default TaxonomyTreeFilter;
