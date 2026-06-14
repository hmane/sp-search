import * as React from 'react';
import { TreeView } from 'devextreme-react/tree-view';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import 'spfx-toolkit/lib/utilities/context/pnpImports/taxonomy';
import styles from './SpSearchFilters.module.scss';
import type { IActiveFilter, IFilterConfig, IRefinerValue, IReplaceRefinerValuesPayload } from '@interfaces/index';

/** Shape of a PnP taxonomy term tree node returned by getAllChildrenAsOrderedTree */
interface IPnPTermTreeNode {
  id: string;
  defaultLabel?: string;
  labels?: Array<{ name: string; isDefault: boolean }>;
  children?: IPnPTermTreeNode[];
}

/** Typed shape for PnP taxonomy term store API accessed via SPContext.sp.termStore */
interface IPnPTermStoreApi {
  getTermById(id: string): () => Promise<{ set?: { id?: string } } | undefined>;
  sets: {
    getById(id: string): {
      getAllChildrenAsOrderedTree(): Promise<IPnPTermTreeNode[]>;
    };
  };
}

interface ITaxonomyTreeItem {
  id: string;
  text: string;
  token: string;
  items?: ITaxonomyTreeItem[];
}

export interface ITaxonomyTreeFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
  /**
   * Batched callback (Task 1 foundation). Components migrating in Tasks 2-5
   * will switch from per-delta `onToggleRefiner` to a single batched call here.
   */
  onReplaceRefinerValues?: (payload: IReplaceRefinerValuesPayload) => void;
}

function extractGuid(value: string): string | undefined {
  const match = value.match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/);
  return match ? match[0] : undefined;
}

function getDefaultLabel(labels: Array<{ name: string; isDefault: boolean }> | undefined): string {
  if (!labels || labels.length === 0) {
    return '';
  }
  const found = labels.find((label) => label.isDefault);
  return found ? found.name : labels[0].name;
}

function buildToken(termId: string): string {
  return 'GP0|#' + termId;
}

function buildCountMap(values: IRefinerValue[]): Map<string, number> {
  const map = new Map<string, number>();
  for (let i = 0; i < values.length; i++) {
    const guid = extractGuid(values[i].value);
    if (guid) {
      map.set(guid, values[i].count);
    }
  }
  return map;
}

function mapTermTree(
  term: IPnPTermTreeNode,
  countMap: Map<string, number>,
  showCount: boolean
): ITaxonomyTreeItem {
  const label = term.defaultLabel || getDefaultLabel(term.labels) || term.id;
  const count = countMap.get(term.id);
  const text = showCount && typeof count === 'number'
    ? label + ' (' + String(count) + ')'
    : label;

  const children = term.children && term.children.length
    ? term.children.map((child: IPnPTermTreeNode) => mapTermTree(child, countMap, showCount))
    : undefined;

  return {
    id: term.id,
    text,
    token: buildToken(term.id),
    items: children,
  };
}

function buildTokenMap(items: ITaxonomyTreeItem[] | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (!items) {
    return map;
  }
  const stack: ITaxonomyTreeItem[] = items.slice();
  while (stack.length > 0) {
    const item = stack.pop() as ITaxonomyTreeItem;
    map.set(item.id, item.token);
    if (item.items && item.items.length > 0) {
      for (let i = 0; i < item.items.length; i++) {
        stack.push(item.items[i]);
      }
    }
  }
  return map;
}

function buildLabelMap(items: ITaxonomyTreeItem[] | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (!items) {
    return map;
  }
  const stack: ITaxonomyTreeItem[] = items.slice();
  while (stack.length > 0) {
    const item = stack.pop() as ITaxonomyTreeItem;
    // Strip count suffix like " (5)" from text to get clean label
    const text = item.text.replace(/\s*\(\d+\)$/, '');
    map.set(item.id, text);
    if (item.items && item.items.length > 0) {
      for (let i = 0; i < item.items.length; i++) {
        stack.push(item.items[i]);
      }
    }
  }
  return map;
}

function buildFallbackItems(values: IRefinerValue[], showCount: boolean): ITaxonomyTreeItem[] {
  return values.map((value) => {
    const guid = extractGuid(value.value) || value.value;
    const label = value.name || value.value;
    const text = showCount ? label + ' (' + String(value.count) + ')' : label;
    return {
      id: guid,
      text,
      token: value.value || buildToken(guid),
    };
  });
}

/**
 * Pure helper — builds the batched onReplaceRefinerValues payload for the
 * taxonomy filter from the TreeView's currently-selected node keys.
 * Exported so the unit test can drive the conversion logic without
 * mounting a real DevExtreme TreeView.
 *
 * Falls back to `GP0|#<key>` when the tokenMap has no entry (e.g. lazy
 * selection of a node before its parent subtree mapped a token).
 */
export function buildTaxonomyBatchPayload(input: {
  filterName: string;
  selectedKeys: string[];
  tokenMap: Map<string, string>;
  labelMap: Map<string, string>;
  operator: 'AND' | 'OR';
}): IReplaceRefinerValuesPayload {
  const { filterName, selectedKeys, tokenMap, labelMap, operator } = input;
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

const TaxonomyTreeFilter: React.FC<ITaxonomyTreeFilterProps> = (props: ITaxonomyTreeFilterProps): React.ReactElement => {
  const { filterName, values, config, activeFilters, onToggleRefiner, onReplaceRefinerValues } = props;

  const showCount: boolean = config ? config.showCount : true;
  const operator: 'AND' | 'OR' = config ? config.operator : 'OR';
  const includeChildren = config?.includeChildren !== false;
  const selectionMode = config?.multiValues === false ? 'single' : 'multiple';

  const [treeItems, setTreeItems] = React.useState<ITaxonomyTreeItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

  const countMap = React.useMemo(() => buildCountMap(values), [values]);

  // Tracks whether the term-store fetch has already populated the tree
  // for the current term-set scope. Persists across renders so the
  // loadTree effect can early-out once it has hydrated successfully,
  // independent of how often `values` changes per search response.
  const termTreeHydratedRef = React.useRef<string | undefined>(undefined);

  // Probe to detect the first transition from empty -> non-empty refiner
  // buckets. Used as a dep so the auto-derive-termSetId path can fire
  // when the first search response arrives, without re-running on
  // subsequent responses (which would wipe the user's selection — Issue G).
  const hasValues: boolean = values.length > 0;

  React.useEffect(() => {
    let cancelled = false;

    // Issue G: re-fetching the term tree on every search response
    // remounts the DevExtreme TreeView and wipes the visible selection.
    // Skip if we've already hydrated for this term-set scope.
    const scopeKey: string = config?.termSetId || (hasValues ? 'auto' : 'pending');
    if (termTreeHydratedRef.current === scopeKey) {
      return;
    }

    async function loadTree(): Promise<void> {
      setIsLoading(true);
      setErrorMessage(undefined);

      try {
        let termSetId = config?.termSetId;
        if (!termSetId) {
          const firstGuid = values.length > 0 ? extractGuid(values[0].value) : undefined;
          if (firstGuid) {
            const termStoreApi = (SPContext.sp as unknown as { termStore: IPnPTermStoreApi }).termStore;
            const term = await termStoreApi.getTermById(firstGuid)();
            termSetId = term?.set?.id;
          }
        }

        if (termSetId) {
          const termStoreApi = (SPContext.sp as unknown as { termStore: IPnPTermStoreApi }).termStore;
          const tree = await termStoreApi.sets.getById(termSetId).getAllChildrenAsOrderedTree();
          const mapped = tree.map((term: IPnPTermTreeNode) => mapTermTree(term, countMap, showCount));
          if (!cancelled) {
            setTreeItems(mapped);
            termTreeHydratedRef.current = scopeKey;
          }
        } else {
          if (!cancelled) {
            setTreeItems(buildFallbackItems(values, showCount));
            // Don't mark hydrated — we still want to retry once refiner
            // values arrive so we can derive the term set ID.
          }
        }
      } catch (error) {
        SPContext.logger.warn('TaxonomyTreeFilter: failed to load term tree', { error });
        if (!cancelled) {
          setTreeItems(buildFallbackItems(values, showCount));
          setErrorMessage('Failed to load taxonomy terms. Showing available refiners instead.');
        }
      } finally {
        if (!cancelled) {
          setIsLoading(false);
        }
      }
    }

    loadTree().catch(() => { /* fire-and-forget */ });

    return () => {
      cancelled = true;
    };
    // Intentionally NOT depending on `values`, `countMap`, or `showCount`.
    // Re-running this effect on every search response would rebuild the
    // DevExtreme TreeView from scratch and wipe the user's visible
    // selection (Issue G). The tree shape only changes when the term set
    // ID changes; counts in labels go stale between term-set switches,
    // which is the accepted trade-off until PR2/Task 11 replaces this
    // component with PnP TaxonomyPicker.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [config?.termSetId, hasValues]);

  // Reset hydration scope when the configured term set ID changes so a
  // fresh load fires for the new scope.
  React.useEffect(function (): void {
    termTreeHydratedRef.current = undefined;
  }, [config?.termSetId]);

  const tokenMap = React.useMemo(() => buildTokenMap(treeItems), [treeItems]);
  const labelMap = React.useMemo(() => buildLabelMap(treeItems), [treeItems]);

  const selectedKeys = React.useMemo(() => {
    const selected: string[] = [];
    for (let i = 0; i < activeFilters.length; i++) {
      const filter = activeFilters[i];
      if (filter.filterName !== filterName) {
        continue;
      }
      const guid = extractGuid(filter.value);
      selected.push(guid || filter.value);
    }
    return selected;
  }, [activeFilters, filterName]);

  function handleSelectionChanged(e: { component?: { getSelectedNodeKeys(): string[] } }): void {
    if (!e || !e.component) {
      return;
    }
    const keys: string[] = e.component.getSelectedNodeKeys();

    // Issue A: emit the full intended selection in one batched call so the
    // parent doesn't process N per-delta toggles against stale closures.
    if (onReplaceRefinerValues) {
      onReplaceRefinerValues(
        buildTaxonomyBatchPayload({
          filterName,
          selectedKeys: keys,
          tokenMap,
          labelMap,
          operator,
        })
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
        onToggleRefiner({
          filterName,
          value: currentValues[i],
          operator,
        });
      }
    }
  }

  return (
    <div className={styles.taxonomyTreeContainer}>
      {isLoading && (
        <div className={styles.taxonomyTreeLoading}>Loading taxonomy...</div>
      )}
      {errorMessage && (
        <div className={styles.taxonomyTreeError}>{errorMessage}</div>
      )}
      <TreeView
        items={treeItems}
        keyExpr="id"
        displayExpr="text"
        searchEnabled={true}
        searchMode="contains"
        showCheckBoxesMode="normal"
        selectNodesRecursive={includeChildren}
        selectionMode={selectionMode}
        selectByClick={true}
        selectedItemKeys={selectedKeys}
        onItemSelectionChanged={handleSelectionChanged}
      />
    </div>
  );
};

export default TaxonomyTreeFilter;
