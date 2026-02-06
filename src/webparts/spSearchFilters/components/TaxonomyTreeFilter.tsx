import * as React from 'react';
import { TreeView } from 'devextreme-react/tree-view';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import 'spfx-toolkit/lib/utilities/context/pnpImports/taxonomy';
import styles from './SpSearchFilters.module.scss';
import type { IActiveFilter, IFilterConfig, IRefinerValue } from '@interfaces/index';

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
  term: any,
  countMap: Map<string, number>,
  showCount: boolean
): ITaxonomyTreeItem {
  const label = term.defaultLabel || getDefaultLabel(term.labels) || term.id;
  const count = countMap.get(term.id);
  const text = showCount && typeof count === 'number'
    ? label + ' (' + String(count) + ')'
    : label;

  const children = term.children && term.children.length
    ? term.children.map((child: any) => mapTermTree(child, countMap, showCount))
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

const TaxonomyTreeFilter: React.FC<ITaxonomyTreeFilterProps> = (props: ITaxonomyTreeFilterProps): React.ReactElement => {
  const { filterName, values, config, activeFilters, onToggleRefiner } = props;

  const showCount: boolean = config ? config.showCount : true;
  const operator: 'AND' | 'OR' = config ? config.operator : 'OR';
  const includeChildren = config?.includeChildren !== false;

  const [treeItems, setTreeItems] = React.useState<ITaxonomyTreeItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

  const countMap = React.useMemo(() => buildCountMap(values), [values]);

  React.useEffect(() => {
    let cancelled = false;

    async function loadTree(): Promise<void> {
      setIsLoading(true);
      setErrorMessage(undefined);

      try {
        let termSetId = config?.termSetId;
        if (!termSetId) {
          const firstGuid = values.length > 0 ? extractGuid(values[0].value) : undefined;
          if (firstGuid) {
            const term = await (SPContext.sp as any).termStore.getTermById(firstGuid)();
            termSetId = term?.set?.id;
          }
        }

        if (termSetId) {
          const tree = await (SPContext.sp as any).termStore.sets.getById(termSetId).getAllChildrenAsOrderedTree();
          const mapped = tree.map((term: any) => mapTermTree(term, countMap, showCount));
          if (!cancelled) {
            setTreeItems(mapped);
          }
        } else {
          if (!cancelled) {
            setTreeItems(buildFallbackItems(values, showCount));
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

    void loadTree();

    return () => {
      cancelled = true;
    };
  }, [values, config?.termSetId, showCount, countMap]);

  const tokenMap = React.useMemo(() => buildTokenMap(treeItems), [treeItems]);

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

  function handleSelectionChanged(e: any): void {
    if (!e || !e.component) {
      return;
    }
    const keys: string[] = e.component.getSelectedNodeKeys();
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
        selectionMode="multiple"
        selectByClick={true}
        selectedItemKeys={selectedKeys}
        onItemSelectionChanged={handleSelectionChanged}
      />
    </div>
  );
};

export default TaxonomyTreeFilter;
