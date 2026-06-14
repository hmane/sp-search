import * as React from 'react';
import { TaxonomyPicker, IPickerTerms } from '@pnp/spfx-controls-react/lib/TaxonomyPicker';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
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

/** Refinement tokens for taxonomy values are emitted as `GP0|#<guid>`. */
function extractGuid(taxonomyToken: string): string | undefined {
  const match = /^GP0\|#([0-9a-fA-F-]+)/.exec(taxonomyToken);
  return match ? match[1] : undefined;
}

function buildToken(guid: string): string {
  return 'GP0|#' + guid;
}

/**
 * TaxonomyTreeFilter — renders a Fluent-themed PnP TaxonomyPicker.
 *
 * Replaces the previous DevExtreme TreeView implementation (which read as
 * "search box with flat options below"). TaxonomyPicker handles all term-store
 * interaction internally — no more loadTree race wiping the visible selection
 * (Issue G) and no more buildToken/labelMap/tokenMap bookkeeping.
 *
 * Selection state derives synchronously from `activeFilters` via
 * `initialValues` so the picker stays in sync after every search re-fire.
 */
const TaxonomyTreeFilter: React.FC<ITaxonomyTreeFilterProps> = (
  props: ITaxonomyTreeFilterProps
): React.ReactElement => {
  const { filterName, config, activeFilters, onToggleRefiner, onReplaceRefinerValues } = props;
  const operator: 'AND' | 'OR' = config && config.operator === 'AND' ? 'AND' : 'OR';

  if (!config || !config.termSetId) {
    return (
      <div className={styles.taxonomyTreeError}>
        Taxonomy filter is not configured — set <strong>Term set ID</strong> in the
        property pane to use this control.
      </div>
    );
  }

  const termSetId: string = config.termSetId;

  const initialTerms: IPickerTerms = React.useMemo(function (): IPickerTerms {
    const seeded: IPickerTerms = [] as IPickerTerms;
    for (let i = 0; i < activeFilters.length; i++) {
      const filter = activeFilters[i];
      if (filter.filterName !== filterName) {
        continue;
      }
      const guid = extractGuid(filter.value) || filter.value;
      seeded.push({
        key: guid,
        name: filter.displayValue || guid,
        path: '',
        termSet: termSetId,
      });
    }
    return seeded;
  }, [activeFilters, filterName, termSetId]);

  function handleChange(terms?: IPickerTerms): void {
    const next: IPickerTerms = terms || ([] as IPickerTerms);

    if (onReplaceRefinerValues) {
      const batched: IActiveFilter[] = [];
      for (let i = 0; i < next.length; i++) {
        batched.push({
          filterName,
          value: buildToken(next[i].key),
          displayValue: next[i].name,
          operator,
        });
      }
      onReplaceRefinerValues({ filterName, values: batched });
      return;
    }

    // Fallback path (no batched callback wired): per-delta toggles.
    const previousGuids: string[] = [];
    for (let i = 0; i < activeFilters.length; i++) {
      if (activeFilters[i].filterName === filterName) {
        previousGuids.push(extractGuid(activeFilters[i].value) || activeFilters[i].value);
      }
    }
    const nextGuids: string[] = next.map(function (t): string { return t.key; });

    for (let i = 0; i < nextGuids.length; i++) {
      if (previousGuids.indexOf(nextGuids[i]) < 0) {
        onToggleRefiner({
          filterName,
          value: buildToken(nextGuids[i]),
          displayValue: next[i].name,
          operator,
        });
      }
    }
    for (let i = 0; i < previousGuids.length; i++) {
      if (nextGuids.indexOf(previousGuids[i]) < 0) {
        onToggleRefiner({
          filterName,
          value: buildToken(previousGuids[i]),
          operator,
        });
      }
    }
  }

  return (
    <div className={styles.taxonomyTreeContainer}>
      <TaxonomyPicker
        allowMultipleSelections={config.multiValues !== false}
        termsetNameOrID={termSetId}
        panelTitle={config.displayName || 'Pick terms'}
        label=""
        context={SPContext.spfxContext as never}
        initialValues={initialTerms}
        onChange={handleChange}
        isTermSetSelectable={false}
      />
    </div>
  );
};

export default TaxonomyTreeFilter;
