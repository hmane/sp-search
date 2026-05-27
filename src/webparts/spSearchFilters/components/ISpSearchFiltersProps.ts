import type { StoreApi } from 'zustand/vanilla';
import type { ISearchStore } from '@interfaces/index';

export interface ISpSearchFiltersProps {
  store: StoreApi<ISearchStore> | undefined;
  applyMode: 'instant' | 'manual';
  showClearAll: boolean;
  enableVisualFilterBuilder: boolean;
  // T4.D5 — edit-mode-only validation MessageBar at component root. Renders
  // a Did-You-Mean-style warning when any filter row references a managed
  // property typo against the cached schema, and flags malformed refinement
  // filter rows (missing operator, range-without-comma, unsupported op).
  isEditMode?: boolean;
  // T4.D12 — current search context ID, used to look up cross-web-part
  // preset suggestions in the registry.
  searchContextId?: string;
  // T4.D12 — apply the suggested filter rows to the web part's
  // `filtersCollection` property. Called from the preset-suggestion
  // MessageBar's "Apply" button. The web part handles the property write
  // + property pane re-render.
  onApplyPresetFilters?: (filterRows: Array<{
    managedProperty: string;
    label: string;
    urlAlias?: string;
    filterType: string;
  }>) => void;
}
