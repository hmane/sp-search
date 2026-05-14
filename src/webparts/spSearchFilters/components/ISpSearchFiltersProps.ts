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
}
