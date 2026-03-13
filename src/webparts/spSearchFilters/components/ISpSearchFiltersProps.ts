import type { StoreApi } from 'zustand/vanilla';
import type { ISearchStore } from '@interfaces/index';

export interface ISpSearchFiltersProps {
  store: StoreApi<ISearchStore> | undefined;
  applyMode: 'instant' | 'manual';
  showClearAll: boolean;
  enableVisualFilterBuilder: boolean;
}
