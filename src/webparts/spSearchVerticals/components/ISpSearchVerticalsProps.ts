import { type StoreApi } from 'zustand/vanilla';
import { type ISearchStore } from '@interfaces/index';
import { type IReadonlyTheme } from '@microsoft/sp-component-base';

export interface ISpSearchVerticalsProps {
  store: StoreApi<ISearchStore>;
  showCounts: boolean;
  hideEmptyVerticals: boolean;
  tabStyle: 'tabs' | 'pills' | 'underline';
  theme: IReadonlyTheme | undefined;
  // T3.D7 — edit-mode-only validation MessageBar at component root.
  // Renders when any vertical references a `dataProviderId` not in
  // the registered providers list (typo Did-You-Mean for the silent-
  // fallback failure mode `SearchOrchestrator.ts` produces today).
  isEditMode?: boolean;
}
