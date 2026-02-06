import { StoreApi } from 'zustand/vanilla';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ISearchStore } from '@interfaces/index';
import { SearchOrchestrator } from '@orchestrator/SearchOrchestrator';

export interface ISpSearchResultsProps {
  store: StoreApi<ISearchStore>;
  orchestrator: SearchOrchestrator | undefined;
  searchContextId: string;
  theme: IReadonlyTheme | undefined;
  showResultCount: boolean;
  showSortDropdown: boolean;
  enableSelection: boolean;
  defaultLayout: string;
  pageSize: number;
}
