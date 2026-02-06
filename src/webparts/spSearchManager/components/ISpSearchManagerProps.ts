import { StoreApi } from 'zustand/vanilla';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ISearchStore } from '@interfaces/index';
import { SearchManagerService } from '@services/index';

export interface ISpSearchManagerProps {
  store: StoreApi<ISearchStore>;
  service: SearchManagerService;
  theme: IReadonlyTheme | undefined;
  mode: 'standalone' | 'panel';
}
