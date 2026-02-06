import { type StoreApi } from 'zustand/vanilla';
import { type ISearchStore } from '@interfaces/index';
import { type IReadonlyTheme } from '@microsoft/sp-component-base';

export interface ISpSearchVerticalsProps {
  store: StoreApi<ISearchStore>;
  showCounts: boolean;
  hideEmptyVerticals: boolean;
  tabStyle: 'tabs' | 'pills' | 'underline';
  theme: IReadonlyTheme | undefined;
}
