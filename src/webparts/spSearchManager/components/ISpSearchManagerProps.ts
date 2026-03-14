import { StoreApi } from 'zustand/vanilla';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISearchStore } from '@interfaces/index';
import { SearchManagerService } from '@services/index';
import type { ICoverageProfile } from '@services/SearchCoverageService';

export interface ISpSearchManagerProps {
  store: StoreApi<ISearchStore>;
  service: SearchManagerService;
  theme: IReadonlyTheme | undefined;
  variant?: 'user' | 'admin';
  searchContextId?: string;
  mode: 'standalone' | 'panel';
  defaultTab?: 'saved' | 'history' | 'collections' | 'coverage' | 'health' | 'insights';
  headerTitle?: string;
  hideHeader?: boolean;
  /** Optional — when omitted, SPContext.context.context is used as fallback */
  context?: WebPartContext;
  enableSavedSearches?: boolean;
  enableSharedSearches?: boolean;
  enableCollections?: boolean;
  enableHistory?: boolean;
  enableCoverage?: boolean;
  coverageSourcePageUrl?: string;
  coverageProfiles?: ICoverageProfile[];
  enableHealth?: boolean;
  enableInsights?: boolean;
  enableAnnotations?: boolean;
  maxHistoryItems?: number;
  showResetAction?: boolean;
  showSaveAction?: boolean;
  onRequestClose?: () => void;
}
