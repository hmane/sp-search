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
  defaultTab?: 'saved' | 'history' | 'collections' | 'health' | 'insights' | 'dashboard';
  headerTitle?: string;
  hideHeader?: boolean;
  /** Optional — when omitted, SPContext.context.context is used as fallback */
  context?: WebPartContext;
  enableSavedSearches?: boolean;
  enableSharedSearches?: boolean;
  enableCollections?: boolean;
  enableHistory?: boolean;
  coverageProfiles?: ICoverageProfile[];
  enableHealth?: boolean;
  enableInsights?: boolean;
  enableAnnotations?: boolean;
  enableDashboard?: boolean;
  expectedSiteUrls?: string[];
  maxHistoryItems?: number;
  showResetAction?: boolean;
  showSaveAction?: boolean;
  onRequestClose?: () => void;
  // T4.D5 — edit-mode-only validation MessageBar at component root. Default
  // false (production / view mode); set true from the web part `render()`
  // via `this.displayMode === DisplayMode.Edit`.
  isEditMode?: boolean;
  // Tenant root URL — drives "different tenant" detection in the validators.
  // Optional: when omitted the validators fall back to URL-shape-only checks.
  tenantRoot?: string;
}
