import { StoreApi } from 'zustand/vanilla';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ISearchStore } from '@interfaces/index';
import { SearchOrchestrator } from '@orchestrator/SearchOrchestrator';
import { GraphOrgService } from './GraphOrgService';
import { TitleDisplayMode } from './documentTitleUtils';
import type { IResultLinkConfig } from './resultLink';
import type { IColumnConfigItem } from './ColumnConfigField/columnConfig';

export interface ISelectedPropertyColumn {
  property: string;
  alias: string;
}

export interface ISpSearchResultsProps {
  store: StoreApi<ISearchStore>;
  orchestrator: SearchOrchestrator | undefined;
  searchContextId: string;
  /** Absolute URL of the current SharePoint site — used to build ISearchContext for action providers. */
  siteUrl: string;
  theme: IReadonlyTheme | undefined;
  showResultCount: boolean;
  showSortDropdown: boolean;
  showDeleteConfirmation: boolean;
  enablePreviewPanel: boolean;
  hideWebPartWhenNoResults: boolean;
  titleDisplayMode: TitleDisplayMode;
  defaultLayout: string;
  pageSize: number;
  /** True when the page is in SharePoint edit mode. Enables admin diagnostic notices. */
  isEditMode: boolean;
  /** Admin-configured retrievable/display properties from selectedPropertiesCollection. */
  selectedPropertyColumns: ISelectedPropertyColumn[];
  /**
   * Data Grid metadata columns. Title remains fixed and is not part of this list.
   *
   * Stream B / Phase 1: the richer `IColumnConfigItem` shape carries
   * alias / width / visibility / renderer / maxLength / seeMoreLink /
   * multiValueSeparator. `renderer: ''` is the migration sentinel — it routes
   * through today's auto-detect path so pre-Phase-1 pages render identically.
   */
  gridPropertyColumns: IColumnConfigItem[];
  /** Compact view metadata columns. Title remains fixed; these control the additional compact fields. */
  compactPropertyColumns: ISelectedPropertyColumn[];
  /** KQL query template from the property pane — used for edit-mode validation only. */
  queryTemplate: string;
  /** Graph org service for manager/direct-reports lookups in People layout. Undefined when Graph is unavailable. */
  graphOrgService?: GraphOrgService;
  /** Result link behaviour config (Stream C / #7). Drives `resolveResultLink` + DocumentTitleHoverCard click flow. */
  linkConfig: IResultLinkConfig;
}
