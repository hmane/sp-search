import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import type { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle,
  PropertyPaneLabel,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { StoreApi } from 'zustand/vanilla';
import { spfxToolkitStylesLoaded } from '../../styles/loadSpfxToolkitStyles';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { configureLegacyPnPBaseUrl } from 'spfx-toolkit/lib/utilities/context/urlSanitizer';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { ensurePnpPropertyControlStyles } from '../../styles/pnpPropertyControlsFix';

import * as strings from 'SpSearchResultsWebPartStrings';
import SpSearchResults from './components/SpSearchResults';
import { ISpSearchResultsProps, ISelectedPropertyColumn } from './components/ISpSearchResultsProps';
import {
  IColumnConfigItem,
  ILegacyColumnItem,
  normalizeColumnConfigItem,
} from './components/ColumnConfigField/columnConfig';
import { PropertyPaneColumnConfigField } from './components/ColumnConfigField/ColumnConfigField';
import { AudienceGate, parseAudienceGroups } from '../../utilities/AudienceGate';
import { SearchContextIdBannerWrapper } from '../../utilities/SearchContextIdMismatchBanner';
import { SPDebugProvider } from 'spfx-toolkit/lib/components/debug';
import { propertyPaneSearchContextIdField } from '../../propertyPaneControls/PropertyPaneSearchContextIdField';
// T4.D11 — context-sensitive help link surface.
import { propertyPaneGroupHelp } from '../../propertyPaneControls/propertyPaneGroupHelp';
// T3.D10 — initialization-order diagnostic; records this web part's
// registration + the filterConfig length at first-search time so the
// Results component can surface an edit-mode warning when Filters
// arrives late.
import { recordWebPartInit, recordFirstSearch } from '@store/utils/initOrderDiagnostic';
import { ISearchStore } from '@interfaces/index';
import {
  getStore,
  getOrchestrator,
  initializeSearchContext,
  incrementContextRef,
  decrementContextRef
} from '@store/store';
import { SearchOrchestrator } from '@orchestrator/SearchOrchestrator';
import { registerBuiltInActions } from './registerBuiltInActions';
import { SharePointSearchProvider, GraphSearchProvider } from '@providers/index';
import { PropertyPaneSchemaHelper } from '../../propertyPaneControls/PropertyPaneSchemaHelper';
import { SCENARIO_PRESETS } from './presets/searchPresets';
// T4.D12 — cross-web-part preset propagation. Results publishes
// filterSuggestions for the Filters web part to consume via an edit-mode
// MessageBar.
import { recordPresetSuggestion } from '@store/utils/presetSuggestionRegistry';
import { GraphOrgService } from './components/GraphOrgService';
import { TitleDisplayMode } from './components/documentTitleUtils';
import { DebugCollector } from '@store/debug';

// Bundle DevExtreme CSS — injected via style-loader at runtime.
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('devextreme/dist/css/dx.common.css');
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('devextreme/dist/css/dx.light.css');
// eslint-disable-next-line @typescript-eslint/no-unused-vars
const _ensureStyles = spfxToolkitStylesLoaded;

export interface ISpSearchResultsWebPartProps {
  searchContextId: string;
  queryTemplate: string;
  selectedProperties: string;
  selectedPropertiesCollection: ISelectedPropertyItem[];
  /**
   * Stream B / Phase 3 — IColumnConfigItem[] replacing the legacy
   * `{ uniqueId, property }` shape (Phase 1 did this for grid; Phase 3
   * brings compact along). Migration is handled at read time by
   * `normalizeColumnConfigItem`, so stored pages don't need to be migrated.
   */
  compactPropertiesCollection: IColumnConfigItem[];
  /**
   * Stream B / Phase 1 — column-config items, replacing the legacy
   * `{ uniqueId, property }` shape. Migration is handled at read time by
   * `normalizeColumnConfigItem`, so stored pages don't need to be migrated.
   */
  gridPropertiesCollection: IColumnConfigItem[];
  /**
   * Stream B / Phase 3 — when true (default), the DataGrid layout shows its
   * built-in column chooser. Columns with `visibility === 'always'` never
   * appear in the chooser; the rest pre-check based on
   * `visibility === 'defaultOn'`.
   */
  showColumnChooser: boolean;
  resultSourceId: string;
  refinementFilters: string;
  refinementFiltersCollection: IRefinementFilterItem[];
  collapseSpecification: string;
  enableQueryRules: boolean;
  trimDuplicates: boolean;
  pageSize: number;
  showPaging: boolean;
  pageRange: number;
  sortablePropertiesCollection: ISortCollectionItem[];
  defaultLayout: string;
  showResultCount: boolean;
  showSortDropdown: boolean;
  showDeleteConfirmation: boolean;
  enablePreviewPanel: boolean;
  hideWebPartWhenNoResults: boolean;
  /**
   * Admin-supplied HTML rendered in place of the default contextual empty
   * state when a search returns zero results. Empty = use the default
   * messaging. Sanitized via spfx-toolkit's `sanitizeHtml` before render.
   */
  emptyResultsMessage: string;
  titleDisplayMode: TitleDisplayMode;
  // Stream C / #7 — result link behaviour. Defaults preserve today's behaviour:
  // 'panel' keeps the existing DocumentTitleHoverCard Modal-for-previewables;
  // 'file' / 'displayForm' keep today's URL resolution.
  resultClickTarget: 'panel' | 'newTab' | 'sameTab' | 'sidePanel';
  documentLinkMode: 'file' | 'propertiesForm';
  listItemLinkMode: 'displayForm' | 'editForm';
  searchScope: string;
  searchScopePath: string;
  showListLayout: boolean;
  showCompactLayout: boolean;
  showGridLayout: boolean;
  showCardLayout: boolean;
  showPeopleLayout: boolean;
  showGalleryLayout: boolean;
  /** Active scenario preset. 'custom' means individual toggles are in control. */
  layoutPreset: string;
  /** Stream D / #10 — comma/newline-separated Azure AD group IDs. Empty = visible to everyone. */
  audienceGroups: string;
}

interface ISortCollectionItem {
  uniqueId: string;
  property: string;
  label: string;
  direction: string;
}

interface ISelectedPropertyItem {
  uniqueId: string;
  property: string;
  alias: string;
}

interface IRefinementFilterItem {
  uniqueId: string;
  property: string;
  operator: string;
  value: string;
}

function normalizeCollectionValue<T>(raw: T[] | string | undefined): T[] {
  if (Array.isArray(raw)) {
    return raw;
  }
  if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw) as T[];
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }
  return [];
}

function isTitleProperty(property: string): boolean {
  const normalized = (property || '').trim().toLowerCase();
  return normalized === 'title' || normalized === 'filename';
}

export default class SpSearchResultsWebPart extends BaseClientSideWebPart<ISpSearchResultsWebPartProps> {

  private _theme: IReadonlyTheme | undefined;
  private _store: StoreApi<ISearchStore> | undefined;
  private _orchestrator: SearchOrchestrator | undefined;
  private _graphOrgService: GraphOrgService | undefined;

  private _getWebAbsoluteUrl(): string {
    return this.context?.pageContext?.web?.absoluteUrl || '';
  }

  private _getSiteAbsoluteUrl(): string {
    return this.context?.pageContext?.site?.absoluteUrl || '';
  }

  public render(): void {
    if (!this._store) {
      return;
    }

    const selectedPropertyColumns: ISelectedPropertyColumn[] = this._getSelectedPropertyColumns();
    const compactPropertyColumns: IColumnConfigItem[] = this._getCompactPropertyColumns();
    const gridPropertyColumns: IColumnConfigItem[] = this._getGridPropertyColumns();
    const contextId: string = this.properties.searchContextId || 'default';
    const element: React.ReactElement<ISpSearchResultsProps> = React.createElement(
      SpSearchResults,
      {
        store: this._store,
        orchestrator: this._orchestrator,
        searchContextId: contextId,
        siteUrl: this._getWebAbsoluteUrl(),
        theme: this._theme,
        showResultCount: this.properties.showResultCount,
        showSortDropdown: this.properties.showSortDropdown,
        showDeleteConfirmation: this.properties.showDeleteConfirmation !== false,
        enablePreviewPanel: this.properties.enablePreviewPanel !== false,
        hideWebPartWhenNoResults: this.properties.hideWebPartWhenNoResults === true,
        emptyResultsMessage: this.properties.emptyResultsMessage || '',
        showColumnChooser: this.properties.showColumnChooser !== false,
        titleDisplayMode: this.properties.titleDisplayMode || 'wrap',
        defaultLayout: this.properties.defaultLayout,
        pageSize: this.properties.pageSize,
        isEditMode: this.displayMode === DisplayMode.Edit,
        selectedPropertyColumns,
        gridPropertyColumns,
        compactPropertyColumns,
        queryTemplate: this.properties.queryTemplate || '{searchTerms}',
        graphOrgService: this._graphOrgService,
        // Stream C / #7 — assemble the link-behaviour config with safe defaults
        // (preserves today's behaviour byte-for-byte for missing properties).
        linkConfig: {
          clickTarget: this.properties.resultClickTarget || 'panel',
          documentLinkMode: this.properties.documentLinkMode || 'file',
          listItemLinkMode: this.properties.listItemLinkMode || 'displayForm',
        }
      }
    );

    // Stream D / #10 — wrap with AudienceGate so the web part hides itself
    // when the current user isn't in any of the configured groups.
    const audienceGroups = parseAudienceGroups(this.properties.audienceGroups);
    const gatedElement: React.ReactElement = React.createElement(
      AudienceGate,
      { audienceGroups, store: this._store },
      element
    );

    // T3.D2 — edit-mode banner warning admins when the contextId doesn't
    // match peer web parts on the same page.
    const bannerWrapped: React.ReactElement = React.createElement(
      SearchContextIdBannerWrapper,
      {
        webPartId: this.instanceId,
        contextId: this.properties.searchContextId || 'default',
        webPartLabel: 'SP Search Results',
        isEditMode: this.displayMode === DisplayMode.Edit,
      },
      gatedElement
    );

    // SPDebug — toolkit's debug runtime + lazy-loaded panel. See SpSearchBox
    // for the per-web-part-state-isolation note. Coexists with this web
    // part's existing DebugFab + DebugPanel (those live inside the React
    // tree below — DebugCollector is window-backed so they share state).
    const wrappedElement: React.ReactElement = React.createElement(
      SPDebugProvider,
      { logger: SPContext.logger, allowInProduction: false },
      bannerWrapped
    );

    ReactDom.render(wrappedElement, this.domElement);
  }

  protected async onInit(): Promise<void> {
    ensurePnpPropertyControlStyles();

    // Initialize SPContext for PnPjs
    // Cast needed: spfx-toolkit uses SPFx 1.21.1 types; this project uses 1.22.2
    await SPContext.basic(this.context as unknown as Parameters<typeof SPContext.basic>[0], 'SPSearchResults');
    // Strip _layouts/15 contamination from the PnP v2 base URL bundled with
    // @pnp/spfx-controls-react + patch global fetch. Both calls are idempotent.
    configureLegacyPnPBaseUrl(this.context);

    const contextId: string = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);
    this._orchestrator = getOrchestrator(contextId);
    // T3.D1 — refcount holder for this web part instance.
    incrementContextRef(contextId);

    if (this._store) {
      // Register the SharePoint Search data provider (idempotent — skips if already registered by another web part)
      const provider = new SharePointSearchProvider();
      const dataProviders = this._store.getState().registries.dataProviders;
      if (!dataProviders.get(provider.id)) {
        dataProviders.register(provider);
      }

      // Register the Graph Search data provider.
      // Requires the Microsoft Graph Sites.Read.All (or equivalent) permission
      // to have been approved in the SharePoint API Access management page.
      // If the client can't be obtained (permission denied, not yet approved),
      // we skip silently — SharePoint Search remains the fallback.
      try {
        const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
        this._graphOrgService = new GraphOrgService(graphClient);
        const graphProvider = new GraphSearchProvider(graphClient);
        if (!dataProviders.get(graphProvider.id)) {
          dataProviders.register(graphProvider);
        }

        // People-specific Graph provider — registered under a separate ID so
        // a "People" vertical can route to it via dataProviderId: 'graph-people'
        const graphPeopleProvider = new GraphSearchProvider(graphClient, {
          id: 'graph-people',
          entityTypes: ['person']
        });
        if (!dataProviders.get(graphPeopleProvider.id)) {
          dataProviders.register(graphPeopleProvider);
        }
      } catch {
        // Graph client unavailable — continue without Graph provider.
        // SP Search will handle all verticals as fallback.
      }

      registerBuiltInActions(this._store.getState().registries.actions);

      // Freeze registries AFTER all providers/actions are registered.
      // This prevents mid-session mutations. Must happen here (in the Results
      // web part) because it loads LAST — other web parts may still be
      // registering when the first search executes.
      if (this._orchestrator) {
        this._orchestrator.freezeRegistries();
      }
    }

    // Apply the configured default layout only when the store is still on the
    // slice default ('list'). This preserves a layout restored from URL sync
    // (e.g. ?l=grid on page refresh) instead of forcing it back to the web
    // part's configured starting view.
    if (this.properties.defaultLayout && this._store) {
      const state: ISearchStore = this._store.getState();
      if (state.activeLayoutKey === 'list' && this.properties.defaultLayout !== 'list') {
        state.setLayout(this.properties.defaultLayout);
      }
    }

    // Apply configured page size
    if (this.properties.pageSize && this._store) {
      const state: ISearchStore = this._store.getState();
      if (state.pageSize !== this.properties.pageSize) {
        this._store.setState({ pageSize: this.properties.pageSize });
      }
    }

    // Migrate legacy string properties to collection format (one-time)
    if (this.properties.selectedProperties &&
        (!this.properties.selectedPropertiesCollection || this.properties.selectedPropertiesCollection.length === 0)) {
      this._migrateSelectedPropertiesToCollection();
    }
    if (this.properties.refinementFilters &&
        (!this.properties.refinementFiltersCollection || this.properties.refinementFiltersCollection.length === 0)) {
      this._migrateRefinementFiltersToCollection();
    }

    // Initialize layoutPreset for existing installs that don't have it persisted yet.
    if (!this.properties.layoutPreset) {
      this.properties.layoutPreset = 'custom';
    }
    this._normalizeDefaultLayoutProperty();

    // Sync all property pane settings to the store BEFORE initializeSearchContext
    // because initializeSearchContext triggers the first search
    this._syncQueryTemplateToStore();
    this._syncSortablePropertiesToStore();
    this._syncSelectedPropertiesToStore();
    this._syncRefinementFiltersToStore();
    this._syncSearchConfigToStore();
    this._syncScopeToStore();
    this._syncAvailableLayoutsToStore();

    // Initialize the search context (history service, orchestrator, etc.)
    // This triggers the first search — all store config must be set above
    await initializeSearchContext(contextId, this.context);

    // T3.D10 — record this web part's registration + measure the
    // filterConfig length at first-search time so the edit-mode
    // diagnostic can surface an init-order MessageBar when Filters
    // arrives late.
    recordWebPartInit(contextId, 'SpSearchResultsWebPart');

    // Always trigger a search after initialization. The initial search inside
    // initializeSearchContext may have been skipped if another web part
    // (e.g., Filters) called initializeSearchContext before the Results web
    // part registered the data provider. This call is safe — triggerSearch
    // cancels any pending/in-flight search before starting a new one.
    if (this._orchestrator) {
      const filterConfigLength = this._store
        ? (this._store.getState().filterConfig || []).length
        : 0;
      recordFirstSearch(contextId, filterConfigLength);
      this._orchestrator.triggerSearch().catch(function noop(): void { /* handled in orchestrator */ });
    }
    DebugCollector.registerWebPart('SPSearchResultsWebPart', this.properties as unknown as Record<string, unknown>);
  }

  private _getSelectedPropertyColumns(): ISelectedPropertyColumn[] {
    const raw = this._normalizeSelectedPropertiesCollection();

    return raw.map((item: ISelectedPropertyItem): ISelectedPropertyColumn => ({
      property: item.property || '',
      alias: item.alias || ''
    }));
  }

  private _getCompactPropertyColumns(): IColumnConfigItem[] {
    return this._getCompactLayoutProperties();
  }

  private _getGridPropertyColumns(): IColumnConfigItem[] {
    return this._getGridLayoutProperties();
  }

  private _normalizeSelectedPropertiesCollection(): ISelectedPropertyItem[] {
    const raw = normalizeCollectionValue<ISelectedPropertyItem>(this.properties.selectedPropertiesCollection);
    const result: ISelectedPropertyItem[] = [];
    const seen = new Set<string>();
    let titleItem: ISelectedPropertyItem | undefined;

    for (let i: number = 0; i < raw.length; i++) {
      const property = String(raw[i].property || '').trim();
      if (!property) {
        continue;
      }
      const lookup = property.toLowerCase();
      if (seen.has(lookup)) {
        continue;
      }
      seen.add(lookup);

      const normalizedItem: ISelectedPropertyItem = {
        uniqueId: raw[i].uniqueId || ('sp-' + String(i)),
        property,
        alias: String(raw[i].alias || '').trim()
      };

      if (isTitleProperty(property)) {
        titleItem = {
          uniqueId: normalizedItem.uniqueId,
          property: 'Title',
          alias: normalizedItem.alias || 'Name'
        };
      } else {
        result.push(normalizedItem);
      }
    }

    if (!titleItem) {
      titleItem = { uniqueId: 'sp-title', property: 'Title', alias: 'Name' };
    }

    const normalized = [titleItem, ...result];
    this.properties.selectedPropertiesCollection = normalized;
    return normalized;
  }

  private _getLayoutPropertyOptions(): Array<{ key: string; text: string }> {
    return this._normalizeSelectedPropertiesCollection()
      .filter((item: ISelectedPropertyItem) => !isTitleProperty(item.property))
      .map((item: ISelectedPropertyItem) => ({
        key: item.property,
        text: item.alias ? (item.alias + ' (' + item.property + ')') : item.property
      }));
  }

  private _getCompactLayoutProperties(): IColumnConfigItem[] {
    // Stream B / Phase 3 — same shape + migration approach as the grid path
    // in `_getGridLayoutProperties`. The Compact layout's tight visual
    // budget caps the visible columns at 4; the editor list reflects that.
    const raw = normalizeCollectionValue<Partial<IColumnConfigItem> & ILegacyColumnItem>(
      this.properties.compactPropertiesCollection as unknown as Array<Partial<IColumnConfigItem> & ILegacyColumnItem>
    );

    const masterAliasMap = new Map<string, { property: string; alias: string }>();
    const master = this._normalizeSelectedPropertiesCollection();
    for (let i: number = 0; i < master.length; i++) {
      const property = master[i].property;
      if (!isTitleProperty(property)) {
        masterAliasMap.set(property.toLowerCase(), {
          property,
          alias: master[i].alias || property,
        });
      }
    }

    const result: IColumnConfigItem[] = [];
    const seen = new Set<string>();

    for (let i: number = 0; i < raw.length; i++) {
      const property = String(raw[i].property || '').trim();
      const lookup = property.toLowerCase();
      if (!property || isTitleProperty(property) || seen.has(lookup) || !masterAliasMap.has(lookup)) {
        continue;
      }
      seen.add(lookup);

      const normalized: IColumnConfigItem = normalizeColumnConfigItem(raw[i]);
      const aliasOnRaw = typeof raw[i].alias === 'string' ? String(raw[i].alias).trim() : '';
      if (!aliasOnRaw) {
        const masterEntry = masterAliasMap.get(lookup);
        if (masterEntry && masterEntry.alias && masterEntry.alias !== property) {
          normalized.alias = masterEntry.alias;
        }
      }
      result.push(normalized);
    }

    // Fallback: seed from the Compact-friendly defaults if no items configured.
    if (result.length === 0) {
      const defaults = ['Author', 'LastModifiedTime', 'Size', 'FileType'];
      for (let i: number = 0; i < defaults.length; i++) {
        const lookup = defaults[i].toLowerCase();
        const masterEntry = masterAliasMap.get(lookup);
        if (masterEntry) {
          result.push(
            normalizeColumnConfigItem({
              uniqueId: 'compact-fallback-' + String(i),
              property: masterEntry.property,
              alias: masterEntry.alias,
            })
          );
        }
      }
    }

    this.properties.compactPropertiesCollection = result;
    return result;
  }

  private _getGridLayoutProperties(): IColumnConfigItem[] {
    // Stream B / Phase 1 — accepts both legacy `{ uniqueId, property }` items
    // and the new `IColumnConfigItem` shape. Whichever is on disk is wrapped
    // via `normalizeColumnConfigItem`; missing aliases are seeded from
    // `selectedPropertiesCollection` so out-of-box display labels (e.g.
    // "Modified", "Type") survive migration.
    const raw = normalizeCollectionValue<Partial<IColumnConfigItem> & ILegacyColumnItem>(
      this.properties.gridPropertiesCollection as unknown as Array<Partial<IColumnConfigItem> & ILegacyColumnItem>
    );

    const masterAliasMap = new Map<string, { property: string; alias: string }>();
    const master = this._normalizeSelectedPropertiesCollection();
    for (let i: number = 0; i < master.length; i++) {
      const property = master[i].property;
      if (!isTitleProperty(property)) {
        masterAliasMap.set(property.toLowerCase(), {
          property,
          alias: master[i].alias || property,
        });
      }
    }

    const result: IColumnConfigItem[] = [];
    const seen = new Set<string>();

    for (let i: number = 0; i < raw.length; i++) {
      const property = String(raw[i].property || '').trim();
      const lookup = property.toLowerCase();
      if (!property || isTitleProperty(property) || seen.has(lookup) || !masterAliasMap.has(lookup)) {
        continue;
      }
      seen.add(lookup);

      const normalized: IColumnConfigItem = normalizeColumnConfigItem(raw[i]);
      // Seed alias from the master collection when the column item carries no
      // explicit alias (i.e. legacy items where normalizer fell back to property).
      const aliasOnRaw = typeof raw[i].alias === 'string' ? String(raw[i].alias).trim() : '';
      if (!aliasOnRaw) {
        const masterEntry = masterAliasMap.get(lookup);
        if (masterEntry && masterEntry.alias && masterEntry.alias !== property) {
          normalized.alias = masterEntry.alias;
        }
      }
      result.push(normalized);
    }

    // Fallback: seed from selectedPropertiesCollection when no grid items configured.
    if (result.length === 0) {
      let idx = 0;
      masterAliasMap.forEach((entry) => {
        result.push(
          normalizeColumnConfigItem({
            uniqueId: 'lp-fallback-' + String(idx),
            property: entry.property,
            alias: entry.alias,
          })
        );
        idx++;
      });
    }

    this.properties.gridPropertiesCollection = result;
    return result;
  }

  private _syncQueryTemplateToStore(): void {
    if (!this._store) {
      return;
    }
    const template: string = this.properties.queryTemplate || '{searchTerms}';
    const state: ISearchStore = this._store.getState();
    if (state.queryTemplate !== template) {
      this._store.setState({ queryTemplate: template });
    }
  }

  private _syncSortablePropertiesToStore(): void {
    if (!this._store) {
      return;
    }

    const raw = normalizeCollectionValue<ISortCollectionItem>(this.properties.sortablePropertiesCollection);

    const sortableProperties = (raw || []).map((item: ISortCollectionItem) => ({
      property: item.property,
      label: item.label,
      direction: item.direction || 'Ascending'
    }));

    const state: ISearchStore = this._store.getState();
    if (JSON.stringify(state.sortableProperties) !== JSON.stringify(sortableProperties)) {
      this._store.setState({ sortableProperties });
    }
  }

  private _syncSelectedPropertiesToStore(): void {
    if (!this._store) {
      return;
    }

    const raw = this._normalizeSelectedPropertiesCollection();

    const propertiesString = (raw || [])
      .map((item: ISelectedPropertyItem) => item.property)
      .filter(Boolean)
      .join(',');

    // Keep legacy property in sync for backward compatibility
    this.properties.selectedProperties = propertiesString;

    const state = this._store.getState();
    if (state.selectedProperties !== propertiesString) {
      this._store.setState({ selectedProperties: propertiesString });
    }
  }

  private _syncRefinementFiltersToStore(): void {
    if (!this._store) {
      return;
    }

    const raw = normalizeCollectionValue<IRefinementFilterItem>(this.properties.refinementFiltersCollection);

    const filtersString = (raw || [])
      .map((item: IRefinementFilterItem) => {
        if (!item.property || !item.value) {
          return '';
        }
        const op = (item.operator || 'equals').toLowerCase();
        switch (op) {
          case 'range':
            return item.property + ':range(' + item.value + ')';
          case 'contains':
            return item.property + ':string("*' + item.value + '*")';
          case 'beginswith':
            return item.property + ':string("' + item.value + '*")';
          default:
            // 'equals' and any other — simple quoted value
            return item.property + ':"' + item.value + '"';
        }
      })
      .filter(Boolean)
      .join(',');

    // Keep legacy property in sync for backward compatibility
    this.properties.refinementFilters = filtersString;

    const state = this._store.getState();
    if (state.refinementFilters !== filtersString) {
      this._store.setState({ refinementFilters: filtersString });
    }
  }

  /**
   * Migrate legacy comma-separated selectedProperties string to collection format.
   */
  private _migrateSelectedPropertiesToCollection(): void {
    const raw = this.properties.selectedProperties;
    if (!raw) {
      return;
    }
    const properties = raw.split(',').map((p: string) => p.trim()).filter(Boolean);
    if (properties.length === 0) {
      return;
    }
    this.properties.selectedPropertiesCollection = properties.map(
      (prop: string, idx: number) => ({
        uniqueId: 'sp-' + String(idx),
        property: prop,
        alias: ''
      })
    );
  }

  /**
   * Migrate legacy FQL refinementFilters string to collection format.
   */
  private _migrateRefinementFiltersToCollection(): void {
    const raw = this.properties.refinementFilters;
    if (!raw) {
      return;
    }
    const fqlPattern = /^(\w+):(equals|contains|range|beginsWith)\((.+)\)$/;
    const filters = raw.split(',').map((f: string) => f.trim()).filter(Boolean);
    if (filters.length === 0) {
      return;
    }
    this.properties.refinementFiltersCollection = filters.map(
      (filter: string, idx: number) => {
        const match = filter.match(fqlPattern);
        if (match) {
          return {
            uniqueId: 'rf-' + String(idx),
            property: match[1],
            operator: match[2],
            value: match[3].replace(/^"(.*)"$/, '$1')
          };
        }
        const colonIdx = filter.indexOf(':');
        if (colonIdx > 0) {
          return {
            uniqueId: 'rf-' + String(idx),
            property: filter.substring(0, colonIdx),
            operator: 'equals',
            value: filter.substring(colonIdx + 1)
          };
        }
        return {
          uniqueId: 'rf-' + String(idx),
          property: filter,
          operator: 'equals',
          value: ''
        };
      }
    );
  }

  /**
   * Sync search scope to the store based on the searchScope property.
   */
  private _syncScopeToStore(): void {
    if (!this._store) {
      return;
    }

    const scopeType = this.properties.searchScope || 'all';
    const state = this._store.getState();
    let scopeId: string;
    let scopeLabel: string;
    let kqlPath: string | undefined;

    switch (scopeType) {
      case 'currentsite': {
        const webUrl = this._getWebAbsoluteUrl();
        scopeId = 'currentsite';
        scopeLabel = 'This site';
        kqlPath = webUrl ? 'Path:"' + webUrl + '"' : undefined;
        break;
      }
      case 'currentcollection': {
        const siteUrl = this._getSiteAbsoluteUrl();
        scopeId = 'currentcollection';
        scopeLabel = 'This site collection';
        kqlPath = siteUrl ? 'Path:"' + siteUrl + '"' : undefined;
        break;
      }
      case 'custom': {
        const customPath = this.properties.searchScopePath || '';
        scopeId = 'custom';
        scopeLabel = 'Custom scope';
        kqlPath = customPath ? 'Path:"' + customPath + '"' : undefined;
        break;
      }
      default:
        scopeId = 'all';
        scopeLabel = 'All SharePoint';
        kqlPath = undefined;
        break;
    }

    // Always sync the full scope object — the URL sync middleware may have
    // set only the scope id (without kqlPath) if it initialized before this
    // web part. Compare both id and kqlPath to ensure the path restriction
    // is always present for scoped searches.
    if (state.scope.id !== scopeId || state.scope.kqlPath !== kqlPath) {
      state.setScope({ id: scopeId, label: scopeLabel, kqlPath });
    }
  }

  /**
   * Sync search configuration properties (result source, query rules, etc.) to the store.
   * Note: selectedProperties and refinementFilters are handled by dedicated sync methods.
   */
  private _syncSearchConfigToStore(): void {
    if (!this._store) {
      return;
    }

    const updates: Partial<ISearchStore> = {};
    const state = this._store.getState();

    const resultSourceId = this.properties.resultSourceId || '';
    if (state.resultSourceId !== resultSourceId) {
      updates.resultSourceId = resultSourceId;
    }

    const enableQueryRules = this.properties.enableQueryRules !== false;
    if (state.enableQueryRules !== enableQueryRules) {
      updates.enableQueryRules = enableQueryRules;
    }

    const trimDuplicates = this.properties.trimDuplicates !== false;
    if (state.trimDuplicates !== trimDuplicates) {
      updates.trimDuplicates = trimDuplicates;
    }

    const collapseSpecification = this.properties.collapseSpecification || '';
    if (state.collapseSpecification !== collapseSpecification) {
      updates.collapseSpecification = collapseSpecification;
    }

    const showPaging = this.properties.showPaging !== false;
    if (state.showPaging !== showPaging) {
      updates.showPaging = showPaging;
    }

    const pageRange = this.properties.pageRange || 5;
    if (state.pageRange !== pageRange) {
      updates.pageRange = pageRange;
    }

    if (Object.keys(updates).length > 0) {
      this._store.setState(updates);
    }
  }

  /**
   * Applies a named scenario preset by overwriting all Results web part properties
   * that the preset governs: layout visibility, default layout, query template,
   * selected properties, and sortable properties.
   *
   * Fields owned by the Filters / Verticals web parts (filterConfig,
   * dataProviderId) are NOT written here — the property pane label shows
   * suggestions for those so the admin can configure them manually.
   *
   * The admin can further customize any individual field after selecting a preset;
   * touching a layout toggle or defaultLayout automatically reverts to 'custom'.
   */
  private _applyScenarioPreset(presetId: string): void {
    const preset = SCENARIO_PRESETS[presetId];
    if (!preset) {
      return;
    }

    // Layout properties
    this.properties.defaultLayout      = preset.defaultLayout;
    this.properties.showListLayout     = preset.showListLayout;
    this.properties.showCompactLayout  = preset.showCompactLayout;
    this.properties.showGridLayout     = preset.showGridLayout;
    this.properties.showCardLayout     = preset.showCardLayout;
    this.properties.showPeopleLayout   = preset.showPeopleLayout;
    this.properties.showGalleryLayout  = preset.showGalleryLayout;

    // Query template
    this.properties.queryTemplate = preset.queryTemplate;

    // Selected properties — map to the collection format expected by the property pane
    this.properties.selectedPropertiesCollection = preset.selectedProperties.map(
      (p, idx) => ({ uniqueId: 'preset-sp-' + String(idx), property: p.property, alias: p.alias })
    );

    // Stream B / Phase 3 — compact items now carry the IColumnConfigItem
    // shape. Map alias through from the matching selectedProperty when
    // available, so out-of-box labels (Modified / Type / Size) survive.
    const presetAliasMap = new Map<string, string>(
      preset.selectedProperties.map((p) => [p.property.toLowerCase(), p.alias])
    );
    this.properties.compactPropertiesCollection = preset.compactProperties.map(
      (p, idx) =>
        normalizeColumnConfigItem({
          uniqueId: 'preset-compact-' + String(idx),
          property: p.property,
          alias: presetAliasMap.get(p.property.toLowerCase()) || p.property,
        })
    );

    this.properties.gridPropertiesCollection = preset.selectedProperties
      .filter((p) => !isTitleProperty(p.property))
      .map((p, idx) =>
        normalizeColumnConfigItem({
          uniqueId: 'preset-grid-' + String(idx),
          property: p.property,
          alias: p.alias,
        })
      );

    // Sortable properties — map to the collection format
    this.properties.sortablePropertiesCollection = preset.sortableProperties.map(
      (s, idx) => ({ uniqueId: 'preset-sort-' + String(idx), property: s.property, label: s.label, direction: s.direction })
    );

    // T4.D12 — record the preset's filter suggestions in the cross-web-part
    // registry so the Filters web part can offer to apply them via an
    // edit-mode MessageBar. Peers subscribed to the registry re-render and
    // surface the "Apply N filters from preset" CTA.
    recordPresetSuggestion(this.properties.searchContextId || 'default', {
      id: preset.id,
      label: preset.label,
      filterSuggestions: preset.filterSuggestions,
      recordedAt: 0, // overwritten by recordPresetSuggestion
    });
  }

  private _getAvailableLayoutsFromProperties(): string[] {
    const available: string[] = [];
    if (this.properties.showListLayout !== false) { available.push('list'); }
    if (this.properties.showCompactLayout !== false) { available.push('compact'); }
    if (this.properties.showGridLayout !== false) { available.push('grid'); }
    if (this.properties.showCardLayout === true) { available.push('card'); }
    if (this.properties.showPeopleLayout === true) { available.push('people'); }
    if (this.properties.showGalleryLayout === true) { available.push('gallery'); }
    if (available.length === 0) { available.push('list'); }
    return available;
  }

  private _normalizeDefaultLayoutProperty(): boolean {
    const available = this._getAvailableLayoutsFromProperties();
    const current = this.properties.defaultLayout || 'list';
    if (available.indexOf(current) >= 0) {
      return false;
    }
    this.properties.defaultLayout = available[0];
    return true;
  }

  private _buildDefaultLayoutOptions(): Array<{ key: string; text: string; iconProps: { officeFabricIconFontName: string } }> {
    const available = this._getAvailableLayoutsFromProperties();
    const allOptions = [
      { key: 'list', text: strings.ListLayoutText, iconProps: { officeFabricIconFontName: 'List' } },
      { key: 'compact', text: strings.CompactLayoutText, iconProps: { officeFabricIconFontName: 'GridViewSmall' } },
      { key: 'grid', text: strings.GridLayoutText, iconProps: { officeFabricIconFontName: 'Table' } },
      { key: 'card', text: strings.CardLayoutText, iconProps: { officeFabricIconFontName: 'ContactCard' } },
      { key: 'people', text: strings.PeopleLayoutText, iconProps: { officeFabricIconFontName: 'People' } },
      { key: 'gallery', text: strings.GalleryLayoutText, iconProps: { officeFabricIconFontName: 'PictureLibrary' } }
    ];

    return allOptions.filter((option) => available.indexOf(option.key) >= 0);
  }

  private _buildPresetOptions(): Array<{ key: string; text: string; iconProps: { officeFabricIconFontName: string } }> {
    // T4.D2 — surface all 9 presets unconditionally. The People preset
    // is shown even when Graph isn't available so admins can discover the
    // capability and select it; `_isPeoplePresetBlocked` then renders a
    // graceful warning explaining what's missing.
    return [
      { key: 'custom',          text: 'Custom',         iconProps: { officeFabricIconFontName: 'Settings' } },
      { key: 'general',         text: 'General',        iconProps: { officeFabricIconFontName: 'Search' } },
      { key: 'documents',       text: 'Documents',      iconProps: { officeFabricIconFontName: 'DocLibrary' } },
      { key: 'hub-search',      text: 'Hub Search',     iconProps: { officeFabricIconFontName: 'Globe' } },
      { key: 'knowledge-base',  text: 'Knowledge Base', iconProps: { officeFabricIconFontName: 'BookAnswers' } },
      { key: 'policy-search',   text: 'Policy Search',  iconProps: { officeFabricIconFontName: 'Shield' } },
      { key: 'news',            text: 'News',           iconProps: { officeFabricIconFontName: 'News' } },
      { key: 'media',           text: 'Media',          iconProps: { officeFabricIconFontName: 'Photo2' } },
      { key: 'people',          text: 'People',         iconProps: { officeFabricIconFontName: 'Group' } }
    ];
  }

  /**
   * T4.D2 — true when the admin has selected `people` but no Graph
   * client is available (most commonly: the tenant hasn't approved the
   * Microsoft Graph permission at SharePoint admin → API access).
   * Drives the graceful-warning label below the preset picker.
   */
  private _isPeoplePresetBlocked(): boolean {
    return this.properties.layoutPreset === 'people' && !this._graphOrgService;
  }

  private _shouldShowSpecializedViews(): boolean {
    if (this.properties.showCardLayout === true ||
      this.properties.showPeopleLayout === true ||
      this.properties.showGalleryLayout === true) {
      return true;
    }

    const presetId = this.properties.layoutPreset || 'custom';
    return ['people', 'news', 'media', 'knowledge-base', 'hub-search'].indexOf(presetId) >= 0;
  }

  private _syncAvailableLayoutsToStore(): void {
    if (!this._store) {
      return;
    }
    const available = this._getAvailableLayoutsFromProperties();
    this._store.getState().setAvailableLayouts(available);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    this._theme = currentTheme;

    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
      this.domElement.style.setProperty('--link', semanticColors.link || '');
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
    }
  }

  protected onDispose(): void {
    // T3.D1 — drop the refcount before unmounting React. See SearchBox.
    const contextId: string = this.properties.searchContextId || 'default';
    decrementContextRef(contextId);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'searchContextId') {
      const contextId: string = this.properties.searchContextId || 'default';
      this._store = getStore(contextId);
      this._orchestrator = getOrchestrator(contextId);
    }

    if (propertyPath === 'queryTemplate' || propertyPath === 'searchContextId') {
      this._syncQueryTemplateToStore();
    }

    if (propertyPath === 'sortablePropertiesCollection' || propertyPath === 'searchContextId') {
      this._syncSortablePropertiesToStore();
    }

    if (propertyPath === 'selectedPropertiesCollection' || propertyPath === 'searchContextId') {
      this._getCompactLayoutProperties();
      this._getGridLayoutProperties();
      this._syncSelectedPropertiesToStore();
      this.render();
      this.context.propertyPane.refresh();
    }

    if (propertyPath === 'compactPropertiesCollection' || propertyPath === 'gridPropertiesCollection' || propertyPath === 'titleDisplayMode') {
      this.render();
    }

    if (propertyPath === 'refinementFiltersCollection' || propertyPath === 'searchContextId') {
      this._syncRefinementFiltersToStore();
    }

    if (propertyPath === 'pageSize' && this._store) {
      this._store.setState({ pageSize: this.properties.pageSize, currentPage: 1 });
    }

    if (propertyPath === 'defaultLayout' && this._store) {
      // Manual default layout change → abandon any active preset
      if (this.properties.layoutPreset !== 'custom') {
        this.properties.layoutPreset = 'custom';
        this.context.propertyPane.refresh();
      }
      this._normalizeDefaultLayoutProperty();
      const state: ISearchStore = this._store.getState();
      state.setLayout(this.properties.defaultLayout);
    }

    if (propertyPath === 'layoutPreset') {
      this._applyScenarioPreset(this.properties.layoutPreset);
      this._normalizeDefaultLayoutProperty();
      this._syncAvailableLayoutsToStore();
      if (this._store) {
        this._store.getState().setLayout(this.properties.defaultLayout || 'list');
      }
      // Refresh the property pane so all individual toggles show their updated values.
      this.context.propertyPane.refresh();
    }

    // Sync remaining search configuration properties
    if (['resultSourceId', 'enableQueryRules', 'trimDuplicates',
      'collapseSpecification', 'showPaging', 'pageRange', 'searchContextId'].indexOf(propertyPath) >= 0) {
      this._syncSearchConfigToStore();
    }

    if (propertyPath === 'searchScope' || propertyPath === 'searchScopePath' || propertyPath === 'searchContextId') {
      this._syncScopeToStore();
    }

    const layoutProps = ['showListLayout', 'showCompactLayout', 'showGridLayout',
      'showCardLayout', 'showPeopleLayout', 'showGalleryLayout', 'searchContextId'];
    if (layoutProps.indexOf(propertyPath) >= 0) {
      // Individual toggle change → revert to Custom so the preset radio stays accurate
      if (propertyPath !== 'searchContextId' && this.properties.layoutPreset !== 'custom') {
        this.properties.layoutPreset = 'custom';
      }
      const defaultChanged = this._normalizeDefaultLayoutProperty();
      this._syncAvailableLayoutsToStore();
      if (this._store && defaultChanged) {
        this._store.getState().setLayout(this.properties.defaultLayout || 'list');
      }
      if (propertyPath !== 'searchContextId' || defaultChanged) {
        this.context.propertyPane.refresh();
      }
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        // ─── Page 1: Data Source ───────────────────────────
        {
          header: {
            description: strings.DataSourcePageHeader
          },
          groups: [
            // T4.D2 — "Get started" hoists the scenario preset picker to
            // the first field on page 1, ahead of any data-source detail.
            // T3.D4 — `searchContextId` follows the picker in the same group
            // so both first-impression knobs are visible without a scroll.
            {
              groupName: strings.GetStartedGroupName,
              groupFields: [
                // T4.D11 — context-sensitive help link.
                propertyPaneGroupHelp('quick-start', 'Help: Quick Start presets'),
                PropertyPaneChoiceGroup('layoutPreset', {
                  label: strings.ScenarioPresetLabel,
                  options: this._buildPresetOptions()
                }),
                PropertyPaneLabel('layoutPresetHint', {
                  text: strings.ScenarioPresetHint
                }),
                // T4.D2 — graceful warning when admin selects `people` but
                // Graph isn't available. The preset still applies; this
                // label tells admins what to fix.
                ...(this._isPeoplePresetBlocked() ? [
                  PropertyPaneLabel('layoutPresetPeopleWarning', {
                    text: strings.ScenarioPresetPeopleWarning
                  })
                ] : []),
                propertyPaneSearchContextIdField()
              ]
            },
            {
              groupName: strings.DataGroupName,
              groupFields: [
                // T4.D11 — context-sensitive help link.
                propertyPaneGroupHelp('results-data', 'Help: Search scope and managed properties'),
                PropertyPaneDropdown('searchScope', {
                  label: strings.SearchScopeLabel,
                  options: [
                    { key: 'all', text: strings.ScopeAllText },
                    { key: 'currentsite', text: strings.ScopeCurrentSiteText },
                    { key: 'currentcollection', text: strings.ScopeCurrentCollectionText },
                    { key: 'custom', text: strings.ScopeCustomText }
                  ],
                  selectedKey: this.properties.searchScope || 'all'
                }),
                ...(this.properties.searchScope === 'custom' ? [
                  PropertyPaneTextField('searchScopePath', {
                    label: strings.SearchScopePathLabel,
                    description: strings.SearchScopePathDescription,
                    placeholder: 'https://contoso.sharepoint.com/sites/hr'
                  })
                ] : []),
                PropertyPaneSchemaHelper('queryTemplate', {
                  label: strings.QueryTemplateLabel,
                  description: strings.QueryTemplateDescription,
                  value: this.properties.queryTemplate || '',
                  filterHint: 'queryable',
                  applyOnEnter: true,
                }),
                PropertyPaneTextField('resultSourceId', {
                  label: strings.ResultSourceIdLabel,
                  description: strings.ResultSourceIdDescription
                }),
                PropertyFieldCollectionData('selectedPropertiesCollection', {
                  key: 'selectedPropertiesCollection',
                  label: strings.SelectedPropertiesLabel,
                  panelHeader: strings.SelectedPropertiesPanelHeader,
                  manageBtnLabel: strings.SelectedPropertiesManageBtn,
                  value: this._normalizeSelectedPropertiesCollection(),
                  enableSorting: true,
                  fields: [
                    {
                      id: 'property',
                      title: strings.SelectedPropertyColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'LastModifiedTime'
                    },
                    {
                      id: 'alias',
                      title: strings.SelectedPropertyAliasColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'Date Modified'
                    }
                  ]
                }),
                // Stream B / Phase 3 — quiet-deprecation note. `alias` here
                // is no longer read by the column-render path; per-column
                // display labels live on the Grid / Compact column editors.
                PropertyPaneLabel('selectedPropertiesAliasDeprecation', {
                  text: strings.SelectedPropertiesAliasDeprecationNote
                })
              ]
            },
            {
              groupName: strings.SortGroupName,
              groupFields: [
                PropertyFieldCollectionData('sortablePropertiesCollection', {
                  key: 'sortablePropertiesCollection',
                  label: strings.SortFieldLabel,
                  panelHeader: strings.SortPanelHeader,
                  manageBtnLabel: strings.SortManageBtn,
                  value: normalizeCollectionValue<ISortCollectionItem>(this.properties.sortablePropertiesCollection),
                  enableSorting: true,
                  fields: [
                    {
                      id: 'property',
                      title: strings.SortPropertyColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'LastModifiedTime'
                    },
                    {
                      id: 'label',
                      title: strings.SortLabelColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'Date Modified'
                    },
                    {
                      id: 'direction',
                      title: strings.SortDirectionColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: [
                        { key: 'Ascending', text: strings.SortAscending },
                        { key: 'Descending', text: strings.SortDescending }
                      ]
                    }
                  ]
                })
              ]
            },
            {
              groupName: strings.PaginationGroupName,
              groupFields: [
                PropertyPaneSlider('pageSize', {
                  label: strings.PageSizeLabel,
                  min: 5,
                  max: 100,
                  value: this.properties.pageSize || 25,
                  step: 5,
                  showValue: true
                }),
                PropertyPaneToggle('showPaging', {
                  label: strings.ShowPagingLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneSlider('pageRange', {
                  label: strings.PageRangeLabel,
                  min: 3,
                  max: 10,
                  value: this.properties.pageRange || 5,
                  step: 1,
                  showValue: true
                })
              ]
            }
          ]
        },
        // ─── Page 2: Display ──────────────────────────────
        {
          header: {
            description: strings.DisplayPageHeader
          },
          groups: [
            {
              groupName: strings.MainLayoutsGroupName,
              groupFields: [
                // T4.D11 — context-sensitive help link.
                propertyPaneGroupHelp('results-layouts', 'Help: Layouts and presets'),
                // T4.D2 — preset picker moved to page-1 "Get started" group.
                PropertyPaneChoiceGroup('defaultLayout', {
                  label: strings.DefaultLayoutLabel,
                  options: this._buildDefaultLayoutOptions()
                }),
                PropertyPaneToggle('showListLayout', {
                  label: 'Show List view',
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('showCompactLayout', {
                  label: 'Show Compact List view',
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('showGridLayout', {
                  label: 'Show Data Grid view',
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                })
              ]
            },
            ...(this._shouldShowSpecializedViews() ? [{
              groupName: strings.AdvancedLayoutsGroupName,
              groupFields: [
                PropertyPaneToggle('showCardLayout', {
                  label: 'Show Card view',
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('showPeopleLayout', {
                  label: 'Show People view',
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('showGalleryLayout', {
                  label: 'Show Gallery view',
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                })
              ]
            }] : []),
            ...(this.properties.showCompactLayout !== false ? [{
              groupName: strings.CompactViewGroupName,
              groupFields: [
                // Stream B / Phase 3 — Compact adopts the same custom field
                // as the grid. Renderer dispatch flows through the unified
                // `renderCell.tsx` module via `CompactLayout`.
                PropertyPaneColumnConfigField('compactPropertiesCollection', {
                  label: strings.CompactPropertiesLabel,
                  value: this._getCompactLayoutProperties(),
                  availableProperties: this._getLayoutPropertyOptions(),
                })
              ]
            }] : []),
            ...(this.properties.showGridLayout !== false ? [{
              groupName: strings.GridViewGroupName,
              groupFields: [
                // Stream B / Phase 1 — custom column-config field with a
                // side-panel editor in place of PnP's flat collection-data
                // table. Phase 2 added richText / tags / number / boolean
                // renderer types; Phase 3 added the column-chooser toggle.
                PropertyPaneColumnConfigField('gridPropertiesCollection', {
                  label: strings.GridPropertiesLabel,
                  value: this._getGridLayoutProperties(),
                  availableProperties: this._getLayoutPropertyOptions(),
                }),
                PropertyPaneToggle('showColumnChooser', {
                  label: strings.ShowColumnChooserLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                })
              ]
            }] : []),
            {
              groupName: strings.BehaviorGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('titleDisplayMode', {
                  label: strings.TitleDisplayModeLabel,
                  options: [
                    { key: 'ellipsis', text: strings.TitleDisplayEllipsisText },
                    { key: 'middle', text: strings.TitleDisplayMiddleText },
                    { key: 'wrap', text: strings.TitleDisplayWrapText }
                  ]
                }),
                PropertyPaneToggle('showSortDropdown', {
                  label: strings.ShowSortDropdownLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('showResultCount', {
                  label: strings.ShowResultCountLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('enablePreviewPanel', {
                  label: strings.ShowPreviewPanelLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('hideWebPartWhenNoResults', {
                  label: strings.HideWebPartWhenNoResultsLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneTextField('emptyResultsMessage', {
                  label: strings.EmptyResultsMessageLabel,
                  description: strings.EmptyResultsMessageDescription,
                  multiline: true,
                  rows: 4,
                  resizable: true
                }),
              ]
            },
            // ─── Result link behaviour (Stream C / #7) ──────────
            {
              groupName: strings.ResultLinkGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('resultClickTarget', {
                  label: strings.ResultClickTargetLabel,
                  options: [
                    { key: 'panel',     text: strings.ResultClickTargetPanelText },
                    { key: 'newTab',    text: strings.ResultClickTargetNewTabText },
                    { key: 'sameTab',   text: strings.ResultClickTargetSameTabText },
                    { key: 'sidePanel', text: strings.ResultClickTargetSidePanelText }
                  ]
                }),
                ...((this.properties.resultClickTarget || 'panel') !== 'sidePanel' ? [
                  PropertyPaneDropdown('documentLinkMode', {
                    label: strings.DocumentLinkModeLabel,
                    options: [
                      { key: 'file',            text: strings.DocumentLinkModeFileText },
                      { key: 'propertiesForm', text: strings.DocumentLinkModePropertiesFormText }
                    ]
                  }),
                  PropertyPaneDropdown('listItemLinkMode', {
                    label: strings.ListItemLinkModeLabel,
                    options: [
                      { key: 'displayForm', text: strings.ListItemLinkModeDisplayFormText },
                      { key: 'editForm',    text: strings.ListItemLinkModeEditFormText }
                    ]
                  })
                ] : [])
              ]
            }
          ]
        },
        // ─── Page 3: Connections ───────────────────────────
        // T3.D4 — searchContextId moved to page-1 group-1; this page now
        // hosts audience targeting only.
        {
          header: {
            description: strings.ConnectionsPageHeader
          },
          groups: [
            // Stream D / #10 — per-web-part audience targeting.
            {
              groupName: strings.AudienceTargetingGroupName,
              groupFields: [
                PropertyPaneTextField('audienceGroups', {
                  label: strings.AudienceTargetingLabel,
                  description: strings.AudienceTargetingDescription,
                  multiline: true,
                  rows: 3,
                  resizable: true
                })
              ]
            }
          ]
        },
        // ─── Page 4: Advanced ──────────────────────────────
        {
          header: {
            description: strings.AdvancedPageHeader
          },
          groups: [
            {
              groupName: strings.QuerySettingsGroupName,
              groupFields: [
                PropertyPaneToggle('enableQueryRules', {
                  label: strings.EnableQueryRulesLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('trimDuplicates', {
                  label: strings.TrimDuplicatesLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                })
              ]
            },
            {
              groupName: strings.AdvancedGroupName,
              groupFields: [
                PropertyFieldCollectionData('refinementFiltersCollection', {
                  key: 'refinementFiltersCollection',
                  label: strings.RefinementFiltersLabel,
                  panelHeader: strings.RefinementFiltersPanelHeader,
                  manageBtnLabel: strings.RefinementFiltersManageBtn,
                  value: normalizeCollectionValue<IRefinementFilterItem>(this.properties.refinementFiltersCollection),
                  enableSorting: false,
                  fields: [
                    {
                      id: 'property',
                      title: strings.RefinementFilterPropertyColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'FileType'
                    },
                    {
                      id: 'operator',
                      title: strings.RefinementFilterOperatorColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: [
                        { key: 'equals', text: strings.FqlOperatorEquals },
                        { key: 'contains', text: strings.FqlOperatorContains },
                        { key: 'range', text: strings.FqlOperatorRange },
                        { key: 'beginsWith', text: strings.FqlOperatorBeginsWith }
                      ]
                    },
                    {
                      id: 'value',
                      title: strings.RefinementFilterValueColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'docx'
                    }
                  ]
                }),
                // T4.D3 — schema-helper-backed editor with sortable filter.
                // The browser pivot opens to the Sortable tab and admins
                // can browse only sortable managed properties; the field's
                // `onGetErrorMessage` adds did-you-mean / non-sortable
                // validation against the cached schema (closes the audit's
                // "Collapse spec rejects non-sortable" acceptance signal).
                PropertyPaneSchemaHelper('collapseSpecification', {
                  label: strings.CollapseSpecificationLabel,
                  description: strings.CollapseSpecificationDescription,
                  value: this.properties.collapseSpecification || '',
                  filterHint: 'sortable',
                  validation: { requireSortable: true },
                }),
                PropertyPaneToggle('showDeleteConfirmation', {
                  label: strings.ShowDeleteConfirmationLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
