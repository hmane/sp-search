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

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'SpSearchResultsWebPartStrings';
import SpSearchResults from './components/SpSearchResults';
import { ISpSearchResultsProps, ISelectedPropertyColumn } from './components/ISpSearchResultsProps';
import { ISearchStore } from '@interfaces/index';
import {
  getStore,
  getOrchestrator,
  initializeSearchContext
} from '@store/store';
import { SearchOrchestrator } from '@orchestrator/SearchOrchestrator';
import { registerBuiltInActions } from './registerBuiltInActions';
import { SharePointSearchProvider, GraphSearchProvider } from '@providers/index';
import { PropertyPaneSchemaHelper } from '../../propertyPaneControls/PropertyPaneSchemaHelper';
import { SCENARIO_PRESETS } from './presets/searchPresets';
import { GraphOrgService } from './components/GraphOrgService';
import { TitleDisplayMode } from './components/documentTitleUtils';

// Bundle DevExtreme CSS — injected via style-loader at runtime.
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('devextreme/dist/css/dx.common.css');
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('devextreme/dist/css/dx.light.css');
void spfxToolkitStylesLoaded;

export interface ISpSearchResultsWebPartProps {
  searchContextId: string;
  queryTemplate: string;
  selectedProperties: string;
  selectedPropertiesCollection: ISelectedPropertyItem[];
  compactPropertiesCollection: ILayoutPropertyItem[];
  gridPropertiesCollection: ILayoutPropertyItem[];
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
  titleDisplayMode: TitleDisplayMode;
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

interface ILayoutPropertyItem {
  uniqueId: string;
  property: string;
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
    const compactPropertyColumns: ISelectedPropertyColumn[] = this._getCompactPropertyColumns();
    const gridPropertyColumns: ISelectedPropertyColumn[] = this._getGridPropertyColumns();
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
        titleDisplayMode: this.properties.titleDisplayMode || 'wrap',
        defaultLayout: this.properties.defaultLayout,
        pageSize: this.properties.pageSize,
        isEditMode: this.displayMode === DisplayMode.Edit,
        selectedPropertyColumns,
        gridPropertyColumns,
        compactPropertyColumns,
        queryTemplate: this.properties.queryTemplate || '{searchTerms}',
        graphOrgService: this._graphOrgService
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // Initialize SPContext for PnPjs
    await SPContext.basic(this.context, 'SPSearchResults');

    const contextId: string = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);
    this._orchestrator = getOrchestrator(contextId);

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

    // Always trigger a search after initialization. The initial search inside
    // initializeSearchContext may have been skipped if another web part
    // (e.g., Filters) called initializeSearchContext before the Results web
    // part registered the data provider. This call is safe — triggerSearch
    // cancels any pending/in-flight search before starting a new one.
    if (this._orchestrator) {
      this._orchestrator.triggerSearch().catch(function noop(): void { /* handled in orchestrator */ });
    }
  }

  private _getSelectedPropertyColumns(): ISelectedPropertyColumn[] {
    const raw = this._normalizeSelectedPropertiesCollection();

    return raw.map((item: ISelectedPropertyItem): ISelectedPropertyColumn => ({
      property: item.property || '',
      alias: item.alias || ''
    }));
  }

  private _getCompactPropertyColumns(): ISelectedPropertyColumn[] {
    return this._mapLayoutPropertyColumns(this._getCompactLayoutProperties());
  }

  private _getGridPropertyColumns(): ISelectedPropertyColumn[] {
    return this._mapLayoutPropertyColumns(this._getGridLayoutProperties());
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

  private _normalizeLayoutPropertyCollection(raw: ILayoutPropertyItem[], fallbackProperties: string[] = []): ILayoutPropertyItem[] {
    const options = this._getLayoutPropertyOptions();
    const allowed = new Set<string>(options.map((item) => String(item.key).toLowerCase()));
    const result: ILayoutPropertyItem[] = [];
    const seen = new Set<string>();

    for (let i: number = 0; i < raw.length; i++) {
      const property = String(raw[i].property || '').trim();
      const lookup = property.toLowerCase();
      if (!property || isTitleProperty(property) || seen.has(lookup) || !allowed.has(lookup)) {
        continue;
      }
      seen.add(lookup);
      result.push({
        uniqueId: raw[i].uniqueId || ('lp-' + String(i)),
        property
      });
    }

    if (result.length === 0 && fallbackProperties.length > 0) {
      for (let i: number = 0; i < fallbackProperties.length; i++) {
        const property = String(fallbackProperties[i] || '').trim();
        const lookup = property.toLowerCase();
        if (!property || seen.has(lookup) || !allowed.has(lookup)) {
          continue;
        }
        seen.add(lookup);
        result.push({
          uniqueId: 'lp-fallback-' + String(i),
          property
        });
      }
    }

    return result;
  }

  private _getCompactLayoutProperties(): ILayoutPropertyItem[] {
    const raw = normalizeCollectionValue<ILayoutPropertyItem>(this.properties.compactPropertiesCollection);
    const normalized = this._normalizeLayoutPropertyCollection(raw, [
      'Author',
      'LastModifiedTime',
      'Size',
      'FileType'
    ]);
    this.properties.compactPropertiesCollection = normalized;
    return normalized;
  }

  private _getGridLayoutProperties(): ILayoutPropertyItem[] {
    const raw = normalizeCollectionValue<ILayoutPropertyItem>(this.properties.gridPropertiesCollection);
    const selected = this._normalizeSelectedPropertiesCollection()
      .map((item) => item.property)
      .filter((property) => !isTitleProperty(property));
    const normalized = this._normalizeLayoutPropertyCollection(raw, selected);
    this.properties.gridPropertiesCollection = normalized;
    return normalized;
  }

  private _mapLayoutPropertyColumns(layoutProperties: ILayoutPropertyItem[]): ISelectedPropertyColumn[] {
    const master = this._normalizeSelectedPropertiesCollection();
    const masterMap = new Map<string, ISelectedPropertyItem>();
    for (let i: number = 0; i < master.length; i++) {
      masterMap.set(master[i].property.toLowerCase(), master[i]);
    }

    return layoutProperties
      .map((item: ILayoutPropertyItem): ISelectedPropertyColumn | undefined => {
        const match = masterMap.get(String(item.property || '').toLowerCase());
        if (!match) {
          return undefined;
        }
        return {
          property: match.property,
          alias: match.alias || match.property
        };
      })
      .filter((item: ISelectedPropertyColumn | undefined): item is ISelectedPropertyColumn => !!item);
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

    this.properties.compactPropertiesCollection = preset.compactProperties.map(
      (p, idx) => ({ uniqueId: 'preset-compact-' + String(idx), property: p.property })
    );

    this.properties.gridPropertiesCollection = preset.selectedProperties
      .filter((p) => !isTitleProperty(p.property))
      .map((p, idx) => ({ uniqueId: 'preset-grid-' + String(idx), property: p.property }));

    // Sortable properties — map to the collection format
    this.properties.sortablePropertiesCollection = preset.sortableProperties.map(
      (s, idx) => ({ uniqueId: 'preset-sort-' + String(idx), property: s.property, label: s.label, direction: s.direction })
    );
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
    const options = [
      { key: 'custom',          text: 'Custom',         iconProps: { officeFabricIconFontName: 'Settings' } },
      { key: 'general',         text: 'General',        iconProps: { officeFabricIconFontName: 'Search' } },
      { key: 'documents',       text: 'Documents',      iconProps: { officeFabricIconFontName: 'DocLibrary' } },
      { key: 'hub-search',      text: 'Hub Search',     iconProps: { officeFabricIconFontName: 'Globe' } },
      { key: 'knowledge-base',  text: 'Knowledge Base', iconProps: { officeFabricIconFontName: 'BookAnswers' } },
      { key: 'policy-search',   text: 'Policy Search',  iconProps: { officeFabricIconFontName: 'Shield' } },
      { key: 'news',            text: 'News',           iconProps: { officeFabricIconFontName: 'News' } },
      { key: 'media',           text: 'Media',          iconProps: { officeFabricIconFontName: 'Photo2' } }
    ];

    const currentPreset = this.properties.layoutPreset || 'custom';
    const shouldShowPeople = !!this._graphOrgService || currentPreset === 'people';
    if (shouldShowPeople) {
      options.push({ key: 'people', text: 'People', iconProps: { officeFabricIconFontName: 'Group' } });
    }

    return options;
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
            {
              groupName: strings.DataGroupName,
              groupFields: [
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
                PropertyPaneChoiceGroup('layoutPreset', {
                  label: strings.ScenarioPresetLabel,
                  options: this._buildPresetOptions()
                }),
                PropertyPaneLabel('layoutPresetHint', {
                  text: strings.ScenarioPresetHint
                }),
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
                PropertyFieldCollectionData('compactPropertiesCollection', {
                  key: 'compactPropertiesCollection',
                  label: strings.CompactPropertiesLabel,
                  panelHeader: strings.CompactPropertiesPanelHeader,
                  manageBtnLabel: strings.CompactPropertiesManageBtn,
                  value: this._getCompactLayoutProperties(),
                  enableSorting: true,
                  fields: [
                    {
                      id: 'property',
                      title: strings.SelectedPropertyColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: this._getLayoutPropertyOptions()
                    }
                  ]
                })
              ]
            }] : []),
            ...(this.properties.showGridLayout !== false ? [{
              groupName: strings.GridViewGroupName,
              groupFields: [
                PropertyFieldCollectionData('gridPropertiesCollection', {
                  key: 'gridPropertiesCollection',
                  label: strings.GridPropertiesLabel,
                  panelHeader: strings.GridPropertiesPanelHeader,
                  manageBtnLabel: strings.GridPropertiesManageBtn,
                  value: this._getGridLayoutProperties(),
                  enableSorting: true,
                  fields: [
                    {
                      id: 'property',
                      title: strings.SelectedPropertyColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: this._getLayoutPropertyOptions()
                    }
                  ]
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
              ]
            }
          ]
        },
        // ─── Page 3: Connections ───────────────────────────
        {
          header: {
            description: strings.ConnectionsPageHeader
          },
          groups: [
            {
              groupName: strings.ConnectionsGroupName,
              groupFields: [
                PropertyPaneTextField('searchContextId', {
                  label: strings.SearchContextIdLabel,
                  description: strings.SearchContextIdDescription,
                  onGetErrorMessage: (value: string): string => {
                    if (!value || value.trim() === '') {
                      return 'Required — enter an ID to connect search web parts on this page (e.g. "hr-search").';
                    }
                    return '';
                  }
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
                PropertyPaneTextField('collapseSpecification', {
                  label: strings.CollapseSpecificationLabel,
                  description: strings.CollapseSpecificationDescription
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
