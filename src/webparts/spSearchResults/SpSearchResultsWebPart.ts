import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
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
import { SPComponentLoader } from '@microsoft/sp-loader';
import { StoreApi } from 'zustand/vanilla';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'SpSearchResultsWebPartStrings';
import SpSearchResults from './components/SpSearchResults';
import { ISpSearchResultsProps } from './components/ISpSearchResultsProps';
import { ISearchStore } from '@interfaces/index';
import {
  getStore,
  getOrchestrator,
  initializeSearchContext
} from '@store/store';
import { SearchOrchestrator } from '@orchestrator/SearchOrchestrator';
import { registerBuiltInActions } from './registerBuiltInActions';
import { SharePointSearchProvider } from '@providers/index';
import { PropertyPaneSchemaHelper } from '../../propertyPaneControls/PropertyPaneSchemaHelper';

export interface ISpSearchResultsWebPartProps {
  searchContextId: string;
  queryTemplate: string;
  selectedProperties: string;
  selectedPropertiesCollection: ISelectedPropertyItem[];
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
  searchScope: string;
  searchScopePath: string;
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

export default class SpSearchResultsWebPart extends BaseClientSideWebPart<ISpSearchResultsWebPartProps> {

  private _theme: IReadonlyTheme | undefined;
  private _store: StoreApi<ISearchStore> | undefined;
  private _orchestrator: SearchOrchestrator | undefined;

  public render(): void {
    if (!this._store) {
      return;
    }

    const contextId: string = this.properties.searchContextId || 'default';
    const element: React.ReactElement<ISpSearchResultsProps> = React.createElement(
      SpSearchResults,
      {
        store: this._store,
        orchestrator: this._orchestrator,
        searchContextId: contextId,
        theme: this._theme,
        showResultCount: this.properties.showResultCount,
        showSortDropdown: this.properties.showSortDropdown,
        defaultLayout: this.properties.defaultLayout,
        pageSize: this.properties.pageSize
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // Load DevExtreme CSS from CDN
    SPComponentLoader.loadCss('https://cdn3.devexpress.com/jslib/22.2.3/css/dx.common.css');
    SPComponentLoader.loadCss('https://cdn3.devexpress.com/jslib/22.2.3/css/dx.light.css');

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

      registerBuiltInActions(this._store.getState().registries.actions);

      // Freeze registries AFTER all providers/actions are registered.
      // This prevents mid-session mutations. Must happen here (in the Results
      // web part) because it loads LAST — other web parts may still be
      // registering when the first search executes.
      if (this._orchestrator) {
        this._orchestrator.freezeRegistries();
      }
    }

    // Apply default layout if configured
    if (this.properties.defaultLayout && this._store) {
      const state: ISearchStore = this._store.getState();
      if (state.activeLayoutKey !== this.properties.defaultLayout) {
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

    // Sync all property pane settings to the store BEFORE initializeSearchContext
    // because initializeSearchContext triggers the first search
    this._syncQueryTemplateToStore();
    this._syncSortablePropertiesToStore();
    this._syncSelectedPropertiesToStore();
    this._syncRefinementFiltersToStore();
    this._syncSearchConfigToStore();
    this._syncScopeToStore();

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

    // Handle both array (from property pane) and JSON string (from PnP PowerShell provisioning)
    let raw = this.properties.sortablePropertiesCollection;
    if (typeof raw === 'string') {
      try { raw = JSON.parse(raw); } catch { raw = []; }
    }

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

    // Handle both array (from property pane) and JSON string (from PnP PowerShell provisioning)
    let raw = this.properties.selectedPropertiesCollection;
    if (typeof raw === 'string') {
      try { raw = JSON.parse(raw); } catch { raw = []; }
    }

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

    // Handle both array (from property pane) and JSON string (from PnP PowerShell provisioning)
    let raw = this.properties.refinementFiltersCollection;
    if (typeof raw === 'string') {
      try { raw = JSON.parse(raw); } catch { raw = []; }
    }

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
        const webUrl = this.context.pageContext.web.absoluteUrl;
        scopeId = 'currentsite';
        scopeLabel = 'This site';
        kqlPath = 'Path:"' + webUrl + '"';
        break;
      }
      case 'currentcollection': {
        const siteUrl = this.context.pageContext.site.absoluteUrl;
        scopeId = 'currentcollection';
        scopeLabel = 'This site collection';
        kqlPath = 'Path:"' + siteUrl + '"';
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
      this._syncSelectedPropertiesToStore();
    }

    if (propertyPath === 'refinementFiltersCollection' || propertyPath === 'searchContextId') {
      this._syncRefinementFiltersToStore();
    }

    if (propertyPath === 'pageSize' && this._store) {
      this._store.setState({ pageSize: this.properties.pageSize });
    }

    if (propertyPath === 'defaultLayout' && this._store) {
      const state: ISearchStore = this._store.getState();
      state.setLayout(this.properties.defaultLayout);
    }

    // Sync remaining search configuration properties
    if (['resultSourceId', 'enableQueryRules', 'trimDuplicates',
      'collapseSpecification', 'showPaging', 'pageRange', 'searchContextId'].indexOf(propertyPath) >= 0) {
      this._syncSearchConfigToStore();
    }

    if (propertyPath === 'searchScope' || propertyPath === 'searchScopePath' || propertyPath === 'searchContextId') {
      this._syncScopeToStore();
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
                PropertyPaneTextField('searchContextId', {
                  label: strings.SearchContextIdLabel,
                  description: strings.SearchContextIdDescription
                }),
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
                  value: this.properties.selectedPropertiesCollection,
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
                PropertyFieldCollectionData('refinementFiltersCollection', {
                  key: 'refinementFiltersCollection',
                  label: strings.RefinementFiltersLabel,
                  panelHeader: strings.RefinementFiltersPanelHeader,
                  manageBtnLabel: strings.RefinementFiltersManageBtn,
                  value: this.properties.refinementFiltersCollection,
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
                })
              ]
            },
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
                }),
                PropertyPaneSlider('pageSize', {
                  label: strings.PageSizeLabel,
                  min: 5,
                  max: 100,
                  value: this.properties.pageSize || 25,
                  step: 5,
                  showValue: true
                })
              ]
            },
            {
              groupName: strings.PaginationGroupName,
              groupFields: [
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
              groupName: strings.LayoutGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('defaultLayout', {
                  label: strings.DefaultLayoutLabel,
                  options: [
                    { key: 'list', text: strings.ListLayoutText, iconProps: { officeFabricIconFontName: 'List' } },
                    { key: 'compact', text: strings.CompactLayoutText, iconProps: { officeFabricIconFontName: 'GridViewSmall' } },
                    { key: 'grid', text: strings.GridLayoutText, iconProps: { officeFabricIconFontName: 'GridViewMedium' } },
                    { key: 'card', text: strings.CardLayoutText, iconProps: { officeFabricIconFontName: 'ContactCard' } },
                    { key: 'people', text: strings.PeopleLayoutText, iconProps: { officeFabricIconFontName: 'People' } },
                    { key: 'gallery', text: strings.GalleryLayoutText, iconProps: { officeFabricIconFontName: 'PictureLibrary' } }
                  ]
                })
              ]
            },
            {
              groupName: strings.SortGroupName,
              groupFields: [
                PropertyPaneToggle('showSortDropdown', {
                  label: strings.ShowSortDropdownLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyFieldCollectionData('sortablePropertiesCollection', {
                  key: 'sortablePropertiesCollection',
                  label: strings.SortFieldLabel,
                  panelHeader: strings.SortPanelHeader,
                  manageBtnLabel: strings.SortManageBtn,
                  value: this.properties.sortablePropertiesCollection,
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
              groupName: strings.GeneralGroupName,
              groupFields: [
                PropertyPaneToggle('showResultCount', {
                  label: strings.ShowResultCountLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                })
              ]
            }
          ]
        },
        // ─── Page 3: About ────────────────────────────────
        {
          header: {
            description: strings.AboutPageHeader
          },
          groups: [
            {
              groupName: strings.AboutGroupName,
              groupFields: [
                PropertyPaneLabel('', {
                  text: 'SP Search Results v1.0 — Enterprise SharePoint search with 6 layouts, bulk actions, and URL deep linking.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
