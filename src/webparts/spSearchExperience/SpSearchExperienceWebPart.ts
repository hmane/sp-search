import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import type { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { StoreApi } from 'zustand/vanilla';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { configureLegacyPnPBaseUrl } from 'spfx-toolkit/lib/utilities/context/urlSanitizer';
import { SPDebugProvider } from 'spfx-toolkit/lib/components/debug';
import { spfxToolkitStylesLoaded } from '../../styles/loadSpfxToolkitStyles';

import * as strings from 'SpSearchExperienceWebPartStrings';
import * as resultsStrings from 'SpSearchResultsWebPartStrings';
import * as filtersStrings from 'SpSearchFiltersWebPartStrings';
import SpSearchExperience from './components/SpSearchExperience';
import type { FiltersPlacement, ISpSearchExperienceProps } from './components/ISpSearchExperienceProps';
import SpSearchResults from '../spSearchResults/components/SpSearchResults';
import type { ISelectedPropertyColumn, ISpSearchResultsProps } from '../spSearchResults/components/ISpSearchResultsProps';
import SpSearchFilters from '../spSearchFilters/components/SpSearchFilters';
import type { ISpSearchFiltersProps } from '../spSearchFilters/components/ISpSearchFiltersProps';
import {
  IColumnConfigItem,
  IColumnPropertyOption,
  ILegacyColumnItem,
  normalizeColumnConfigItem,
} from '../spSearchResults/components/ColumnConfigField/columnConfig';
import { PropertyPaneColumnConfigField } from '../spSearchResults/components/ColumnConfigField/ColumnConfigField';
import { normalizeSelectedPropertyItems } from '../spSearchResults/components/selectedPropertiesConfig';
import { TitleDisplayMode } from '../spSearchResults/components/documentTitleUtils';
import { GraphOrgService } from '../spSearchResults/components/GraphOrgService';
import { registerBuiltInActions } from '../spSearchResults/registerBuiltInActions';
import { registerBuiltInFilterTypes } from '../spSearchFilters/registerBuiltInFilterTypes';
import { SCENARIO_PRESETS } from '../spSearchResults/presets/searchPresets';
import {
  PropertyPaneCollectionData,
  CustomCollectionFieldType,
} from '../../propertyPaneControls/collectionData/PropertyPaneCollectionData';
import { PropertyPaneFiltersCollection } from '../../propertyPaneControls/filtersCollection/PropertyPaneFiltersCollection';
import { PropertyPaneSchemaHelper } from '../../propertyPaneControls/PropertyPaneSchemaHelper';
import {
  propertyPaneSearchContextIdField,
  SEARCH_CONTEXT_ID_GROUP_NAME,
} from '../../propertyPaneControls/PropertyPaneSearchContextIdField';
import { propertyPaneGroupHelp } from '../../propertyPaneControls/propertyPaneGroupHelp';
import { AudienceGate, parseAudienceGroups } from '../../utilities/AudienceGate';
import { SearchContextIdBannerWrapper } from '../../utilities/SearchContextIdMismatchBanner';
import { DebugCollector } from '@store/debug';
import type { IFilterConfig, ISearchStore } from '@interfaces/index';
import { SharePointSearchProvider, GraphSearchProvider } from '@providers/index';
import {
  getOrchestrator,
  getStore,
  initializeSearchContext,
  incrementContextRef,
  decrementContextRef,
} from '@store/store';
import { SearchOrchestrator } from '@orchestrator/SearchOrchestrator';
import { recordFirstSearch, recordWebPartInit } from '@store/utils/initOrderDiagnostic';
import { recordPresetSuggestion } from '@store/utils/presetSuggestionRegistry';
import { sanitizeUrlAlias } from '@store/utils/filterUrlAliases';
import { spLog } from '@store/utils/spLog';

// Bundle DevExtreme CSS for the embedded Filters component.
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('devextreme/dist/css/dx.common.css');
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('devextreme/dist/css/dx.light.css');
// eslint-disable-next-line @typescript-eslint/no-unused-vars
const _ensureStyles = spfxToolkitStylesLoaded;

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

interface IFilterCollectionItem {
  uniqueId: string;
  managedProperty: string;
  displayName: string;
  urlAlias?: string;
  filterType: string;
  operator: string;
  maxValues: number;
  defaultExpanded: boolean;
  showCount: boolean;
  sortBy: string;
  sortDirection: string;
  multiValues: boolean;
  dependsOn?: string;
  showWhenParentHasValue?: boolean;
  hideZeroCountValues?: boolean;
  resetWhenParentChanges?: boolean;
  trueLabel?: string;
  falseLabel?: string;
  invertBoolean?: boolean;
  defaultValue?: boolean;
  dataType?: 'auto' | 'text' | 'choiceMulti' | 'lookup' | 'calculated' |
             'datetime' | 'yesno' | 'number';
  valueSplitDelimiter?: string;
  audience?: string;
}

export interface ISpSearchExperienceWebPartProps {
  searchContextId: string;
  filtersPlacement: FiltersPlacement;
  filtersWidth: number;
  queryTemplate: string;
  selectedPropertiesCollection: ISelectedPropertyItem[];
  compactPropertiesCollection: IColumnConfigItem[];
  gridPropertiesCollection: IColumnConfigItem[];
  showColumnChooser: boolean;
  resultSourceId: string;
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
  emptyResultsMessage: string;
  titleDisplayMode: TitleDisplayMode;
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
  layoutPreset: string;
  filtersCollection: IFilterCollectionItem[];
  applyMode: 'instant' | 'manual';
  operatorBetweenFilters: 'AND' | 'OR';
  showClearAll: boolean;
  enableVisualFilterBuilder: boolean;
  audienceGroups: string;
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

function normalizeFiltersCollectionValue(raw: IFilterCollectionItem[] | string | undefined): IFilterCollectionItem[] {
  return normalizeCollectionValue<IFilterCollectionItem>(raw);
}

function isTitleProperty(property: string): boolean {
  const normalized = (property || '').trim().toLowerCase();
  return normalized === 'title' || normalized === 'filename';
}

function normalizeManagedPropertyForFilter(
  managedProperty: string,
  filterType: IFilterConfig['filterType']
): string {
  if (filterType === 'people' && managedProperty === 'Author') {
    return 'AuthorOWSUSER';
  }

  return managedProperty;
}

const DATA_FORMAT_FILTER_TYPES: ReadonlyArray<string> = [
  'checkbox',
  'tagbox',
  'dropdown',
  'text',
];

function isDataFormatRelevant(filterType: string): boolean {
  return DATA_FORMAT_FILTER_TYPES.indexOf(filterType) !== -1;
}

function normalizeDataType(
  raw: IFilterCollectionItem['dataType'] | undefined,
  filterType: IFilterConfig['filterType']
): IFilterConfig['dataType'] | undefined {
  if (!isDataFormatRelevant(filterType)) {
    return undefined;
  }
  if (!raw || raw === 'auto') {
    return undefined;
  }
  return raw;
}

function normalizeSplitDelimiter(
  raw: string | undefined,
  filterType: IFilterConfig['filterType']
): string | undefined {
  if (!isDataFormatRelevant(filterType)) {
    return undefined;
  }
  if (typeof raw !== 'string' || raw.length === 0) {
    return undefined;
  }
  return raw;
}

export default class SpSearchExperienceWebPart extends BaseClientSideWebPart<ISpSearchExperienceWebPartProps> {
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

    const contextId = this.properties.searchContextId || 'default';
    const selectedPropertyColumns = this._getSelectedPropertyColumns();
    const compactPropertyColumns = this._getCompactPropertyColumns();
    const gridPropertyColumns = this._getGridPropertyColumns();
    const resultsElement: React.ReactElement<ISpSearchResultsProps> = React.createElement(
      SpSearchResults,
      {
        store: this._store,
        orchestrator: this._orchestrator,
        searchContextId: contextId,
        siteUrl: this._getWebAbsoluteUrl(),
        theme: this._theme,
        showResultCount: this.properties.showResultCount !== false,
        showSortDropdown: this.properties.showSortDropdown !== false,
        showDeleteConfirmation: this.properties.showDeleteConfirmation !== false,
        enablePreviewPanel: this.properties.enablePreviewPanel !== false,
        hideWebPartWhenNoResults: this.properties.hideWebPartWhenNoResults === true,
        emptyResultsMessage: this.properties.emptyResultsMessage || '',
        showColumnChooser: this.properties.showColumnChooser !== false,
        titleDisplayMode: this.properties.titleDisplayMode || 'wrap',
        defaultLayout: this.properties.defaultLayout || 'list',
        pageSize: this.properties.pageSize || 10,
        isEditMode: this.displayMode === DisplayMode.Edit,
        selectedPropertyColumns,
        gridPropertyColumns,
        compactPropertyColumns,
        queryTemplate: this.properties.queryTemplate || '{searchTerms}',
        graphOrgService: this._graphOrgService,
        linkConfig: {
          clickTarget: this.properties.resultClickTarget || 'panel',
          documentLinkMode: this.properties.documentLinkMode || 'file',
          listItemLinkMode: this.properties.listItemLinkMode || 'displayForm',
        },
      }
    );

    const filtersElement: React.ReactElement<ISpSearchFiltersProps> = React.createElement(
      SpSearchFilters,
      {
        store: this._store,
        applyMode: this.properties.applyMode || 'instant',
        showClearAll: this.properties.showClearAll !== false,
        enableVisualFilterBuilder: !!this.properties.enableVisualFilterBuilder,
        isEditMode: this.displayMode === DisplayMode.Edit,
        searchContextId: contextId,
        onApplyPresetFilters: (filterRows): void => {
          const existing = normalizeFiltersCollectionValue(this.properties.filtersCollection);
          const haveProps = new Set(existing.map((r) => (r.managedProperty || '').toLowerCase()));
          const additions: IFilterCollectionItem[] = filterRows
            .filter((r) => !haveProps.has(r.managedProperty.toLowerCase()))
            .map((r, idx): IFilterCollectionItem => ({
              uniqueId: 'preset-filter-' + String(Date.now()) + '-' + String(idx),
              managedProperty: r.managedProperty,
              displayName: r.label,
              urlAlias: r.urlAlias,
              filterType: r.filterType,
              operator: '=',
              maxValues: 10,
              defaultExpanded: true,
              showCount: true,
              sortBy: 'count',
              sortDirection: 'desc',
              multiValues: true,
            }));
          if (additions.length === 0) {
            return;
          }
          this.properties.filtersCollection = existing.concat(additions);
          this._syncFilterConfigToStore();
          this.context.propertyPane.refresh();
          this.render();
        },
      }
    );

    const experienceElement: React.ReactElement<ISpSearchExperienceProps> = React.createElement(
      SpSearchExperience,
      {
        resultsElement,
        filtersElement,
        filtersPlacement: this.properties.filtersPlacement || 'right',
        filtersWidth: this.properties.filtersWidth || 360,
      }
    );

    const audienceGroups = parseAudienceGroups(this.properties.audienceGroups);
    const gatedElement: React.ReactElement = React.createElement(
      AudienceGate,
      { audienceGroups, store: this._store },
      experienceElement
    );

    const bannerWrapped: React.ReactElement = React.createElement(
      SearchContextIdBannerWrapper,
      {
        webPartId: this.instanceId,
        contextId,
        webPartLabel: 'SP Search Results + Filters',
        isEditMode: this.displayMode === DisplayMode.Edit,
      },
      gatedElement
    );

    const wrappedElement: React.ReactElement = React.createElement(
      SPDebugProvider,
      { logger: SPContext.logger, allowInProduction: false },
      bannerWrapped
    );

    ReactDom.render(wrappedElement, this.domElement);
  }

  protected async onInit(): Promise<void> {
    try {
      await SPContext.basic(this.context as unknown as Parameters<typeof SPContext.basic>[0], 'SPSearchExperience');
      configureLegacyPnPBaseUrl(this.context);

      const contextId = this.properties.searchContextId || 'default';
      this._store = getStore(contextId);
      this._orchestrator = getOrchestrator(contextId);
      incrementContextRef(contextId);

      if (this._store) {
        const provider = new SharePointSearchProvider();
        const dataProviders = this._store.getState().registries.dataProviders;
        if (!dataProviders.get(provider.id)) {
          dataProviders.register(provider);
        }

        try {
          const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
          this._graphOrgService = new GraphOrgService(graphClient);
          const graphProvider = new GraphSearchProvider(graphClient);
          if (!dataProviders.get(graphProvider.id)) {
            dataProviders.register(graphProvider);
          }

          const graphPeopleProvider = new GraphSearchProvider(graphClient, {
            id: 'graph-people',
            entityTypes: ['person'],
          });
          if (!dataProviders.get(graphPeopleProvider.id)) {
            dataProviders.register(graphPeopleProvider);
          }
        } catch {
          // Graph client unavailable — SharePoint Search remains the fallback.
        }

        registerBuiltInActions(this._store.getState().registries.actions);
        registerBuiltInFilterTypes(this._store.getState().registries.filterTypes);
      }

      if (!this.properties.layoutPreset) {
        this.properties.layoutPreset = 'custom';
      }
      this._normalizeDefaultLayoutProperty();

      this._syncFilterConfigToStore();
      this._syncOperatorToStore();
      this._syncQueryTemplateToStore();
      this._syncSortablePropertiesToStore();
      this._syncSelectedPropertiesToStore();
      this._syncRefinementFiltersToStore();
      this._syncSearchConfigToStore();
      this._syncScopeToStore();
      this._syncAvailableLayoutsToStore();
      this._syncDefaultLayoutAndPageSizeToStore();

      recordWebPartInit(contextId, 'SpSearchFiltersWebPart');
      recordWebPartInit(contextId, 'SpSearchResultsWebPart');
      recordWebPartInit(contextId, 'SpSearchExperienceWebPart');

      await initializeSearchContext(contextId, this.context);

      if (this._orchestrator) {
        const filterConfigLength = this._store
          ? (this._store.getState().filterConfig || []).length
          : 0;
        recordFirstSearch(contextId, filterConfigLength);
        this._orchestrator.triggerSearch().catch(function noop(): void { /* handled in orchestrator */ });
      }

      DebugCollector.registerWebPart('SPSearchExperienceWebPart', this.properties as unknown as Record<string, unknown>);
    } catch (err) {
      spLog.error('SPSearchExperience onInit failed', { error: err });
      throw err;
    }
  }

  private _getSelectedPropertyColumns(): ISelectedPropertyColumn[] {
    const raw = this._normalizeSelectedPropertiesCollection();

    return raw.map((item: ISelectedPropertyItem): ISelectedPropertyColumn => ({
      property: item.property || '',
      alias: item.alias || '',
    }));
  }

  private _normalizeSelectedPropertiesCollection(): ISelectedPropertyItem[] {
    const raw = normalizeCollectionValue<ISelectedPropertyItem>(this.properties.selectedPropertiesCollection);
    const normalized = normalizeSelectedPropertyItems(raw, true);
    this.properties.selectedPropertiesCollection = normalized;
    return normalized;
  }

  private _getLayoutPropertyOptions(): IColumnPropertyOption[] {
    return this._normalizeSelectedPropertiesCollection()
      .filter((item: ISelectedPropertyItem) => !isTitleProperty(item.property))
      .map((item: ISelectedPropertyItem) => ({
        key: item.property,
        text: item.alias ? (item.alias + ' (' + item.property + ')') : item.property,
        alias: item.alias || item.property,
      }));
  }

  private _getCompactPropertyColumns(): IColumnConfigItem[] {
    return this._getCompactLayoutProperties();
  }

  private _getGridPropertyColumns(): IColumnConfigItem[] {
    return this._getGridLayoutProperties();
  }

  private _getCompactLayoutProperties(): IColumnConfigItem[] {
    const raw = normalizeCollectionValue<Partial<IColumnConfigItem> & ILegacyColumnItem>(
      this.properties.compactPropertiesCollection as unknown as Array<Partial<IColumnConfigItem> & ILegacyColumnItem>
    );
    const masterAliasMap = this._getMasterAliasMap();
    const result = this._normalizeColumnCollection(raw, masterAliasMap);

    if (result.length === 0) {
      const defaults = ['Author', 'LastModifiedTime', 'Size', 'FileType'];
      for (let i = 0; i < defaults.length; i++) {
        const masterEntry = masterAliasMap.get(defaults[i].toLowerCase());
        if (masterEntry) {
          result.push(normalizeColumnConfigItem({
            uniqueId: 'experience-compact-fallback-' + String(i),
            property: masterEntry.property,
            alias: masterEntry.alias,
          }));
        }
      }
    }

    this.properties.compactPropertiesCollection = result;
    return result;
  }

  private _getGridLayoutProperties(): IColumnConfigItem[] {
    const raw = normalizeCollectionValue<Partial<IColumnConfigItem> & ILegacyColumnItem>(
      this.properties.gridPropertiesCollection as unknown as Array<Partial<IColumnConfigItem> & ILegacyColumnItem>
    );
    const masterAliasMap = this._getMasterAliasMap();
    const result = this._normalizeColumnCollection(raw, masterAliasMap);

    if (result.length === 0) {
      let idx = 0;
      masterAliasMap.forEach((entry) => {
        result.push(normalizeColumnConfigItem({
          uniqueId: 'experience-grid-fallback-' + String(idx),
          property: entry.property,
          alias: entry.alias,
        }));
        idx++;
      });
    }

    this.properties.gridPropertiesCollection = result;
    return result;
  }

  private _getMasterAliasMap(): Map<string, { property: string; alias: string }> {
    const masterAliasMap = new Map<string, { property: string; alias: string }>();
    const master = this._normalizeSelectedPropertiesCollection();
    for (let i = 0; i < master.length; i++) {
      const property = master[i].property;
      if (!isTitleProperty(property)) {
        masterAliasMap.set(property.toLowerCase(), {
          property,
          alias: master[i].alias || property,
        });
      }
    }
    return masterAliasMap;
  }

  private _normalizeColumnCollection(
    raw: Array<Partial<IColumnConfigItem> & ILegacyColumnItem>,
    masterAliasMap: Map<string, { property: string; alias: string }>
  ): IColumnConfigItem[] {
    const result: IColumnConfigItem[] = [];
    const seen = new Set<string>();

    for (let i = 0; i < raw.length; i++) {
      const property = String(raw[i].property || '').trim();
      const lookup = property.toLowerCase();
      if (!property || isTitleProperty(property) || seen.has(lookup) || !masterAliasMap.has(lookup)) {
        continue;
      }
      seen.add(lookup);

      const normalized = normalizeColumnConfigItem(raw[i]);
      const aliasOnRaw = typeof raw[i].alias === 'string' ? String(raw[i].alias).trim() : '';
      if (!aliasOnRaw) {
        const masterEntry = masterAliasMap.get(lookup);
        if (masterEntry && masterEntry.alias && masterEntry.alias !== property) {
          normalized.alias = masterEntry.alias;
        }
      }
      result.push(normalized);
    }

    return result;
  }

  private _syncFilterConfigToStore(): void {
    if (!this._store) {
      return;
    }

    const rawFilters = normalizeFiltersCollectionValue(this.properties.filtersCollection);
    if (!rawFilters || rawFilters.length === 0) {
      const state = this._store.getState();
      if (state.filterConfig.length > 0) {
        this._store.setState({ filterConfig: [], activeFilters: [] });
      }
      return;
    }

    const filterConfig: IFilterConfig[] = rawFilters.map((item: IFilterCollectionItem) => {
      const filterType = (item.filterType || 'checkbox') as IFilterConfig['filterType'];

      return {
        id: item.uniqueId,
        managedProperty: normalizeManagedPropertyForFilter(item.managedProperty, filterType),
        displayName: item.displayName,
        urlAlias: sanitizeUrlAlias(item.urlAlias),
        filterType,
        operator: (item.operator || 'OR') as IFilterConfig['operator'],
        maxValues: item.maxValues || 10,
        defaultExpanded: item.defaultExpanded !== false,
        showCount: item.showCount !== false,
        sortBy: (item.sortBy || 'count') as IFilterConfig['sortBy'],
        sortDirection: (item.sortDirection || 'desc') as IFilterConfig['sortDirection'],
        multiValues: item.multiValues !== false,
        dependsOn: item.dependsOn || undefined,
        showWhenParentHasValue: item.showWhenParentHasValue === true,
        hideZeroCountValues: item.hideZeroCountValues === true,
        resetWhenParentChanges: item.resetWhenParentChanges === true,
        trueLabel: item.trueLabel || undefined,
        falseLabel: item.falseLabel || undefined,
        invertBoolean: item.invertBoolean === true,
        defaultValue: filterType === 'toggle' && typeof item.defaultValue === 'boolean'
          ? item.defaultValue
          : undefined,
        dataType: normalizeDataType(item.dataType, filterType),
        valueSplitDelimiter: normalizeSplitDelimiter(item.valueSplitDelimiter, filterType),
        audienceGroups: item.audience
          ? item.audience.split(',').map((s: string) => s.trim()).filter(Boolean)
          : undefined,
      };
    });

    const state = this._store.getState();
    if (JSON.stringify(state.filterConfig) !== JSON.stringify(filterConfig)) {
      this._store.setState({ filterConfig });
    }
  }

  private _syncOperatorToStore(): void {
    if (!this._store) {
      return;
    }
    const operator: 'AND' | 'OR' = this.properties.operatorBetweenFilters === 'OR' ? 'OR' : 'AND';
    const state = this._store.getState();
    if (state.operatorBetweenFilters !== operator) {
      state.setOperatorBetweenFilters(operator);
    }
  }

  private _syncQueryTemplateToStore(): void {
    if (!this._store) {
      return;
    }
    const template = this.properties.queryTemplate || '{searchTerms}';
    const state = this._store.getState();
    if (state.queryTemplate !== template) {
      this._store.setState({ queryTemplate: template });
    }
  }

  private _syncSortablePropertiesToStore(): void {
    if (!this._store) {
      return;
    }

    const raw = normalizeCollectionValue<ISortCollectionItem>(this.properties.sortablePropertiesCollection);
    const sortableProperties = raw.map((item: ISortCollectionItem) => ({
      property: item.property,
      label: item.label,
      direction: item.direction || 'Ascending',
    }));
    const state = this._store.getState();
    if (JSON.stringify(state.sortableProperties) !== JSON.stringify(sortableProperties)) {
      this._store.setState({ sortableProperties });
    }
  }

  private _syncSelectedPropertiesToStore(): void {
    if (!this._store) {
      return;
    }

    const raw = this._normalizeSelectedPropertiesCollection();
    const propertiesString = raw
      .map((item: ISelectedPropertyItem) => item.property)
      .filter(Boolean)
      .join(',');
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
    const filtersString = raw
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
            return item.property + ':"' + item.value + '"';
        }
      })
      .filter(Boolean)
      .join(',');

    const state = this._store.getState();
    if (state.refinementFilters !== filtersString) {
      this._store.setState({ refinementFilters: filtersString });
    }
  }

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

    if (state.scope.id !== scopeId || state.scope.kqlPath !== kqlPath) {
      state.setScope({ id: scopeId, label: scopeLabel, kqlPath });
    }
  }

  private _syncDefaultLayoutAndPageSizeToStore(): void {
    if (!this._store) {
      return;
    }

    const state = this._store.getState();
    const defaultLayout = this.properties.defaultLayout || 'list';
    if (state.activeLayoutKey === 'list' && defaultLayout !== 'list') {
      state.setLayout(defaultLayout);
    }

    const pageSize = this.properties.pageSize || 10;
    if (state.pageSize !== pageSize) {
      this._store.setState({ pageSize });
    }
  }

  private _applyScenarioPreset(presetId: string): void {
    const preset = SCENARIO_PRESETS[presetId];
    if (!preset) {
      return;
    }

    this.properties.defaultLayout = preset.defaultLayout;
    this.properties.showListLayout = preset.showListLayout;
    this.properties.showCompactLayout = preset.showCompactLayout;
    this.properties.showGridLayout = preset.showGridLayout;
    this.properties.showCardLayout = preset.showCardLayout;
    this.properties.showPeopleLayout = preset.showPeopleLayout;
    this.properties.showGalleryLayout = preset.showGalleryLayout;
    this.properties.queryTemplate = preset.queryTemplate;
    this.properties.selectedPropertiesCollection = preset.selectedProperties.map(
      (p, idx) => ({ uniqueId: 'preset-sp-' + String(idx), property: p.property, alias: p.alias })
    );

    const presetAliasMap = new Map<string, string>(
      preset.selectedProperties.map((p) => [p.property.toLowerCase(), p.alias])
    );
    this.properties.compactPropertiesCollection = preset.compactProperties.map(
      (p, idx) => normalizeColumnConfigItem({
        uniqueId: 'preset-compact-' + String(idx),
        property: p.property,
        alias: presetAliasMap.get(p.property.toLowerCase()) || p.property,
      })
    );
    this.properties.gridPropertiesCollection = preset.selectedProperties
      .filter((p) => !isTitleProperty(p.property))
      .map((p, idx) => normalizeColumnConfigItem({
        uniqueId: 'preset-grid-' + String(idx),
        property: p.property,
        alias: p.alias,
      }));
    this.properties.sortablePropertiesCollection = preset.sortableProperties.map(
      (s, idx) => ({ uniqueId: 'preset-sort-' + String(idx), property: s.property, label: s.label, direction: s.direction })
    );

    recordPresetSuggestion(this.properties.searchContextId || 'default', {
      id: preset.id,
      label: preset.label,
      filterSuggestions: preset.filterSuggestions,
      recordedAt: 0,
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
      { key: 'list', text: resultsStrings.ListLayoutText, iconProps: { officeFabricIconFontName: 'List' } },
      { key: 'compact', text: resultsStrings.CompactLayoutText, iconProps: { officeFabricIconFontName: 'GridViewSmall' } },
      { key: 'grid', text: resultsStrings.GridLayoutText, iconProps: { officeFabricIconFontName: 'Table' } },
      { key: 'card', text: resultsStrings.CardLayoutText, iconProps: { officeFabricIconFontName: 'ContactCard' } },
      { key: 'people', text: resultsStrings.PeopleLayoutText, iconProps: { officeFabricIconFontName: 'People' } },
      { key: 'gallery', text: resultsStrings.GalleryLayoutText, iconProps: { officeFabricIconFontName: 'PictureLibrary' } },
    ];

    return allOptions.filter((option) => available.indexOf(option.key) >= 0);
  }

  private _buildPresetOptions(): Array<{ key: string; text: string; iconProps: { officeFabricIconFontName: string } }> {
    return [
      { key: 'custom', text: 'Custom', iconProps: { officeFabricIconFontName: 'Settings' } },
      { key: 'general', text: 'General', iconProps: { officeFabricIconFontName: 'Search' } },
      { key: 'documents', text: 'Documents', iconProps: { officeFabricIconFontName: 'DocLibrary' } },
      { key: 'hub-search', text: 'Hub Search', iconProps: { officeFabricIconFontName: 'Globe' } },
      { key: 'knowledge-base', text: 'Knowledge Base', iconProps: { officeFabricIconFontName: 'BookAnswers' } },
      { key: 'policy-search', text: 'Policy Search', iconProps: { officeFabricIconFontName: 'Shield' } },
      { key: 'news', text: 'News', iconProps: { officeFabricIconFontName: 'News' } },
      { key: 'media', text: 'Media', iconProps: { officeFabricIconFontName: 'Photo2' } },
      { key: 'people', text: 'People', iconProps: { officeFabricIconFontName: 'Group' } },
    ];
  }

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
    this._store.getState().setAvailableLayouts(this._getAvailableLayoutsFromProperties());
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
    const contextId = this.properties.searchContextId || 'default';
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
      const contextId = this.properties.searchContextId || 'default';
      this._store = getStore(contextId);
      this._orchestrator = getOrchestrator(contextId);
    }

    if (propertyPath === 'filtersCollection' || propertyPath === 'searchContextId') {
      this._syncFilterConfigToStore();
    }

    if (propertyPath === 'operatorBetweenFilters' || propertyPath === 'searchContextId') {
      this._syncOperatorToStore();
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

    if (propertyPath === 'compactPropertiesCollection' ||
      propertyPath === 'gridPropertiesCollection' ||
      propertyPath === 'titleDisplayMode' ||
      propertyPath === 'filtersPlacement' ||
      propertyPath === 'filtersWidth') {
      this.render();
    }

    if (propertyPath === 'refinementFiltersCollection' || propertyPath === 'searchContextId') {
      this._syncRefinementFiltersToStore();
    }

    if (propertyPath === 'pageSize' && this._store) {
      this._store.setState({ pageSize: this.properties.pageSize || 10, currentPage: 1 });
    }

    if (propertyPath === 'defaultLayout' && this._store) {
      if (this.properties.layoutPreset !== 'custom') {
        this.properties.layoutPreset = 'custom';
        this.context.propertyPane.refresh();
      }
      this._normalizeDefaultLayoutProperty();
      this._store.getState().setLayout(this.properties.defaultLayout || 'list');
    }

    if (propertyPath === 'layoutPreset') {
      this._applyScenarioPreset(this.properties.layoutPreset);
      this._normalizeDefaultLayoutProperty();
      this._syncAvailableLayoutsToStore();
      if (this._store) {
        this._store.getState().setLayout(this.properties.defaultLayout || 'list');
      }
      this.context.propertyPane.refresh();
    }

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
        {
          header: {
            description: strings.ExperiencePageHeader,
          },
          groups: [
            {
              groupName: SEARCH_CONTEXT_ID_GROUP_NAME,
              groupFields: [
                propertyPaneSearchContextIdField(),
              ],
            },
            {
              groupName: strings.ExperienceLayoutGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('filtersPlacement', {
                  label: strings.FiltersPlacementLabel,
                  options: [
                    { key: 'right', text: strings.FiltersPlacementRight, iconProps: { officeFabricIconFontName: 'AlignRight' } },
                    { key: 'left', text: strings.FiltersPlacementLeft, iconProps: { officeFabricIconFontName: 'AlignLeft' } },
                    { key: 'top', text: strings.FiltersPlacementTop, iconProps: { officeFabricIconFontName: 'AlignCenter' } },
                  ],
                }),
                PropertyPaneSlider('filtersWidth', {
                  label: strings.FiltersWidthLabel,
                  min: 260,
                  max: 520,
                  value: this.properties.filtersWidth || 360,
                  step: 20,
                  showValue: true,
                }),
              ],
            },
            {
              groupName: resultsStrings.GetStartedGroupName,
              groupFields: [
                propertyPaneGroupHelp('quick-start', 'Help: Quick Start presets'),
                PropertyPaneChoiceGroup('layoutPreset', {
                  label: resultsStrings.ScenarioPresetLabel,
                  options: this._buildPresetOptions(),
                }),
                PropertyPaneLabel('layoutPresetHint', {
                  text: strings.ScenarioPresetHint,
                }),
                ...(this._isPeoplePresetBlocked() ? [
                  PropertyPaneLabel('layoutPresetPeopleWarning', {
                    text: resultsStrings.ScenarioPresetPeopleWarning,
                  }),
                ] : []),
              ],
            },
          ],
        },
        {
          header: {
            description: strings.ResultsPageHeader,
          },
          groups: [
            {
              groupName: resultsStrings.DataGroupName,
              groupFields: [
                propertyPaneGroupHelp('results-data', 'Help: Search scope and managed properties'),
                PropertyPaneDropdown('searchScope', {
                  label: resultsStrings.SearchScopeLabel,
                  options: [
                    { key: 'all', text: resultsStrings.ScopeAllText },
                    { key: 'currentsite', text: resultsStrings.ScopeCurrentSiteText },
                    { key: 'currentcollection', text: resultsStrings.ScopeCurrentCollectionText },
                    { key: 'custom', text: resultsStrings.ScopeCustomText },
                  ],
                  selectedKey: this.properties.searchScope || 'all',
                }),
                ...(this.properties.searchScope === 'custom' ? [
                  PropertyPaneTextField('searchScopePath', {
                    label: resultsStrings.SearchScopePathLabel,
                    description: resultsStrings.SearchScopePathDescription,
                    placeholder: 'https://contoso.sharepoint.com/sites/hr',
                  }),
                ] : []),
                PropertyPaneSchemaHelper('queryTemplate', {
                  label: resultsStrings.QueryTemplateLabel,
                  description: resultsStrings.QueryTemplateDescription,
                  value: this.properties.queryTemplate || '',
                  filterHint: 'queryable',
                  applyOnEnter: true,
                }),
                PropertyPaneTextField('resultSourceId', {
                  label: resultsStrings.ResultSourceIdLabel,
                  description: resultsStrings.ResultSourceIdDescription,
                }),
                PropertyPaneCollectionData('selectedPropertiesCollection', {
                  key: 'selectedPropertiesCollection',
                  label: resultsStrings.SelectedPropertiesLabel,
                  panelHeader: resultsStrings.SelectedPropertiesPanelHeader,
                  manageBtnLabel: resultsStrings.SelectedPropertiesManageBtn,
                  value: this._normalizeSelectedPropertiesCollection(),
                  enableSorting: true,
                  fields: [
                    {
                      id: 'property',
                      title: resultsStrings.SelectedPropertyColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'LastModifiedTime',
                    },
                    {
                      id: 'alias',
                      title: resultsStrings.SelectedPropertyAliasColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'Date Modified',
                    },
                  ],
                }),
                PropertyPaneLabel('selectedPropertiesAliasDeprecation', {
                  text: resultsStrings.SelectedPropertiesAliasDeprecationNote,
                }),
              ],
            },
            {
              groupName: resultsStrings.SortGroupName,
              groupFields: [
                PropertyPaneCollectionData('sortablePropertiesCollection', {
                  key: 'sortablePropertiesCollection',
                  label: resultsStrings.SortFieldLabel,
                  panelHeader: resultsStrings.SortPanelHeader,
                  manageBtnLabel: resultsStrings.SortManageBtn,
                  value: normalizeCollectionValue<ISortCollectionItem>(this.properties.sortablePropertiesCollection),
                  enableSorting: true,
                  fields: [
                    {
                      id: 'property',
                      title: resultsStrings.SortPropertyColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'LastModifiedTime',
                    },
                    {
                      id: 'label',
                      title: resultsStrings.SortLabelColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'Date Modified',
                    },
                    {
                      id: 'direction',
                      title: resultsStrings.SortDirectionColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: [
                        { key: 'Ascending', text: resultsStrings.SortAscending },
                        { key: 'Descending', text: resultsStrings.SortDescending },
                      ],
                    },
                  ],
                }),
              ],
            },
            {
              groupName: resultsStrings.PaginationGroupName,
              groupFields: [
                PropertyPaneSlider('pageSize', {
                  label: resultsStrings.PageSizeLabel,
                  min: 5,
                  max: 100,
                  value: this.properties.pageSize || 10,
                  step: 5,
                  showValue: true,
                }),
                PropertyPaneToggle('showPaging', {
                  label: resultsStrings.ShowPagingLabel,
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
                PropertyPaneSlider('pageRange', {
                  label: resultsStrings.PageRangeLabel,
                  min: 3,
                  max: 10,
                  value: this.properties.pageRange || 5,
                  step: 1,
                  showValue: true,
                }),
              ],
            },
          ],
        },
        {
          header: {
            description: resultsStrings.DisplayPageHeader,
          },
          groups: [
            {
              groupName: resultsStrings.MainLayoutsGroupName,
              groupFields: [
                propertyPaneGroupHelp('results-layouts', 'Help: Layouts and presets'),
                PropertyPaneChoiceGroup('defaultLayout', {
                  label: resultsStrings.DefaultLayoutLabel,
                  options: this._buildDefaultLayoutOptions(),
                }),
                PropertyPaneToggle('showListLayout', {
                  label: 'Show List view',
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
                PropertyPaneToggle('showCompactLayout', {
                  label: 'Show Compact List view',
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
                PropertyPaneToggle('showGridLayout', {
                  label: 'Show Data Grid view',
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
              ],
            },
            ...(this._shouldShowSpecializedViews() ? [{
              groupName: resultsStrings.AdvancedLayoutsGroupName,
              groupFields: [
                PropertyPaneToggle('showCardLayout', {
                  label: 'Show Card view',
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
                PropertyPaneToggle('showPeopleLayout', {
                  label: 'Show People view',
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
                PropertyPaneToggle('showGalleryLayout', {
                  label: 'Show Gallery view',
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
              ],
            }] : []),
            ...(this.properties.showCompactLayout !== false ? [{
              groupName: resultsStrings.CompactViewGroupName,
              groupFields: [
                PropertyPaneColumnConfigField('compactPropertiesCollection', {
                  label: resultsStrings.CompactPropertiesLabel,
                  value: this._getCompactLayoutProperties(),
                  availableProperties: this._getLayoutPropertyOptions(),
                }),
              ],
            }] : []),
            ...(this.properties.showGridLayout !== false ? [{
              groupName: resultsStrings.GridViewGroupName,
              groupFields: [
                PropertyPaneColumnConfigField('gridPropertiesCollection', {
                  label: resultsStrings.GridPropertiesLabel,
                  value: this._getGridLayoutProperties(),
                  availableProperties: this._getLayoutPropertyOptions(),
                }),
                PropertyPaneToggle('showColumnChooser', {
                  label: resultsStrings.ShowColumnChooserLabel,
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
              ],
            }] : []),
            {
              groupName: resultsStrings.BehaviorGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('titleDisplayMode', {
                  label: resultsStrings.TitleDisplayModeLabel,
                  options: [
                    { key: 'ellipsis', text: resultsStrings.TitleDisplayEllipsisText },
                    { key: 'middle', text: resultsStrings.TitleDisplayMiddleText },
                    { key: 'wrap', text: resultsStrings.TitleDisplayWrapText },
                  ],
                }),
                PropertyPaneToggle('showSortDropdown', {
                  label: resultsStrings.ShowSortDropdownLabel,
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
                PropertyPaneToggle('showResultCount', {
                  label: resultsStrings.ShowResultCountLabel,
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
                PropertyPaneToggle('enablePreviewPanel', {
                  label: resultsStrings.ShowPreviewPanelLabel,
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
                PropertyPaneToggle('hideWebPartWhenNoResults', {
                  label: resultsStrings.HideWebPartWhenNoResultsLabel,
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
                PropertyPaneTextField('emptyResultsMessage', {
                  label: resultsStrings.EmptyResultsMessageLabel,
                  description: resultsStrings.EmptyResultsMessageDescription,
                  multiline: true,
                  rows: 4,
                  resizable: true,
                }),
              ],
            },
            {
              groupName: resultsStrings.ResultLinkGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('resultClickTarget', {
                  label: resultsStrings.ResultClickTargetLabel,
                  options: [
                    { key: 'panel', text: resultsStrings.ResultClickTargetPanelText },
                    { key: 'newTab', text: resultsStrings.ResultClickTargetNewTabText },
                    { key: 'sameTab', text: resultsStrings.ResultClickTargetSameTabText },
                    { key: 'sidePanel', text: resultsStrings.ResultClickTargetSidePanelText },
                  ],
                }),
                ...((this.properties.resultClickTarget || 'panel') !== 'sidePanel' ? [
                  PropertyPaneDropdown('documentLinkMode', {
                    label: resultsStrings.DocumentLinkModeLabel,
                    options: [
                      { key: 'file', text: resultsStrings.DocumentLinkModeFileText },
                      { key: 'propertiesForm', text: resultsStrings.DocumentLinkModePropertiesFormText },
                    ],
                  }),
                  PropertyPaneDropdown('listItemLinkMode', {
                    label: resultsStrings.ListItemLinkModeLabel,
                    options: [
                      { key: 'displayForm', text: resultsStrings.ListItemLinkModeDisplayFormText },
                      { key: 'editForm', text: resultsStrings.ListItemLinkModeEditFormText },
                    ],
                  }),
                ] : []),
              ],
            },
          ],
        },
        {
          header: {
            description: strings.FiltersPageHeader,
          },
          groups: [
            {
              groupName: filtersStrings.FiltersGroupName,
              groupFields: [
                propertyPaneGroupHelp('filters-config', 'Help: Configure refiners and filter types'),
                PropertyPaneFiltersCollection('filtersCollection', {
                  label: filtersStrings.FiltersFieldLabel,
                  panelHeader: filtersStrings.FiltersPanelHeader,
                  manageButtonLabel: filtersStrings.FiltersManageBtn,
                  value: normalizeFiltersCollectionValue(this.properties.filtersCollection),
                }),
              ],
            },
            {
              groupName: filtersStrings.BehaviorGroupName,
              groupFields: [
                propertyPaneGroupHelp('filters-behavior', 'Help: Apply mode and Clear All behaviour'),
                PropertyPaneChoiceGroup('applyMode', {
                  label: filtersStrings.ApplyModeLabel,
                  options: [
                    { key: 'instant', text: filtersStrings.ApplyModeInstant },
                    { key: 'manual', text: filtersStrings.ApplyModeManual },
                  ],
                }),
                PropertyPaneChoiceGroup('operatorBetweenFilters', {
                  label: filtersStrings.OperatorLabel,
                  options: [
                    { key: 'AND', text: 'AND' },
                    { key: 'OR', text: 'OR' },
                  ],
                }),
                PropertyPaneToggle('showClearAll', {
                  label: filtersStrings.ShowClearAllLabel,
                  onText: filtersStrings.ToggleOnText,
                  offText: filtersStrings.ToggleOffText,
                }),
                PropertyPaneToggle('enableVisualFilterBuilder', {
                  label: filtersStrings.EnableVisualFilterBuilderLabel,
                  onText: filtersStrings.ToggleOnText,
                  offText: filtersStrings.ToggleOffText,
                }),
              ],
            },
          ],
        },
        {
          header: {
            description: strings.AudiencePageHeader,
          },
          groups: [
            {
              groupName: resultsStrings.AudienceTargetingGroupName,
              groupFields: [
                PropertyPaneTextField('audienceGroups', {
                  label: resultsStrings.AudienceTargetingLabel,
                  description: resultsStrings.AudienceTargetingDescription,
                  multiline: true,
                  rows: 3,
                  resizable: true,
                }),
              ],
            },
          ],
        },
        {
          header: {
            description: strings.AdvancedPageHeader,
          },
          groups: [
            {
              groupName: resultsStrings.QuerySettingsGroupName,
              groupFields: [
                PropertyPaneToggle('enableQueryRules', {
                  label: resultsStrings.EnableQueryRulesLabel,
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
                PropertyPaneToggle('trimDuplicates', {
                  label: resultsStrings.TrimDuplicatesLabel,
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
              ],
            },
            {
              groupName: resultsStrings.AdvancedGroupName,
              groupFields: [
                PropertyPaneCollectionData('refinementFiltersCollection', {
                  key: 'refinementFiltersCollection',
                  label: resultsStrings.RefinementFiltersLabel,
                  panelHeader: resultsStrings.RefinementFiltersPanelHeader,
                  manageBtnLabel: resultsStrings.RefinementFiltersManageBtn,
                  value: normalizeCollectionValue<IRefinementFilterItem>(this.properties.refinementFiltersCollection),
                  enableSorting: false,
                  fields: [
                    {
                      id: 'property',
                      title: resultsStrings.RefinementFilterPropertyColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'FileType',
                    },
                    {
                      id: 'operator',
                      title: resultsStrings.RefinementFilterOperatorColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: [
                        { key: 'equals', text: resultsStrings.FqlOperatorEquals },
                        { key: 'contains', text: resultsStrings.FqlOperatorContains },
                        { key: 'range', text: resultsStrings.FqlOperatorRange },
                        { key: 'beginsWith', text: resultsStrings.FqlOperatorBeginsWith },
                      ],
                    },
                    {
                      id: 'value',
                      title: resultsStrings.RefinementFilterValueColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'docx',
                    },
                  ],
                }),
                PropertyPaneSchemaHelper('collapseSpecification', {
                  label: resultsStrings.CollapseSpecificationLabel,
                  description: resultsStrings.CollapseSpecificationDescription,
                  value: this.properties.collapseSpecification || '',
                  filterHint: 'sortable',
                  validation: { requireSortable: true },
                }),
                PropertyPaneToggle('showDeleteConfirmation', {
                  label: resultsStrings.ShowDeleteConfirmationLabel,
                  onText: resultsStrings.ToggleOnText,
                  offText: resultsStrings.ToggleOffText,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
