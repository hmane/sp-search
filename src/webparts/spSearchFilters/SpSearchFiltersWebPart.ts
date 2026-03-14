import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { StoreApi } from 'zustand/vanilla';
import { spfxToolkitStylesLoaded } from '../../styles/loadSpfxToolkitStyles';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'SpSearchFiltersWebPartStrings';
import SpSearchFilters from './components/SpSearchFilters';
import type { ISpSearchFiltersProps } from './components/ISpSearchFiltersProps';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { getStore, initializeSearchContext } from '@store/store';
import type { ISearchStore, IFilterConfig } from '@interfaces/index';
import { registerBuiltInFilterTypes } from './registerBuiltInFilterTypes';
import { SharePointSearchProvider } from '@providers/index';
import { sanitizeUrlAlias } from '@store/utils/filterUrlAliases';

// Bundle DevExtreme CSS — injected via style-loader at runtime.
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('devextreme/dist/css/dx.common.css');
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('devextreme/dist/css/dx.light.css');
void spfxToolkitStylesLoaded;

export interface ISpSearchFiltersWebPartProps {
  searchContextId: string;
  filtersCollection: IFilterCollectionItem[];
  applyMode: 'instant' | 'manual';
  operatorBetweenFilters: 'AND' | 'OR';
  showClearAll: boolean;
  enableVisualFilterBuilder: boolean;
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
}

function normalizeFiltersCollectionValue(
  raw: IFilterCollectionItem[] | string | undefined
): IFilterCollectionItem[] {
  if (Array.isArray(raw)) {
    return raw;
  }
  if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw) as IFilterCollectionItem[];
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }
  return [];
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

export default class SpSearchFiltersWebPart extends BaseClientSideWebPart<ISpSearchFiltersWebPartProps> {

  private _store: StoreApi<ISearchStore> | undefined;

  public render(): void {
    const element: React.ReactElement<ISpSearchFiltersProps> = React.createElement(
      SpSearchFilters,
      {
        store: this._store,
        applyMode: this.properties.applyMode || 'instant',
        showClearAll: this.properties.showClearAll !== false,
        enableVisualFilterBuilder: !!this.properties.enableVisualFilterBuilder
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    try {
      await SPContext.basic(this.context, 'SPSearchFilters');
      const contextId: string = this.properties.searchContextId || 'default';
      this._store = getStore(contextId);

      // Register the SharePoint Search data provider (idempotent — skips if already registered by another web part)
      const provider = new SharePointSearchProvider();
      const dataProviders = this._store.getState().registries.dataProviders;
      if (!dataProviders.get(provider.id)) {
        dataProviders.register(provider);
      }

      // Register all built-in filter types (checkbox, daterange, toggle, tagbox, slider, taxonomy, people)
      const filterTypes = this._store.getState().registries.filterTypes;
      registerBuiltInFilterTypes(filterTypes);

      // Sync filter configuration and operator to the store BEFORE initializing context,
      // so the first search (triggered by initializeSearchContext) includes refiner properties
      this._syncFilterConfigToStore();
      this._syncOperatorToStore();

      // Initialize the shared search context (ensures library bundle's SPContext is ready)
      // Idempotent — if already initialized by another web part, this is a no-op
      await initializeSearchContext(contextId, this.context);
    } catch (err) {
      console.error('[SPSearchFilters] onInit failed:', err);
      throw err;
    }
  }

  private _syncFilterConfigToStore(): void {
    if (!this._store) {
      return;
    }

    const rawFilters = normalizeFiltersCollectionValue(this.properties.filtersCollection);

    if (!rawFilters || rawFilters.length === 0) {
      const state: ISearchStore = this._store.getState();
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
        multiValues: item.multiValues !== false
      };
    });

    const state: ISearchStore = this._store.getState();
    if (JSON.stringify(state.filterConfig) !== JSON.stringify(filterConfig)) {
      this._store.setState({ filterConfig });
    }
  }

  private _syncOperatorToStore(): void {
    if (!this._store) {
      return;
    }
    const operator: 'AND' | 'OR' = (this.properties.operatorBetweenFilters === 'OR') ? 'OR' : 'AND';
    const state: ISearchStore = this._store.getState();
    if (state.operatorBetweenFilters !== operator) {
      state.setOperatorBetweenFilters(operator);
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
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
    }

    if (propertyPath === 'filtersCollection' || propertyPath === 'searchContextId') {
      this._syncFilterConfigToStore();
    }

    if (propertyPath === 'operatorBetweenFilters' || propertyPath === 'searchContextId') {
      this._syncOperatorToStore();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.FiltersPageHeader
          },
          groups: [
            {
              groupName: strings.FiltersGroupName,
              groupFields: [
                PropertyFieldCollectionData('filtersCollection', {
                  key: 'filtersCollection',
                  label: strings.FiltersFieldLabel,
                  panelHeader: strings.FiltersPanelHeader,
                  manageBtnLabel: strings.FiltersManageBtn,
                  value: normalizeFiltersCollectionValue(this.properties.filtersCollection),
                  enableSorting: true,
                  fields: [
                    {
                      id: 'managedProperty',
                      title: strings.FilterPropertyColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'RefinableString00'
                    },
                    {
                      id: 'displayName',
                      title: strings.FilterDisplayNameColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'File Type'
                    },
                    {
                      id: 'urlAlias',
                      title: strings.FilterUrlAliasColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'ft'
                    },
                    {
                      id: 'filterType',
                      title: strings.FilterTypeColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: [
                        { key: 'checkbox', text: strings.FilterTypeCheckbox },
                        { key: 'daterange', text: strings.FilterTypeDateRange },
                        { key: 'text', text: strings.FilterTypeText },
                        { key: 'people', text: strings.FilterTypePeople },
                        { key: 'taxonomy', text: strings.FilterTypeTaxonomy },
                        { key: 'slider', text: strings.FilterTypeSlider },
                        { key: 'tagbox', text: strings.FilterTypeTagBox },
                        { key: 'toggle', text: strings.FilterTypeToggle }
                      ]
                    },
                    {
                      id: 'operator',
                      title: strings.FilterOperatorColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: false,
                      options: [
                        { key: 'OR', text: 'OR' },
                        { key: 'AND', text: 'AND' }
                      ]
                    },
                    {
                      id: 'maxValues',
                      title: strings.FilterMaxValuesColumn,
                      type: CustomCollectionFieldType.number,
                      required: false,
                      defaultValue: 10
                    },
                    {
                      id: 'showCount',
                      title: strings.FilterShowCountColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: true
                    },
                    {
                      id: 'defaultExpanded',
                      title: strings.FilterExpandedColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: true
                    },
                    {
                      id: 'sortBy',
                      title: strings.FilterSortByColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: false,
                      options: [
                        { key: 'count', text: strings.FilterSortByCount },
                        { key: 'alphabetical', text: strings.FilterSortByName }
                      ]
                    },
                    {
                      id: 'sortDirection',
                      title: strings.FilterSortDirectionColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: false,
                      options: [
                        { key: 'desc', text: strings.FilterSortDescending },
                        { key: 'asc', text: strings.FilterSortAscending }
                      ]
                    },
                    {
                      id: 'multiValues',
                      title: strings.FilterMultiValuesColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: true
                    }
                  ]
                }),
              ]
            },
            {
              groupName: strings.BehaviorGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('applyMode', {
                  label: strings.ApplyModeLabel,
                  options: [
                    { key: 'instant', text: strings.ApplyModeInstant },
                    { key: 'manual', text: strings.ApplyModeManual }
                  ]
                }),
                PropertyPaneChoiceGroup('operatorBetweenFilters', {
                  label: strings.OperatorLabel,
                  options: [
                    { key: 'AND', text: 'AND' },
                    { key: 'OR', text: 'OR' }
                  ]
                }),
                PropertyPaneToggle('showClearAll', {
                  label: strings.ShowClearAllLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                })
              ]
            }
          ]
        },
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
                  value: this.properties.searchContextId || 'default',
                  onGetErrorMessage: (value: string): string => {
                    if (!value || value.trim() === '') {
                      return 'Required — must match the Search Context ID set on the Search Results web part.';
                    }
                    return '';
                  }
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.AdvancedPageHeader
          },
          groups: [
            {
              groupName: strings.AdvancedGroupName,
              groupFields: [
                PropertyPaneToggle('enableVisualFilterBuilder', {
                  label: strings.EnableVisualFilterBuilderLabel,
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
