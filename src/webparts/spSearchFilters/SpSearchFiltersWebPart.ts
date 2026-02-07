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

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'SpSearchFiltersWebPartStrings';
import SpSearchFilters from './components/SpSearchFilters';
import type { ISpSearchFiltersProps } from './components/ISpSearchFiltersProps';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { getStore } from '@store/store';
import type { ISearchStore, IFilterConfig } from '@interfaces/index';
import { registerBuiltInFilterTypes } from './registerBuiltInFilterTypes';

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
  filterType: string;
  operator: string;
  maxValues: number;
  defaultExpanded: boolean;
  showCount: boolean;
  sortBy: string;
}

export default class SpSearchFiltersWebPart extends BaseClientSideWebPart<ISpSearchFiltersWebPartProps> {

  private _store: StoreApi<ISearchStore> | undefined;

  public render(): void {
    const element: React.ReactElement<ISpSearchFiltersProps> = React.createElement(
      SpSearchFilters,
      {
        store: this._store,
        applyMode: this.properties.applyMode || 'instant',
        operatorBetweenFilters: this.properties.operatorBetweenFilters || 'AND',
        showClearAll: this.properties.showClearAll !== false,
        enableVisualFilterBuilder: !!this.properties.enableVisualFilterBuilder
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await SPContext.basic(this.context, 'SPSearchFilters');
    const contextId: string = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);

    // Register all built-in filter types (checkbox, daterange, toggle, tagbox, slider, taxonomy, people)
    const filterTypes = this._store.getState().registries.filterTypes;
    registerBuiltInFilterTypes(filterTypes);

    // Sync filter configuration to the store
    this._syncFilterConfigToStore();
  }

  private _syncFilterConfigToStore(): void {
    if (!this._store) {
      return;
    }

    if (!this.properties.filtersCollection || this.properties.filtersCollection.length === 0) {
      const state: ISearchStore = this._store.getState();
      if (state.filterConfig.length > 0) {
        this._store.setState({ filterConfig: [], activeFilters: [] });
      }
      return;
    }

    const filterConfig: IFilterConfig[] = this.properties.filtersCollection.map((item: IFilterCollectionItem) => ({
      id: item.uniqueId,
      managedProperty: item.managedProperty,
      displayName: item.displayName,
      filterType: (item.filterType || 'checkbox') as IFilterConfig['filterType'],
      operator: (item.operator || 'OR') as IFilterConfig['operator'],
      maxValues: item.maxValues || 10,
      defaultExpanded: item.defaultExpanded !== false,
      showCount: item.showCount !== false,
      sortBy: (item.sortBy || 'count') as IFilterConfig['sortBy'],
      sortDirection: 'desc' as const
    }));

    const state: ISearchStore = this._store.getState();
    if (JSON.stringify(state.filterConfig) !== JSON.stringify(filterConfig)) {
      this._store.setState({ filterConfig });
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
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('searchContextId', {
                  label: strings.SearchContextIdLabel,
                  description: strings.SearchContextIdDescription,
                  value: this.properties.searchContextId || 'default'
                }),
                PropertyFieldCollectionData('filtersCollection', {
                  key: 'filtersCollection',
                  label: strings.FiltersFieldLabel,
                  panelHeader: strings.FiltersPanelHeader,
                  manageBtnLabel: strings.FiltersManageBtn,
                  value: this.properties.filtersCollection,
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
                      id: 'filterType',
                      title: strings.FilterTypeColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: [
                        { key: 'checkbox', text: strings.FilterTypeCheckbox },
                        { key: 'daterange', text: strings.FilterTypeDateRange },
                        { key: 'slider', text: strings.FilterTypeSlider },
                        { key: 'people', text: strings.FilterTypePeople },
                        { key: 'taxonomy', text: strings.FilterTypeTaxonomy },
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
                    }
                  ]
                }),
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
                }),
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
