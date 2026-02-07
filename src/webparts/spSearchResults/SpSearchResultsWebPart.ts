import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
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
import { PropertyPaneSchemaHelper } from '../../propertyPaneControls/PropertyPaneSchemaHelper';

export interface ISpSearchResultsWebPartProps {
  searchContextId: string;
  queryTemplate: string;
  selectedProperties: string;
  pageSize: number;
  sortablePropertiesCollection: ISortCollectionItem[];
  defaultLayout: string;
  showResultCount: boolean;
  showSortDropdown: boolean;
  enableSelection: boolean;
}

interface ISortCollectionItem {
  uniqueId: string;
  property: string;
  label: string;
  direction: string;
}

export default class SpSearchResultsWebPart extends BaseClientSideWebPart<ISpSearchResultsWebPartProps> {

  private _theme: IReadonlyTheme | undefined;
  private _store: StoreApi<ISearchStore> | undefined;
  private _orchestrator: SearchOrchestrator | undefined;

  public render(): void {
    const contextId: string = this.properties.searchContextId || 'default';
    const element: React.ReactElement<ISpSearchResultsProps> = React.createElement(
      SpSearchResults,
      {
        store: this._store as StoreApi<ISearchStore>,
        orchestrator: this._orchestrator,
        searchContextId: contextId,
        theme: this._theme,
        showResultCount: this.properties.showResultCount,
        showSortDropdown: this.properties.showSortDropdown,
        enableSelection: this.properties.enableSelection,
        defaultLayout: this.properties.defaultLayout,
        pageSize: this.properties.pageSize
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
      registerBuiltInActions(this._store.getState().registries.actions);
    }

    // Initialize the search context (history service, orchestrator, etc.)
    await initializeSearchContext(contextId);

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

    // Sync query template to the store
    this._syncQueryTemplateToStore();

    // Sync sortable properties to the store
    this._syncSortablePropertiesToStore();
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

    const sortableProperties = (this.properties.sortablePropertiesCollection || []).map((item: ISortCollectionItem) => ({
      property: item.property,
      label: item.label,
      direction: item.direction || 'Ascending'
    }));

    const state: ISearchStore = this._store.getState();
    if (JSON.stringify(state.sortableProperties) !== JSON.stringify(sortableProperties)) {
      this._store.setState({ sortableProperties });
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

    if (propertyPath === 'pageSize' && this._store) {
      this._store.setState({ pageSize: this.properties.pageSize });
    }

    if (propertyPath === 'defaultLayout' && this._store) {
      const state: ISearchStore = this._store.getState();
      state.setLayout(this.properties.defaultLayout);
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
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneTextField('searchContextId', {
                  label: strings.SearchContextIdLabel,
                  description: strings.SearchContextIdDescription
                }),
                PropertyPaneSchemaHelper('queryTemplate', {
                  label: strings.QueryTemplateLabel,
                  description: strings.QueryTemplateDescription,
                  value: this.properties.queryTemplate || '',
                  filterHint: 'queryable',
                  applyOnEnter: true,
                }),
                PropertyPaneSchemaHelper('selectedProperties', {
                  label: strings.SelectedPropertiesLabel,
                  description: strings.SelectedPropertiesDescription,
                  value: this.properties.selectedProperties || '',
                  multiline: true,
                  rows: 4,
                  filterHint: 'retrievable',
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
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('defaultLayout', {
                  label: strings.DefaultLayoutLabel,
                  options: [
                    { key: 'list', text: strings.ListLayoutText, iconProps: { officeFabricIconFontName: 'List' } },
                    { key: 'compact', text: strings.CompactLayoutText, iconProps: { officeFabricIconFontName: 'GridViewSmall' } }
                  ]
                }),
                PropertyPaneToggle('showResultCount', {
                  label: strings.ShowResultCountLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
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
                }),
                PropertyPaneToggle('enableSelection', {
                  label: strings.EnableSelectionLabel,
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
