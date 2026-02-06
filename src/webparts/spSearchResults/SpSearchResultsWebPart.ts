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
import { spfi, SPFx } from '@pnp/sp';

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

export interface ISpSearchResultsWebPartProps {
  searchContextId: string;
  queryTemplate: string;
  selectedProperties: string;
  pageSize: number;
  defaultLayout: string;
  showResultCount: boolean;
  showSortDropdown: boolean;
  enableSelection: boolean;
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
    const contextId: string = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);
    this._orchestrator = getOrchestrator(contextId);

    // Initialize PnPjs with SPFx context
    const sp = spfi().using(SPFx(this.context));

    // Initialize the search context (history service, orchestrator, etc.)
    await initializeSearchContext(contextId, sp);

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
                PropertyPaneTextField('queryTemplate', {
                  label: strings.QueryTemplateLabel,
                  description: strings.QueryTemplateDescription
                }),
                PropertyPaneTextField('selectedProperties', {
                  label: strings.SelectedPropertiesLabel,
                  description: strings.SelectedPropertiesDescription,
                  multiline: true,
                  rows: 4
                }),
                PropertyPaneSlider('pageSize', {
                  label: strings.PageSizeLabel,
                  min: 5,
                  max: 100,
                  value: 25,
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
