import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { StoreApi } from 'zustand/vanilla';

import * as strings from 'SpSearchManagerWebPartStrings';
import SpSearchManager from './components/SpSearchManager';
import { ISpSearchManagerProps } from './components/ISpSearchManagerProps';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { ISearchStore } from '@interfaces/index';
import { getStore, initializeSearchContext } from '@store/store';
import { SharePointSearchProvider } from '@providers/index';
import { SearchManagerService } from '@services/index';

export interface ISpSearchManagerWebPartProps {
  searchContextId: string;
  mode: 'standalone' | 'panel';
  enableSavedSearches: boolean;
  enableSharedSearches: boolean;
  enableCollections: boolean;
  enableHistory: boolean;
  enableAnnotations: boolean;
  maxHistoryItems: number;
}

export default class SpSearchManagerWebPart extends BaseClientSideWebPart<ISpSearchManagerWebPartProps> {

  private _theme: IReadonlyTheme | undefined;
  private _store: StoreApi<ISearchStore> | undefined;
  private _service: SearchManagerService | undefined;

  public render(): void {
    if (!this._store || !this._service) {
      return;
    }

    const element: React.ReactElement<ISpSearchManagerProps> = React.createElement(
      SpSearchManager,
      {
        store: this._store,
        service: this._service,
        theme: this._theme,
        mode: this.properties.mode || 'standalone',
        context: this.context,
        enableSavedSearches: this.properties.enableSavedSearches !== false,
        enableSharedSearches: this.properties.enableSharedSearches !== false,
        enableCollections: this.properties.enableCollections !== false,
        enableHistory: this.properties.enableHistory !== false,
        enableAnnotations: !!this.properties.enableAnnotations,
        maxHistoryItems: this.properties.maxHistoryItems || 50
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // Initialize SPContext for PnPjs
    await SPContext.basic(this.context, 'SPSearchManager');

    // Get or create the shared Zustand store
    const contextId: string = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);

    // Register the SharePoint Search data provider (idempotent — skips if already registered by another web part)
    const provider = new SharePointSearchProvider();
    const dataProviders = this._store.getState().registries.dataProviders;
    if (!dataProviders.get(provider.id)) {
      dataProviders.register(provider);
    }

    // Initialize the shared search context (ensures library bundle's SPContext is ready)
    // Idempotent — if already initialized by another web part, this is a no-op
    await initializeSearchContext(contextId, this.context);

    // Create and initialize the SearchManagerService (uses SPContext.sp internally)
    this._service = new SearchManagerService();
    await this._service.initialize();
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

    const palette = currentTheme.palette;
    if (palette) {
      this.domElement.style.setProperty('--themePrimary', palette.themePrimary || '');
      this.domElement.style.setProperty('--themeDarkAlt', palette.themeDarkAlt || '');
      this.domElement.style.setProperty('--themeDark', palette.themeDark || '');
      this.domElement.style.setProperty('--neutralLight', palette.neutralLight || '');
      this.domElement.style.setProperty('--neutralLighter', palette.neutralLighter || '');
      this.domElement.style.setProperty('--neutralSecondary', palette.neutralSecondary || '');
      this.domElement.style.setProperty('--white', palette.white || '');
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
              groupName: strings.ConnectionGroupName,
              groupFields: [
                PropertyPaneTextField('searchContextId', {
                  label: strings.SearchContextIdLabel,
                  description: strings.SearchContextIdDescription
                }),
                PropertyPaneChoiceGroup('mode', {
                  label: strings.DisplayModeLabel,
                  options: [
                    { key: 'standalone', text: strings.ModeStandalone },
                    { key: 'panel', text: strings.ModePanel }
                  ]
                })
              ]
            },
            {
              groupName: strings.FeaturesGroupName,
              groupFields: [
                PropertyPaneToggle('enableSavedSearches', {
                  label: strings.EnableSavedSearchesLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('enableSharedSearches', {
                  label: strings.EnableSharedSearchesLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('enableCollections', {
                  label: strings.EnableCollectionsLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('enableHistory', {
                  label: strings.EnableHistoryLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('enableAnnotations', {
                  label: strings.EnableAnnotationsLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneSlider('maxHistoryItems', {
                  label: strings.MaxHistoryItemsLabel,
                  min: 10,
                  max: 200,
                  step: 10,
                  showValue: true,
                  value: this.properties.maxHistoryItems || 50
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
