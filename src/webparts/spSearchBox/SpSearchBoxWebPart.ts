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

import * as strings from 'SpSearchBoxWebPartStrings';
import SpSearchBox from './components/SpSearchBox';
import { ISpSearchBoxProps } from './components/ISpSearchBoxProps';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { ISearchStore, ISearchScope } from '@interfaces/index';
import { getStore, initializeSearchContext } from '@store/store';
import { SharePointSearchProvider } from '@providers/index';
import { spfi, SPFx } from '@pnp/sp';

export interface ISpSearchBoxWebPartProps {
  searchContextId: string;
  placeholder: string;
  debounceMs: number;
  searchBehavior: 'onEnter' | 'onButton' | 'both';
  enableScopeSelector: boolean;
  searchScopes: ISearchScope[];
  enableSuggestions: boolean;
  enableSearchManager: boolean;
}

export default class SpSearchBoxWebPart extends BaseClientSideWebPart<ISpSearchBoxWebPartProps> {

  private _store: StoreApi<ISearchStore> | undefined;
  private _theme: IReadonlyTheme | undefined;

  public render(): void {
    if (!this._store) {
      return;
    }

    const element: React.ReactElement<ISpSearchBoxProps> = React.createElement(
      SpSearchBox,
      {
        store: this._store,
        placeholder: this.properties.placeholder || 'Search...',
        debounceMs: this.properties.debounceMs || 300,
        searchBehavior: this.properties.searchBehavior || 'both',
        enableScopeSelector: !!this.properties.enableScopeSelector,
        searchScopes: this.properties.searchScopes || [],
        enableSuggestions: !!this.properties.enableSuggestions,
        enableSearchManager: !!this.properties.enableSearchManager,
        theme: this._theme,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // Initialize SPContext for PnPjs
    await SPContext.basic(this.context, 'SPSearchBox');

    // Get or create the shared Zustand store
    const contextId = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);

    // Register the SharePoint Search data provider
    const provider = new SharePointSearchProvider(SPContext.sp);
    const dataProviders = this._store.getState().registries.dataProviders;
    if (!dataProviders.get(provider.id)) {
      dataProviders.register(provider);
    }

    // Initialize the shared search context (orchestrator + manager service)
    // This is idempotent - if already initialized by another web part, it's a no-op
    const sp = spfi().using(SPFx(this.context));
    await initializeSearchContext(contextId, sp);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    this._theme = currentTheme;

    if (!currentTheme) {
      return;
    }

    const semanticColors = currentTheme.semanticColors;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
      this.domElement.style.setProperty('--link', semanticColors.link || '');
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
      this.domElement.style.setProperty('--inputBackground', semanticColors.inputBackground || '');
      this.domElement.style.setProperty('--inputBorder', semanticColors.inputBorder || '');
      this.domElement.style.setProperty('--inputBorderHovered', semanticColors.inputBorderHovered || '');
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
    // Unmount React component tree
    // Note: We don't stop the shared orchestrator here - it's managed by the registry
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('searchContextId', {
                  label: strings.SearchContextIdFieldLabel,
                  description: strings.SearchContextIdFieldDescription,
                }),
                PropertyPaneTextField('placeholder', {
                  label: strings.PlaceholderFieldLabel,
                }),
                PropertyPaneSlider('debounceMs', {
                  label: strings.DebounceMsFieldLabel,
                  min: 100,
                  max: 2000,
                  step: 50,
                  showValue: true,
                  value: this.properties.debounceMs || 300,
                }),
                PropertyPaneChoiceGroup('searchBehavior', {
                  label: strings.SearchBehaviorFieldLabel,
                  options: [
                    { key: 'onEnter', text: strings.SearchBehaviorOnEnter },
                    { key: 'onButton', text: strings.SearchBehaviorOnButton },
                    { key: 'both', text: strings.SearchBehaviorBoth },
                  ]
                }),
                PropertyPaneToggle('enableScopeSelector', {
                  label: strings.EnableScopeSelectorFieldLabel,
                }),
                PropertyPaneToggle('enableSuggestions', {
                  label: strings.EnableSuggestionsFieldLabel,
                }),
                PropertyPaneToggle('enableSearchManager', {
                  label: 'Enable saved searches button',
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
