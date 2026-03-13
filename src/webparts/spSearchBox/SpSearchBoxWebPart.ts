import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneLabel,
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
import { getStore, initializeSearchContext, getManagerService } from '@store/store';
import { SharePointSearchProvider } from '@providers/index';
import { registerBuiltInSuggestions } from './registerBuiltInSuggestions';

export interface ISpSearchBoxWebPartProps {
  searchContextId: string;
  placeholder: string;
  debounceMs: number;
  searchBehavior: 'onEnter' | 'onButton' | 'both';
  resetSearchOnClear: boolean;
  enableScopeSelector: boolean;
  searchScopes: ISearchScope[];
  enableSuggestions: boolean;
  enableSharePointSuggestions: boolean;
  enableRecentSuggestions: boolean;
  enablePopularSuggestions: boolean;
  enableQuickResults: boolean;
  enablePropertySuggestions: boolean;
  suggestionsPerGroup: number;
  enableQueryBuilder: boolean;
  enableKqlMode: boolean;
  enableSearchManager: boolean;
  searchInNewPage: boolean;
  newPageUrl: string;
  newPageOpenBehavior: 'sameTab' | 'newTab';
  newPageParameterLocation: 'queryString' | 'hash';
  newPageQueryParameter: string;
  queryInputTransformation: string;
}

function normalizeSearchScopes(raw: unknown): ISearchScope[] {
  if (Array.isArray(raw)) {
    return raw as ISearchScope[];
  }
  if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw) as ISearchScope[];
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }
  return [];
}

export default class SpSearchBoxWebPart extends BaseClientSideWebPart<ISpSearchBoxWebPartProps> {

  private _store: StoreApi<ISearchStore> | undefined;
  private _theme: IReadonlyTheme | undefined;

  private _getWebAbsoluteUrl(): string {
    return this.context?.pageContext?.web?.absoluteUrl || '';
  }

  public render(): void {
    if (!this._store) {
      return;
    }

    const element: React.ReactElement<ISpSearchBoxProps> = React.createElement(
      SpSearchBox,
      {
        store: this._store,
        searchContextId: this.properties.searchContextId || 'default',
        siteUrl: this._getWebAbsoluteUrl(),
        placeholder: this.properties.placeholder || 'Search...',
        debounceMs: this.properties.debounceMs || 300,
        searchBehavior: this.properties.searchBehavior || 'both',
        resetSearchOnClear: this.properties.resetSearchOnClear !== false,
        enableScopeSelector: !!this.properties.enableScopeSelector,
        searchScopes: normalizeSearchScopes(this.properties.searchScopes),
        enableSuggestions: !!this.properties.enableSuggestions,
        enableSharePointSuggestions: this.properties.enableSharePointSuggestions !== false,
        enableRecentSuggestions: this.properties.enableRecentSuggestions !== false,
        enablePopularSuggestions: this.properties.enablePopularSuggestions !== false,
        enableQuickResults: this.properties.enableQuickResults !== false,
        enablePropertySuggestions: this.properties.enablePropertySuggestions !== false,
        suggestionsPerGroup: this.properties.suggestionsPerGroup || 5,
        enableQueryBuilder: !!this.properties.enableQueryBuilder,
        enableKqlMode: !!this.properties.enableKqlMode,
        enableSearchManager: !!this.properties.enableSearchManager,
        searchInNewPage: !!this.properties.searchInNewPage,
        newPageUrl: this.properties.newPageUrl || '',
        newPageOpenBehavior: this.properties.newPageOpenBehavior || 'sameTab',
        newPageParameterLocation: this.properties.newPageParameterLocation || 'queryString',
        newPageQueryParameter: this.properties.newPageQueryParameter || 'q',
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

    // Register the SharePoint Search data provider (uses SPContext.sp internally)
    const provider = new SharePointSearchProvider();
    const dataProviders = this._store.getState().registries.dataProviders;
    if (!dataProviders.get(provider.id)) {
      dataProviders.register(provider);
    }

    // Sync queryInputTransformation to store — applies before each search execution
    const transformation = this.properties.queryInputTransformation || '{searchTerms}';
    if (this._store.getState().queryInputTransformation !== transformation) {
      this._store.getState().setQueryInputTransformation(transformation);
    }

    // Initialize the shared search context (orchestrator + manager service)
    // This is idempotent - if already initialized by another web part, it's a no-op
    await initializeSearchContext(contextId, this.context);

    // Register built-in suggestion providers (Recent, Trending, ManagedProperty)
    const managerService = getManagerService(contextId);
    if (managerService) {
      const suggestions = this._store.getState().registries.suggestions;
      const dataProviders = this._store.getState().registries.dataProviders;
      registerBuiltInSuggestions(suggestions, managerService, dataProviders);
    }

    this.properties.searchScopes = normalizeSearchScopes(this.properties.searchScopes);
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

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected onPropertyPaneFieldChanged(propertyPath: string, _oldValue: any, _newValue: any): void {
    if (propertyPath === 'queryInputTransformation' && this._store) {
      const transformation = this.properties.queryInputTransformation || '{searchTerms}';
      this._store.getState().setQueryInputTransformation(transformation);
    }

    if (
      propertyPath === 'enableScopeSelector' ||
      propertyPath === 'searchInNewPage' ||
      propertyPath === 'enableSuggestions'
    ) {
      this.context.propertyPane.refresh();
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
        // ─── Page 1: Search Box Settings ──────────────────
        {
          header: {
            description: strings.SearchPageHeader
          },
          groups: [
            {
              groupName: strings.SearchGroupName,
              groupFields: [
                PropertyPaneTextField('placeholder', {
                  label: strings.PlaceholderFieldLabel,
                }),
                PropertyPaneChoiceGroup('searchBehavior', {
                  label: strings.SearchBehaviorFieldLabel,
                  options: [
                    { key: 'onEnter', text: strings.SearchBehaviorOnEnter },
                    { key: 'onButton', text: strings.SearchBehaviorOnButton },
                    { key: 'both', text: strings.SearchBehaviorBoth },
                  ]
                }),
                PropertyPaneToggle('resetSearchOnClear', {
                  label: strings.ResetSearchOnClearLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                })
              ]
            },
            {
              groupName: strings.QueryGroupName,
              groupFields: [
                PropertyPaneTextField('queryInputTransformation', {
                  label: strings.QueryInputTransformationLabel,
                  description: strings.QueryInputTransformationDescription
                })
              ]
            },
            {
              groupName: strings.NavigationGroupName,
              groupFields: [
                PropertyPaneToggle('searchInNewPage', {
                  label: strings.SearchInNewPageLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                ...(this.properties.searchInNewPage ? [
                  PropertyPaneTextField('newPageUrl', {
                    label: strings.NewPageUrlLabel,
                    description: strings.NewPageUrlDescription,
                    onGetErrorMessage: (value: string): string => {
                      if (!value || value.trim() === '') {
                        return strings.NewPageUrlRequiredMessage;
                      }
                      return '';
                    }
                  }),
                  PropertyPaneChoiceGroup('newPageOpenBehavior', {
                    label: strings.NewPageOpenBehaviorLabel,
                    options: [
                      { key: 'sameTab', text: strings.NewPageOpenBehaviorSameTab },
                      { key: 'newTab', text: strings.NewPageOpenBehaviorNewTab }
                    ]
                  }),
                  PropertyPaneChoiceGroup('newPageParameterLocation', {
                    label: strings.NewPageParameterLocationLabel,
                    options: [
                      { key: 'queryString', text: strings.NewPageParameterLocationQueryString },
                      { key: 'hash', text: strings.NewPageParameterLocationHash }
                    ]
                  }),
                  PropertyPaneTextField('newPageQueryParameter', {
                    label: strings.NewPageQueryParameterLabel,
                    description: strings.NewPageQueryParameterDescription
                  })
                ] : [])
              ]
            }
          ]
        },
        {
          header: {
            description: strings.SuggestionsPageHeader
          },
          groups: [
            {
              groupName: strings.SuggestionsGroupName,
              groupFields: [
                PropertyPaneToggle('enableSuggestions', {
                  label: strings.EnableSuggestionsFieldLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                ...(this.properties.enableSuggestions ? [
                  PropertyPaneSlider('suggestionsPerGroup', {
                    label: strings.SuggestionsPerGroupLabel,
                    min: 1,
                    max: 10,
                    step: 1,
                    showValue: true,
                    value: this.properties.suggestionsPerGroup || 5,
                  }),
                  PropertyPaneToggle('enableSharePointSuggestions', {
                    label: strings.EnableSharePointSuggestionsLabel,
                    onText: strings.ToggleOnText,
                    offText: strings.ToggleOffText
                  }),
                  PropertyPaneToggle('enableRecentSuggestions', {
                    label: strings.EnableRecentSuggestionsLabel,
                    onText: strings.ToggleOnText,
                    offText: strings.ToggleOffText
                  }),
                  PropertyPaneToggle('enablePopularSuggestions', {
                    label: strings.EnablePopularSuggestionsLabel,
                    onText: strings.ToggleOnText,
                    offText: strings.ToggleOffText
                  }),
                  PropertyPaneToggle('enableQuickResults', {
                    label: strings.EnableQuickResultsLabel,
                    onText: strings.ToggleOnText,
                    offText: strings.ToggleOffText
                  }),
                  PropertyPaneToggle('enablePropertySuggestions', {
                    label: strings.EnablePropertySuggestionsLabel,
                    onText: strings.ToggleOnText,
                    offText: strings.ToggleOffText
                  })
                ] : [])
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
              groupName: strings.ConnectionGroupName,
              groupFields: [
                PropertyPaneTextField('searchContextId', {
                  label: strings.SearchContextIdFieldLabel,
                  description: strings.SearchContextIdFieldDescription,
                  onGetErrorMessage: (value: string): string => {
                    if (!value || value.trim() === '') {
                      return 'Required — must match the Search Context ID set on the Search Results web part.';
                    }
                    return '';
                  }
                })
              ]
            },
            {
              groupName: strings.ScopeGroupName,
              groupFields: [
                PropertyPaneToggle('enableScopeSelector', {
                  label: strings.EnableScopeSelectorFieldLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                ...(this.properties.enableScopeSelector ? [
                  PropertyPaneLabel('scopeInfo', {
                    text: strings.ScopeInfoLabel
                  })
                ] : [])
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
                PropertyPaneSlider('debounceMs', {
                  label: strings.DebounceMsFieldLabel,
                  min: 100,
                  max: 2000,
                  step: 50,
                  showValue: true,
                  value: this.properties.debounceMs || 300,
                }),
                PropertyPaneToggle('enableQueryBuilder', {
                  label: strings.EnableQueryBuilderFieldLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('enableKqlMode', {
                  label: strings.EnableKqlModeFieldLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('enableSearchManager', {
                  label: strings.EnableSearchManagerFieldLabel,
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
