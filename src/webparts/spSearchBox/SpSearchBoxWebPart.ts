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
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { StoreApi } from 'zustand/vanilla';
import { spfxToolkitStylesLoaded } from '../../styles/loadSpfxToolkitStyles';

import * as strings from 'SpSearchBoxWebPartStrings';
import SpSearchBox from './components/SpSearchBox';
import { ISpSearchBoxProps } from './components/ISpSearchBoxProps';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { ISearchStore, ISearchScope } from '@interfaces/index';
import { getStore, initializeSearchContext, getManagerService, incrementContextRef, decrementContextRef } from '@store/store';
import { SharePointSearchProvider } from '@providers/index';
import { registerBuiltInSuggestions } from './registerBuiltInSuggestions';
import { DebugCollector } from '@store/debug';
import { ensurePnpPropertyControlStyles } from '../../styles/pnpPropertyControlsFix';
// T4.D8 — shared validator for the newPageQueryParameter URL-key field.
import { validateNewPageQueryParameter } from '../../propertyPaneControls/fieldValidation';
// T4.D11 — context-sensitive help link helper.
import { propertyPaneGroupHelp } from '../../propertyPaneControls/propertyPaneGroupHelp';
import { DisplayMode } from '@microsoft/sp-core-library';
import { AudienceGate, parseAudienceGroups } from '../../utilities/AudienceGate';
import { SearchContextIdBannerWrapper } from '../../utilities/SearchContextIdMismatchBanner';
import { SPDebugProvider } from 'spfx-toolkit/lib/components/debug';
import {
  propertyPaneSearchContextIdField,
  SEARCH_CONTEXT_ID_GROUP_NAME,
} from '../../propertyPaneControls/PropertyPaneSearchContextIdField';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
const _ensureStyles = spfxToolkitStylesLoaded;

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
  /** Stream D / #10 — comma/newline-separated Azure AD group IDs. Empty = visible to everyone. */
  audienceGroups: string;
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

    const audienceGroups = parseAudienceGroups(this.properties.audienceGroups);
    const innerElement: React.ReactElement<ISpSearchBoxProps> = React.createElement(
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
        managerService: getManagerService(this.properties.searchContextId || 'default'),
      }
    );

    // Stream D / #10 — wrap with AudienceGate so the web part hides itself
    // when the current user isn't in any of the configured groups.
    const gatedElement: React.ReactElement = React.createElement(
      AudienceGate,
      { audienceGroups, store: this._store },
      innerElement
    );

    // T3.D2 — edit-mode banner above the gated tree warning admins when
    // this web part's searchContextId doesn't match other search web parts
    // on the page. View-mode renders nothing.
    const bannerWrapped: React.ReactElement = React.createElement(
      SearchContextIdBannerWrapper,
      {
        webPartId: this.instanceId,
        contextId: this.properties.searchContextId || 'default',
        webPartLabel: 'SP Search Box',
        isEditMode: this.displayMode === DisplayMode.Edit,
      },
      gatedElement
    );

    // SPDebug — toolkit's debug runtime + lazy-loaded panel.
    // Per-web-part state isolation: each web part bundles its own copy of
    // the toolkit, so each SPDebugProvider has its own SPDebugStore. URL
    // activation (`?debug=1` / `?isDebug=1`) toggles every web part on the
    // page in lockstep; keyboard shortcut `Ctrl+Alt+D` toggles the panel of
    // whichever web part has focus. Coexists with the project's existing
    // DebugCollector + DebugFab (window-backed) on the Results web part.
    const element: React.ReactElement = React.createElement(
      SPDebugProvider,
      { logger: SPContext.logger, allowInProduction: false },
      bannerWrapped
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    ensurePnpPropertyControlStyles();

    // Initialize SPContext for PnPjs
    // Cast needed: spfx-toolkit uses SPFx 1.21.1 types; this project uses 1.22.2
    await SPContext.basic(this.context as unknown as Parameters<typeof SPContext.basic>[0], 'SPSearchBox');

    // Get or create the shared Zustand store
    const contextId = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);
    // T3.D1 — register this web part as a refcount holder. Drops in onDispose.
    incrementContextRef(contextId);

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
      // Safe to register after initializeSearchContext — suggestion providers are
      // UI-only (not involved in search execution) so the suggestion registry is
      // intentionally NOT frozen by SearchOrchestrator.freezeRegistries().
      registerBuiltInSuggestions(suggestions, managerService, dataProviders);
    }

    this.properties.searchScopes = normalizeSearchScopes(this.properties.searchScopes);
    DebugCollector.registerWebPart('SPSearchBoxWebPart', this.properties as unknown as Record<string, unknown>);
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
    // T3.D1 — decrement the per-context refcount BEFORE unmounting the
    // React tree. When the last web part on the context unmounts, the
    // deferred dispose tears down URL sync, orchestrator, and the
    // window-backed context entry. Cross-page SPA navigation order is
    // handled by the microtask deferral in the registry.
    const contextId = this.properties.searchContextId || 'default';
    decrementContextRef(contextId);
    // Unmount React component tree
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
            // T3.D4 — searchContextId is the first field every admin sees
            // on every search web part. Hoisted to page-1 / group-1 via the
            // shared helper so the label / description / required-error
            // string match across all six panes.
            {
              groupName: SEARCH_CONTEXT_ID_GROUP_NAME,
              groupFields: [
                propertyPaneSearchContextIdField()
              ]
            },
            {
              groupName: strings.SearchGroupName,
              groupFields: [
                propertyPaneGroupHelp('box-search', 'Help: Search input behaviour'),
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
                propertyPaneGroupHelp('box-navigation', 'Help: Same-page vs new-page navigation'),
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
                      const trimmed = value.trim();
                      if (!trimmed.startsWith('/') && !trimmed.startsWith('https://') && !trimmed.startsWith('http://')) {
                        return 'URL must start with /, https://, or http://';
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
                    description: strings.NewPageQueryParameterDescription,
                    // T4.D8 — alphanumeric + dash + underscore only.
                    // Special URL characters (?, &, =, #, space) corrupt
                    // the query string the search box constructs.
                    onGetErrorMessage: validateNewPageQueryParameter
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
                propertyPaneGroupHelp('box-suggestions', 'Help: Search suggestions and quick results'),
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
            // T3.D4 — searchContextId moved to page-1 group-1 via the shared
            // helper. The Connections page now hosts only audience targeting.
            // Stream D / #10 — per-web-part audience targeting.
            {
              groupName: strings.AudienceTargetingGroupName,
              groupFields: [
                PropertyPaneTextField('audienceGroups', {
                  label: strings.AudienceTargetingLabel,
                  description: strings.AudienceTargetingDescription,
                  multiline: true,
                  rows: 3,
                  resizable: true
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
                  PropertyFieldCollectionData('searchScopes', {
                    key: 'searchScopes',
                    label: 'Search scopes',
                    panelHeader: 'Configure search scopes',
                    panelDescription: 'Define the scopes available in the scope selector dropdown. Each scope needs a unique ID, a display label, and an optional KQL path filter.',
                    manageBtnLabel: 'Manage scopes',
                    value: this.properties.searchScopes as unknown as Array<Record<string, unknown>>,
                    fields: [
                      {
                        id: 'id',
                        title: 'ID',
                        type: CustomCollectionFieldType.string,
                        required: true,
                        placeholder: 'e.g. allsites'
                      },
                      {
                        id: 'label',
                        title: 'Label',
                        type: CustomCollectionFieldType.string,
                        required: true,
                        placeholder: 'e.g. All sites'
                      },
                      {
                        id: 'kqlPath',
                        title: 'KQL Path',
                        type: CustomCollectionFieldType.string,
                        required: false,
                        placeholder: 'e.g. path:"https://tenant.sharepoint.com/sites/hr"'
                      }
                    ]
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
