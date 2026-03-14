import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { Icon } from '@fluentui/react/lib/Icon';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPPermission } from '@microsoft/sp-page-context';
import { StoreApi } from 'zustand/vanilla';
import { spfxToolkitStylesLoaded } from '../../styles/loadSpfxToolkitStyles';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'SpSearchManagerWebPartStrings';
import SpSearchManager from './components/SpSearchManager';
import { ISpSearchManagerProps } from './components/ISpSearchManagerProps';
import styles from './components/SpSearchManager.module.scss';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { ISearchStore } from '@interfaces/index';
import { getStore, initializeSearchContext } from '@store/store';
import { SharePointSearchProvider } from '@providers/index';
import { SearchManagerService } from '@services/index';
import { ICoverageProfile, normalizeCoverageProfile } from '@services/SearchCoverageService';

void spfxToolkitStylesLoaded;

export interface ISpSearchManagerWebPartProps {
  searchContextId: string;
  coverageSourcePageUrl: string;
  mode: 'standalone' | 'panel';
  defaultTab: 'saved' | 'history' | 'collections' | 'coverage' | 'health' | 'insights';
  enableSavedSearches: boolean;
  enableSharedSearches: boolean;
  enableCollections: boolean;
  enableHistory: boolean;
  enableCoverage: boolean;
  coverageProfilesCollection: ICoverageProfileCollectionItem[];
  enableHealth: boolean;
  enableInsights: boolean;
  enableAnnotations: boolean;
  maxHistoryItems: number;
  showResetAction: boolean;
  showSaveAction: boolean;
}

interface ICoverageProfileCollectionItem {
  uniqueId?: string;
  title: string;
  description?: string;
  queryTemplate?: string;
  resultSourceId?: string;
  sourceUrls?: string;
  contentTypeIds?: string;
  excludePaths?: string;
  includeFolders?: boolean;
  trimDuplicates?: boolean;
  refinementFilters?: string;
}

export default class SpSearchManagerWebPart extends BaseClientSideWebPart<ISpSearchManagerWebPartProps> {

  private _theme: IReadonlyTheme | undefined;
  private _store: StoreApi<ISearchStore> | undefined;
  private _service: SearchManagerService | undefined;
  private _hasAdminAccess: boolean = true;

  public render(): void {
    if (!this._hasAdminAccess) {
      ReactDom.render(
        React.createElement(
          'div',
          { className: styles.accessDenied },
          React.createElement(Icon, { iconName: 'Lock', className: styles.accessDeniedIcon }),
          React.createElement('h2', { className: styles.accessDeniedTitle }, strings.AccessDeniedTitle),
          React.createElement('p', { className: styles.accessDeniedDescription }, strings.AccessDeniedDescription)
        ),
        this.domElement
      );
      return;
    }

    if (!this._store || !this._service) {
      return;
    }

    const element: React.ReactElement<ISpSearchManagerProps> = React.createElement(
      SpSearchManager,
      {
        store: this._store,
        service: this._service,
        theme: this._theme,
        variant: 'admin',
        searchContextId: this.properties.searchContextId || 'default',
        mode: 'standalone',
        defaultTab: this.properties.defaultTab || 'coverage',
        headerTitle: 'Admin Search Manager',
        context: this.context,
        enableSavedSearches: false,
        enableSharedSearches: false,
        enableCollections: false,
        enableHistory: false,
        enableCoverage: this.properties.enableCoverage !== false,
        coverageSourcePageUrl: this.properties.coverageSourcePageUrl || '',
        coverageProfiles: this._normalizeCoverageProfiles(),
        enableHealth: this.properties.enableHealth !== false,
        enableInsights: this.properties.enableInsights !== false,
        enableAnnotations: false,
        maxHistoryItems: 0,
        showResetAction: false,
        showSaveAction: false
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    const webPermissions = this.context.pageContext.web.permissions;
    this._hasAdminAccess = !!(webPermissions && webPermissions.hasPermission(SPPermission.manageWeb));

    if (!this._hasAdminAccess) {
      return;
    }

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
                PropertyPaneTextField('coverageSourcePageUrl', {
                  label: strings.CoverageSourcePageUrlLabel,
                  description: strings.CoverageSourcePageUrlDescription,
                  placeholder: '/sites/search/SitePages/Search.aspx'
                })
              ]
            },
            {
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('defaultTab', {
                  label: strings.DefaultTabLabel,
                  options: [
                    { key: 'coverage', text: strings.DefaultTabCoverage },
                    { key: 'health', text: strings.DefaultTabHealth },
                    { key: 'insights', text: strings.DefaultTabInsights }
                  ]
                })
              ]
            },
            {
              groupName: strings.SectionsGroupName,
              groupFields: [
                PropertyPaneToggle('enableCoverage', {
                  label: strings.EnableCoverageLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('enableHealth', {
                  label: strings.EnableHealthLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('enableInsights', {
                  label: strings.EnableInsightsLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                })
              ]
            },
            {
              groupName: strings.MonitoringGroupName,
              groupFields: [
                PropertyFieldCollectionData('coverageProfilesCollection', {
                  key: 'coverageProfilesCollection',
                  label: strings.CoverageProfilesLabel,
                  panelHeader: strings.CoverageProfilesPanelHeader,
                  manageBtnLabel: strings.CoverageProfilesManageButton,
                  value: this.properties.coverageProfilesCollection || [],
                  enableSorting: true,
                  fields: [
                    {
                      id: 'title',
                      title: strings.CoverageProfileTitleColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'Default site content'
                    },
                    {
                      id: 'description',
                      title: strings.CoverageProfileDescriptionColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'Primary document libraries included in the search experience'
                    },
                    {
                      id: 'sourceUrls',
                      title: strings.CoverageProfileSourceUrlsColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: '/sites/demo/Shared Documents, /sites/demo/Policies'
                    },
                    {
                      id: 'contentTypeIds',
                      title: strings.CoverageProfileContentTypesColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: '0x0101, 0x01010007FF3E058DEF4E058DEF4E'
                    },
                    {
                      id: 'excludePaths',
                      title: strings.CoverageProfileExcludePathsColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: '/sites/demo/Shared Documents/Archive'
                    },
                    {
                      id: 'queryTemplate',
                      title: strings.CoverageProfileQueryTemplateColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: '{searchTerms} IsDocument:1'
                    },
                    {
                      id: 'resultSourceId',
                      title: strings.CoverageProfileResultSourceColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'GUID'
                    },
                    {
                      id: 'refinementFilters',
                      title: strings.CoverageProfileRefinementFiltersColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'FileType:or("docx","pdf")'
                    },
                    {
                      id: 'includeFolders',
                      title: strings.CoverageProfileIncludeFoldersColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: false
                    },
                    {
                      id: 'trimDuplicates',
                      title: strings.CoverageProfileTrimDuplicatesColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: true
                    }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _normalizeCoverageProfiles(): ICoverageProfile[] {
    return (this.properties.coverageProfilesCollection || [])
      .map(function (profile, index): ICoverageProfile {
        return normalizeCoverageProfile({
          id: (profile.uniqueId || 'coverage-profile-' + String(index + 1)).trim(),
          title: profile.title,
          description: profile.description,
          queryTemplate: profile.queryTemplate,
          resultSourceId: profile.resultSourceId,
          sourceUrls: profile.sourceUrls,
          contentTypeIds: profile.contentTypeIds,
          excludePaths: profile.excludePaths,
          includeFolders: !!profile.includeFolders,
          trimDuplicates: profile.trimDuplicates !== false,
          refinementFilters: profile.refinementFilters
        });
      })
      .filter(function (profile): boolean {
        return profile.sourceUrls.length > 0;
      });
  }
}
