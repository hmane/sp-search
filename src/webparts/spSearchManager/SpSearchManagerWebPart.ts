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
import type { ISpSearchManagerProps } from './components/ISpSearchManagerProps';
import styles from './components/SpSearchManager.module.scss';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { ISearchStore } from '@interfaces/index';
import { getStore, initializeSearchContext, incrementContextRef, decrementContextRef } from '@store/store';
import { SharePointSearchProvider } from '@providers/index';
import { SearchManagerService } from '@services/index';
import { ensurePnpPropertyControlStyles } from '../../styles/pnpPropertyControlsFix';
// T4.D8 — shared property-pane field validator (expectedSiteUrls).
// Centralised so AdminManager (which inherits this property-pane builder)
// gets identical validation copy.
import {
  validateExpectedSiteUrlsField,
} from '../../propertyPaneControls/fieldValidation';
// T4.D11 — context-sensitive help link helper.
import { propertyPaneGroupHelp } from '../../propertyPaneControls/propertyPaneGroupHelp';
import { ICoverageProfile, normalizeCoverageProfile } from '@services/SearchCoverageService';
import { DebugCollector } from '@store/debug';
import { DisplayMode } from '@microsoft/sp-core-library';
import { SearchContextIdBannerWrapper } from '../../utilities/SearchContextIdMismatchBanner';
import { SPDebugProvider } from 'spfx-toolkit/lib/components/debug';
import {
  propertyPaneSearchContextIdField,
  SEARCH_CONTEXT_ID_GROUP_NAME,
} from '../../propertyPaneControls/PropertyPaneSearchContextIdField';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
const _ensureStyles = spfxToolkitStylesLoaded;

export interface ISpSearchManagerWebPartProps {
  searchContextId: string;
  mode: 'standalone' | 'panel';
  defaultTab: 'saved' | 'history' | 'collections' | 'health' | 'insights' | 'dashboard';
  enableSavedSearches: boolean;
  enableSharedSearches: boolean;
  enableCollections: boolean;
  enableHistory: boolean;
  coverageProfilesCollection: ICoverageProfileCollectionItem[];
  enableHealth: boolean;
  enableInsights: boolean;
  enableDashboard: boolean;
  enableAnnotations: boolean;
  expectedSiteUrls: string;
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

  /** T3.D2 — surface name shown in the mismatch banner. Overridden by AdminManager. */
  protected _getWebPartLabel(): string {
    return 'SP Search Manager';
  }

  /**
   * T4.D6 — variant flag for the rendered surface.
   *
   * Base class returns 'user' (end-user saved/shared/collections/history
   * surface). AdminManager overrides to return 'admin' (coverage/health/
   * insights). Driving the variant via a virtual lets subclasses fork the
   * `<SpSearchManager>` props without copying the whole render block.
   */
  protected _getVariant(): 'user' | 'admin' {
    return 'user';
  }

  /**
   * T4.D6 — props the variant projects into the `<SpSearchManager>` element.
   * Base class produces user-tab props (saved/shared/collections/history
   * enabled per manifest defaults; admin fields force-false). AdminManager
   * subclass overrides this hook to return the admin-tab projection.
   */
  protected _buildManagerProps(): ISpSearchManagerProps {
    return {
      store: this._store as StoreApi<ISearchStore>,
      service: this._service as SearchManagerService,
      theme: this._theme,
      variant: 'user',
      searchContextId: this.properties.searchContextId || 'default',
      mode: this.properties.mode || 'panel',
      defaultTab: this.properties.defaultTab || 'saved',
      headerTitle: 'Search Manager',
      context: this.context,
      // User-tab toggles — projected from properties (manifest defaults
      // enable saved/shared/collections/history). Admins flip them via
      // the property pane.
      enableSavedSearches: this.properties.enableSavedSearches !== false,
      enableSharedSearches: this.properties.enableSharedSearches !== false,
      enableCollections: this.properties.enableCollections !== false,
      enableHistory: this.properties.enableHistory !== false,
      enableAnnotations: this.properties.enableAnnotations === true,
      maxHistoryItems: this.properties.maxHistoryItems || 50,
      showResetAction: this.properties.showResetAction === true,
      showSaveAction: this.properties.showSaveAction === true,
      // Admin-only surface — force-disabled in the user variant. Even if
      // the property pane somehow has these flags set true, the user
      // variant renders without admin tabs.
      coverageProfiles: [],
      enableHealth: false,
      enableInsights: false,
      enableDashboard: false,
      expectedSiteUrls: [],
      // T4.D5 — edit-mode validators don't fire in user variant since
      // none of the admin fields are exposed; still pass isEditMode so
      // the mismatch banner can fire.
      isEditMode: this.displayMode === DisplayMode.Edit,
      tenantRoot: this.context.pageContext.web.absoluteUrl,
    };
  }

  public render(): void {
    let element: React.ReactElement;

    if (!this._hasAdminAccess && this._getVariant() === 'admin') {
      element = React.createElement(
        'div',
        { className: styles.accessDenied },
        React.createElement(Icon, { iconName: 'Lock', className: styles.accessDeniedIcon }),
        React.createElement('h2', { className: styles.accessDeniedTitle }, strings.AccessDeniedTitle),
        React.createElement('p', { className: styles.accessDeniedDescription }, strings.AccessDeniedDescription)
      );
    } else if (!this._store || !this._service) {
      return;
    } else {
      element = React.createElement(SpSearchManager, this._buildManagerProps());
    }

    // T3.D2 — edit-mode mismatch banner. AdminManager extends this class
    // and overrides `_getWebPartLabel()` so the banner copy names the
    // correct surface in each subclass.
    const bannerWrapped: React.ReactElement = React.createElement(
      SearchContextIdBannerWrapper,
      {
        webPartId: this.instanceId,
        contextId: this.properties.searchContextId || 'default',
        webPartLabel: this._getWebPartLabel(),
        isEditMode: this.displayMode === DisplayMode.Edit,
      },
      element
    );

    // SPDebug — toolkit's debug runtime + lazy-loaded panel. See SpSearchBox
    // for the per-web-part-state-isolation note. AdminManager extends this
    // class so it inherits the wrapping automatically.
    const wrappedElement: React.ReactElement = React.createElement(
      SPDebugProvider,
      { logger: SPContext.logger, allowInProduction: false },
      bannerWrapped
    );

    ReactDom.render(wrappedElement, this.domElement);
  }

  protected async onInit(): Promise<void> {
    ensurePnpPropertyControlStyles();

    const webPermissions = this.context.pageContext.web.permissions;
    // Cast needed: spfx-toolkit uses SPFx 1.21.1 types; this project uses 1.22.2
    this._hasAdminAccess = !!(webPermissions && (webPermissions as unknown as { hasPermission(perm: unknown): boolean }).hasPermission(SPPermission.manageWeb));

    if (!this._hasAdminAccess) {
      return;
    }

    // Initialize SPContext for PnPjs
    // Cast needed: spfx-toolkit uses SPFx 1.21.1 types; this project uses 1.22.2
    await SPContext.basic(this.context as unknown as Parameters<typeof SPContext.basic>[0], 'SPSearchManager');

    // Get or create the shared Zustand store
    const contextId: string = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);
    // T3.D1 — refcount holder. AdminManager extends this class so the
    // increment fires once per AdminManager instance too via super.onInit().
    incrementContextRef(contextId);

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
    DebugCollector.registerWebPart('SPSearchManagerWebPart', this.properties as unknown as Record<string, unknown>);
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
    // T3.D1 — drop refcount before unmounting. AdminManager subclass
    // inherits this method; the increment happens once per instance in
    // the inherited onInit(), so the single decrement here is correct.
    const contextId: string = this.properties.searchContextId || 'default';
    decrementContextRef(contextId);
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
            description: this._getPropertyPaneHeaderDescription()
          },
          groups: this._buildPropertyPaneGroups()
        }
      ]
    };
  }

  /**
   * T4.D6 — header description shown at the top of the property pane.
   * Base class returns the user-facing copy; AdminManager overrides.
   */
  protected _getPropertyPaneHeaderDescription(): string {
    return strings.PropertyPaneDescription;
  }

  /**
   * T4.D6 — property pane groups for the user-facing Manager.
   *
   * The base class (user-facing Manager) exposes only end-user toggles:
   * saved/shared/collections/history + maxHistoryItems. Admin fields
   * (`health` / `insights` / `coverageProfilesCollection` /
   * `expectedSiteUrls` / `enableDashboard`) live exclusively on the
   * AdminManager subclass's property pane — the audit's "suppress all
   * user-facing toggles from Admin Manager's pane at all" applies
   * symmetrically: admin fields don't appear in user-Manager's pane either.
   */
  protected _buildPropertyPaneGroups(): IPropertyPaneConfiguration['pages'][0]['groups'] {
    return [
      // T3.D4 — searchContextId is the first field every admin sees on
      // every search web part. Shared helper.
      {
        groupName: SEARCH_CONTEXT_ID_GROUP_NAME,
        groupFields: [
          propertyPaneSearchContextIdField()
        ]
      },
      {
        groupName: strings.DisplayGroupName,
        groupFields: [
          PropertyPaneChoiceGroup('defaultTab', {
            label: strings.DefaultTabLabel,
            options: [
              { key: 'saved', text: 'Saved Searches' },
              { key: 'history', text: 'History' },
              { key: 'collections', text: 'Collections' }
            ]
          })
        ]
      },
      {
        groupName: 'User tabs',
        groupFields: [
          propertyPaneGroupHelp('manager-user-tabs', 'Help: User-facing Manager tabs'),
          PropertyPaneToggle('enableSavedSearches', {
            label: 'Show Saved Searches tab',
            onText: strings.ToggleOnText,
            offText: strings.ToggleOffText
          }),
          PropertyPaneToggle('enableSharedSearches', {
            label: 'Show Shared Searches tab',
            onText: strings.ToggleOnText,
            offText: strings.ToggleOffText
          }),
          PropertyPaneToggle('enableCollections', {
            label: 'Show Collections tab',
            onText: strings.ToggleOnText,
            offText: strings.ToggleOffText
          }),
          PropertyPaneToggle('enableHistory', {
            label: 'Show History tab',
            onText: strings.ToggleOnText,
            offText: strings.ToggleOffText
          })
        ]
      }
    ];
  }

  /**
   * T4.D6 — helper that returns the admin-only property pane groups. The
   * base-class user-facing property pane does NOT include these groups;
   * AdminManager subclass calls this helper from its overridden
   * `_buildPropertyPaneGroups()`.
   */
  protected _buildAdminPropertyPaneGroups(): IPropertyPaneConfiguration['pages'][0]['groups'] {
    return [
      // T3.D4 — same searchContextId field; appears first in the admin pane.
      {
        groupName: SEARCH_CONTEXT_ID_GROUP_NAME,
        groupFields: [
          propertyPaneSearchContextIdField()
        ]
      },
      {
        groupName: strings.DisplayGroupName,
        groupFields: [
          PropertyPaneChoiceGroup('defaultTab', {
            label: strings.DefaultTabLabel,
            options: [
              { key: 'dashboard', text: strings.DefaultTabDashboard },
              { key: 'health', text: strings.DefaultTabHealth },
              { key: 'insights', text: strings.DefaultTabInsights }
            ]
          })
        ]
      },
      {
        groupName: strings.SectionsGroupName,
        groupFields: [
          PropertyPaneToggle('enableDashboard', {
            label: strings.EnableDashboardLabel,
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
        groupName: 'Gap Analysis',
        groupFields: [
          PropertyPaneTextField('expectedSiteUrls', {
            label: 'Expected Site URLs (one per line)',
            multiline: true,
            rows: 5,
            description: 'Enter site URLs to monitor for content coverage. One URL per line.',
            // T4.D8 — per-line URL validator. Reports the first offending
            // line so admins can locate the bad row in the textarea.
            onGetErrorMessage: validateExpectedSiteUrlsField
          })
        ]
      },
      {
        groupName: strings.MonitoringGroupName,
        groupFields: [
          propertyPaneGroupHelp('adminmgr-coverage', 'Help: Coverage profiles and monitoring'),
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
    ];
  }

  /** T4.D6 — protected so AdminManager subclass can reuse the normaliser. */
  protected _normalizeCoverageProfiles(): ICoverageProfile[] {
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
