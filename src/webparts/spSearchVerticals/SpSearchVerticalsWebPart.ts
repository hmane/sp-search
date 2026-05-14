import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { type StoreApi } from 'zustand/vanilla';
import { spfxToolkitStylesLoaded } from '../../styles/loadSpfxToolkitStyles';

import { PropertyFieldCollectionData, CustomCollectionFieldType, type ICustomCollectionField } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'SpSearchVerticalsWebPartStrings';
import SpSearchVerticals from './components/SpSearchVerticals';
import { type ISpSearchVerticalsProps } from './components/ISpSearchVerticalsProps';
import { type ISearchStore, type IVerticalDefinition } from '@interfaces/index';
import { getStore, initializeSearchContext, incrementContextRef, decrementContextRef } from '@store/store';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { SharePointSearchProvider } from '@providers/index';
import { ensurePnpPropertyControlStyles } from '../../styles/pnpPropertyControlsFix';
import { DebugCollector } from '@store/debug';
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

export interface ISpSearchVerticalsWebPartProps {
  searchContextId: string;
  verticals: string; // Legacy: JSON-serialized IVerticalDefinition[]
  verticalsCollection: IVerticalCollectionItem[]; // Collection data from property pane
  defaultVertical: string;
  showCounts: boolean;
  hideEmptyVerticals: boolean;
  tabStyle: 'tabs' | 'pills' | 'underline';
  /** Stream D / #10 — comma/newline-separated Azure AD group IDs. Empty = visible to everyone. */
  audienceGroups: string;
}

interface IVerticalCollectionItem {
  uniqueId: string;
  key: string;
  label: string;
  iconName: string;
  queryTemplate: string;
  resultSourceId: string;
  /** ID of the data provider to use for this vertical, e.g. 'graph-people'. */
  dataProviderId: string;
  /** Layout key to activate automatically when this vertical is selected. */
  defaultLayout: string;
  sortOrder: number;
  isLink: boolean;
  linkUrl: string;
  openBehavior: string;
  audience: string;
}

export default class SpSearchVerticalsWebPart extends BaseClientSideWebPart<ISpSearchVerticalsWebPartProps> {

  private _store: StoreApi<ISearchStore> | undefined;
  private _theme: IReadonlyTheme | undefined;

  public render(): void {
    if (!this._store) {
      return;
    }

    const innerElement: React.ReactElement<ISpSearchVerticalsProps> = React.createElement(
      SpSearchVerticals,
      {
        store: this._store,
        showCounts: this.properties.showCounts,
        hideEmptyVerticals: this.properties.hideEmptyVerticals,
        tabStyle: this.properties.tabStyle || 'tabs',
        theme: this._theme,
        isEditMode: this.displayMode === DisplayMode.Edit
      }
    );

    // Stream D / #10 — wrap with AudienceGate so the web part hides itself
    // when the current user isn't in any of the configured groups.
    const audienceGroups = parseAudienceGroups(this.properties.audienceGroups);
    const gatedElement: React.ReactElement = React.createElement(
      AudienceGate,
      { audienceGroups, store: this._store },
      innerElement
    );

    // T3.D2 — edit-mode mismatch banner above the gated tree.
    const bannerWrapped: React.ReactElement = React.createElement(
      SearchContextIdBannerWrapper,
      {
        webPartId: this.instanceId,
        contextId: this.properties.searchContextId || 'default',
        webPartLabel: 'SP Search Verticals',
        isEditMode: this.displayMode === DisplayMode.Edit,
      },
      gatedElement
    );

    // SPDebug — toolkit's debug runtime + lazy-loaded panel. See SpSearchBox
    // for the per-web-part-state-isolation note.
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
    await SPContext.basic(this.context as unknown as Parameters<typeof SPContext.basic>[0], 'SPSearchVerticals');

    const contextId: string = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);
    // T3.D1 — refcount holder.
    incrementContextRef(contextId);

    // Register the SharePoint Search data provider (idempotent — skips if already registered by another web part)
    const provider = new SharePointSearchProvider();
    const dataProviders = this._store.getState().registries.dataProviders;
    if (!dataProviders.get(provider.id)) {
      dataProviders.register(provider);
    }

    // Migrate legacy JSON to collection data if needed
    if (this.properties.verticals && (!this.properties.verticalsCollection || this.properties.verticalsCollection.length === 0)) {
      this._migrateJsonToCollection();
    }

    this._syncVerticalsToStore();

    // Initialize the shared search context (ensures library bundle's SPContext is ready)
    // Idempotent — if already initialized by another web part, this is a no-op
    await initializeSearchContext(contextId, this.context);
    DebugCollector.registerWebPart('SPSearchVerticalsWebPart', this.properties as unknown as Record<string, unknown>);
  }

  private _migrateJsonToCollection(): void {
    try {
      const parsed: IVerticalDefinition[] = JSON.parse(this.properties.verticals);
      if (Array.isArray(parsed) && parsed.length > 0) {
        this.properties.verticalsCollection = parsed.map((v: IVerticalDefinition, idx: number) => ({
          uniqueId: `v-${idx}`,
          key: v.key,
          label: v.label,
          iconName: v.iconName || '',
          queryTemplate: v.queryTemplate || '',
          resultSourceId: v.resultSourceId || '',
          dataProviderId: v.dataProviderId || '',
          defaultLayout: v.defaultLayout || '',
          sortOrder: v.sortOrder ?? idx + 1,
          isLink: !!v.isLink,
          linkUrl: v.linkUrl || '',
          openBehavior: v.openBehavior || 'currentTab',
          audience: v.audienceGroups ? v.audienceGroups.join(',') : ''
        }));
      }
    } catch {
      // Invalid JSON, ignore
    }
  }

  private _syncVerticalsToStore(): void {
    if (!this._store) {
      return;
    }

    let parsed: IVerticalDefinition[] = [];

    // Prefer collection data; fall back to legacy JSON
    if (this.properties.verticalsCollection && this.properties.verticalsCollection.length > 0) {
      // Use array index as sortOrder — PropertyFieldCollectionData drag-and-drop
      // reorders the array, so array position is the canonical order
      parsed = this.properties.verticalsCollection.map((item: IVerticalCollectionItem, idx: number) => ({
        key: item.key,
        label: item.label,
        iconName: item.iconName || undefined,
        queryTemplate: item.queryTemplate || undefined,
        resultSourceId: item.resultSourceId || undefined,
        dataProviderId: item.dataProviderId || undefined,
        defaultLayout: item.defaultLayout || undefined,
        isLink: !!item.isLink,
        linkUrl: item.linkUrl || undefined,
        openBehavior: (item.openBehavior as 'currentTab' | 'newTab') || undefined,
        audienceGroups: item.audience ? item.audience.split(',').map((s: string) => s.trim()).filter(Boolean) : undefined,
        sortOrder: idx + 1
      }));
    } else if (this.properties.verticals) {
      try {
        parsed = JSON.parse(this.properties.verticals) as IVerticalDefinition[];
        // Legacy JSON: sort by explicit sortOrder
        parsed.sort(function (a: IVerticalDefinition, b: IVerticalDefinition): number {
          return a.sortOrder - b.sortOrder;
        });
      } catch {
        parsed = [];
      }
    }

    const state: ISearchStore = this._store.getState();
    if (JSON.stringify(state.verticals) !== JSON.stringify(parsed)) {
      this._store.setState({ verticals: parsed });
      if (parsed.length > 0 && !state.currentVerticalKey) {
        const defaultKey = this.properties.defaultVertical || parsed[0].key;
        state.setVertical(defaultKey);
      }
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    this._theme = currentTheme;
  }

  protected onDispose(): void {
    // T3.D1 — drop refcount before unmounting React.
    const contextId: string = this.properties.searchContextId || 'default';
    decrementContextRef(contextId);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (this._store) {
      const contextId: string = this.properties.searchContextId || 'default';
      this._store = getStore(contextId);
      this._syncVerticalsToStore();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Build dropdown options from configured verticals
    const verticalOptions = (this.properties.verticalsCollection || [])
      .filter((v: IVerticalCollectionItem) => !v.isLink)
      .map((v: IVerticalCollectionItem) => ({ key: v.key, text: v.label || v.key }));
    // PropertyFieldCollectionData field configs have a broad union type; build them in steps
    // so TypeScript doesn't over-narrow the base array and reject advanced dropdown fields.
    const verticalFields: ICustomCollectionField[] = [
      {
        id: 'key',
        title: strings.VerticalKeyColumn,
        type: CustomCollectionFieldType.string,
        required: true,
        placeholder: 'documents'
      },
      {
        id: 'label',
        title: strings.VerticalLabelColumn,
        type: CustomCollectionFieldType.string,
        required: true,
        placeholder: 'Documents'
      },
      {
        id: 'iconName',
        title: strings.VerticalIconColumn,
        type: CustomCollectionFieldType.string,
        required: false,
        placeholder: 'Document'
      },
      {
        id: 'queryTemplate',
        title: strings.VerticalQueryColumn,
        type: CustomCollectionFieldType.string,
        required: false,
        placeholder: '{searchTerms} IsDocument:1'
      },
      {
        id: 'isLink',
        title: strings.VerticalIsLinkColumn,
        type: CustomCollectionFieldType.boolean,
        required: false,
        defaultValue: false
      },
      {
        id: 'linkUrl',
        title: strings.VerticalLinkUrlColumn,
        type: CustomCollectionFieldType.string,
        required: false,
        placeholder: 'https://...'
      }
    ];

    verticalFields.push(
      {
        id: 'resultSourceId',
        title: strings.VerticalSourceColumn,
        type: CustomCollectionFieldType.string,
        required: false,
        placeholder: 'GUID'
      },
      {
        id: 'dataProviderId',
        title: strings.VerticalProviderColumn,
        type: CustomCollectionFieldType.dropdown,
        required: false,
        options: [
          { key: '', text: 'SharePoint Search (default)' },
          { key: 'sharepoint-search', text: 'SharePoint Search' },
          { key: 'graph-search', text: 'Graph Search (files)' },
          { key: 'graph-people', text: 'Graph Search (people)' }
        ]
      },
      {
        id: 'defaultLayout',
        title: strings.VerticalDefaultLayoutColumn,
        type: CustomCollectionFieldType.dropdown,
        required: false,
        options: [
          { key: '', text: 'Inherit active layout' },
          { key: 'list', text: 'List' },
          { key: 'compact', text: 'Compact' },
          { key: 'grid', text: 'Grid' },
          { key: 'card', text: 'Card' },
          { key: 'people', text: 'People' },
          { key: 'gallery', text: 'Gallery' }
        ]
      },
      {
        id: 'openBehavior',
        title: strings.VerticalOpenBehaviorColumn,
        type: CustomCollectionFieldType.dropdown,
        required: false,
        options: [
          { key: 'currentTab', text: strings.OpenBehaviorCurrentTab },
          { key: 'newTab', text: strings.OpenBehaviorNewTab }
        ]
      },
      {
        id: 'audience',
        title: strings.VerticalAudienceColumn,
        type: CustomCollectionFieldType.string,
        required: false,
        placeholder: strings.VerticalAudiencePlaceholder
      }
    );

    return {
      pages: [
        // ─── Page 1: Verticals Configuration ────────────
        {
          header: {
            description: strings.VerticalsPageHeader
          },
          groups: [
            // T3.D4 — searchContextId is the first field every admin sees
            // on every search web part. Shared helper.
            {
              groupName: SEARCH_CONTEXT_ID_GROUP_NAME,
              groupFields: [
                propertyPaneSearchContextIdField()
              ]
            },
            {
              groupName: strings.VerticalsGroupName,
              groupFields: [
                PropertyPaneLabel('verticalsIntro', {
                  text: strings.VerticalsIntroLabel
                }),
                PropertyFieldCollectionData('verticalsCollection', {
                  key: 'verticalsCollection',
                  label: strings.VerticalsFieldLabel,
                  panelHeader: strings.VerticalsPanelHeader,
                  manageBtnLabel: strings.VerticalsManageBtn,
                  value: this.properties.verticalsCollection,
                  enableSorting: true,
                  fields: verticalFields,
                  disabled: false
                }),
                ...(verticalOptions.length > 0 ? [
                  PropertyPaneDropdown('defaultVertical', {
                    label: strings.DefaultVerticalLabel,
                    options: verticalOptions
                  })
                ] : [])
              ]
            }
          ]
        },
        // ─── Page 2: Display ────────────────────────────
        {
          header: {
            description: strings.DisplayPageHeader
          },
          groups: [
            {
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneToggle('showCounts', {
                  label: strings.ShowCountsFieldLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneToggle('hideEmptyVerticals', {
                  label: strings.HideEmptyVerticalsFieldLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                }),
                PropertyPaneChoiceGroup('tabStyle', {
                  label: strings.TabStyleFieldLabel,
                  options: [
                    { key: 'tabs', text: strings.TabStyleTabs },
                    { key: 'pills', text: strings.TabStylePills },
                    { key: 'underline', text: strings.TabStyleUnderline }
                  ]
                })
              ]
            },
            // Stream D / #10 — per-web-part audience targeting. Independent of
            // the per-vertical `audience` column already on the verticals
            // collection (that controls individual tab visibility; this one
            // controls whether the entire tab strip renders).
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
            }
          ]
        }
      ]
    };
  }
}
