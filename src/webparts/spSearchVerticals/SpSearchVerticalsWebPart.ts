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

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'SpSearchVerticalsWebPartStrings';
import SpSearchVerticals from './components/SpSearchVerticals';
import { type ISpSearchVerticalsProps } from './components/ISpSearchVerticalsProps';
import { type ISearchStore, type IVerticalDefinition } from '@interfaces/index';
import { getStore, initializeSearchContext } from '@store/store';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { SharePointSearchProvider } from '@providers/index';

export interface ISpSearchVerticalsWebPartProps {
  searchContextId: string;
  verticals: string; // Legacy: JSON-serialized IVerticalDefinition[]
  verticalsCollection: IVerticalCollectionItem[]; // Collection data from property pane
  defaultVertical: string;
  showCounts: boolean;
  hideEmptyVerticals: boolean;
  tabStyle: 'tabs' | 'pills' | 'underline';
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

    const element: React.ReactElement<ISpSearchVerticalsProps> = React.createElement(
      SpSearchVerticals,
      {
        store: this._store,
        showCounts: this.properties.showCounts,
        hideEmptyVerticals: this.properties.hideEmptyVerticals,
        tabStyle: this.properties.tabStyle || 'tabs',
        theme: this._theme
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // Initialize SPContext for PnPjs
    await SPContext.basic(this.context, 'SPSearchVerticals');

    const contextId: string = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);

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
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
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
    const verticalFields: any[] = [
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
            {
              groupName: strings.ConnectionGroupName,
              groupFields: [
                PropertyPaneTextField('searchContextId', {
                  label: strings.SearchContextIdFieldLabel,
                  description: strings.SearchContextIdFieldDescription
                })
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
            }
          ]
        }
      ]
    };
  }
}
