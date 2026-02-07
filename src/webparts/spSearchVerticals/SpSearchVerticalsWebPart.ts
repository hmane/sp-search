import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { type StoreApi } from 'zustand/vanilla';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'SpSearchVerticalsWebPartStrings';
import SpSearchVerticals from './components/SpSearchVerticals';
import { type ISpSearchVerticalsProps } from './components/ISpSearchVerticalsProps';
import { type ISearchStore, type IVerticalDefinition } from '@interfaces/index';
import { getStore } from '@store/store';

export interface ISpSearchVerticalsWebPartProps {
  searchContextId: string;
  verticals: string; // Legacy: JSON-serialized IVerticalDefinition[]
  verticalsCollection: IVerticalCollectionItem[]; // Collection data from property pane
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
  sortOrder: number;
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

  protected onInit(): Promise<void> {
    const contextId: string = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);

    // Migrate legacy JSON to collection data if needed
    if (this.properties.verticals && (!this.properties.verticalsCollection || this.properties.verticalsCollection.length === 0)) {
      this._migrateJsonToCollection();
    }

    this._syncVerticalsToStore();

    return Promise.resolve();
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
          sortOrder: v.sortOrder ?? idx + 1
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
        state.setVertical(parsed[0].key);
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
                  description: strings.SearchContextIdFieldDescription
                }),
                PropertyFieldCollectionData('verticalsCollection', {
                  key: 'verticalsCollection',
                  label: strings.VerticalsFieldLabel,
                  panelHeader: strings.VerticalsPanelHeader,
                  manageBtnLabel: strings.VerticalsManageBtn,
                  value: this.properties.verticalsCollection,
                  enableSorting: true,
                  fields: [
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
                      id: 'resultSourceId',
                      title: strings.VerticalSourceColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'GUID'
                    },
                    {
                      id: 'sortOrder',
                      title: strings.VerticalSortColumn,
                      type: CustomCollectionFieldType.number,
                      required: true,
                      defaultValue: 1
                    }
                  ]
                }),
                PropertyPaneToggle('showCounts', {
                  label: strings.ShowCountsFieldLabel
                }),
                PropertyPaneToggle('hideEmptyVerticals', {
                  label: strings.HideEmptyVerticalsFieldLabel
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
