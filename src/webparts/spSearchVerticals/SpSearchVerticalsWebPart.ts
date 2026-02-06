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

import * as strings from 'SpSearchVerticalsWebPartStrings';
import SpSearchVerticals from './components/SpSearchVerticals';
import { type ISpSearchVerticalsProps } from './components/ISpSearchVerticalsProps';
import { type ISearchStore, type IVerticalDefinition } from '@interfaces/index';
import { getStore } from '@store/store';

export interface ISpSearchVerticalsWebPartProps {
  searchContextId: string;
  verticals: string; // JSON-serialized IVerticalDefinition[]
  showCounts: boolean;
  hideEmptyVerticals: boolean;
  tabStyle: 'tabs' | 'pills' | 'underline';
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

    // Parse verticals JSON and set in store
    this._syncVerticalsToStore();

    return Promise.resolve();
  }

  private _syncVerticalsToStore(): void {
    if (!this._store) {
      return;
    }

    let parsed: IVerticalDefinition[] = [];
    try {
      const raw: string = this.properties.verticals;
      if (raw) {
        parsed = JSON.parse(raw) as IVerticalDefinition[];
      }
    } catch {
      // Invalid JSON; fall back to empty array
      parsed = [];
    }

    // Sort by sortOrder before storing
    parsed.sort(function (a: IVerticalDefinition, b: IVerticalDefinition): number {
      return a.sortOrder - b.sortOrder;
    });

    const state: ISearchStore = this._store.getState();
    // Only update if verticals actually changed
    if (JSON.stringify(state.verticals) !== JSON.stringify(parsed)) {
      this._store.setState({ verticals: parsed });
      // If no vertical is selected yet and we have verticals, select the first one
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
    // Re-sync verticals when the property pane changes
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
                PropertyPaneTextField('verticals', {
                  label: strings.VerticalsFieldLabel,
                  description: strings.VerticalsFieldDescription,
                  multiline: true,
                  rows: 10
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
