import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { StoreApi } from 'zustand/vanilla';

import * as strings from 'SpSearchFiltersWebPartStrings';
import SpSearchFilters from './components/SpSearchFilters';
import type { ISpSearchFiltersProps } from './components/ISpSearchFiltersProps';
import { getStore } from '@store/store';
import type { ISearchStore } from '@interfaces/index';

export interface ISpSearchFiltersWebPartProps {
  searchContextId: string;
  applyMode: 'instant' | 'manual';
  operatorBetweenFilters: 'AND' | 'OR';
  showClearAll: boolean;
}

export default class SpSearchFiltersWebPart extends BaseClientSideWebPart<ISpSearchFiltersWebPartProps> {

  private _store: StoreApi<ISearchStore> | undefined;

  public render(): void {
    const element: React.ReactElement<ISpSearchFiltersProps> = React.createElement(
      SpSearchFilters,
      {
        store: this._store,
        applyMode: this.properties.applyMode || 'instant',
        operatorBetweenFilters: this.properties.operatorBetweenFilters || 'AND',
        showClearAll: this.properties.showClearAll !== false
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    const contextId: string = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
      this.domElement.style.setProperty('--link', semanticColors.link || '');
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('searchContextId', {
                  label: strings.SearchContextIdLabel,
                  description: strings.SearchContextIdDescription,
                  value: this.properties.searchContextId || 'default'
                }),
                PropertyPaneChoiceGroup('applyMode', {
                  label: strings.ApplyModeLabel,
                  options: [
                    { key: 'instant', text: strings.ApplyModeInstant },
                    { key: 'manual', text: strings.ApplyModeManual }
                  ]
                }),
                PropertyPaneChoiceGroup('operatorBetweenFilters', {
                  label: strings.OperatorLabel,
                  options: [
                    { key: 'AND', text: 'AND' },
                    { key: 'OR', text: 'OR' }
                  ]
                }),
                PropertyPaneToggle('showClearAll', {
                  label: strings.ShowClearAllLabel,
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
