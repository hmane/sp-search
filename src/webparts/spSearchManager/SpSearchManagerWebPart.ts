import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { StoreApi } from 'zustand/vanilla';

import * as strings from 'SpSearchManagerWebPartStrings';
import SpSearchManager from './components/SpSearchManager';
import { ISpSearchManagerProps } from './components/ISpSearchManagerProps';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { ISearchStore } from '@interfaces/index';
import { getStore } from '@store/store';
import { SearchManagerService } from '@services/index';

export interface ISpSearchManagerWebPartProps {
  searchContextId: string;
  mode: 'standalone' | 'panel';
}

export default class SpSearchManagerWebPart extends BaseClientSideWebPart<ISpSearchManagerWebPartProps> {

  private _theme: IReadonlyTheme | undefined;
  private _store: StoreApi<ISearchStore> | undefined;
  private _service: SearchManagerService | undefined;

  public render(): void {
    if (!this._store || !this._service) {
      return;
    }

    const element: React.ReactElement<ISpSearchManagerProps> = React.createElement(
      SpSearchManager,
      {
        store: this._store,
        service: this._service,
        theme: this._theme,
        mode: this.properties.mode || 'standalone'
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // Initialize SPContext for PnPjs
    await SPContext.basic(this.context, 'SPSearchManager');

    // Get or create the shared Zustand store
    const contextId: string = this.properties.searchContextId || 'default';
    this._store = getStore(contextId);

    // Create and initialize the SearchManagerService
    this._service = new SearchManagerService(SPContext.sp);
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('searchContextId', {
                  label: 'Search Context ID',
                  description: 'Enter the same ID as the other search web parts on this page to share state. Leave blank for the default context.'
                }),
                PropertyPaneChoiceGroup('mode', {
                  label: 'Display Mode',
                  options: [
                    { key: 'standalone', text: 'Standalone (full web part)' },
                    { key: 'panel', text: 'Panel (side panel overlay)' }
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
