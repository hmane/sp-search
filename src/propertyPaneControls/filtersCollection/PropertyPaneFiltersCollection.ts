import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  type IPropertyPaneField,
  type IPropertyPaneCustomFieldProps,
  PropertyPaneFieldType,
} from '@microsoft/sp-property-pane';

import FiltersCollectionControl from './FiltersCollectionControl';
import type { IFiltersCollectionItem } from './FiltersCollectionControl';

export type { IFiltersCollectionItem } from './FiltersCollectionControl';

export interface IPropertyPaneFiltersCollectionProps {
  label: string;
  panelHeader: string;
  manageButtonLabel: string;
  value: IFiltersCollectionItem[];
}

/**
 * Master-detail replacement for `PropertyFieldCollectionData` on the
 * Configure refiners surface. Renders a "Manage refiners" button that
 * opens a wide Panel split into a refiner list on the left and a
 * sectioned form on the right. The form only surfaces fields that are
 * relevant to the selected refiner's filterType (see `fieldRelevance.ts`).
 */
export function PropertyPaneFiltersCollection(
  targetProperty: string,
  properties: IPropertyPaneFiltersCollectionProps
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty: targetProperty,
    properties: {
      key: 'filtersCollection_' + targetProperty,
      context: properties,
      onRender: function (
        domElement: HTMLElement,
        context: IPropertyPaneFiltersCollectionProps,
        changeCallback?: (targetProperty?: string, newValue?: unknown) => void
      ): void {
        function handleChange(next: IFiltersCollectionItem[]): void {
          if (changeCallback) {
            changeCallback(targetProperty, next);
          }
        }

        ReactDom.render(
          React.createElement(FiltersCollectionControl, {
            label: context.label,
            panelHeader: context.panelHeader,
            manageButtonLabel: context.manageButtonLabel,
            value: context.value,
            onChange: handleChange,
          }),
          domElement
        );
      },
      onDispose: function (domElement: HTMLElement): void {
        ReactDom.unmountComponentAtNode(domElement);
      },
    },
  };
}
