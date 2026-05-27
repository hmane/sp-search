import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  type IPropertyPaneField,
  type IPropertyPaneCustomFieldProps,
  PropertyPaneFieldType,
} from '@microsoft/sp-property-pane';

import { IColumnConfigItem } from './columnConfig';
import ColumnConfigControl from './ColumnConfigControl';

export interface IColumnConfigFieldProps {
  label: string;
  description?: string;
  value: IColumnConfigItem[];
  availableProperties: Array<{ key: string; text: string }>;
}

/**
 * Custom SPFx property-pane field for the Stream B / Phase 1 column-config
 * editor. Replaces `PropertyFieldCollectionData('gridPropertiesCollection',
 * …)` with a compact in-pane list + a Fluent Panel side editor.
 *
 * Spec: docs/superpowers/specs/2026-05-13-stream-b-column-config-design.md
 */
export function PropertyPaneColumnConfigField(
  targetProperty: string,
  properties: IColumnConfigFieldProps
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty,
    properties: {
      key: 'columnConfig_' + targetProperty,
      context: properties,
      onRender: function (
        domElement: HTMLElement,
        context: IColumnConfigFieldProps,
        changeCallback?: (targetProperty?: string, newValue?: unknown) => void
      ): void {
        function handleChange(newValue: IColumnConfigItem[]): void {
          if (changeCallback) {
            changeCallback(targetProperty, newValue);
          }
        }

        ReactDom.render(
          React.createElement(ColumnConfigControl, {
            label: context.label,
            description: context.description,
            value: context.value,
            availableProperties: context.availableProperties,
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
