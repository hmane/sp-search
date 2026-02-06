import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  type IPropertyPaneField,
  type IPropertyPaneCustomFieldProps,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';

import SchemaHelperControl from './SchemaHelperControl';
import type { SchemaFilterHint } from './SchemaHelperControl';

// ─── Public API ─────────────────────────────────────────────

export interface IPropertyPaneSchemaHelperProps {
  /** Label displayed above the text field */
  label: string;
  /** Description text below the text field */
  description?: string;
  /** Current value of the property (pass this.properties[targetProperty]) */
  value: string;
  /** Whether the text field is multiline (for comma-separated property lists) */
  multiline?: boolean;
  /** Number of rows for multiline text field */
  rows?: number;
  /** Pre-selects a Pivot tab matching this flag in the schema browser */
  filterHint?: SchemaFilterHint;
}

/**
 * PropertyPaneSchemaHelper — creates a custom property pane field that combines
 * a text field with a "Browse Schema" button. The button opens a Panel with a
 * searchable, filterable list of SharePoint managed properties.
 *
 * Usage in getPropertyPaneConfiguration():
 * ```
 * PropertyPaneSchemaHelper('selectedProperties', {
 *   label: 'Selected Properties',
 *   value: this.properties.selectedProperties || '',
 *   multiline: true,
 *   rows: 4,
 *   filterHint: 'retrievable',
 * })
 * ```
 *
 * @param targetProperty - The web part property to bind to (e.g. 'selectedProperties')
 * @param properties - Configuration for the schema helper control
 * @returns SPFx property pane field definition
 */
export function PropertyPaneSchemaHelper(
  targetProperty: string,
  properties: IPropertyPaneSchemaHelperProps
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty: targetProperty,
    properties: {
      key: 'schemaHelper_' + targetProperty,
      context: properties,
      onRender: function (
        domElement: HTMLElement,
        context: IPropertyPaneSchemaHelperProps,
        changeCallback?: (targetProperty?: string, newValue?: unknown) => void
      ): void {
        function handleChange(newValue: string): void {
          if (changeCallback) {
            changeCallback(targetProperty, newValue);
          }
        }

        // ReactDom.render() on an already-mounted component updates props (acts as re-render).
        // SPFx calls onRender() each time getPropertyPaneConfiguration() returns,
        // so context.value always reflects the latest web part property value.
        ReactDom.render(
          React.createElement(SchemaHelperControl, {
            label: context.label,
            description: context.description,
            value: context.value,
            multiline: context.multiline,
            rows: context.rows,
            filterHint: context.filterHint,
            onChange: handleChange,
          }),
          domElement
        );
      },
      onDispose: function (domElement: HTMLElement): void {
        ReactDom.unmountComponentAtNode(domElement);
      }
    }
  };
}
