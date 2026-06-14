import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  type IPropertyPaneField,
  type IPropertyPaneCustomFieldProps,
  PropertyPaneFieldType,
} from '@microsoft/sp-property-pane';

import CollectionDataControl from './CollectionDataControl';
import {
  CustomCollectionFieldType,
  ICustomCollectionField,
  ICollectionDataItem,
} from './types';

export {
  CustomCollectionFieldType,
};
export type { ICustomCollectionField, ICollectionDataItem };

/**
 * API-compatible options for our `PropertyPaneCollectionData` factory.
 * Mirrors the surface of `@pnp/spfx-property-controls` `PropertyFieldCollectionData`
 * so call sites swap with no shape change beyond `manageBtnLabel` → `manageBtnLabel`.
 */
export interface IPropertyPaneCollectionDataInternalProps {
  key: string;
  label: string;
  panelHeader: string;
  panelDescription?: string;
  manageBtnLabel: string;
  /**
   * Persisted rows. Loosely typed so call sites can pass their own typed
   * arrays (`ISortCollectionItem[]`, `IVerticalCollectionItem[]`, etc.)
   * without each interface needing an explicit `[key: string]: unknown`
   * index signature.
   */
  value: unknown[] | undefined;
  fields: ICustomCollectionField[];
  enableSorting?: boolean;
  disabled?: boolean;
}

interface IRenderContext extends IPropertyPaneCollectionDataInternalProps {}

/**
 * Custom replacement for `@pnp/spfx-property-controls/PropertyFieldCollectionData`.
 *
 * Why this exists: the PnP control ships pre-compiled `.module.css` files whose
 * JS `styles` import resolves to `undefined` under SPFx 1.22's `sp-css-loader`.
 * Their `CollectionDataViewer` then crashes inside React reconciliation with
 * "Cannot read properties of undefined (reading 'table')", breaking the entire
 * property pane (the spinner never finishes loading).
 *
 * This control renders a Fluent UI `Panel` with a table whose styles live in
 * a SCSS module under our own build pipeline, so the styles import is always
 * defined at runtime.
 */
export function PropertyPaneCollectionData(
  targetProperty: string,
  properties: IPropertyPaneCollectionDataInternalProps
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty,
    properties: {
      key: properties.key || 'collectionData_' + targetProperty,
      context: properties as IRenderContext,
      onRender: function (
        domElement: HTMLElement,
        context: IRenderContext,
        changeCallback?: (targetProperty?: string, newValue?: unknown) => void
      ): void {
        function handleChange(next: ICollectionDataItem[]): void {
          if (changeCallback) {
            changeCallback(targetProperty, next);
          }
        }
        const safeValue = (Array.isArray(context.value) ? context.value : []) as ICollectionDataItem[];

        ReactDom.render(
          React.createElement(CollectionDataControl, {
            label: context.label,
            panelHeader: context.panelHeader,
            panelDescription: context.panelDescription,
            manageButtonLabel: context.manageBtnLabel,
            value: safeValue,
            fields: context.fields,
            enableSorting: context.enableSorting,
            disabled: context.disabled,
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
