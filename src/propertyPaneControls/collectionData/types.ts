/**
 * API-compatible replacement for @pnp/spfx-property-controls' CustomCollectionFieldType.
 *
 * The PnP control ships pre-compiled CSS modules whose `styles` import returns
 * `undefined` under SPFx 1.22's sp-css-loader, crashing inside CollectionDataViewer
 * with "Cannot read properties of undefined (reading 'table')". This control
 * mirrors the PnP API surface so call sites can swap with minimal change.
 */
export enum CustomCollectionFieldType {
  string = 'string',
  number = 'number',
  boolean = 'boolean',
  dropdown = 'dropdown',
}

export interface ICustomCollectionFieldOption {
  key: string | number;
  text: string;
}

export interface ICustomCollectionField {
  id: string;
  title: string;
  type: CustomCollectionFieldType;
  required?: boolean;
  placeholder?: string;
  options?: ICustomCollectionFieldOption[];
  defaultValue?: unknown;
}

/**
 * Untyped row shape — kept permissive so existing call-site interfaces
 * (e.g. `ISelectedPropertyItem`, `IVerticalCollectionItem`) without an
 * explicit `[key: string]: unknown` index signature satisfy it structurally
 * via `unknown[]` rather than via a strict `Record<string, unknown>` check.
 */
export type ICollectionDataItem = Record<string, unknown>;
