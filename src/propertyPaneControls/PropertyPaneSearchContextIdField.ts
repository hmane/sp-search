import {
  PropertyPaneTextField,
  type IPropertyPaneField,
  type IPropertyPaneTextFieldProps,
} from '@microsoft/sp-property-pane';

/**
 * T3.D4 — single source of truth for the `searchContextId` property-pane
 * field that every SP Search web part exposes at page 1, group 1, first
 * field. Centralising this eliminates the four-way drift we had today
 * (different group placements, three different copy variants) and gives
 * the audit's Required scenarios S1 (multi-context authoring), S2
 * (mismatch detection), S3 (URL prefix collision) a consistent admin
 * surface to lean on.
 *
 * The label + description + error string are exported as constants so
 * that consumers (e.g. cross-pane validators in T3.D2) can reference
 * them without hard-coding the prose.
 */

/** Group name for the dedicated "Search context" property-pane group. */
export const SEARCH_CONTEXT_ID_GROUP_NAME = 'Search context';

/** Pane-field label, identical across all six web parts. */
export const SEARCH_CONTEXT_ID_LABEL = 'Search context ID';

/** Pane-field description — explains the multi-context wiring. */
export const SEARCH_CONTEXT_ID_DESCRIPTION =
  'Identifies which shared search state this web part participates in. ' +
  'Web parts with the same ID on the same page exchange queries, filters, ' +
  'and verticals through one Zustand store. Use distinct IDs (e.g. ' +
  '"hr-search" and "policy-search") to run independent search experiences ' +
  'side by side on a single page.';

/** Single error message admins see when the field is empty. */
export const SEARCH_CONTEXT_ID_REQUIRED_ERROR =
  'Required — enter an ID to connect search web parts on this page (e.g. "hr-search").';

/**
 * Build the `PropertyPaneTextField('searchContextId', ...)` field with the
 * canonical label / description / required-error wiring. Call this as the
 * first field of the first group on page 1 in every web part's
 * `getPropertyPaneConfiguration`.
 */
export function propertyPaneSearchContextIdField(): IPropertyPaneField<IPropertyPaneTextFieldProps> {
  return PropertyPaneTextField('searchContextId', {
    label: SEARCH_CONTEXT_ID_LABEL,
    description: SEARCH_CONTEXT_ID_DESCRIPTION,
    onGetErrorMessage: (value: string): string => {
      if (!value || value.trim() === '') {
        return SEARCH_CONTEXT_ID_REQUIRED_ERROR;
      }
      return '';
    },
  });
}
