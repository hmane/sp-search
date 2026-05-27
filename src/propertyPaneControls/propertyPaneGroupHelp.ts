/**
 * T4.D11 — context-sensitive help links for property pane groups.
 *
 * SPFx doesn't expose a per-group "help icon" slot, so we add a
 * `PropertyPaneLink` as the first or last field of each group. The
 * link target is a deep anchor in `docs/admin-guide.md` (hosted in the
 * repo's GitHub raw view by default; admins can override the base URL
 * via `setPropertyPaneHelpBaseUrl` for tenants that mirror docs
 * elsewhere).
 *
 * Usage:
 *
 * ```typescript
 * import { propertyPaneGroupHelp } from '../../propertyPaneControls/propertyPaneGroupHelp';
 *
 * {
 *   groupName: 'Layouts',
 *   groupFields: [
 *     propertyPaneGroupHelp('results-layouts', 'Layouts and presets help'),
 *     // ... rest of the group fields
 *   ]
 * }
 * ```
 */

import {
  PropertyPaneLink,
  type IPropertyPaneField,
  type IPropertyPaneLinkProps,
} from '@microsoft/sp-property-pane';

/** Default — points at the repo's main-branch admin-guide.md. */
const DEFAULT_BASE_URL = 'https://github.com/hmane/sp-search/blob/main/docs/admin-guide.md';

let baseUrl: string = DEFAULT_BASE_URL;

/**
 * Override the help-link base URL. Useful for tenants that host
 * SP Search docs internally (e.g. a SharePoint wiki). Call once at
 * web part `onInit()` before `getPropertyPaneConfiguration` runs.
 */
export function setPropertyPaneHelpBaseUrl(url: string): void {
  if (url && url.trim().length > 0) {
    baseUrl = url.trim();
  }
}

/**
 * Build a `PropertyPaneLink` pointing at a deep anchor in the admin
 * guide. Use as a field inside `IPropertyPaneGroup.groupFields` —
 * either at the start (header link) or end (footer link) of the group.
 */
export function propertyPaneGroupHelp(
  anchorId: string,
  linkText: string
): IPropertyPaneField<IPropertyPaneLinkProps> {
  const href = baseUrl + '#' + anchorId;
  return PropertyPaneLink('help-' + anchorId, {
    text: linkText,
    href,
    target: '_blank',
  });
}
