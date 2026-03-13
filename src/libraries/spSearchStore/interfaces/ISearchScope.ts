/**
 * Search scope definition — maps to KQL path restrictions
 * or result source IDs for scoping queries.
 */
export interface ISearchScope {
  id: string;
  label: string;
  /** KQL path restriction, e.g. "Path:https://contoso.sharepoint.com/sites/hr" */
  kqlPath?: string;
  /** SharePoint result source GUID */
  resultSourceId?: string;
}

/**
 * Suggestion item displayed in the Search Box dropdown.
 */
export interface ISuggestion {
  displayText: string;
  /** Group label: "Recent", "People", "Files", "Trending" */
  groupName: string;
  iconName?: string;
  filePath?: string;
  secondaryText?: string;
  action?: () => void;
  removeAction?: () => Promise<void> | void;
  removeLabel?: string;
}
