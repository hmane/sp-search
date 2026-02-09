import { IFilterConfig } from './IFilterTypes';

/**
 * Vertical tab definition — each vertical can override the
 * query template, result source, data provider, and filters.
 */
export interface IVerticalDefinition {
  key: string;
  label: string;
  iconName?: string;
  /** Override global query template for this vertical */
  queryTemplate?: string;
  /** SharePoint result source GUID (SharePoint provider only) */
  resultSourceId?: string;
  /** Per-vertical data provider override, e.g. "graph" for People vertical */
  dataProviderId?: string;
  /** Vertical-specific filter configuration */
  filterConfig?: IFilterConfig[];
  /** Azure AD security group IDs for audience targeting */
  audienceGroups?: string[];
  /** If true, clicking the tab navigates to linkUrl instead of filtering */
  isLink?: boolean;
  /** URL to navigate to when isLink is true */
  linkUrl?: string;
  /** How to open the link: current tab or new tab */
  openBehavior?: 'currentTab' | 'newTab';
  sortOrder: number;
}
