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
  sortOrder: number;
}
