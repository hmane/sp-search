/**
 * Promoted result rule — stored in SearchConfiguration list
 * with ConfigType: PromotedResult. Evaluated client-side
 * after results load.
 */
export interface IPromotedResultRule {
  id: number;
  matchType: 'contains' | 'equals' | 'regex' | 'kql';
  matchValue: string;
  promotedItems: IPromotedResultDisplay[];
  /** Azure AD security group IDs for audience targeting */
  audienceGroups: string[];
  /** Optional start date for time-limited promotions */
  startDate: Date | undefined;
  /** Optional end date for time-limited promotions */
  endDate: Date | undefined;
  /** Restrict to specific verticals only */
  verticalScope: string[] | undefined;
  isActive: boolean;
}

/**
 * Display data for a single promoted result item.
 */
export interface IPromotedResultDisplay {
  url: string;
  title: string;
  description?: string;
  imageUrl?: string;
  /** Rank order within the promoted results block */
  position: number;
}
