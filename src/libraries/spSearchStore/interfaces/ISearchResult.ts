/**
 * Structured person information used throughout SP Search
 * for authors, editors, and user refiners.
 */
export interface IPersonaInfo {
  displayText: string;
  email: string;
  imageUrl?: string;
}

/**
 * Sort field configuration for server-side sorting via SortList.
 */
export interface ISortField {
  property: string;
  direction: 'Ascending' | 'Descending';
}

/**
 * Normalized search result — every data provider maps raw results
 * to this interface before dispatching to resultSlice.
 */
export interface ISearchResult {
  /** Unique key — WorkId or DocId from search results */
  key: string;
  title: string;
  url: string;
  /** HitHighlightedSummary with <mark> tags for keyword hits */
  summary: string;
  author: IPersonaInfo;
  created: string;   // ISO Date
  modified: string;  // ISO Date
  fileType: string;
  fileSize: number;
  siteName: string;
  siteUrl: string;
  thumbnailUrl: string;
  /** Dynamic managed property bag — raw values from search API */
  properties: Record<string, unknown>;
  // Collapsing / Thread folding
  isCollapsedGroup?: boolean;
  childResults?: ISearchResult[];
  groupCount?: number;
}
