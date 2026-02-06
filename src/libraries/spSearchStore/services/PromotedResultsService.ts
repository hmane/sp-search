import {
  IPromotedResultRule,
  IPromotedResultDisplay
} from '@interfaces/index';
import { IPromotedResultItem } from '@interfaces/IStoreSlices';
import { isInAudience } from './AudienceService';

/**
 * PromotedResultsService — evaluates promoted result rules against the current query
 * and returns matching promoted items. Rules are loaded from SearchConfiguration list.
 *
 * Evaluation:
 * 1. Filter out inactive rules
 * 2. Filter by date range (if specified)
 * 3. Filter by vertical scope (if specified)
 * 4. Match query text against rule matchValue using matchType
 * 5. Merge matched items, cap at maxResults
 */

/** Default maximum promoted results per query */
const DEFAULT_MAX_PROMOTED = 3;

/**
 * Evaluate a single rule's match condition against the query text.
 */
function evaluateMatch(queryText: string, matchType: string, matchValue: string): boolean {
  const normalizedQuery = queryText.toLowerCase().trim();
  const normalizedValue = matchValue.toLowerCase().trim();

  switch (matchType) {
    case 'contains':
      return normalizedQuery.indexOf(normalizedValue) >= 0;

    case 'equals':
      return normalizedQuery === normalizedValue;

    case 'regex':
      try {
        const regex = new RegExp(matchValue, 'i');
        return regex.test(queryText);
      } catch {
        return false;
      }

    case 'kql':
      // Simple KQL matching: split by spaces, all terms must be present
      const terms = normalizedValue.split(/\s+/).filter(function (t: string): boolean { return t.length > 0; });
      for (let i = 0; i < terms.length; i++) {
        if (normalizedQuery.indexOf(terms[i]) < 0) {
          return false;
        }
      }
      return terms.length > 0;

    default:
      return false;
  }
}

/**
 * Check if a rule is currently within its date range.
 */
function isWithinDateRange(
  startDate: string | undefined,
  endDate: string | undefined
): boolean {
  const now = new Date();

  if (startDate) {
    const start = new Date(startDate);
    if (isNaN(start.getTime()) || now < start) {
      return false;
    }
  }

  if (endDate) {
    const end = new Date(endDate);
    if (isNaN(end.getTime()) || now > end) {
      return false;
    }
  }

  return true;
}

/**
 * Evaluate all promoted result rules against the current query.
 *
 * @param rules - Loaded rules from SearchConfiguration
 * @param queryText - Current search query
 * @param currentVertical - Current vertical key
 * @param maxResults - Max promoted results to return (default 3)
 * @param userGroupIds - Current user's Azure AD group IDs for audience targeting
 * @returns Array of promoted result items to display
 */
export function evaluatePromotedResults(
  rules: IPromotedResultRule[],
  queryText: string,
  currentVertical: string,
  maxResults: number = DEFAULT_MAX_PROMOTED,
  userGroupIds?: string[]
): IPromotedResultItem[] {
  if (!rules || rules.length === 0 || !queryText || queryText === '*') {
    return [];
  }

  const matchedItems: Array<IPromotedResultDisplay & { ruleId: number }> = [];

  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];

    // Skip inactive rules
    if (!rule.isActive) {
      continue;
    }

    // Check date range
    if (!isWithinDateRange(
      rule.startDate ? String(rule.startDate) : undefined,
      rule.endDate ? String(rule.endDate) : undefined
    )) {
      continue;
    }

    // Check vertical scope
    if (rule.verticalScope && rule.verticalScope.length > 0) {
      let verticalMatch = false;
      for (let v = 0; v < rule.verticalScope.length; v++) {
        if (rule.verticalScope[v] === currentVertical) {
          verticalMatch = true;
          break;
        }
      }
      if (!verticalMatch) {
        continue;
      }
    }

    // Audience targeting: skip rules the current user is not targeted for
    if (rule.audienceGroups && rule.audienceGroups.length > 0) {
      if (!userGroupIds || !isInAudience(rule.audienceGroups, userGroupIds)) {
        continue;
      }
    }

    // Evaluate match condition
    if (!evaluateMatch(queryText, rule.matchType, rule.matchValue)) {
      continue;
    }

    // Add all promoted items from this rule
    for (let j = 0; j < rule.promotedItems.length; j++) {
      matchedItems.push({
        ...rule.promotedItems[j],
        ruleId: rule.id,
      });
    }
  }

  // Sort by position, then cap at maxResults
  matchedItems.sort(function (a, b): number {
    return a.position - b.position;
  });

  // Deduplicate by URL
  const seen = new Set<string>();
  const result: IPromotedResultItem[] = [];

  for (let i = 0; i < matchedItems.length && result.length < maxResults; i++) {
    const item = matchedItems[i];
    if (!seen.has(item.url)) {
      seen.add(item.url);
      result.push({
        title: item.title,
        url: item.url,
        description: item.description,
        iconUrl: item.imageUrl,
      });
    }
  }

  return result;
}
