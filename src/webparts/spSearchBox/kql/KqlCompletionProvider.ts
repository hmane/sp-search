import type { IKqlCompletionContext, IKqlCompletion } from './KqlTypes';
import type { IManagedProperty, IRefiner } from '@interfaces/index';

/** Maximum number of completions to return per category. */
const MAX_COMPLETIONS: number = 12;

/**
 * Static well-known values for common managed properties.
 * Mirrors the pattern from ManagedPropertyProvider._getKnownValues().
 */
const KNOWN_VALUES: Record<string, string[]> = {
  filetype: ['docx', 'xlsx', 'pptx', 'pdf', 'aspx', 'msg', 'txt', 'csv', 'jpg', 'png', 'mp4', 'one'],
  fileextension: ['docx', 'xlsx', 'pptx', 'pdf', 'aspx', 'msg', 'txt', 'csv', 'jpg', 'png', 'mp4', 'one'],
  contentclass: [
    'STS_ListItem_DocumentLibrary',
    'STS_Site',
    'STS_Web',
    'STS_ListItem',
    'STS_ListItem_GenericList',
    'STS_ListItem_Events',
    'STS_List_850',
  ],
  iscontainer: ['true', 'false'],
  isdocument: ['true', 'false'],
};

/**
 * Boolean connective completions.
 */
const CONNECTIVE_COMPLETIONS: IKqlCompletion[] = [
  { insertText: 'AND ', displayText: 'AND', completionType: 'keyword', description: 'Both conditions must match', iconName: 'MergeDuplicate' },
  { insertText: 'OR ', displayText: 'OR', completionType: 'keyword', description: 'Either condition can match', iconName: 'BranchMerge' },
  { insertText: 'NOT ', displayText: 'NOT', completionType: 'keyword', description: 'Exclude matching results', iconName: 'Cancel' },
  { insertText: '(', displayText: '( ... )', completionType: 'keyword', description: 'Group conditions', iconName: 'Code' },
];

/**
 * Returns auto-completion items based on the parsed completion context,
 * available schema properties, and current refiner values.
 *
 * All computation is synchronous — no API calls. Data must be pre-loaded.
 */
export function getCompletions(
  context: IKqlCompletionContext,
  schema: IManagedProperty[],
  refiners: IRefiner[]
): IKqlCompletion[] {
  switch (context.type) {
    case 'PropertyName':
      return getPropertyCompletions(context.partialText, schema);
    case 'PropertyValue':
      return getValueCompletions(context.partialText, context.propertyName || '', refiners);
    case 'BooleanConnective':
      return getConnectiveCompletions(context.partialText);
    case 'FreeText':
      return getPropertyCompletions(context.partialText, schema);
    default:
      return [];
  }
}

/**
 * Property name completions — fuzzy match against schema.
 * Auto-appends ':' to the insert text.
 */
function getPropertyCompletions(partial: string, schema: IManagedProperty[]): IKqlCompletion[] {
  const normalized: string = partial.toLowerCase();
  const exactPrefix: IKqlCompletion[] = [];
  const containsMatch: IKqlCompletion[] = [];

  for (let i: number = 0; i < schema.length; i++) {
    const prop: IManagedProperty = schema[i];
    if (!prop.queryable) {
      continue;
    }

    const nameLower: string = prop.name.toLowerCase();
    const aliasLower: string = (prop.alias || prop.name).toLowerCase();

    let isExact: boolean = false;
    let isContains: boolean = false;

    if (normalized.length === 0) {
      isExact = true;
    } else if (nameLower.indexOf(normalized) === 0 || aliasLower.indexOf(normalized) === 0) {
      isExact = true;
    } else if (nameLower.indexOf(normalized) >= 0 || aliasLower.indexOf(normalized) >= 0) {
      isContains = true;
    }

    if (!isExact && !isContains) {
      continue;
    }

    const displayText: string = prop.alias && prop.alias !== prop.name
      ? prop.alias + ' (' + prop.name + ')'
      : prop.name;

    const completion: IKqlCompletion = {
      insertText: prop.name + ':',
      displayText: displayText,
      completionType: 'property',
      propertyType: normalizePropertyType(prop.type),
      iconName: getPropertyIcon(prop.type),
      description: prop.refinable ? 'Refinable' : (prop.sortable ? 'Sortable' : undefined),
    };

    if (isExact) {
      exactPrefix.push(completion);
    } else {
      containsMatch.push(completion);
    }

    if (exactPrefix.length + containsMatch.length >= MAX_COMPLETIONS) {
      break;
    }
  }

  // Exact prefix matches first, then contains matches
  const result: IKqlCompletion[] = exactPrefix.concat(containsMatch);
  return result.slice(0, MAX_COMPLETIONS);
}

/**
 * Property value completions — from refiner data + static well-known values.
 */
function getValueCompletions(partial: string, propertyName: string, refiners: IRefiner[]): IKqlCompletion[] {
  const completions: IKqlCompletion[] = [];
  const normalizedPartial: string = partial.toLowerCase();
  const seen: Set<string> = new Set();

  // Source 1: Live refiner values (highest priority — they have counts)
  const propLower: string = propertyName.toLowerCase();
  for (let r: number = 0; r < refiners.length; r++) {
    if (refiners[r].filterName.toLowerCase() === propLower) {
      const values = refiners[r].values;
      for (let v: number = 0; v < values.length; v++) {
        const refVal = values[v];
        const nameLower: string = refVal.name.toLowerCase();
        if (normalizedPartial.length === 0 || nameLower.indexOf(normalizedPartial) >= 0) {
          const insertText: string = refVal.name.indexOf(' ') >= 0
            ? '"' + refVal.name + '"'
            : refVal.name;
          if (!seen.has(insertText.toLowerCase())) {
            seen.add(insertText.toLowerCase());
            completions.push({
              insertText: insertText,
              displayText: refVal.name,
              completionType: 'value',
              count: refVal.count,
              iconName: 'Tag',
            });
          }
        }
        if (completions.length >= MAX_COMPLETIONS) {
          break;
        }
      }
      break;
    }
  }

  // Source 2: Static well-known values
  const knownValues: string[] | undefined = KNOWN_VALUES[propLower];
  if (knownValues && completions.length < MAX_COMPLETIONS) {
    for (let k: number = 0; k < knownValues.length; k++) {
      const val: string = knownValues[k];
      if (!seen.has(val.toLowerCase()) && (normalizedPartial.length === 0 || val.toLowerCase().indexOf(normalizedPartial) >= 0)) {
        seen.add(val.toLowerCase());
        completions.push({
          insertText: val,
          displayText: val,
          completionType: 'value',
          iconName: 'Tag',
        });
        if (completions.length >= MAX_COMPLETIONS) {
          break;
        }
      }
    }
  }

  // Sort: items with counts first (desc), then alphabetical
  completions.sort(function (a: IKqlCompletion, b: IKqlCompletion): number {
    if (a.count !== undefined && b.count !== undefined) {
      return b.count - a.count;
    }
    if (a.count !== undefined) {
      return -1;
    }
    if (b.count !== undefined) {
      return 1;
    }
    return a.displayText.localeCompare(b.displayText);
  });

  return completions;
}

/**
 * Boolean connective completions — filter by partial text.
 */
function getConnectiveCompletions(partial: string): IKqlCompletion[] {
  if (partial.length === 0) {
    return CONNECTIVE_COMPLETIONS;
  }

  const normalized: string = partial.toUpperCase();
  const results: IKqlCompletion[] = [];

  for (let i: number = 0; i < CONNECTIVE_COMPLETIONS.length; i++) {
    const c: IKqlCompletion = CONNECTIVE_COMPLETIONS[i];
    if (c.displayText.indexOf(normalized) === 0) {
      results.push(c);
    }
  }

  return results;
}

/**
 * Normalizes SP managed property type to a display string.
 */
function normalizePropertyType(type: string): string {
  if (!type) {
    return 'Text';
  }
  const upper: string = type.toLowerCase();
  if (upper === 'datetime' || upper === 'date') {
    return 'DateTime';
  }
  if (upper === 'integer' || upper === 'int32' || upper === 'int64') {
    return 'Integer';
  }
  if (upper === 'double' || upper === 'decimal' || upper === 'float') {
    return 'Double';
  }
  if (upper === 'boolean' || upper === 'yesno') {
    return 'Boolean';
  }
  return 'Text';
}

/**
 * Returns a Fluent UI icon name based on property type.
 */
function getPropertyIcon(type: string): string {
  const normalized: string = normalizePropertyType(type);
  switch (normalized) {
    case 'DateTime': return 'Calendar';
    case 'Integer': return 'NumberField';
    case 'Double': return 'NumberField';
    case 'Boolean': return 'ToggleRight';
    default: return 'Variable2';
  }
}
