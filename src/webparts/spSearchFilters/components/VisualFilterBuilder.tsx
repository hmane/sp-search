import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { createLazyComponent } from 'spfx-toolkit/lib/utilities/lazyLoader';
import styles from './SpSearchFilters.module.scss';
import type { IActiveFilter, IFilterConfig, IRefiner, IRefinerValue } from '@interfaces/index';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const FilterBuilder: any = createLazyComponent(
  () => import('devextreme-react/filter-builder').then((module) => ({
    default: module.FilterBuilder,
  })) as any,
  { errorMessage: 'Failed to load filter builder' }
);

type FilterBuilderValue = any;
type FieldType = 'string' | 'number' | 'date' | 'boolean';

interface IVisualFilterBuilderField {
  dataField: string;
  caption: string;
  dataType: FieldType;
  filterOperations: string[];
  defaultFilterOperation?: string;
  lookup?: {
    dataSource: Array<{ value: string; text: string }>;
    displayExpr: string;
    valueExpr: string;
  };
  editorOptions?: Record<string, unknown>;
}

interface IParsedDateRange {
  from: Date;
  to: Date;
}

function decodeHexToken(value: string): string {
  const stripped = value.charAt(0) === '"' && value.charAt(value.length - 1) === '"'
    ? value.substring(1, value.length - 1)
    : value;

  if (stripped.indexOf('\u01C2\u01C2') !== 0) {
    return stripped;
  }

  const hex = stripped.substring(2);
  if (!/^[0-9a-fA-F]+$/.test(hex) || hex.length % 2 !== 0) {
    return stripped;
  }

  try {
    const bytes: number[] = [];
    for (let i = 0; i < hex.length; i += 2) {
      bytes.push(parseInt(hex.substring(i, i + 2), 16));
    }
    return new TextDecoder().decode(new Uint8Array(bytes));
  } catch {
    return stripped;
  }
}

function extractEmailFromClaim(claim: string): string | undefined {
  const parts = claim.split('|');
  const last = parts[parts.length - 1];
  return last && last.indexOf('@') >= 0 ? last : undefined;
}

function resolveReadableLookupText(
  rawValue: string,
  fieldLookup: Map<string, string> | undefined
): string | undefined {
  const direct = fieldLookup?.get(rawValue);
  if (direct) {
    return direct;
  }

  const decoded = decodeHexToken(rawValue);
  if (decoded && decoded !== rawValue) {
    return decoded;
  }

  return extractEmailFromClaim(rawValue);
}

export interface IVisualFilterBuilderProps {
  refiners: IRefiner[];
  filterConfig: IFilterConfig[];
  activeFilters: IActiveFilter[];
  operatorBetweenFilters: 'AND' | 'OR';
  onApplyFilters: (filters: IActiveFilter[]) => void;
  onCancel: () => void;
}

/**
 * Determine field data type from the filter config and refiner values.
 */
function resolveFieldType(config: IFilterConfig | undefined): FieldType {
  if (!config) {
    return 'string';
  }
  const ft = config.filterType;
  if (ft === 'daterange') {
    return 'date';
  }
  if (ft === 'slider') {
    return 'number';
  }
  if (ft === 'toggle') {
    return 'boolean';
  }
  return 'string';
}

/**
 * Build field definitions for the DevExtreme FilterBuilder from available refiners.
 */
function buildFieldsFromRefiners(
  refiners: IRefiner[],
  filterConfig: IFilterConfig[],
  activeFilters: IActiveFilter[]
): IVisualFilterBuilderField[] {
  const configMap: Map<string, IFilterConfig> = new Map();
  for (let i = 0; i < filterConfig.length; i++) {
    configMap.set(filterConfig[i].managedProperty, filterConfig[i]);
  }

  const activeDisplayMap = new Map<string, Map<string, string>>();
  for (let i = 0; i < activeFilters.length; i++) {
    const filter = activeFilters[i];
    if (!filter.displayValue) {
      continue;
    }
    let perField = activeDisplayMap.get(filter.filterName);
    if (!perField) {
      perField = new Map<string, string>();
      activeDisplayMap.set(filter.filterName, perField);
    }
    perField.set(filter.value, filter.displayValue);
  }

  const fields: IVisualFilterBuilderField[] = [];
  for (let i = 0; i < refiners.length; i++) {
    const refiner = refiners[i];
    const config = configMap.get(refiner.filterName);
    const dataType = resolveFieldType(config);
    const caption = config ? config.displayName : refiner.filterName;

    let filterOperations: string[];
    let defaultFilterOperation: string | undefined;

    if (dataType === 'date') {
      filterOperations = ['=', '<>', '>', '>=', '<', '<=', 'between'];
      defaultFilterOperation = 'between';
    } else if (dataType === 'number') {
      filterOperations = ['=', '<>', '>', '>=', '<', '<=', 'between'];
      defaultFilterOperation = '=';
    } else if (dataType === 'boolean') {
      filterOperations = ['=', '<>'];
      defaultFilterOperation = '=';
    } else {
      filterOperations = ['=', '<>', 'contains', 'notcontains', 'anyof'];
      defaultFilterOperation = '=';
    }

    const field: IVisualFilterBuilderField = {
      dataField: refiner.filterName,
      caption,
      dataType,
      filterOperations,
      defaultFilterOperation,
    };

    // For string/checkbox/tagbox/taxonomy/people types, provide lookup values from refiner
    if (dataType === 'string') {
      const lookupData: Array<{ value: string; text: string }> = [];
      const seenValues = new Set<string>();

      for (let j = 0; j < refiner.values.length; j++) {
        const rv: IRefinerValue = refiner.values[j];
        const displayText = (rv.name && rv.name !== rv.value) ? rv.name : decodeHexToken(rv.value);
        lookupData.push({
          value: rv.value,
          text: displayText + (rv.count > 0 ? ' (' + String(rv.count) + ')' : ''),
        });
        seenValues.add(rv.value);
      }

      const activeDisplayValues = activeDisplayMap.get(refiner.filterName);
      if (activeDisplayValues) {
        activeDisplayValues.forEach(function (displayText: string, value: string): void {
          if (seenValues.has(value)) {
            return;
          }
          lookupData.push({
            value,
            text: displayText,
          });
        });
      }

      if (lookupData.length > 0) {
        field.lookup = {
          dataSource: lookupData,
          displayExpr: 'text',
          valueExpr: 'value',
        };
        field.editorOptions = {
          width: '100%',
          dropDownOptions: {
            width: 340,
            maxWidth: 'min(90vw, 340px)',
          },
        };
      }
    }

    fields.push(field);
  }

  fields.sort((a, b) => a.caption.localeCompare(b.caption));
  return fields;
}

/**
 * Escape a string value for FQL/refinement token usage.
 */
function escapeRefinementValue(value: string): string {
  return '"' + value.replace(/"/g, '""') + '"';
}

/**
 * Build an FQL date range token.
 */
function buildDateRangeToken(start: Date, end: Date): string {
  return 'range(datetime("' + start.toISOString() + '"), datetime("' + end.toISOString() + '"))';
}

/**
 * Build an FQL numeric range token.
 */
function buildNumericRangeToken(min: number, max: number): string {
  return 'range(decimal(' + String(min) + '), decimal(' + String(max) + '))';
}

function stripQuotedValue(value: string): string {
  if (value.indexOf('string("') === 0 && value.lastIndexOf('")') === value.length - 2) {
    return value.substring(8, value.length - 2);
  }
  if (value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') {
    return value.substring(1, value.length - 1);
  }
  return value;
}

function resolveLookupValue(
  rawValue: string,
  lookupValues: Set<string> | undefined
): string {
  if (!lookupValues || lookupValues.size === 0) {
    return stripQuotedValue(rawValue);
  }

  if (lookupValues.has(rawValue)) {
    return rawValue;
  }

  const stripped = stripQuotedValue(rawValue);
  if (lookupValues.has(stripped)) {
    return stripped;
  }

  return stripped;
}

function parseNumericRangeToken(rawValue: string): { min?: number; max?: number } | undefined {
  if (rawValue.indexOf('range(') !== 0 || rawValue[rawValue.length - 1] !== ')') {
    return undefined;
  }

  const inner = rawValue.substring(6, rawValue.length - 1);
  const parts = inner.split(',');
  if (parts.length < 2) {
    return undefined;
  }

  function parsePart(value: string): number | undefined {
    const trimmed = value.trim();
    if (trimmed === 'min' || trimmed === 'max') {
      return undefined;
    }
    const match = trimmed.match(/decimal\(([^)]+)\)/i);
    const candidate = match ? match[1] : trimmed;
    const parsed = parseFloat(candidate);
    return isNaN(parsed) ? undefined : parsed;
  }

  return {
    min: parsePart(parts[0]),
    max: parsePart(parts[1]),
  };
}

function parseDateRangeToken(rawValue: string): IParsedDateRange | undefined {
  const regex = /^range\(datetime\("([^"]+)"\),\s*datetime\("([^"]+)"\)\)$/;
  const match = rawValue.match(regex);
  if (!match) {
    return undefined;
  }

  const from = new Date(match[1]);
  const to = new Date(match[2]);
  if (isNaN(from.getTime()) || isNaN(to.getTime())) {
    return undefined;
  }

  return { from, to };
}

function buildExpressionFromActiveFilters(
  activeFilters: IActiveFilter[],
  filterConfig: IFilterConfig[],
  fieldTypeMap: Map<string, FieldType>,
  fieldLookupValueMap: Map<string, Set<string>>,
  operatorBetweenFilters: 'AND' | 'OR'
): FilterBuilderValue {
  if (!activeFilters || activeFilters.length === 0) {
    return null;
  }

  const configMap = new Map<string, IFilterConfig>();
  for (let i = 0; i < filterConfig.length; i++) {
    configMap.set(filterConfig[i].managedProperty, filterConfig[i]);
  }

  const grouped = new Map<string, IActiveFilter[]>();
  for (let i = 0; i < activeFilters.length; i++) {
    const filter = activeFilters[i];
    const existing = grouped.get(filter.filterName);
    if (existing) {
      existing.push(filter);
    } else {
      grouped.set(filter.filterName, [filter]);
    }
  }

  const expressions: FilterBuilderValue[] = [];

  grouped.forEach(function (filters: IActiveFilter[], filterName: string): void {
    const config = configMap.get(filterName);
    const fieldType = fieldTypeMap.get(filterName) || 'string';

    if (fieldType === 'date') {
      const parsed = parseDateRangeToken(filters[0].value);
      if (parsed) {
        expressions.push([filterName, 'between', [parsed.from, parsed.to]]);
      }
      return;
    }

    if (fieldType === 'number') {
      const parsed = parseNumericRangeToken(filters[0].value);
      if (parsed && parsed.min !== undefined && parsed.max !== undefined) {
        expressions.push([filterName, 'between', [parsed.min, parsed.max]]);
      }
      return;
    }

    if (fieldType === 'boolean') {
      const raw = stripQuotedValue(filters[0].value).toLowerCase();
      expressions.push([filterName, '=', raw === '1' || raw === 'true']);
      return;
    }

    if (config?.filterType === 'text') {
      expressions.push([filterName, 'contains', stripQuotedValue(filters[0].value)]);
      return;
    }

    if (filters.length > 1) {
      const orExpression: FilterBuilderValue[] = [];
      for (let i = 0; i < filters.length; i++) {
        if (i > 0) {
          orExpression.push('or');
        }
        orExpression.push([
          filterName,
          '=',
          resolveLookupValue(filters[i].value, fieldLookupValueMap.get(filterName))
        ]);
      }
      expressions.push(orExpression);
      return;
    }

    expressions.push([filterName, '=', resolveLookupValue(filters[0].value, fieldLookupValueMap.get(filterName))]);
  });

  if (expressions.length === 0) {
    return null;
  }

  if (expressions.length === 1) {
    return expressions[0];
  }

  const joinOperator = operatorBetweenFilters === 'OR' ? 'or' : 'and';

  const groupedExpression: FilterBuilderValue[] = [];
  for (let i = 0; i < expressions.length; i++) {
    if (i > 0) {
      groupedExpression.push(joinOperator);
    }
    groupedExpression.push(expressions[i]);
  }

  return groupedExpression;
}

/**
 * Convert a single DevExtreme FilterBuilder condition to an IActiveFilter.
 * Returns undefined if the condition cannot be converted.
 */
function conditionToFilter(
  field: string,
  operator: string,
  value: any,
  dataType: FieldType,
  fieldLookupTextMap: Map<string, Map<string, string>>
): IActiveFilter | undefined {
  if (value === null || value === undefined) {
    return undefined;
  }

  const op = operator.toLowerCase();
  let refinementValue: string;
  let displayValue: string | undefined;
  const fieldLookup = fieldLookupTextMap.get(field);

  if (dataType === 'boolean') {
    refinementValue = value ? '1' : '0';
    displayValue = value ? 'Yes' : 'No';
  } else if (dataType === 'date') {
    if (op === 'between' && Array.isArray(value) && value.length >= 2) {
      const start = value[0] instanceof Date ? value[0] : new Date(value[0]);
      const end = value[1] instanceof Date ? value[1] : new Date(value[1]);
      if (!isNaN(start.getTime()) && !isNaN(end.getTime())) {
        refinementValue = buildDateRangeToken(start, end);
        displayValue = start.toLocaleDateString() + ' - ' + end.toLocaleDateString();
      } else {
        return undefined;
      }
    } else {
      const dateValue = value instanceof Date ? value : new Date(value);
      if (isNaN(dateValue.getTime())) {
        return undefined;
      }
      if (op === '>=' || op === '>') {
        const farFuture = new Date('2099-12-31T23:59:59Z');
        refinementValue = buildDateRangeToken(dateValue, farFuture);
      } else if (op === '<=' || op === '<') {
        const farPast = new Date('1900-01-01T00:00:00Z');
        refinementValue = buildDateRangeToken(farPast, dateValue);
      } else {
        // Equals: create a one-day range
        const dayStart = new Date(dateValue);
        dayStart.setHours(0, 0, 0, 0);
        const dayEnd = new Date(dateValue);
        dayEnd.setHours(23, 59, 59, 999);
        refinementValue = buildDateRangeToken(dayStart, dayEnd);
      }
      displayValue = dateValue.toLocaleDateString();
    }
  } else if (dataType === 'number') {
    if (op === 'between' && Array.isArray(value) && value.length >= 2) {
      refinementValue = buildNumericRangeToken(Number(value[0]), Number(value[1]));
      displayValue = String(value[0]) + ' - ' + String(value[1]);
    } else if (op === '>=' || op === '>') {
      refinementValue = buildNumericRangeToken(Number(value), Number.MAX_SAFE_INTEGER);
      displayValue = String(value);
    } else if (op === '<=' || op === '<') {
      refinementValue = buildNumericRangeToken(0, Number(value));
      displayValue = String(value);
    } else {
      refinementValue = String(value);
      displayValue = String(value);
    }
  } else {
    // String type
    if (op === 'anyof' && Array.isArray(value)) {
      // Multiple values — return multiple filters via caller
      return undefined;
    }
    if (op === 'contains' || op === 'notcontains') {
      refinementValue = escapeRefinementValue(String(value));
      displayValue = String(value);
    } else {
      refinementValue = String(value);
      displayValue = resolveReadableLookupText(String(value), fieldLookup);
    }
  }

  return {
    filterName: field,
    value: refinementValue,
    displayValue,
    operator: 'AND',
  };
}

/**
 * Recursively extract IActiveFilter[] from a DevExtreme FilterBuilder value.
 * Handles AND/OR groups and nested expressions.
 */
function extractFiltersFromExpression(
  expression: FilterBuilderValue,
  fieldTypeMap: Map<string, FieldType>,
  fieldLookupTextMap: Map<string, Map<string, string>>
): IActiveFilter[] {
  if (!expression) {
    return [];
  }

  if (!Array.isArray(expression)) {
    return [];
  }

  // Negation: ['!', subExpression]
  if (expression.length === 2 && expression[0] === '!') {
    // Negations can't be represented as simple refinement filters — skip
    return [];
  }

  // Single condition: [field, operator, value]
  if (expression.length >= 3 && typeof expression[0] === 'string' && typeof expression[1] === 'string') {
    const field = expression[0] as string;
    const operator = expression[1] as string;
    const value = expression[2];
    const dataType = fieldTypeMap.get(field) || 'string';
    const fieldLookup = fieldLookupTextMap.get(field);

    // Handle 'anyof' as multiple filters
    if (operator.toLowerCase() === 'anyof' && Array.isArray(value)) {
      const filters: IActiveFilter[] = [];
      for (let i = 0; i < value.length; i++) {
        filters.push({
          filterName: field,
          value: String(value[i]),
          displayValue: resolveReadableLookupText(String(value[i]), fieldLookup),
          operator: 'OR',
        });
      }
      return filters;
    }

    const filter = conditionToFilter(field, operator, value, dataType, fieldLookupTextMap);
    return filter ? [filter] : [];
  }

  // Group: [condition1, 'and'|'or', condition2, 'and'|'or', ...]
  const filters: IActiveFilter[] = [];
  let groupOperator: 'AND' | 'OR' = 'AND';

  for (let i = 0; i < expression.length; i++) {
    if (i % 2 === 1) {
      // Odd indices are operators
      const op = typeof expression[i] === 'string' ? (expression[i] as string).toUpperCase() : 'AND';
      if (op === 'OR') {
        groupOperator = 'OR';
      }
    } else {
      // Even indices are conditions/sub-groups
      const subFilters = extractFiltersFromExpression(expression[i], fieldTypeMap, fieldLookupTextMap);
      for (let j = 0; j < subFilters.length; j++) {
        filters.push(subFilters[j]);
      }
    }
  }

  // Apply group operator to all extracted filters
  if (groupOperator === 'OR') {
    for (let i = 0; i < filters.length; i++) {
      filters[i] = {
        filterName: filters[i].filterName,
        value: filters[i].value,
        displayValue: filters[i].displayValue,
        operator: 'OR',
      };
    }
  }

  return filters;
}

/**
 * Generate a human-readable KQL preview from the filter builder expression.
 */
function buildKqlPreview(
  expression: FilterBuilderValue,
  fieldCaptionMap: Map<string, string>,
  fieldTypeMap: Map<string, FieldType>,
  fieldLookupTextMap: Map<string, Map<string, string>>
): string {
  if (!expression) {
    return '';
  }
  if (!Array.isArray(expression)) {
    return '';
  }

  // Negation
  if (expression.length === 2 && expression[0] === '!') {
    const inner = buildKqlPreview(expression[1], fieldCaptionMap, fieldTypeMap, fieldLookupTextMap);
    return inner ? 'NOT (' + inner + ')' : '';
  }

  // Single condition
  if (expression.length >= 3 && typeof expression[0] === 'string' && typeof expression[1] === 'string') {
    const field = expression[0] as string;
    const operator = expression[1] as string;
    const value = expression[2];
    const caption = fieldCaptionMap.get(field) || field;
    const dataType = fieldTypeMap.get(field) || 'string';
    const fieldLookup = fieldLookupTextMap.get(field);

    if (operator.toLowerCase() === 'between' && Array.isArray(value) && value.length >= 2) {
      const start = dataType === 'date'
        ? new Date(value[0]).toLocaleDateString()
        : String(value[0]);
      const end = dataType === 'date'
        ? new Date(value[1]).toLocaleDateString()
        : String(value[1]);
      return caption + ' between ' + start + ' and ' + end;
    }

    if (operator.toLowerCase() === 'anyof' && Array.isArray(value)) {
      const labels = value.map(function (entry: unknown): string {
        const raw = String(entry);
        return resolveReadableLookupText(raw, fieldLookup) || raw;
      });
      return caption + ' in (' + labels.join(', ') + ')';
    }

    let displayValue: string;
    if (dataType === 'boolean') {
      displayValue = value ? 'Yes' : 'No';
    } else if (dataType === 'date' && value) {
      displayValue = new Date(value).toLocaleDateString();
    } else {
      const raw = String(value || '');
      displayValue = resolveReadableLookupText(raw, fieldLookup) || raw;
    }

    return caption + ' ' + operator + ' ' + displayValue;
  }

  // Group
  const parts: string[] = [];
  let groupOp = 'AND';
  for (let i = 0; i < expression.length; i++) {
    if (i % 2 === 1) {
      groupOp = typeof expression[i] === 'string' ? (expression[i] as string).toUpperCase() : 'AND';
    } else {
      const part = buildKqlPreview(expression[i], fieldCaptionMap, fieldTypeMap, fieldLookupTextMap);
      if (part) {
        parts.push(part);
      }
    }
  }

  if (parts.length === 0) {
    return '';
  }
  if (parts.length === 1) {
    return parts[0];
  }
  return '(' + parts.join(' ' + groupOp + ' ') + ')';
}

const VisualFilterBuilder: React.FC<IVisualFilterBuilderProps> = (props: IVisualFilterBuilderProps): React.ReactElement => {
  const { refiners, filterConfig, activeFilters, operatorBetweenFilters, onApplyFilters, onCancel } = props;

  const fields: IVisualFilterBuilderField[] = React.useMemo(
    () => buildFieldsFromRefiners(refiners, filterConfig, activeFilters),
    [refiners, filterConfig, activeFilters]
  );

  const fieldTypeMap: Map<string, FieldType> = React.useMemo(() => {
    const map = new Map<string, FieldType>();
    for (let i = 0; i < fields.length; i++) {
      map.set(fields[i].dataField, fields[i].dataType);
    }
    return map;
  }, [fields]);

  const fieldCaptionMap: Map<string, string> = React.useMemo(() => {
    const map = new Map<string, string>();
    for (let i = 0; i < fields.length; i++) {
      map.set(fields[i].dataField, fields[i].caption);
    }
    return map;
  }, [fields]);

  const fieldLookupValueMap: Map<string, Set<string>> = React.useMemo(() => {
    const map = new Map<string, Set<string>>();
    for (let i = 0; i < fields.length; i++) {
      const field = fields[i];
      const values = new Set<string>();
      if (field.lookup?.dataSource) {
        for (let j = 0; j < field.lookup.dataSource.length; j++) {
          values.add(String(field.lookup.dataSource[j].value));
        }
      }
      map.set(field.dataField, values);
    }
    return map;
  }, [fields]);

  const fieldLookupTextMap: Map<string, Map<string, string>> = React.useMemo(() => {
    const map = new Map<string, Map<string, string>>();
    for (let i = 0; i < fields.length; i++) {
      const field = fields[i];
      const values = new Map<string, string>();
      if (field.lookup?.dataSource) {
        for (let j = 0; j < field.lookup.dataSource.length; j++) {
          values.set(
            String(field.lookup.dataSource[j].value),
            String(field.lookup.dataSource[j].text).replace(/\s+\(\d+\)$/, '')
          );
        }
      }
      map.set(field.dataField, values);
    }
    return map;
  }, [fields]);

  const initialBuilderValue = React.useMemo(
    () => buildExpressionFromActiveFilters(activeFilters, filterConfig, fieldTypeMap, fieldLookupValueMap, operatorBetweenFilters),
    [activeFilters, fieldLookupValueMap, fieldTypeMap, filterConfig, operatorBetweenFilters]
  );

  const [builderValue, setBuilderValue] = React.useState<FilterBuilderValue>(initialBuilderValue);
  const [kqlPreview, setKqlPreview] = React.useState<string>('');

  React.useEffect((): void => {
    setBuilderValue(initialBuilderValue);
  }, [initialBuilderValue]);

  React.useEffect(() => {
    const preview = buildKqlPreview(builderValue, fieldCaptionMap, fieldTypeMap, fieldLookupTextMap);
    setKqlPreview(preview);
  }, [builderValue, fieldCaptionMap, fieldLookupTextMap, fieldTypeMap]);

  function handleApply(): void {
    const filters = extractFiltersFromExpression(builderValue, fieldTypeMap, fieldLookupTextMap);
    onApplyFilters(filters);
  }

  function handleClear(): void {
    setBuilderValue(null);
    setKqlPreview('');
  }

  if (fields.length === 0) {
    return (
      <div className={styles.visualFilterBuilderPanel}>
        <div className={styles.visualFilterBuilderEmpty}>
          <Icon iconName="FilterSolid" className={styles.visualFilterBuilderEmptyIcon} />
          <span>No filter fields available. Perform a search to see available refiners.</span>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.visualFilterBuilderPanel}>
      <div className={styles.visualFilterBuilderHeader}>
        <Icon iconName="Filter" />
        <span>Visual Filter Builder</span>
      </div>
      <div className={styles.visualFilterBuilderBody}>
        <FilterBuilder
          fields={fields as any}
          value={builderValue}
          onValueChanged={(e: { value?: FilterBuilderValue }): void => {
            setBuilderValue(e.value || null);
          }}
          allowHierarchicalFields={false}
          groupOperationDescriptions={{
            and: 'AND — all conditions must match',
            or: 'OR — any condition can match',
            notAnd: 'NOT AND',
            notOr: 'NOT OR',
          }}
          filterOperationDescriptions={{
            equal: 'Equals',
            notEqual: 'Does not equal',
            greaterThan: 'Greater than',
            greaterThanOrEqual: 'Greater than or equal',
            lessThan: 'Less than',
            lessThanOrEqual: 'Less than or equal',
            between: 'Between',
            contains: 'Contains',
            notContains: 'Does not contain',
          }}
        />
      </div>
      {kqlPreview && (
        <div className={styles.visualFilterBuilderPreview}>
          <div className={styles.visualFilterBuilderPreviewLabel}>Filter expression preview</div>
          <div className={styles.visualFilterBuilderPreviewText}>{kqlPreview}</div>
        </div>
      )}
      <div className={styles.visualFilterBuilderActions}>
        <DefaultButton text="Cancel" onClick={onCancel} />
        <DefaultButton text="Clear" onClick={handleClear} />
        <PrimaryButton text="Apply Filters" onClick={handleApply} disabled={!kqlPreview} />
      </div>
    </div>
  );
};

export default VisualFilterBuilder;
