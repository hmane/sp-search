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
  };
}

export interface IVisualFilterBuilderProps {
  refiners: IRefiner[];
  filterConfig: IFilterConfig[];
  activeFilters: IActiveFilter[];
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
  filterConfig: IFilterConfig[]
): IVisualFilterBuilderField[] {
  const configMap: Map<string, IFilterConfig> = new Map();
  for (let i = 0; i < filterConfig.length; i++) {
    configMap.set(filterConfig[i].managedProperty, filterConfig[i]);
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
    if (dataType === 'string' && refiner.values.length > 0) {
      const lookupData: Array<{ value: string; text: string }> = [];
      for (let j = 0; j < refiner.values.length; j++) {
        const rv: IRefinerValue = refiner.values[j];
        lookupData.push({
          value: rv.value,
          text: rv.name + (rv.count > 0 ? ' (' + String(rv.count) + ')' : ''),
        });
      }
      field.lookup = { dataSource: lookupData };
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

/**
 * Convert a single DevExtreme FilterBuilder condition to an IActiveFilter.
 * Returns undefined if the condition cannot be converted.
 */
function conditionToFilter(
  field: string,
  operator: string,
  value: any,
  dataType: FieldType
): IActiveFilter | undefined {
  if (value === null || value === undefined) {
    return undefined;
  }

  const op = operator.toLowerCase();
  let refinementValue: string;

  if (dataType === 'boolean') {
    refinementValue = value ? '1' : '0';
  } else if (dataType === 'date') {
    if (op === 'between' && Array.isArray(value) && value.length >= 2) {
      const start = value[0] instanceof Date ? value[0] : new Date(value[0]);
      const end = value[1] instanceof Date ? value[1] : new Date(value[1]);
      if (!isNaN(start.getTime()) && !isNaN(end.getTime())) {
        refinementValue = buildDateRangeToken(start, end);
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
    }
  } else if (dataType === 'number') {
    if (op === 'between' && Array.isArray(value) && value.length >= 2) {
      refinementValue = buildNumericRangeToken(Number(value[0]), Number(value[1]));
    } else if (op === '>=' || op === '>') {
      refinementValue = buildNumericRangeToken(Number(value), Number.MAX_SAFE_INTEGER);
    } else if (op === '<=' || op === '<') {
      refinementValue = buildNumericRangeToken(0, Number(value));
    } else {
      refinementValue = String(value);
    }
  } else {
    // String type
    if (op === 'anyof' && Array.isArray(value)) {
      // Multiple values — return multiple filters via caller
      return undefined;
    }
    if (op === 'contains' || op === 'notcontains') {
      refinementValue = escapeRefinementValue(String(value));
    } else {
      refinementValue = String(value);
    }
  }

  return {
    filterName: field,
    value: refinementValue,
    operator: 'AND',
  };
}

/**
 * Recursively extract IActiveFilter[] from a DevExtreme FilterBuilder value.
 * Handles AND/OR groups and nested expressions.
 */
function extractFiltersFromExpression(
  expression: FilterBuilderValue,
  fieldTypeMap: Map<string, FieldType>
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

    // Handle 'anyof' as multiple filters
    if (operator.toLowerCase() === 'anyof' && Array.isArray(value)) {
      const filters: IActiveFilter[] = [];
      for (let i = 0; i < value.length; i++) {
        filters.push({
          filterName: field,
          value: String(value[i]),
          operator: 'OR',
        });
      }
      return filters;
    }

    const filter = conditionToFilter(field, operator, value, dataType);
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
      const subFilters = extractFiltersFromExpression(expression[i], fieldTypeMap);
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
  fieldTypeMap: Map<string, FieldType>
): string {
  if (!expression) {
    return '';
  }
  if (!Array.isArray(expression)) {
    return '';
  }

  // Negation
  if (expression.length === 2 && expression[0] === '!') {
    const inner = buildKqlPreview(expression[1], fieldCaptionMap, fieldTypeMap);
    return inner ? 'NOT (' + inner + ')' : '';
  }

  // Single condition
  if (expression.length >= 3 && typeof expression[0] === 'string' && typeof expression[1] === 'string') {
    const field = expression[0] as string;
    const operator = expression[1] as string;
    const value = expression[2];
    const caption = fieldCaptionMap.get(field) || field;
    const dataType = fieldTypeMap.get(field) || 'string';

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
      return caption + ' in (' + value.join(', ') + ')';
    }

    let displayValue: string;
    if (dataType === 'boolean') {
      displayValue = value ? 'Yes' : 'No';
    } else if (dataType === 'date' && value) {
      displayValue = new Date(value).toLocaleDateString();
    } else {
      displayValue = String(value || '');
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
      const part = buildKqlPreview(expression[i], fieldCaptionMap, fieldTypeMap);
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
  const { refiners, filterConfig, onApplyFilters, onCancel } = props;

  const fields: IVisualFilterBuilderField[] = React.useMemo(
    () => buildFieldsFromRefiners(refiners, filterConfig),
    [refiners, filterConfig]
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

  const [builderValue, setBuilderValue] = React.useState<FilterBuilderValue>(null);
  const [kqlPreview, setKqlPreview] = React.useState<string>('');

  React.useEffect(() => {
    const preview = buildKqlPreview(builderValue, fieldCaptionMap, fieldTypeMap);
    setKqlPreview(preview);
  }, [builderValue, fieldCaptionMap, fieldTypeMap]);

  function handleApply(): void {
    const filters = extractFiltersFromExpression(builderValue, fieldTypeMap);
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
