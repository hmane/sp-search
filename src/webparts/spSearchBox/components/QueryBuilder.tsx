import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { createLazyComponent } from 'spfx-toolkit/lib/utilities/lazyLoader';
import type { IManagedProperty } from '@interfaces/index';
import styles from './SpSearchBox.module.scss';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const FilterBuilder: any = createLazyComponent(
  () => import('devextreme-react/filter-builder').then((module) => ({
    default: module.FilterBuilder,
  })) as any,
  { errorMessage: 'Failed to load query builder' }
);

type FilterBuilderValue = any;

type FieldType = 'string' | 'number' | 'date' | 'boolean';

interface IQueryBuilderField {
  dataField: string;
  caption: string;
  dataType: FieldType;
  filterOperations: string[];
  defaultFilterOperation?: string;
}

export interface IQueryBuilderProps {
  properties: IManagedProperty[];
  isLoading: boolean;
  errorMessage?: string;
  onApply: (kql: string) => void;
  onClear: () => void;
}

function normalizeFieldType(type: string | undefined): FieldType {
  const normalized = (type || '').toLowerCase();
  if (normalized.indexOf('date') >= 0 || normalized.indexOf('time') >= 0) {
    return 'date';
  }
  if (normalized.indexOf('int') >= 0 || normalized.indexOf('double') >= 0 || normalized.indexOf('decimal') >= 0 || normalized.indexOf('number') >= 0) {
    return 'number';
  }
  if (normalized.indexOf('bool') >= 0) {
    return 'boolean';
  }
  return 'string';
}

function buildFields(properties: IManagedProperty[]): IQueryBuilderField[] {
  const fields: IQueryBuilderField[] = [];
  for (let i = 0; i < properties.length; i++) {
    const prop = properties[i];
    if (!prop.queryable) {
      continue;
    }
    const dataType = normalizeFieldType(prop.type);
    let filterOperations: string[] = [];
    let defaultFilterOperation: string | undefined;

    if (dataType === 'string') {
      filterOperations = ['contains', 'notcontains', 'startswith', 'endswith', '=', '<>'];
      defaultFilterOperation = 'contains';
    } else if (dataType === 'number' || dataType === 'date') {
      filterOperations = ['=', '<>', '>', '>=', '<', '<=', 'between'];
      defaultFilterOperation = '=';
    } else {
      filterOperations = ['=', '<>'];
      defaultFilterOperation = '=';
    }

    fields.push({
      dataField: prop.name,
      caption: prop.alias || prop.name,
      dataType,
      filterOperations,
      defaultFilterOperation,
    });
  }

  fields.sort((a, b) => a.caption.localeCompare(b.caption));
  return fields;
}

function escapeString(value: string): string {
  return value.replace(/"/g, '""');
}

function formatValue(value: any, dataType: FieldType): string {
  if (value === null || value === undefined) {
    return '""';
  }
  if (dataType === 'number') {
    return String(value);
  }
  if (dataType === 'boolean') {
    return value ? 'true' : 'false';
  }
  if (dataType === 'date') {
    const dateValue = value instanceof Date ? value : new Date(value);
    if (!isNaN(dateValue.getTime())) {
      return 'datetime("' + dateValue.toISOString() + '")';
    }
  }
  return '"' + escapeString(String(value)) + '"';
}

function formatValueList(values: any[], dataType: FieldType): string[] {
  const formatted: string[] = [];
  for (let i = 0; i < values.length; i++) {
    formatted.push(formatValue(values[i], dataType));
  }
  return formatted;
}

function buildCondition(field: string, operator: string, value: any, dataType: FieldType): string {
  const op = operator.toLowerCase();

  if (op === 'isblank') {
    return 'NOT ' + field + ':*';
  }
  if (op === 'isnotblank') {
    return field + ':*';
  }

  if (op === 'between' && Array.isArray(value) && value.length >= 2) {
    const start = formatValue(value[0], dataType);
    const end = formatValue(value[1], dataType);
    return '(' + field + '>=' + start + ' AND ' + field + '<=' + end + ')';
  }

  if ((op === 'anyof' || op === 'noneof') && Array.isArray(value)) {
    const list = formatValueList(value, dataType);
    const expression = field + ':(' + list.join(' OR ') + ')';
    return op === 'noneof' ? 'NOT ' + expression : expression;
  }

  if (op === 'contains') {
    return field + ':' + formatValue(value, 'string');
  }
  if (op === 'notcontains') {
    return 'NOT ' + field + ':' + formatValue(value, 'string');
  }
  if (op === 'startswith') {
    const raw = String(value || '');
    const escaped = escapeString(raw);
    return field + ':"' + escaped + '*"';
  }
  if (op === 'endswith') {
    const raw = String(value || '');
    const escaped = escapeString(raw);
    return field + ':"*' + escaped + '"';
  }

  if (op === '=') {
    return field + '=' + formatValue(value, dataType);
  }
  if (op === '<>') {
    return 'NOT ' + field + '=' + formatValue(value, dataType);
  }
  if (op === '>' || op === '>=' || op === '<' || op === '<=') {
    return field + op + formatValue(value, dataType);
  }

  return '';
}

function buildKql(expression: FilterBuilderValue, fieldMap: Map<string, FieldType>): string {
  if (!expression) {
    return '';
  }
  if (Array.isArray(expression)) {
    if (expression.length === 2 && expression[0] === '!') {
      const inner = buildKql(expression[1], fieldMap);
      return inner ? 'NOT (' + inner + ')' : '';
    }
    if (expression.length >= 3 && typeof expression[0] === 'string' && typeof expression[1] === 'string') {
      const field = expression[0] as string;
      const operator = expression[1] as string;
      const value = expression[2];
      const dataType = fieldMap.get(field) || 'string';
      return buildCondition(field, operator, value, dataType);
    }

    let result = '';
    for (let i = 0; i < expression.length; i++) {
      const token = expression[i];
      if (i % 2 === 0) {
        const part = buildKql(token, fieldMap);
        if (part) {
          result = result ? result + ' ' + part : part;
        }
      } else {
        const op = typeof token === 'string' ? token.toUpperCase() : 'AND';
        if (result) {
          result = result + ' ' + op + ' ';
        }
      }
    }
    if (result.indexOf(' ') >= 0) {
      return '(' + result + ')';
    }
    return result;
  }
  return '';
}

const QueryBuilder: React.FC<IQueryBuilderProps> = (props: IQueryBuilderProps): React.ReactElement => {
  const { properties, isLoading, errorMessage, onApply, onClear } = props;

  const fields = React.useMemo(() => buildFields(properties), [properties]);

  const fieldMap = React.useMemo(() => {
    const map = new Map<string, FieldType>();
    for (let i = 0; i < fields.length; i++) {
      map.set(fields[i].dataField, fields[i].dataType);
    }
    return map;
  }, [fields]);

  const [builderValue, setBuilderValue] = React.useState<FilterBuilderValue>(null);
  const [kqlPreview, setKqlPreview] = React.useState<string>('');

  React.useEffect(() => {
    const next = buildKql(builderValue, fieldMap);
    setKqlPreview(next);
  }, [builderValue, fieldMap]);

  function handleApply(): void {
    if (kqlPreview) {
      onApply(kqlPreview);
    }
  }

  function handleClear(): void {
    setBuilderValue(null);
    setKqlPreview('');
    onClear();
  }

  return (
    <div className={styles.queryBuilderPanel}>
      <div className={styles.queryBuilderHeader}>
        <span>Advanced Query Builder</span>
      </div>
      {isLoading && (
        <div className={styles.queryBuilderLoading}>
          <Spinner size={SpinnerSize.small} label="Loading managed properties..." />
        </div>
      )}
      {!isLoading && errorMessage && (
        <div className={styles.queryBuilderError} role="alert">{errorMessage}</div>
      )}
      {!isLoading && !errorMessage && fields.length === 0 && (
        <div className={styles.queryBuilderEmpty} role="status">No managed properties available.</div>
      )}
      {!isLoading && fields.length > 0 && (
        <FilterBuilder
          fields={fields as any}
          value={builderValue}
          onValueChanged={(e: { value?: FilterBuilderValue }): void => {
            setBuilderValue(e.value || null);
          }}
        />
      )}
      <div className={styles.queryBuilderPreview}>
        <div className={styles.queryBuilderPreviewLabel}>KQL Preview</div>
        <textarea
          className={styles.queryBuilderPreviewText}
          readOnly={true}
          value={kqlPreview}
          aria-label="KQL Preview"
        />
      </div>
      <div className={styles.queryBuilderActions}>
        <DefaultButton text="Clear" onClick={handleClear} />
        <PrimaryButton text="Apply" onClick={handleApply} disabled={!kqlPreview} />
      </div>
    </div>
  );
};

export default QueryBuilder;
