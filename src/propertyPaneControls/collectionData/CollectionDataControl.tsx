import * as React from 'react';
import { Label } from '@fluentui/react/lib/Label';
import { DefaultButton, PrimaryButton, IconButton, ActionButton } from '@fluentui/react/lib/Button';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Stack } from '@fluentui/react/lib/Stack';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import type { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Icon } from '@fluentui/react/lib/Icon';

import {
  CustomCollectionFieldType,
  ICustomCollectionField,
  ICollectionDataItem,
} from './types';
import styles from './CollectionDataControl.module.scss';

export interface ICollectionDataControlProps {
  label: string;
  panelHeader: string;
  panelDescription?: string;
  manageButtonLabel: string;
  value: ICollectionDataItem[];
  fields: ICustomCollectionField[];
  enableSorting?: boolean;
  disabled?: boolean;
  onChange: (newValue: ICollectionDataItem[]) => void;
}

type ErrorMap = Record<number, Record<string, string>>;

const ROW_KEY_PROP = '__sp_search_row_key';

interface IInternalRow extends ICollectionDataItem {
  [ROW_KEY_PROP]: string;
}

let rowKeyCounter = 0;
function nextRowKey(): string {
  rowKeyCounter += 1;
  return 'row_' + String(rowKeyCounter) + '_' + String(Date.now());
}

function cloneWithKeys(items: ICollectionDataItem[]): IInternalRow[] {
  return items.map((item) => ({ ...item, [ROW_KEY_PROP]: nextRowKey() }));
}

function stripKeys(rows: IInternalRow[]): ICollectionDataItem[] {
  return rows.map((row) => {
    const copy: ICollectionDataItem = { ...row };
    delete copy[ROW_KEY_PROP];
    return copy;
  });
}

function defaultValueForField(field: ICustomCollectionField): unknown {
  if (field.defaultValue !== undefined) {
    return field.defaultValue;
  }
  switch (field.type) {
    case CustomCollectionFieldType.boolean:
      return false;
    case CustomCollectionFieldType.number:
      return undefined;
    case CustomCollectionFieldType.dropdown:
      return field.options && field.options.length > 0 ? field.options[0].key : '';
    default:
      return '';
  }
}

function makeEmptyRow(fields: ICustomCollectionField[]): IInternalRow {
  const row: IInternalRow = { [ROW_KEY_PROP]: nextRowKey() };
  for (const field of fields) {
    row[field.id] = defaultValueForField(field);
  }
  return row;
}

function validateRow(
  row: IInternalRow,
  fields: ICustomCollectionField[]
): Record<string, string> {
  const errors: Record<string, string> = {};
  for (const field of fields) {
    if (!field.required) {
      continue;
    }
    const v = row[field.id];
    if (field.type === CustomCollectionFieldType.boolean) {
      continue;
    }
    if (v === undefined || v === null || v === '') {
      errors[field.id] = field.title + ' is required';
    }
  }
  return errors;
}

// ─── Field editor (declared before CollectionDataControl so the renderer
//     reference resolves before use) ───────────────────────────────────

interface ICollectionFieldEditorProps {
  field: ICustomCollectionField;
  value: unknown;
  onChange: (next: unknown) => void;
  hasError: boolean;
}

const CollectionFieldEditor: React.FC<ICollectionFieldEditorProps> = ({
  field,
  value,
  onChange,
  hasError,
}) => {
  switch (field.type) {
    case CustomCollectionFieldType.boolean:
      return (
        <Toggle
          checked={Boolean(value)}
          onChange={(_e, checked): void => onChange(checked === true)}
          ariaLabel={field.title}
        />
      );

    case CustomCollectionFieldType.number:
      return (
        <TextField
          type='number'
          value={value === undefined || value === null ? '' : String(value)}
          placeholder={field.placeholder}
          onChange={(_e, v): void => {
            if (v === undefined || v === '') {
              onChange(undefined);
              return;
            }
            const n = Number(v);
            onChange(Number.isNaN(n) ? undefined : n);
          }}
          ariaLabel={field.title}
          errorMessage={hasError ? ' ' : undefined}
        />
      );

    case CustomCollectionFieldType.dropdown: {
      const options: IDropdownOption[] = (field.options || []).map((o) => ({
        key: o.key,
        text: o.text,
      }));
      return (
        <Dropdown
          selectedKey={(value as string | number | undefined) ?? undefined}
          options={options}
          onChange={(_e, option): void => onChange(option?.key)}
          ariaLabel={field.title}
          errorMessage={hasError ? ' ' : undefined}
        />
      );
    }

    case CustomCollectionFieldType.string:
    default:
      return (
        <TextField
          value={value === undefined || value === null ? '' : String(value)}
          placeholder={field.placeholder}
          onChange={(_e, v): void => onChange(v ?? '')}
          ariaLabel={field.title}
          errorMessage={hasError ? ' ' : undefined}
        />
      );
  }
};

export const CollectionDataControl: React.FC<ICollectionDataControlProps> = (props) => {
  const {
    label,
    panelHeader,
    panelDescription,
    manageButtonLabel,
    value,
    fields,
    enableSorting,
    disabled,
    onChange,
  } = props;

  const [isPanelOpen, setIsPanelOpen] = React.useState<boolean>(false);
  const [draft, setDraft] = React.useState<IInternalRow[]>([]);
  const [errors, setErrors] = React.useState<ErrorMap>({});
  const [footerError, setFooterError] = React.useState<string>('');

  const openPanel = React.useCallback((): void => {
    setDraft(cloneWithKeys(Array.isArray(value) ? value : []));
    setErrors({});
    setFooterError('');
    setIsPanelOpen(true);
  }, [value]);

  const closePanel = React.useCallback((): void => {
    setIsPanelOpen(false);
  }, []);

  const updateCell = React.useCallback(
    (rowIndex: number, fieldId: string, newValue: unknown): void => {
      setDraft((prev) => {
        const next = prev.slice();
        next[rowIndex] = { ...next[rowIndex], [fieldId]: newValue };
        return next;
      });
      setErrors((prev) => {
        if (!prev[rowIndex] || !prev[rowIndex][fieldId]) {
          return prev;
        }
        const nextRowErrors = { ...prev[rowIndex] };
        delete nextRowErrors[fieldId];
        const nextMap: ErrorMap = { ...prev };
        if (Object.keys(nextRowErrors).length === 0) {
          delete nextMap[rowIndex];
        } else {
          nextMap[rowIndex] = nextRowErrors;
        }
        return nextMap;
      });
    },
    []
  );

  const addRow = React.useCallback((): void => {
    setDraft((prev) => prev.concat(makeEmptyRow(fields)));
  }, [fields]);

  const deleteRow = React.useCallback((rowIndex: number): void => {
    setDraft((prev) => prev.filter((_, i) => i !== rowIndex));
    setErrors({});
  }, []);

  const moveRow = React.useCallback(
    (rowIndex: number, direction: -1 | 1): void => {
      setDraft((prev) => {
        const target = rowIndex + direction;
        if (target < 0 || target >= prev.length) {
          return prev;
        }
        const next = prev.slice();
        const [moved] = next.splice(rowIndex, 1);
        next.splice(target, 0, moved);
        return next;
      });
      setErrors({});
    },
    []
  );

  const commitAndClose = React.useCallback((): void => {
    const nextErrors: ErrorMap = {};
    let anyError = false;
    for (let i = 0; i < draft.length; i++) {
      const rowErrors = validateRow(draft[i], fields);
      if (Object.keys(rowErrors).length > 0) {
        nextErrors[i] = rowErrors;
        anyError = true;
      }
    }
    if (anyError) {
      setErrors(nextErrors);
      setFooterError('Please fill in all required fields before saving.');
      return;
    }
    onChange(stripKeys(draft));
    setIsPanelOpen(false);
  }, [draft, fields, onChange]);

  const summary: string = !Array.isArray(value) || value.length === 0
    ? 'No items configured'
    : value.length + ' item' + (value.length === 1 ? '' : 's') + ' configured';

  return (
    <div className={styles.host}>
      <Label className={styles.label}>{label}</Label>
      <div className={styles.triggerRow}>
        <span className={styles.summary}>{summary}</span>
        <DefaultButton
          iconProps={{ iconName: 'Settings' }}
          text={manageButtonLabel}
          onClick={openPanel}
          disabled={disabled === true}
        />
      </div>

      <Panel
        isOpen={isPanelOpen}
        onDismiss={closePanel}
        headerText={panelHeader}
        type={PanelType.custom}
        customWidth='900px'
        isLightDismiss={false}
        closeButtonAriaLabel='Close'
        onRenderFooterContent={((): React.ReactElement => (
          <Stack horizontal verticalAlign='center' tokens={{ childrenGap: 8 }} className={styles.footerBar}>
            {footerError && <span className={styles.footerError}>{footerError}</span>}
            <DefaultButton text='Cancel' onClick={closePanel} />
            <PrimaryButton text='Save' onClick={commitAndClose} />
          </Stack>
        )) as unknown as import('@fluentui/react/lib/Panel').IPanelProps['onRenderFooterContent']}
        isFooterAtBottom
      >
        <div className={styles.panelBody}>
          {panelDescription && <p className={styles.panelDescription}>{panelDescription}</p>}

          <div className={styles.tableScroll}>
            <table className={styles.table} role='grid'>
              <thead className={styles.tableHead}>
                <tr>
                  {fields.map((f) => (
                    <th key={f.id} className={styles.tableHeadCell} scope='col'>
                      {f.title}
                      {f.required && <span className={styles.required} aria-hidden='true'>*</span>}
                    </th>
                  ))}
                  <th className={`${styles.tableHeadCell} ${styles.tableHeadActions}`} scope='col' aria-label='Row actions' />
                </tr>
              </thead>
              <tbody>
                {draft.length === 0 && (
                  <tr>
                    <td className={styles.emptyRow} colSpan={fields.length + 1}>
                      <Icon iconName='Info' style={{ marginRight: 6 }} />
                      No items yet. Click <strong>Add</strong> below to start.
                    </td>
                  </tr>
                )}
                {draft.map((row, rowIndex) => {
                  const rowErrors = errors[rowIndex] || {};
                  return (
                    <tr key={row[ROW_KEY_PROP] as string} className={styles.tableRow}>
                      {fields.map((field) => {
                        const fieldError = rowErrors[field.id];
                        return (
                          <td key={field.id} className={styles.tableCell}>
                            <CollectionFieldEditor
                              field={field}
                              value={row[field.id]}
                              onChange={(v): void => updateCell(rowIndex, field.id, v)}
                              hasError={!!fieldError}
                            />
                            {fieldError && <div className={styles.fieldError}>{fieldError}</div>}
                          </td>
                        );
                      })}
                      <td className={styles.tableCellActions}>
                        {enableSorting && (
                          <>
                            <IconButton
                              iconProps={{ iconName: 'Up' }}
                              ariaLabel='Move up'
                              title='Move up'
                              disabled={rowIndex === 0}
                              onClick={(): void => moveRow(rowIndex, -1)}
                            />
                            <IconButton
                              iconProps={{ iconName: 'Down' }}
                              ariaLabel='Move down'
                              title='Move down'
                              disabled={rowIndex === draft.length - 1}
                              onClick={(): void => moveRow(rowIndex, 1)}
                            />
                          </>
                        )}
                        <IconButton
                          iconProps={{ iconName: 'Delete' }}
                          ariaLabel='Delete row'
                          title='Delete row'
                          onClick={(): void => deleteRow(rowIndex)}
                        />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div className={styles.addRow}>
            <ActionButton
              iconProps={{ iconName: 'Add' }}
              text='Add'
              onClick={addRow}
            />
          </div>
        </div>
      </Panel>
    </div>
  );
};

export default CollectionDataControl;
