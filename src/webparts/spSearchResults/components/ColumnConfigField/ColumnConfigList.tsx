import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { IColumnConfigItem, ColumnRenderer, ColumnVisibility } from './columnConfig';
import styles from './ColumnConfigField.module.scss';

export interface IColumnConfigListProps {
  items: IColumnConfigItem[];
  onEdit: (uniqueId: string) => void;
  onRemove: (uniqueId: string) => void;
  onMove: (uniqueId: string, direction: -1 | 1) => void;
}

const RENDERER_CHIP_LABEL: Record<ColumnRenderer, string> = {
  '': 'Auto',
  text: 'Text',
  date: 'Date',
  fileType: 'File type',
  fileSize: 'File size',
  url: 'URL',
  persona: 'Person',
  richText: 'Rich text',
  number: 'Number',
  tags: 'Tags',
  boolean: 'Boolean',
};

function visibilityChip(visibility: ColumnVisibility): { text: string; className: string; tooltip: string } {
  if (visibility === 'always') {
    return { text: 'Always', className: styles.chipAlways, tooltip: 'Always shown — not in the column chooser' };
  }
  if (visibility === 'defaultOff') {
    return { text: 'Off', className: styles.chipDefaultOff, tooltip: 'Hidden by default — admins can opt in' };
  }
  return { text: 'On', className: '', tooltip: 'Shown by default — admins can hide' };
}

export const ColumnConfigList: React.FC<IColumnConfigListProps> = (props) => {
  const { items, onEdit, onRemove, onMove } = props;

  if (items.length === 0) {
    return <p className={styles.columnListDescription}>No columns configured yet.</p>;
  }

  return (
    <div className={styles.columnList} role="list">
      {items.map((item, index) => {
        const visibility = visibilityChip(item.visibility);
        const canMoveUp = index > 0;
        const canMoveDown = index < items.length - 1;
        return (
          <div key={item.uniqueId} className={styles.columnRow} role="listitem">
            <div className={styles.columnRowReorder}>
              <IconButton
                iconProps={{ iconName: 'ChevronUpSmall' }}
                ariaLabel={'Move up: ' + item.property}
                disabled={!canMoveUp}
                onClick={(): void => onMove(item.uniqueId, -1)}
                styles={{ root: { height: 20, width: 20 } }}
              />
              <IconButton
                iconProps={{ iconName: 'ChevronDownSmall' }}
                ariaLabel={'Move down: ' + item.property}
                disabled={!canMoveDown}
                onClick={(): void => onMove(item.uniqueId, 1)}
                styles={{ root: { height: 20, width: 20 } }}
              />
            </div>
            <TooltipHost content={visibility.tooltip}>
              <span className={[styles.chip, visibility.className].filter(Boolean).join(' ')}>{visibility.text}</span>
            </TooltipHost>
            <div className={styles.columnRowMain}>
              <span className={styles.columnRowProperty}>{item.property}</span>
              {item.alias && item.alias !== item.property && (
                <span className={styles.columnRowAlias}>{item.alias}</span>
              )}
            </div>
            <div className={styles.columnRowChips}>
              <span className={styles.chip}>{RENDERER_CHIP_LABEL[item.renderer] || 'Auto'}</span>
            </div>
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              ariaLabel={'Edit column ' + item.property}
              onClick={(): void => onEdit(item.uniqueId)}
            />
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              ariaLabel={'Remove column ' + item.property}
              onClick={(): void => onRemove(item.uniqueId)}
            />
          </div>
        );
      })}
    </div>
  );
};

export default ColumnConfigList;
