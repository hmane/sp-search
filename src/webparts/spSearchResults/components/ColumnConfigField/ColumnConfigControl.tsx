import * as React from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import {
  IColumnConfigItem,
  normalizeColumnConfigItem,
  generateColumnUniqueId,
} from './columnConfig';
import ColumnConfigList from './ColumnConfigList';
import ColumnConfigPanel from './ColumnConfigPanel';
import styles from './ColumnConfigField.module.scss';

export interface IColumnConfigControlProps {
  label: string;
  description?: string;
  value: IColumnConfigItem[];
  /** Available managed properties (sourced from selectedPropertiesCollection). */
  availableProperties: Array<{ key: string; text: string }>;
  onChange: (newValue: IColumnConfigItem[]) => void;
}

interface IEditorState {
  isOpen: boolean;
  /** Item currently being edited; the empty `IColumnConfigItem` in add-mode. */
  editing: IColumnConfigItem | null;
  /** uniqueId of the existing item being edited; undefined in add-mode. */
  editingId?: string;
}

const EMPTY_NEW_ITEM = (): IColumnConfigItem =>
  normalizeColumnConfigItem({ uniqueId: generateColumnUniqueId(), property: '' });

export const ColumnConfigControl: React.FC<IColumnConfigControlProps> = (props) => {
  const { label, description, value, availableProperties, onChange } = props;

  const [editor, setEditor] = React.useState<IEditorState>({ isOpen: false, editing: null });

  const handleAdd = (): void => {
    setEditor({ isOpen: true, editing: EMPTY_NEW_ITEM(), editingId: undefined });
  };

  const handleEdit = (uniqueId: string): void => {
    const found = value.find((item) => item.uniqueId === uniqueId);
    if (found) {
      setEditor({ isOpen: true, editing: { ...found }, editingId: uniqueId });
    }
  };

  const handleCancel = (): void => {
    setEditor({ isOpen: false, editing: null });
  };

  const handleSave = (item: IColumnConfigItem): void => {
    const normalized = normalizeColumnConfigItem(item);
    if (editor.editingId) {
      // Edit existing — keep position.
      onChange(value.map((existing) => existing.uniqueId === editor.editingId ? normalized : existing));
    } else {
      // Add new — append.
      onChange([...value, normalized]);
    }
    setEditor({ isOpen: false, editing: null });
  };

  const handleRemove = (uniqueId: string): void => {
    onChange(value.filter((item) => item.uniqueId !== uniqueId));
  };

  const handleMove = (uniqueId: string, direction: -1 | 1): void => {
    const index = value.findIndex((item) => item.uniqueId === uniqueId);
    if (index < 0) {
      return;
    }
    const target = index + direction;
    if (target < 0 || target >= value.length) {
      return;
    }
    const next = value.slice();
    const [moved] = next.splice(index, 1);
    next.splice(target, 0, moved);
    onChange(next);
  };

  // Properties already used by *other* columns — drives the disabled state in
  // the panel's property dropdown so add-mode can't pick a duplicate.
  const takenProperties = value
    .filter((item) => item.uniqueId !== editor.editingId)
    .map((item) => item.property.toLowerCase());

  return (
    <div>
      <label className={styles.columnListLabel}>{label}</label>
      {description && <p className={styles.columnListDescription}>{description}</p>}
      <ColumnConfigList
        items={value}
        onEdit={handleEdit}
        onRemove={handleRemove}
        onMove={handleMove}
      />
      <DefaultButton
        className={styles.addButton}
        iconProps={{ iconName: 'Add' }}
        text="Add column"
        onClick={handleAdd}
      />
      <ColumnConfigPanel
        isOpen={editor.isOpen}
        initialItem={editor.editing}
        availableProperties={availableProperties}
        takenProperties={takenProperties}
        onSave={handleSave}
        onCancel={handleCancel}
      />
    </div>
  );
};

export default ColumnConfigControl;
