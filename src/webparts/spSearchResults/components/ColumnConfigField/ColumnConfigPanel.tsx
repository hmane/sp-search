import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react/lib/ChoiceGroup';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import {
  IColumnConfigItem,
  IColumnPropertyOption,
  ColumnRenderer,
  ColumnVisibility,
  MultiValueSeparator,
  PHASE_1_RENDERERS,
  applyColumnPropertySelection,
} from './columnConfig';
import styles from './ColumnConfigField.module.scss';

export interface IColumnConfigPanelProps {
  /** Open state. Closed = unmounted heavy content. */
  isOpen: boolean;
  /** Item being edited; null when add-mode. */
  initialItem: IColumnConfigItem | null;
  /** Available properties (from selectedPropertiesCollection). */
  availableProperties: IColumnPropertyOption[];
  /** Already-taken properties (so add-mode disallows duplicates). */
  takenProperties: string[];
  onSave: (item: IColumnConfigItem) => void;
  onCancel: () => void;
}

const RENDERER_LABELS: Record<ColumnRenderer, string> = {
  '': 'Auto-detect',
  text: 'Text',
  date: 'Date',
  fileType: 'File type',
  fileSize: 'File size',
  url: 'URL',
  persona: 'Person',
  // Phase-2 renderers — defined so the dropdown can label saved values from
  // future builds, but not exposed in Phase-1's PHASE_1_RENDERERS list.
  richText: 'Rich text',
  number: 'Number',
  tags: 'Tags',
  boolean: 'Boolean',
};

const VISIBILITY_OPTIONS: IChoiceGroupOption[] = [
  { key: 'always', text: 'Always shown (not in column chooser)' },
  { key: 'defaultOn', text: 'Shown by default (admin can hide)' },
  { key: 'defaultOff', text: 'Hidden by default (admin can show)' },
];

const SEPARATOR_OPTIONS: IDropdownOption[] = [
  { key: 'comma', text: 'Comma-separated' },
  { key: 'newline', text: 'One per line' },
  { key: 'semicolon', text: 'Semicolon-separated' },
  { key: 'pill', text: 'Pills' },
];

function rendererSupportsMaxLength(renderer: ColumnRenderer): boolean {
  return renderer === 'text' || renderer === 'richText' || renderer === 'url';
}

function rendererSupportsSeparator(renderer: ColumnRenderer): boolean {
  return renderer === 'tags';
}

export const ColumnConfigPanel: React.FC<IColumnConfigPanelProps> = (props) => {
  const { isOpen, initialItem, availableProperties, takenProperties, onSave, onCancel } = props;

  const [draft, setDraft] = React.useState<IColumnConfigItem | null>(initialItem);

  React.useEffect((): void => {
    setDraft(initialItem);
  }, [initialItem, isOpen]);

  if (!draft) {
    return null;
  }

  const isAddMode = !initialItem?.property;

  const propertyOptions: IDropdownOption[] = availableProperties.map((p) => {
    const taken = takenProperties.indexOf(p.key.toLowerCase()) >= 0
      && p.key.toLowerCase() !== (initialItem?.property || '').toLowerCase();
    return {
      key: p.key,
      text: p.text + (taken ? ' (already in use)' : ''),
      data: p,
      disabled: taken,
    };
  });

  const rendererOptions: IDropdownOption[] = PHASE_1_RENDERERS.map((r) => ({
    key: r,
    text: RENDERER_LABELS[r],
  }));

  const handleSave = (): void => {
    if (!draft.property) {
      return;
    }
    onSave(draft);
  };

  const update = (patch: Partial<IColumnConfigItem>): void => {
    setDraft((current) => current ? ({ ...current, ...patch }) : current);
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onCancel}
      type={PanelType.medium}
      headerText={isAddMode ? 'Add column' : ('Edit column: ' + (draft.alias || draft.property))}
      closeButtonAriaLabel="Close"
      onRenderFooterContent={((): React.ReactElement => (
        <div className={styles.panelFooter}>
          <PrimaryButton onClick={handleSave} disabled={!draft.property}>Save</PrimaryButton>
          <DefaultButton onClick={onCancel}>Cancel</DefaultButton>
        </div>
      )) as unknown as React.ComponentProps<typeof Panel>['onRenderFooterContent']}
      isFooterAtBottom={true}
    >
      <div className={styles.panelSection}>
        <Dropdown
          label="Managed property"
          required={true}
          selectedKey={draft.property}
          options={propertyOptions}
          onChange={(_e, option): void => {
            if (option) {
              const source = (option.data as IColumnPropertyOption | undefined) || {
                key: String(option.key),
                text: option.text,
              };
              setDraft((current) => current ? applyColumnPropertySelection(current, source) : current);
            }
          }}
          disabled={!isAddMode}
        />
        <TextField
          label="Display label (alias)"
          value={draft.alias}
          placeholder={draft.property}
          onChange={(_e, newValue): void => update({ alias: newValue || '' })}
        />
        <TextField
          label="Width (px)"
          type="number"
          value={draft.width === undefined ? '' : String(draft.width)}
          placeholder="0 = auto"
          onChange={(_e, newValue): void => {
            const parsed = newValue ? parseInt(newValue, 10) : 0;
            update({ width: isFinite(parsed) && parsed > 0 ? parsed : undefined });
          }}
        />
        <ChoiceGroup
          label="Visibility"
          selectedKey={draft.visibility}
          options={VISIBILITY_OPTIONS}
          onChange={(_e, option): void => {
            if (option) {
              update({ visibility: option.key as ColumnVisibility });
            }
          }}
        />
        <Dropdown
          label="Renderer"
          selectedKey={draft.renderer}
          options={rendererOptions}
          onChange={(_e, option): void => {
            if (option) {
              update({ renderer: option.key as ColumnRenderer });
            }
          }}
        />
        {rendererSupportsMaxLength(draft.renderer) && (
          <>
            <TextField
              label="Maximum length (characters)"
              type="number"
              value={draft.maxLength === undefined ? '' : String(draft.maxLength)}
              placeholder="0 = no truncation"
              onChange={(_e, newValue): void => {
                const parsed = newValue ? parseInt(newValue, 10) : 0;
                update({ maxLength: isFinite(parsed) && parsed > 0 ? parsed : undefined });
              }}
            />
            <Toggle
              label='Append "See more" link when truncated'
              checked={!!draft.seeMoreLink}
              onChange={(_e, checked): void => update({ seeMoreLink: !!checked })}
            />
          </>
        )}
        {rendererSupportsSeparator(draft.renderer) && (
          <Dropdown
            label="Multi-value separator"
            selectedKey={draft.multiValueSeparator || 'comma'}
            options={SEPARATOR_OPTIONS}
            onChange={(_e, option): void => {
              if (option) {
                update({ multiValueSeparator: option.key as MultiValueSeparator });
              }
            }}
          />
        )}
      </div>
    </Panel>
  );
};

export default ColumnConfigPanel;
