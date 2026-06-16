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
  BadgeColor,
  BADGE_COLORS,
  PHASE_1_RENDERERS,
  applyColumnPropertySelection,
  IBadgeColorRule,
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
  tags: 'Tags / badges',
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
  { key: 'badge', text: 'Badges (colored)' },
];

const BADGE_COLOR_OPTIONS: IDropdownOption[] = BADGE_COLORS.map((c) => ({
  key: c,
  text: c.charAt(0).toUpperCase() + c.slice(1),
}));

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

  const updateRule = (idx: number, patch: Partial<IBadgeColorRule>): void => {
    const next = (draft.valueColorMap || []).map((r, i) => (i === idx ? { ...r, ...patch } : r));
    update({ valueColorMap: next });
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
          <>
            <TextField
              label="Split character"
              value={draft.splitDelimiter || ''}
              placeholder=";  ,  |"
              maxLength={8}
              description={'Character to split the value on. Use \\n for a new line. Leave blank for the default ( , and ; ).'}
              onChange={(_e, newValue): void => update({ splitDelimiter: newValue || undefined })}
            />
            <Dropdown
              label="Display style"
              selectedKey={draft.multiValueSeparator || 'comma'}
              options={SEPARATOR_OPTIONS}
              onChange={(_e, option): void => {
                if (option) {
                  update({ multiValueSeparator: option.key as MultiValueSeparator });
                }
              }}
            />
            {draft.multiValueSeparator === 'badge' && (
              <div className={styles.panelSection}>
                <Toggle
                  label="Auto-color other values"
                  checked={draft.autoColorUnmapped !== false}
                  onChange={(_e, checked): void => update({ autoColorUnmapped: !!checked })}
                />
                {/* Index keys are intentional: rules have no stable id, and a
                    value-based key would change on every keystroke and remount
                    the focused TextField. Inputs are controlled, so values stay
                    correct on add/remove. */}
                {(draft.valueColorMap || []).map((rule, idx) => (
                  <div key={'rule-' + String(idx)} className={styles.badgeRuleRow}>
                    <TextField
                      label={idx === 0 ? 'Value' : undefined}
                      ariaLabel={idx === 0 ? undefined : 'Value for rule ' + String(idx + 1)}
                      value={rule.value}
                      placeholder="Approved"
                      onChange={(_e, v): void => updateRule(idx, { value: v || '' })}
                    />
                    <Dropdown
                      label={idx === 0 ? 'Color' : undefined}
                      ariaLabel={idx === 0 ? undefined : 'Color for rule ' + String(idx + 1)}
                      selectedKey={rule.color}
                      options={BADGE_COLOR_OPTIONS}
                      onChange={(_e, option): void => { if (option) { updateRule(idx, { color: option.key as BadgeColor }); } }}
                    />
                    <TextField
                      label={idx === 0 ? 'Icon (optional)' : undefined}
                      ariaLabel={idx === 0 ? undefined : 'Icon for rule ' + String(idx + 1)}
                      value={rule.icon || ''}
                      placeholder="CheckMark"
                      onChange={(_e, v): void => updateRule(idx, { icon: v || undefined })}
                    />
                    <DefaultButton
                      text="Remove"
                      ariaLabel={'Remove rule ' + String(idx + 1)}
                      onClick={(): void => {
                        update({ valueColorMap: (draft.valueColorMap || []).filter((_r, i) => i !== idx) });
                      }}
                    />
                  </div>
                ))}
                <DefaultButton
                  text="Add value color"
                  iconProps={{ iconName: 'Add' }}
                  onClick={(): void => {
                    update({ valueColorMap: [...(draft.valueColorMap || []), { value: '', color: 'blue' }] });
                  }}
                />
              </div>
            )}
          </>
        )}
      </div>
    </Panel>
  );
};

export default ColumnConfigPanel;
