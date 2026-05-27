import * as React from 'react';
import { Label } from '@fluentui/react/lib/Label';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import type { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { SpinButton } from '@fluentui/react/lib/SpinButton';
import { DefaultButton, PrimaryButton, IconButton, ActionButton } from '@fluentui/react/lib/Button';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Stack } from '@fluentui/react/lib/Stack';
import { Icon } from '@fluentui/react/lib/Icon';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Separator } from '@fluentui/react/lib/Separator';

import {
  getRelevantSections,
  isFieldRelevant,
} from './fieldRelevance';
import type { FilterEditorSection } from './fieldRelevance';
import styles from './FiltersCollectionControl.module.scss';

// ─── Public types ───────────────────────────────────────────

export interface IFiltersCollectionItem {
  uniqueId: string;
  managedProperty: string;
  displayName: string;
  urlAlias?: string;
  filterType: string;
  operator: string;
  maxValues: number;
  defaultExpanded: boolean;
  showCount: boolean;
  sortBy: string;
  sortDirection: string;
  multiValues: boolean;
  dependsOn?: string;
  showWhenParentHasValue?: boolean;
  hideZeroCountValues?: boolean;
  resetWhenParentChanges?: boolean;
  trueLabel?: string;
  falseLabel?: string;
  invertBoolean?: boolean;
  audience?: string;
}

export interface IFiltersCollectionControlProps {
  label: string;
  panelHeader: string;
  manageButtonLabel: string;
  value: IFiltersCollectionItem[];
  onChange: (next: IFiltersCollectionItem[]) => void;
}

// ─── Constants ──────────────────────────────────────────────

const FILTER_TYPE_OPTIONS: IDropdownOption[] = [
  { key: 'checkbox', text: 'Checkbox list' },
  { key: 'dropdown', text: 'Dropdown' },
  { key: 'daterange', text: 'Date range' },
  { key: 'text', text: 'Text query' },
  { key: 'people', text: 'People' },
  { key: 'taxonomy', text: 'Taxonomy' },
  { key: 'slider', text: 'Slider' },
  { key: 'tagbox', text: 'Tag box' },
  { key: 'toggle', text: 'Toggle' },
];

const OPERATOR_OPTIONS: IDropdownOption[] = [
  { key: 'OR', text: 'OR (match any selection)' },
  { key: 'AND', text: 'AND (match every selection)' },
];

const SORT_BY_OPTIONS: IDropdownOption[] = [
  { key: 'count', text: 'Count' },
  { key: 'alphabetical', text: 'Alphabetical' },
];

const SORT_DIR_OPTIONS: IDropdownOption[] = [
  { key: 'desc', text: 'Descending' },
  { key: 'asc', text: 'Ascending' },
];

const FILTER_TYPE_TEXT: { [k: string]: string } = {
  checkbox: 'Checkbox',
  dropdown: 'Dropdown',
  daterange: 'Date range',
  text: 'Text',
  people: 'People',
  taxonomy: 'Taxonomy',
  slider: 'Slider',
  tagbox: 'Tag box',
  toggle: 'Toggle',
};

function makeBlankRefiner(): IFiltersCollectionItem {
  return {
    uniqueId: 'refiner-' + String(Date.now()) + '-' + Math.random().toString(36).slice(2, 8),
    managedProperty: '',
    displayName: '',
    urlAlias: '',
    filterType: 'checkbox',
    operator: 'OR',
    maxValues: 10,
    defaultExpanded: true,
    showCount: true,
    sortBy: 'count',
    sortDirection: 'desc',
    multiValues: true,
    dependsOn: '',
    showWhenParentHasValue: false,
    hideZeroCountValues: false,
    resetWhenParentChanges: false,
    trueLabel: '',
    falseLabel: '',
    invertBoolean: false,
    audience: '',
  };
}

// ─── Collapsible section ───────────────────────────────────

interface ISectionProps {
  title: string;
  defaultCollapsed?: boolean;
  children: React.ReactNode;
}

const Section: React.FC<ISectionProps> = (props) => {
  const [collapsed, setCollapsed] = React.useState<boolean>(!!props.defaultCollapsed);
  return (
    <div className={styles.section}>
      <button
        type='button'
        className={styles.sectionHeader}
        onClick={(): void => setCollapsed(!collapsed)}
        aria-expanded={!collapsed}
      >
        <Icon iconName={collapsed ? 'ChevronRight' : 'ChevronDown'} className={styles.sectionChevron} />
        <span className={styles.sectionTitle}>{props.title}</span>
      </button>
      {!collapsed && (
        <div className={styles.sectionBody}>
          {props.children}
          <Separator />
        </div>
      )}
    </div>
  );
};

// ─── Right-pane form ─────────────────────────────────────────

interface IFilterEditorFormProps {
  refiner: IFiltersCollectionItem;
  sections: FilterEditorSection[];
  onPatch: (patch: Partial<IFiltersCollectionItem>) => void;
}

const FilterEditorForm: React.FC<IFilterEditorFormProps> = (formProps) => {
  const { refiner, sections, onPatch } = formProps;
  const ft = refiner.filterType;

  function renderBasic(): React.ReactElement {
    return (
      <Section title='Basic' key='basic'>
        <TextField
          label='Managed property'
          required
          value={refiner.managedProperty}
          onChange={(_e, v): void => onPatch({ managedProperty: v || '' })}
          placeholder='RefinableString00'
          description='SharePoint refinable managed property to refine on.'
        />
        <TextField
          label='Display name'
          required
          value={refiner.displayName}
          onChange={(_e, v): void => onPatch({ displayName: v || '' })}
          placeholder='File Type'
          description='Shown above the filter on the page.'
        />
        <TextField
          label='URL alias'
          value={refiner.urlAlias || ''}
          onChange={(_e, v): void => onPatch({ urlAlias: v || '' })}
          placeholder='ft'
          description='Short URL parameter name for sharing/deep-linking. Leave blank to auto-generate.'
        />
        <Dropdown
          label='Filter type'
          required
          selectedKey={refiner.filterType}
          options={FILTER_TYPE_OPTIONS}
          onChange={(_e, opt): void => { if (opt) { onPatch({ filterType: String(opt.key) }); } }}
        />
        <Toggle
          label='Expanded by default'
          checked={refiner.defaultExpanded}
          onChange={(_e, checked): void => onPatch({ defaultExpanded: !!checked })}
          onText='Yes'
          offText='No'
        />
      </Section>
    );
  }

  function renderDisplay(): React.ReactElement {
    return (
      <Section title='Display' key='display'>
        {isFieldRelevant('maxValues', ft) && (
          <SpinButton
            label='Max values to show'
            value={String(refiner.maxValues)}
            min={1}
            max={500}
            step={1}
            onChange={(_e, v): void => {
              const n = v !== undefined ? Number(v) : refiner.maxValues;
              onPatch({ maxValues: Number.isFinite(n) ? n : refiner.maxValues });
            }}
          />
        )}
        {isFieldRelevant('showCount', ft) && (
          <Toggle
            label='Show counts'
            checked={refiner.showCount}
            onChange={(_e, checked): void => onPatch({ showCount: !!checked })}
            onText='Yes' offText='No'
          />
        )}
        {isFieldRelevant('hideZeroCountValues', ft) && (
          <Toggle
            label='Hide zero-count values'
            checked={!!refiner.hideZeroCountValues}
            onChange={(_e, checked): void => onPatch({ hideZeroCountValues: !!checked })}
            onText='Yes' offText='No'
          />
        )}
        {isFieldRelevant('sortBy', ft) && (
          <Dropdown
            label='Sort values by'
            selectedKey={refiner.sortBy}
            options={SORT_BY_OPTIONS}
            onChange={(_e, opt): void => { if (opt) { onPatch({ sortBy: String(opt.key) }); } }}
          />
        )}
        {isFieldRelevant('sortDirection', ft) && (
          <Dropdown
            label='Sort direction'
            selectedKey={refiner.sortDirection}
            options={SORT_DIR_OPTIONS}
            onChange={(_e, opt): void => { if (opt) { onPatch({ sortDirection: String(opt.key) }); } }}
          />
        )}
      </Section>
    );
  }

  function renderBehavior(): React.ReactElement {
    return (
      <Section title='Behavior' key='behavior'>
        {isFieldRelevant('operator', ft) && (
          <Dropdown
            label='Combine selections with'
            selectedKey={refiner.operator}
            options={OPERATOR_OPTIONS}
            onChange={(_e, opt): void => { if (opt) { onPatch({ operator: String(opt.key) }); } }}
          />
        )}
        {isFieldRelevant('multiValues', ft) && (
          <Toggle
            label='Allow multiple selections'
            checked={refiner.multiValues}
            onChange={(_e, checked): void => onPatch({ multiValues: !!checked })}
            onText='Yes' offText='No'
          />
        )}
      </Section>
    );
  }

  function renderToggleLabels(): React.ReactElement {
    return (
      <Section title='Toggle labels' key='toggleLabels'>
        <TextField
          label='True label'
          value={refiner.trueLabel || ''}
          onChange={(_e, v): void => onPatch({ trueLabel: v || '' })}
          placeholder='Yes'
          description='Caption shown when the toggle is ON.'
        />
        <TextField
          label='False label'
          value={refiner.falseLabel || ''}
          onChange={(_e, v): void => onPatch({ falseLabel: v || '' })}
          placeholder='No'
          description='Caption shown when the toggle is OFF.'
        />
        <Toggle
          label='Invert boolean'
          checked={!!refiner.invertBoolean}
          onChange={(_e, checked): void => onPatch({ invertBoolean: !!checked })}
          onText='Inverted (true means false)'
          offText='Normal'
        />
      </Section>
    );
  }

  function renderConditional(): React.ReactElement {
    const hasParent = !!(refiner.dependsOn && refiner.dependsOn.trim());
    return (
      <Section title='Conditional visibility' key='conditional' defaultCollapsed={!hasParent}>
        <TextField
          label='Depends on'
          value={refiner.dependsOn || ''}
          onChange={(_e, v): void => onPatch({ dependsOn: v || '' })}
          placeholder='ContentType'
          description='Managed property of the parent refiner. Leave blank for no dependency.'
        />
        {hasParent && (
          <>
            <Toggle
              label='Show only after parent has a selection'
              checked={!!refiner.showWhenParentHasValue}
              onChange={(_e, checked): void => onPatch({ showWhenParentHasValue: !!checked })}
              onText='Yes' offText='No'
            />
            <Toggle
              label='Reset this refiner when parent changes'
              checked={!!refiner.resetWhenParentChanges}
              onChange={(_e, checked): void => onPatch({ resetWhenParentChanges: !!checked })}
              onText='Yes' offText='No'
            />
          </>
        )}
        {!hasParent && (
          <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
            Set a parent managed property to enable conditional visibility.
          </MessageBar>
        )}
      </Section>
    );
  }

  function renderAudience(): React.ReactElement {
    return (
      <Section title='Audience targeting' key='audience' defaultCollapsed={!refiner.audience}>
        <TextField
          label='Visible to (Azure AD group IDs)'
          value={refiner.audience || ''}
          onChange={(_e, v): void => onPatch({ audience: v || '' })}
          placeholder='Comma-separated group IDs; leave blank for everyone'
          multiline
          rows={2}
        />
      </Section>
    );
  }

  return (
    <div className={styles.editorForm}>
      <div className={styles.editorHeader}>
        <span className={styles.editorHeaderTitle}>
          {refiner.displayName || '(unnamed refiner)'}
        </span>
        <span className={styles.editorHeaderType}>
          {FILTER_TYPE_TEXT[refiner.filterType] || refiner.filterType}
        </span>
      </div>
      {sections.map((s): React.ReactElement | null => {
        switch (s) {
          case 'basic': return renderBasic();
          case 'display': return renderDisplay();
          case 'behavior': return renderBehavior();
          case 'toggleLabels': return renderToggleLabels();
          case 'conditional': return renderConditional();
          case 'audience': return renderAudience();
          default: return null;
        }
      })}
    </div>
  );
};

// ─── Top-level control ──────────────────────────────────────

const FiltersCollectionControl: React.FC<IFiltersCollectionControlProps> = (props) => {
  const { label, panelHeader, manageButtonLabel, value, onChange } = props;

  const [isPanelOpen, setIsPanelOpen] = React.useState<boolean>(false);
  // Draft state isolates uncommitted edits from the parent property bag
  // so admins can Cancel out of a session without polluting the saved JSON.
  const [draft, setDraft] = React.useState<IFiltersCollectionItem[]>(value);
  const [selectedId, setSelectedId] = React.useState<string | undefined>(undefined);

  function openPanel(): void {
    setDraft(value);
    setSelectedId(value.length > 0 ? value[0].uniqueId : undefined);
    setIsPanelOpen(true);
  }

  function closePanel(): void {
    setIsPanelOpen(false);
  }

  function commitAndClose(): void {
    onChange(draft);
    setIsPanelOpen(false);
  }

  function addRefiner(): void {
    const next = draft.slice();
    const fresh = makeBlankRefiner();
    next.push(fresh);
    setDraft(next);
    setSelectedId(fresh.uniqueId);
  }

  function deleteRefiner(uniqueId: string): void {
    const next = draft.filter((r) => r.uniqueId !== uniqueId);
    setDraft(next);
    if (selectedId === uniqueId) {
      setSelectedId(next.length > 0 ? next[0].uniqueId : undefined);
    }
  }

  function moveRefiner(uniqueId: string, direction: -1 | 1): void {
    const idx = draft.findIndex((r) => r.uniqueId === uniqueId);
    if (idx < 0) { return; }
    const target = idx + direction;
    if (target < 0 || target >= draft.length) { return; }
    const next = draft.slice();
    const [removed] = next.splice(idx, 1);
    next.splice(target, 0, removed);
    setDraft(next);
  }

  function updateSelected(patch: Partial<IFiltersCollectionItem>): void {
    if (!selectedId) { return; }
    setDraft(draft.map((r) => (r.uniqueId === selectedId ? { ...r, ...patch } : r)));
  }

  const selected: IFiltersCollectionItem | undefined =
    selectedId !== undefined ? draft.find((r) => r.uniqueId === selectedId) : undefined;

  const sectionsForSelected: FilterEditorSection[] = selected
    ? getRelevantSections(selected.filterType)
    : [];

  // ─── Render: trigger (text field + button) ─────────────────
  const summary: string = value.length === 0
    ? 'No refiners configured'
    : value.length + ' refiner' + (value.length === 1 ? '' : 's') + ' configured';

  return (
    <div className={styles.host}>
      <Label>{label}</Label>
      <div className={styles.triggerRow}>
        <span className={styles.summary}>{summary}</span>
        <DefaultButton
          iconProps={{ iconName: 'Settings' }}
          text={manageButtonLabel}
          onClick={openPanel}
        />
      </div>

      <Panel
        isOpen={isPanelOpen}
        onDismiss={closePanel}
        headerText={panelHeader}
        type={PanelType.custom}
        customWidth='960px'
        isLightDismiss={false}
        closeButtonAriaLabel='Close'
        onRenderFooterContent={(): React.ReactElement => (
          <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign='end'>
            <DefaultButton text='Cancel' onClick={closePanel} />
            <PrimaryButton text='Save' onClick={commitAndClose} />
          </Stack>
        )}
        isFooterAtBottom
      >
        <div className={styles.panelBody}>
          {/* ─── Left rail: refiner list ────────────────────── */}
          <div className={styles.leftRail} role='listbox' aria-label='Refiners'>
            {draft.length === 0 && (
              <div className={styles.emptyHint}>
                No refiners yet. Click <strong>Add refiner</strong> to start.
              </div>
            )}
            {draft.map((r, idx) => {
              const isSelected = r.uniqueId === selectedId;
              const incomplete = !r.managedProperty || !r.displayName;
              const typeLabel = FILTER_TYPE_TEXT[r.filterType] || r.filterType;
              return (
                <div
                  key={r.uniqueId}
                  className={[styles.refinerRow, isSelected ? styles.refinerRowSelected : ''].join(' ').trim()}
                  role='option'
                  aria-selected={isSelected}
                  tabIndex={0}
                  onClick={(): void => setSelectedId(r.uniqueId)}
                  onKeyDown={(e): void => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault();
                      setSelectedId(r.uniqueId);
                    }
                  }}
                >
                  <div className={styles.refinerRowMain}>
                    <div className={styles.refinerName}>
                      {r.displayName || <span className={styles.refinerNameMuted}>(no display name)</span>}
                      {incomplete && (
                        <Icon
                          iconName='Warning'
                          className={styles.incompleteIcon}
                          title='Missing managed property or display name'
                          aria-label='Incomplete refiner'
                        />
                      )}
                    </div>
                    <div className={styles.refinerMeta}>
                      <span className={styles.typeChip}>{typeLabel}</span>
                      {r.managedProperty && (
                        <span className={styles.refinerProp}>{r.managedProperty}</span>
                      )}
                    </div>
                  </div>
                  <div className={styles.refinerActions}>
                    <IconButton
                      iconProps={{ iconName: 'Up' }}
                      title='Move up'
                      ariaLabel='Move up'
                      disabled={idx === 0}
                      onClick={(e): void => { e.stopPropagation(); moveRefiner(r.uniqueId, -1); }}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Down' }}
                      title='Move down'
                      ariaLabel='Move down'
                      disabled={idx === draft.length - 1}
                      onClick={(e): void => { e.stopPropagation(); moveRefiner(r.uniqueId, 1); }}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title='Delete refiner'
                      ariaLabel='Delete refiner'
                      onClick={(e): void => { e.stopPropagation(); deleteRefiner(r.uniqueId); }}
                    />
                  </div>
                </div>
              );
            })}
            <ActionButton
              iconProps={{ iconName: 'Add' }}
              text='Add refiner'
              onClick={addRefiner}
              className={styles.addButton}
            />
          </div>

          {/* ─── Right pane: editor ─────────────────────────── */}
          <div className={styles.rightPane}>
            {!selected && (
              <div className={styles.emptyState}>
                <Icon iconName='Filter' className={styles.emptyStateIcon} />
                <div>Select a refiner on the left, or add a new one.</div>
              </div>
            )}
            {selected && (
              <FilterEditorForm
                key={selected.uniqueId}
                refiner={selected}
                sections={sectionsForSelected}
                onPatch={updateSelected}
              />
            )}
          </div>
        </div>
      </Panel>
    </div>
  );
};

export default FiltersCollectionControl;
