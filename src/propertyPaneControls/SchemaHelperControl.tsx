import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { IconButton } from '@fluentui/react/lib/Button';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Link } from '@fluentui/react/lib/Link';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';

import { fetchManagedProperties } from '@services/index';
import type { ISchemaResult } from '@services/index';
import type { IManagedProperty } from '@interfaces/index';
import styles from './SchemaHelperControl.module.scss';

// ─── Types ──────────────────────────────────────────────────

export type SchemaFilterHint = 'refinable' | 'sortable' | 'retrievable' | 'queryable';

export interface ISchemaHelperControlProps {
  label: string;
  description?: string;
  value: string;
  multiline?: boolean;
  rows?: number;
  /** Pre-selects a Pivot tab matching this flag */
  filterHint?: SchemaFilterHint;
  onChange: (newValue: string) => void;
}

// ─── Constants ──────────────────────────────────────────────

const FLAG_LABELS: Array<{ key: keyof IManagedProperty; short: string; full: string }> = [
  { key: 'queryable', short: 'Q', full: 'Queryable' },
  { key: 'retrievable', short: 'R', full: 'Retrievable' },
  { key: 'refinable', short: 'Re', full: 'Refinable' },
  { key: 'sortable', short: 'S', full: 'Sortable' },
];

const PIVOT_MAP: Record<string, keyof IManagedProperty | undefined> = {
  all: undefined,
  refinable: 'refinable',
  sortable: 'sortable',
  retrievable: 'retrievable',
  queryable: 'queryable',
};

// ─── Component ──────────────────────────────────────────────

const SchemaHelperControl: React.FC<ISchemaHelperControlProps> = function SchemaHelperControl(
  props: ISchemaHelperControlProps
): React.ReactElement {
  const { label, description, value, multiline, rows, filterHint, onChange } = props;

  // State
  const [isPanelOpen, setIsPanelOpen] = React.useState<boolean>(false);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [schemaResult, setSchemaResult] = React.useState<ISchemaResult | undefined>(undefined);
  const [searchFilter, setSearchFilter] = React.useState<string>('');
  const [activeTab, setActiveTab] = React.useState<string>(filterHint || 'all');

  // ─── Handlers ───────────────────────────────────────────

  function handleTextFieldChange(
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void {
    onChange(newValue !== undefined ? newValue : '');
  }

  function handleBrowseClick(): void {
    setIsLoading(true);
    setIsPanelOpen(true);
    setSearchFilter('');
    setActiveTab(filterHint || 'all');

    fetchManagedProperties()
      .then(function (result: ISchemaResult): void {
        setSchemaResult(result);
        setIsLoading(false);
      })
      .catch(function (): void {
        setSchemaResult({ status: 'error', properties: [], errorMessage: 'Failed to fetch schema' });
        setIsLoading(false);
      });
  }

  function handlePanelDismiss(): void {
    setIsPanelOpen(false);
  }

  function handleSearchChange(
    _event?: React.ChangeEvent<HTMLInputElement>,
    newValue?: string
  ): void {
    setSearchFilter(newValue || '');
  }

  function handleSearchClear(): void {
    setSearchFilter('');
  }

  function handlePivotChange(item?: PivotItem): void {
    if (item && item.props.itemKey) {
      setActiveTab(item.props.itemKey);
    }
  }

  function handlePropertyClick(propertyName: string): void {
    if (multiline) {
      // For multiline (comma-separated): append to existing value
      const trimmed = value.trim();
      if (trimmed.length === 0) {
        onChange(propertyName);
      } else {
        // Check if already present
        const existing = trimmed.split(',').map(function (s: string): string { return s.trim(); });
        if (existing.indexOf(propertyName) < 0) {
          onChange(trimmed + ', ' + propertyName);
        }
      }
    } else {
      // For single-line: replace value
      onChange(propertyName);
      setIsPanelOpen(false);
    }
  }

  function handleRefreshClick(event: React.MouseEvent): void {
    event.preventDefault();
    setIsLoading(true);
    fetchManagedProperties(true)
      .then(function (result: ISchemaResult): void {
        setSchemaResult(result);
        setIsLoading(false);
      })
      .catch(function (): void {
        setSchemaResult({ status: 'error', properties: [], errorMessage: 'Failed to refresh schema' });
        setIsLoading(false);
      });
  }

  // ─── Computed: filtered properties ──────────────────────

  const filteredProperties: IManagedProperty[] = React.useMemo(function (): IManagedProperty[] {
    if (!schemaResult || schemaResult.properties.length === 0) {
      return [];
    }

    let result = schemaResult.properties;

    // Apply Pivot tab filter
    const flagKey = PIVOT_MAP[activeTab];
    if (flagKey) {
      result = result.filter(function (p: IManagedProperty): boolean {
        return (p as unknown as Record<string, unknown>)[flagKey] === true;
      });
    }

    // Apply search text filter
    if (searchFilter.length > 0) {
      const needle = searchFilter.toLowerCase();
      result = result.filter(function (p: IManagedProperty): boolean {
        const nameMatch = p.name.toLowerCase().indexOf(needle) >= 0;
        const aliasMatch = p.alias ? p.alias.toLowerCase().indexOf(needle) >= 0 : false;
        return nameMatch || aliasMatch;
      });
    }

    return result;
  }, [schemaResult, activeTab, searchFilter]);

  // ─── Check for unauthorized state ───────────────────────

  const isUnauthorized = schemaResult && schemaResult.status === 'unauthorized';

  // ─── Render ─────────────────────────────────────────────

  return (
    <div className={styles.schemaHelper}>
      <div className={styles.schemaHelperRow}>
        <div className={styles.schemaHelperTextField}>
          <TextField
            label={label}
            description={description}
            value={value}
            onChange={handleTextFieldChange}
            multiline={multiline}
            rows={rows}
          />
        </div>
        <TooltipHost
          content={
            isUnauthorized
              ? 'Requires Search Admin or Site Collection Admin permissions'
              : 'Browse search schema managed properties'
          }
        >
          <IconButton
            className={styles.browseButton}
            iconProps={{ iconName: 'DeveloperTools' }}
            title="Browse Schema"
            ariaLabel="Browse Schema"
            onClick={handleBrowseClick}
          />
        </TooltipHost>
      </div>

      {isUnauthorized && (
        <MessageBar
          messageBarType={MessageBarType.info}
          className={styles.unauthorizedMessage}
        >
          Contact your SharePoint admin for managed property names.
        </MessageBar>
      )}

      <Panel
        isOpen={isPanelOpen}
        onDismiss={handlePanelDismiss}
        type={PanelType.medium}
        headerText="Search Schema - Managed Properties"
        isLightDismiss={true}
      >
        <div className={styles.schemaPanel}>
          {isLoading && (
            <div className={styles.loadingContainer}>
              <Spinner size={SpinnerSize.large} label="Loading schema..." />
            </div>
          )}

          {!isLoading && schemaResult && schemaResult.status === 'unauthorized' && (
            <MessageBar messageBarType={MessageBarType.warning}>
              Requires Search Admin or Site Collection Admin permissions. Contact your SharePoint admin for managed property names.
            </MessageBar>
          )}

          {!isLoading && schemaResult && schemaResult.status === 'error' && (
            <MessageBar messageBarType={MessageBarType.error}>
              {schemaResult.errorMessage || 'Failed to load schema.'}
            </MessageBar>
          )}

          {!isLoading && schemaResult && schemaResult.status === 'success' && (
            <>
              <div className={styles.propertySearch}>
                <SearchBox
                  placeholder="Filter properties by name or alias..."
                  value={searchFilter}
                  onChange={handleSearchChange}
                  onClear={handleSearchClear}
                />
              </div>

              <div className={styles.pivotBar}>
                <Pivot
                  selectedKey={activeTab}
                  onLinkClick={handlePivotChange}
                >
                  <PivotItem headerText="All" itemKey="all" />
                  <PivotItem headerText="Queryable" itemKey="queryable" />
                  <PivotItem headerText="Retrievable" itemKey="retrievable" />
                  <PivotItem headerText="Refinable" itemKey="refinable" />
                  <PivotItem headerText="Sortable" itemKey="sortable" />
                </Pivot>
              </div>

              <div className={styles.resultCount}>
                {String(filteredProperties.length) + ' of ' + String(schemaResult.properties.length) + ' properties'}
              </div>

              {/* Header row */}
              <div className={styles.headerRow}>
                <span className={styles.headerName}>Name</span>
                <span className={styles.headerAlias}>Alias</span>
                <span className={styles.headerType}>Type</span>
                <div className={styles.flagHeader}>
                  {FLAG_LABELS.map(function (flag): React.ReactElement {
                    return (
                      <TooltipHost key={flag.key} content={flag.full}>
                        <span className={styles.flagHeaderLabel}>{flag.short}</span>
                      </TooltipHost>
                    );
                  })}
                </div>
              </div>

              {/* Property list */}
              <div className={styles.propertyList}>
                {filteredProperties.length === 0 && (
                  <div className={styles.emptyMessage}>
                    No properties match the current filter.
                  </div>
                )}
                {filteredProperties.map(function (prop: IManagedProperty): React.ReactElement {
                  return (
                    <div
                      key={prop.name}
                      className={styles.propertyRow}
                      onClick={function (): void { handlePropertyClick(prop.name); }}
                      role="button"
                      tabIndex={0}
                      onKeyDown={function (e: React.KeyboardEvent): void {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault();
                          handlePropertyClick(prop.name);
                        }
                      }}
                      title={'Click to insert ' + prop.name}
                    >
                      <span className={styles.propertyName}>{prop.name}</span>
                      <span className={styles.propertyAlias}>
                        {prop.alias && prop.alias !== prop.name ? prop.alias : ''}
                      </span>
                      <span className={styles.typeTag}>{prop.type}</span>
                      <div className={styles.flagsContainer}>
                        {FLAG_LABELS.map(function (flag): React.ReactElement {
                          const isSet = (prop as unknown as Record<string, unknown>)[flag.key] === true;
                          if (isSet) {
                            return (
                              <TooltipHost key={flag.key} content={flag.full}>
                                <span className={styles.flagIcon}>
                                  <Icon iconName="CheckMark" />
                                </span>
                              </TooltipHost>
                            );
                          }
                          return <span key={flag.key} className={styles.flagIconEmpty} />;
                        })}
                      </div>
                    </div>
                  );
                })}
              </div>

              <div className={styles.refreshLink}>
                <Link onClick={handleRefreshClick}>
                  Refresh Schema
                </Link>
              </div>
            </>
          )}
        </div>
      </Panel>
    </div>
  );
};

export default SchemaHelperControl;
