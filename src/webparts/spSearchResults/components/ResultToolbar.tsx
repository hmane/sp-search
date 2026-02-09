import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { ISortField, ISortableProperty } from '@interfaces/index';
import styles from './SpSearchResults.module.scss';

export interface IResultToolbarProps {
  totalCount: number;
  activeLayoutKey: string;
  sort: ISortField | undefined;
  sortableProperties: ISortableProperty[];
  showResultCount: boolean;
  showSortDropdown: boolean;
  onLayoutChange: (key: string) => void;
  onSortChange: (sort: ISortField) => void;
}

/** Sort preset key for "Relevance" (no explicit sort) */
const SORT_RELEVANCE: string = 'relevance';

/**
 * Builds dropdown options from admin-configured sortable properties.
 * Always includes "Relevance" as the first option.
 * Falls back to default presets if no sortable properties are configured.
 */
function buildSortOptions(sortableProperties: ISortableProperty[]): IDropdownOption[] {
  const options: IDropdownOption[] = [
    { key: SORT_RELEVANCE, text: 'Relevance' }
  ];

  if (sortableProperties.length > 0) {
    sortableProperties.forEach(function (sp: ISortableProperty): void {
      options.push({
        key: sp.property + ':' + sp.direction,
        text: sp.label
      });
    });
  } else {
    // Fallback: default presets when no admin-configured sort fields
    options.push(
      { key: 'LastModifiedTime:Descending', text: 'Date (newest)' },
      { key: 'LastModifiedTime:Ascending', text: 'Date (oldest)' },
      { key: 'DisplayAuthor:Ascending', text: 'Author A\u2013Z' }
    );
  }

  return options;
}

/**
 * Determines the currently selected sort dropdown key from the store's ISortField.
 */
function getSortKey(sort: ISortField | undefined): string {
  if (!sort || sort.property === 'Rank') {
    return SORT_RELEVANCE;
  }
  return sort.property + ':' + sort.direction;
}

/**
 * Maps a dropdown key back to a sort field for the store.
 */
function mapSortKey(key: string): ISortField {
  if (key === SORT_RELEVANCE) {
    return { property: 'Rank', direction: 'Ascending' };
  }
  const parts: string[] = key.split(':');
  return {
    property: parts[0],
    direction: (parts[1] as 'Ascending' | 'Descending') || 'Ascending'
  };
}

/**
 * Formats a total count into a user-friendly string.
 * e.g. 1250 => "About 1,250 results"
 */
function formatResultCount(count: number): string {
  if (count === 0) {
    return 'No results found';
  }
  if (count === 1) {
    return '1 result';
  }
  // Format with locale-aware thousands separators
  const formatted: string = count.toLocaleString();
  // Use approximate wording for large counts (SharePoint TotalRows is an estimate)
  if (count >= 100) {
    return '\u2248 ' + formatted + ' results';
  }
  return formatted + ' results';
}

const ResultToolbar: React.FC<IResultToolbarProps> = (props) => {
  const {
    totalCount,
    activeLayoutKey,
    sort,
    sortableProperties,
    showResultCount,
    showSortDropdown,
    onLayoutChange,
    onSortChange
  } = props;

  const sortOptions: IDropdownOption[] = React.useMemo(
    function (): IDropdownOption[] {
      return buildSortOptions(sortableProperties);
    },
    [sortableProperties]
  );

  const handleSortChange = React.useCallback(
    (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
      if (option) {
        onSortChange(mapSortKey(String(option.key)));
      }
    },
    [onSortChange]
  );

  const handleListLayout = React.useCallback((): void => {
    onLayoutChange('list');
  }, [onLayoutChange]);

  const handleCompactLayout = React.useCallback((): void => {
    onLayoutChange('compact');
  }, [onLayoutChange]);

  const handleCardLayout = React.useCallback((): void => {
    onLayoutChange('card');
  }, [onLayoutChange]);

  const handlePeopleLayout = React.useCallback((): void => {
    onLayoutChange('people');
  }, [onLayoutChange]);

  const handleDataGridLayout = React.useCallback((): void => {
    onLayoutChange('datagrid');
  }, [onLayoutChange]);

  const handleGalleryLayout = React.useCallback((): void => {
    onLayoutChange('gallery');
  }, [onLayoutChange]);

  return (
    <div className={styles.toolbar}>
      <div className={styles.toolbarLeft}>
        {showResultCount && (
          <span className={styles.resultCount} aria-live="polite" role="status">{formatResultCount(totalCount)}</span>
        )}
        {showSortDropdown && (
          <Dropdown
            className={styles.sortDropdown}
            options={sortOptions}
            selectedKey={getSortKey(sort)}
            onChange={handleSortChange}
            ariaLabel="Sort results"
          />
        )}
      </div>
      <div className={styles.toolbarRight}>
        <IconButton
          className={activeLayoutKey === 'list' ? styles.layoutButtonActive : styles.layoutButton}
          iconProps={{ iconName: 'List' }}
          title="List view"
          ariaLabel="List view"
          checked={activeLayoutKey === 'list'}
          onClick={handleListLayout}
        />
        <IconButton
          className={activeLayoutKey === 'compact' ? styles.layoutButtonActive : styles.layoutButton}
          iconProps={{ iconName: 'GridViewSmall' }}
          title="Compact view"
          ariaLabel="Compact view"
          checked={activeLayoutKey === 'compact'}
          onClick={handleCompactLayout}
        />
        <IconButton
          className={activeLayoutKey === 'card' ? styles.layoutButtonActive : styles.layoutButton}
          iconProps={{ iconName: 'GridViewMedium' }}
          title="Card view"
          ariaLabel="Card view"
          checked={activeLayoutKey === 'card'}
          onClick={handleCardLayout}
        />
        <IconButton
          className={activeLayoutKey === 'people' ? styles.layoutButtonActive : styles.layoutButton}
          iconProps={{ iconName: 'People' }}
          title="People view"
          ariaLabel="People view"
          checked={activeLayoutKey === 'people'}
          onClick={handlePeopleLayout}
        />
        <IconButton
          className={activeLayoutKey === 'datagrid' ? styles.layoutButtonActive : styles.layoutButton}
          iconProps={{ iconName: 'Table' }}
          title="DataGrid view"
          ariaLabel="DataGrid view"
          checked={activeLayoutKey === 'datagrid'}
          onClick={handleDataGridLayout}
        />
        <IconButton
          className={activeLayoutKey === 'gallery' ? styles.layoutButtonActive : styles.layoutButton}
          iconProps={{ iconName: 'PhotoCollection' }}
          title="Gallery view"
          ariaLabel="Gallery view"
          checked={activeLayoutKey === 'gallery'}
          onClick={handleGalleryLayout}
        />
      </div>
    </div>
  );
};

export default ResultToolbar;
