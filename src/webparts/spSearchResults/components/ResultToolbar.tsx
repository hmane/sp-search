import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { ISortField } from '@interfaces/index';
import styles from './SpSearchResults.module.scss';

export interface IResultToolbarProps {
  totalCount: number;
  activeLayoutKey: string;
  sort: ISortField | undefined;
  showResultCount: boolean;
  showSortDropdown: boolean;
  onLayoutChange: (key: string) => void;
  onSortChange: (sort: ISortField) => void;
}

/** Sort preset key for "Relevance" (no explicit sort) */
const SORT_RELEVANCE: string = 'relevance';
const SORT_DATE_NEWEST: string = 'date-newest';
const SORT_DATE_OLDEST: string = 'date-oldest';
const SORT_AUTHOR_AZ: string = 'author-az';

const sortOptions: IDropdownOption[] = [
  { key: SORT_RELEVANCE, text: 'Relevance' },
  { key: SORT_DATE_NEWEST, text: 'Date (newest)' },
  { key: SORT_DATE_OLDEST, text: 'Date (oldest)' },
  { key: SORT_AUTHOR_AZ, text: 'Author A\u2013Z' }
];

/**
 * Determines the currently selected sort dropdown key from the store's ISortField.
 */
function getSortKey(sort: ISortField | undefined): string {
  if (!sort) {
    return SORT_RELEVANCE;
  }
  if (sort.property === 'Write' || sort.property === 'LastModifiedTime') {
    return sort.direction === 'Descending' ? SORT_DATE_NEWEST : SORT_DATE_OLDEST;
  }
  if (sort.property === 'Author' || sort.property === 'DisplayAuthor') {
    return SORT_AUTHOR_AZ;
  }
  return SORT_RELEVANCE;
}

/**
 * Maps a dropdown key to a sort field for the store.
 */
function mapSortKey(key: string): ISortField {
  switch (key) {
    case SORT_DATE_NEWEST:
      return { property: 'LastModifiedTime', direction: 'Descending' };
    case SORT_DATE_OLDEST:
      return { property: 'LastModifiedTime', direction: 'Ascending' };
    case SORT_AUTHOR_AZ:
      return { property: 'DisplayAuthor', direction: 'Ascending' };
    default:
      // Relevance — represented as an ascending rank sort
      return { property: 'Rank', direction: 'Ascending' };
  }
}

/**
 * Formats a total count into a user-friendly string.
 * e.g. 1250 => "About 1,250 results"
 */
function formatResultCount(count: number): string {
  if (count === 0) {
    return 'No results';
  }
  if (count === 1) {
    return '1 result';
  }
  // Manual thousands formatting for ES5 target compatibility
  const parts: string[] = [];
  let remaining: number = count;
  while (remaining > 0) {
    const chunk: number = remaining % 1000;
    remaining = Math.floor(remaining / 1000);
    if (remaining > 0) {
      // Zero-pad to 3 digits
      let chunkStr: string = String(chunk);
      while (chunkStr.length < 3) {
        chunkStr = '0' + chunkStr;
      }
      parts.unshift(chunkStr);
    } else {
      parts.unshift(String(chunk));
    }
  }
  return 'About ' + parts.join(',') + ' results';
}

const ResultToolbar: React.FC<IResultToolbarProps> = (props) => {
  const {
    totalCount,
    activeLayoutKey,
    sort,
    showResultCount,
    showSortDropdown,
    onLayoutChange,
    onSortChange
  } = props;

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
