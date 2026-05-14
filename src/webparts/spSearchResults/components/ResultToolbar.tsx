import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { ISortField, ISortableProperty } from '@interfaces/index';
import styles from './SpSearchResults.module.scss';

export interface IResultToolbarProps {
  totalCount: number;
  activeLayoutKey: string;
  availableLayouts: string[];
  sort: ISortField | undefined;
  sortableProperties: ISortableProperty[];
  showResultCount: boolean;
  showSortDropdown: boolean;
  onLayoutChange: (key: string) => void;
  onSortChange: (sort: ISortField) => void;
  /** Called on button hover to warm the webpack chunk before the user clicks. */
  onPreloadLayout?: (key: string) => void;
  // T2.D11 — layout-agnostic export. When `onExportCsv` / `onExportXlsx`
  // are provided, an Export IconButton appears with a contextual menu.
  // The "Selection only (N rows)" item only renders when
  // `selectionCount > 0` (consumes T2.D2 bulk-selection state).
  onExportCsv?: (selectionOnly: boolean) => void;
  onExportXlsx?: (selectionOnly: boolean) => void;
  selectionCount?: number;
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

function getLayoutTooltip(layoutKey: string): string {
  switch (layoutKey) {
    case 'list':
      return 'List view: titles, summaries, and metadata in a readable stack.';
    case 'compact':
      return 'Compact view: dense rows for quick scanning with minimal detail.';
    case 'card':
      return 'Card view: visual tiles with thumbnails and richer content.';
    case 'people':
      return 'People view: profile cards with contact and org details.';
    case 'grid':
      return 'DataGrid view: resizable columns, chooser, export, and fullscreen.';
    case 'gallery':
      return 'Gallery view: large previews for image and media-heavy content.';
    default:
      return 'Switch result layout.';
  }
}

function renderLayoutButton(
  key: string,
  iconName: string,
  label: string,
  activeLayoutKey: string,
  onLayoutChange: (key: string) => void,
  onPreloadLayout?: (key: string) => void
): React.ReactElement {
  return (
    <TooltipHost key={key} content={getLayoutTooltip(key)}>
      <IconButton
        className={activeLayoutKey === key ? styles.layoutButtonActive : styles.layoutButton}
        iconProps={{ iconName }}
        title={label}
        ariaLabel={label}
        checked={activeLayoutKey === key}
        onClick={(): void => { onLayoutChange(key); }}
        onMouseEnter={onPreloadLayout ? (): void => { onPreloadLayout(key); } : undefined}
      />
    </TooltipHost>
  );
}

const ResultToolbar: React.FC<IResultToolbarProps> = (props) => {
  const {
    totalCount,
    activeLayoutKey,
    availableLayouts,
    sort,
    sortableProperties,
    showResultCount,
    showSortDropdown,
    onLayoutChange,
    onSortChange,
    onPreloadLayout,
    onExportCsv,
    onExportXlsx,
    selectionCount,
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

  return (
    <div className={styles.toolbar}>
      <div className={styles.toolbarLeft}>
        {showResultCount && (
          <span className={styles.resultCount} aria-live="polite" role="status">
            <span className={styles.resultCountIcon}><Icon iconName="Search" /></span>
            <span className={styles.resultCountText}>{formatResultCount(totalCount)}</span>
          </span>
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
        {availableLayouts.indexOf('list') >= 0 && (
          renderLayoutButton('list', 'List', 'List view', activeLayoutKey, onLayoutChange)
        )}
        {availableLayouts.indexOf('compact') >= 0 && (
          renderLayoutButton('compact', 'GridViewSmall', 'Compact view', activeLayoutKey, onLayoutChange)
        )}
        {availableLayouts.indexOf('card') >= 0 && (
          renderLayoutButton('card', 'GridViewMedium', 'Card view', activeLayoutKey, onLayoutChange, onPreloadLayout)
        )}
        {availableLayouts.indexOf('people') >= 0 && (
          renderLayoutButton('people', 'People', 'People view', activeLayoutKey, onLayoutChange, onPreloadLayout)
        )}
        {availableLayouts.indexOf('grid') >= 0 && (
          renderLayoutButton('grid', 'Table', 'DataGrid view', activeLayoutKey, onLayoutChange, onPreloadLayout)
        )}
        {availableLayouts.indexOf('gallery') >= 0 && (
          renderLayoutButton('gallery', 'PhotoCollection', 'Gallery view', activeLayoutKey, onLayoutChange, onPreloadLayout)
        )}
        {/* T2.D11 — layout-agnostic export menu. Renders whenever the
            web part wires `onExportCsv` / `onExportXlsx`. The DataGrid
            layout has its own Toolbar export so we keep this menu visible
            on every layout; admins can choose either path. */}
        {(onExportCsv || onExportXlsx) && (
          <IconButton
            iconProps={{ iconName: 'Download' }}
            title="Export results"
            ariaLabel="Export results"
            menuProps={{
              items: [
                ...(onExportCsv ? [{
                  key: 'csv-all',
                  text: 'Export all to CSV',
                  iconProps: { iconName: 'TextDocument' },
                  onClick: (): void => { onExportCsv(false); },
                }] : []),
                ...(onExportXlsx ? [{
                  key: 'xlsx-all',
                  text: 'Export all to XLSX',
                  iconProps: { iconName: 'ExcelDocument' },
                  onClick: (): void => { onExportXlsx(false); },
                }] : []),
                ...((selectionCount && selectionCount > 0) ? [
                  { key: 'sel-divider', itemType: 1 as 0 | 1 | 2 | 3 },
                  ...(onExportCsv ? [{
                    key: 'csv-sel',
                    text: 'Selection only (' + selectionCount + ' rows) to CSV',
                    iconProps: { iconName: 'TextDocument' },
                    onClick: (): void => { onExportCsv(true); },
                  }] : []),
                  ...(onExportXlsx ? [{
                    key: 'xlsx-sel',
                    text: 'Selection only (' + selectionCount + ' rows) to XLSX',
                    iconProps: { iconName: 'ExcelDocument' },
                    onClick: (): void => { onExportXlsx(true); },
                  }] : []),
                ] : []),
              ],
            }}
          />
        )}
      </div>
    </div>
  );
};

export default ResultToolbar;
