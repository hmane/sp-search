import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { getFileTypeIconProps } from '@fluentui/react-file-type-icons';
import { ISearchResult } from '@interfaces/index';
import { formatShortDate, formatFileSize, stripHtml, formatTitleText, isImageType, TitleDisplayMode } from './documentTitleUtils';
import { resolveResultLink, type IResultLinkConfig } from './resultLink';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import type { IColumnConfigItem, ColumnRenderer } from './ColumnConfigField/columnConfig';
import {
  renderText,
  renderRichText,
  renderNumber,
  renderFileSize,
  renderBoolean,
  renderTags,
  renderPersona,
  renderDate,
  renderUrl,
  renderFileType,
} from './renderCell';
import AddToCollectionButton from './AddToCollectionButton';
import styles from './SpSearchResults.module.scss';

export interface ICompactLayoutProps {
  items: ISearchResult[];
  searchContextId: string;
  /** Stream B / Phase 3 — full IColumnConfigItem[]. */
  compactPropertyColumns: IColumnConfigItem[];
  titleDisplayMode: TitleDisplayMode;
  onItemClick?: (item: ISearchResult, position: number) => void;
  // Stream C / #7
  linkConfig: IResultLinkConfig;
  onOpenInSidePanel?: (item: ISearchResult) => void;
  // T2.D2 — bulk-selection wiring shared with List + DataGrid.
  bulkSelection: string[];
  onToggleSelect: (itemKey: string) => void;
}

/**
 * Compact-layout cell kinds — a subset of the DataGrid's `ColumnKind`
 * adapted for the tight 4-column grid layout. The auto-detect path keeps
 * today's behaviour (plain text for author/site, short-date for dates,
 * etc.); explicit renderers route through `renderCell.tsx`.
 */
type CompactKind =
  | 'author'    // auto-detected — plain text + tooltip (Compact doesn't have room for an avatar)
  | 'date'      // auto-detected — short date
  | 'fileSize'  // auto-detected — formatFileSize
  | 'fileType'  // auto-detected — uppercase badge
  | 'site'      // auto-detected — plain text
  | 'text'      // auto-detected — plain text
  // explicit-renderer kinds dispatch via renderCell.tsx
  | 'persona' | 'richText' | 'number' | 'tags' | 'boolean' | 'url' | 'rendererText' | 'rendererDate' | 'rendererFileSize' | 'rendererFileType';

interface ICompactColumn {
  property: string;
  label: string;
  kind: CompactKind;
  width: string;
  config: IColumnConfigItem;
}

function widthForKind(kind: CompactKind): string {
  switch (kind) {
    case 'author':
    case 'persona':
      return '160px';
    case 'date':
    case 'rendererDate':
      return '110px';
    case 'fileSize':
    case 'rendererFileSize':
    case 'number':
      return '80px';
    case 'fileType':
    case 'rendererFileType':
    case 'boolean':
      return '72px';
    case 'site':
      return '130px';
    case 'tags':
      return '160px';
    default:
      return '130px';
  }
}

function autoDetectCompactKind(property: string): CompactKind {
  const normalized = property.toLowerCase();
  if (normalized === 'author' || normalized === 'authorowsuser' || normalized === 'displayauthor') {
    return 'author';
  }
  if (normalized === 'lastmodifiedtime' || normalized === 'modified' || normalized === 'created') {
    return 'date';
  }
  if (normalized === 'size' || normalized === 'filesize') {
    return 'fileSize';
  }
  if (normalized === 'filetype' || normalized === 'fileextension') {
    return 'fileType';
  }
  if (normalized === 'sitename' || normalized === 'sitetitle') {
    return 'site';
  }
  return 'text';
}

function kindFromExplicitRenderer(renderer: ColumnRenderer): CompactKind | undefined {
  switch (renderer) {
    case 'persona':  return 'persona';
    case 'richText': return 'richText';
    case 'number':   return 'number';
    case 'tags':     return 'tags';
    case 'boolean':  return 'boolean';
    case 'url':      return 'url';
    case 'text':     return 'rendererText';
    case 'date':     return 'rendererDate';
    case 'fileSize': return 'rendererFileSize';
    case 'fileType': return 'rendererFileType';
    default:         return undefined;
  }
}

function getCompactColumns(columns: IColumnConfigItem[]): ICompactColumn[] {
  const result: ICompactColumn[] = [];
  const seen = new Set<string>();

  for (let i: number = 0; i < columns.length; i++) {
    const config = columns[i];
    const property = (config.property || '').trim();
    if (!property) {
      continue;
    }
    const lookup = property.toLowerCase();
    if (lookup === 'title' || lookup === 'filename' || seen.has(lookup)) {
      continue;
    }
    seen.add(lookup);

    const explicit = kindFromExplicitRenderer(config.renderer);
    const kind: CompactKind = explicit !== undefined ? explicit : autoDetectCompactKind(property);
    result.push({
      property,
      label: config.alias || property,
      kind,
      width: widthForKind(kind),
      config,
    });
  }

  return result.slice(0, 4);
}

function resolveCompactValue(item: ISearchResult, property: string): unknown {
  const normalized = property.toLowerCase();
  switch (normalized) {
    case 'author':
    case 'authorowsuser':
    case 'displayauthor':
      return item.author?.displayText || item.properties[property] || '';
    case 'lastmodifiedtime':
    case 'modified':
      return item.modified || item.properties[property] || '';
    case 'created':
      return item.created || item.properties[property] || '';
    case 'size':
    case 'filesize':
      return item.fileSize || item.properties[property] || 0;
    case 'filetype':
    case 'fileextension':
      return item.fileType || item.properties[property] || '';
    case 'sitename':
    case 'sitetitle':
      return item.siteName || item.properties[property] || '';
    default:
      return item.properties[property] || '';
  }
}

function renderCompactCell(item: ISearchResult, column: ICompactColumn): React.ReactNode {
  const value = resolveCompactValue(item, column.property);

  switch (column.kind) {
    // Auto-detected kinds — preserve today's compact-friendly rendering.
    case 'author':
    case 'site':
    case 'text':
      return <span className={styles.compactMetaText}>{String(value || '')}</span>;
    case 'date':
      return <span className={styles.compactMetaText}>{formatShortDate(String(value || ''))}</span>;
    case 'fileSize':
      return <span className={`${styles.compactMetaText} ${styles.compactMetaAlignRight}`}>{formatFileSize(Number(value || 0))}</span>;
    case 'fileType':
      return value ? (
        <span className={styles.compactFileTypeBadge}>
          {String(value).toUpperCase()}
        </span>
      ) : '';
    // Explicit-renderer kinds dispatch via renderCell.tsx (Stream B / Phase 2 + 3).
    case 'persona':           return renderPersona(value, column.config);
    case 'richText':          return renderRichText(value, column.config);
    case 'number':            return renderNumber(value, column.config);
    case 'tags':              return renderTags(value, column.config);
    case 'boolean':           return renderBoolean(value, column.config);
    case 'url':               return renderUrl(value, column.config);
    case 'rendererText':      return renderText(value, column.config);
    case 'rendererDate':      return renderDate(value, column.config);
    case 'rendererFileSize':  return renderFileSize(value, column.config);
    case 'rendererFileType':  return renderFileType(value, column.config);
    default:
      return <span className={styles.compactMetaText}>{String(value || '')}</span>;
  }
}

const CompactLayout: React.FC<ICompactLayoutProps> = (props) => {
  const { items, searchContextId, compactPropertyColumns, titleDisplayMode, onItemClick, linkConfig, onOpenInSidePanel, bulkSelection, onToggleSelect } = props;
  const columns = React.useMemo(
    (): ICompactColumn[] => getCompactColumns(compactPropertyColumns),
    [compactPropertyColumns]
  );
  // T2.D2 — prepend a 24px selection-checkbox column to the grid template.
  const layoutTemplate = React.useMemo((): string => {
    return ['24px', '20px', 'minmax(0, 1fr)', ...columns.map((column) => column.width)].join(' ');
  }, [columns]);
  const selectionSet = React.useMemo(() => new Set(bulkSelection), [bulkSelection]);

  return (
    <div className={styles.compactTable} role="table" aria-label="Search results">
      <div className={styles.compactHeader} role="row" style={{ gridTemplateColumns: layoutTemplate }}>
        <div className={styles.compactHeaderSelect} role="columnheader" aria-label="Select" />
        <div className={styles.compactHeaderIcon} role="columnheader" aria-label="File type" />
        <div className={styles.compactHeaderTitle} role="columnheader">Name</div>
        {columns.map((column) => (
          <div
            key={column.property}
            className={styles.compactHeaderMeta}
            role="columnheader"
            style={{ width: column.width }}
            title={column.label}
          >
            {column.label}
          </div>
        ))}
      </div>
      {items.map((item: ISearchResult, index: number) => {
        const tooltipText: string = stripHtml(item.summary) || item.title;
        const linkProps = resolveResultLink(item, linkConfig);
        const isSelected = selectionSet.has(item.key);

        return (
          <div
            key={item.key}
            className={`${styles.compactRow}${isSelected ? ' ' + styles.compactRowSelected : ''}`}
            role="row"
            title={tooltipText}
            style={{ gridTemplateColumns: layoutTemplate }}
          >
            <div className={styles.compactSelect} role="cell">
              {/* T2.D2 — selection checkbox; stopPropagation keeps the row's
                  title-anchor click from firing when toggling. */}
              <input
                type="checkbox"
                className={styles.compactSelectCheckbox}
                checked={isSelected}
                onChange={(): void => onToggleSelect(item.key)}
                onClick={(e): void => e.stopPropagation()}
                aria-label={isSelected ? 'Deselect ' + item.title : 'Select ' + item.title}
              />
            </div>
            <div className={styles.compactIcon} role="cell">
              {isImageType(item) && item.thumbnailUrl ? (
                // Stream C / #8 — show the image itself for image results.
                <img className={styles.compactIconImage} src={item.thumbnailUrl} alt="" loading="lazy" />
              ) : (
                <Icon {...getFileTypeIconProps({ extension: item.fileType || '', size: 16 })} />
              )}
            </div>
            <div className={styles.compactTitle} role="cell">
              <div className={styles.compactTitleInner}>
                <DocumentTitleHoverCard
                  item={item}
                  position={index + 1}
                  onItemClick={onItemClick}
                  hostDisplay="block"
                  clickTarget={linkConfig.clickTarget}
                  onOpenInSidePanel={onOpenInSidePanel}
                >
                  {(handleClick): React.ReactNode => (
                    <a
                      href={linkProps.href}
                      target={linkProps.target}
                      rel={linkProps.rel}
                      className={titleDisplayMode === 'wrap' ? styles.compactTitleLinkWrap : undefined}
                      onClick={handleClick}
                    >
                      {formatTitleText(item.title, titleDisplayMode)}
                    </a>
                  )}
                </DocumentTitleHoverCard>
                <AddToCollectionButton
                  item={item}
                  searchContextId={searchContextId}
                />
              </div>
            </div>
            {columns.map((column) => (
              <div
                key={column.property}
                className={styles.compactMetaCell}
                role="cell"
                style={{ width: column.width }}
              >
                {renderCompactCell(item, column)}
              </div>
            ))}
          </div>
        );
      })}
    </div>
  );
};

export default CompactLayout;
