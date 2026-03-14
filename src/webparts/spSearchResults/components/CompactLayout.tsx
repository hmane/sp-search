import * as React from 'react';
import { FileTypeIcon, IconType, ImageSize } from '@pnp/spfx-controls-react/lib/FileTypeIcon';
import { ISearchResult } from '@interfaces/index';
import { formatFileSize, formatShortDate, stripHtml, getResultAnchorProps, formatTitleText, TitleDisplayMode } from './documentTitleUtils';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import { ISelectedPropertyColumn } from './ISpSearchResultsProps';
import AddToCollectionButton from './AddToCollectionButton';
import styles from './SpSearchResults.module.scss';

export interface ICompactLayoutProps {
  items: ISearchResult[];
  searchContextId: string;
  compactPropertyColumns: ISelectedPropertyColumn[];
  titleDisplayMode: TitleDisplayMode;
  onItemClick?: (item: ISearchResult, position: number) => void;
}

type CompactColumnKind = 'author' | 'date' | 'fileSize' | 'fileType' | 'site' | 'text';

interface ICompactColumnConfig {
  property: string;
  label: string;
  kind: CompactColumnKind;
  width: string;
}

const DEFAULT_COMPACT_COLUMNS: ISelectedPropertyColumn[] = [
  { property: 'Author', alias: 'Author' },
  { property: 'LastModifiedTime', alias: 'Modified' },
  { property: 'Size', alias: 'Size' },
  { property: 'FileType', alias: 'Type' }
];

function getCompactColumns(columns: ISelectedPropertyColumn[]): ICompactColumnConfig[] {
  const source = columns.length > 0 ? columns : DEFAULT_COMPACT_COLUMNS;
  const seen = new Set<string>();
  const result: ICompactColumnConfig[] = [];

  for (let i: number = 0; i < source.length; i++) {
    const property = (source[i].property || '').trim();
    if (!property) {
      continue;
    }
    const lookup = property.toLowerCase();
    if (lookup === 'title' || lookup === 'filename' || seen.has(lookup)) {
      continue;
    }
    seen.add(lookup);

    if (lookup === 'author' || lookup === 'authorowsuser' || lookup === 'displayauthor') {
      result.push({ property, label: source[i].alias || 'Author', kind: 'author', width: '140px' });
      continue;
    }
    if (lookup === 'lastmodifiedtime' || lookup === 'modified' || lookup === 'created') {
      result.push({ property, label: source[i].alias || 'Modified', kind: 'date', width: '110px' });
      continue;
    }
    if (lookup === 'size' || lookup === 'filesize') {
      result.push({ property, label: source[i].alias || 'Size', kind: 'fileSize', width: '72px' });
      continue;
    }
    if (lookup === 'filetype' || lookup === 'fileextension') {
      result.push({ property, label: source[i].alias || 'Type', kind: 'fileType', width: '72px' });
      continue;
    }
    if (lookup === 'sitename' || lookup === 'sitetitle') {
      result.push({ property, label: source[i].alias || 'Site', kind: 'site', width: '130px' });
      continue;
    }
    result.push({ property, label: source[i].alias || property, kind: 'text', width: '130px' });
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

function renderCompactCell(item: ISearchResult, column: ICompactColumnConfig): React.ReactNode {
  const value = resolveCompactValue(item, column.property);

  switch (column.kind) {
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
    default:
      return <span className={styles.compactMetaText}>{String(value || '')}</span>;
  }
}

const CompactLayout: React.FC<ICompactLayoutProps> = (props) => {
  const { items, searchContextId, compactPropertyColumns, titleDisplayMode, onItemClick } = props;
  const columns = React.useMemo(
    (): ICompactColumnConfig[] => getCompactColumns(compactPropertyColumns),
    [compactPropertyColumns]
  );
  const layoutTemplate = React.useMemo((): string => {
    return ['20px', 'minmax(0, 1fr)', ...columns.map((column) => column.width)].join(' ');
  }, [columns]);

  return (
    <div className={styles.compactTable} role="table" aria-label="Search results">
      <div className={styles.compactHeader} role="row" style={{ gridTemplateColumns: layoutTemplate }}>
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
        const linkProps = getResultAnchorProps(item);

        return (
          <div
            key={item.key}
            className={styles.compactRow}
            role="row"
            title={tooltipText}
            style={{ gridTemplateColumns: layoutTemplate }}
          >
            <div className={styles.compactIcon} role="cell">
              <FileTypeIcon type={IconType.image} path={item.url} size={ImageSize.small} />
            </div>
            <div className={styles.compactTitle} role="cell">
              <div className={styles.compactTitleInner}>
                <DocumentTitleHoverCard item={item} position={index + 1} onItemClick={onItemClick} hostDisplay="block">
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
