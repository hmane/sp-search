import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { FileTypeIcon, IconType, ImageSize } from '@pnp/spfx-controls-react/lib/FileTypeIcon';
import { ISearchResult } from '@interfaces/index';
import { formatFileSize, formatShortDate, stripHtml } from './documentTitleUtils';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import styles from './SpSearchResults.module.scss';

export interface ICompactLayoutProps {
  items: ISearchResult[];
  onItemClick?: (item: ISearchResult, position: number) => void;
}

const CompactLayout: React.FC<ICompactLayoutProps> = (props) => {
  const { items, onItemClick } = props;

  return (
    <div className={styles.compactTable} role="table" aria-label="Search results">
      <div className={styles.compactHeader} role="row">
        <div className={styles.compactHeaderIcon} role="columnheader" aria-label="File type" />
        <div className={styles.compactHeaderTitle} role="columnheader">Name</div>
        <div className={styles.compactHeaderAuthor} role="columnheader">Author</div>
        <div className={styles.compactHeaderDate} role="columnheader">
          <Icon iconName="Clock" style={{ fontSize: 11, marginRight: 4 }} />
          Modified
        </div>
        <div className={styles.compactHeaderSize} role="columnheader">Size</div>
        <div className={styles.compactHeaderFileType} role="columnheader">Type</div>
      </div>
      {items.map((item: ISearchResult, index: number) => {
        const sizeDisplay: string = formatFileSize(item.fileSize);
        const tooltipText: string = stripHtml(item.summary) || item.title;

        return (
          <div key={item.key} className={styles.compactRow} role="row" title={tooltipText}>
            <div className={styles.compactIcon} role="cell">
              <FileTypeIcon type={IconType.image} path={item.url} size={ImageSize.small} />
            </div>
            <div className={styles.compactTitle} role="cell">
              <DocumentTitleHoverCard item={item} position={index + 1} onItemClick={onItemClick}>
                {(handleClick): React.ReactNode => (
                  <a
                    href={item.url}
                    target="_blank"
                    rel="noopener noreferrer"
                    onClick={handleClick}
                  >
                    {item.title}
                  </a>
                )}
              </DocumentTitleHoverCard>
            </div>
            <div className={styles.compactAuthor} role="cell">
              {item.author ? item.author.displayText : ''}
            </div>
            <div className={styles.compactDate} role="cell">
              {formatShortDate(item.modified)}
            </div>
            <div className={styles.compactSize} role="cell">
              {sizeDisplay}
            </div>
            <div className={styles.compactFileType} role="cell">
              {item.fileType ? (
                <span className={styles.compactFileTypeBadge}>
                  {item.fileType.toUpperCase()}
                </span>
              ) : ''}
            </div>
          </div>
        );
      })}
    </div>
  );
};

export default CompactLayout;
