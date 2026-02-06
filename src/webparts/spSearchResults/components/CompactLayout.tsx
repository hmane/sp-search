import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { ISearchResult } from '@interfaces/index';
import styles from './SpSearchResults.module.scss';

export interface ICompactLayoutProps {
  items: ISearchResult[];
  onItemClick?: (item: ISearchResult, position: number) => void;
}

/**
 * Maps a file extension to a Fluent UI icon name (compact version).
 */
function getFileTypeIcon(fileType: string): string {
  const ft: string = (fileType || '').toLowerCase();
  switch (ft) {
    case 'docx': case 'doc': return 'WordDocument';
    case 'xlsx': case 'xls': return 'ExcelDocument';
    case 'pptx': case 'ppt': return 'PowerPointDocument';
    case 'pdf': return 'PDF';
    case 'one': case 'onetoc2': return 'OneNoteLogo';
    case 'html': case 'htm': case 'aspx': return 'FileHTML';
    case 'txt': return 'TextDocument';
    case 'jpg': case 'jpeg': case 'png': case 'gif': case 'bmp': case 'svg': return 'FileImage';
    case 'mp4': case 'avi': case 'mov': return 'Video';
    default: return 'Page';
  }
}

/**
 * Formats an ISO date string into a short date format.
 */
function formatShortDate(isoDate: string): string {
  if (!isoDate) {
    return '';
  }
  try {
    const d: Date = new Date(isoDate);
    if (isNaN(d.getTime())) {
      return '';
    }
    const month: number = d.getMonth() + 1;
    const day: number = d.getDate();
    const year: number = d.getFullYear();
    return month + '/' + day + '/' + year;
  } catch {
    return '';
  }
}

/**
 * Strips HTML tags from a summary string for tooltip display.
 */
function stripHtml(html: string): string {
  if (!html) {
    return '';
  }
  return html.replace(/<[^>]*>/g, '');
}

const CompactLayout: React.FC<ICompactLayoutProps> = (props) => {
  const { items, onItemClick } = props;

  const handleLinkClick = React.useCallback(
    (item: ISearchResult, position: number): void => {
      if (onItemClick) {
        onItemClick(item, position);
      }
    },
    [onItemClick]
  );

  return (
    <ul className={styles.compactList} role="list">
      {items.map((item: ISearchResult, index: number) => {
        const tooltipContent: string = stripHtml(item.summary);

        return (
          <TooltipHost
            key={item.key}
            content={tooltipContent || item.title}
            calloutProps={{ gapSpace: 4 }}
          >
            <li className={styles.compactRow} role="listitem">
              <div className={styles.compactIcon}>
                <Icon iconName={getFileTypeIcon(item.fileType)} />
              </div>
              <div className={styles.compactTitle}>
                <a
                  href={item.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  onClick={(): void => { handleLinkClick(item, index + 1); }}
                >
                  {item.title}
                </a>
              </div>
              <div className={styles.compactAuthor}>
                {item.author ? item.author.displayText : ''}
              </div>
              <div className={styles.compactDate}>
                {formatShortDate(item.modified)}
              </div>
              <div className={styles.compactFileType}>
                {item.fileType || ''}
              </div>
            </li>
          </TooltipHost>
        );
      })}
    </ul>
  );
};

export default CompactLayout;
