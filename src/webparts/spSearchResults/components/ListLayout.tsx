import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { ISearchResult } from '@interfaces/index';
import styles from './SpSearchResults.module.scss';

export interface IListLayoutProps {
  items: ISearchResult[];
  enableSelection: boolean;
  selectedKeys: string[];
  onToggleSelection: (key: string, multiSelect: boolean) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}

/**
 * Maps a file extension to a Fluent UI icon name.
 */
function getFileTypeIcon(fileType: string): string {
  const ft: string = (fileType || '').toLowerCase();
  switch (ft) {
    case 'docx': case 'doc': return 'WordDocument';
    case 'xlsx': case 'xls': return 'ExcelDocument';
    case 'pptx': case 'ppt': return 'PowerPointDocument';
    case 'pdf': return 'PDF';
    case 'one': case 'onetoc2': return 'OneNoteLogo';
    case 'vsdx': case 'vsd': return 'VisioDocument';
    case 'html': case 'htm': case 'aspx': return 'FileHTML';
    case 'txt': return 'TextDocument';
    case 'zip': case 'rar': case '7z': return 'ZipFolder';
    case 'jpg': case 'jpeg': case 'png': case 'gif': case 'bmp': case 'svg': return 'FileImage';
    case 'mp4': case 'avi': case 'mov': case 'wmv': return 'Video';
    case 'mp3': case 'wav': return 'MusicInCollectionFill';
    case 'csv': return 'ExcelDocument';
    case 'msg': case 'eml': return 'Mail';
    default: return 'Page';
  }
}

/**
 * Extracts a breadcrumb-style URL display from a full URL.
 * e.g. "https://contoso.sharepoint.com/sites/hr/docs/guide.pdf" => "contoso.sharepoint.com > sites > hr > docs"
 */
function formatUrlBreadcrumb(url: string): string {
  try {
    // Remove protocol
    let cleaned: string = url.replace(/^https?:\/\//, '');
    // Remove query/hash
    const qIdx: number = cleaned.indexOf('?');
    if (qIdx >= 0) {
      cleaned = cleaned.substring(0, qIdx);
    }
    const hIdx: number = cleaned.indexOf('#');
    if (hIdx >= 0) {
      cleaned = cleaned.substring(0, hIdx);
    }
    // Split into segments and remove the file name (last segment)
    const segments: string[] = cleaned.split('/');
    if (segments.length > 1) {
      segments.pop(); // remove file name
    }
    return segments.join(' \u203A ');
  } catch {
    return url;
  }
}

/**
 * Sanitizes HitHighlightedSummary HTML from SharePoint Search API.
 * Strips all tags except safe formatting tags used for hit highlighting.
 */
function sanitizeSummaryHtml(html: string): string {
  // Allow only safe tags used by SharePoint search highlighting
  return html.replace(/<\/?(?!(?:b|strong|em|i|mark|c0|ddd)\b)[^>]*>/gi, '');
}

/**
 * Formats an ISO date string into a short locale-friendly format.
 */
function formatDate(isoDate: string): string {
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

const ListLayout: React.FC<IListLayoutProps> = (props) => {
  const { items, enableSelection, selectedKeys, onToggleSelection, onItemClick } = props;

  const handleCheckboxChange = React.useCallback(
    (key: string, ev?: React.FormEvent<HTMLElement | HTMLInputElement>): void => {
      // Check for ctrlKey via the native event if available
      const nativeEvent = ev?.nativeEvent as KeyboardEvent | undefined;
      const multiSelect: boolean = !!(nativeEvent && nativeEvent.ctrlKey);
      onToggleSelection(key, multiSelect);
    },
    [onToggleSelection]
  );

  const handleLinkClick = React.useCallback(
    (item: ISearchResult, position: number): void => {
      if (onItemClick) {
        onItemClick(item, position);
      }
    },
    [onItemClick]
  );

  return (
    <ul className={styles.resultList} role="list">
      {items.map((item: ISearchResult, index: number) => {
        const isSelected: boolean = selectedKeys.indexOf(item.key) >= 0;
        const cardClasses: string = styles.resultCard +
          (isSelected ? ' ' + styles.resultCardSelected : '');
        const checkboxClasses: string = styles.resultCheckbox +
          (isSelected ? ' ' + styles.resultCheckboxVisible : '');

        return (
          <li key={item.key} className={cardClasses} role="listitem">
            {enableSelection && (
              <div className={checkboxClasses}>
                <Checkbox
                  checked={isSelected}
                  onChange={(_ev?: React.FormEvent<HTMLElement | HTMLInputElement>): void => {
                    handleCheckboxChange(item.key, _ev);
                  }}
                  ariaLabel={'Select ' + item.title}
                />
              </div>
            )}
            <div className={styles.resultIcon}>
              <Icon iconName={getFileTypeIcon(item.fileType)} />
            </div>
            <div className={styles.resultBody}>
              <h3 className={styles.resultTitle}>
                <a
                  href={item.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  onClick={(): void => { handleLinkClick(item, index + 1); }}
                >
                  {item.title}
                </a>
              </h3>
              <p className={styles.resultUrl}>{formatUrlBreadcrumb(item.url)}</p>
              {item.summary && (
                <div
                  className={styles.resultSummary}
                  dangerouslySetInnerHTML={{ __html: sanitizeSummaryHtml(item.summary) }}
                />
              )}
              <div className={styles.resultMeta}>
                {item.author && item.author.displayText && (
                  <span className={styles.metaItem}>
                    <Icon iconName="Contact" style={{ fontSize: 12 }} />
                    {item.author.displayText}
                  </span>
                )}
                {item.author && item.author.displayText && item.modified && (
                  <span className={styles.metaSeparator} />
                )}
                {item.modified && (
                  <span className={styles.metaItem}>
                    <Icon iconName="Calendar" style={{ fontSize: 12 }} />
                    {formatDate(item.modified)}
                  </span>
                )}
                {item.siteName && (
                  <>
                    <span className={styles.metaSeparator} />
                    <span className={styles.metaItem}>
                      <Icon iconName="SharePointLogo" style={{ fontSize: 12 }} />
                      {item.siteName}
                    </span>
                  </>
                )}
              </div>
            </div>
          </li>
        );
      })}
    </ul>
  );
};

export default ListLayout;
