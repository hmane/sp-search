import * as React from 'react';
import { Card, Header, Content } from 'spfx-toolkit/lib/components/Card/components';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { IContextualMenuProps, IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { ISearchResult } from '@interfaces/index';
import styles from './SpSearchResults.module.scss';

export interface ICardLayoutProps {
  items: ISearchResult[];
  onPreviewItem?: (item: ISearchResult) => void;
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

/**
 * Strips HTML tags from a summary string for safe text display.
 */
function stripHtml(html: string): string {
  if (!html) {
    return '';
  }
  return html.replace(/<[^>]*>/g, '');
}

/**
 * Formats a file size in bytes into a human-readable string.
 */
function formatFileSize(bytes: number): string {
  if (!bytes || bytes <= 0) {
    return '';
  }
  if (bytes < 1024) {
    return bytes + ' B';
  }
  if (bytes < 1048576) {
    return Math.round(bytes / 1024) + ' KB';
  }
  return (bytes / 1048576).toFixed(1) + ' MB';
}

/**
 * Single card item rendered inside the grid.
 */
const CardItem: React.FC<{
  item: ISearchResult;
  position: number;
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}> = (cardItemProps) => {
  const { item, position, onPreviewItem, onItemClick } = cardItemProps;
  const summaryText: string = stripHtml(item.summary);
  const sizeDisplay: string = formatFileSize(item.fileSize);

  const handlePreviewClick = React.useCallback((): void => {
    if (onPreviewItem) {
      onPreviewItem(item);
    }
  }, [item, onPreviewItem]);

  const handleOpenInNewTab = React.useCallback((): void => {
    if (onItemClick) {
      onItemClick(item, position);
    }
    window.open(item.url, '_blank', 'noopener,noreferrer');
  }, [item, position, onItemClick]);

  const handleTitleClick = React.useCallback((): void => {
    if (onItemClick) {
      onItemClick(item, position);
    }
  }, [item, position, onItemClick]);

  const handleCopyLink = React.useCallback((): void => {
    if (navigator.clipboard) {
      navigator.clipboard.writeText(item.url).catch((): void => {
        // Silently fail — no clipboard API in older browsers
      });
    }
  }, [item.url]);

  const menuItems: IContextualMenuItem[] = React.useMemo((): IContextualMenuItem[] => {
    const items: IContextualMenuItem[] = [
      {
        key: 'open',
        text: 'Open in new tab',
        iconProps: { iconName: 'OpenInNewTab' },
        onClick: handleOpenInNewTab
      },
      {
        key: 'copyLink',
        text: 'Copy link',
        iconProps: { iconName: 'Link' },
        onClick: handleCopyLink
      }
    ];
    if (onPreviewItem) {
      items.unshift({
        key: 'preview',
        text: 'Preview',
        iconProps: { iconName: 'View' },
        onClick: handlePreviewClick
      });
    }
    return items;
  }, [handleOpenInNewTab, handleCopyLink, handlePreviewClick, onPreviewItem]);

  const menuProps: IContextualMenuProps = React.useMemo(
    (): IContextualMenuProps => ({ items: menuItems }),
    [menuItems]
  );

  return (
    <div className={styles.cardItem}>
      <Card
        id={item.key}
        elevation={1}
        allowExpand={false}
        allowMaximize={false}
        defaultExpanded={true}
        size="regular"
      >
        <Header hideExpandButton={true} hideMaximizeButton={true}>
          <div className={styles.cardHeader}>
            <div className={styles.cardHeaderLeft}>
              <span className={styles.cardFileIcon}>
                <Icon iconName={getFileTypeIcon(item.fileType)} />
              </span>
              <a
                className={styles.cardTitleLink}
                href={item.url}
                target="_blank"
                rel="noopener noreferrer"
                title={item.title}
                onClick={handleTitleClick}
              >
                {item.title}
              </a>
            </div>
            <IconButton
              className={styles.cardMoreButton}
              iconProps={{ iconName: 'More' }}
              title="More actions"
              ariaLabel="More actions"
              menuProps={menuProps}
            />
          </div>
        </Header>
        <Content padding="compact">
          <div className={styles.cardContent}>
            {/* Thumbnail or fallback icon */}
            <div className={styles.cardThumbnailContainer}>
              {item.thumbnailUrl ? (
                <img
                  className={styles.cardThumbnail}
                  src={item.thumbnailUrl}
                  alt={item.title}
                  loading="lazy"
                />
              ) : (
                <div className={styles.cardThumbnailFallback}>
                  <Icon iconName={getFileTypeIcon(item.fileType)} />
                </div>
              )}
            </div>

            {/* Summary text */}
            {summaryText && (
              <p className={styles.cardSummary}>{summaryText}</p>
            )}

            {/* Metadata row */}
            <div className={styles.cardMeta}>
              {item.author && item.author.displayText && (
                <span className={styles.cardMetaItem}>
                  <Icon iconName="Contact" style={{ fontSize: 12 }} />
                  {item.author.displayText}
                </span>
              )}
              {item.modified && (
                <span className={styles.cardMetaItem}>
                  <Icon iconName="Calendar" style={{ fontSize: 12 }} />
                  {formatDate(item.modified)}
                </span>
              )}
              {item.siteName && (
                <span className={styles.cardMetaItem}>
                  <Icon iconName="SharePointLogo" style={{ fontSize: 12 }} />
                  {item.siteName}
                </span>
              )}
              {sizeDisplay && (
                <span className={styles.cardMetaItem}>
                  <Icon iconName="Page" style={{ fontSize: 12 }} />
                  {sizeDisplay}
                </span>
              )}
            </div>
          </div>
        </Content>
      </Card>
    </div>
  );
};

/**
 * CardLayout — renders search results as cards in a responsive CSS grid.
 * Uses the spfx-toolkit Card component for consistent card presentation.
 *
 * Grid columns:
 *  - Desktop (>= 1024px): 3 columns
 *  - Tablet (>= 640px): 2 columns
 *  - Mobile (< 640px): 1 column
 */
const CardLayout: React.FC<ICardLayoutProps> = (props) => {
  const { items, onPreviewItem, onItemClick } = props;

  return (
    <div className={styles.cardGrid} role="list">
      {items.map((item: ISearchResult, index: number) => (
        <CardItem
          key={item.key}
          item={item}
          position={index + 1}
          onPreviewItem={onPreviewItem}
          onItemClick={onItemClick}
        />
      ))}
    </div>
  );
};

export default CardLayout;
