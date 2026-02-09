import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { IContextualMenuProps, IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { FileTypeIcon, IconType, ImageSize } from '@pnp/spfx-controls-react/lib/FileTypeIcon';
import { ISearchResult } from '@interfaces/index';
import { formatFileSize } from './documentTitleUtils';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import styles from './SpSearchResults.module.scss';

export interface IGalleryLayoutProps {
  items: ISearchResult[];
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}

/**
 * Determines if a file type is an image.
 */
function isImageType(fileType: string): boolean {
  const ft: string = (fileType || '').toLowerCase();
  return ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].indexOf(ft) >= 0;
}

/**
 * Single gallery item component.
 */
const GalleryItem: React.FC<{
  item: ISearchResult;
  position: number;
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}> = (galleryProps) => {
  const { item, position, onPreviewItem, onItemClick } = galleryProps;
  const isImage: boolean = isImageType(item.fileType);

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

  const handleImageClick = React.useCallback((): void => {
    if (onItemClick) {
      onItemClick(item, position);
    }
    if (onPreviewItem) {
      onPreviewItem(item);
    }
  }, [item, position, onItemClick, onPreviewItem]);

  const handleCopyLink = React.useCallback((): void => {
    if (navigator.clipboard) {
      navigator.clipboard.writeText(item.url).catch((): void => {
        // Silently fail
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
    <div className={styles.galleryItem} role="listitem">
      {/* Thumbnail area */}
      <div
        className={styles.galleryThumbnailContainer}
        onClick={handleImageClick}
        role="button"
        tabIndex={0}
        onKeyDown={(e: React.KeyboardEvent): void => {
          if (e.key === 'Enter' || e.key === ' ') {
            handleImageClick();
          }
        }}
      >
        {item.thumbnailUrl ? (
          <img
            className={styles.galleryThumbnail}
            src={item.thumbnailUrl}
            alt={item.title}
            loading="lazy"
          />
        ) : (
          <div className={styles.galleryThumbnailFallback}>
            <FileTypeIcon type={IconType.image} path={item.url} size={ImageSize.large} />
          </div>
        )}
        {/* Hover overlay */}
        <div className={styles.galleryOverlay}>
          <Icon iconName={isImage ? 'ZoomIn' : 'View'} style={{ fontSize: 24 }} />
        </div>
      </div>

      {/* Info bar */}
      <div className={styles.galleryInfo}>
        <div className={styles.galleryTitle} title={item.title}>
          <DocumentTitleHoverCard item={item} position={position} onItemClick={onItemClick}>
            {(handleClick): React.ReactNode => (
              <a
                href={item.url}
                target="_blank"
                rel="noopener noreferrer"
                onClick={(e: React.MouseEvent): void => {
                  e.stopPropagation();
                  handleClick(e);
                }}
              >
                {item.title}
              </a>
            )}
          </DocumentTitleHoverCard>
        </div>
        <div className={styles.galleryMeta}>
          {item.fileType && (
            <span className={styles.galleryFileType}>{item.fileType.toUpperCase()}</span>
          )}
          {item.fileSize > 0 && (
            <span className={styles.galleryFileSize}>{formatFileSize(item.fileSize)}</span>
          )}
        </div>
        <IconButton
          className={styles.galleryMoreButton}
          iconProps={{ iconName: 'More' }}
          title="More actions"
          ariaLabel="More actions"
          menuProps={menuProps}
        />
      </div>
    </div>
  );
};

/**
 * GalleryLayout — renders search results as a photo-gallery grid.
 * Ideal for image/video/document results with thumbnails.
 *
 * Grid columns:
 *  - Desktop (>= 1024px): 4 columns
 *  - Tablet (>= 640px): 3 columns
 *  - Mobile (< 640px): 2 columns
 */
const GalleryLayout: React.FC<IGalleryLayoutProps> = (props) => {
  const { items, onPreviewItem, onItemClick } = props;

  return (
    <div className={styles.galleryGrid} role="list">
      {items.map((item: ISearchResult, index: number) => (
        <GalleryItem
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

export default GalleryLayout;
