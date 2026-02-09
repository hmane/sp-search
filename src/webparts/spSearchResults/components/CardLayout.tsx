import * as React from 'react';
import {
  DocumentCard,
  DocumentCardPreview,
  DocumentCardActivity,
  DocumentCardType,
} from '@fluentui/react/lib/DocumentCard';
import { ImageFit } from '@fluentui/react/lib/Image';
import { FileTypeIcon, IconType, ImageSize } from '@pnp/spfx-controls-react/lib/FileTypeIcon';
import { ISearchResult } from '@interfaces/index';
import { formatRelativeDate } from './documentTitleUtils';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import styles from './SpSearchResults.module.scss';

export interface ICardLayoutProps {
  items: ISearchResult[];
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}

/**
 * Builds a user photo URL for the activity section.
 * Uses SharePoint's userphoto.aspx endpoint which works cross-site.
 */
function getUserPhotoUrl(email: string): string {
  if (!email) {
    return '';
  }
  return '/_layouts/15/userphoto.aspx?size=S&accountname=' + encodeURIComponent(email);
}

/**
 * Renders the file type icon as an image element for the preview fallback.
 */
const FileTypeIconPreview: React.FC<{ url: string }> = (iconProps) => {
  return (
    <div className={styles.docCardIconPreview}>
      <FileTypeIcon type={IconType.image} path={iconProps.url} size={ImageSize.large} />
    </div>
  );
};

/**
 * Single document card item.
 */
const CardItem: React.FC<{
  item: ISearchResult;
  position: number;
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}> = (cardItemProps) => {
  const { item, position, onItemClick } = cardItemProps;

  // Build activity string (modified date + file type)
  const activityParts: string[] = [];
  if (item.modified) {
    activityParts.push('Modified ' + formatRelativeDate(item.modified));
  }
  if (item.fileType) {
    activityParts.push(item.fileType.toUpperCase());
  }
  const activityText: string = activityParts.join(' \u00B7 ');

  // Build people array for DocumentCardActivity
  const people: { name: string; profileImageSrc: string }[] = [];
  if (item.author && item.author.displayText) {
    people.push({
      name: item.author.displayText,
      profileImageSrc: getUserPhotoUrl(item.author.email),
    });
  }

  return (
    <div className={styles.docCardItem}>
      <DocumentCard
        type={DocumentCardType.normal}
        aria-label={item.title}
      >
        {/* Preview image or file type icon fallback */}
        {item.thumbnailUrl ? (
          <DocumentCardPreview
            previewImages={[
              {
                name: item.title,
                previewImageSrc: item.thumbnailUrl,
                imageFit: ImageFit.cover,
                height: 196,
              },
            ]}
          />
        ) : (
          <FileTypeIconPreview url={item.url} />
        )}

        {/* Document title with HoverCard */}
        <div className={styles.docCardTitleWrapper} title={item.title}>
          <DocumentTitleHoverCard item={item} position={position} onItemClick={onItemClick} hostDisplay="block">
            {(handleClick): React.ReactNode => (
              <a
                href={item.url}
                target="_blank"
                rel="noopener noreferrer"
                className={styles.docCardTitleLink}
                onClick={handleClick}
              >
                {item.title}
              </a>
            )}
          </DocumentTitleHoverCard>
        </div>

        {/* Activity: author persona + modified date */}
        {people.length > 0 && (
          <DocumentCardActivity
            activity={activityText}
            people={people}
          />
        )}
      </DocumentCard>
    </div>
  );
};

/**
 * CardLayout — renders search results as Fluent UI DocumentCards in a responsive grid.
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
