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
import { formatRelativeDate, formatDateTime, getResultAnchorProps, formatTitleText, TitleDisplayMode } from './documentTitleUtils';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import AddToCollectionButton from './AddToCollectionButton';
import styles from './SpSearchResults.module.scss';

export interface ICardLayoutProps {
  items: ISearchResult[];
  searchContextId: string;
  titleDisplayMode: TitleDisplayMode;
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
  searchContextId: string;
  titleDisplayMode: TitleDisplayMode;
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}> = (cardItemProps) => {
  const { item, position, searchContextId, titleDisplayMode, onItemClick } = cardItemProps;
  const linkProps = getResultAnchorProps(item);

  // Build activity string (modified date + file type)
  const activityParts: string[] = [];
  if (item.modified) {
    activityParts.push('Modified ' + formatRelativeDate(item.modified));
  }
  if (item.fileType) {
    activityParts.push(item.fileType.toUpperCase());
  }
  const activityText: string = activityParts.join(' \u00B7 ');
  const activityTooltip: string = item.modified ? formatDateTime(item.modified) : '';

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
          <div className={styles.docCardTitleBar}>
            <DocumentTitleHoverCard item={item} position={position} onItemClick={onItemClick} hostDisplay="block">
              {(handleClick): React.ReactNode => (
                <a
                  href={linkProps.href}
                  target={linkProps.target}
                  rel={linkProps.rel}
                  className={styles.docCardTitleLink}
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

        {/* Activity: author persona + modified date */}
        {people.length > 0 && (
          <div title={activityTooltip}>
            <DocumentCardActivity
              activity={activityText}
              people={people}
            />
          </div>
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
  const { items, searchContextId, titleDisplayMode, onPreviewItem, onItemClick } = props;

  return (
    <div className={styles.cardGrid} role="list">
      {items.map((item: ISearchResult, index: number) => (
        <CardItem
          key={item.key}
          item={item}
          position={index + 1}
          searchContextId={searchContextId}
          titleDisplayMode={titleDisplayMode}
          onPreviewItem={onPreviewItem}
          onItemClick={onItemClick}
        />
      ))}
    </div>
  );
};

export default CardLayout;
