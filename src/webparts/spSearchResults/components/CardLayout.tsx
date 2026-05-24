import * as React from 'react';
import {
  DocumentCard,
  DocumentCardPreview,
  DocumentCardActivity,
  DocumentCardType,
} from '@fluentui/react/lib/DocumentCard';
import { ImageFit } from '@fluentui/react/lib/Image';
import { Icon } from '@fluentui/react/lib/Icon';
import { getFileTypeIconProps } from '@fluentui/react-file-type-icons';
import { ISearchResult } from '@interfaces/index';
import { formatRelativeDate, formatDateTime, formatTitleText, TitleDisplayMode } from './documentTitleUtils';
import { resolveResultLink, type IResultLinkConfig } from './resultLink';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import AddToCollectionButton from './AddToCollectionButton';
import styles from './SpSearchResults.module.scss';

export interface ICardLayoutProps {
  items: ISearchResult[];
  searchContextId: string;
  titleDisplayMode: TitleDisplayMode;
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
  // Stream C / #7
  linkConfig: IResultLinkConfig;
  onOpenInSidePanel?: (item: ISearchResult) => void;
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
const FileTypeIconPreview: React.FC<{ extension: string }> = (iconProps) => {
  return (
    <div className={styles.docCardIconPreview}>
      <Icon {...getFileTypeIconProps({ extension: iconProps.extension || '', size: 48 })} />
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
  linkConfig: IResultLinkConfig;
  onOpenInSidePanel?: (item: ISearchResult) => void;
}> = (cardItemProps) => {
  const { item, position, searchContextId, titleDisplayMode, onItemClick, linkConfig, onOpenInSidePanel } = cardItemProps;
  const linkProps = resolveResultLink(item, linkConfig);

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
          <FileTypeIconPreview extension={item.fileType || ''} />
        )}

        {/* Document title with HoverCard */}
        <div className={styles.docCardTitleWrapper} title={item.title}>
          <div className={styles.docCardTitleBar}>
            <DocumentTitleHoverCard
              item={item}
              position={position}
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
                  data-interception="off"
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
  const { items, searchContextId, titleDisplayMode, onPreviewItem, onItemClick, linkConfig, onOpenInSidePanel } = props;

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
          linkConfig={linkConfig}
          onOpenInSidePanel={onOpenInSidePanel}
        />
      ))}
    </div>
  );
};

export default CardLayout;
