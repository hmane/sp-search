import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { HoverCard, HoverCardType } from '@fluentui/react/lib/HoverCard';
import { Modal } from '@fluentui/react/lib/Modal';
import { FileTypeIcon, IconType, ImageSize } from '@pnp/spfx-controls-react/lib/FileTypeIcon';
import { UserPersona as _UserPersona } from 'spfx-toolkit/lib/components/UserPersona';
import { LazyVersionHistory as _LazyVersionHistory } from 'spfx-toolkit/lib/components/lazy';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const UserPersona: any = _UserPersona;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const VersionHistory: any = _LazyVersionHistory;
import { ISearchResult } from '@interfaces/index';
import { formatFileSize, formatDateTime, buildPreviewUrl, buildFormUrl } from './documentTitleUtils';
import styles from './SpSearchResults.module.scss';

export interface IDocumentTitleHoverCardProps {
  item: ISearchResult;
  position: number;
  onItemClick?: (item: ISearchResult, position: number) => void;
  children: (handleClick: (e: React.MouseEvent) => void) => React.ReactNode;
  hostDisplay?: 'inline' | 'block';
  disabled?: boolean;
}

const DocumentTitleHoverCard: React.FC<IDocumentTitleHoverCardProps> = (props) => {
  const { item, position, onItemClick, children, hostDisplay, disabled } = props;
  const [previewItem, setPreviewItem] = React.useState<ISearchResult | undefined>(undefined);
  const [formModalUrl, setFormModalUrl] = React.useState<string | undefined>(undefined);
  const [formModalTitle, setFormModalTitle] = React.useState<string>('');
  const [versionHistoryItem, setVersionHistoryItem] = React.useState<ISearchResult | undefined>(undefined);

  const handleDismissPreview = React.useCallback((): void => {
    setPreviewItem(undefined);
  }, []);

  const handleClick = React.useCallback((e: React.MouseEvent): void => {
    if (onItemClick) {
      onItemClick(item, position);
    }
    const previewUrl = buildPreviewUrl(item);
    if (previewUrl) {
      e.preventDefault();
      setPreviewItem(item);
    }
  }, [item, position, onItemClick]);

  const renderPlainCard = React.useCallback((): JSX.Element => {
    const sizeDisplay: string = formatFileSize(item.fileSize);
    const viewUrl: string | undefined = buildFormUrl(item, 4);
    const editUrl: string | undefined = buildFormUrl(item, 6);
    const hasVersionHistory: boolean = !!(item.properties.ListId && item.properties.ListItemID);
    const hasActions: boolean = !!(viewUrl || editUrl || hasVersionHistory);

    return (
      <div className={styles.resultTitleHoverCard}>
        {/* Header: file icon + title + size */}
        <div className={styles.hoverCardHeader}>
          <div className={styles.hoverCardFileIcon}>
            <FileTypeIcon type={IconType.image} path={item.url} size={ImageSize.small} />
          </div>
          <div className={styles.hoverCardTitleInfo}>
            <p className={styles.hoverCardTitle}>{item.title}</p>
            {sizeDisplay && (
              <span className={styles.hoverCardFileSize}>{sizeDisplay}</span>
            )}
          </div>
        </div>

        <hr className={styles.hoverCardDivider} />

        {/* Thumbnail preview */}
        {item.thumbnailUrl && (
          <div className={styles.hoverCardThumbnail}>
            <img src={item.thumbnailUrl} alt="" loading="lazy" />
          </div>
        )}

        {/* Author with persona */}
        {item.author && item.author.displayText && (
          <div className={styles.hoverCardPersonaRow}>
            <UserPersona
              userIdentifier={item.author.email || item.author.displayText}
              displayName={item.author.displayText}
              size={32}
              displayMode="avatar"
            />
            <div className={styles.hoverCardPersonaInfo}>
              <span className={styles.hoverCardPersonaName}>
                {item.author.displayText}
              </span>
              {item.created && (
                <span className={styles.hoverCardDateLabel}>
                  {'Created ' + formatDateTime(item.created)}
                </span>
              )}
              {item.modified && (
                <span className={styles.hoverCardDateLabel}>
                  {'Modified ' + formatDateTime(item.modified)}
                </span>
              )}
            </div>
          </div>
        )}

        <hr className={styles.hoverCardDivider} />

        {/* Meta: file type + site name */}
        <div className={styles.hoverCardMeta}>
          <div className={styles.hoverCardMetaLeft}>
            {item.fileType && (
              <span className={styles.hoverCardMetaItem}>
                <Icon iconName="Page" style={{ fontSize: 12 }} />
                {item.fileType.toUpperCase()}
              </span>
            )}
            {item.siteName && (
              <span className={styles.hoverCardMetaItem}>
                <Icon iconName="SharePointLogo" style={{ fontSize: 12 }} />
                {item.siteName}
              </span>
            )}
          </div>
        </div>

        {/* Action links: View, Edit, Version History */}
        {hasActions && (
          <>
            <hr className={styles.hoverCardDivider} />
            <div className={styles.hoverCardActions}>
              {viewUrl && (
                <a
                  href={viewUrl}
                  className={styles.hoverCardActionLink}
                  onClick={(e: React.MouseEvent): void => {
                    e.preventDefault();
                    e.stopPropagation();
                    setFormModalUrl(viewUrl);
                    setFormModalTitle('View: ' + item.title);
                  }}
                >
                  <Icon iconName="View" style={{ fontSize: 12 }} />
                  {'View item'}
                </a>
              )}
              {editUrl && (
                <a
                  href={editUrl}
                  className={styles.hoverCardActionLink}
                  onClick={(e: React.MouseEvent): void => {
                    e.preventDefault();
                    e.stopPropagation();
                    setFormModalUrl(editUrl);
                    setFormModalTitle('Edit: ' + item.title);
                  }}
                >
                  <Icon iconName="Edit" style={{ fontSize: 12 }} />
                  {'Edit item'}
                </a>
              )}
              {hasVersionHistory && (
                <a
                  href="#"
                  className={styles.hoverCardActionLink}
                  onClick={(e: React.MouseEvent): void => {
                    e.preventDefault();
                    e.stopPropagation();
                    setVersionHistoryItem(item);
                  }}
                >
                  <Icon iconName="History" style={{ fontSize: 12 }} />
                  {'View history'}
                </a>
              )}
            </div>
          </>
        )}
      </div>
    );
  }, [item]);

  const display: string = hostDisplay === 'block' ? 'block' : 'inline';

  return (
    <>
      {disabled ? (
        children(handleClick)
      ) : (
        <HoverCard
          type={HoverCardType.plain}
          plainCardProps={{ onRenderPlainCard: renderPlainCard }}
          instantOpenOnClick={false}
          styles={{ host: { display } }}
        >
          {children(handleClick)}
        </HoverCard>
      )}

      {/* Document preview modal */}
      {previewItem && (
        <Modal
          isOpen={true}
          onDismiss={handleDismissPreview}
          isBlocking={true}
          styles={{
            main: {
              width: '90vw',
              maxWidth: '1280px',
              height: '90vh',
              padding: 0,
              display: 'flex',
              flexDirection: 'column',
            },
          }}
        >
          <div className={styles.previewModalHeader}>
            <span className={styles.previewModalTitle}>{previewItem.title}</span>
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              ariaLabel="Close preview"
              onClick={handleDismissPreview}
            />
          </div>
          <div className={styles.previewModalFrame}>
            <iframe
              src={buildPreviewUrl(previewItem)}
              title={previewItem.title}
              // eslint-disable-next-line react/no-unknown-property
              allowFullScreen
            />
          </div>
        </Modal>
      )}

      {/* View/Edit form modal */}
      {formModalUrl && (
        <Modal
          isOpen={true}
          onDismiss={(): void => { setFormModalUrl(undefined); }}
          isBlocking={false}
          styles={{
            main: {
              width: '90vw',
              maxWidth: '800px',
              height: '85vh',
              padding: 0,
              display: 'flex',
              flexDirection: 'column',
            },
          }}
        >
          <div className={styles.previewModalHeader}>
            <span className={styles.previewModalTitle}>{formModalTitle}</span>
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              ariaLabel="Close"
              onClick={(): void => { setFormModalUrl(undefined); }}
            />
          </div>
          <div className={styles.previewModalFrame}>
            <iframe
              src={formModalUrl + '&IsDlg=1'}
              title={formModalTitle}
            />
          </div>
        </Modal>
      )}

      {/* Version history (lazy-loaded) */}
      {versionHistoryItem && versionHistoryItem.properties.ListId && versionHistoryItem.properties.ListItemID && (
        <VersionHistory
          listId={versionHistoryItem.properties.ListId as string}
          itemId={Number(versionHistoryItem.properties.ListItemID)}
          onClose={(): void => { setVersionHistoryItem(undefined); }}
          allowCopyLink={true}
        />
      )}
    </>
  );
};

export default DocumentTitleHoverCard;
