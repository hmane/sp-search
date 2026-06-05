import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { HoverCard, HoverCardType } from '@fluentui/react/lib/HoverCard';
import { Modal } from '@fluentui/react/lib/Modal';
import { getFileTypeIconProps } from '@fluentui/react-file-type-icons';
import { UserPersona as _UserPersona } from 'spfx-toolkit/lib/components/UserPersona';
import { LazyVersionHistory as _LazyVersionHistory } from './LazyVersionHistory';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const UserPersona: any = _UserPersona;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const VersionHistory: any = _LazyVersionHistory;
import { ISearchResult } from '@interfaces/index';
import { formatFileSize, formatDateTime, buildPreviewUrl, buildFormUrl, isImageType } from './documentTitleUtils';
import type { ResultClickTarget } from './resultLink';
import styles from './SpSearchResults.module.scss';

export interface IDocumentTitleHoverCardProps {
  item: ISearchResult;
  position: number;
  onItemClick?: (item: ISearchResult, position: number) => void;
  children: (handleClick: (e: React.MouseEvent) => void) => React.ReactNode;
  hostDisplay?: 'inline' | 'block';
  disabled?: boolean;
  /**
   * Stream C / #7. When omitted, behaves as 'panel' (today's behaviour):
   * previewable files open the centred preview Modal on click. In
   * 'newTab'/'sameTab' the Modal is suppressed (`<a>` navigates). In
   * 'sidePanel' the click is intercepted and `onOpenInSidePanel` is invoked
   * (to call `store.setPreviewItem(item)` and open `ResultDetailPanel`).
   */
  clickTarget?: ResultClickTarget;
  onOpenInSidePanel?: (item: ISearchResult) => void;
}

const DocumentTitleHoverCard: React.FC<IDocumentTitleHoverCardProps> = (props) => {
  const { item, position, onItemClick, children, hostDisplay, disabled, clickTarget, onOpenInSidePanel } = props;
  const [previewItem, setPreviewItem] = React.useState<ISearchResult | undefined>(undefined);
  const [versionHistoryItem, setVersionHistoryItem] = React.useState<ISearchResult | undefined>(undefined);

  const handleDismissPreview = React.useCallback((): void => {
    setPreviewItem(undefined);
  }, []);

  const handleClick = React.useCallback((e: React.MouseEvent): void => {
    // Always log the click (history) — regardless of mode.
    if (onItemClick) {
      onItemClick(item, position);
    }

    const mode: ResultClickTarget = clickTarget || 'panel';

    // sidePanel — intercept and open ResultDetailPanel via the parent callback.
    if (mode === 'sidePanel' && onOpenInSidePanel) {
      e.preventDefault();
      e.nativeEvent.preventDefault();
      e.stopPropagation();
      onOpenInSidePanel(item);
      return;
    }

    // panel (default) — today's behaviour: previewable files → centred Modal.
    if (mode === 'panel') {
      const previewUrl = buildPreviewUrl(item);
      if (previewUrl) {
        e.preventDefault();
        e.nativeEvent.preventDefault();
        e.stopPropagation();
        setPreviewItem(item);
      }
      return;
    }

    // newTab / sameTab — let the <a> navigate naturally. No Modal even for previewable files.
  }, [item, position, onItemClick, clickTarget, onOpenInSidePanel]);

  const openFormInNewTab = React.useCallback((url: string): void => {
    window.open(url, '_blank', 'noopener,noreferrer');
  }, []);

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
            <Icon {...getFileTypeIconProps({ extension: item.fileType || '', size: 16 })} />
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
                <button
                  type="button"
                  className={styles.hoverCardActionLink}
                  onClick={(): void => {
                    openFormInNewTab(viewUrl);
                  }}
                >
                  <Icon iconName="View" style={{ fontSize: 12 }} />
                  {'View item'}
                </button>
              )}
              {editUrl && (
                <button
                  type="button"
                  className={styles.hoverCardActionLink}
                  onClick={(): void => {
                    openFormInNewTab(editUrl);
                  }}
                >
                  <Icon iconName="Edit" style={{ fontSize: 12 }} />
                  {'Edit item'}
                </button>
              )}
              {hasVersionHistory && (
                <button
                  type="button"
                  className={styles.hoverCardActionLink}
                  onClick={(): void => {
                    setVersionHistoryItem(item);
                  }}
                >
                  <Icon iconName="History" style={{ fontSize: 12 }} />
                  {'View history'}
                </button>
              )}
            </div>
          </>
        )}
      </div>
    );
  }, [item, openFormInNewTab]);

  const display: string = hostDisplay === 'block' ? 'block' : 'inline';
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const plainCardPropsCompat = { onRenderPlainCard: renderPlainCard } as any;

  return (
    <>
      {disabled ? (
        children(handleClick)
      ) : (
        <HoverCard
          type={HoverCardType.plain}
          plainCardProps={plainCardPropsCompat}
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
              width: 'calc(100vw - 48px)',
              maxWidth: 'calc(100vw - 48px)',
              height: 'calc(100vh - 48px)',
              padding: 0,
              display: 'flex',
              flexDirection: 'column',
            },
            scrollableContent: {
              display: 'flex',
              flexDirection: 'column',
              width: '100%',
              height: '100%'
            }
          }}
        >
          <div className={styles.previewModalHeader}>
            <span className={styles.previewModalTitle}>{previewItem.title}</span>
            <div className={styles.previewModalActions}>
              <TooltipHost content="Open in new tab">
                <IconButton
                  iconProps={{ iconName: 'OpenInNewTab' }}
                  ariaLabel="Open in new tab"
                  onClick={(): void => { window.open(previewItem.url, '_blank', 'noopener,noreferrer'); }}
                />
              </TooltipHost>
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                ariaLabel="Close preview"
                onClick={handleDismissPreview}
              />
            </div>
          </div>
          <div className={styles.previewModalFrame}>
            {((): React.ReactElement => {
              const ext: string = (previewItem.fileType || '').toLowerCase();
              const urlExt: string = (previewItem.url || '').toLowerCase().split('?')[0].split('#')[0];
              const isPdf: boolean = ext === 'pdf' || urlExt.endsWith('.pdf');

              if (isImageType(previewItem)) {
                // Stream C / #8 — render the image directly (clean fullscreen
                // view) instead of an iframe wrapping `?web=1`.
                return (
                  <img
                    className={styles.previewModalImage}
                    src={previewItem.url}
                    alt={previewItem.title}
                  />
                );
              }

              if (isPdf) {
                // <embed> uses the browser's native PDF plugin directly —
                // sidesteps Chrome's "This page has been blocked by Chrome"
                // failure mode that hits sandboxed iframes loading PDFs.
                // No iframe means no sandbox-vs-PDF-viewer conflict, and no
                // top-frame-navigation risk either (no scripting context).
                return (
                  <embed
                    src={previewItem.url}
                    type="application/pdf"
                    width="100%"
                    height="100%"
                  />
                );
              }

              // Office / text / csv / json / xml — WopiFrame in a sandboxed
              // iframe. allow-scripts + allow-same-origin: Office Online runtime.
              // allow-popups: "Open in app" links. allow-top-navigation is
              // deliberately omitted so the iframe can't break out of the Modal.
              return (
                <iframe
                  src={buildPreviewUrl(previewItem)}
                  title={previewItem.title}
                  sandbox="allow-scripts allow-same-origin allow-popups allow-forms"
                  // eslint-disable-next-line react/no-unknown-property
                  allowFullScreen
                />
              );
            })()}
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
