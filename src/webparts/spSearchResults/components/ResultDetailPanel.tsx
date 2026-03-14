import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Icon } from '@fluentui/react/lib/Icon';
import { DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { UserPersona } from 'spfx-toolkit/lib/components/UserPersona';
import { LazyVersionHistory as _LazyVersionHistory } from './LazyVersionHistory';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const VersionHistory: any = _LazyVersionHistory;
import { ISearchResult } from '@interfaces/index';
import { formatRelativeDate, formatDateTime, formatFileSize, buildFormUrl } from './documentTitleUtils';
import styles from './SpSearchResults.module.scss';

export interface IResultDetailPanelProps {
  isOpen: boolean;
  item: ISearchResult | undefined;
  onDismiss: () => void;
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
 * Determines whether this file type supports WOPI preview.
 */
function supportsWopiPreview(fileType: string): boolean {
  const ft: string = (fileType || '').toLowerCase();
  const supported: string[] = [
    'docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt',
    'pdf', 'one', 'onetoc2', 'vsdx', 'vsd',
    'txt', 'csv'
  ];
  return supported.indexOf(ft) >= 0;
}

/**
 * Builds the preview URL for the detail panel iframe.
 * PDFs: direct file URL (browser native PDF viewer, no frame breakout).
 * Office docs: WopiFrame interactivepreview.
 */
function buildPreviewUrl(item: ISearchResult): string {
  const docUrl: string = item.url || '';
  if (!docUrl) {
    return '';
  }
  const ft: string = (item.fileType || '').toLowerCase();
  if (ft === 'pdf') {
    return docUrl;
  }
  const siteUrl: string = item.siteUrl || '';
  if (!siteUrl) {
    return '';
  }
  return siteUrl + '/_layouts/15/WopiFrame.aspx?sourcedoc=' +
    encodeURIComponent(docUrl) + '&action=interactivepreview';
}

/**
 * Shimmer loading placeholder for when the panel is opening.
 */
const PanelShimmer: React.FC = () => (
  <div className={styles.detailPanelShimmer}>
    <Shimmer
      shimmerElements={[
        { type: ShimmerElementType.line, height: 24, width: '60%' }
      ]}
      width="100%"
    />
    <Shimmer
      shimmerElements={[
        { type: ShimmerElementType.line, height: 200, width: '100%' }
      ]}
      width="100%"
      style={{ marginTop: 16 }}
    />
    <Shimmer
      shimmerElements={[
        { type: ShimmerElementType.circle, height: 40 },
        { type: ShimmerElementType.gap, width: 12 },
        { type: ShimmerElementType.line, height: 16, width: '40%' }
      ]}
      width="100%"
      style={{ marginTop: 20 }}
    />
    <Shimmer
      shimmerElements={[
        { type: ShimmerElementType.line, height: 14, width: '50%' }
      ]}
      width="100%"
      style={{ marginTop: 12 }}
    />
  </div>
);

/**
 * ResultDetailPanel — a slide-out panel that shows the full details of a
 * selected search result, including document preview, metadata, and actions.
 */
const ResultDetailPanel: React.FC<IResultDetailPanelProps> = (props) => {
  const { isOpen, item, onDismiss } = props;
  const [isPreviewLoaded, setIsPreviewLoaded] = React.useState<boolean>(false);
  const [linkCopied, setLinkCopied] = React.useState<boolean>(false);
  const [versionHistoryItem, setVersionHistoryItem] = React.useState<ISearchResult | undefined>(undefined);

  React.useEffect((): void => {
    setIsPreviewLoaded(false);
    setLinkCopied(false);
    setVersionHistoryItem(undefined);
  }, [item]);

  const handleIframeLoad = React.useCallback((): void => {
    setIsPreviewLoaded(true);
  }, []);

  // ─── Action handlers ───────────────────────────────────
  const handleOpenInBrowser = React.useCallback((): void => {
    if (item) {
      window.open(item.url, '_blank', 'noopener,noreferrer');
    }
  }, [item]);

  const handleDownload = React.useCallback((): void => {
    if (item) {
      const downloadUrl: string = item.url.indexOf('?') >= 0
        ? item.url + '&download=1'
        : item.url + '?download=1';
      window.open(downloadUrl, '_blank', 'noopener,noreferrer');
    }
  }, [item]);

  const handleCopyLink = React.useCallback((): void => {
    if (item && navigator.clipboard) {
      navigator.clipboard.writeText(item.url).then(function (): void {
        setLinkCopied(true);
        setTimeout(function (): void { setLinkCopied(false); }, 2000);
      }).catch((): void => {
        // Silently fail
      });
    }
  }, [item]);

  const handleOpenForm = React.useCallback((url: string): void => {
    window.open(url, '_blank', 'noopener,noreferrer');
  }, []);

  // ─── Render header content ─────────────────────────────
  const onRenderNavigationContent = React.useCallback(
    (): React.ReactElement => {
      return (
        <div className={styles.detailPanelNav}>
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            ariaLabel="Close panel"
            onClick={onDismiss}
            title="Close"
          />
        </div>
      );
    },
    [onDismiss]
  );

  if (!item) {
    return (
      <Panel
        isOpen={isOpen}
        type={PanelType.medium}
        onDismiss={onDismiss}
        isLightDismiss={true}
        hasCloseButton={true}
        headerText="Result Details"
      >
        <PanelShimmer />
      </Panel>
    );
  }

  const canPreview: boolean = supportsWopiPreview(item.fileType);
  const previewUrl: string = canPreview ? buildPreviewUrl(item) : '';
  const hasAuthorEmail: boolean = !!(item.author && item.author.email);
  const fileSizeStr: string = formatFileSize(item.fileSize);
  const viewUrl: string | undefined = buildFormUrl(item, 4);
  const editUrl: string | undefined = buildFormUrl(item, 6);
  const hasVersionHistory: boolean = !!(item.properties.ListId && item.properties.ListItemID);
  const hasFormLinks: boolean = !!(viewUrl || editUrl || hasVersionHistory);

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.medium}
      onDismiss={onDismiss}
      isLightDismiss={true}
      hasCloseButton={false}
      onRenderNavigationContent={onRenderNavigationContent}
    >
      <div className={styles.detailPanel}>
        {/* ─── Header Section ───────────────────────────── */}
        <div className={styles.detailPanelHeader}>
          <div className={styles.detailPanelTitleRow}>
            <span className={styles.detailPanelFileIcon}>
              <Icon iconName={getFileTypeIcon(item.fileType)} />
            </span>
            <div className={styles.detailPanelTitleGroup}>
              <h2 className={styles.detailPanelTitle}>{item.title}</h2>
              {item.fileType && (
                <span className={styles.detailPanelFileTypeBadge}>
                  {item.fileType.toUpperCase()}
                </span>
              )}
            </div>
          </div>

          {/* Quick action bar */}
          <div className={styles.detailPanelActionBar}>
            <DefaultButton
              text="Open"
              iconProps={{ iconName: 'OpenInNewTab' }}
              onClick={handleOpenInBrowser}
              className={styles.detailPanelActionPrimary}
            />
            <IconButton
              iconProps={{ iconName: 'Download' }}
              title="Download"
              ariaLabel="Download"
              onClick={handleDownload}
              className={styles.detailPanelActionIcon}
            />
            <TooltipHost content={linkCopied ? 'Copied!' : 'Copy link'}>
              <IconButton
                iconProps={{ iconName: linkCopied ? 'CheckMark' : 'Link' }}
                ariaLabel="Copy link"
                onClick={handleCopyLink}
                className={linkCopied ? styles.detailPanelActionIconSuccess : styles.detailPanelActionIcon}
              />
            </TooltipHost>
          </div>
        </div>

        {/* ─── Preview Section ──────────────────────────── */}
        {canPreview && previewUrl && (
          <div className={styles.detailPanelPreview}>
            {!isPreviewLoaded && (
              <div className={styles.detailPanelPreviewLoading}>
                <Shimmer
                  shimmerElements={[
                    { type: ShimmerElementType.line, height: 300, width: '100%' }
                  ]}
                  width="100%"
                />
              </div>
            )}
            <iframe
              className={styles.previewFrame}
              src={previewUrl}
              title={'Preview: ' + item.title}
              onLoad={handleIframeLoad}
              sandbox="allow-scripts allow-same-origin allow-forms allow-popups"
              style={{ display: isPreviewLoaded ? 'block' : 'none' }}
            />
          </div>
        )}

        {/* ─── Non-previewable fallback ─────────────────── */}
        {!canPreview && (
          <div className={styles.detailPanelNoPreview}>
            <Icon iconName={getFileTypeIcon(item.fileType)} style={{ fontSize: 48 }} />
            <p>Preview is not available for this file type.</p>
            <DefaultButton
              text="Open in browser"
              iconProps={{ iconName: 'OpenInNewTab' }}
              onClick={handleOpenInBrowser}
            />
          </div>
        )}

        {/* ─── Author & Dates Section ─────────────────── */}
        <div className={styles.detailPanelAuthorSection}>
          {hasAuthorEmail ? (
            <UserPersona
              userIdentifier={item.author.email}
              displayName={item.author.displayText}
              size={40}
              displayMode="avatarAndName"
              showSecondaryText={true}
              showLivePersona={false}
            />
          ) : (
            <div className={styles.detailPanelAuthorFallback}>
              <Icon iconName="Contact" className={styles.detailPanelAuthorFallbackIcon} />
              <span>{item.author?.displayText || 'Unknown'}</span>
            </div>
          )}

          <div className={styles.detailPanelDates}>
            {item.modified && (
              <div className={styles.detailPanelDateRow}>
                <Icon iconName="Edit" className={styles.detailPanelDateIcon} />
                <span className={styles.detailPanelDateLabel}>Modified</span>
                <TooltipHost content={formatDateTime(item.modified)}>
                  <span className={styles.detailPanelDateValue}>
                    {formatRelativeDate(item.modified)}
                  </span>
                </TooltipHost>
              </div>
            )}
            {item.created && (
              <div className={styles.detailPanelDateRow}>
                <Icon iconName="Calendar" className={styles.detailPanelDateIcon} />
                <span className={styles.detailPanelDateLabel}>Created</span>
                <TooltipHost content={formatDateTime(item.created)}>
                  <span className={styles.detailPanelDateValue}>
                    {formatRelativeDate(item.created)}
                  </span>
                </TooltipHost>
              </div>
            )}
          </div>
        </div>

        {/* ─── Properties Section ─────────────────────── */}
        <div className={styles.detailPanelPropsSection}>
          <h3 className={styles.detailPanelPropsTitle}>Properties</h3>
          <div className={styles.detailPanelPropsList}>
            <div className={styles.detailPanelPropRow}>
              <Icon iconName={getFileTypeIcon(item.fileType)} className={styles.detailPanelPropIcon} />
              <span className={styles.detailPanelPropLabel}>Type</span>
              <span className={styles.detailPanelPropValue}>
                {item.fileType ? item.fileType.toUpperCase() : '\u2014'}
              </span>
            </div>
            <div className={styles.detailPanelPropRow}>
              <Icon iconName="HardDrive" className={styles.detailPanelPropIcon} />
              <span className={styles.detailPanelPropLabel}>Size</span>
              <span className={styles.detailPanelPropValue}>
                {fileSizeStr || '\u2014'}
              </span>
            </div>
            {item.siteName && (
              <div className={styles.detailPanelPropRow}>
                <Icon iconName="Globe" className={styles.detailPanelPropIcon} />
                <span className={styles.detailPanelPropLabel}>Site</span>
                <a
                  href={item.siteUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className={styles.metadataSiteLink}
                >
                  {item.siteName}
                </a>
              </div>
            )}
          </div>
        </div>

        {/* ─── View / Edit / History Links ────────────── */}
        {hasFormLinks && (
          <div className={styles.detailPanelFormLinks}>
            {viewUrl && (
              <button
                type="button"
                className={styles.detailPanelFormLink}
                onClick={(): void => {
                  handleOpenForm(viewUrl);
                }}
              >
                <Icon iconName="View" className={styles.detailPanelFormLinkIcon} />
                View item
              </button>
            )}
            {editUrl && (
              <button
                type="button"
                className={styles.detailPanelFormLink}
                onClick={(): void => {
                  handleOpenForm(editUrl);
                }}
              >
                <Icon iconName="Edit" className={styles.detailPanelFormLinkIcon} />
                Edit item
              </button>
            )}
            {hasVersionHistory && (
              <button
                type="button"
                className={styles.detailPanelFormLink}
                onClick={(): void => { setVersionHistoryItem(item); }}
              >
                <Icon iconName="History" className={styles.detailPanelFormLinkIcon} />
                View history
              </button>
            )}
          </div>
        )}
      </div>

      {/* Version history (lazy-loaded) */}
      {versionHistoryItem && versionHistoryItem.properties.ListId && versionHistoryItem.properties.ListItemID && (
        <VersionHistory
          listId={versionHistoryItem.properties.ListId as string}
          itemId={Number(versionHistoryItem.properties.ListItemID)}
          onClose={(): void => { setVersionHistoryItem(undefined); }}
          allowCopyLink={true}
        />
      )}
    </Panel>
  );
};

export default ResultDetailPanel;
