import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Icon } from '@fluentui/react/lib/Icon';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import FormContainer from 'spfx-toolkit/lib/components/spForm/FormContainer/FormContainer';
import FormItem from 'spfx-toolkit/lib/components/spForm/FormItem/FormItem';
import FormLabel from 'spfx-toolkit/lib/components/spForm/FormLabel/FormLabel';
import FormValue from 'spfx-toolkit/lib/components/spForm/FormValue/FormValue';
import { ISearchResult } from '@interfaces/index';
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
 * Formats a file size in bytes into a human-readable string.
 */
function formatFileSize(bytes: number): string {
  if (!bytes || bytes <= 0) {
    return 'Unknown';
  }
  if (bytes < 1024) {
    return bytes + ' B';
  }
  if (bytes < 1048576) {
    return Math.round(bytes / 1024) + ' KB';
  }
  if (bytes < 1073741824) {
    return (bytes / 1048576).toFixed(1) + ' MB';
  }
  return (bytes / 1073741824).toFixed(2) + ' GB';
}

/**
 * Formats an ISO date string into a readable format.
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
    const hours: number = d.getHours();
    const minutes: number = d.getMinutes();
    const minuteStr: string = minutes < 10 ? '0' + minutes : String(minutes);
    const ampm: string = hours >= 12 ? 'PM' : 'AM';
    const hour12: number = hours % 12 || 12;
    return month + '/' + day + '/' + year + ' ' + hour12 + ':' + minuteStr + ' ' + ampm;
  } catch {
    return '';
  }
}

/**
 * Returns the file type display text (uppercase extension).
 */
function formatFileType(fileType: string): string {
  if (!fileType) {
    return 'Unknown';
  }
  return fileType.toUpperCase();
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
 * Builds the WOPI preview URL for SharePoint documents.
 * Uses the _layouts/15/WopiFrame.aspx endpoint with interactivepreview action.
 */
function buildPreviewUrl(item: ISearchResult): string {
  const siteUrl: string = item.siteUrl || '';
  const docUrl: string = item.url || '';
  if (!siteUrl || !docUrl) {
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
        { type: ShimmerElementType.line, height: 16, width: '40%' }
      ]}
      width="100%"
      style={{ marginTop: 16 }}
    />
    <Shimmer
      shimmerElements={[
        { type: ShimmerElementType.line, height: 16, width: '70%' }
      ]}
      width="100%"
      style={{ marginTop: 8 }}
    />
    <Shimmer
      shimmerElements={[
        { type: ShimmerElementType.line, height: 16, width: '50%' }
      ]}
      width="100%"
      style={{ marginTop: 8 }}
    />
  </div>
);

/**
 * ResultDetailPanel — a slide-out panel that shows the full details of a
 * selected search result, including document preview, metadata, and actions.
 *
 * This component is intended to be lazy-loaded:
 * ```typescript
 * const ResultDetailPanel = React.lazy(() => import('./ResultDetailPanel'));
 * ```
 */
const ResultDetailPanel: React.FC<IResultDetailPanelProps> = (props) => {
  const { isOpen, item, onDismiss } = props;
  const [isPreviewLoaded, setIsPreviewLoaded] = React.useState<boolean>(false);

  // Reset preview loaded state when item changes
  React.useEffect((): void => {
    setIsPreviewLoaded(false);
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
      // Construct the download URL — append ?download=1 for SharePoint documents
      const downloadUrl: string = item.url.indexOf('?') >= 0
        ? item.url + '&download=1'
        : item.url + '?download=1';
      window.open(downloadUrl, '_blank', 'noopener,noreferrer');
    }
  }, [item]);

  const handleCopyLink = React.useCallback((): void => {
    if (item && navigator.clipboard) {
      navigator.clipboard.writeText(item.url).catch((): void => {
        // Silently fail if clipboard API not available
      });
    }
  }, [item]);

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
            <h2 className={styles.detailPanelTitle}>{item.title}</h2>
          </div>
          <div className={styles.detailPanelFileTypeBadge}>
            {formatFileType(item.fileType)}
          </div>
          <PrimaryButton
            className={styles.detailPanelOpenButton}
            text="Open"
            iconProps={{ iconName: 'OpenInNewTab' }}
            onClick={handleOpenInBrowser}
          />
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

        {/* ─── Metadata Section ─────────────────────────── */}
        <div className={styles.metadataSection}>
          <h3 className={styles.metadataSectionTitle}>Details</h3>
          <FormContainer labelWidth="120px">
            <FormItem>
              <FormLabel>Author</FormLabel>
              <FormValue>
                <span>
                  {item.author && item.author.displayText
                    ? item.author.displayText
                    : 'Unknown'}
                </span>
              </FormValue>
            </FormItem>

            <FormItem>
              <FormLabel>Modified</FormLabel>
              <FormValue>
                <span>{formatDate(item.modified)}</span>
              </FormValue>
            </FormItem>

            <FormItem>
              <FormLabel>Created</FormLabel>
              <FormValue>
                <span>{formatDate(item.created)}</span>
              </FormValue>
            </FormItem>

            <FormItem>
              <FormLabel>File Type</FormLabel>
              <FormValue>
                <span className={styles.metadataFileType}>
                  <Icon
                    iconName={getFileTypeIcon(item.fileType)}
                    style={{ fontSize: 14, marginRight: 6 }}
                  />
                  {formatFileType(item.fileType)}
                </span>
              </FormValue>
            </FormItem>

            <FormItem>
              <FormLabel>File Size</FormLabel>
              <FormValue>
                <span>{formatFileSize(item.fileSize)}</span>
              </FormValue>
            </FormItem>

            <FormItem>
              <FormLabel>Site</FormLabel>
              <FormValue>
                <span>
                  {item.siteName ? (
                    <a
                      href={item.siteUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className={styles.metadataSiteLink}
                    >
                      {item.siteName}
                    </a>
                  ) : (
                    'Unknown'
                  )}
                </span>
              </FormValue>
            </FormItem>
          </FormContainer>
        </div>

        {/* ─── Action Buttons ───────────────────────────── */}
        <div className={styles.detailPanelActions}>
          <DefaultButton
            text="Open in browser"
            iconProps={{ iconName: 'OpenInNewTab' }}
            onClick={handleOpenInBrowser}
          />
          <DefaultButton
            text="Download"
            iconProps={{ iconName: 'Download' }}
            onClick={handleDownload}
          />
          <DefaultButton
            text="Copy link"
            iconProps={{ iconName: 'Link' }}
            onClick={handleCopyLink}
          />
        </div>
      </div>
    </Panel>
  );
};

export default ResultDetailPanel;
