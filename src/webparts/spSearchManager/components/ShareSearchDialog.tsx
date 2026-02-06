import * as React from 'react';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Icon } from '@fluentui/react/lib/Icon';
import { ISavedSearch } from '@interfaces/index';
import styles from './SpSearchManager.module.scss';

export interface IShareSearchDialogProps {
  isOpen: boolean;
  search: ISavedSearch | undefined;
  onDismiss: () => void;
}

/**
 * ShareSearchDialog -- dialog for sharing a saved search.
 * Provides two tabs:
 *   1. Copy Link -- displays the full search URL with a copy-to-clipboard button
 *   2. Share to Users -- placeholder for future People Picker integration
 */
const ShareSearchDialog: React.FC<IShareSearchDialogProps> = (props) => {
  const { isOpen, search, onDismiss } = props;

  // ─── Local state ──────────────────────────────────────────
  const [copied, setCopied] = React.useState<boolean>(false);
  const copyTimeoutRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  // ─── Cleanup copy timeout on unmount ──────────────────────
  React.useEffect(function (): () => void {
    return function cleanup(): void {
      if (copyTimeoutRef.current !== undefined) {
        clearTimeout(copyTimeoutRef.current);
      }
    };
  }, []);

  // ─── Reset copied state when dialog opens ─────────────────
  React.useEffect(function (): void {
    if (isOpen) {
      setCopied(false);
    }
  }, [isOpen]);

  // ─── Handlers ─────────────────────────────────────────────

  function handleCopyLink(): void {
    if (!search || !search.searchUrl) {
      return;
    }

    // Build full URL: if the searchUrl is relative, prepend location.origin
    let fullUrl = search.searchUrl;
    if (fullUrl.startsWith('/')) {
      fullUrl = window.location.origin + fullUrl;
    }

    // Copy to clipboard
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(fullUrl)
        .then(function (): void {
          setCopied(true);
          // Reset after 3 seconds
          if (copyTimeoutRef.current !== undefined) {
            clearTimeout(copyTimeoutRef.current);
          }
          copyTimeoutRef.current = setTimeout(function (): void {
            setCopied(false);
            copyTimeoutRef.current = undefined;
          }, 3000);
        })
        .catch(function noop(): void { /* swallow */ });
    } else {
      // Fallback for older browsers: create a temporary textarea
      const textarea = document.createElement('textarea');
      textarea.value = fullUrl;
      textarea.style.position = 'fixed';
      textarea.style.left = '-9999px';
      document.body.appendChild(textarea);
      textarea.select();
      try {
        document.execCommand('copy');
        setCopied(true);
        if (copyTimeoutRef.current !== undefined) {
          clearTimeout(copyTimeoutRef.current);
        }
        copyTimeoutRef.current = setTimeout(function (): void {
          setCopied(false);
          copyTimeoutRef.current = undefined;
        }, 3000);
      } catch {
        // Swallow copy errors
      }
      document.body.removeChild(textarea);
    }
  }

  // ─── Build the share URL ──────────────────────────────────
  let shareUrl = '';
  if (search) {
    shareUrl = search.searchUrl || '';
    if (shareUrl.startsWith('/')) {
      shareUrl = window.location.origin + shareUrl;
    }
  }

  return (
    <Dialog
      hidden={!isOpen}
      onDismiss={onDismiss}
      dialogContentProps={{
        type: DialogType.largeHeader,
        title: search ? 'Share: ' + search.title : 'Share search'
      }}
      modalProps={{ isBlocking: false }}
      minWidth={480}
    >
      <div className={styles.shareDialogContent}>
        <Pivot aria-label="Share options">
          {/* Copy Link tab */}
          <PivotItem headerText="Copy Link" itemIcon="Link">
            <div className={styles.shareLinkContainer}>
              {/* Success message */}
              {copied && (
                <div className={styles.successMessage}>
                  <Icon iconName="StatusCircleCheckmark" className={styles.successIcon} />
                  <span>Link copied to clipboard</span>
                </div>
              )}

              {/* URL display + copy button */}
              <div className={styles.shareLinkInput}>
                <div className={styles.shareLinkUrl}>
                  <TextField
                    value={shareUrl}
                    readOnly={true}
                    borderless={false}
                  />
                </div>
                <PrimaryButton
                  iconProps={{ iconName: copied ? 'StatusCircleCheckmark' : 'Copy' }}
                  text={copied ? 'Copied' : 'Copy'}
                  onClick={handleCopyLink}
                  disabled={!shareUrl}
                />
              </div>

              <p style={{ fontSize: '12px', color: '#605e5c', margin: 0 }}>
                Anyone with this link can use the same search parameters.
              </p>
            </div>
          </PivotItem>

          {/* Share to Users tab */}
          <PivotItem headerText="Share to Users" itemIcon="People">
            <div className={styles.shareUserContainer}>
              <div className={styles.shareUserPlaceholder}>
                <Icon iconName="People" className={styles.shareUserPlaceholderIcon} />
                <p>
                  People-based sharing will be available in a future update.
                  For now, use the Copy Link tab to share the search URL directly.
                </p>
              </div>
            </div>
          </PivotItem>
        </Pivot>
      </div>

      <DialogFooter>
        <DefaultButton onClick={onDismiss} text="Close" />
      </DialogFooter>
    </Dialog>
  );
};

export default ShareSearchDialog;
