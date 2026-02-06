import * as React from 'react';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISavedSearch } from '@interfaces/index';
import { SearchManagerService } from '@services/index';
import styles from './SpSearchManager.module.scss';

export interface IShareSearchDialogProps {
  isOpen: boolean;
  search: ISavedSearch | undefined;
  onDismiss: () => void;
  service?: SearchManagerService;
  context?: WebPartContext;
  onShareComplete?: () => void;
}

/**
 * ShareSearchDialog — tabbed dialog for sharing a saved search.
 * Tabs: Copy Link, Email, Teams, Share to Users
 */
const ShareSearchDialog: React.FC<IShareSearchDialogProps> = function ShareSearchDialog(props) {
  const { isOpen, search, onDismiss, service, context, onShareComplete } = props;

  // ─── Local state ──────────────────────────────────────────
  const [copied, setCopied] = React.useState<boolean>(false);
  const [emailSent, setEmailSent] = React.useState<boolean>(false);
  const [teamsSent, setTeamsSent] = React.useState<boolean>(false);
  const [selectedUsers, setSelectedUsers] = React.useState<string[]>([]);
  const [isSharing, setIsSharing] = React.useState<boolean>(false);
  const [shareError, setShareError] = React.useState<string | undefined>(undefined);
  const [shareSuccess, setShareSuccess] = React.useState<boolean>(false);
  const copyTimeoutRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  // ─── Cleanup copy timeout on unmount ──────────────────────
  React.useEffect(function (): () => void {
    return function cleanup(): void {
      if (copyTimeoutRef.current !== undefined) {
        clearTimeout(copyTimeoutRef.current);
      }
    };
  }, []);

  // ─── Reset state when dialog opens ─────────────────────────
  React.useEffect(function (): void {
    if (isOpen) {
      setCopied(false);
      setEmailSent(false);
      setTeamsSent(false);
      setSelectedUsers([]);
      setIsSharing(false);
      setShareError(undefined);
      setShareSuccess(false);
    }
  }, [isOpen]);

  // ─── Build the share URL ──────────────────────────────────
  let shareUrl = '';
  if (search) {
    shareUrl = search.searchUrl || '';
    if (shareUrl.startsWith('/')) {
      shareUrl = window.location.origin + shareUrl;
    }
  }

  // ─── Copy to clipboard helper ─────────────────────────────
  function copyToClipboard(text: string): void {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text)
        .then(function (): void {
          showCopied();
        })
        .catch(function noop(): void { /* swallow */ });
    } else {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      textarea.style.position = 'fixed';
      textarea.style.left = '-9999px';
      document.body.appendChild(textarea);
      textarea.select();
      try {
        document.execCommand('copy');
        showCopied();
      } catch {
        // Swallow copy errors
      }
      document.body.removeChild(textarea);
    }
  }

  function showCopied(): void {
    setCopied(true);
    if (copyTimeoutRef.current !== undefined) {
      clearTimeout(copyTimeoutRef.current);
    }
    copyTimeoutRef.current = setTimeout(function (): void {
      setCopied(false);
      copyTimeoutRef.current = undefined;
    }, 3000);
  }

  // ─── Copy Link handler ────────────────────────────────────
  function handleCopyLink(): void {
    if (!shareUrl) {
      return;
    }
    copyToClipboard(shareUrl);
  }

  // ─── Email handler ────────────────────────────────────────
  function handleShareEmail(): void {
    if (!search) {
      return;
    }

    const subject = 'Shared Search: ' + search.title;
    const body =
      'Check out this search: ' + search.title + '\n\n' +
      'Search query: ' + search.queryText + '\n\n' +
      (search.resultCount > 0 ? 'Results found: ' + search.resultCount + '\n\n' : '') +
      'Open in SharePoint:\n' + shareUrl;

    window.open(
      'mailto:?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(body),
      '_self'
    );
    setEmailSent(true);
  }

  // ─── Teams handler ────────────────────────────────────────
  function handleShareTeams(): void {
    if (!search) {
      return;
    }

    const message = 'Check out this search: ' + search.title + ' - ' + shareUrl;
    const teamsUrl = 'https://teams.microsoft.com/l/chat/0/0?message=' + encodeURIComponent(message);
    window.open(teamsUrl, '_blank');
    setTeamsSent(true);
  }

  // ─── Share to Users handler ───────────────────────────────
  function handlePeopleChanged(items: Array<{ secondaryText?: string }>): void {
    const emails: string[] = [];
    for (let i = 0; i < items.length; i++) {
      if (items[i].secondaryText) {
        emails.push(items[i].secondaryText as string);
      }
    }
    setSelectedUsers(emails);
    setShareError(undefined);
    setShareSuccess(false);
  }

  function handleShareToUsers(): void {
    if (!search || !service || selectedUsers.length === 0) {
      return;
    }

    setIsSharing(true);
    setShareError(undefined);
    setShareSuccess(false);

    service.shareToUsers(search.id, selectedUsers)
      .then(function (): void {
        setIsSharing(false);
        setShareSuccess(true);
        setSelectedUsers([]);
        if (onShareComplete) {
          onShareComplete();
        }
      })
      .catch(function (err: unknown): void {
        setIsSharing(false);
        const message = err instanceof Error ? err.message : 'Failed to share search';
        setShareError(message);
      });
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
      minWidth={520}
    >
      <div className={styles.shareDialogContent}>
        <Pivot aria-label="Share options">
          {/* ── Copy Link tab ───────────────────────────── */}
          <PivotItem headerText="Copy Link" itemIcon="Link">
            <div className={styles.shareLinkContainer}>
              {copied && (
                <div className={styles.successMessage}>
                  <Icon iconName="StatusCircleCheckmark" className={styles.successIcon} />
                  <span>Link copied to clipboard</span>
                </div>
              )}

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

          {/* ── Email tab ───────────────────────────────── */}
          <PivotItem headerText="Email" itemIcon="Mail">
            <div className={styles.shareLinkContainer}>
              {emailSent && (
                <div className={styles.successMessage}>
                  <Icon iconName="StatusCircleCheckmark" className={styles.successIcon} />
                  <span>Email client opened</span>
                </div>
              )}

              <p style={{ fontSize: '13px', color: '#323130', margin: '0 0 12px 0' }}>
                Open your email client with a pre-filled message containing the search link and details.
              </p>

              <PrimaryButton
                iconProps={{ iconName: 'Mail' }}
                text="Open in Email"
                onClick={handleShareEmail}
                disabled={!shareUrl}
              />
            </div>
          </PivotItem>

          {/* ── Teams tab ───────────────────────────────── */}
          <PivotItem headerText="Teams" itemIcon="TeamsLogo">
            <div className={styles.shareLinkContainer}>
              {teamsSent && (
                <div className={styles.successMessage}>
                  <Icon iconName="StatusCircleCheckmark" className={styles.successIcon} />
                  <span>Teams chat opened</span>
                </div>
              )}

              <p style={{ fontSize: '13px', color: '#323130', margin: '0 0 12px 0' }}>
                Open a new Teams chat with the search link pre-filled in the message box.
              </p>

              <PrimaryButton
                iconProps={{ iconName: 'TeamsLogo' }}
                text="Open in Teams"
                onClick={handleShareTeams}
                disabled={!shareUrl}
              />
            </div>
          </PivotItem>

          {/* ── Share to Users tab ──────────────────────── */}
          <PivotItem headerText="Users" itemIcon="People">
            <div className={styles.shareUserContainer}>
              {shareSuccess && (
                <MessageBar
                  messageBarType={MessageBarType.success}
                  onDismiss={function (): void { setShareSuccess(false); }}
                  dismissButtonAriaLabel="Close"
                >
                  Search shared successfully
                </MessageBar>
              )}

              {shareError && (
                <MessageBar
                  messageBarType={MessageBarType.error}
                  onDismiss={function (): void { setShareError(undefined); }}
                  dismissButtonAriaLabel="Close"
                >
                  {shareError}
                </MessageBar>
              )}

              {context && service ? (
                <>
                  <p style={{ fontSize: '13px', color: '#323130', margin: '0 0 12px 0' }}>
                    Share this search with specific people. They will see it in their Shared Searches tab.
                  </p>

                  <PeoplePicker
                    context={context as never}
                    titleText="Select people"
                    personSelectionLimit={10}
                    showtooltip={true}
                    required={false}
                    onChange={handlePeopleChanged}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={300}
                  />

                  <div style={{ marginTop: '16px' }}>
                    {isSharing ? (
                      <Spinner size={SpinnerSize.small} label="Sharing..." />
                    ) : (
                      <PrimaryButton
                        iconProps={{ iconName: 'Share' }}
                        text="Share"
                        onClick={handleShareToUsers}
                        disabled={selectedUsers.length === 0}
                      />
                    )}
                  </div>
                </>
              ) : (
                <div className={styles.shareUserPlaceholder}>
                  <Icon iconName="People" className={styles.shareUserPlaceholderIcon} />
                  <p>
                    People-based sharing requires additional configuration.
                    Use the Copy Link tab to share the search URL directly.
                  </p>
                </div>
              )}
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
