import * as React from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import type { IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import type { IActionProvider, ISearchContext, ISearchResult } from '@interfaces/index';
import { buildShareLines, copyTextToClipboard, normalizeUrl } from '@providers/actions/actionUtils';
import styles from './SpSearchResults.module.scss';

export interface IBulkActionsToolbarProps {
  selectedItems: ISearchResult[];
  actions: IActionProvider[];
  context: ISearchContext;
  onClearSelection: () => void;
}

function formatSelectionCount(count: number): string {
  if (count === 1) {
    return '1 item selected';
  }
  return String(count) + ' items selected';
}

const BulkActionsToolbar: React.FC<IBulkActionsToolbarProps> = (props) => {
  const { selectedItems, actions, context, onClearSelection } = props;
  const [runningActionId, setRunningActionId] = React.useState<string | undefined>(undefined);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

  React.useEffect(() => {
    setErrorMessage(undefined);
  }, [selectedItems.length]);

  const bulkActions = React.useMemo(() => {
    return actions.filter((action) => action.isBulkEnabled);
  }, [actions]);

  function isActionApplicable(action: IActionProvider): boolean {
    if (selectedItems.length === 0) {
      return false;
    }
    if (action.id === 'compare') {
      return selectedItems.length >= 2 && selectedItems.length <= 3;
    }
    for (let i = 0; i < selectedItems.length; i++) {
      if (!action.isApplicable(selectedItems[i])) {
        return false;
      }
    }
    return true;
  }

  function handleActionClick(action: IActionProvider): void {
    if (runningActionId) {
      return;
    }
    setRunningActionId(action.id);
    setErrorMessage(undefined);

    action.execute(selectedItems, context)
      .then(() => {
        setRunningActionId(undefined);
      })
      .catch((error) => {
        const message = error instanceof Error ? error.message : 'Action failed';
        setErrorMessage(message);
        setRunningActionId(undefined);
      });
  }

  function buildShareUrls(): string[] {
    const urls: string[] = [];
    for (let i = 0; i < selectedItems.length; i++) {
      const url = normalizeUrl(selectedItems[i].url);
      if (url) {
        urls.push(url);
      }
    }
    return urls;
  }

  function handleShareCopy(): void {
    setErrorMessage(undefined);
    const urls = buildShareUrls();
    if (urls.length === 0) {
      setErrorMessage('No valid URLs to share.');
      return;
    }
    copyTextToClipboard(urls.join('\n'))
      .catch(function (): void {
        setErrorMessage('Failed to copy links.');
      });
  }

  function handleShareEmail(): void {
    setErrorMessage(undefined);
    const lines = buildShareLines(selectedItems);
    if (lines.length === 0) {
      setErrorMessage('No valid items to share.');
      return;
    }
    const subject = 'Shared items (' + String(selectedItems.length) + ')';
    const body = 'Shared items:\n\n' + lines.join('\n');
    window.open(
      'mailto:?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(body),
      '_self'
    );
  }

  function handleShareTeams(): void {
    setErrorMessage(undefined);
    const lines = buildShareLines(selectedItems);
    if (lines.length === 0) {
      setErrorMessage('No valid items to share.');
      return;
    }
    const message = 'Shared items:\n' + lines.join('\n');
    const teamsUrl = 'https://teams.microsoft.com/l/chat/0/0?message=' + encodeURIComponent(message);
    window.open(teamsUrl, '_blank');
  }

  const shareMenuItems: IContextualMenuItem[] = React.useMemo(() => {
    return [
      {
        key: 'copyLinks',
        text: 'Copy links',
        iconProps: { iconName: 'Copy' },
        onClick: handleShareCopy
      },
      {
        key: 'email',
        text: 'Email',
        iconProps: { iconName: 'Mail' },
        onClick: handleShareEmail
      },
      {
        key: 'teams',
        text: 'Teams',
        iconProps: { iconName: 'TeamsLogo' },
        onClick: handleShareTeams
      }
    ];
  }, [selectedItems]);

  if (selectedItems.length === 0) {
    return null;
  }

  return (
    <div className={styles.bulkToolbar} role="region" aria-label="Bulk actions">
      <div className={styles.bulkToolbarLeft}>
        <Icon iconName="MultiSelect" className={styles.bulkToolbarIcon} />
        <span className={styles.bulkToolbarCount}>{formatSelectionCount(selectedItems.length)}</span>
        <DefaultButton
          text="Clear"
          onClick={onClearSelection}
          className={styles.bulkToolbarClear}
        />
      </div>
      <div className={styles.bulkToolbarRight}>
        {bulkActions.length === 0 && (
          <span className={styles.bulkToolbarEmpty}>No bulk actions available</span>
        )}
        {bulkActions.map((action) => {
          const applicable = isActionApplicable(action);
          const disabled = !applicable || !!runningActionId;
          if (action.id === 'share') {
            return (
              <DefaultButton
                key={action.id}
                text={action.label}
                iconProps={{ iconName: action.iconName }}
                disabled={disabled}
                className={styles.bulkToolbarAction}
                menuProps={{ items: shareMenuItems }}
              />
            );
          }

          return (
            <DefaultButton
              key={action.id}
              text={action.label}
              iconProps={{ iconName: action.iconName }}
              onClick={() => handleActionClick(action)}
              disabled={disabled}
              className={styles.bulkToolbarAction}
            />
          );
        })}
      </div>
      {errorMessage && (
        <div className={styles.bulkToolbarError} role="alert">{errorMessage}</div>
      )}
    </div>
  );
};

export default BulkActionsToolbar;
