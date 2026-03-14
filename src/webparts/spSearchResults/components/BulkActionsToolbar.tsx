import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { TextField } from '@fluentui/react/lib/TextField';
import type { IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import type { IActionProvider, ISearchCollection, ISearchContext, ISearchResult } from '@interfaces/index';
import { buildShareLines, copyTextToClipboard, normalizeUrl } from '@providers/actions/actionUtils';
import { getManagerService } from '@store/store';
import styles from './SpSearchResults.module.scss';

export interface IBulkActionsToolbarProps {
  selectedItems: ISearchResult[];
  actions: IActionProvider[];
  context: ISearchContext;
  onClearSelection: () => void;
}

const CREATE_COLLECTION_KEY = '__create_new_collection__';

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
  const [showCollectionDialog, setShowCollectionDialog] = React.useState<boolean>(false);
  const [availableCollections, setAvailableCollections] = React.useState<ISearchCollection[]>([]);
  const [selectedCollectionKey, setSelectedCollectionKey] = React.useState<string | undefined>(undefined);
  const [newCollectionName, setNewCollectionName] = React.useState<string>('');
  const [isLoadingCollections, setIsLoadingCollections] = React.useState<boolean>(false);
  const [collectionError, setCollectionError] = React.useState<string | undefined>(undefined);

  React.useEffect(() => {
    setErrorMessage(undefined);
  }, [selectedItems.length]);

  const bulkActions = React.useMemo(() => {
    return actions.filter((action) => action.isBulkEnabled);
  }, [actions]);

  const collectionOptions: IDropdownOption[] = React.useMemo(() => {
    const names = getUniqueCollectionNames(availableCollections);
    const options: IDropdownOption[] = [];
    for (let i = 0; i < names.length; i++) {
      options.push({
        key: names[i],
        text: names[i]
      });
    }
    options.push({
      key: CREATE_COLLECTION_KEY,
      text: 'Create new collection'
    });
    return options;
  }, [availableCollections]);

  const isCreatingNewCollection = selectedCollectionKey === CREATE_COLLECTION_KEY;
  const canSubmitCollection = React.useMemo(() => {
    if (isLoadingCollections || runningActionId === 'pin') {
      return false;
    }

    if (isCreatingNewCollection) {
      return newCollectionName.trim().length > 0;
    }

    return !!selectedCollectionKey;
  }, [isCreatingNewCollection, isLoadingCollections, newCollectionName, runningActionId, selectedCollectionKey]);

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

  function handleCollectionNameChange(
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    value?: string
  ): void {
    setNewCollectionName(value || '');
  }

  function handleCollectionOptionChange(
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void {
    if (!option) {
      return;
    }

    setSelectedCollectionKey(String(option.key));
    setCollectionError(undefined);
  }

  function handleDismissCollectionDialog(): void {
    if (runningActionId === 'pin') {
      return;
    }

    setShowCollectionDialog(false);
    setCollectionError(undefined);
    setIsLoadingCollections(false);
    setSelectedCollectionKey(undefined);
    setNewCollectionName('');
  }

  function handleOpenCollectionDialog(): void {
    const service = getManagerService(context.searchContextId);
    if (!service) {
      setErrorMessage('Search Manager is not initialized.');
      return;
    }

    setShowCollectionDialog(true);
    setCollectionError(undefined);
    setAvailableCollections([]);
    setSelectedCollectionKey(undefined);
    setNewCollectionName('');
    setIsLoadingCollections(true);

    service.loadCollections()
      .then((collections) => {
        const uniqueNames = getUniqueCollectionNames(collections);
        setAvailableCollections(collections);
        setSelectedCollectionKey(uniqueNames.length > 0 ? uniqueNames[0] : CREATE_COLLECTION_KEY);
        setIsLoadingCollections(false);
      })
      .catch((error) => {
        const message = error instanceof Error ? error.message : 'Failed to load collections.';
        setCollectionError(message);
        setSelectedCollectionKey(CREATE_COLLECTION_KEY);
        setIsLoadingCollections(false);
      });
  }

  async function handleConfirmCollection(): Promise<void> {
    const service = getManagerService(context.searchContextId);
    if (!service) {
      setCollectionError('Search Manager is not initialized.');
      return;
    }

    const existingNames = getUniqueCollectionNames(availableCollections);
    const trimmedNewCollectionName = newCollectionName.trim();
    const targetCollectionName = isCreatingNewCollection
      ? trimmedNewCollectionName
      : (selectedCollectionKey || '');

    if (!targetCollectionName) {
      setCollectionError('Choose a collection or create a new one.');
      return;
    }

    const uniqueItems = getUniqueItemsForCollection(selectedItems);
    if (uniqueItems.length === 0) {
      setCollectionError('No valid result links were selected.');
      return;
    }

    setRunningActionId('pin');
    setCollectionError(undefined);
    setErrorMessage(undefined);

    try {
      if (existingNames.indexOf(targetCollectionName) < 0) {
        await service.createCollection(targetCollectionName);
      }

      for (let i = 0; i < uniqueItems.length; i++) {
        const item = uniqueItems[i];
        await service.pinToCollection(
          targetCollectionName,
          item.url,
          item.title || item.url,
          item.properties || {}
        );
      }

      setShowCollectionDialog(false);
      setAvailableCollections([]);
      setSelectedCollectionKey(undefined);
      setNewCollectionName('');
      setRunningActionId(undefined);
      onClearSelection();
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to add items to the collection.';
      setCollectionError(message);
      setRunningActionId(undefined);
    }
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
    <>
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

            if (action.id === 'pin') {
              return (
                <DefaultButton
                  key={action.id}
                  text={action.label}
                  iconProps={{ iconName: action.iconName }}
                  onClick={handleOpenCollectionDialog}
                  disabled={disabled}
                  className={styles.bulkToolbarAction}
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
      <Dialog
        hidden={!showCollectionDialog}
        onDismiss={handleDismissCollectionDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Add selected results to a collection'
        }}
        modalProps={{ isBlocking: true }}
      >
        <div className={styles.bulkToolbarDialogIntro}>
          Collections save the results you found. Saved searches keep the query, filters, and vertical you used.
        </div>
        {collectionError && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            {collectionError}
          </MessageBar>
        )}
        {isLoadingCollections ? (
          <div className={styles.bulkToolbarDialogSpinner}>
            <Spinner size={SpinnerSize.medium} label="Loading collections..." />
          </div>
        ) : (
          <div className={styles.bulkToolbarDialogForm}>
            <Dropdown
              label="Add to"
              options={collectionOptions}
              selectedKey={selectedCollectionKey}
              onChange={handleCollectionOptionChange}
              placeholder="Select a collection"
            />
            {isCreatingNewCollection && (
              <TextField
                label="New collection name"
                value={newCollectionName}
                onChange={handleCollectionNameChange}
                placeholder="For example: Q1 review set"
                autoFocus={true}
              />
            )}
            <div className={styles.bulkToolbarDialogHint}>
              {String(selectedItems.length) + ' selected result' + (selectedItems.length === 1 ? '' : 's') + ' will be added.'}
            </div>
          </div>
        )}
        <DialogFooter>
          <PrimaryButton
            onClick={handleConfirmCollection}
            text={runningActionId === 'pin' ? 'Adding...' : 'Add to Collection'}
            disabled={!canSubmitCollection}
          />
          <DefaultButton
            onClick={handleDismissCollectionDialog}
            text="Cancel"
            disabled={runningActionId === 'pin'}
          />
        </DialogFooter>
      </Dialog>
    </>
  );
};

function getUniqueCollectionNames(collections: ISearchCollection[]): string[] {
  const names: string[] = [];
  const seenNames = new Set<string>();

  for (let i = 0; i < collections.length; i++) {
    const collectionName = collections[i].collectionName;
    if (!collectionName || seenNames.has(collectionName)) {
      continue;
    }

    names.push(collectionName);
    seenNames.add(collectionName);
  }

  return names;
}

function getUniqueItemsForCollection(items: ISearchResult[]): ISearchResult[] {
  const uniqueItems: ISearchResult[] = [];
  const seenUrls = new Set<string>();

  for (let i = 0; i < items.length; i++) {
    const normalizedUrl = normalizeUrl(items[i].url);
    if (!normalizedUrl || seenUrls.has(normalizedUrl)) {
      continue;
    }

    uniqueItems.push(items[i]);
    seenUrls.add(normalizedUrl);
  }

  return uniqueItems;
}

export default BulkActionsToolbar;
