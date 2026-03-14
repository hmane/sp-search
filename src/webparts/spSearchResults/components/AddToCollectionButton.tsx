import * as React from 'react';
import { DefaultButton, IconButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { TextField } from '@fluentui/react/lib/TextField';
import type { ISearchCollection, ISearchResult } from '@interfaces/index';
import { getManagerService } from '@store/store';
import styles from './SpSearchResults.module.scss';

const CREATE_COLLECTION_KEY = '__create_new_collection__';

export interface IAddToCollectionButtonProps {
  item: ISearchResult;
  searchContextId: string;
  buttonClassName?: string;
}

const AddToCollectionButton: React.FC<IAddToCollectionButtonProps> = (props) => {
  const { item, searchContextId, buttonClassName } = props;
  const [isOpen, setIsOpen] = React.useState<boolean>(false);
  const [availableCollections, setAvailableCollections] = React.useState<ISearchCollection[]>([]);
  const [selectedCollectionKey, setSelectedCollectionKey] = React.useState<string | undefined>(undefined);
  const [newCollectionName, setNewCollectionName] = React.useState<string>('');
  const [isLoadingCollections, setIsLoadingCollections] = React.useState<boolean>(false);
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

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
  const canSubmit = React.useMemo(() => {
    if (isLoadingCollections || isSaving || !item.url) {
      return false;
    }
    if (isCreatingNewCollection) {
      return newCollectionName.trim().length > 0;
    }
    return !!selectedCollectionKey;
  }, [isCreatingNewCollection, isLoadingCollections, isSaving, item.url, newCollectionName, selectedCollectionKey]);

  function handleOpenDialog(event: React.MouseEvent<HTMLElement>): void {
    event.preventDefault();
    event.stopPropagation();

    const service = getManagerService(searchContextId);
    if (!service) {
      setErrorMessage('Search Manager is not initialized.');
      setIsOpen(true);
      return;
    }

    setIsOpen(true);
    setErrorMessage(undefined);
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
        setErrorMessage(message);
        setSelectedCollectionKey(CREATE_COLLECTION_KEY);
        setIsLoadingCollections(false);
      });
  }

  function handleDismiss(): void {
    if (isSaving) {
      return;
    }

    setIsOpen(false);
    setErrorMessage(undefined);
    setAvailableCollections([]);
    setSelectedCollectionKey(undefined);
    setNewCollectionName('');
  }

  function handleCollectionOptionChange(
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void {
    if (!option) {
      return;
    }

    setSelectedCollectionKey(String(option.key));
    setErrorMessage(undefined);
  }

  function handleCollectionNameChange(
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    value?: string
  ): void {
    setNewCollectionName(value || '');
  }

  async function handleConfirm(): Promise<void> {
    const service = getManagerService(searchContextId);
    if (!service) {
      setErrorMessage('Search Manager is not initialized.');
      return;
    }

    const existingNames = getUniqueCollectionNames(availableCollections);
    const targetCollectionName = isCreatingNewCollection
      ? newCollectionName.trim()
      : (selectedCollectionKey || '');

    if (!targetCollectionName) {
      setErrorMessage('Choose a collection or create a new one.');
      return;
    }

    if (!item.url) {
      setErrorMessage('This result does not have a valid URL.');
      return;
    }

    setIsSaving(true);
    setErrorMessage(undefined);

    try {
      if (existingNames.indexOf(targetCollectionName) < 0) {
        await service.createCollection(targetCollectionName);
      }

      await service.pinToCollection(
        targetCollectionName,
        item.url,
        item.title || item.url,
        item.properties || {}
      );

      setIsSaving(false);
      setIsOpen(false);
      setAvailableCollections([]);
      setSelectedCollectionKey(undefined);
      setNewCollectionName('');
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to add the result to the collection.';
      setErrorMessage(message);
      setIsSaving(false);
    }
  }

  return (
    <>
      <IconButton
        iconProps={{ iconName: 'FabricFolder' }}
        title="Add to collection"
        ariaLabel="Add to collection"
        className={buttonClassName || styles.resultInlineAction}
        onClick={handleOpenDialog}
      />
      <Dialog
        hidden={!isOpen}
        onDismiss={handleDismiss}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Add result to a collection'
        }}
        modalProps={{ isBlocking: true }}
      >
        <div className={styles.bulkToolbarDialogIntro}>
          Collections keep the results you found. Choose an existing collection or create a new one for this result.
        </div>
        {errorMessage && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            {errorMessage}
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
                placeholder="For example: Search follow-up"
                autoFocus={true}
              />
            )}
            <div className={styles.bulkToolbarDialogHint}>
              {item.title || item.url}
            </div>
          </div>
        )}
        <DialogFooter>
          <PrimaryButton
            onClick={handleConfirm}
            text={isSaving ? 'Adding...' : 'Add to Collection'}
            disabled={!canSubmit}
          />
          <DefaultButton
            onClick={handleDismiss}
            text="Cancel"
            disabled={isSaving}
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

export default AddToCollectionButton;
