import * as React from 'react';
import { StoreApi } from 'zustand/vanilla';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import {
  ISearchCollection,
  ICollectionItem,
  ISearchStore
} from '@interfaces/index';
import { SearchManagerService } from '@services/index';
import TagBadges from './TagBadges';
import ResultAnnotations from './ResultAnnotations';
import styles from './SpSearchManager.module.scss';

export interface ISearchCollectionsProps {
  store: StoreApi<ISearchStore>;
  service: SearchManagerService;
  collections: ISearchCollection[];
  onDataChanged: () => void;
}

/**
 * SearchCollections -- displays collections/pinboards.
 * Each collection is expandable to show its pinned items with tag annotations.
 * Supports creating/deleting collections, removing pinned items,
 * editing per-item tags, and filtering collections by tag.
 */
const SearchCollections: React.FC<ISearchCollectionsProps> = (props) => {
  const { service, collections, onDataChanged } = props;

  // ─── Local state ──────────────────────────────────────────
  const [expandedCollections, setExpandedCollections] = React.useState<Record<string, boolean>>({});
  const [showCreateDialog, setShowCreateDialog] = React.useState<boolean>(false);
  const [newCollectionName, setNewCollectionName] = React.useState<string>('');
  const [isCreating, setIsCreating] = React.useState<boolean>(false);
  const [deleteTarget, setDeleteTarget] = React.useState<ISearchCollection | undefined>(undefined);
  const [isDeleting, setIsDeleting] = React.useState<boolean>(false);
  const [filterTag, setFilterTag] = React.useState<string | undefined>(undefined);

  // ─── Computed values ────────────────────────────────────────

  // Collect all unique tags across all collections for the filter dropdown
  const allTags: string[] = React.useMemo(function (): string[] {
    const tagSet: Set<string> = new Set();
    for (let i = 0; i < collections.length; i++) {
      for (let j = 0; j < collections[i].tags.length; j++) {
        tagSet.add(collections[i].tags[j]);
      }
    }
    const result: string[] = Array.from(tagSet);
    result.sort();
    return result;
  }, [collections]);

  // Filter collections by selected tag
  const filteredCollections: ISearchCollection[] = React.useMemo(function (): ISearchCollection[] {
    if (!filterTag) {
      return collections;
    }
    const filtered: ISearchCollection[] = [];
    for (let i = 0; i < collections.length; i++) {
      if (collections[i].tags.indexOf(filterTag) >= 0) {
        filtered.push(collections[i]);
      }
    }
    return filtered;
  }, [collections, filterTag]);

  // Build filter dropdown options
  const filterOptions: IDropdownOption[] = React.useMemo(function (): IDropdownOption[] {
    const opts: IDropdownOption[] = [{ key: '__all__', text: 'All collections' }];
    for (let i = 0; i < allTags.length; i++) {
      opts.push({ key: allTags[i], text: allTags[i] });
    }
    return opts;
  }, [allTags]);

  // ─── Handlers ─────────────────────────────────────────────

  function handleToggleCollection(collectionName: string): void {
    setExpandedCollections(function (prev): Record<string, boolean> {
      return {
        ...prev,
        [collectionName]: !prev[collectionName]
      };
    });
  }

  function handleCreateClick(): void {
    setShowCreateDialog(true);
    setNewCollectionName('');
  }

  function handleCreateCancel(): void {
    setShowCreateDialog(false);
    setNewCollectionName('');
  }

  function handleCreateNameChange(
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void {
    setNewCollectionName(newValue !== undefined ? newValue : '');
  }

  function handleCreateConfirm(): void {
    if (!newCollectionName.trim()) {
      return;
    }

    setIsCreating(true);
    service.createCollection(newCollectionName.trim())
      .then(function (): void {
        setShowCreateDialog(false);
        setNewCollectionName('');
        setIsCreating(false);
        onDataChanged();
      })
      .catch(function (): void {
        setIsCreating(false);
      });
  }

  function handleCreateKeyDown(event: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>): void {
    if (event.key === 'Enter') {
      handleCreateConfirm();
    } else if (event.key === 'Escape') {
      handleCreateCancel();
    }
  }

  function handleDeleteClick(collection: ISearchCollection, event: React.MouseEvent): void {
    event.stopPropagation();
    setDeleteTarget(collection);
  }

  function handleDeleteConfirm(): void {
    if (!deleteTarget) {
      return;
    }

    setIsDeleting(true);
    service.deleteCollection(deleteTarget.collectionName)
      .then(function (): void {
        setDeleteTarget(undefined);
        setIsDeleting(false);
        onDataChanged();
      })
      .catch(function (): void {
        setIsDeleting(false);
      });
  }

  function handleDeleteCancel(): void {
    setDeleteTarget(undefined);
  }

  function handleUnpinItem(itemId: number, event: React.MouseEvent): void {
    event.stopPropagation();
    event.preventDefault();
    service.unpinFromCollection(itemId)
      .then(function (): void {
        onDataChanged();
      })
      .catch(function noop(): void { /* swallow */ });
  }

  function handleTagsChanged(itemId: number, tags: string[]): void {
    service.updateItemTags(itemId, tags)
      .then(function (): void {
        onDataChanged();
      })
      .catch(function noop(): void { /* swallow */ });
  }

  function handleFilterTagChange(
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void {
    if (option) {
      setFilterTag(option.key === '__all__' ? undefined : option.key as string);
    }
  }

  // ─── Empty state ──────────────────────────────────────────
  if (!collections || collections.length === 0) {
    return (
      <div>
        <div className={styles.collectionToolbar}>
          <PrimaryButton
            iconProps={{ iconName: 'Add' }}
            text="New collection"
            onClick={handleCreateClick}
          />
        </div>

        <div className={styles.emptyState}>
          <div className={styles.emptyIcon}>
            <Icon iconName="FabricFolder" />
          </div>
          <h3 className={styles.emptyTitle}>No collections</h3>
          <p className={styles.emptyDescription}>
            Create a collection to organize and pin important search results for quick access.
          </p>
        </div>

        {/* Create dialog */}
        <Dialog
          hidden={!showCreateDialog}
          onDismiss={handleCreateCancel}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Create new collection'
          }}
          modalProps={{ isBlocking: true }}
        >
          <div className={styles.dialogForm}>
            <TextField
              label="Collection name"
              value={newCollectionName}
              onChange={handleCreateNameChange}
              onKeyDown={handleCreateKeyDown}
              autoFocus={true}
              placeholder="Enter a name for the collection"
              required={true}
            />
          </div>
          <DialogFooter>
            <PrimaryButton
              onClick={handleCreateConfirm}
              text="Create"
              disabled={isCreating || !newCollectionName.trim()}
            />
            <DefaultButton
              onClick={handleCreateCancel}
              text="Cancel"
            />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  return (
    <div>
      {/* Toolbar */}
      <div className={styles.collectionToolbar}>
        {allTags.length > 0 && (
          <Dropdown
            placeholder="Filter by tag"
            options={filterOptions}
            selectedKey={filterTag || '__all__'}
            onChange={handleFilterTagChange}
            className={styles.tagFilterDropdown}
          />
        )}
        <PrimaryButton
          iconProps={{ iconName: 'Add' }}
          text="New collection"
          onClick={handleCreateClick}
        />
      </div>

      {/* Collections list */}
      <div className={styles.collectionsList}>
        {filteredCollections.map(function (collection): React.ReactElement {
          const isExpanded = expandedCollections[collection.collectionName] === true;
          // Filter out placeholder items (empty URL = collection placeholder)
          const pinnedItems: ICollectionItem[] = [];
          for (let i = 0; i < collection.items.length; i++) {
            if (collection.items[i].url) {
              pinnedItems.push(collection.items[i]);
            }
          }
          const itemCount = pinnedItems.length;

          return (
            <div key={collection.collectionName} className={styles.collectionItem}>
              {/* Collection header */}
              <div
                className={styles.collectionHeader}
                onClick={function (): void { handleToggleCollection(collection.collectionName); }}
                role="button"
                aria-expanded={isExpanded}
              >
                <Icon
                  iconName="ChevronRight"
                  className={
                    isExpanded
                      ? styles.collectionChevron + ' ' + styles.collectionChevronExpanded
                      : styles.collectionChevron
                  }
                />
                <div className={styles.collectionIcon}>
                  <Icon iconName="FabricFolder" />
                </div>
                <div className={styles.collectionBody}>
                  <p className={styles.collectionName}>{collection.collectionName}</p>
                  {/* Collection-level tag summary (collapsed view) */}
                  {!isExpanded && collection.tags.length > 0 && (
                    <TagBadges tags={collection.tags} maxVisible={3} />
                  )}
                </div>
                <span className={styles.collectionCount}>
                  {itemCount === 1 ? '1 item' : String(itemCount) + ' items'}
                </span>
                <div className={styles.collectionActions}>
                  <IconButton
                    iconProps={{ iconName: 'Delete' }}
                    title="Delete collection"
                    ariaLabel={'Delete collection ' + collection.collectionName}
                    onClick={function (e: React.MouseEvent<HTMLButtonElement>): void {
                      handleDeleteClick(collection, e as unknown as React.MouseEvent);
                    }}
                  />
                </div>
              </div>

              {/* Expanded: pinned items with tags */}
              {isExpanded && (
                <div className={styles.collectionItems}>
                  {pinnedItems.length === 0 && (
                    <div className={styles.emptyDescription}>
                      No items pinned to this collection yet.
                    </div>
                  )}
                  {pinnedItems.map(function (item, idx): React.ReactElement {
                    return (
                      <div key={item.url + '-' + String(idx)}>
                        <div className={styles.collectionPinnedItem}>
                          <div className={styles.pinnedItemIcon}>
                            <Icon iconName="Pin" />
                          </div>
                          <a
                            href={item.url}
                            target="_blank"
                            rel="noopener noreferrer"
                            className={styles.pinnedItemTitle}
                          >
                            {item.title || item.url}
                          </a>
                          <div className={styles.pinnedItemActions}>
                            <IconButton
                              iconProps={{ iconName: 'Unpin' }}
                              title="Remove from collection"
                              ariaLabel={'Remove ' + (item.title || 'item') + ' from collection'}
                              onClick={function (e: React.MouseEvent<HTMLButtonElement>): void {
                                handleUnpinItem(item.id, e as unknown as React.MouseEvent);
                              }}
                            />
                          </div>
                        </div>
                        <ResultAnnotations
                          itemId={item.id}
                          tags={item.tags}
                          onTagsChanged={handleTagsChanged}
                        />
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* Create collection dialog */}
      <Dialog
        hidden={!showCreateDialog}
        onDismiss={handleCreateCancel}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Create new collection'
        }}
        modalProps={{ isBlocking: true }}
      >
        <div className={styles.dialogForm}>
          <TextField
            label="Collection name"
            value={newCollectionName}
            onChange={handleCreateNameChange}
            onKeyDown={handleCreateKeyDown}
            autoFocus={true}
            placeholder="Enter a name for the collection"
            required={true}
          />
        </div>
        <DialogFooter>
          <PrimaryButton
            onClick={handleCreateConfirm}
            text="Create"
            disabled={isCreating || !newCollectionName.trim()}
          />
          <DefaultButton
            onClick={handleCreateCancel}
            text="Cancel"
          />
        </DialogFooter>
      </Dialog>

      {/* Delete collection confirmation */}
      <Dialog
        hidden={!deleteTarget}
        onDismiss={handleDeleteCancel}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete collection',
          subText: deleteTarget
            ? 'Are you sure you want to delete the collection "' + deleteTarget.collectionName + '" and all its pinned items? This action cannot be undone.'
            : ''
        }}
        modalProps={{ isBlocking: true }}
      >
        {isDeleting && (
          <div className={styles.loadingContainer}>
            <Spinner size={SpinnerSize.medium} label="Deleting collection..." />
          </div>
        )}
        <DialogFooter>
          <PrimaryButton
            onClick={handleDeleteConfirm}
            text="Delete"
            disabled={isDeleting}
          />
          <DefaultButton
            onClick={handleDeleteCancel}
            text="Cancel"
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default SearchCollections;
