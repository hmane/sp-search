/**
 * Shared ContextualMenu items builder for per-row ECB ("...") menus.
 *
 * Powers the inline action menu on List + DataGrid layouts so the same
 * Open / Download / Copy link surface is exposed from one place. The
 * DataGrid surface adds View / Edit / Delete via item-level permissions
 * — those branches inline directly in DataGridContent because they
 * depend on the `permissionCache` + `recycleSearchResult` helpers
 * that live there. List's menu is the smaller common-case subset.
 */

import type { IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import type { ISearchResult } from '@interfaces/index';
import { buildDownloadUrl, copyTextToClipboard } from '@providers/actions/actionUtils';
import { buildBrowserOpenUrl } from './documentTitleUtils';

export interface IBuildRowActionMenuOptions {
  /** Row position (1-based) reported to onItemClick for analytics. */
  position?: number;
  /** Telemetry / preview-panel hook fired on Open. */
  onItemClick?: (item: ISearchResult, position: number) => void;
  /** Optional toast trigger fired after Copy link succeeds. */
  onCopyLinkSuccess?: (url: string) => void;
  /** Opens the row's Add to collection dialog. */
  onAddToCollection?: (event?: { preventDefault?: () => void; stopPropagation?: () => void }) => void;
}

export function buildAddToCollectionMenuItem(
  onAddToCollection: (event?: { preventDefault?: () => void; stopPropagation?: () => void }) => void
): IContextualMenuItem {
  return {
    key: 'addToCollection',
    text: 'Add to collection',
    iconProps: { iconName: 'FabricFolder' },
    onClick: (event?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>): void => {
      onAddToCollection(event);
    },
  };
}

/**
 * Build the canonical Open / Download / Copy-link menu items for a row.
 * Returns an array consumable as the `items` prop of `IContextualMenu`.
 */
export function buildRowActionMenu(
  item: ISearchResult,
  options: IBuildRowActionMenuOptions = {}
): IContextualMenuItem[] {
  const position: number = options.position ?? 1;
  const onItemClick = options.onItemClick;
  const onCopyLinkSuccess = options.onCopyLinkSuccess;
  const menuItems: IContextualMenuItem[] = [
    {
      key: 'open',
      text: 'Open in new tab',
      iconProps: { iconName: 'OpenInNewTab' },
      onClick: (): void => {
        if (onItemClick) {
          onItemClick(item, position);
        }
        window.open(buildBrowserOpenUrl(item), '_blank', 'noopener,noreferrer');
      },
    },
  ];

  if (options.onAddToCollection) {
    menuItems.push(buildAddToCollectionMenuItem(options.onAddToCollection));
  }

  menuItems.push(
    {
      key: 'download',
      text: 'Download',
      iconProps: { iconName: 'Download' },
      onClick: (): void => {
        const downloadUrl = buildDownloadUrl(item.url);
        if (downloadUrl) {
          window.open(downloadUrl, '_blank', 'noopener,noreferrer');
        }
      },
    },
    {
      key: 'copyLink',
      text: 'Copy link',
      iconProps: { iconName: 'Link' },
      onClick: (): void => {
        copyTextToClipboard(item.url)
          .then((): void => {
            if (onCopyLinkSuccess) {
              onCopyLinkSuccess(item.url);
            }
          })
          .catch((): void => { /* silent — caller decides UX */ });
      },
    },
  );

  return menuItems;
}
