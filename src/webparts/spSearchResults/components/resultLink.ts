import type { ISearchResult } from '@interfaces/index';
import { buildBrowserOpenUrl, buildFormUrl } from './documentTitleUtils';

/**
 * Result link behaviour util (Stream C / #7).
 * Design: docs/superpowers/specs/2026-05-12-result-link-behaviour-design.md.
 *
 * Two pure functions consumed by all five result layouts and by
 * `DocumentTitleHoverCard.handleClick`:
 *
 * - `classifyResult(item)` — content-type discriminator from
 *   `item.properties.contentclass` (+ `IsDocument` + `FileExtension` fallbacks).
 * - `resolveResultLink(item, cfg)` — `{ href, target?, rel?, openInPanel }`
 *   computed from the classification + the three new property-pane settings.
 *   Falls back safely to `item.url` whenever a form URL can't be built
 *   (missing `ListId` / `ListItemID` — e.g. external-connector items).
 */

export type ResultKind = 'document' | 'page' | 'listItem' | 'site' | 'folder' | 'other';

export type ResultClickTarget = 'panel' | 'newTab' | 'sameTab' | 'sidePanel';
export type DocumentLinkMode = 'file' | 'propertiesForm';
export type ListItemLinkMode = 'displayForm' | 'editForm';

export interface IResultLinkConfig {
  clickTarget: ResultClickTarget;
  documentLinkMode: DocumentLinkMode;
  listItemLinkMode: ListItemLinkMode;
}

export interface IResolvedResultLink {
  href: string;
  target?: string;
  rel?: string;
  /**
   * True iff the title-link click should open `ResultDetailPanel` via
   * `store.getState().setPreviewItem(item)` instead of navigating. Set only
   * when `clickTarget === 'sidePanel'`. (The `panel` mode preserves today's
   * `DocumentTitleHoverCard` Modal-for-previewables behaviour and does NOT
   * set this — that interception lives in `DocumentTitleHoverCard.handleClick`.)
   */
  openInPanel: boolean;
}

/**
 * Classify a result by content type. Order-dependent: a `.aspx` file in a
 * document library is classified as `page`, not `document`.
 */
export function classifyResult(item: ISearchResult): ResultKind {
  const props = (item.properties || {}) as Record<string, unknown>;
  const cc: string = String(props.contentclass || '').trim();
  const ext: string = String(
    item.fileType || props.FileExtension || props.SecondaryFileExtension || ''
  ).toLowerCase();
  const isDoc: boolean = String(props.IsDocument || '').toLowerCase() === 'true';

  // Pages first — a .aspx in a doc library is a page.
  if (ext === 'aspx' || cc === 'STS_ListItem_WebPageLibrary' || cc === 'STS_ListItem_851') {
    return 'page';
  }
  if (isDoc || cc.indexOf('STS_ListItem_DocumentLibrary') === 0 || cc === 'STS_Document') {
    return 'document';
  }
  if (cc === 'STS_Site' || cc === 'STS_Web') {
    return 'site';
  }
  // STS_List / STS_List_* (without ListItem) is a container, not an item.
  if (cc === 'STS_List' || cc.indexOf('STS_List_') === 0) {
    return 'folder';
  }
  if (cc.indexOf('STS_ListItem') === 0) {
    return 'listItem';
  }
  return 'other';
}

/**
 * Resolve a result's click destination + anchor attributes from its
 * classification and the admin config.
 *
 * Defaults (`clickTarget='panel'`, `documentLinkMode='file'`,
 * `listItemLinkMode='displayForm'`) reproduce today's behaviour byte-for-byte
 * at the `<a>` level (target=_blank, href=document's browser-open URL or
 * the item URL). The Modal-for-previewables happens via
 * `DocumentTitleHoverCard.handleClick`, not here.
 */
export function resolveResultLink(item: ISearchResult, cfg: IResultLinkConfig): IResolvedResultLink {
  const kind: ResultKind = classifyResult(item);
  const openInPanel: boolean = cfg.clickTarget === 'sidePanel';

  // ── href: resolved by classification + per-type mode, with safe fallback ──
  let href: string;
  const fallback: string = item.url || '#';
  if (kind === 'document') {
    if (cfg.documentLinkMode === 'propertiesForm') {
      href = buildFormUrl(item, 4) || fallback; // PageType 4 = DispForm
    } else {
      href = buildBrowserOpenUrl(item); // 'file' (default) — Office Online via ?web=1, PDF inline, else item.url
    }
  } else if (kind === 'listItem') {
    if (cfg.listItemLinkMode === 'editForm') {
      href = buildFormUrl(item, 6) || fallback; // PageType 6 = EditForm
    } else {
      href = fallback; // displayForm — search returns item.url as DispForm already
    }
  } else {
    // page / site / folder / other → just navigate to the item URL.
    href = fallback;
  }

  // ── target / rel from clickTarget ──
  let target: string | undefined;
  let rel: string | undefined;
  if (cfg.clickTarget === 'newTab' || cfg.clickTarget === 'panel') {
    // newTab → always new-tab navigation.
    // panel → preserves today's behaviour: new-tab for non-previewables; the
    // DocumentTitleHoverCard Modal preventDefaults previewables.
    target = '_blank';
    rel = 'noopener noreferrer';
  }
  // sameTab / sidePanel → no target (sameTab navigates in-place; sidePanel
  // intercepts the click via `openInPanel`).

  return { href, target, rel, openInPanel };
}
