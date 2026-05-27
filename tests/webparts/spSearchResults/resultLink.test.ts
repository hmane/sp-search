import type { ISearchResult } from '../../../src/libraries/spSearchStore/interfaces/ISearchResult';
import {
  classifyResult,
  resolveResultLink,
  type ResultKind,
  type IResultLinkConfig,
} from '../../../src/webparts/spSearchResults/components/resultLink';

/**
 * Tests for the result-link resolution util (Stream C / #7).
 * Design: docs/superpowers/specs/2026-05-12-result-link-behaviour-design.md.
 *
 * Two pure functions:
 * - `classifyResult(item)` → 'document' | 'page' | 'listItem' | 'site' | 'folder' | 'other'
 *   from `item.properties` (contentclass, FileExtension, IsDocument).
 * - `resolveResultLink(item, cfg)` → { href, target?, rel?, openInPanel } —
 *   dispatches by classification + the three click-behaviour config values.
 */

function makeItem(props: Partial<ISearchResult> & { properties?: Record<string, unknown> } = {}): ISearchResult {
  return {
    key: 'k',
    title: 't',
    url: props.url ?? 'https://contoso.sharepoint.com/sites/x/Doc.docx',
    summary: '',
    author: { displayText: '', email: '' },
    created: '',
    modified: '',
    fileType: props.fileType ?? '',
    fileSize: 0,
    siteName: '',
    siteUrl: props.siteUrl ?? 'https://contoso.sharepoint.com/sites/x',
    thumbnailUrl: '',
    properties: props.properties ?? {},
    ...props,
  } as ISearchResult;
}

const baseCfg: IResultLinkConfig = {
  clickTarget: 'newTab',
  documentLinkMode: 'file',
  listItemLinkMode: 'displayForm',
};

describe('classifyResult', () => {
  it.each<[string, Partial<ISearchResult> & { properties?: Record<string, unknown> }, ResultKind]>([
    // Pages take precedence over docs (a .aspx in a doc library is a page).
    ['.aspx file → page', { fileType: 'aspx', properties: { contentclass: 'STS_ListItem_DocumentLibrary' } }, 'page'],
    ['ContentClass STS_ListItem_WebPageLibrary → page', { properties: { contentclass: 'STS_ListItem_WebPageLibrary' } }, 'page'],

    // Documents — non-aspx file in a doc library, or IsDocument=true.
    ['.docx in doc library → document', { fileType: 'docx', properties: { contentclass: 'STS_ListItem_DocumentLibrary' } }, 'document'],
    ['IsDocument=true → document', { properties: { IsDocument: 'true', FileExtension: 'pdf' } }, 'document'],
    ['ContentClass STS_Document → document', { properties: { contentclass: 'STS_Document' } }, 'document'],

    // List items — STS_ListItem_* but not a doc library or page library.
    ['STS_ListItem_GenericList → listItem', { properties: { contentclass: 'STS_ListItem_GenericList' } }, 'listItem'],
    ['STS_ListItem (bare) → listItem', { properties: { contentclass: 'STS_ListItem' } }, 'listItem'],

    // Sites / folders / other.
    ['STS_Site → site', { properties: { contentclass: 'STS_Site' } }, 'site'],
    ['STS_Web → site', { properties: { contentclass: 'STS_Web' } }, 'site'],
    ['STS_List_DocumentLibrary → folder', { properties: { contentclass: 'STS_List_DocumentLibrary' } }, 'folder'],
    ['empty / unknown → other', {}, 'other'],
    ['ExternalContentItem → other', { properties: { contentclass: 'ExternalContentItem' } }, 'other'],
  ])('%s', (_label, partial, expected) => {
    expect(classifyResult(makeItem(partial))).toBe(expected);
  });
});

describe('resolveResultLink', () => {
  // Common test items.
  const docItem = makeItem({
    url: 'https://contoso.sharepoint.com/sites/x/Shared%20Documents/Spec.docx',
    fileType: 'docx',
    properties: {
      contentclass: 'STS_ListItem_DocumentLibrary',
      SPSiteURL: 'https://contoso.sharepoint.com/sites/x',
      ListId: '11111111-1111-1111-1111-111111111111',
      ListItemID: '42',
    },
  });
  const listItem = makeItem({
    url: 'https://contoso.sharepoint.com/sites/x/Lists/Items/DispForm.aspx?ID=7',
    properties: {
      contentclass: 'STS_ListItem_GenericList',
      SPSiteURL: 'https://contoso.sharepoint.com/sites/x',
      ListId: '22222222-2222-2222-2222-222222222222',
      ListItemID: '7',
    },
  });
  const pageItem = makeItem({
    url: 'https://contoso.sharepoint.com/sites/x/SitePages/Home.aspx',
    fileType: 'aspx',
    properties: { contentclass: 'STS_ListItem_WebPageLibrary' },
  });

  describe('clickTarget = newTab (default)', () => {
    it('document + file mode → buildBrowserOpenUrl URL, target=_blank, openInPanel=false', () => {
      const r = resolveResultLink(docItem, { ...baseCfg, clickTarget: 'newTab', documentLinkMode: 'file' });
      expect(r.href).toContain('Spec.docx');
      expect(r.target).toBe('_blank');
      expect(r.rel).toBe('noopener noreferrer');
      expect(r.openInPanel).toBe(false);
    });

    it('document + propertiesForm mode → listform.aspx?PageType=4', () => {
      const r = resolveResultLink(docItem, { ...baseCfg, clickTarget: 'newTab', documentLinkMode: 'propertiesForm' });
      expect(r.href).toContain('/_layouts/15/listform.aspx');
      expect(r.href).toContain('PageType=4');
      expect(r.href).toContain('ListId=');
      expect(r.href).toContain('ID=42');
      expect(r.target).toBe('_blank');
      expect(r.openInPanel).toBe(false);
    });

    it('list item + displayForm mode → item.url (which IS the DispForm)', () => {
      const r = resolveResultLink(listItem, { ...baseCfg, clickTarget: 'newTab', listItemLinkMode: 'displayForm' });
      expect(r.href).toBe(listItem.url);
      expect(r.target).toBe('_blank');
      expect(r.openInPanel).toBe(false);
    });

    it('list item + editForm mode → listform.aspx?PageType=6', () => {
      const r = resolveResultLink(listItem, { ...baseCfg, clickTarget: 'newTab', listItemLinkMode: 'editForm' });
      expect(r.href).toContain('/_layouts/15/listform.aspx');
      expect(r.href).toContain('PageType=6');
      expect(r.href).toContain('ID=7');
      expect(r.openInPanel).toBe(false);
    });

    it('page → item.url, openInPanel=false', () => {
      const r = resolveResultLink(pageItem, { ...baseCfg, clickTarget: 'newTab' });
      expect(r.href).toBe(pageItem.url);
      expect(r.openInPanel).toBe(false);
    });

    it('missing ListId on editForm → falls back to item.url (safe)', () => {
      const itemNoIds = makeItem({
        url: 'https://contoso.sharepoint.com/sites/x/Lists/Items/DispForm.aspx?ID=99',
        properties: { contentclass: 'STS_ListItem_GenericList' /* no ListId / ListItemID */ },
      });
      const r = resolveResultLink(itemNoIds, { ...baseCfg, clickTarget: 'newTab', listItemLinkMode: 'editForm' });
      expect(r.href).toBe(itemNoIds.url);
      expect(r.openInPanel).toBe(false);
    });
  });

  describe('clickTarget = sameTab', () => {
    it('document → no target/rel set; same-tab navigation', () => {
      const r = resolveResultLink(docItem, { ...baseCfg, clickTarget: 'sameTab' });
      expect(r.target).toBeUndefined();
      expect(r.rel).toBeUndefined();
      expect(r.openInPanel).toBe(false);
    });
  });

  describe('clickTarget = panel', () => {
    it('document → resolved URL set, no _blank target (Modal handles previewables via DocumentTitleHoverCard; non-previewable navigates)', () => {
      const r = resolveResultLink(docItem, { ...baseCfg, clickTarget: 'panel' });
      expect(r.href).toContain('Spec.docx');
      // panel mode = today's behaviour: previewable files open the Modal (handled by DocumentTitleHoverCard.handleClick),
      // non-previewable files navigate. The <a> has a href but no _blank (the Modal preventDefaults previewables).
      expect(r.openInPanel).toBe(false);
    });
  });

  describe('clickTarget = sidePanel', () => {
    it('any kind → openInPanel=true (ResultDetailPanel)', () => {
      expect(resolveResultLink(docItem, { ...baseCfg, clickTarget: 'sidePanel' }).openInPanel).toBe(true);
      expect(resolveResultLink(listItem, { ...baseCfg, clickTarget: 'sidePanel' }).openInPanel).toBe(true);
      expect(resolveResultLink(pageItem, { ...baseCfg, clickTarget: 'sidePanel' }).openInPanel).toBe(true);
    });

    it('href is still the resolved URL (so middle-click / Ctrl-click open it)', () => {
      const r = resolveResultLink(listItem, { ...baseCfg, clickTarget: 'sidePanel', listItemLinkMode: 'editForm' });
      expect(r.href).toContain('PageType=6');
      expect(r.openInPanel).toBe(true);
    });
  });
});
