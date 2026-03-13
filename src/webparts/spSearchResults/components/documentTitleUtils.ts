import { ISearchResult } from '@interfaces/index';

export type TitleDisplayMode = 'ellipsis' | 'middle' | 'wrap';

/**
 * Extensions that use WopiFrame interactivepreview (Office Online).
 */
export const OFFICE_EXTENSIONS: string[] = [
  'docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt',
  'docm', 'dotx', 'xlsm', 'xltx', 'pptm', 'potx',
];

/** All extensions that can be previewed in-browser. */
export const PREVIEWABLE_EXTENSIONS: string[] = [
  ...OFFICE_EXTENSIONS,
  'pdf',
  'png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg',
  'txt', 'csv', 'json', 'xml',
];

/**
 * Formats a file size in bytes into a human-readable string.
 */
export function formatFileSize(bytes: number): string {
  if (!bytes || bytes <= 0) {
    return '';
  }
  if (bytes < 1024) {
    return bytes + ' B';
  }
  if (bytes < 1048576) {
    return Math.round(bytes / 1024) + ' KB';
  }
  return (bytes / 1048576).toFixed(1) + ' MB';
}

/**
 * Formats a relative time string (e.g. "2 days ago", "3 hours ago").
 * Falls back to short date for older items.
 */
export function formatRelativeDate(isoDate: string): string {
  if (!isoDate) {
    return '';
  }
  try {
    const d: Date = new Date(isoDate);
    if (isNaN(d.getTime())) {
      return '';
    }
    const now: number = Date.now();
    const diff: number = now - d.getTime();
    const minutes: number = Math.floor(diff / 60000);
    const hours: number = Math.floor(diff / 3600000);
    const days: number = Math.floor(diff / 86400000);

    if (minutes < 1) {
      return 'Just now';
    }
    if (minutes < 60) {
      return minutes + (minutes === 1 ? ' min ago' : ' mins ago');
    }
    if (hours < 24) {
      return hours + (hours === 1 ? ' hour ago' : ' hours ago');
    }
    if (days < 7) {
      return days + (days === 1 ? ' day ago' : ' days ago');
    }
    if (days < 30) {
      const weeks: number = Math.floor(days / 7);
      return weeks + (weeks === 1 ? ' week ago' : ' weeks ago');
    }
    const month: number = d.getMonth() + 1;
    const day: number = d.getDate();
    const year: number = d.getFullYear();
    return month + '/' + day + '/' + year;
  } catch {
    return '';
  }
}

/**
 * Formats an ISO date string into "MMM D, YYYY at H:MM AM/PM" format.
 */
export function formatDateTime(isoDate: string): string {
  if (!isoDate) {
    return '';
  }
  try {
    const d: Date = new Date(isoDate);
    if (isNaN(d.getTime())) {
      return '';
    }
    const dateStr: string = d.toLocaleDateString(undefined, {
      month: 'short',
      day: 'numeric',
      year: 'numeric',
    });
    const timeStr: string = d.toLocaleTimeString([], {
      hour: 'numeric',
      minute: '2-digit',
    });
    return dateStr + ' at ' + timeStr;
  } catch {
    return '';
  }
}

/**
 * Formats an ISO date string into a short date format (M/D/YYYY).
 */
export function formatShortDate(isoDate: string): string {
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
    return month + '/' + day + '/' + year;
  } catch {
    return '';
  }
}

/**
 * Extracts a breadcrumb-style URL display from a full URL.
 * e.g. "https://contoso.sharepoint.com/sites/hr/docs/guide.pdf" => "contoso.sharepoint.com > sites > hr > docs"
 */
export function formatUrlBreadcrumb(url: string): string {
  try {
    let cleaned: string = url.replace(/^https?:\/\//, '');
    const qIdx: number = cleaned.indexOf('?');
    if (qIdx >= 0) {
      cleaned = cleaned.substring(0, qIdx);
    }
    const hIdx: number = cleaned.indexOf('#');
    if (hIdx >= 0) {
      cleaned = cleaned.substring(0, hIdx);
    }
    const segments: string[] = cleaned.split('/');
    if (segments.length > 1) {
      segments.pop();
    }
    return segments.join(' \u203A ');
  } catch {
    return url;
  }
}

/**
 * Sanitizes HitHighlightedSummary HTML from SharePoint Search API.
 * Strips all tags except safe formatting tags used for hit highlighting.
 */
export function sanitizeSummaryHtml(html: string): string {
  return html.replace(/<\/?(?!(?:b|strong|em|i|mark|c0|ddd)\b)[^>]*>/gi, '');
}

/**
 * Strips all HTML tags from a string.
 */
export function stripHtml(html: string): string {
  if (!html) {
    return '';
  }
  return html.replace(/<[^>]*>/g, '');
}

/**
 * Formats a title according to the admin-configured display mode.
 * "middle" preserves the tail/extension for long filenames.
 */
export function formatTitleText(title: string, mode: TitleDisplayMode): string {
  if (!title || mode !== 'middle') {
    return title;
  }

  if (title.length <= 40) {
    return title;
  }

  const dotIndex = title.lastIndexOf('.');
  if (dotIndex > 0 && dotIndex < title.length - 1) {
    const extension = title.substring(dotIndex);
    const base = title.substring(0, dotIndex);
    if (base.length <= 28) {
      return title;
    }
    return base.substring(0, 22) + '...' + base.substring(Math.max(base.length - 10, 22)) + extension;
  }

  return title.substring(0, 24) + '...' + title.substring(Math.max(title.length - 12, 24));
}

/**
 * Extracts the file extension from a URL, ignoring query strings.
 * Returns empty string for URLs without an extension.
 */
function getExtFromUrl(url: string): string {
  const clean = url.split('?')[0].split('#')[0];
  const filename = clean.substring(clean.lastIndexOf('/') + 1);
  const dot = filename.lastIndexOf('.');
  return dot >= 0 ? filename.substring(dot + 1).toLowerCase() : '';
}

function getExtFromName(name: string): string {
  if (!name) {
    return '';
  }
  const clean = name.split('?')[0].split('#')[0];
  const dot = clean.lastIndexOf('.');
  return dot >= 0 ? clean.substring(dot + 1).toLowerCase() : '';
}

function getFileExtension(item: ISearchResult): string {
  const managedPropertyExt = String(
    item.fileType ||
    item.properties.FileExtension ||
    item.properties.SecondaryFileExtension ||
    item.properties.FileLeafRef ||
    item.properties.FileName ||
    item.properties.Filename ||
    ''
  ).toLowerCase();

  return managedPropertyExt || getExtFromUrl(item.url || '') || getExtFromName(item.title || '');
}

/**
 * Builds the preview iframe URL for a search result.
 * PDFs and Office docs use WopiFrame.aspx (embeds cleanly in iframes).
 * Images/text use ?web=1.
 * Returns undefined for non-previewable file types.
 */
export function buildPreviewUrl(item: ISearchResult): string | undefined {
  if (!item.url) {
    return undefined;
  }
  const ext: string = getFileExtension(item);
  if (!ext || PREVIEWABLE_EXTENSIONS.indexOf(ext) < 0) {
    return undefined;
  }

  if (ext === 'pdf') {
    // PDFs: embed the file directly. SharePoint serves PDFs with
    // Content-Disposition: inline and X-Frame-Options: SAMEORIGIN so they
    // render via the browser's native PDF viewer inside the iframe.
    // WopiFrame (both embedview and interactivepreview) navigates the top
    // frame for PDFs, breaking out of any modal/panel.
    return item.url;
  }

  if (OFFICE_EXTENSIONS.indexOf(ext) >= 0) {
    // Office docs: WopiFrame interactivepreview (Office Online editor in iframe).
    const siteUrl = item.siteUrl;
    if (siteUrl) {
      return siteUrl + '/_layouts/15/WopiFrame.aspx?sourcedoc=' +
        encodeURIComponent(item.url) + '&action=interactivepreview';
    }
    // Fallback: derive the site URL by parsing item.url
    try {
      const parsed = new URL(item.url);
      const segs: string[] = parsed.pathname.split('/').filter(Boolean);
      const first: string = (segs[0] || '').toLowerCase();
      const sitePath: string =
        (first === 'sites' || first === 'teams' || first === 'personal') && segs.length >= 2
          ? '/' + segs[0] + '/' + segs[1]
          : '';
      return parsed.origin + sitePath + '/_layouts/15/WopiFrame.aspx?sourcedoc=' +
        encodeURIComponent(parsed.pathname) + '&action=interactivepreview';
    } catch {
      return undefined;
    }
  }

  const separator: string = item.url.indexOf('?') >= 0 ? '&' : '?';
  return item.url + separator + 'web=1';
}

export function getResultAnchorProps(item: ISearchResult): {
  href: string;
  target?: string;
  rel?: string;
} {
  if (buildPreviewUrl(item)) {
    return { href: '#' };
  }

  return {
    href: item.url || '#',
    target: '_blank',
    rel: 'noopener noreferrer'
  };
}

/**
 * Builds a browser-open URL that mirrors normal SharePoint library behavior
 * more closely than a direct binary link.
 */
export function buildBrowserOpenUrl(item: ISearchResult): string {
  if (!item.url) {
    return '#';
  }

  const ext: string = getFileExtension(item);
  if (ext === 'pdf') {
    return item.url;
  }

  if (PREVIEWABLE_EXTENSIONS.indexOf(ext) >= 0 || OFFICE_EXTENSIONS.indexOf(ext) >= 0) {
    const separator: string = item.url.indexOf('?') >= 0 ? '&' : '?';
    return item.url + separator + 'web=1';
  }

  return item.url;
}

/**
 * Builds a SharePoint OOB list form URL.
 * PageType 4 = DispForm (view), PageType 6 = EditForm (edit).
 */
export function buildFormUrl(item: ISearchResult, pageType: number): string | undefined {
  const siteUrl = item.properties.SPSiteURL as string | undefined;
  const listId = item.properties.ListId as string | undefined;
  const itemId = item.properties.ListItemID as string | undefined;
  if (!siteUrl || !listId || !itemId) {
    return undefined;
  }
  return siteUrl + '/_layouts/15/listform.aspx?PageType=' + pageType +
    '&ListId=' + encodeURIComponent(listId) + '&ID=' + encodeURIComponent(itemId);
}
