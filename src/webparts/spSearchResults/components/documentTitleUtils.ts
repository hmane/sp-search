import { ISearchResult } from '@interfaces/index';

/** Office extensions that use WopiFrame for preview. */
export const OFFICE_EXTENSIONS: string[] = [
  'docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt',
  'docm', 'dotx', 'xlsm', 'xltx', 'pptm', 'potx',
];

/** All extensions that can be previewed in-browser. */
export const PREVIEWABLE_EXTENSIONS: string[] = [
  ...OFFICE_EXTENSIONS,
  'pdf', 'png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg',
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
 * Builds the preview iframe URL for a search result.
 * Office docs use WopiFrame.aspx; others use ?web=1.
 * Returns undefined for non-previewable file types.
 */
export function buildPreviewUrl(item: ISearchResult): string | undefined {
  if (!item.fileType || !item.url) {
    return undefined;
  }
  const ext: string = item.fileType.toLowerCase();
  if (PREVIEWABLE_EXTENSIONS.indexOf(ext) < 0) {
    return undefined;
  }

  if (OFFICE_EXTENSIONS.indexOf(ext) >= 0) {
    try {
      const parsed = new URL(item.url);
      const pathSegments: string[] = parsed.pathname.split('/').filter(Boolean);
      let sitePath: string = '';
      const first: string = (pathSegments[0] || '').toLowerCase();
      if ((first === 'sites' || first === 'teams' || first === 'personal') && pathSegments.length >= 2) {
        sitePath = '/' + pathSegments[0] + '/' + pathSegments[1];
      }
      return parsed.origin + sitePath + '/_layouts/15/WopiFrame.aspx?sourcedoc=' + encodeURIComponent(parsed.pathname) + '&action=interactivepreview';
    } catch {
      return undefined;
    }
  }

  const separator: string = item.url.indexOf('?') >= 0 ? '&' : '?';
  return item.url + separator + 'web=1';
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
