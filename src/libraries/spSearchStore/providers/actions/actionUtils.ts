import { ISearchResult } from '@interfaces/index';

export function normalizeUrl(rawUrl: string): string {
  if (!rawUrl) {
    return '';
  }
  const url: string = rawUrl.trim();
  if (url.indexOf('http://') === 0 || url.indexOf('https://') === 0) {
    return url;
  }
  if (url.indexOf('//') === 0) {
    return window.location.protocol + url;
  }
  if (url.charAt(0) === '/') {
    return window.location.origin + url;
  }
  return url;
}

export function buildDownloadUrl(rawUrl: string): string {
  const url: string = normalizeUrl(rawUrl);
  if (!url) {
    return '';
  }
  return url.indexOf('?') >= 0 ? url + '&download=1' : url + '?download=1';
}

export function copyTextToClipboard(text: string): Promise<void> {
  if (!text) {
    return Promise.resolve();
  }

  if (navigator.clipboard && navigator.clipboard.writeText) {
    return navigator.clipboard.writeText(text);
  }

  return new Promise(function (resolve, reject): void {
    try {
      const textarea: HTMLTextAreaElement = document.createElement('textarea');
      textarea.value = text;
      textarea.style.position = 'fixed';
      textarea.style.left = '-9999px';
      document.body.appendChild(textarea);
      textarea.select();
      const success = document.execCommand('copy');
      document.body.removeChild(textarea);
      if (success) {
        resolve();
      } else {
        reject(new Error('Copy failed'));
      }
    } catch (error) {
      reject(error as Error);
    }
  });
}

export function buildShareLines(items: ISearchResult[]): string[] {
  return items.map(function (item): string {
    const url = normalizeUrl(item.url);
    if (item.title) {
      return item.title + ' - ' + url;
    }
    return url;
  });
}
