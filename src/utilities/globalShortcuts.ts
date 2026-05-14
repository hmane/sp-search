/**
 * T2.D9 — global keyboard shortcuts.
 *
 * Each web part registers shortcut handlers via the `useGlobalShortcuts`
 * React hook. The hook attaches a single `document.keydown` listener
 * per mount and dispatches to the registered handlers when the key
 * matches AND `shouldHandleShortcut(event.target)` returns true.
 *
 * Cross-bundle coordination: shortcuts that affect another web part
 * (e.g. `/` focuses the search box) use a window-dispatched
 * `CustomEvent('sp-search:focus-search-box')` so the target web part
 * subscribes regardless of which bundle the keydown fired from.
 */

import * as React from 'react';

const TEXT_INPUT_TYPES = new Set<string>([
  'text', 'search', 'email', 'url', 'number', 'password', 'tel',
]);

/**
 * Returns true when the event target is NOT inside a text-entry surface
 * (HTMLInputElement of a text-y type, HTMLTextAreaElement, role="textbox",
 * or contentEditable). Used to gate global shortcuts so typing in the
 * search box doesn't fire `/` again.
 */
export function shouldHandleShortcut(target: EventTarget | null): boolean {
  if (!target) { return true; }
  const el = target as HTMLElement;

  if (el.tagName === 'TEXTAREA') { return false; }

  if (el.tagName === 'INPUT') {
    const inputEl = el as HTMLInputElement;
    const type = (inputEl.type || 'text').toLowerCase();
    return !TEXT_INPUT_TYPES.has(type);
  }

  if (el.getAttribute && el.getAttribute('role') === 'textbox') { return false; }

  if (el.isContentEditable) { return false; }

  return true;
}

// Custom event names. Window-dispatched so any web part bundle can listen
// without needing a shared module-level singleton.
export const SP_SEARCH_FOCUS_SEARCH_BOX = 'sp-search:focus-search-box';
export const SP_SEARCH_OPEN_SHORTCUT_HELP = 'sp-search:open-shortcut-help';

export interface IShortcutBinding {
  /** Key value from KeyboardEvent.key (e.g. '/', '?', 'Escape'). */
  key: string;
  /** Optional shift/ctrl/alt/meta requirement. Default false for each. */
  shift?: boolean;
  ctrl?: boolean;
  alt?: boolean;
  meta?: boolean;
  /** Action to run when the key matches AND `shouldHandleShortcut` passes. */
  handler: (event: KeyboardEvent) => void;
}

/**
 * Install global keyboard shortcuts on `document.keydown`. The hook
 * automatically excludes typing inside text inputs / textareas /
 * contentEditable elements via `shouldHandleShortcut`. Bindings array
 * is captured at install time; pass a stable reference (useMemo) to
 * avoid re-installing on every render.
 */
export function useGlobalShortcuts(bindings: IShortcutBinding[]): void {
  React.useEffect((): (() => void) => {
    const handleKeyDown = (event: KeyboardEvent): void => {
      if (!shouldHandleShortcut(event.target)) { return; }
      for (let i = 0; i < bindings.length; i++) {
        const b = bindings[i];
        if (event.key !== b.key) { continue; }
        if ((b.shift || false) !== event.shiftKey) { continue; }
        if ((b.ctrl || false) !== event.ctrlKey) { continue; }
        if ((b.alt || false) !== event.altKey) { continue; }
        if ((b.meta || false) !== event.metaKey) { continue; }
        b.handler(event);
        return;
      }
    };
    document.addEventListener('keydown', handleKeyDown);
    return (): void => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [bindings]);
}

/**
 * Dispatch the "focus the search box" event. Any mounted SP Search Box
 * web part listens for this and calls `.focus()` on its input.
 */
export function dispatchFocusSearchBox(): void {
  if (typeof window === 'undefined') { return; }
  window.dispatchEvent(new CustomEvent(SP_SEARCH_FOCUS_SEARCH_BOX));
}

/**
 * Dispatch the "open shortcut help" event. The Manager (or Box) web
 * part hosts the modal; whichever is mounted handles it.
 */
export function dispatchOpenShortcutHelp(): void {
  if (typeof window === 'undefined') { return; }
  window.dispatchEvent(new CustomEvent(SP_SEARCH_OPEN_SHORTCUT_HELP));
}
