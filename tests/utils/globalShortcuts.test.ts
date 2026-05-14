/**
 * T2.D9 — global keyboard shortcuts. Tests the predicate logic that
 * decides whether a keydown event should fire a shortcut: shortcuts
 * are suppressed when focus is in a text input, textarea, contenteditable,
 * or any element with `[role="textbox"]`.
 */

import { shouldHandleShortcut } from '../../src/utilities/globalShortcuts';

function elementWith(props: { tagName?: string; type?: string; role?: string; isContentEditable?: boolean }): HTMLElement {
  const tag = (props.tagName || 'div').toLowerCase();
  const el = document.createElement(tag);
  if (props.type) { el.setAttribute('type', props.type); }
  if (props.role) { el.setAttribute('role', props.role); }
  // jsdom doesn't preserve isContentEditable via setAttribute alone — set the property.
  if (props.isContentEditable) {
    Object.defineProperty(el, 'isContentEditable', { value: true, configurable: true });
  }
  return el;
}

describe('shouldHandleShortcut', () => {
  it('returns true for a body/div target (outside any input)', () => {
    expect(shouldHandleShortcut(document.body)).toBe(true);
    expect(shouldHandleShortcut(elementWith({ tagName: 'div' }))).toBe(true);
  });

  it('returns false when target is an INPUT', () => {
    expect(shouldHandleShortcut(elementWith({ tagName: 'input' }))).toBe(false);
  });

  it('returns false when target is a TEXTAREA', () => {
    expect(shouldHandleShortcut(elementWith({ tagName: 'textarea' }))).toBe(false);
  });

  it('returns true for input[type=checkbox] / radio / button (non-text)', () => {
    expect(shouldHandleShortcut(elementWith({ tagName: 'input', type: 'checkbox' }))).toBe(true);
    expect(shouldHandleShortcut(elementWith({ tagName: 'input', type: 'radio' }))).toBe(true);
    expect(shouldHandleShortcut(elementWith({ tagName: 'input', type: 'button' }))).toBe(true);
  });

  it('returns false for input[type=text/search/email/url/number/password]', () => {
    expect(shouldHandleShortcut(elementWith({ tagName: 'input', type: 'text' }))).toBe(false);
    expect(shouldHandleShortcut(elementWith({ tagName: 'input', type: 'search' }))).toBe(false);
    expect(shouldHandleShortcut(elementWith({ tagName: 'input', type: 'email' }))).toBe(false);
    expect(shouldHandleShortcut(elementWith({ tagName: 'input', type: 'url' }))).toBe(false);
    expect(shouldHandleShortcut(elementWith({ tagName: 'input', type: 'number' }))).toBe(false);
    expect(shouldHandleShortcut(elementWith({ tagName: 'input', type: 'password' }))).toBe(false);
  });

  it('returns false when element has role="textbox"', () => {
    expect(shouldHandleShortcut(elementWith({ tagName: 'div', role: 'textbox' }))).toBe(false);
  });

  it('returns false when element is contentEditable', () => {
    expect(shouldHandleShortcut(elementWith({ tagName: 'div', isContentEditable: true }))).toBe(false);
  });

  it('returns true for null target (no focused element)', () => {
    expect(shouldHandleShortcut(null)).toBe(true);
  });
});
