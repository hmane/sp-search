/**
 * T2.D9 — keyboard-shortcut help modal. Listens for the global
 * `sp-search:open-shortcut-help` event (dispatched by `?` from any web
 * part) and renders a Fluent Modal listing every shipped shortcut.
 *
 * Cross-bundle singleton (mirrors `DebugFabHost`). Every user-facing
 * web part mounts `<ShortcutHelpModalHost />` and the first one to
 * render claims the window-backed owner flag; non-owners short-circuit
 * to a null render and skip binding installation. On unmount the
 * owner releases the flag and a low-frequency poll lets another
 * mounted host take over. This means a page with Results+Manager
 * but no SearchBox still gets a working `?` modal and `/` focus.
 */

import * as React from 'react';
import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import {
  SP_SEARCH_OPEN_SHORTCUT_HELP,
  useGlobalShortcuts,
  dispatchOpenShortcutHelp,
  dispatchFocusSearchBox,
  type IShortcutBinding,
} from './globalShortcuts';

const OWNER_KEY = '__sp_search_shortcut_help_owner__';
const CLAIM_POLL_MS = 500;

interface IWindowWithOwner {
  [OWNER_KEY]?: string;
}

function tryClaim(instanceId: string): boolean {
  if (typeof window === 'undefined') { return false; }
  const win = window as unknown as IWindowWithOwner;
  if (!win[OWNER_KEY] || win[OWNER_KEY] === instanceId) {
    win[OWNER_KEY] = instanceId;
    return true;
  }
  return false;
}

function release(instanceId: string): void {
  if (typeof window === 'undefined') { return; }
  const win = window as unknown as IWindowWithOwner;
  if (win[OWNER_KEY] === instanceId) {
    win[OWNER_KEY] = undefined;
  }
}

interface IShortcutRow {
  keys: string[];
  description: string;
}

const SHORTCUTS: IShortcutRow[] = [
  { keys: ['/'], description: 'Focus the search box' },
  { keys: ['?'], description: 'Open this shortcut help' },
  { keys: ['Esc'], description: 'Close the open panel or dialog' },
  { keys: ['Enter'], description: 'Open the focused result' },
  { keys: ['Alt', '←'], description: 'Previous result (detail panel open)' },
  { keys: ['Alt', '→'], description: 'Next result (detail panel open)' },
  { keys: ['j'], description: 'Move focus to the next result (planned)' },
  { keys: ['k'], description: 'Move focus to the previous result (planned)' },
];

/**
 * Host component — install once per page (typically inside the
 * SearchBox web part). Renders nothing visible until the help event
 * fires.
 */
export const ShortcutHelpModalHost: React.FC = () => {
  const instanceIdRef = React.useRef<string>('shortcut-help-' + Math.random().toString(36).substring(2));
  const [isOwner, setIsOwner] = React.useState<boolean>(() => tryClaim(instanceIdRef.current));
  const [open, setOpen] = React.useState<boolean>(false);

  // Poll for ownership when we're not the owner. The current owner may
  // release the flag during a SPA navigation; this lets the next mounted
  // host take over.
  React.useEffect((): (() => void) | undefined => {
    if (isOwner) { return undefined; }
    const intervalId = window.setInterval((): void => {
      if (tryClaim(instanceIdRef.current)) {
        setIsOwner(true);
      }
    }, CLAIM_POLL_MS);
    return (): void => { window.clearInterval(intervalId); };
  }, [isOwner]);

  // Release the flag on unmount.
  React.useEffect((): (() => void) => {
    const id = instanceIdRef.current;
    return (): void => { release(id); };
  }, []);

  // Install the global shortcut bindings that drive the help modal + the
  // search-box-focus event. Non-owners pass an empty array so no duplicate
  // keydown handlers land on the document.
  const bindings = React.useMemo<IShortcutBinding[]>(() => isOwner ? [
    {
      key: '?',
      shift: true, // '?' is shift+/ on US layouts
      handler: (e: KeyboardEvent): void => {
        e.preventDefault();
        dispatchOpenShortcutHelp();
      },
    },
    {
      key: '/',
      handler: (e: KeyboardEvent): void => {
        e.preventDefault();
        dispatchFocusSearchBox();
      },
    },
  ] : [], [isOwner]);
  useGlobalShortcuts(bindings);

  // Listen for the open event — only when owner.
  React.useEffect((): (() => void) | undefined => {
    if (!isOwner) { return undefined; }
    const handler = (): void => { setOpen(true); };
    window.addEventListener(SP_SEARCH_OPEN_SHORTCUT_HELP, handler);
    return (): void => { window.removeEventListener(SP_SEARCH_OPEN_SHORTCUT_HELP, handler); };
  }, [isOwner]);

  if (!isOwner) { return null; }

  return (
    <Modal
      isOpen={open}
      onDismiss={(): void => setOpen(false)}
      isBlocking={false}
      containerClassName="sp-search-shortcut-help"
    >
      <div style={{ padding: 24, minWidth: 320, maxWidth: 480 }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
          <h2 style={{ margin: 0, fontSize: 18 }}>Keyboard shortcuts</h2>
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            ariaLabel="Close shortcut help"
            onClick={(): void => setOpen(false)}
          />
        </div>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 14 }}>
          <tbody>
            {SHORTCUTS.map((s): React.ReactElement => (
              <tr key={s.keys.join('+')}>
                <td style={{ padding: '6px 12px 6px 0', whiteSpace: 'nowrap' }}>
                  {s.keys.map((k, idx): React.ReactElement => (
                    <span
                      key={String(idx) + k}
                      style={{
                        display: 'inline-block',
                        padding: '2px 8px',
                        marginRight: 4,
                        border: '1px solid var(--neutralTertiaryAlt, #c8c6c4)',
                        borderRadius: 3,
                        backgroundColor: 'var(--neutralLighter, #faf9f8)',
                        fontFamily: 'Consolas, Monaco, monospace',
                        fontSize: 13,
                      }}
                    >
                      {k}
                    </span>
                  ))}
                </td>
                <td style={{ padding: '6px 0', color: 'var(--bodyText, #323130)' }}>{s.description}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <p style={{ marginTop: 16, marginBottom: 0, color: 'var(--neutralSecondary, #605e5c)', fontSize: 12 }}>
          Shortcuts don&apos;t fire while typing in a text box.
        </p>
      </div>
    </Modal>
  );
};

export default ShortcutHelpModalHost;
