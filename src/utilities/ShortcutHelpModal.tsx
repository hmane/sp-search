/**
 * T2.D9 — keyboard-shortcut help modal. Listens for the global
 * `sp-search:open-shortcut-help` event (dispatched by `?` from any web
 * part) and renders a Fluent Modal listing every shipped shortcut.
 *
 * Mounted via `<ShortcutHelpModalHost />` in the SearchBox web part
 * shell (the most common entry point on a search page). Only one host
 * should be mounted per page; multiple hosts harmlessly all show the
 * same modal.
 */

import * as React from 'react';
import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import {
  SP_SEARCH_OPEN_SHORTCUT_HELP,
  useGlobalShortcuts,
  dispatchOpenShortcutHelp,
  dispatchFocusSearchBox,
} from './globalShortcuts';

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
  const [open, setOpen] = React.useState<boolean>(false);

  // Install the global shortcut bindings that drive the help modal +
  // the search-box-focus event. `?` opens this modal; `/` dispatches a
  // separate event the SearchBox component listens for.
  const bindings = React.useMemo(() => [
    {
      key: '?',
      shift: true, // '?' is shift+/ on US layouts; e.key === '?' already implies shift in jsdom but be explicit
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
  ], []);
  useGlobalShortcuts(bindings);

  // Listen for the open event (dispatched by the `?` handler above
  // OR by any other web part that wants to trigger help programmatically).
  React.useEffect((): (() => void) => {
    const handler = (): void => { setOpen(true); };
    window.addEventListener(SP_SEARCH_OPEN_SHORTCUT_HELP, handler);
    return (): void => { window.removeEventListener(SP_SEARCH_OPEN_SHORTCUT_HELP, handler); };
  }, []);

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
                        border: '1px solid #c8c6c4',
                        borderRadius: 3,
                        backgroundColor: '#faf9f8',
                        fontFamily: 'Consolas, Monaco, monospace',
                        fontSize: 13,
                      }}
                    >
                      {k}
                    </span>
                  ))}
                </td>
                <td style={{ padding: '6px 0', color: '#323130' }}>{s.description}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <p style={{ marginTop: 16, marginBottom: 0, color: '#605e5c', fontSize: 12 }}>
          Shortcuts don&apos;t fire while typing in a text box.
        </p>
      </div>
    </Modal>
  );
};

export default ShortcutHelpModalHost;
