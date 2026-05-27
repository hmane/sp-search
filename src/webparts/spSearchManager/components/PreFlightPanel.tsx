/**
 * T4.D9 — Pre-Flight tab for the Admin Manager.
 *
 * Runs `runTenantReadinessScan()` and renders a status grid with one row
 * per check (green / yellow / red icon + title + message + Fix-this link
 * when applicable). Designed to be screenshot-ready: a single image of
 * this tab is enough evidence for a peer hand-off or support ticket.
 */

import * as React from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { runTenantReadinessScan, type IReadinessReport, type ReadinessStatus } from '@services/index';
import styles from './SpSearchManager.module.scss';

const STATUS_ICON: Record<ReadinessStatus, { name: string; color: string }> = {
  green: { name: 'CheckMark', color: '#107c10' },
  yellow: { name: 'Warning', color: '#d29200' },
  red: { name: 'ErrorBadge', color: '#a4262c' },
};

const PreFlightPanel: React.FC = () => {
  const [report, setReport] = React.useState<IReadinessReport | undefined>(undefined);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);

  const abortRef = React.useRef<AbortController | undefined>(undefined);

  const load = React.useCallback((): void => {
    if (abortRef.current) { abortRef.current.abort(); }
    const controller = new AbortController();
    abortRef.current = controller;
    setIsLoading(true);
    setError(undefined);
    runTenantReadinessScan(controller.signal)
      .then((r): void => {
        if (controller.signal.aborted) { return; }
        setReport(r);
        setIsLoading(false);
      })
      .catch((err): void => {
        if (controller.signal.aborted) { return; }
        setError(err instanceof Error ? err.message : 'Pre-flight scan failed');
        setIsLoading(false);
      });
  }, []);

  React.useEffect((): (() => void) => {
    load();
    return (): void => {
      if (abortRef.current) { abortRef.current.abort(); }
    };
  }, [load]);

  if (isLoading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size={SpinnerSize.large} label="Running tenant-readiness scan..." />
      </div>
    );
  }

  if (error || !report) {
    return (
      <div style={{ padding: 16 }}>
        <MessageBar messageBarType={MessageBarType.error}>{error || 'No report data.'}</MessageBar>
        <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Retry" onClick={load} style={{ marginTop: 12 }} />
      </div>
    );
  }

  return (
    <div style={{ padding: 16 }}>
      {/* Toolbar */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 18 }}>Tenant readiness</h2>
          <p style={{ margin: '4px 0 0 0', color: '#605e5c', fontSize: 13 }}>
            {report.allGreen
              ? 'All checks green — this tenant is ready for SP Search.'
              : report.redCount + ' check(s) failed. Address red rows before launching to end users.'}
            <span style={{ marginLeft: 8, color: '#888', fontSize: 11 }}>
              Last scan: {report.generatedAt.toLocaleTimeString()}
            </span>
          </p>
        </div>
        <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Re-run scan" onClick={load} />
      </div>

      {/* Status grid — one row per check, screenshot-ready */}
      <div role="table" aria-label="Tenant readiness checklist" style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {report.checks.map((check): React.ReactElement => {
          const icon = STATUS_ICON[check.status];
          return (
            <div
              key={check.id}
              role="row"
              style={{
                display: 'flex',
                alignItems: 'flex-start',
                gap: 12,
                padding: 12,
                border: '1px solid #edebe9',
                borderLeft: '4px solid ' + icon.color,
                borderRadius: 4,
                backgroundColor: '#ffffff',
              }}
            >
              <Icon
                iconName={icon.name}
                style={{ color: icon.color, fontSize: 20, marginTop: 2, flexShrink: 0 }}
                aria-label={check.status}
              />
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 600, marginBottom: 4 }}>{check.title}</div>
                <div style={{ fontSize: 13, color: '#323130', marginBottom: check.fix ? 6 : 0 }}>
                  {check.message}
                </div>
                {check.fix && (
                  <a
                    href={check.fix.href}
                    target="_blank"
                    rel="noopener noreferrer"
                    style={{ fontSize: 13, color: '#0078d4', textDecoration: 'none' }}
                  >
                    <Icon iconName="OpenInNewWindow" style={{ fontSize: 12, marginRight: 4 }} />
                    {check.fix.text}
                  </a>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default PreFlightPanel;
