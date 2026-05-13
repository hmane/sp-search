import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DebugCollector } from '@store/debug';
import type { IDebugEvent, IQueryDebugInfo, DebugEventType, INetworkEvent } from '@store/debug';
import type { StoreApi } from 'zustand/vanilla';
import type { ISearchStore } from '@interfaces/index';
import styles from './DebugPanel.module.scss';

// --- Helpers ---

function formatTime(ts: number): string {
  const d = new Date(ts);
  const ms = String(d.getMilliseconds());
  return d.toLocaleTimeString([], { hour12: false }) + '.' + ('000' + ms).slice(-3);
}

function timingClass(ms: number): string {
  if (ms < 500) return styles.timingGreen;
  if (ms <= 2000) return styles.timingYellow;
  return styles.timingRed;
}

const LOG_TYPE_CLASSES: Record<DebugEventType, string> = {
  SEARCH: styles.logTypeSearch,
  FILTER: styles.logTypeFilter,
  VERTICAL: styles.logTypeVertical,
  URL: styles.logTypeUrl,
  ERROR: styles.logTypeError,
  INIT: styles.logTypeInit,
};

// --- Collapsible JSON Tree ---

interface IJsonTreeProps {
  data: unknown;
  depth?: number;
  defaultExpanded?: boolean;
}

const JsonTree: React.FC<IJsonTreeProps> = ({ data, depth = 0, defaultExpanded = true }) => {
  const [expanded, setExpanded] = React.useState(defaultExpanded);

  if (data === null || data === undefined) {
    return <span className={styles.jsonNull}>{String(data)}</span>;
  }
  if (typeof data === 'string') {
    return <span className={styles.jsonString}>"{data}"</span>;
  }
  if (typeof data === 'number') {
    return <span className={styles.jsonNumber}>{data}</span>;
  }
  if (typeof data === 'boolean') {
    return <span className={styles.jsonBoolean}>{String(data)}</span>;
  }

  if (Array.isArray(data)) {
    if (data.length === 0) {
      return <span className={styles.jsonBracket}>[]</span>;
    }
    return (
      <span>
        <span className={styles.jsonToggle} onClick={() => setExpanded(!expanded)}>
          {expanded ? '\u25BC' : '\u25B6'}
        </span>
        <span className={styles.jsonBracket}>[{data.length}]</span>
        {expanded && data.map((item, i) => (
          <div key={i} className={styles.jsonRow}>
            <span className={styles.jsonKey}>{i}: </span>
            <JsonTree data={item} depth={depth + 1} defaultExpanded={depth < 1} />
          </div>
        ))}
      </span>
    );
  }

  if (typeof data === 'object') {
    const keys = Object.keys(data as Record<string, unknown>);
    if (keys.length === 0) {
      return <span className={styles.jsonBracket}>{'{}'}</span>;
    }
    const obj = data as Record<string, unknown>;
    return (
      <span>
        <span className={styles.jsonToggle} onClick={() => setExpanded(!expanded)}>
          {expanded ? '\u25BC' : '\u25B6'}
        </span>
        <span className={styles.jsonBracket}>{'{'}{keys.length}{'}'}</span>
        {expanded && keys.map((key: string) => (
          <div key={key} className={styles.jsonRow}>
            <span className={styles.jsonKey}>{key}: </span>
            <JsonTree data={obj[key]} depth={depth + 1} defaultExpanded={depth < 1} />
          </div>
        ))}
      </span>
    );
  }

  return <span>{String(data)}</span>;
};

// --- Query Tab ---

const QueryTab: React.FC = () => {
  const [query, setQuery] = React.useState<IQueryDebugInfo | undefined>(DebugCollector.getLastQuery());

  React.useEffect(() => {
    return DebugCollector.subscribe(() => {
      setQuery(DebugCollector.getLastQuery());
    });
  }, []);

  if (!query) {
    return <div style={{ color: '#888', padding: 16 }}>No search executed yet.</div>;
  }

  return (
    <div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>KQL</span>
        <span className={styles.queryValue}>{query.kql || '(empty)'}</span>
      </div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>Query Template</span>
        <span className={styles.queryValue}>{query.queryTemplate}</span>
      </div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>Result Source ID</span>
        <span className={styles.queryValue}>{query.resultSourceId || '(default)'}</span>
      </div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>Provider</span>
        <span className={styles.queryValue}>{query.providerId}</span>
      </div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>Timing</span>
        <span className={styles.queryValue}>
          {query.duration !== undefined ? (
            <>
              {query.duration}ms
              <span className={`${styles.timingBadge} ${timingClass(query.duration)}`}>
                {query.duration}ms
              </span>
            </>
          ) : 'In progress...'}
        </span>
      </div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>Results</span>
        <span className={styles.queryValue}>
          {query.itemsReturned !== undefined
            ? `${query.itemsReturned} of ${query.totalCount} (page ${query.currentPage}, size ${query.pageSize})`
            : 'Pending...'}
        </span>
      </div>
      {query.refinementFilters.length > 0 && (
        <div className={styles.queryField}>
          <span className={styles.queryLabel}>Refinement Filters</span>
          <span className={styles.queryValue}>
            <JsonTree data={query.refinementFilters} defaultExpanded={true} />
          </span>
        </div>
      )}
      {query.refiners.length > 0 && (
        <div className={styles.queryField}>
          <span className={styles.queryLabel}>Refiners</span>
          <span className={styles.queryValue}>
            <JsonTree data={query.refiners} defaultExpanded={false} />
          </span>
        </div>
      )}
      {query.error && (
        <div className={styles.queryField}>
          <span className={styles.queryLabel} style={{ color: '#ff5050' }}>Error</span>
          <span className={styles.queryValue} style={{ color: '#ff5050' }}>{query.error}</span>
        </div>
      )}
      {query.request && (
        <div className={styles.stateSection}>
          <div className={styles.stateSectionHeader}>Request (ISearchQuery)</div>
          <JsonTree data={query.request} defaultExpanded={false} />
        </div>
      )}
      {query.response && (
        <div className={styles.stateSection}>
          <div className={styles.stateSectionHeader}>Response Summary</div>
          <JsonTree data={query.response} defaultExpanded={false} />
        </div>
      )}
    </div>
  );
};

// --- State Tab (with change highlight) ---

interface IStateTabProps {
  store: StoreApi<ISearchStore>;
}

const StateTab: React.FC<IStateTabProps> = ({ store }) => {
  const [snapshot, setSnapshot] = React.useState<Record<string, unknown>>({});
  const [changedKeys, setChangedKeys] = React.useState<Set<string>>(new Set());
  const prevSnapshotRef = React.useRef<Record<string, unknown>>({});

  React.useEffect(() => {
    function update(): void {
      const s = store.getState();
      const next: Record<string, unknown> = {
        querySlice: {
          queryText: s.queryText,
          scope: s.scope,
          queryTemplate: s.queryTemplate,
        },
        filterSlice: {
          activeFilters: s.activeFilters,
          displayRefinersCount: s.displayRefiners.length,
          filterConfigCount: s.filterConfig.length,
          operatorBetweenFilters: s.operatorBetweenFilters,
        },
        verticalSlice: {
          currentVerticalKey: s.currentVerticalKey,
          verticalsCount: s.verticals.length,
          verticalCounts: s.verticalCounts,
        },
        resultSlice: {
          itemCount: s.items.length,
          totalCount: s.totalCount,
          currentPage: s.currentPage,
          pageSize: s.pageSize,
          sort: s.sort,
          promotedResultsCount: s.promotedResults.length,
          isLoading: s.isLoading,
        },
        uiSlice: {
          activeLayoutKey: s.activeLayoutKey,
          availableLayouts: s.availableLayouts,
        },
        registries: {
          dataProviders: s.registries.dataProviders.getAll().map((p) => p.id),
          actions: s.registries.actions.getAll().map((a) => a.id),
          layouts: s.registries.layouts.getAll().map((l) => l.id),
          filterTypes: s.registries.filterTypes.getAll().map((f) => f.id),
          suggestions: s.registries.suggestions.getAll().map((sp) => sp.id),
        },
      };

      // Detect changed top-level slice keys for highlight flash
      const prev = prevSnapshotRef.current;
      const newChanged = new Set<string>();
      for (const key of Object.keys(next)) {
        if (JSON.stringify(prev[key]) !== JSON.stringify(next[key])) {
          newChanged.add(key);
        }
      }
      prevSnapshotRef.current = next;
      setSnapshot(next);

      if (newChanged.size > 0) {
        setChangedKeys(newChanged);
        setTimeout(() => setChangedKeys(new Set()), 2000);
      }
    }
    update();
    return store.subscribe(update);
  }, [store]);

  return (
    <div>
      {Object.keys(snapshot).map((key: string) => (
        <div
          key={key}
          className={`${styles.stateSection}${changedKeys.has(key) ? ' ' + styles.jsonHighlight : ''}`}
        >
          <JsonTree data={{ [key]: snapshot[key] }} defaultExpanded={true} />
        </div>
      ))}
    </div>
  );
};

// --- Config Tab ---

const ConfigTab: React.FC = () => {
  const [configs, setConfigs] = React.useState<Array<{ name: string; properties: Record<string, unknown> }>>([]);

  React.useEffect(() => {
    function update(): void {
      const map = DebugCollector.getWebPartConfigs();
      const arr: Array<{ name: string; properties: Record<string, unknown> }> = [];
      map.forEach((config) => {
        arr.push({ name: config.componentName, properties: config.properties });
      });
      setConfigs(arr);
    }
    update();
    return DebugCollector.subscribe(update);
  }, []);

  if (configs.length === 0) {
    return <div style={{ color: '#888', padding: 16 }}>No web parts registered yet.</div>;
  }

  return (
    <div>
      {configs.map((config) => (
        <div key={config.name} className={styles.stateSection}>
          <div className={styles.stateSectionHeader}>{config.name}</div>
          <JsonTree data={config.properties} defaultExpanded={false} />
        </div>
      ))}
    </div>
  );
};

// --- Log Tab ---

const ALL_EVENT_TYPES: DebugEventType[] = ['SEARCH', 'FILTER', 'VERTICAL', 'URL', 'ERROR', 'INIT'];

const LogTab: React.FC = () => {
  const [events, setEvents] = React.useState<IDebugEvent[]>(DebugCollector.getEvents());
  const [enabledTypes, setEnabledTypes] = React.useState<Set<DebugEventType>>(new Set(ALL_EVENT_TYPES));

  React.useEffect(() => {
    return DebugCollector.subscribe(() => {
      setEvents([...DebugCollector.getEvents()]);
    });
  }, []);

  const toggleType = (type: DebugEventType): void => {
    setEnabledTypes((prev) => {
      const next = new Set(prev);
      if (next.has(type)) {
        next.delete(type);
      } else {
        next.add(type);
      }
      return next;
    });
  };

  const filtered = events.filter((e) => enabledTypes.has(e.type));

  return (
    <div>
      <div className={styles.logFilters}>
        {ALL_EVENT_TYPES.map((type) => (
          <label key={type} className={styles.logFilterCheckbox}>
            <input
              type="checkbox"
              checked={enabledTypes.has(type)}
              onChange={() => toggleType(type)}
            />
            <span className={LOG_TYPE_CLASSES[type]}>{type}</span>
          </label>
        ))}
      </div>
      {filtered.map((event) => (
        <div
          key={event.id}
          className={`${styles.logEntry}${event.type === 'ERROR' ? ' ' + styles.logEntryError : ''}`}
        >
          <span className={styles.logTimestamp}>{formatTime(event.timestamp)}</span>
          <span className={`${styles.logType} ${LOG_TYPE_CLASSES[event.type]}`}>{event.type}</span>
          <span className={styles.logData}>{JSON.stringify(event.data)}</span>
        </div>
      ))}
      {filtered.length === 0 && (
        <div style={{ color: '#888', padding: 16 }}>No events matching filter.</div>
      )}
    </div>
  );
};

// --- Main Panel ---

type TabKey = 'query' | 'state' | 'config' | 'log' | 'network';

// T5.D2 — Network tab thresholds for the timing badge colour.
const TIMING_FAST_MS = 400;     // green
const TIMING_SLOW_MS = 1200;    // yellow above this; red is anything slower

function timingBadgeClass(durationMs: number | undefined): string {
  if (durationMs === undefined) {
    return styles.timingYellow;
  }
  if (durationMs <= TIMING_FAST_MS) {
    return styles.timingGreen;
  }
  if (durationMs <= TIMING_SLOW_MS) {
    return styles.timingYellow;
  }
  return styles.timingRed;
}

const NetworkTab: React.FC = () => {
  const [events, setEvents] = React.useState<INetworkEvent[]>(DebugCollector.getNetworkEvents());

  React.useEffect(() => {
    return DebugCollector.subscribe(() => {
      setEvents([...DebugCollector.getNetworkEvents()]);
    });
  }, []);

  if (events.length === 0) {
    return (
      <div style={{ color: '#888', padding: 16 }}>
        No network calls captured yet. Trigger a search to populate this tab.
        Buffer cap: 50 most-recent calls.
      </div>
    );
  }

  return (
    <div>
      {events.map((event) => {
        const statusClass = event.status === 'error'
          ? styles.logTypeNetworkError
          : event.status === 'aborted'
            ? styles.logTypeNetworkAborted
            : styles.logTypeNetwork;
        const kindLabel = event.kind === 'verticalCount'
          ? 'vertical:' + (event.verticalKey || '?')
          : 'search';
        return (
          <div
            key={event.id}
            className={`${styles.logEntry}${event.status === 'error' ? ' ' + styles.logEntryError : ''}`}
          >
            <span className={styles.logTimestamp}>{formatTime(event.timestamp)}</span>
            <span className={`${styles.logType} ${statusClass}`}>{kindLabel}</span>
            <span className={`${styles.timingBadge} ${timingBadgeClass(event.durationMs)}`}>
              {event.durationMs === undefined ? 'in-flight' : event.durationMs + 'ms'}
            </span>
            <span className={styles.logData}>
              {event.status === 'ok'
                ? 'p' + event.currentPage + '×' + event.pageSize +
                  ' → ' + (event.itemCount ?? 0) + '/' + (event.totalCount ?? 0)
                : (event.errorMessage || event.status)
              }
              {' · '}
              {event.providerId}
            </span>
          </div>
        );
      })}
    </div>
  );
};

interface IDebugPanelProps {
  store: StoreApi<ISearchStore>;
  onClose: () => void;
}

const TABS: Array<{ key: TabKey; label: string }> = [
  { key: 'query', label: 'Query' },
  { key: 'network', label: 'Network' },
  { key: 'state', label: 'State' },
  { key: 'config', label: 'Config' },
  { key: 'log', label: 'Log' },
];

const DEFAULT_HEIGHT = 0.6;

const DebugPanel: React.FC<IDebugPanelProps> = ({ store, onClose }) => {
  const [activeTab, setActiveTab] = React.useState<TabKey>('query');
  const [height, setHeight] = React.useState(DEFAULT_HEIGHT);
  const panelRef = React.useRef<HTMLDivElement>(null);
  const dragRef = React.useRef<{ startY: number; startHeight: number } | null>(null);

  const onDragStart = React.useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    dragRef.current = { startY: e.clientY, startHeight: height };
    const onMove = (me: MouseEvent): void => {
      if (!dragRef.current) return;
      const delta = dragRef.current.startY - me.clientY;
      const newHeight = Math.min(0.9, Math.max(0.2, dragRef.current.startHeight + delta / window.innerHeight));
      setHeight(newHeight);
    };
    const onUp = (): void => {
      dragRef.current = null;
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }, [height]);

  const renderTab = (): React.ReactElement => {
    switch (activeTab) {
      case 'query': return <QueryTab />;
      case 'network': return <NetworkTab />;
      case 'state': return <StateTab store={store} />;
      case 'config': return <ConfigTab />;
      case 'log': return <LogTab />;
    }
  };

  return (
    <div
      ref={panelRef}
      className={styles.debugPanel}
      style={{ height: `${Math.round(height * 100)}vh` }}
    >
      <div className={styles.dragHandle} onMouseDown={onDragStart} />
      <div className={styles.header}>
        <span className={styles.headerTitle}>SP Search Debug</span>
        <div className={styles.headerActions}>
          <button onClick={onClose} title="Close" type="button">
            <Icon iconName="Cancel" />
          </button>
        </div>
      </div>
      <div className={styles.tabBar}>
        {TABS.map((tab) => (
          <button
            key={tab.key}
            className={`${styles.tab}${activeTab === tab.key ? ' ' + styles.tabActive : ''}`}
            onClick={() => setActiveTab(tab.key)}
            type="button"
          >
            {tab.label}
          </button>
        ))}
      </div>
      <div className={styles.tabContent}>
        {renderTab()}
      </div>
    </div>
  );
};

export default DebugPanel;
