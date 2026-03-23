import type { IDebugEvent, IQueryDebugInfo, IWebPartDebugConfig, DebugEventType } from './IDebugTypes';

const COLLECTOR_KEY = '__sp_search_debug_collector__';
const MAX_EVENTS = 200;

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const _win = window as any;

interface IDebugCollectorState {
  events: IDebugEvent[];
  lastQuery: IQueryDebugInfo | undefined;
  webPartConfigs: Map<string, IWebPartDebugConfig>;
  nextEventId: number;
  listeners: Set<() => void>;
}

function getState(): IDebugCollectorState {
  if (!_win[COLLECTOR_KEY]) {
    _win[COLLECTOR_KEY] = {
      events: [],
      lastQuery: undefined,
      webPartConfigs: new Map(),
      nextEventId: 1,
      listeners: new Set(),
    };
  }
  return _win[COLLECTOR_KEY];
}

function notify(): void {
  const state = getState();
  state.listeners.forEach((fn) => fn());
}

let _active: boolean | undefined;

export const DebugCollector = {
  isActive(): boolean {
    if (_active === undefined) {
      try {
        _active = window.location.search.indexOf('debug=1') >= 0;
      } catch {
        _active = false;
      }
    }
    return _active;
  },

  logEvent(type: DebugEventType, data: Record<string, unknown>): void {
    if (!DebugCollector.isActive()) return;
    const state = getState();
    const event: IDebugEvent = {
      id: state.nextEventId++,
      type,
      timestamp: Date.now(),
      data,
    };
    state.events.unshift(event);
    if (state.events.length > MAX_EVENTS) {
      state.events.length = MAX_EVENTS;
    }
    notify();
  },

  setLastQuery(info: IQueryDebugInfo): void {
    if (!DebugCollector.isActive()) return;
    const state = getState();
    state.lastQuery = info;
    notify();
  },

  registerWebPart(name: string, properties: Record<string, unknown>): void {
    if (!DebugCollector.isActive()) return;
    const state = getState();
    state.webPartConfigs.set(name, {
      componentName: name,
      properties: { ...properties },
      registeredAt: Date.now(),
    });
    DebugCollector.logEvent('INIT', { componentName: name });
    notify();
  },

  getEvents(): IDebugEvent[] {
    return getState().events;
  },

  getLastQuery(): IQueryDebugInfo | undefined {
    return getState().lastQuery;
  },

  getWebPartConfigs(): Map<string, IWebPartDebugConfig> {
    return getState().webPartConfigs;
  },

  subscribe(listener: () => void): () => void {
    const state = getState();
    state.listeners.add(listener);
    return (): void => { state.listeners.delete(listener); };
  },
};
