export { DebugCollector } from './DebugCollector';
export type { IDebugEvent, INetworkEvent, IQueryDebugInfo, IWebPartDebugConfig, DebugEventType } from './IDebugTypes';
// T5.D4 — extensible tab registry consumed by DebugPanel.
export {
  registerDebugTab,
  getRegisteredDebugTabs,
  unregisterDebugTab,
} from './debugTabRegistry';
export type { IDebugTabContext, IDebugTabRegistration, IRegisterDebugTabOptions } from './debugTabRegistry';
