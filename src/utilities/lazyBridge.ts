import * as React from 'react';
import { createLazyComponent } from 'spfx-toolkit/lib/utilities/lazyLoader';

/**
 * Type bridge for createLazyComponent.
 *
 * spfx-toolkit uses its own @types/react which is structurally incompatible
 * with this project's @types/react (different React.Key type).
 * This helper casts through `unknown` to satisfy both sides.
 */
type LazyImportFn = () => Promise<{ default: React.ComponentType<Record<string, unknown>> }>;

interface ILazyBridgeOptions {
  errorMessage: string;
  minLoadingTime?: number;
}

export function lazyBridge(
  importFn: LazyImportFn,
  options: ILazyBridgeOptions
): React.ComponentType<Record<string, unknown>> {
  return createLazyComponent(
    importFn as unknown as Parameters<typeof createLazyComponent>[0],
    options
  ) as unknown as React.ComponentType<Record<string, unknown>>;
}
