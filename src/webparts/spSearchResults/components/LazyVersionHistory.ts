import * as React from 'react';
import { lazyBridge } from '../../../utilities/lazyBridge';

export const LazyVersionHistory = lazyBridge(
  () => import(/* webpackChunkName: 'VersionHistory' */ 'spfx-toolkit/lib/components/VersionHistory').then(m => ({ default: m.VersionHistory as unknown as React.ComponentType<Record<string, unknown>> })),
  {
    errorMessage: 'Failed to load Version History component',
    minLoadingTime: 200,
  }
);
