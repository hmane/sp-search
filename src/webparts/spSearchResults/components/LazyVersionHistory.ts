import * as React from 'react';
import { createLazyComponent } from 'spfx-toolkit/lib/utilities/lazyLoader';

export const LazyVersionHistory: React.FC<any> = createLazyComponent(
  () => import('spfx-toolkit/lib/components/VersionHistory').then(m => ({ default: m.VersionHistory })),
  {
    errorMessage: 'Failed to load Version History component',
    minLoadingTime: 200,
  }
);
