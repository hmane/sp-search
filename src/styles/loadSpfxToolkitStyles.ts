// Force execution of toolkit CSS modules from app entrypoints.
// SPFx ship builds can be inconsistent about package-internal CSS side effects,
// especially for linked packages, so we import the concrete style assets here.

// eslint-disable-next-line @typescript-eslint/no-require-imports
require('spfx-toolkit/lib/components/Card/card.css');
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('spfx-toolkit/lib/components/ErrorBoundary/ErrorBoundary.css');
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('spfx-toolkit/lib/components/UserPersona/UserPersona.css');
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('spfx-toolkit/lib/components/VersionHistory/VersionHistory.css');

export const spfxToolkitStylesLoaded: boolean = true;
