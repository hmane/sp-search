// Mock for @pnp/sp and its sub-path augmentations.
// PnP packages use ESM import syntax and require a live SharePoint context,
// neither of which is available in the Jest/jsdom test environment.
// Tests that need PnP behavior should mock at the service boundary instead.
module.exports = {};
