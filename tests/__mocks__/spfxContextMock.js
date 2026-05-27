// Mock for spfx-toolkit's SPContext utility.
// SPContext wraps the SPFx context object and configures PnPjs — it requires
// a live SharePoint page context which is unavailable in Jest/jsdom.
// Tests that need SPContext behavior should mock at the service boundary
// (e.g. SPContext.http.get / SPContext.sp.*) via the jest.fn() stubs below.
const SPContext = {
  basic: jest.fn(),
  isReady: jest.fn(() => true),
  sp: {},
  spPessimistic: {},
  spCached: {},
  webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/test',
  http: {
    get: jest.fn(),
    post: jest.fn(),
  },
  logger: {
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
    success: jest.fn(),
  },
};

module.exports = { SPContext };
