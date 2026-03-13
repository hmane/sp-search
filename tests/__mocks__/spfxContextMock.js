// Mock for spfx-toolkit's SPContext utility.
// SPContext wraps the SPFx context object and configures PnPjs — it requires
// a live SharePoint page context which is unavailable in Jest/jsdom.
// Tests that need SPContext behavior should mock at the service boundary.
const SPContext = {
  basic: jest.fn(),
  sp: {},
  spPessimistic: {},
};

module.exports = { SPContext };
