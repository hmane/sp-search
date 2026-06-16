// Mock for `*WebPartStrings` localized-strings modules.
//
// Returns an empty object, so an unmocked `strings.SomeKey` resolves to
// `undefined` (the original, pre-identity-proxy behavior). This is intentionally
// NOT the identity Proxy used for SCSS modules (styleMock.js): a strings module
// must yield falsy values for unmocked keys so a test rendering a
// strings-consuming component doesn't get truthy placeholder strings that mask a
// missing-localization bug.
module.exports = {};
