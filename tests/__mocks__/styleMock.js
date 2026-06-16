// Identity proxy — SCSS-module class names resolve to their literal key string.
// e.g. styles.gridBadge === 'gridBadge', styles['gridBadge--green'] === 'gridBadge--green'
//
// __esModule: true prevents ts-jest's __importDefault from wrapping this in
// { default: proxy }, which would make `import styles from '*.scss'` resolve to
// the string 'default' instead of the proxy.
// The 'default' key returns the proxy itself for the same reason.
const RESERVED = { __esModule: true };

const proxy = new Proxy(RESERVED, {
  get: function (target, prop) {
    if (prop === '__esModule') {
      return true;
    }
    if (prop === 'default') {
      return proxy;
    }
    if (typeof prop === 'string') {
      return prop;
    }
    return undefined;
  },
});

module.exports = proxy;
