'use strict';

const webpack = require('webpack');
const path = require('path');
const fs = require('fs');

// In gulpfile.js __dirname was the project root. This file lives in
// config/, so resolve one level up to reach the project root.
const projectRoot = path.resolve(__dirname, '..');

/**
 * Heft webpack patch for SP Search.
 *
 * Migrated from the gulpfile.js `additionalConfiguration` callback.
 * Signature: receives the generated webpack config, returns the patched config.
 */
module.exports = function (webpackConfig) {
  const isProduction = webpackConfig.mode === 'production';
  const projectNodeModules = path.resolve(projectRoot, 'node_modules');

  // ---------------------------------------------------------------------------
  // Shared dependency aliases — force linked packages (spfx-toolkit etc.) to
  // share the host app's dependency tree and avoid duplicate React instances.
  // ---------------------------------------------------------------------------
  const sharedDependencyAliases = {
    react: path.resolve(projectNodeModules, 'react'),
    'react-dom': path.resolve(projectNodeModules, 'react-dom'),
    'react-hook-form': path.resolve(projectNodeModules, 'react-hook-form'),
    '@fluentui/react': path.resolve(projectNodeModules, '@fluentui/react'),
    '@fluentui/utilities': path.resolve(projectNodeModules, '@fluentui/utilities'),
    '@fluentui/merge-styles': path.resolve(projectNodeModules, '@fluentui/merge-styles'),
    '@fluentui/react-focus': path.resolve(projectNodeModules, '@fluentui/react-focus'),
    devextreme: path.resolve(projectNodeModules, 'devextreme'),
    'devextreme-react': path.resolve(projectNodeModules, 'devextreme-react'),
    inferno: path.resolve(projectNodeModules, 'inferno'),
    tslib: path.resolve(projectNodeModules, 'tslib'),
    zustand: path.resolve(projectNodeModules, 'zustand'),
  };

  // ---------------------------------------------------------------------------
  // Path aliases (match tsconfig.json paths)
  // ---------------------------------------------------------------------------
  webpackConfig.resolve = webpackConfig.resolve || {};
  webpackConfig.resolve.alias = {
    ...webpackConfig.resolve.alias,
    ...sharedDependencyAliases,
    '@store': path.resolve(projectRoot, 'lib/libraries/spSearchStore'),
    '@interfaces': path.resolve(projectRoot, 'lib/libraries/spSearchStore/interfaces'),
    '@services': path.resolve(projectRoot, 'lib/libraries/spSearchStore/services'),
    '@providers': path.resolve(projectRoot, 'lib/libraries/spSearchStore/providers'),
    '@registries': path.resolve(projectRoot, 'lib/libraries/spSearchStore/registries'),
    '@orchestrator': path.resolve(projectRoot, 'lib/libraries/spSearchStore/orchestrator'),
    '@webparts': path.resolve(projectRoot, 'lib/webparts'),
  };

  // ---------------------------------------------------------------------------
  // Module resolution
  // ---------------------------------------------------------------------------
  webpackConfig.resolve.modules = [
    ...(webpackConfig.resolve.modules || []),
    'node_modules',
  ];

  // ---------------------------------------------------------------------------
  // Module rules setup
  // ---------------------------------------------------------------------------
  webpackConfig.module = webpackConfig.module || {};
  webpackConfig.module.rules = webpackConfig.module.rules || [];

  // ---------------------------------------------------------------------------
  // Exclude project lib/ from source-map-loader
  // Heft's source-map-loader rule tries to resolve .module.scss imports in
  // compiled JS files, which fails because CSS is handled by webpack's SASS
  // rules at bundle time, not as separate files in lib/.
  // Match ANY rule referencing source-map-loader (string, use array, or loader prop).
  // ---------------------------------------------------------------------------
  const libDir = path.resolve(projectRoot, 'lib');
  function hasSourceMapLoader(rule) {
    if (!rule) return false;
    // Direct loader string
    if (typeof rule.loader === 'string' && rule.loader.indexOf('source-map-loader') >= 0) return true;
    // use array
    if (Array.isArray(rule.use)) {
      return rule.use.some(function (u) {
        var loader = typeof u === 'string' ? u : (u && u.loader ? u.loader : '');
        return loader.indexOf('source-map-loader') >= 0;
      });
    }
    // Single use string
    if (typeof rule.use === 'string' && rule.use.indexOf('source-map-loader') >= 0) return true;
    return false;
  }
  function addExclude(rule, dir) {
    rule.exclude = rule.exclude
      ? (Array.isArray(rule.exclude) ? [...rule.exclude, dir] : [rule.exclude, dir])
      : [dir];
  }
  webpackConfig.module.rules.forEach(function (rule) {
    if (hasSourceMapLoader(rule)) {
      addExclude(rule, libDir);
    }
    // Also check oneOf groups
    if (Array.isArray(rule.oneOf)) {
      rule.oneOf.forEach(function (inner) {
        if (hasSourceMapLoader(inner)) {
          addExclude(inner, libDir);
        }
      });
    }
  });

  // ---------------------------------------------------------------------------
  // DevExtreme CSS handling
  //
  // Root cause: SPFx CSS rules can process DevExtreme CSS before our custom
  // rule does. When css-loader resolves url() inside dx.light.css it imports
  // binary icon fonts (dxicons.woff2, etc.). If webpack has no matching font
  // loader at that point, it tries to parse the binary as JS.
  //
  // Strategy:
  //   1. Patch every SPFx CSS rule to exclude dxCssDir.
  //   2. Add an exact icon-font rule for devextreme/dist/css/icons/*.
  //   3. Add our own exclusive DevExtreme CSS rule with css-loader url:true
  //      so the font URLs resolve through the explicit font rule.
  //   4. Keep a generic font asset rule as a final safety net.
  // ---------------------------------------------------------------------------
  const dxCssDir = path.resolve(projectRoot, 'node_modules/devextreme/dist/css');
  const dxCssIconsDir = path.resolve(projectRoot, 'node_modules/devextreme/dist/css/icons');
  const toolkitPackageDir = path.resolve(projectRoot, 'node_modules/spfx-toolkit');
  const toolkitRealPackageDir = fs.existsSync(toolkitPackageDir)
    ? fs.realpathSync(toolkitPackageDir)
    : toolkitPackageDir;
  const toolkitCssDirs = Array.from(
    new Set([
      path.join(toolkitPackageDir, 'lib'),
      path.join(toolkitPackageDir, 'esm'),
      path.join(toolkitRealPackageDir, 'lib'),
      path.join(toolkitRealPackageDir, 'esm'),
    ])
  );

  // Helper: does this rule's test match .css files?
  function ruleMatchesCss(r) {
    if (!r || !r.test) return false;
    try {
      if (r.test instanceof RegExp) return r.test.test('dummy.css');
      if (typeof r.test === 'string') return r.test === '.css' || r.test.indexOf('css') >= 0;
    } catch (e) { /* ignore */ }
    return false;
  }

  // Helper: append paths to a rule's exclude list
  function excludePaths(r, paths, idx, label) {
    const prev = r.exclude;
    r.exclude = prev
      ? (Array.isArray(prev) ? [...prev, ...paths] : [prev, ...paths])
      : [...paths];
    console.log('[SP Search] Excluded ' + label + ' from rule[' + idx + ']');
  }

  // Patch every CSS rule at the top level
  (webpackConfig.module.rules || []).forEach((r, idx) => {
    if (ruleMatchesCss(r)) {
      excludePaths(r, [dxCssDir], idx, 'devextreme CSS');
      excludePaths(r, toolkitCssDirs, idx, 'spfx-toolkit CSS');
    }
    // Also patch inside any oneOf groups
    if (Array.isArray(r.oneOf)) {
      r.oneOf.forEach((inner, innerIdx) => {
        if (ruleMatchesCss(inner)) {
          const ruleLabel = idx + '.oneOf[' + innerIdx + ']';
          excludePaths(inner, [dxCssDir], ruleLabel, 'devextreme CSS');
          excludePaths(inner, toolkitCssDirs, ruleLabel, 'spfx-toolkit CSS');
        }
      });
    }
  });

  // ---------------------------------------------------------------------------
  // DevExtreme icon font rule — webpack 5 asset/resource (replaces file-loader)
  // ---------------------------------------------------------------------------
  const dxIconFontRule = {
    test: /\.(woff2?|ttf|eot|svg)(\?.*)?$/i,
    include: [dxCssIconsDir],
    type: 'asset/resource',
    generator: {
      filename: 'devextreme-icons/[name]_[contenthash][ext]',
    },
  };
  webpackConfig.module.rules.unshift(dxIconFontRule);

  // ---------------------------------------------------------------------------
  // Exclusive DevExtreme CSS rule
  // ---------------------------------------------------------------------------
  webpackConfig.module.rules.push({
    test: /\.css$/,
    include: [dxCssDir],
    use: [
      require.resolve('style-loader'),
      {
        loader: require.resolve('css-loader'),
        options: { url: true, import: false },
      },
    ],
  });

  // ---------------------------------------------------------------------------
  // spfx-toolkit CSS rule — explicitly bundle CSS shipped by the linked
  // spfx-toolkit package. SPFx's default CSS rules do not reliably pick up
  // package CSS from symlinked file: dependencies during ship builds.
  // ---------------------------------------------------------------------------
  webpackConfig.module.rules.push({
    test: /\.css$/,
    include: toolkitCssDirs,
    sideEffects: true,
    use: [
      require.resolve('style-loader'),
      {
        loader: require.resolve('css-loader'),
        options: { url: true, import: true },
      },
    ],
  });

  // ---------------------------------------------------------------------------
  // Safety-net font rule — handles any woff/woff2/ttf/eot that still ends up
  // as a webpack module dep (e.g. from a non-CSS import or a rule we missed).
  // Placed via unshift so it is evaluated FIRST, and also injected inside any
  // oneOf group so it wins even if SPFx uses oneOf for asset routing.
  // ---------------------------------------------------------------------------
  const fontRule = {
    test: /\.(woff2?|ttf|eot)(\?.*)?$/i,
    type: 'asset/resource',
  };
  webpackConfig.module.rules.unshift(fontRule);
  (webpackConfig.module.rules || []).forEach(r => {
    if (Array.isArray(r.oneOf)) {
      r.oneOf.unshift(dxIconFontRule);
      r.oneOf.unshift(fontRule);
    }
  });

  // ---------------------------------------------------------------------------
  // Disable webpack's internal source-map-loader
  // SPFx's Heft build adds `devtool: 'source-map'` which causes webpack to
  // inject an internal source-map-loader rule. This rule processes lib/ JS files
  // and tries to resolve .module.scss imports, which don't exist as separate
  // files (CSS is bundled at webpack time via the SASS pipeline).
  // Fix: set devtool to false (disables the internal loader), then use
  // SourceMapDevToolPlugin for production source maps.
  // ---------------------------------------------------------------------------
  // Find and patch the source-map-loader rule to exclude project lib/ files.
  // SPFx adds a source-map-loader rule with enforce:'pre' that processes all JS.
  // When it hits compiled JS that imports .module.scss, it fails because those
  // files don't exist separately (CSS is bundled by the SASS pipeline).
  webpackConfig.module.rules.forEach(function (rule, idx) {
    if (!rule || !rule.use) return;
    var loaderPath = typeof rule.use === 'object' && rule.use.loader ? rule.use.loader : '';
    if (typeof rule.use === 'string') loaderPath = rule.use;
    if (loaderPath.indexOf('source-map-loader') >= 0 && rule.enforce === 'pre') {
      addExclude(rule, libDir);
      console.log('[SP Search] Excluded lib/ from source-map-loader rule[' + idx + ']');
    }
  });

  // ---------------------------------------------------------------------------
  // Plugins
  // ---------------------------------------------------------------------------
  webpackConfig.plugins = webpackConfig.plugins || [];

  // Ignore unnecessary DevExtreme locales (only keep en)
  webpackConfig.plugins.push(
    new webpack.IgnorePlugin({
      resourceRegExp: /^\.\/locale$/,
      contextRegExp: /devextreme/,
    })
  );

  // Ignore moment.js locales if moment is used
  webpackConfig.plugins.push(
    new webpack.IgnorePlugin({
      resourceRegExp: /^\.\/locale$/,
      contextRegExp: /moment$/,
    })
  );

  // ---------------------------------------------------------------------------
  // Production vs Development settings
  // ---------------------------------------------------------------------------
  if (isProduction) {
    // Production optimizations
    webpackConfig.optimization = {
      ...webpackConfig.optimization,
      usedExports: true,
      moduleIds: 'deterministic',
      chunkIds: 'deterministic',
    };

    // Production source maps — use plugin instead of devtool to avoid
    // webpack's internal source-map-loader which fails on .module.scss imports
    webpackConfig.devtool = false;
    webpackConfig.plugins.push(
      new webpack.SourceMapDevToolPlugin({
        filename: '[file].map',
        append: false, // hidden source maps (not linked in bundle)
      })
    );

    // Bundle analyzer (only when ANALYZE env var is set)
    if (process.env.ANALYZE) {
      const bundleAnalyzer = require('webpack-bundle-analyzer');
      webpackConfig.plugins.push(
        new bundleAnalyzer.BundleAnalyzerPlugin({
          analyzerMode: 'static',
          reportFilename: path.join(projectRoot, 'temp', 'stats', 'bundle-report.html'),
          openAnalyzer: false,
          generateStatsFile: true,
          statsFilename: path.join(projectRoot, 'temp', 'stats', 'bundle-stats.json'),
          logLevel: 'warn',
        })
      );
    }

    console.log('[SP Search] Production build - Optimized');
  } else {
    // Development optimizations
    webpackConfig.optimization = {
      ...webpackConfig.optimization,
      moduleIds: 'named',
      chunkIds: 'named',
    };

    // Filesystem cache for faster rebuilds
    webpackConfig.cache = {
      type: 'filesystem',
      buildDependencies: {
        config: [__filename, path.resolve(projectRoot, 'tsconfig.json')],
      },
      cacheDirectory: path.resolve(projectRoot, 'node_modules/.cache/webpack'),
      name: 'spfx-dev-cache',
    };

    // Development source maps — use devtool:false + plugin to avoid
    // webpack's internal source-map-loader which fails on .module.scss imports
    webpackConfig.devtool = false;
    webpackConfig.plugins.push(
      new webpack.EvalSourceMapDevToolPlugin({
        moduleFilenameTemplate: '[resource-path]',
      })
    );

    console.log('[SP Search] Development build - Fast compilation with filesystem cache');
  }

  return webpackConfig;
};
