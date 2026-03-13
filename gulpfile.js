'use strict';

// Core imports
const build = require('@microsoft/sp-build-web');
const bundleAnalyzer = require('webpack-bundle-analyzer');
const webpack = require('webpack');
const path = require('path');
const fs = require('fs');
const { task } = require('gulp');

// Fast serve configuration
const { addFastServe } = require('spfx-fast-serve-helpers');

addFastServe(build, {
  serve: {
    open: false,
    port: 4321,
    https: true,
  },
});

// Disable SPFx warnings
build.addSuppression(/Warning - \[sass\]/g);
build.addSuppression(/Warning - lint.*/g);

// Main webpack configuration
build.configureWebpack.mergeConfig({
  additionalConfiguration: generatedConfiguration => {
    const isProduction = build.getConfig().production;

    // Configure path aliases to match tsconfig.json
    // These resolve at both TypeScript compilation and webpack bundling
    generatedConfiguration.resolve = generatedConfiguration.resolve || {};
    generatedConfiguration.resolve.alias = {
      ...generatedConfiguration.resolve.alias,
      '@store': path.resolve(__dirname, 'lib/libraries/spSearchStore'),
      '@interfaces': path.resolve(__dirname, 'lib/libraries/spSearchStore/interfaces'),
      '@services': path.resolve(__dirname, 'lib/libraries/spSearchStore/services'),
      '@providers': path.resolve(__dirname, 'lib/libraries/spSearchStore/providers'),
      '@registries': path.resolve(__dirname, 'lib/libraries/spSearchStore/registries'),
      '@orchestrator': path.resolve(__dirname, 'lib/libraries/spSearchStore/orchestrator'),
      '@webparts': path.resolve(__dirname, 'lib/webparts'),
    };

    // Enhanced module resolution
    generatedConfiguration.resolve.modules = [
      ...(generatedConfiguration.resolve.modules || []),
      'node_modules',
    ];

    // Tree-shaking optimizations
    generatedConfiguration.module = generatedConfiguration.module || {};
    generatedConfiguration.module.rules = generatedConfiguration.module.rules || [];

    // DevExtreme optimization: use individual component imports for tree-shaking
    generatedConfiguration.module.rules.push({
      test: /node_modules[\\/]devextreme-react[\\/].*.js$/,
      sideEffects: false,
    });

    // DevExtreme core: only keep what's imported
    generatedConfiguration.module.rules.push({
      test: /node_modules[\\/]devextreme[\\/](?!dist[\\/]css).*.js$/,
      sideEffects: false,
    });

    // Bundle DevExtreme CSS without breaking the SPFx build.
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

    const dxCssDir = path.resolve(__dirname, 'node_modules/devextreme/dist/css');
    const dxCssIconsDir = path.resolve(__dirname, 'node_modules/devextreme/dist/css/icons');

    // Helper: does this rule's test match .css files?
    function ruleMatchesCss(r) {
      if (!r || !r.test) return false;
      try {
        if (r.test instanceof RegExp) return r.test.test('dummy.css');
        if (typeof r.test === 'string') return r.test === '.css' || r.test.indexOf('css') >= 0;
      } catch (e) { /* ignore */ }
      return false;
    }

    // Helper: append dxCssDir to a rule's exclude list
    function excludeDxCss(r, idx) {
      const prev = r.exclude;
      r.exclude = prev
        ? (Array.isArray(prev) ? [...prev, dxCssDir] : [prev, dxCssDir])
        : [dxCssDir];
      console.log('[SP Search] Excluded devextreme CSS from rule[' + idx + ']');
    }

    // Patch every CSS rule at the top level
    (generatedConfiguration.module.rules || []).forEach((r, idx) => {
      if (ruleMatchesCss(r)) {
        excludeDxCss(r, idx);
      }
      // Also patch inside any oneOf groups
      if (Array.isArray(r.oneOf)) {
        r.oneOf.forEach((inner, innerIdx) => {
          if (ruleMatchesCss(inner)) {
            excludeDxCss(inner, idx + '.oneOf[' + innerIdx + ']');
          }
        });
      }
    });

    const dxIconFontRule = {
      test: /\.(woff2?|ttf|eot|svg)(\?.*)?$/i,
      include: [dxCssIconsDir],
      use: [
        {
          loader: require.resolve('file-loader'),
          options: {
            name: 'devextreme-icons/[name]_[contenthash].[ext]'
          }
        }
      ]
    };
    generatedConfiguration.module.rules.unshift(dxIconFontRule);

    // Exclusive DevExtreme CSS rule.
    generatedConfiguration.module.rules.push({
      test: /\.css$/,
      include: [dxCssDir],
      use: [
        require.resolve('style-loader'),
        {
          loader: require.resolve('css-loader'),
          options: { url: true, import: false }
        },
      ],
    });

    // Safety-net font rule — handles any woff/woff2/ttf/eot that still ends up
    // as a webpack module dep (e.g. from a non-CSS import or a rule we missed).
    // Placed via unshift so it is evaluated FIRST, and also injected inside any
    // oneOf group so it wins even if SPFx uses oneOf for asset routing.
    const fontRule = {
      test: /\.(woff2?|ttf|eot)(\?.*)?$/i,
      type: 'asset/resource',
    };
    generatedConfiguration.module.rules.unshift(fontRule);
    (generatedConfiguration.module.rules || []).forEach(r => {
      if (Array.isArray(r.oneOf)) {
        r.oneOf.unshift(dxIconFontRule);
        r.oneOf.unshift(fontRule);
      }
    });

    // Bundle optimization plugins
    generatedConfiguration.plugins = generatedConfiguration.plugins || [];

    // Ignore unnecessary DevExtreme locales (only keep en)
    generatedConfiguration.plugins.push(
      new webpack.IgnorePlugin({
        resourceRegExp: /^\.\/locale$/,
        contextRegExp: /devextreme/
      })
    );

    // Ignore moment.js locales if moment is used
    generatedConfiguration.plugins.push(
      new webpack.IgnorePlugin({
        resourceRegExp: /^\.\/locale$/,
        contextRegExp: /moment$/
      })
    );

    if (isProduction) {
      // Production optimizations
      generatedConfiguration.optimization = {
        ...generatedConfiguration.optimization,
        usedExports: true,
        moduleIds: 'deterministic',
        chunkIds: 'deterministic',
      };

      // Production source maps
      generatedConfiguration.devtool = 'hidden-source-map';

      // Bundle analyzer (only when ANALYZE env var is set)
      if (process.env.ANALYZE) {
        generatedConfiguration.plugins.push(
          new bundleAnalyzer.BundleAnalyzerPlugin({
            analyzerMode: 'static',
            reportFilename: path.join(__dirname, 'temp', 'stats', 'bundle-report.html'),
            openAnalyzer: false,
            generateStatsFile: true,
            statsFilename: path.join(__dirname, 'temp', 'stats', 'bundle-stats.json'),
            logLevel: 'warn',
          })
        );
      }

      console.log('🏗️  Production build - Optimized for SP Search');
    } else {
      // Development optimizations
      generatedConfiguration.optimization = {
        ...generatedConfiguration.optimization,
        moduleIds: 'named',
        chunkIds: 'named',
      };

      // Filesystem cache for faster rebuilds
      generatedConfiguration.cache = {
        type: 'filesystem',
        buildDependencies: {
          config: [__filename, path.resolve(__dirname, 'tsconfig.json')],
        },
        cacheDirectory: path.resolve(__dirname, 'node_modules/.cache/webpack'),
        name: 'spfx-dev-cache',
      };

      // Development source maps
      generatedConfiguration.devtool = 'eval-cheap-module-source-map';

      console.log('🔧 Development build - Fast compilation with filesystem cache');
    }

    return generatedConfiguration;
  },
});

// Utility tasks
task('clean-cache', done => {
  console.log('🧹 Clearing build caches...');
  const cacheDir = path.join(__dirname, 'node_modules/.cache');

  if (fs.existsSync(cacheDir)) {
    fs.rmSync(cacheDir, { recursive: true, force: true });
    console.log('✅ Cache cleared successfully');
  } else {
    console.log('ℹ️  No cache found');
  }
  done();
});

task('analyze-bundle', done => {
  const reportPath = path.join(__dirname, 'temp', 'stats', 'bundle-report.html');

  if (fs.existsSync(reportPath)) {
    console.log(`📊 Bundle report: ${reportPath}`);
  } else {
    console.log('❌ Run production build first');
  }
  done();
});

// Clean all build artifacts
task('clean-all', done => {
  console.log('🧹 Cleaning all build artifacts...');

  const dirsToClean = [
    'lib',
    'dist',
    'temp',
    'release',
    'sharepoint/solution',
    'node_modules/.cache'
  ];

  dirsToClean.forEach(dir => {
    const fullPath = path.join(__dirname, dir);
    if (fs.existsSync(fullPath)) {
      fs.rmSync(fullPath, { recursive: true, force: true });
      console.log(`  ✓ Removed ${dir}`);
    }
  });

  console.log('✅ Clean complete\n');
  done();
});

// Initialize build
build.initialize(require('gulp'));
