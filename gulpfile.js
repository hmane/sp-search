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
