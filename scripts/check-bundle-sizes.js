#!/usr/bin/env node
/**
 * Per-web-part bundle size breach gate (Foundations Found.D7, amended via Task 3.5).
 * Discovers Heft production assets via regex (`<stem>_<contenthash>.js`,
 * minimum 8 hex chars; strict — unhashed bundles are rejected),
 * compares against config/bundle-budgets.json, exits non-zero on breach.
 * Emits release/analysis-logs/bundle-sizes.json for the per-PR attribution dashboard.
 */

const fs = require('fs');
const path = require('path');

const REPO_ROOT = path.resolve(__dirname, '..');
const ASSETS_DIR = path.join(REPO_ROOT, 'release', 'assets');
const BUDGETS_PATH = path.join(REPO_ROOT, 'config', 'bundle-budgets.json');
const REPORT_PATH = path.join(REPO_ROOT, 'release', 'analysis-logs', 'bundle-sizes.json');

if (!fs.existsSync(BUDGETS_PATH)) {
  console.error(`[bundle-gate] missing budgets file: ${BUDGETS_PATH}`);
  process.exit(2);
}

const { budgets } = JSON.parse(fs.readFileSync(BUDGETS_PATH, 'utf8'));
if (!budgets || typeof budgets !== 'object') {
  console.error(`[bundle-gate] budgets file is missing the "budgets" key: ${BUDGETS_PATH}`);
  process.exit(2);
}

if (!fs.existsSync(ASSETS_DIR)) {
  console.error(`[bundle-gate] missing assets directory: ${ASSETS_DIR} (run npm run package first)`);
  process.exit(2);
}

const assetFiles = fs
  .readdirSync(ASSETS_DIR)
  .filter((f) => f.endsWith('.js') && !f.endsWith('.LICENSE.txt'));

const breaches = [];
const report = { schemaVersion: 1, generatedAt: new Date().toISOString(), webParts: {} };

for (const [name, budget] of Object.entries(budgets)) {
  // Derive stem: "sp-search-filters-web-part.js" -> "sp-search-filters-web-part"
  const stem = name.replace(/\.js$/, '');
  // Strict: require exact stem + _<hexhash> (minimum 8 hex chars). Heft currently
  // emits 20-char hashes; the 8-char floor allows for future Heft drift while
  // rejecting partial hashes from corrupt builds and unhashed debug bundles.
  const pattern = new RegExp(`^${stem.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}_[a-f0-9]{8,}\\.js$`);
  const matches = assetFiles.filter((f) => pattern.test(f));

  if (matches.length === 0) {
    console.error(`[bundle-gate] missing hashed asset matching pattern ${stem}_<hash>.js (run npm run package for production output; npm run package:debug emits unhashed bundles which this gate intentionally rejects)`);
    process.exit(2);
  }
  if (matches.length > 1) {
    console.error(`[bundle-gate] ambiguous match for ${stem}: found ${matches.join(', ')} (run heft clean before npm run package)`);
    process.exit(2);
  }

  const matchedAsset = matches[0];
  const file = path.join(ASSETS_DIR, matchedAsset);
  const actual = fs.statSync(file).size;
  const delta = actual - budget;
  const pct = ((actual / budget) * 100).toFixed(1);
  const status = actual <= budget ? 'PASS' : 'BREACH';
  console.log(`[${status}] ${matchedAsset}: ${actual.toLocaleString()} bytes (budget ${budget.toLocaleString()}, ${pct}%, delta ${delta >= 0 ? '+' : ''}${delta.toLocaleString()})`);
  report.webParts[name] = { matchedAsset, actual, budget, delta, pctOfBudget: Number(pct), status };
  if (status === 'BREACH') breaches.push({ name, matchedAsset, actual, budget, delta });
}

fs.mkdirSync(path.dirname(REPORT_PATH), { recursive: true });
fs.writeFileSync(REPORT_PATH, JSON.stringify(report, null, 2));
console.log(`[bundle-gate] report written: ${REPORT_PATH}`);

if (breaches.length > 0) {
  console.error(`\n[bundle-gate] FAILED: ${breaches.length} web part(s) exceed budget`);
  for (const b of breaches) {
    console.error(`  ${b.matchedAsset}: ${b.actual.toLocaleString()} bytes exceeds ${b.budget.toLocaleString()} budget by ${b.delta.toLocaleString()} bytes`);
  }
  process.exit(1);
}

console.log(`\n[bundle-gate] OK: all ${Object.keys(budgets).length} web parts within budget`);
process.exit(0);
