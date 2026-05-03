#!/usr/bin/env node
/**
 * Per-web-part bundle size breach gate (Foundations Found.D7).
 * Reads release/assets/sp-search-*-web-part.js, compares against
 * config/bundle-budgets.json, exits non-zero on breach.
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
const breaches = [];
const report = { schemaVersion: 1, generatedAt: new Date().toISOString(), webParts: {} };

for (const [name, budget] of Object.entries(budgets)) {
  const file = path.join(ASSETS_DIR, name);
  if (!fs.existsSync(file)) {
    console.error(`[bundle-gate] missing asset: ${file} (run heft build --production first)`);
    process.exit(2);
  }
  const actual = fs.statSync(file).size;
  const delta = actual - budget;
  const pct = ((actual / budget) * 100).toFixed(1);
  const status = actual <= budget ? 'PASS' : 'BREACH';
  console.log(`[${status}] ${name}: ${actual.toLocaleString()} bytes (budget ${budget.toLocaleString()}, ${pct}%, delta ${delta >= 0 ? '+' : ''}${delta.toLocaleString()})`);
  report.webParts[name] = { actual, budget, delta, pctOfBudget: Number(pct), status };
  if (status === 'BREACH') breaches.push({ name, actual, budget, delta });
}

fs.mkdirSync(path.dirname(REPORT_PATH), { recursive: true });
fs.writeFileSync(REPORT_PATH, JSON.stringify(report, null, 2));
console.log(`[bundle-gate] report written: ${REPORT_PATH}`);

if (breaches.length > 0) {
  console.error(`\n[bundle-gate] FAILED: ${breaches.length} web part(s) exceed budget`);
  for (const b of breaches) {
    console.error(`  ${b.name}: ${b.actual.toLocaleString()} bytes exceeds ${b.budget.toLocaleString()} budget by ${b.delta.toLocaleString()} bytes`);
  }
  process.exit(1);
}

console.log(`\n[bundle-gate] OK: all ${Object.keys(budgets).length} web parts within budget`);
process.exit(0);
