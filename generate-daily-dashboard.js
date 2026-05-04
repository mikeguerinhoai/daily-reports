#!/usr/bin/env node
/**
 * HOAi Daily Dashboard Generator
 *
 * Reads a daily data JSON payload + config, injects into dashboard-template.html,
 * writes a self-contained interactive HTML dashboard.
 *
 * Usage:
 *   node daily-reports/generate-daily-dashboard.js                    # Latest JSON in data/
 *   node daily-reports/generate-daily-dashboard.js --date 2026-04-05  # Specific date
 *   node daily-reports/generate-daily-dashboard.js --json path/to.json
 */

const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, 'data');
const CONFIG_PATH = path.join(ROOT, 'daily-report-config.json');
const TEMPLATE_PATH = path.join(ROOT, 'dashboard-template.html');
const OUTPUT_DIR = path.join(ROOT, 'output');

function findLatestJson() {
  if (!fs.existsSync(DATA_DIR)) return null;
  const files = fs.readdirSync(DATA_DIR)
    .filter(f => /^daily-report-\d{4}-\d{2}-\d{2}\.json$/.test(f))
    .sort()
    .reverse();
  return files.length ? path.join(DATA_DIR, files[0]) : null;
}

function main() {
  console.log('[generate-daily-dashboard] Starting...');

  // Parse args
  const args = process.argv.slice(2);
  let jsonPath;
  const dateIdx = args.indexOf('--date');
  const jsonIdx = args.indexOf('--json');

  if (jsonIdx >= 0 && args[jsonIdx + 1]) {
    jsonPath = path.resolve(args[jsonIdx + 1]);
  } else if (dateIdx >= 0 && args[dateIdx + 1]) {
    jsonPath = path.join(DATA_DIR, `daily-report-${args[dateIdx + 1]}.json`);
  } else {
    jsonPath = findLatestJson();
  }

  if (!jsonPath || !fs.existsSync(jsonPath)) {
    console.error(`ERROR: Data file not found: ${jsonPath || '(none in data/)'}. Run fetch-daily-data.py first.`);
    process.exit(1);
  }

  // 1. Read daily data JSON
  const dailyData = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
  const reportDate = dailyData.report_date || 'unknown';
  console.log(`  Data: ${path.basename(jsonPath)} (${reportDate})`);
  console.log(`  Voice companies: ${Object.keys(dailyData.voice || {}).length}`);
  console.log(`  SMS companies: ${Object.keys(dailyData.sms || {}).length}`);
  console.log(`  Webchat companies: ${Object.keys(dailyData.webchat || {}).length}`);
  console.log(`  Alerts: ${(dailyData.alerts || []).length}`);

  // 2. Read config
  if (!fs.existsSync(CONFIG_PATH)) {
    console.error(`ERROR: ${CONFIG_PATH} not found.`);
    process.exit(1);
  }
  const config = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));

  // 3. Merge: inject config subset into data payload
  const merged = {
    ...dailyData,
    config: {
      contracts: config.per_customer_contracts || {},
      utilization_thresholds: config.utilization_thresholds || {},
      voice_benchmarks: config.voice_benchmarks || {},
      sms_benchmarks: config.sms_benchmarks || {},
      webchat_benchmarks: config.webchat_benchmarks || {},
      brand: config.brand || {},
      trend_thresholds: config.trend_thresholds || {},
      voice_packages: config.voice_packages || {},
      sms_pricing_tiers: config.sms_pricing_tiers || {},
      per_customer_packages: config.per_customer_packages || {},
      time_savings: config.time_savings || {},
      test_companies: config.test_companies || []
    }
  };

  // 4. Read template
  if (!fs.existsSync(TEMPLATE_PATH)) {
    console.error(`ERROR: ${TEMPLATE_PATH} not found.`);
    process.exit(1);
  }
  let html = fs.readFileSync(TEMPLATE_PATH, 'utf8');

  // 5. Inject — escape </script> tags in data
  const dataJson = JSON.stringify(merged, null, 2).replace(/<\/script>/gi, '<\\/script>');
  html = html.replace('const DAILY_DATA = {};', `const DAILY_DATA = ${dataJson};`);

  // 6. Write output — nested folder structure: output/YYYY-MM/YYYY-MM-DD/Dashboard/
  const monthDir = reportDate.slice(0, 7); // "2026-04"
  const dashboardDir = path.join(OUTPUT_DIR, monthDir, reportDate, 'Dashboard');
  fs.mkdirSync(dashboardDir, { recursive: true });
  const outPath = path.join(dashboardDir, `HOAi_Daily_Dashboard_${reportDate}.html`);
  fs.writeFileSync(outPath, html, 'utf8');

  // Backward-compat flat copy
  const flatPath = path.join(OUTPUT_DIR, `HOAi_Daily_Dashboard_${reportDate}.html`);
  fs.copyFileSync(outPath, flatPath);

  const sizeMb = (Buffer.byteLength(html) / (1024 * 1024)).toFixed(2);
  console.log(`  Output: ${outPath} (${sizeMb} MB)`);
  console.log('  Done.');
}

main();
