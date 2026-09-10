// Cross-browser regression check for the wizard: inject already-extracted
// W-9 data (skips the OCR pass — see __e2eInjectW9Data in script.js), insert
// into the bundled MSA template, apply manual details, download, then verify
// the History view. Runs headed (visible window) against Chrome, Edge,
// Firefox, and WebKit (Safari's engine — real Safari only runs on macOS, so
// this is the closest available substitute on Windows).
//
// Requires the app to be served over HTTP in e2e mode (auth bypass + the
// injection hook are both gated on VITE_E2E_BYPASS_AUTH) —
// start it with `npx vite --mode e2e` from the Document-Generator directory
// before running this.
//
// Usage: node tests/e2e-cross-browser.js [chrome|edge|firefox|webkit ...]
//   (no args = run all four, one after another)

import { chromium, firefox, webkit } from 'playwright';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUTPUT_DIR = path.join(__dirname, 'output');
const BASE = process.env.BASE_URL || 'http://localhost:5173/index.html';

const TARGETS = {
  chrome: () => chromium.launch({ headless: false, channel: 'chrome' }),
  edge: () => chromium.launch({ headless: false, channel: 'msedge' }),
  firefox: () => firefox.launch({ headless: false }),
  webkit: () => webkit.launch({ headless: false }),
};

function visibleStep(page) {
  return page.evaluate(() => {
    const active = document.querySelector('.wizard-step.active');
    return active ? active.dataset.step : null;
  });
}

async function waitForStep(page, step, timeoutMs) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if ((await visibleStep(page)) === step) return true;
    await page.waitForTimeout(300);
  }
  return false;
}

async function runOn(name) {
  console.log(`\n############ ${name.toUpperCase()} ############`);
  const browser = await TARGETS[name]();
  const page = await (await browser.newContext()).newPage();
  const pageErrors = [];
  const consoleErrors = [];
  const NOISE = /Estimating resolution as|Image too small to scale|Line cannot be recognized|willReadFrequently/;
  page.on('pageerror', (err) => pageErrors.push(err.stack || err.message));
  page.on('console', (msg) => {
    if (msg.type() === 'error' && !NOISE.test(msg.text())) consoleErrors.push(msg.text());
  });

  try {
    console.log('=== Load & dismiss intro ===');
    await page.addInitScript(() => sessionStorage.setItem('docgen_intro_reloaded', '1'));
    await page.goto(BASE);
    await page.waitForFunction(() => !!window.pdfjsLib, undefined, { timeout: 30000 });

    console.log('=== Step 1: inject W-9 data (skips OCR — see __e2eInjectW9Data), insert into bundled MSA template ===');
    await page.waitForFunction(() => typeof window.__e2eInjectW9Data === 'function', undefined, { timeout: 10000 });
    await page.evaluate(() => window.__e2eInjectW9Data({
      company_name: 'Acme Consulting LLC',
      tax_id_number: '12-3456789',
      selected_tax_id_type: 'EIN',
      ein_number: '123456789',
      street_address: '123 Main St',
      city: 'Springfield',
      state_zip: 'CA 94016',
    }));
    await page.waitForFunction(() => !document.getElementById('step1NextBtn').disabled, undefined, { timeout: 10000 });
    await page.click('#step1NextBtn');
    console.log('  reached step 2:', await waitForStep(page, '2', 30000));

    console.log('=== Step 2: manual details ===');
    await page.fill('#manualRep', 'Jordan Smith');
    await page.fill('#manualRole', 'Senior Consultant');
    await page.fill('#manualLocation', 'Remote - San Francisco, CA');
    await page.evaluate(() => {
      const picker = document.getElementById('manualStartDatePicker');
      picker.value = '2026-09-01';
      picker.dispatchEvent(new Event('change', { bubbles: true }));
    });
    await page.fill('#manualBillingRate', '125');
    await page.click('#applyManualBtn');
    console.log('  reached step 3:', await waitForStep(page, '3', 15000));
    await page.waitForTimeout(500);

    console.log('=== Step 3: downloads ===');
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    const [pdfDownload] = await Promise.all([
      page.waitForEvent('download', { timeout: 15000 }),
      page.click('#downloadPdfBtn'),
    ]);
    console.log('  PDF filename:', pdfDownload.suggestedFilename());
    await pdfDownload.saveAs(path.join(OUTPUT_DIR, `e2e_out_${name}.pdf`));

    const [wordDownload] = await Promise.all([
      page.waitForEvent('download', { timeout: 15000 }),
      page.click('#downloadWordBtn'),
    ]);
    console.log('  Word filename:', wordDownload.suggestedFilename());
    await wordDownload.saveAs(path.join(OUTPUT_DIR, `e2e_out_${name}.docx`));

    console.log('=== History ===');
    await page.click('#navHistoryLink');
    await page.waitForTimeout(500);
    const historyState = await page.evaluate(() => ({
      mainStageHidden: document.getElementById('mainStage').hidden,
      historyViewHidden: document.getElementById('historyView').hidden,
    }));
    console.log('  view state:', JSON.stringify(historyState));

    const items = await page.$$eval('.history-item', (els) =>
      els.map((el) => el.querySelector('.history-item-name')?.textContent));
    console.log('  entries:', JSON.stringify(items));

    const firstItem = await page.$('.history-item');
    if (firstItem) {
      const [historyDownload] = await Promise.all([
        page.waitForEvent('download', { timeout: 10000 }),
        firstItem.$eval('[data-action="pdf"]', (btn) => btn.click()),
      ]);
      console.log('  history PDF filename:', historyDownload.suggestedFilename());
      await firstItem.$eval('[data-action="delete"]', (btn) => btn.click());
      await page.waitForTimeout(400);
      console.log('  entries after delete:', await page.$$eval('.history-item', (els) => els.length));
    } else {
      console.log('  WARNING: no history item found');
    }

    await page.click('#navGeneratorLink');
    await page.waitForTimeout(300);
  } catch (err) {
    pageErrors.push(`RUN FAILED: ${err.stack || err.message}`);
  }

  console.log('\n--- Errors ---');
  console.log('Page errors:', JSON.stringify(pageErrors, null, 2));
  console.log('Console errors:', JSON.stringify(consoleErrors, null, 2));

  await browser.close();
  return pageErrors.length === 0 && consoleErrors.length === 0;
}

(async () => {
  const requested = process.argv.slice(2);
  const names = requested.length ? requested : Object.keys(TARGETS);
  const invalid = names.filter((n) => !TARGETS[n]);
  if (invalid.length) {
    console.error(`Unknown target(s): ${invalid.join(', ')}. Valid: ${Object.keys(TARGETS).join(', ')}`);
    process.exit(2);
  }

  const results = {};
  for (const name of names) {
    results[name] = await runOn(name);
  }

  console.log('\n============ SUMMARY ============');
  for (const [name, ok] of Object.entries(results)) {
    console.log(`  ${name.padEnd(8)} ${ok ? 'PASS' : 'FAIL'}`);
  }

  process.exit(Object.values(results).every(Boolean) ? 0 : 1);
})();
