// Full end-to-end regression check for the wizard: upload a W-9, extract,
// insert data into the bundled MSA template, apply manual details, download,
// then verify the History view. Requires the app to be served over HTTP (ES
// modules are blocked under file://) — start it with `python -m http.server
// 8000` from the Document-Generator directory before running this.
//
// Usage: node tests/e2e.js
// Edit W9_PATH below to point at a real fixture file first.

import { chromium } from 'playwright';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUTPUT_DIR = path.join(__dirname, 'output');
const BASE = 'http://localhost:5173/index.html';
const W9_PATH = 'C:/Users/Lenovo/Downloads/W9 (1).pdf';

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

(async () => {
  const browser = await chromium.launch();
  const page = await (await browser.newContext()).newPage();
  const pageErrors = [];
  const consoleErrors = [];
  page.on('pageerror', (err) => pageErrors.push(err.stack || err.message));
  page.on('console', (msg) => { if (msg.type() === 'error') consoleErrors.push(msg.text()); });

  console.log('=== Load & dismiss intro ===');
  await page.goto(BASE);
  await page.waitForFunction(() => !!window.pdfjsLib, { timeout: 30000 });
  const elevator = await page.$('#elevator');
  if (elevator) await elevator.click({ force: true });
  await page.waitForTimeout(1500);

  console.log('=== Step 1: upload W-9, extract, insert data into bundled MSA template ===');
  await page.setInputFiles('#fileInput', W9_PATH);
  await page.click('#extractDataBtn');
  await page.waitForFunction(() => !document.getElementById('step1NextBtn').disabled, { timeout: 90000 });
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
  await page.waitForTimeout(500); // let saveToHistory finish

  console.log('=== Step 3: downloads ===');
  const [pdfDownload] = await Promise.all([
    page.waitForEvent('download', { timeout: 15000 }),
    page.click('#downloadPdfBtn'),
  ]);
  console.log('  PDF filename:', pdfDownload.suggestedFilename());
  await pdfDownload.saveAs(path.join(OUTPUT_DIR, 'e2e_out.pdf'));

  const [wordDownload] = await Promise.all([
    page.waitForEvent('download', { timeout: 15000 }),
    page.click('#downloadWordBtn'),
  ]);
  console.log('  Word filename:', wordDownload.suggestedFilename());
  await wordDownload.saveAs(path.join(OUTPUT_DIR, 'e2e_out.docx'));

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

  console.log('\n=== Errors ===');
  console.log('Page errors:', JSON.stringify(pageErrors, null, 2));
  console.log('Console errors:', JSON.stringify(consoleErrors, null, 2));

  await browser.close();
  process.exit(pageErrors.length || consoleErrors.length ? 1 : 0);
})();
