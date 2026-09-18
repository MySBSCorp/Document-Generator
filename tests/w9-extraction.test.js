import { chromium } from 'playwright';
import { createServer } from 'vite';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const TEST_DATA_DIR = path.join(ROOT, 'test_data');
const SNAPSHOT_FILE = path.join(__dirname, 'w9-extraction-snapshot.json');
const PORT = 5190;
const BASE = `http://localhost:${PORT}/index.html`;
const UPDATE_SNAPSHOT = process.env.UPDATE_SNAPSHOT === '1';
const FIELD_KEYS = ['companyName', 'taxId', 'taxIdType', 'address'];

async function startDevServer() {
  const server = await createServer({
    root: ROOT,
    mode: 'e2e',
    // This is a scripted, one-shot test run — it doesn't need live reload,
    // and any file change picked up mid-run (an editor autosave, a
    // formatter, another process touching the project while the suite is
    // in flight) would otherwise reload the page mid-extraction and corrupt
    // that fixture's result. Disabling HMR and the underlying watcher makes
    // the run immune to that regardless of what's causing the change.
    server: { port: PORT, strictPort: true, hmr: false, watch: null },
  });
  await server.listen();
  return server;
}

async function extractOne(browser, filePath) {
  const context = await browser.newContext();
  const page = await context.newPage();
  const pageErrors = [];
  page.on('pageerror', (err) => pageErrors.push(err.message));

  await page.goto(BASE);
  const elevator = await page.$('#elevator');
  if (elevator) await elevator.click({ force: true }).catch(() => {});
  await page.waitForFunction(() => !document.getElementById('introOverlay'), undefined, { timeout: 15000 });

  await page.setInputFiles('#fileInput', filePath);
  await page.waitForFunction(() => !document.getElementById('extractDataBtn').disabled, undefined, { timeout: 15000 });
  await page.click('#extractDataBtn');

  await page.waitForFunction(
    () => {
      const step2Active = document.querySelector('.wizard-step[data-step="2"]')?.classList.contains('active');
      const isError = document.getElementById('status')?.classList.contains('error');
      return step2Active || isError;
    },
    undefined,
    { timeout: 180000 }
  );

  const result = await page.evaluate(() => ({
    status: document.getElementById('status')?.textContent || '',
    isError: document.getElementById('status')?.classList.contains('error') || false,
    companyName: document.getElementById('w9CompanyName')?.value || '',
    taxId: document.getElementById('w9TaxId')?.value || '',
    taxIdType: document.getElementById('w9TaxIdType')?.value || '',
    address: document.getElementById('w9Address')?.value || '',
  }));
  result.pageErrors = pageErrors;

  await context.close();
  return result;
}

function objectivelyPassed(r) {
  return !r.isError && r.pageErrors.length === 0 && FIELD_KEYS.every((key) => !!r[key]);
}

function fieldsEqual(a, b) {
  return FIELD_KEYS.every((key) => a[key] === b[key]);
}

function snapshotOf(r) {
  return Object.fromEntries(FIELD_KEYS.map((key) => [key, r[key]]));
}

async function main() {
  if (!fs.existsSync(TEST_DATA_DIR)) {
    console.log(`No test_data/ directory found at ${TEST_DATA_DIR}`);
    process.exit(1);
  }

  const files = fs
    .readdirSync(TEST_DATA_DIR)
    .filter((f) => /\.(pdf|jpe?g|png|webp|heic|heif)$/i.test(f))
    .sort();

  if (!files.length) {
    console.log('No test documents found in test_data/.');
    return;
  }

  let baseline = {};
  if (fs.existsSync(SNAPSHOT_FILE)) {
    baseline = JSON.parse(fs.readFileSync(SNAPSHOT_FILE, 'utf8'));
  }
  const nextBaseline = { ...baseline };

  console.log(`Starting dev server (--mode e2e) on port ${PORT}...`);
  const server = await startDevServer();

  let browser;
  let exitCode = 0;

  try {
    browser = await chromium.launch();

    for (const file of files) {
      const filePath = path.join(TEST_DATA_DIR, file);
      process.stdout.write(`${file} ... `);

      let result;
      try {
        result = await extractOne(browser, filePath);
      } catch (err) {
        try {
          result = await extractOne(browser, filePath);
        } catch (retryErr) {
          console.log(`ERROR (${retryErr.message})`);
          exitCode = 1;
          continue;
        }
      }

      const passed = objectivelyPassed(result);
      const prior = baseline[file];

      if (!prior) {
        if (passed) {
          const snap = snapshotOf(result);
          console.log(
            `NEW, extraction succeeded — added to baseline (please eyeball once): ${JSON.stringify(snap)}`
          );
          nextBaseline[file] = snap;
        } else {
          console.log(`NEW, extraction FAILED: ${result.status}`);
          exitCode = 1;
        }
        continue;
      }

      if (!passed) {
        console.log(`REGRESSION — extraction failed: ${result.status}`);
        exitCode = 1;
        if (UPDATE_SNAPSHOT) nextBaseline[file] = snapshotOf(result);
        continue;
      }

      if (!fieldsEqual(prior, result)) {
        console.log('CHANGED vs baseline:');
        for (const key of FIELD_KEYS) {
          if (prior[key] !== result[key]) {
            console.log(`    ${key}: "${prior[key]}" -> "${result[key]}"`);
          }
        }
        exitCode = 1;
        if (UPDATE_SNAPSHOT) nextBaseline[file] = snapshotOf(result);
      } else {
        console.log('PASS (matches baseline)');
      }
    }
  } finally {
    if (browser) await browser.close();
    await server.close();
  }

  fs.writeFileSync(SNAPSHOT_FILE, JSON.stringify(nextBaseline, null, 2) + '\n');

  console.log('');
  if (exitCode === 0) {
    console.log(`All ${files.length} test document(s) match the baseline.`);
  } else {
    console.log(
      'One or more documents changed, failed, or are new — review the lines above.\n' +
        'If a change is intentional, re-run with UPDATE_SNAPSHOT=1 to accept it into the baseline.'
    );
  }
  process.exit(exitCode);
}

main();
