// Captures a few representative mobile screenshots for docs/WORKFLOW.md's
// responsive-design section. Not part of the test suite.
// Usage: node tests/_mobile_screenshots.mjs
import { chromium } from 'playwright';
import { createServer } from 'vite';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const OUT_DIR = path.join(ROOT, 'docs', 'screenshots', 'mobile');
const PORT = 5852;
const BASE = `http://localhost:${PORT}/index.html`;
const W9_PATH = path.join(ROOT, 'test_data', 'w9_01_individual_clean.pdf');

fs.mkdirSync(OUT_DIR, { recursive: true });

async function shot(page, name) {
  await page.screenshot({ path: path.join(OUT_DIR, `${name}.png`) });
  console.log('captured', name);
}

(async () => {
  const server = await createServer({ root: ROOT, mode: 'e2e', server: { port: PORT, strictPort: true, hmr: false, watch: null } });
  await server.listen();
  const browser = await chromium.launch();
  // iPhone SE-ish viewport — the narrowest common phone width in real use.
  const page = await (await browser.newContext({ viewport: { width: 375, height: 667 }, isMobile: true, hasTouch: true })).newPage();

  await page.goto(BASE);
  const elevator = await page.$('#elevator');
  if (elevator) await elevator.click({ force: true }).catch(() => {});
  await page.waitForFunction(() => !document.getElementById('introOverlay'), undefined, { timeout: 15000 }).catch(() => {});
  await page.waitForTimeout(400);
  await shot(page, '01-step1');

  await page.setInputFiles('#fileInput', W9_PATH);
  await page.waitForFunction(() => !document.getElementById('extractDataBtn').disabled, undefined, { timeout: 15000 });
  await page.click('#extractDataBtn');
  await page.waitForFunction(() => document.querySelector('.wizard-step[data-step="2"]')?.classList.contains('active'), undefined, { timeout: 60000 });
  await page.waitForTimeout(500);
  await shot(page, '02-step1-review');

  await page.fill('#manualRep', 'Jordan Smith');
  await page.fill('#manualRole', 'Senior Consultant');
  await page.fill('#manualLocation', 'Remote - Austin, TX');
  await page.fill('#manualStartDate', 'September 1, 2026');
  await page.fill('#manualBillingRate', '95');
  await page.click('#applyManualBtn');
  await page.waitForFunction(() => document.querySelector('.wizard-step[data-step="3"]')?.classList.contains('active'), undefined, { timeout: 30000 });
  await page.waitForTimeout(600);
  await shot(page, '03-step3-download');

  await page.click('#navHistoryLink', { force: true });
  await page.waitForTimeout(500);
  await shot(page, '04-history');

  await browser.close();
  await server.close();
})();
