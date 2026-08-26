// One-off script to capture workflow screenshots for docs/WORKFLOW.md.
// Not part of the test suite. Usage: node tests/_screenshots.mjs
import { chromium } from 'playwright';
import { createServer } from 'vite';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const OUT_DIR = path.join(ROOT, 'docs', 'screenshots');
const PORT = 5191;
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
  const page = await (await browser.newContext({ viewport: { width: 1440, height: 900 } })).newPage();
  page.on('pageerror', (e) => console.log('PAGE ERROR', e.message));

  try {
    await page.goto(BASE);
    await page.waitForFunction(() => !document.getElementById('introOverlay')?.hasAttribute('inert') || true, undefined, { timeout: 15000 }).catch(() => {});
    await page.waitForTimeout(500);
    await shot(page, '01-signed-in-step1-empty');

    const elevator = await page.$('#elevator');
    if (elevator) await elevator.click({ force: true }).catch(() => {});
    await page.waitForFunction(() => !document.getElementById('introOverlay'), undefined, { timeout: 15000 }).catch(() => {});
    await page.waitForTimeout(400);
    await shot(page, '02-step1-upload');

    await page.setInputFiles('#fileInput', W9_PATH);
    await page.waitForFunction(() => !document.getElementById('extractDataBtn').disabled, undefined, { timeout: 15000 });
    await shot(page, '03-step1-file-selected');

    await page.click('#extractDataBtn');
    await page.waitForFunction(
      () => document.querySelector('.wizard-step[data-step="2"]')?.classList.contains('active') ||
            document.getElementById('status')?.classList.contains('error'),
      undefined, { timeout: 120000 }
    );
    await page.waitForTimeout(600);
    await shot(page, '04-step1-extracted-review');

    // Open the W-9 preview tab to capture the side panel.
    const previewToggle = await page.$('#previewToggleBtn');
    if (previewToggle) {
      await previewToggle.click().catch(() => {});
      await page.waitForTimeout(400);
      const w9Tab = await page.$('#previewTabW9');
      if (w9Tab) await w9Tab.click().catch(() => {});
      await page.waitForTimeout(1200);
      await shot(page, '05-preview-panel-w9');
      const genTab = await page.$('#previewTabGenerated');
      if (genTab) await genTab.click().catch(() => {});
      await previewToggle.click().catch(() => {});
      await page.waitForTimeout(300);
    }

    const nextBtn = await page.$('#step1NextBtn');
    if (nextBtn) await nextBtn.click().catch(() => {});
    await page.waitForFunction(() => document.querySelector('.wizard-step[data-step="2"]')?.classList.contains('active'), undefined, { timeout: 15000 }).catch(() => {});
    await page.waitForTimeout(400);
    await shot(page, '06-step2-manual-details-empty');

    await page.fill('#manualRep', 'Jordan Smith').catch(() => {});
    await page.fill('#manualRole', 'Senior GIS Consultant').catch(() => {});
    await page.fill('#manualLocation', 'Remote - Austin, TX').catch(() => {});
    await page.fill('#manualStartDate', 'September 1, 2026').catch(() => {});
    await page.fill('#manualBillingRate', '95').catch(() => {});
    await shot(page, '07-step2-manual-details-filled');

    const applyBtn = await page.$('#applyManualBtn');
    if (applyBtn) await applyBtn.click().catch(() => {});
    await page.waitForFunction(() => document.querySelector('.wizard-step[data-step="3"]')?.classList.contains('active'), undefined, { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(800);
    await shot(page, '08-step3-download');

    const previewToggle2 = await page.$('#previewToggleBtn');
    if (previewToggle2) {
      await previewToggle2.click().catch(() => {});
      await page.waitForTimeout(1500);
      await shot(page, '09-preview-panel-generated-msa');
      // Close the preview panel so it doesn't cover the History view.
      await previewToggle2.click().catch(() => {});
      await page.waitForTimeout(400);
    }

    // History view
    const historyLink = await page.$('#navHistoryLink');
    if (historyLink) {
      await historyLink.click({ force: true }).catch(() => {});
      await page.waitForTimeout(800);
      await shot(page, '10-history-view');
    }
  } catch (err) {
    console.log('SCRIPT ERROR', err.message);
    await shot(page, '99-error-state');
  } finally {
    await browser.close();
    await server.close();
  }
})();
