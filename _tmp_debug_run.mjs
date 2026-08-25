import { chromium } from 'playwright';
import { spawn } from 'child_process';
import path from 'path';

const file = process.argv[2];

const vite = spawn('npx', ['vite', '--mode', 'e2e', '--port', '5202'], { shell: true });
await new Promise((r) => setTimeout(r, 4000));

const browser = await chromium.launch();
const page = await browser.newPage();
page.on('console', (msg) => {
  const t = msg.text();
  if (t.includes('TMP-DEBUG')) console.log(t);
});
await page.goto('http://localhost:5202');
await page.waitForFunction(() => !document.getElementById('introOverlay'), undefined, { timeout: 15000 });
await page.setInputFiles('#fileInput', path.resolve(file));
await page.click('#extractDataBtn', { timeout: 15000 });
await page.waitForFunction(
  () => {
    const s = document.getElementById('status');
    return s && /parsed|error|doesn.t look|could not read|no SSN or EIN/i.test(s.textContent || '');
  },
  undefined,
  { timeout: 120000 }
).catch(() => console.log('[timeout]'));
const status = await page.locator('#status').textContent();
const address = await page.locator('#w9Address').inputValue().catch(() => '');
const company = await page.locator('#w9CompanyName').inputValue().catch(() => '');
const taxId = await page.locator('#w9TaxId').inputValue().catch(() => '');
console.log('STATUS:', status, '| TAXID:', taxId, '| ADDRESS:', address, '| COMPANY:', company);
await browser.close();
vite.kill();
process.exit(0);
