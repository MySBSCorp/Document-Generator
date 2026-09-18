#!/usr/bin/env node
// ---------------------------------------------------------------------------
// Verifies every file under vendor/ against the SHA-512 manifest recorded in
// vendor/README.md.
//
// Why this is a build gate and not a nicety: this app handles SSNs/EINs from
// W-9s entirely in the browser, and its whole security posture rests on
// `script-src 'self'` plus the claim that the vendored copies of pizzip,
// docx, pako, MSAL and Tesseract are byte-for-byte what vendor/README.md
// says they are. Dynamic ES module imports and Tesseract's Worker/WASM loads
// carry no Subresource Integrity, so nothing in the browser re-checks them.
// This script is the only place that check happens — so `npm run build`
// refuses to produce a dist/ if any vendored byte has drifted, and refuses
// to ship a vendor/ file that nobody has audited into the manifest.
//
// Manifest format parsed below (markdown tables in vendor/README.md):
//   | `path/relative/to/vendor` | source URL | fetched date | `<128 hex>` |
// ---------------------------------------------------------------------------

import { createHash } from 'node:crypto';
import { readFile, readdir, stat } from 'node:fs/promises';
import { dirname, join, posix, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const VENDOR_DIR = join(REPO_ROOT, 'vendor');
const MANIFEST = join(VENDOR_DIR, 'README.md');

// Files under vendor/ that are documentation, not shipped code, and so are
// deliberately absent from the hash manifest.
const NOT_CODE = new Set(['README.md']);

// A manifest row: 4 pipe-delimited cells, first backticked (the path) and
// last backticked (the 128-hex SHA-512). Source/date cells are free-form.
const ROW = /^\|\s*`([^`]+)`\s*\|[^|]*\|[^|]*\|\s*`([0-9a-fA-F]{128})`\s*\|\s*$/;

const red = (s) => `\x1b[31m${s}\x1b[0m`;
const green = (s) => `\x1b[32m${s}\x1b[0m`;
const yellow = (s) => `\x1b[33m${s}\x1b[0m`;
const dim = (s) => `\x1b[2m${s}\x1b[0m`;

async function parseManifest() {
  let text;
  try {
    text = await readFile(MANIFEST, 'utf8');
  } catch (err) {
    throw new Error(`cannot read the integrity manifest at vendor/README.md: ${err.message}`);
  }

  const expected = new Map();
  for (const line of text.split('\n')) {
    const match = ROW.exec(line.trim());
    if (!match) continue;
    const [, path, sha] = match;
    // Normalise so 'tesseract/worker.min.js' and './tesseract/worker.min.js'
    // are the same key regardless of how someone wrote the row.
    const key = posix.normalize(path.replace(/^\.\//, ''));
    if (expected.has(key) && expected.get(key) !== sha.toLowerCase()) {
      throw new Error(`vendor/README.md lists two different SHA-512 values for "${key}"`);
    }
    expected.set(key, sha.toLowerCase());
  }

  if (expected.size === 0) {
    throw new Error(
      'no SHA-512 rows found in vendor/README.md. Expected markdown table rows shaped like:\n' +
        '  | `pako-2.1.0.esm.js` | https://... | 2026-08-18 | `<128 hex chars>` |'
    );
  }
  return expected;
}

async function walk(dir, base = dir) {
  const out = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const abs = join(dir, entry.name);
    if (entry.isDirectory()) {
      out.push(...(await walk(abs, base)));
    } else if (entry.isFile()) {
      out.push(posix.normalize(relative(base, abs).split(/[\\/]/).join('/')));
    }
  }
  return out;
}

async function sha512(absPath) {
  // Vendored files top out around 4 MB, so a single read is fine and keeps
  // this script dependency- and stream-free.
  return createHash('sha512').update(await readFile(absPath)).digest('hex');
}

async function main() {
  try {
    await stat(VENDOR_DIR);
  } catch {
    console.error(red('FAIL') + ' vendor/ directory is missing. Nothing to verify.');
    process.exit(1);
  }

  const expected = await parseManifest();
  const present = (await walk(VENDOR_DIR)).filter((p) => !NOT_CODE.has(p));

  const ok = [];
  const mismatched = [];
  const missing = [];
  const unlisted = [];

  for (const [path, wantSha] of expected) {
    if (!present.includes(path)) {
      missing.push(path);
      continue;
    }
    const gotSha = await sha512(join(VENDOR_DIR, path));
    if (gotSha === wantSha) ok.push(path);
    else mismatched.push({ path, wantSha, gotSha });
  }

  for (const path of present) {
    if (!expected.has(path)) unlisted.push(path);
  }

  for (const path of ok) {
    console.log(`${green('ok')}   vendor/${path}`);
  }
  for (const { path, wantSha, gotSha } of mismatched) {
    console.error(`${red('FAIL')} vendor/${path} — SHA-512 does not match the manifest`);
    console.error(dim(`       expected ${wantSha}`));
    console.error(dim(`       actual   ${gotSha}`));
  }
  for (const path of missing) {
    console.error(`${red('FAIL')} vendor/${path} — listed in vendor/README.md but not on disk`);
  }
  for (const path of unlisted) {
    console.error(`${yellow('WARN')} vendor/${path} — on disk but NOT in the vendor/README.md manifest`);
  }

  const failures = mismatched.length + missing.length + unlisted.length;
  console.log('');
  if (failures === 0) {
    console.log(green(`vendor integrity OK — ${ok.length}/${expected.size} files match the manifest.`));
    return;
  }

  console.error(
    red(`vendor integrity FAILED — ${ok.length}/${expected.size} verified, ` +
      `${mismatched.length} mismatched, ${missing.length} missing, ${unlisted.length} unlisted.`)
  );
  console.error('');
  console.error('Do not deploy. Either:');
  console.error('  - restore the vendored files:  git checkout -- vendor/');
  console.error('  - or, if you intentionally updated a library, re-audit it and record the new');
  console.error('    hash in the vendor/README.md table (see its "Updating a version" section):');
  console.error('      shasum -a 512 vendor/<file>      # macOS');
  console.error('      sha512sum vendor/<file>          # Linux');
  process.exit(1);
}

main().catch((err) => {
  console.error(red('FAIL') + ' ' + err.message);
  process.exit(1);
});
