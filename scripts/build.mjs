#!/usr/bin/env node
// ---------------------------------------------------------------------------
// "Build" for a no-build app.
//
// There is no bundler, transpiler or minifier here on purpose (see AUTH.md /
// vendor/README.md — the browser runs exactly what is committed, so any
// tampering shows up in a git diff). What this script does instead is
// assemble an explicit ALLOW-LIST of the files the browser actually needs
// into dist/, so what gets rsynced to EC2 is only that.
//
// The allow-list is the point. An accidental `rsync` of the repo root would
// publish AUTH.md (24 KB describing the auth design), .env, tests/ (which
// contains absolute paths to real W-9 fixtures), and .git/ — all readable by
// anyone who guesses the URL. Copying forward a named set makes publishing a
// new file a deliberate act: add it here or it does not ship.
//
// `npm run build` runs `verify:vendor` first via the "prebuild" script, so a
// dist/ can never be produced from drifted vendor bytes.
// ---------------------------------------------------------------------------

import { execFileSync } from 'node:child_process';
import { cp, mkdir, rm, writeFile, stat, readdir } from 'node:fs/promises';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const DIST = join(REPO_ROOT, 'dist');

// Every file the browser requests, and nothing else. Keep in sync with the
// <script>/<link> tags in index.html and the `import` statements in
// script.js / auth-ui.js.
const FILES = [
  'index.html',
  'style.css',
  'script.js',
  'auth-ui.js',
  'auth.js',
  'authConfig.js',
];

// Whole directories copied as-is. vendor/ is fetched at runtime by
// script.js's imports and by Tesseract's workerPath/corePath/langPath.
const DIRS = ['vendor'];

// Documentation that lives inside a shipped directory but should not be
// served. vendor/README.md is the integrity manifest — useful in git, not on
// the public web.
const EXCLUDE_FROM_DIRS = [/(^|\/)README\.md$/i, /(^|\/)\.DS_Store$/];

const green = (s) => `\x1b[32m${s}\x1b[0m`;
const red = (s) => `\x1b[31m${s}\x1b[0m`;
const dim = (s) => `\x1b[2m${s}\x1b[0m`;

function git(...args) {
  try {
    return execFileSync('git', args, { cwd: REPO_ROOT, encoding: 'utf8' }).trim();
  } catch {
    return 'unknown';
  }
}

async function assertExists(relPath, kind) {
  try {
    const info = await stat(join(REPO_ROOT, relPath));
    if (kind === 'dir' && !info.isDirectory()) throw new Error('not a directory');
    if (kind === 'file' && !info.isFile()) throw new Error('not a file');
  } catch (err) {
    throw new Error(`build allow-list references ${relPath}, but it is missing or wrong (${err.message})`);
  }
}

async function walk(dir, base) {
  const out = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const abs = join(dir, entry.name);
    if (entry.isDirectory()) out.push(...(await walk(abs, base)));
    else if (entry.isFile()) out.push(relative(base, abs).split(/[\\/]/).join('/'));
  }
  return out;
}

async function main() {
  for (const f of FILES) await assertExists(f, 'file');
  for (const d of DIRS) await assertExists(d, 'dir');

  await rm(DIST, { recursive: true, force: true });
  await mkdir(DIST, { recursive: true });

  let count = 0;
  let bytes = 0;

  for (const f of FILES) {
    await cp(join(REPO_ROOT, f), join(DIST, f));
    bytes += (await stat(join(DIST, f))).size;
    count += 1;
  }

  for (const d of DIRS) {
    await cp(join(REPO_ROOT, d), join(DIST, d), {
      recursive: true,
      filter: (src) => {
        const rel = relative(REPO_ROOT, src).split(/[\\/]/).join('/');
        return !EXCLUDE_FROM_DIRS.some((re) => re.test(rel));
      },
    });
    for (const f of await walk(join(DIST, d), DIST)) {
      bytes += (await stat(join(DIST, f))).size;
      count += 1;
    }
  }

  // Dotfile on purpose: the nginx vhost denies /\. so this is readable by ops
  // over SSH (`cat /var/www/docgen/current/.build-info.json`) to confirm which
  // commit is live, without exposing the commit SHA on the public web.
  const info = {
    name: 'document-generator',
    builtAt: new Date().toISOString(),
    gitCommit: git('rev-parse', 'HEAD'),
    gitBranch: git('rev-parse', '--abbrev-ref', 'HEAD'),
    gitDirty: git('status', '--porcelain') !== '',
    files: count,
    bytes,
  };
  await writeFile(join(DIST, '.build-info.json'), JSON.stringify(info, null, 2) + '\n');

  console.log(green(`built dist/ — ${count} files, ${(bytes / 1024 / 1024).toFixed(2)} MB`));
  console.log(dim(`  commit ${info.gitCommit.slice(0, 12)} (${info.gitBranch})${info.gitDirty ? ' + uncommitted changes' : ''}`));
  if (info.gitDirty) {
    console.log(dim('  note: working tree is dirty — dist/ does not match any commit.'));
  }
  console.log(dim('  next: npm run preview   (serve dist/ locally)   or   npm run deploy'));
}

main().catch((err) => {
  console.error(red('build failed: ') + err.message);
  process.exit(1);
});
