#!/usr/bin/env node
// ---------------------------------------------------------------------------
// Minimal zero-dependency static server for local development.
//
// Two reasons this exists rather than `npx serve`:
//
//  1. This app cannot be opened over file:// at all — it loads script.js and
//     auth-ui.js as ES modules, which browsers block on file:// origins, and
//     MSAL needs a real http(s) origin to use as a redirect URI.
//  2. The MIME behaviour has to match the nginx vhost, specifically for
//     vendor/tesseract/eng.traineddata.gz. Tesseract.js fetches that file and
//     gunzips it *itself*. Any server that advertises `Content-Encoding: gzip`
//     on it makes the browser transparently decompress it first, and OCR then
//     fails on already-inflated bytes. This server serves .gz as opaque
//     bytes, exactly as deploy/nginx/docgen.conf does.
//
// Default port is 5500 to match .vscode/settings.json's Live Server config
// and the http://localhost:5500 redirect URI registered in Entra ID.
//
//   node scripts/dev-server.mjs [--port 5500] [--root .] [--host 127.0.0.1]
// ---------------------------------------------------------------------------

import { createReadStream } from 'node:fs';
import { stat } from 'node:fs/promises';
import { createServer } from 'node:http';
import { extname, join, normalize, resolve, sep } from 'node:path';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

function arg(name, fallback) {
  const i = process.argv.indexOf(`--${name}`);
  return i !== -1 && process.argv[i + 1] ? process.argv[i + 1] : fallback;
}

const PORT = Number(arg('port', 5500));
// 127.0.0.1, not 0.0.0.0: this serves an app that processes SSNs/EINs, so it
// should not be reachable from the local network by default.
const HOST = arg('host', '127.0.0.1');
const ROOT = resolve(REPO_ROOT, arg('root', '.'));

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.map': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.wasm': 'application/wasm',
  '.pdf': 'application/pdf',
  // Deliberately a plain byte stream, NOT Content-Encoding: gzip — see the
  // header comment. Tesseract.js inflates this itself.
  '.gz': 'application/gzip',
  '.traineddata': 'application/octet-stream',
};

const server = createServer(async (req, res) => {
  const started = process.hrtime.bigint();
  let status = 200;

  try {
    if (req.method !== 'GET' && req.method !== 'HEAD') {
      status = 405;
      res.writeHead(405, { 'Content-Type': 'text/plain', Allow: 'GET, HEAD' });
      return res.end('Method Not Allowed\n');
    }

    // Strip query (index.html uses ?v=... cache-busters) and decode.
    const rawPath = decodeURIComponent(new URL(req.url, `http://${req.headers.host}`).pathname);

    // Path traversal guard: normalise, then confirm the result is still
    // inside ROOT before touching the filesystem.
    let filePath = resolve(ROOT, '.' + normalize(rawPath));
    if (filePath !== ROOT && !filePath.startsWith(ROOT + sep)) {
      status = 403;
      res.writeHead(403, { 'Content-Type': 'text/plain' });
      return res.end('Forbidden\n');
    }

    let info;
    try {
      info = await stat(filePath);
    } catch {
      info = null;
    }
    if (info?.isDirectory()) {
      filePath = join(filePath, 'index.html');
      try {
        info = await stat(filePath);
      } catch {
        info = null;
      }
    }
    if (!info?.isFile()) {
      status = 404;
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      return res.end(`404 Not Found: ${rawPath}\n`);
    }

    const ext = extname(filePath).toLowerCase();
    const headers = {
      'Content-Type': MIME[ext] ?? 'application/octet-stream',
      'Content-Length': info.size,
      // No caching in dev, so an edit is always the thing you reload.
      'Cache-Control': 'no-store',
      'X-Content-Type-Options': 'nosniff',
    };

    res.writeHead(200, headers);
    if (req.method === 'HEAD') return res.end();
    createReadStream(filePath).pipe(res);
  } catch (err) {
    status = 500;
    if (!res.headersSent) res.writeHead(500, { 'Content-Type': 'text/plain' });
    res.end('Internal Server Error\n');
    console.error('  error:', err.message);
  } finally {
    res.on('close', () => {
      const ms = Number(process.hrtime.bigint() - started) / 1e6;
      console.log(`${String(status).padEnd(3)} ${req.method} ${req.url} ${ms.toFixed(1)}ms`);
    });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`Document Generator dev server`);
  console.log(`  root  ${ROOT}`);
  console.log(`  url   http://localhost:${PORT}/`);
  console.log('');
  console.log(`Sign-in note: http://localhost:${PORT} must be registered as a`);
  console.log(`"Single-page application" redirect URI on the Entra app registration,`);
  console.log(`or MSAL will reject the sign-in redirect. See AUTH.md, Section A.`);
  console.log('');
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`Port ${PORT} is already in use (VS Code Live Server, perhaps?).`);
    console.error(`Either stop that, or run: npm run dev -- --port 5502`);
    console.error(`If you change the port, register the new origin in Entra ID too.`);
    process.exit(1);
  }
  throw err;
});

for (const signal of ['SIGINT', 'SIGTERM']) {
  process.on(signal, () => {
    server.close(() => process.exit(0));
  });
}
