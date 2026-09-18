# Vendored libraries

These files are committed instead of imported live from a CDN, because
dynamic ES module `import` statements have no browser-native Subresource
Integrity mechanism. A live `import 'https://cdn.../+esm'` would mean a
compromised or MITM'd CDN response could run unverified code with access
to the in-memory W-9 data (SSNs/EINs) this app processes. Vendoring +
`script-src 'self'` in the CSP closes that gap: the browser only ever runs
what's in this repo, and any tampering shows up as a `git diff`.

| File | Source | Fetched | SHA-512 |
|---|---|---|---|
| `pizzip-3.2.0.esm.js` | `https://cdn.jsdelivr.net/npm/pizzip@3.2.0/+esm` | 2026-08-18 (trailing `sourceMappingURL` comment stripped 2026-08-19, see below) | `7ef8925d62323ef72b7cadf547268e846fe2e08caf8fbc495736c6cf3f680c66401431052f7e39bfb746701845e6bf5172be59094af3c4225a995933529b1fa5` |
| `docx-8.5.0.esm.js` | `https://cdn.jsdelivr.net/npm/docx@8.5.0/+esm` | 2026-08-18 (trailing `sourceMappingURL` comment stripped 2026-08-19, see below) | `0cb09c29cca3a734e66d3bd96d00007731703703bb202d56fcf04ef1b3cb529aeed111096f0364e62f0e28e244831604c1289006e92de40847d6789471c87a0b` |
| `pako-2.1.0.esm.js` | `https://cdn.jsdelivr.net/npm/pako@2.1.0/dist/pako.es5.min.js/+esm` | 2026-08-18 (trailing `sourceMappingURL` comment stripped 2026-08-19, see below) | `88ff160f549dc2ba7405f9a9a0d18e7c780fd4d3ca25a9b053127da3036801f2f308d489144e2d73594c13a3403aa9d14a5ef121b525e4c4ee91e7b348c03b64` |
| `msal-browser-4.30.0.esm.js` | `https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.30.0/+esm` | 2026-08-18 (patched 2026-08-19 — dynamic-import fix + `sourceMappingURL` comment stripped; patched again 2026-08-20 — `@vite-ignore` added, see below) | `4b03477903ae9026811849f8afe97bed91990c0f9fa05338d4e8315fb387428467d11e68effa27384a2c65a236fc73801c5dce072444f5d79923726dbfd87bc4` |
| `msal-common-15.17.0.esm.js` | `https://cdn.jsdelivr.net/npm/@azure/msal-common@15.17.0/browser/+esm` | 2026-08-18 (trailing `sourceMappingURL` comment stripped 2026-08-19, see below) | `075a58f52bc99985c154b7a812ecdf1fd56e3b419dd8af6da664f38a5b3ee74a7d53e55b0fa4f53321d96dfb09e82aa16d3dd69d17e996a41fed8de8ff34db27` |

`pako` is pizzip's only dependency (zlib compression); `pizzip-3.2.0.esm.js`
imports it via the local relative path `./pako-2.1.0.esm.js`. `docx` has no
further dependencies, jsDelivr's `+esm` build bundles it into one file.

`msal-browser` (the MSAL.js library used for Entra ID sign-in, see
`auth.js`/`authConfig.js`) depends on `msal-common`; the same
relative-import rewrite was applied so `msal-browser-4.30.0.esm.js` imports
`./msal-common-15.17.0.esm.js` instead of a live jsDelivr URL. `4.30.0` is
the version jsDelivr's package metadata tags `"lts"` for
`@azure/msal-browser` (the newer `5.x` line is its current `"latest"`),
picked here for production stability over bleeding-edge features. Neither
file has any further `from"/npm/..."` **static** imports (checked after
the rewrite) — but see the dynamic-import patch below, which that check
did not catch.

### Cleanup: stray `sourceMappingURL` comments stripped (2026-08-19)

Every file in the table above originally ended with a `//# sourceMappingURL=/sm/<hash>.map` comment left over from jsDelivr's `+esm` bundler — pointing at a source map hosted on jsDelivr, not anything present in this repo. Under Vite's dev server this produced a harmless but noisy `ENOENT`-based warning on every startup (Vite tries to load the referenced map for its error-overlay/devtools support and fails since the file was never vendored). Since none of these maps are used by this app in any way, the comment line itself was simply deleted from all five files rather than vendoring five more never-referenced files. No functional code changed; only the SHA-512 values above (recomputed after the deletion).

### Patch: dead dynamic `import()` disarmed for Vite compatibility (2026-08-19)

Adopting Vite (see `package.json`, `AUTH.md`'s "NPM/Vite alternative")
surfaced one thing the `from"..."` grep above doesn't check for: a
**dynamic** `import("/npm/@azure/msal-browser@4.30.0/dist/telemetry/
BrowserPerformanceMeasurement.mjs/+esm")` buried inside MSAL's own
internal, undocumented performance-telemetry code path. It's gated behind
a hardcoded `sessionStorage` key MSAL checks internally — not exposed via
any public `msalConfig` option, and this app never sets that key — so this
was already dead code for us before Vite existed; it just didn't visibly
break anything under plain `<script type="module">` loading the way it
does under Vite's bundler/dev-transform, which both try to statically
resolve every `import()` call they see and fail hard on a live jsDelivr
URL that isn't a real local module.

Fixed by rewriting that one line's argument from a plain string literal to
a string concatenation (`""+"...same string..."`) — this defeats Vite
dev-server's literal-import-analysis pass, which is enough on its own for
`npm run dev`. `npm run build`'s bundler (rolldown) constant-folds the
concatenation back into a literal during its own optimization pass, so
`vite.config.js` additionally lists that exact resolved string under
`build.rollupOptions.external`, telling it to leave that one import
unresolved rather than fail the build. Net behavior is unchanged either
way — it's still exactly the same runtime-only dynamic import upstream
MSAL shipped, which our CSP (`script-src 'self'`) would block if it were
ever somehow reached, wrapped in MSAL's own `try/catch` so a blocked fetch
fails silently.

If re-vendoring a newer version, re-check for this same pattern (search for
`import(` — not just `from"` — and specifically for any `/npm/` or other
external-URL argument) and re-apply the same fix if it's still present.

### Patch: `@vite-ignore` added to silence the dev-server warning (2026-08-20)

The concatenation fix above stops Vite's dev-server transform from
attempting (and failing) to resolve the import, but it still *warns* about
it on every `npm run dev` startup ("The above dynamic import cannot be
analyzed by Vite") — harmless, but noisy. Since this import is confirmed
dead code for this app (see above) and deliberately left unresolvable on
purpose, a `/* @vite-ignore */` comment was added directly before the
`import(` call — this is Vite's own documented way to tell its
import-analysis pass "yes, I know, leave this one alone," rather than
something masking an unintended failure. No behavior change; the SHA-512
above reflects the file after both patches.

## `tesseract/` — Tesseract.js's own runtime fetches

By default, `Tesseract.createWorker()` fetches its OCR worker script, a WASM
core, and English language data from `cdn.jsdelivr.net` at runtime — picking
one of several SIMD/relaxed-SIMD/LSTM WASM core variants via feature
detection. None of that has SRI (dynamic `Worker`/`importScripts` loads can't
carry an `integrity` attribute), so it was the last unverified network
dependency in the app, sitting directly in the path of decoded W-9 images.

Closed by vendoring exactly the three files it needs and pointing
`workerPath`/`corePath`/`langPath` at them explicitly (`script.js`, in
`ocrPdfText`):

| File | Source | Fetched | SHA-512 |
|---|---|---|---|
| `tesseract/worker.min.js` | `https://cdn.jsdelivr.net/npm/tesseract.js@v7.0.0/dist/worker.min.js` | 2026-08-18 | `10ff8b6c62ce15b1960b568a015759f772c789b5864c44e6bfd5859ff70239c3db1d7fcac2e7ef481d0ca0aa9410e96083cd45b96ca23e5775b304d344256ba1` |
| `tesseract/tesseract-core-simd-lstm.wasm.js` | `https://cdn.jsdelivr.net/npm/tesseract.js-core@v7.0.0/tesseract-core-simd-lstm.wasm.js` | 2026-08-18 | `d324f65aec21f7bb514d3fd635d0f63085887bee8e90129b25a8cfa3067155e1ce693c2b14491126435d094ef9a3dd3314fbd87ad5599b4f223ba1871ecb3b36` |
| `tesseract/eng.traineddata.gz` | `https://cdn.jsdelivr.net/npm/@tesseract.js-data/eng/4.0.0_best_int/eng.traineddata.gz` | 2026-08-18 | `3f80cd71e6fea24df0fe3a88da0008d59b83362439eb0a94243d1afa57b09b933314d639bd65bbd121147f2538643067fa00f2efa943f24751cf6b7ee52adf7a` |

Two deliberate trade-offs:

- **`corePath` ends in `.js`**, not a bare directory. That's what tells
  Tesseract.js's internal `getCore.js` to load that exact file instead of
  running its SIMD/relaxed-SIMD feature detection and picking a CDN variant.
  We pinned the **SIMD, LSTM-only** core (`tesseract-core-simd-lstm`) rather
  than the plain non-SIMD build: WASM SIMD has near-universal support in
  evergreen browsers at this point, and the plain build measured ~6-8x
  slower in testing (a scanned W-9 page went from ~15-20s to ~2 minutes of
  OCR) — not worth it for the extra sliver of ancient-browser compatibility.
  The trade actually made is verifiability over relaxed-SIMD (the newest,
  fastest, least-supported variant): we vendor and hash exactly one file
  instead of Tesseract.js's usual six-way runtime pick.
- Confirmed self-contained: `tesseract-core-simd-lstm.wasm.js` embeds its
  WASM binary inline (no separate `.wasm` fetch — checked, no `.wasm` file
  reference in the bundle), so vendoring this one file is sufficient.
- `worker.min.js`'s only remaining CDN references are the *default* fallback
  URLs used when `corePath`/`langPath` aren't provided — since `script.js`
  always passes both explicitly, those fallbacks are dead code paths, not
  live network calls.

With this, `cdn.jsdelivr.net` is no longer referenced anywhere in the app —
the CSP no longer allow-lists it at all.

## Updating a version

1. Fetch the new `+esm` bundle from jsDelivr for the target version.
2. Check it for further `import ... from "/npm/..."` lines (grep for
   `from"`) — each one is another dependency that needs vendoring the same
   way, with its import path rewritten to a local relative path. **Also**
   grep for `import(` (dynamic imports) — `from"` alone misses these, and
   one shipped inside `msal-browser` pointing at a live jsDelivr URL; see
   the "Patch" note above for how that one was handled.
3. Record the new file's SHA-512 (`sha512sum <file>`) in the table above so
   future diffs are auditable.
4. Update the version in the `import` path in `script.js`.
