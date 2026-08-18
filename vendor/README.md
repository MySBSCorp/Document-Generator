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
| `pizzip-3.2.0.esm.js` | `https://cdn.jsdelivr.net/npm/pizzip@3.2.0/+esm` | 2026-08-18 | `ad81f9fe9eae74b1b95d7b3d64d0dedae65bff3602e4b74d953f26e8f5433baec91d4f8ede0763f7f917b013c6d4825ab93436ba9a8aba9843f096083d51124f` |
| `docx-8.5.0.esm.js` | `https://cdn.jsdelivr.net/npm/docx@8.5.0/+esm` | 2026-08-18 | `d0f5982a3578b6b95e32ba2be743d315be907ae1dbb9406b22f3b518d42ab6e14a7afd2ee3ff2b30ce6af8284c6ac6ef3cc058a0015c96e0f67253b7f9fcf995` |
| `pako-2.1.0.esm.js` | `https://cdn.jsdelivr.net/npm/pako@2.1.0/dist/pako.es5.min.js/+esm` | 2026-08-18 | `b57a2e43d7da112a662770fca5a4747e5edfb7db8eb599210ac0e94aa8b866f4af8185ce49a1541ae59ef4c5f14496432de05b7dfd7e8094cd1b99b7e53be6a1` |
| `msal-browser-4.30.0.esm.js` | `https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.30.0/+esm` | 2026-08-18 | `03192d6944de96b001a7a9a7a636f1065fbeb9664dc5881d5ba5da9a0743d649b1182d806eb0c9c142a37f5488a115581f9e7c083198630044dc7a89ccefc150` |
| `msal-common-15.17.0.esm.js` | `https://cdn.jsdelivr.net/npm/@azure/msal-common@15.17.0/browser/+esm` | 2026-08-18 | `fcdfe416216694362e6cecb74c2e5109e0dd0dd509618c5253e4c323729a7d4074fec422265b1123bd8ab24bcdcd99fc096ddb218ab6c2f64176ed961b08601a` |

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
file has any further `from"/npm/..."` imports (checked after the rewrite).

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
   `from"` ) — each one is another dependency that needs vendoring the same
   way, with its import path rewritten to a local relative path.
3. Record the new file's SHA-512 (`sha512sum <file>`) in the table above so
   future diffs are auditable.
4. Update the version in the `import` path in `script.js`.
