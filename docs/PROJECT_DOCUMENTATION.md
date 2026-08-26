# Project Documentation

Full reference documentation for the Document Generator product: what it is, how it's built, and how every piece fits together. For the auth-specific deep dive see [AUTH.md](../AUTH.md); for requirements see [REQUIREMENTS.md](REQUIREMENTS.md); for the step-by-step usage/dev/deploy flow (with screenshots) see [WORKFLOW.md](WORKFLOW.md).

## 1. Overview

Document Generator is a **frontend-only, no-backend** web app that turns an uploaded W-9 tax form into a filled Master Service Agreement (MSA). A user signs in with an approved Microsoft Entra ID account, uploads a W-9, reviews the extracted data, fills in a few extra manual details, and downloads the finished agreement as both PDF and Word — all processed entirely in the browser, with nothing sent to any server. The MSA template itself is a single bundled document, not something the user supplies.

## 2. Project structure

```
Document-Generator/
├── index.html              # markup: auth gate, intro overlay, nav, wizard (3 steps), preview panel, history view
├── style.css                # all styling
├── script.js                # app logic: extraction, generation, preview, history (~3350 lines)
├── authConfig.js             # public MSAL/Entra config, reads Vite env vars (import.meta.env)
├── auth.js                   # MSAL controller (imports vendored MSAL, not npm)
├── auth-ui.js                 # DOM wiring for the auth gate; also handles the E2E test bypass
├── package.json               # real npm project: dev/build/preview/test scripts, real dependencies
├── vite.config.js              # Vite build config
├── .env.example / .env.e2e     # documents VITE_MSAL_* build-time vars and the E2E bypass flag
├── vendor/                     # locally vendored ES modules (see vendor/README.md for why)
│   ├── msal-browser-4.30.0.esm.js, msal-common-15.17.0.esm.js
│   ├── pizzip-3.2.0.esm.js, pako-2.1.0.esm.js
│   ├── docx-8.5.0.esm.js         # vendored but currently unused — see Section 4
│   ├── tesseract/                # OCR worker/core/language data
│   └── README.md                 # source URLs + SHA-512 for every vendored file
├── public/
│   └── assets/
│       ├── msa-template.docx      # the one bundled MSA template, fetched at runtime
│       └── logo.jpg                # SbS Beyond IT logo, shown in the nav (extracted from the template's own embedded image)
├── tests/                      # Playwright end-to-end + extraction-snapshot tests, screenshot script
├── AUTH.md                     # Authentication setup, security rationale, checklist (repo root)
├── docs/
│   ├── WORKFLOW.md                # Usage / dev / test / deploy workflow, with screenshots
│   ├── REQUIREMENTS.md            # Functional & non-functional requirements
│   ├── PROJECT_DOCUMENTATION.md    # this file
│   └── screenshots/                # workflow screenshots referenced by WORKFLOW.md
```

The project now has a **real build step**: `npm install` + `npm run dev` (Vite dev server) or `npm run build` (production bundle). This replaced an earlier no-build, CDN-`<script>`-tag setup — see Section 4 for exactly what changed and why some dependencies are still vendored rather than pulled from npm.

## 3. Architecture

```
                         ┌─────────────────────────┐
                         │      Browser tab         │
                         │                          │
   Entra ID  ◄───────────┤  auth.js / auth-ui.js /  │
  (sign-in,               │   authConfig.js          │
   sign-out,               │      (MSAL.js)           │
   "who's allowed")        └────────────┬─────────────┘
                                         │ gates
                                         ▼
                         ┌─────────────────────────┐
                         │       script.js          │
                         │  (wizard controller,      │
                         │   extraction, generation,  │
                         │   live preview, history)    │
                         └───┬─────────┬─────────┬───┘
                             │         │         │
                    pdf.js   │  Tesseract.js │  pdf-lib / docx-preview /
                 (text/render)│    (OCR)      │  html2canvas / PizZip
                             │         │         │  (preview + generation)
                             ▼         ▼         ▼
                     Everything above runs in-browser.
                     IndexedDB (History) and sessionStorage
                     (MSAL cache, history session ID) are the
                     only persistence — no network calls except
                     to Entra ID and one same-origin fetch for
                     the bundled MSA template.
```

There is no server tier at all. The three architectural layers are:
1. **Auth layer** (`authConfig.js`, `auth.js`, `auth-ui.js`) — gates everything else behind Entra ID sign-in. See Section 5 and [AUTH.md](../AUTH.md).
2. **Application layer** (`script.js`) — the wizard, extraction, live preview, and generation logic. See Section 6.
3. **Presentation layer** (`index.html`, `style.css`) — markup and styling for both layers above.

## 4. Technology stack

| Concern | Technology | How it's loaded |
|---|---|---|
| Build tool | Vite 8 | `npm run dev` / `npm run build` / `npm run preview` — a real build step, not optional |
| Language | Plain JavaScript (ES modules), HTML, CSS | No TypeScript |
| Authentication | Microsoft Entra ID via `@azure/msal-browser` 4.30.0 + `@azure/msal-common` | **Vendored** locally at `vendor/*.esm.js`, imported by relative path in `auth.js` — not via npm, despite being listed in `package.json` |
| PDF text extraction & rendering | `pdfjs-dist` ^3.11.174 | npm bare specifier, resolved by Vite; its worker is loaded via a Vite `?url` asset import |
| OCR | `tesseract.js` ^7.0.0 | npm bare specifier for the library itself, but its worker/core/language files are still **vendored locally** at `vendor/tesseract/*` |
| PDF generation/assembly | `pdf-lib` ^1.17.1 | npm bare specifier |
| DOCX text replacement | `pizzip` ^3.2.0 (+ `pako` ^2.1.0 dependency) | **Vendored**, relative import at the top of `script.js` — text substitution is done via raw XML manipulation of `word/document.xml`, not a DOCX-authoring library |
| Live DOCX preview rendering | `docx-preview` ^0.4.0 | npm bare specifier — renders the generated `.docx` straight into an iframe |
| DOCX → PDF rasterization | `html2canvas` ^1.4.1 | npm bare specifier — screenshots the rendered DOCX pages so they can be assembled into a PDF |
| Local persistence | IndexedDB (history blobs), `sessionStorage` (MSAL cache, history session scoping) | No `localStorage` use for sensitive data |
| Testing | Playwright (`tests/e2e.js`, `tests/w9-extraction.test.js`) | `npm test` / `npm run test:extraction` |
| Hosting | Any static HTTP(S) host, after `npm run build` | No server-side runtime |

**Note on `docx` (the npm package, separate from `docx-preview`)**: it's listed in `package.json` and vendored at `vendor/docx-8.5.0.esm.js`, but nothing in the codebase currently imports it — DOCX editing is done entirely by PizZip + raw XML manipulation. Treat it as a stale dependency until it's either wired up or removed.

**Why some dependencies are vendored instead of pulled straight from npm**: dynamic ES `import` has no browser-native equivalent of Subresource Integrity, so MSAL, PizZip, and pako are vendored locally with their SHA-512 recorded in [vendor/README.md](../vendor/README.md), and the CSP's `script-src 'self'` means only same-origin code can execute at all. Tesseract's worker/core/language files are vendored for the same reason (a dynamically-created `Worker` can't carry an `integrity` attribute either). Everything else (`pdfjs-dist`, `pdf-lib`, `docx-preview`, `html2canvas`) is bundled by Vite at build time, so it ships as part of the same integrity-checked, same-origin bundle rather than being fetched live from anywhere.

## 5. Authentication layer

Covered in full in [AUTH.md](../AUTH.md) — summary here for context:

- **Protocol**: OAuth 2.0 Authorization Code Flow with PKCE, via MSAL.js. No implicit flow, no client secret (this is a public SPA client).
- **Flow**: `loginRedirect`/`logoutRedirect` (not popup — Microsoft's own Cross-Origin-Opener-Policy header breaks popup-completion detection).
- **Scopes**: `openid`, `profile`, `email` only — no Microsoft Graph or other API access is ever requested or held.
- **Who's allowed**: primarily enforced by Entra ID's own "Assignment required" Enterprise Application setting. `authConfig.js` additionally supports an **optional** frontend allow-list (`VITE_ALLOWED_EMAILS`), checked by `auth.js`'s `isEmailAllowed()` — but this is a convenience layer on top of the server-side gate, not a replacement for it; anything in frontend code is always readable by anyone who opens devtools.
- **Session cache**: `sessionStorage` (not `localStorage`) — tokens/account data don't persist once the tab closes.
- **The gate** (`#authGate`): a fixed, full-viewport overlay, backed by the native `inert` attribute on everything behind it (`#introOverlay`, `<nav>`, `#historyView`, `#mainStage`), not just CSS `z-index` — this closes a confirmed keyboard-focus bypass. `inert` is only lifted once `auth-ui.js`'s `showAuthenticated()` actually runs.
- **Blocked sign-ins**: if Entra redirects back with an error, `auth.js` throws `AccessDeniedError` and `auth-ui.js` shows a pop-up warning with a "Retry" that returns to the sign-in screen.
- **Dev/test-only bypass**: when `import.meta.env.DEV` is true (never in a production build) and `VITE_E2E_BYPASS_AUTH=true` is set, `auth-ui.js` skips MSAL entirely and signs in a synthetic `E2E Test User`. Used exclusively by the Playwright test suite and the screenshot script — never present in a production build.

## 6. Application layer (`script.js`)

### 6.1 Wizard controller
**Three steps** (`data-step="1"`, `"2"`, `"3"`), each gated on the previous step's required data being valid:

1. **Upload files** — upload the W-9, click "Extract Data." Extraction success auto-advances into Step 2 and also fills the bundled MSA template with the extracted data behind the scenes.
2. **Manual details** — fields that aren't on a W-9 at all (rep, role, location, start date, billing rate); "Apply Details" finalizes both output documents.
3. **Download** — PDF and Word downloads, plus "Generate New Document" to reset the whole wizard.

There is **no separate "upload MSA template" step** — see Section 6.3.

A minimum-loading-time guard (`waitRemaining`) ensures fast operations still show a spinner briefly rather than flashing, and a decorative intro animation plays once on load (purely cosmetic, `inert` until sign-in succeeds).

### 6.2 W-9 extraction pipeline (Step 1)

1. **Text-layer extraction first** (`extractPdfText`): pdf.js `getTextContent()` plus AcroForm annotation values. If the PDF's fields match the known IRS fillable-W-9 field-name pattern, `extractAcroFormDirectFields` reads values directly by field number (highest-confidence tier) instead of via regex.
2. **OCR fallback** (`ocrPdfText`/`ocrImageFile`) — used whenever there's no usable text layer, the file is an image, or the text doesn't pass a W-9 sanity check:
   - Renders the page to canvas (up to 2 pages for a PDF), then **deskews** it (`deskewCanvas` tries −8°…+8° and keeps whichever angle makes text rows sharpest), then grayscales/contrast-stretches.
   - One full-page OCR pass returns both plain text and structured line/word bounding boxes.
   - **Dynamic label anchoring** (`findTinCropRegion`/`findAddressCropRegion`/`findAnchorLine`/`fuzzyContainsLabel`): searches the whole page's OCR'd lines for the actual SSN/EIN/Address labels (tolerant of OCR typos via a Levenshtein-based fuzzy match), and derives a crop region from wherever they really are — so a shifted, cropped, or rotated scan is still handled. A fixed fraction-of-page-size crop region is kept only as a fallback when no label can be found.
   - **Targeted digit-box OCR** on that crop region, with several defenses layered on top of plain OCR:
     - Ink-density gating rejects a box with too little ink (avoids reading an empty box's own printed border as digits).
     - `whiteOutGridLines` paints over box-cell border strokes before OCR (a border fusing into a digit was a real source of dropped digits).
     - `stretchContrast` improves low-contrast/dim photo scans (skipped on crops that are already high-contrast, to avoid flipping a legible digit into a misread one).
     - **Digit-consensus voting**: up to 12 OCR attempts per box (multiple scales × preprocessing variants × crop widths), only accepting a value with majority agreement; a genuinely ambiguous result is returned as "unknown" rather than guessed, which deliberately blocks the cruder full-text fallback from filling in a wrong answer.
   - Address box OCR follows the same dynamic-anchor-then-fallback crop strategy.
3. **Field parsing** (`parseW9Fields`): combines the AcroForm, digit-box, and address-box results (in that precedence order) with a regex-based full-text fallback (tolerant of OCR noise via flexible whitespace/punctuation matching) for anything not already resolved. Extracts company name, tax ID (SSN or EIN), and address.
4. Every extracted field is shown in an editable summary and re-validated; the user must have at least a tax ID present to proceed. A `w9_debug` localStorage flag (off by default) enables a debug panel with crop-region overlays and confidence scores, for troubleshooting without leaking sensitive data in normal use.

### 6.3 MSA template

The MSA template is a **single bundled document** (`public/assets/msa-template.docx`), fetched once via a same-origin `fetch()` the first time it's needed and cached in memory for the rest of the session (`loadMsaTemplates`/`loadFileAsset`). There is no user-facing template upload anymore, and no placeholder-detection gate to click through — the template is simply assumed correct, and the extracted W-9 data is inserted into it automatically as part of finishing Step 1.

### 6.4 Document generation (Step 2 → Step 3)

Real-world MSA templates aren't fillable forms or merge-tag documents — they're ordinary contracts where the previous contractor's details appear as repeated plain text. So generation is literal find-and-replace, not template-tag substitution, and DOCX is the single source of truth:

- **DOCX replacement** (`applyDocxReplacements`): unzips the `.docx` with PizZip, parses `word/document.xml`, and replaces matching text across `<w:t>` runs — including matches that span multiple runs (common when Word splits a phrase across differently-formatted runs) — plus a separate pass that inserts each manual field's value right after its label.
- **Live preview**: the in-progress DOCX is rendered straight into an iframe via `docx-preview`, updated on a debounce after every relevant keystroke, with independent zoom/pan for the preview panel.
- **PDF generation** (`renderDocxToPdf`): the *same*, already-fully-replaced DOCX is rendered into an off-screen container via `docx-preview`, each page is rasterized to a PNG with `html2canvas`, and those images are assembled into a PDF with `pdf-lib`. This means the PDF is always a flattened image of the final document — the previous contractor's data can never remain present and extractable/copy-pasteable underneath it, and both the PDF and Word downloads are always full, real generations of the same content (never a fallback summary document).

### 6.5 History
- **Storage**: IndexedDB (chosen over `localStorage` because it can store Blobs directly and isn't capped at ~5MB) — stores the actual generated PDF and Word blobs, not just metadata.
- **Scoping & retention**: each browser tab gets a random session ID (`crypto.randomUUID()` in `sessionStorage`). Every read/write is filtered to the current tab's session, and entries from any *other* (closed) session are purged unconditionally both on every page load and every time the History view is opened — specifically so a closed tab's SSN/EIN-bearing documents don't linger indefinitely. Manual per-entry deletion is also available.

## 7. Presentation layer

- `index.html` holds all markup: the auth gate and its states, the blocked-sign-in pop-up, the nav, the 3-step wizard, the live Document Preview side panel, and the history view.
- **Branding**: the nav's `.brand` element is the SbS Beyond IT logo (`public/assets/logo.jpg`) plus a `.brand-text` label, side by side. The logo file was extracted directly from the bundled MSA template's own embedded image (`word/media/image1.jpg` inside `msa-template.docx`) — it's the same artwork the generated documents themselves carry, so the app and its output stay visually consistent by construction rather than by two separately-maintained copies.
- **Document Preview panel** (`#previewPanel`): a toggleable side panel with two tabs.
  - *Generated MSA* — the live `docx-preview`-rendered iframe described in 6.4, with cursor-anchored zoom (30%–200%) and click-drag panning implemented directly against the iframe document.
  - *Uploaded W-9* — shown the moment a file is selected, before extraction even runs: images are shown directly via an object URL; PDFs are rendered page-by-page to canvases with pdf.js (deliberately not an `<iframe src="blob:...pdf">`, since a browser's built-in PDF viewer plugin isn't guaranteed to be available). Has its own independent zoom/pan state.
- `style.css` holds all styling, including the auth gate's intentionally flat design and the preview panel's zoom/pan chrome.
- **Responsive layout**: the desktop split-screen layout (fixed-height, two independently-scrolling panes side by side) is desktop-only, and is deliberately *not* reused as-is on mobile — a dual-clipped-pane layout stacked full-width on a phone just cuts content off partway down each pane with no way to see the rest. Below 900px width (`@media (max-width: 900px)`) the app switches to a normal, naturally-scrolling single column instead, with the nav pinned via `position: sticky` so it stays reachable while a long form scrolls underneath it. Further breakpoints refine things down to small phones (tested to ~320–430px, e.g. an iPhone SE or a Galaxy S21):
  - **640px** — nav padding/gaps tighten, the W-9 summary's label/value pairs and each History entry's name/date/buttons switch from a cramped single row to a stacked layout, and the intro overlay's decorative elevator scene (a fixed 600×400px canvas) gets scaled down with a CSS `transform` so it stays fully on-screen instead of clipping past the viewport edge.
  - **480px** — the "Document Generator" wordmark drops from the nav, keeping just the logo, so the nav links and the signed-in badge always fit on one row instead of wrapping into a ragged second line.
  - A few elements are viewport-relative rather than breakpoint-based: the wizard's "Extracting data…" loading overlay covers the full viewport on mobile (`window.innerWidth <= 900` check in `positionWizardLoading()`) instead of a snapshot of the action panel's on-screen position, since that position is no longer stable once the page itself can scroll; and sign-in/logout confirmation modals were already `width: 100%; max-width: 360px` with padding on their overlay, so they needed no changes to fit any of these widths.
- A Content-Security-Policy `<meta>` tag restricts `script-src` to `'self'` (plus `'wasm-unsafe-eval'` for Tesseract's WASM), and `connect-src`/`frame-src` to Entra's sign-in host only — there is no CDN in the CSP at all, since every script is now either vendored or bundled by Vite, not loaded from a third party at runtime.

## 8. Security model (summary)

See [AUTH.md](../AUTH.md) Section E for the full explanation; the headline points:

1. **No secrets in the frontend, ever** — only public OAuth identifiers exist in the code.
2. **Entra ID's own assignment gate is the real allow-list** — the optional frontend `VITE_ALLOWED_EMAILS` list (Section 5) is a convenience layer on top of it, not a substitute.
3. **No data leaves the browser** — all W-9/MSA processing is client-side; the only network calls are to Entra ID and one same-origin fetch for the bundled template.
4. **Every dependency ships same-origin** — vendored-with-recorded-SHA-512 for MSAL/PizZip/pako/Tesseract's binary assets, bundled by Vite for everything else — none of it is fetched live from a third-party CDN at runtime.
5. **No stored-XSS surface** — every value from outside this codebase (account claims, filenames, extracted W-9 text) is rendered via `textContent`, never `innerHTML`, wherever it's dynamic.
6. **The auth gate is a real DOM-level block** (`inert`), not merely a visual one — closing a confirmed keyboard-focus bypass found during a security review of this codebase.
7. **Generated PDFs can't leak the previous contractor's text** — the PDF is always a flattened image of the fully-replaced document, never an edited copy of an original PDF with old text merely covered up.
8. **Sensitive generated documents don't linger** — History is per-tab-session and purged on every page load and every History view open, once that session ends.

## 9. Known limitations

See [REQUIREMENTS.md](REQUIREMENTS.md) Section 4 ("Constraints") and Section 5 ("Out of scope") for the authoritative list. In short: there is exactly one MSA template, bundled into the app — supporting a different template means replacing that file and updating the constants that describe its previous placeholder text; OCR's digit-box pipeline is tuned to the IRS W-9 (Rev. March 2024) layout; only one uploaded file is actually used even if several are selected; and there is no multi-tenant or server-side audit trail.

## 10. Where to look next

- Setting up or changing authentication → [AUTH.md](../AUTH.md)
- Running, testing, or deploying (with screenshots of the actual app) → [WORKFLOW.md](WORKFLOW.md)
- What the product is supposed to do → [REQUIREMENTS.md](REQUIREMENTS.md)
- Updating a vendored library → [vendor/README.md](../vendor/README.md)
