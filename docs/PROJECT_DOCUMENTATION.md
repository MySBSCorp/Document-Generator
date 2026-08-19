# Project Documentation

Full reference documentation for the Document Generator product: what it is, how it's built, and how every piece fits together. For the auth-specific deep dive see [AUTH.md](AUTH.md); for requirements see [REQUIREMENTS.md](REQUIREMENTS.md); for the step-by-step usage/dev/deploy flow see [WORKFLOW.md](WORKFLOW.md).

## 1. Overview

Document Generator is a **frontend-only, no-backend** web app that turns an uploaded W-9 tax form into a filled Master Service Agreement (MSA). A user signs in with an approved Microsoft Entra ID account, uploads a W-9, uploads an MSA template, fills in a few extra manual details, and downloads the finished agreement as both PDF and Word — all processed entirely in the browser, with nothing sent to any server.

## 2. Project structure

```
Document-Generator/
├── index.html            # All markup: auth gate, nav, wizard, history view
├── style.css              # All styling
├── script.js              # Core app logic: extraction, generation, history (~1770 lines)
├── authConfig.js          # Public MSAL/Entra configuration
├── auth.js                # MSAL controller (session restore, login, logout)
├── auth-ui.js             # DOM wiring between auth.js and index.html
├── .env.example           # Documents the optional VITE_MSAL_* build-time variables
├── .gitignore             # Excludes .env, node_modules
├── vendor/                # Vendored third-party libraries (see vendor/README.md)
│   ├── msal-browser-4.30.0.esm.js
│   ├── msal-common-15.17.0.esm.js
│   ├── docx-8.5.0.esm.js
│   ├── pizzip-3.2.0.esm.js
│   ├── pako-2.1.0.esm.js
│   ├── tesseract/                 # OCR worker/core/language data
│   └── README.md                  # Source URLs + SHA-512 for every vendored file
├── tests/                 # Playwright end-to-end tests
├── AUTH.md                # Authentication setup, security rationale, checklist
├── WORKFLOW.md            # Usage / dev / test / deploy workflow
├── REQUIREMENTS.md        # Functional & non-functional requirements
└── PROJECT_DOCUMENTATION.md  # This file
```

No `node_modules`, bundler, or build step is required to run this app — every dependency is either loaded via a same-origin SRI-verified `<script>` tag or vendored locally as a plain ES module.

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
                         │   history)                 │
                         └───┬─────────┬─────────┬───┘
                             │         │         │
                    pdf.js   │  Tesseract.js │  pdf-lib / docx / PizZip
                 (text/render)│    (OCR)      │   (document generation)
                             │         │         │
                             ▼         ▼         ▼
                     Everything above runs in-browser.
                     IndexedDB (History) and sessionStorage
                     (MSAL cache, history session ID) are the
                     only persistence — no network calls except
                     to Entra ID and, for library integrity
                     checks, one CDN.
```

There is no server tier at all. The three architectural layers are:
1. **Auth layer** (`authConfig.js`, `auth.js`, `auth-ui.js`) — gates everything else behind Entra ID sign-in. See Section 5 and [AUTH.md](AUTH.md).
2. **Application layer** (`script.js`) — the wizard, extraction, and generation logic. See Section 6.
3. **Presentation layer** (`index.html`, `style.css`) — markup and styling for both layers above.

## 4. Technology stack

| Concern | Technology | Notes |
|---|---|---|
| Language | Plain JavaScript (ES modules), HTML, CSS | No TypeScript, no build step |
| Authentication | Microsoft Entra ID via `@azure/msal-browser` 4.30.0 | Vendored locally; see AUTH.md |
| PDF text extraction & rendering | pdf.js 3.11.174 | Loaded via SRI-pinned `<script>` from cdnjs, worker itself re-verified via SHA-512 at runtime |
| OCR | Tesseract.js 7.0.0 | Loaded via SRI-pinned `<script>`; worker/core/language files vendored locally |
| PDF generation/editing | pdf-lib 1.17.1 | Loaded via SRI-pinned `<script>` from cdnjs |
| Word (.docx) generation/editing | docx 8.5.0 + PizZip 3.2.0 (+ pako 2.1.0 dependency) | Vendored as local ES modules |
| Local persistence | IndexedDB (history), `sessionStorage` (MSAL cache, history session scoping) | No `localStorage` use for sensitive data |
| Hosting | Any static HTTP(S) host | No server-side runtime |

## 5. Authentication layer

Covered in full in [AUTH.md](AUTH.md) — summary here for context:

- **Protocol**: OAuth 2.0 Authorization Code Flow with PKCE, via MSAL.js. No implicit flow, no client secret (this is a public SPA client).
- **Flow**: `loginRedirect`/`logoutRedirect` (not popup — Microsoft's own Cross-Origin-Opener-Policy header breaks popup-completion detection; redirect has no such issue, and costs nothing here since sign-in only ever happens before any protected content is reachable).
- **Scopes**: `openid`, `profile`, `email` only — no Microsoft Graph or other API access is ever requested or held.
- **Who's allowed**: enforced entirely by Entra ID's own "Assignment required" Enterprise Application setting. There is **no allow-list of any kind inside this codebase** — deliberately, because anything present in frontend code is always readable by anyone who opens devtools or clones the repo. See AUTH.md Section E.
- **Session cache**: `sessionStorage` (not `localStorage`) — tokens/account data don't persist once the tab closes.
- **The gate** (`#authGate`): a fixed, full-viewport overlay. Critically, it is backed by the native `inert` attribute on everything behind it (`#introOverlay`, `<nav>`, `#historyView`, `#mainStage`), not just CSS `z-index` — this closes a real, confirmed security-review finding where keyboard/programmatic focus could reach the nav's History link behind a purely visual overlay. `inert` is only lifted once `auth-ui.js#showAuthenticated()` actually runs.
- **Blocked sign-ins**: if Entra itself redirects back with an error (rare — it usually shows its own hosted error page instead), `auth.js` throws `AccessDeniedError` and `auth-ui.js` shows a pop-up warning (`#authBlockedModal`) with a "Retry" that returns to the sign-in screen (it does not itself retry sign-in).

## 6. Application layer (`script.js`)

### 6.1 Wizard controller
Four steps, each gated on the previous step's required data being valid (`showWizardStep`, `validateW9Fields`, `validateManualFields`). A minimum-loading-time guard (`waitRemaining`) ensures fast operations still show a spinner briefly rather than flashing, and a decorative intro animation plays once on load (purely cosmetic, has no data role, and is `inert` until sign-in succeeds).

### 6.2 W-9 extraction pipeline (Step 1)
1. **Text-layer extraction first** (`extractPdfText`): pdf.js `getTextContent()` plus AcroForm field values, for W-9s that are real fillable/text PDFs.
2. **OCR fallback** (`ocrPdfText`) only if the text layer looks empty (`HAS_TEXT_THRESHOLD` = 20 chars): renders up to 2 pages to canvas at 3x scale and runs Tesseract on the whole page, **plus** a second, specialized pass (`ocrTinBoxDigits`/`cropCanvasRegion`/`ocrDigitsOnly`) that crops and OCRs just the bordered SSN/EIN digit-box grid with a digit-only whitelist — because whole-page OCR reliably reads prose but reliably garbles those boxed digit cells. Crop coordinates are tuned to the IRS W-9 (Rev. March 2024) layout specifically.
3. **Field parsing** (`parseW9Fields`): regex-based, tolerant of OCR noise (flexible whitespace/punctuation matching via `loosePattern`/`loose`, label-boundary detection, a multi-step address-splitting fallback chain for when the "City, state, ZIP" label doesn't extract cleanly). Extracts company name, tax ID (SSN or EIN, with type inferred from format), street address, city, state+ZIP.
4. Every extracted field is shown in an editable summary and re-validated; the user must have at least a tax ID present to proceed.

### 6.3 MSA template handling (Step 2)
The uploaded template (PDF or `.docx`) is checked, read-only, for the expected placeholder text (`MSA_TEMPLATE_PLACEHOLDERS` — hardcoded to one specific template's previous contractor's company name/tax ID/address) before the user can proceed, so a wrong or unsupported template is caught early with a clear message instead of silently producing garbage output later.

### 6.4 Document generation (Step 3)
The core design decision here: real-world MSA templates aren't fillable forms or merge-tag documents — they're ordinary contracts where the previous contractor's details appear as repeated plain text. So generation is literal find-and-replace, not template-tag substitution:

- **DOCX path** (`applyDocxReplacements`): unzips the `.docx` with PizZip, parses `word/document.xml`, and replaces matching text across `<w:t>` runs — including matches that span multiple runs (common when Word splits a phrase across differently-formatted runs).
- **PDF path** (`applyPdfReplacements`): locates placeholder text via pdf.js and its bounding box, then — because pdf-lib can only *add* drawing instructions, never remove existing ones — **rasterizes any page that needs an edit to a flat PNG image first** via canvas rendering, rebuilds that page in pdf-lib as an image, and only then draws a cover rectangle and the new text on top. This specifically prevents the previous contractor's sensitive text from remaining present and extractable/copy-pasteable underneath a merely-visual cover. Pages with no matches keep their real, selectable text untouched.
- Auto-replacements (from Step 1's data) and manual-field replacements (from Step 3's form) are combined into a **single** replacement pass, not two — because a second pass over an already-rasterized page would find its target label already turned to pixels.
- The "other" format (whichever wasn't the uploaded template) is generated as a **plain summary document from scratch** (`buildSimplePdf`/`buildSimpleDocx`), not a real conversion of the template's layout — a true conversion isn't achievable without a backend.

### 6.5 History
- **Storage**: IndexedDB (chosen over `localStorage` specifically because it can store Blobs directly and isn't capped at ~5MB) — stores the actual generated PDF and Word blobs, not just metadata.
- **Scoping & retention**: each browser tab gets a random session ID (`crypto.randomUUID()` in `sessionStorage`). Every read/write is filtered to the current tab's session, and entries from any *other* (closed) session are purged unconditionally on every page load — not only when the History view happens to be opened — specifically so a closed tab's SSN/EIN-bearing documents don't linger indefinitely. Manual per-entry deletion is also available.

## 7. Presentation layer

- `index.html` holds all markup: the auth gate and its states (loading/unauthenticated/error), the blocked-sign-in pop-up, the nav (with the authenticated-user badge), the 4-step wizard, and the history view.
- `style.css` holds all styling, including the auth gate's intentionally flat (non-boxed-card) design and the modal styling for the blocked-sign-in pop-up.
- A Content-Security-Policy `<meta>` tag (no server available to send a real header) restricts `script-src` to `'self'` plus the one CDN whose `<script>` tags carry Subresource Integrity hashes, and `connect-src`/`frame-src` to Entra's sign-in host only.

## 8. Security model (summary)

See [AUTH.md](AUTH.md) Section E for the full explanation; the headline points:

1. **No secrets in the frontend, ever** — nothing here needs to be, or should ever be, confidential; only public OAuth identifiers exist in the code.
2. **No allow-list in the frontend** — who may sign in is enforced entirely server-side by Entra ID.
3. **No data leaves the browser** — all W-9/MSA processing is client-side; the only network calls are to Entra ID and (for integrity verification / fallback) one CDN.
4. **Every dependency is integrity-verified** — SRI on CDN `<script>` tags, a runtime SHA-512 check on the dynamically-fetched pdf.js worker, and vendored-with-recorded-hash for everything else.
5. **No stored-XSS surface** — every value from outside this codebase (account claims, filenames, extracted W-9 text) is rendered via `textContent`, never `innerHTML`, wherever it's dynamic.
6. **The auth gate is a real DOM-level block** (`inert`), not merely a visual one — closing a confirmed keyboard-focus bypass found during a security review of this codebase.
7. **Sensitive generated documents don't linger** — History is per-tab-session and purged on every page load once that session ends.

## 9. Known limitations

See [REQUIREMENTS.md](REQUIREMENTS.md) Section 4 ("Constraints") and Section 5 ("Out of scope") for the authoritative list. In short: one specific MSA template's placeholder text is what replacement is tuned to; OCR's specialized digit-box pass is tuned to one specific W-9 revision's layout; the "other format" download is a plain summary, not a true layout-preserving conversion; and there is no multi-tenant or server-side audit trail.

## 10. Where to look next

- Setting up or changing authentication → [AUTH.md](AUTH.md)
- Running, testing, or deploying → [WORKFLOW.md](WORKFLOW.md)
- What the product is supposed to do → [REQUIREMENTS.md](REQUIREMENTS.md)
- Updating a vendored library → [vendor/README.md](vendor/README.md)
