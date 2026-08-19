# Project Requirements

What this product must do, and the constraints it must do it within. See [PROJECT_DOCUMENTATION.md](PROJECT_DOCUMENTATION.md) for how these requirements are actually implemented, and [AUTH.md](AUTH.md) for the authentication requirements specifically.

## 1. Purpose

Let an internal team turn a client's W-9 into a filled Master Service Agreement (MSA) in a few clicks, without a company account signing in with a Microsoft work/school account, without standing up a backend, and without ever sending the client's SSN/EIN data off the user's own browser.

## 2. Functional requirements

### 2.1 Authentication
- FR-1: Users must sign in with a Microsoft Entra ID (Azure AD) account before any part of the app is usable.
- FR-2: Only specifically approved accounts may use the app — enforced by Entra's own "Assignment required" Enterprise Application setting, not by a list inside this app's code (see AUTH.md Section E for why).
- FR-3: A returning user with an active session should not have to sign in again on every page load (silent session restore).
- FR-4: The signed-in user's name and email must be visible in the app once authenticated.
- FR-5: A "Sign out" action must be available and must end the session with Entra, not just locally.
- FR-6: If Entra itself refuses a sign-in, the user must see a clear message and be able to retry (return to the sign-in screen), not a broken or blank page.

### 2.2 W-9 data extraction (Step 1)
- FR-7: Accept a W-9 PDF upload (drag-and-drop or file picker).
- FR-8: Extract company/business name, tax ID (SSN or EIN, with type), and address from the uploaded W-9.
- FR-9: Extraction must work for both a normal, text-layer PDF and a scanned/photographed W-9 with no text layer (OCR fallback).
- FR-10: All extracted fields must be shown to the user and be editable before proceeding — extraction is a starting point, not a silent black box.
- FR-11: The user must not be able to proceed past Step 1 without at least a tax ID value present (extracted or manually entered).

### 2.3 MSA template handling (Step 2)
- FR-12: Accept an MSA template upload in either PDF or Word (.docx) format.
- FR-13: Verify the uploaded template actually contains the text this app knows how to replace, before letting the user proceed — and tell them clearly if it doesn't, rather than silently producing a broken document later.

### 2.4 Manual details (Step 3)
- FR-14: Collect fields that never appear on a W-9: contractor representative name(s), role, location, start date, billing rate.
- FR-15: Validate these fields before allowing document generation.
- FR-16: Generating the document must replace all occurrences of the previous contractor's company name, tax ID, and address throughout the template — not just the first occurrence, and not leave any of the old values recoverable from the generated file (including from the underlying PDF text layer, not just what's visible).
- FR-17: The manual-detail values must be inserted next to their corresponding labels in the template.

### 2.5 Output (Step 4)
- FR-18: Produce a downloadable PDF and a downloadable Word (.docx) version of the filled document — both, regardless of which single format the template was uploaded in.
- FR-19: The "other" format (the one that wasn't the uploaded template) may be a plain summary document rather than a true format conversion of the template's exact layout — a true conversion isn't achievable without a backend, and none is planned.

### 2.6 History
- FR-20: Every generated document pair (PDF + Word) must be automatically retained locally so the user can re-download without redoing the wizard.
- FR-21: The user must be able to delete an individual history entry.
- FR-22: History must not persist indefinitely or leak between unrelated browser sessions — see NFR-6.

## 3. Non-functional requirements

### 3.1 Security & privacy
- NFR-1: **No backend, database, or server component of any kind.** All processing (PDF parsing, OCR, document generation) happens entirely in the user's browser. No W-9 or generated-document data is ever transmitted to any server this project controls.
- NFR-2: **No secrets in the frontend.** No client secret, API key, password, or database credential may exist anywhere in this codebase, in any form (hardcoded, encoded, "obfuscated," or otherwise) — see AUTH.md Section E for the full reasoning. Only public OAuth identifiers (client ID, tenant ID) may be present.
- NFR-3: **Least-privilege OAuth scopes.** Only the minimum OIDC scopes needed to establish identity (`openid`/`profile`/`email`) — no Microsoft Graph or other API permission.
- NFR-4: **Content Security Policy** with no `unsafe-inline`/`unsafe-eval`, restricting script/connect/frame sources to exactly what's needed (the app's own origin, the one CDN used for SRI-pinned libraries, and Entra's sign-in endpoint).
- NFR-5: **No stored-XSS surface.** Any value that ultimately comes from outside this codebase — uploaded filenames, extracted W-9 text, the signed-in account's name/email — must be inserted into the DOM only via `textContent` (or equivalent), never `innerHTML`/`insertAdjacentHTML`/`document.write` with untrusted content.
- NFR-6: **Minimize how long sensitive data lingers.** Generated documents (which contain SSN/EIN) stored for the History feature must be scoped to the current browser tab session and purged once that session ends — not retained indefinitely across tabs/devices.
- NFR-7: **The auth gate must be a real block, not merely a visual one.** Content behind the sign-in screen must be unreachable by any interaction method (mouse, keyboard/Tab, or programmatic focus/DOM events), not just hidden from view — see AUTH.md's security-explanation section for the specific fix this requirement drove.
- NFR-8: **No dependency loaded without an integrity guarantee.** Every third-party library either ships with Subresource Integrity (SRI) on its `<script>` tag, or is vendored into the repo with its source URL and hash recorded, so a compromised CDN can't silently inject code that touches SSN/EIN data.

### 3.2 Usability
- NFR-9: The wizard must clearly communicate progress (which step, loading states) and must not let the user proceed past a step with missing/invalid required data.
- NFR-10: Extraction and generation must complete within a reasonable time on typical hardware; long-running operations (OCR, PDF rasterization) must show a loading indicator rather than appearing frozen.

### 3.3 Compatibility
- NFR-11: Must run in current evergreen desktop browsers supporting native ES modules, `<meta>` CSP, `IndexedDB`, and the `inert` HTML attribute — no build/transpilation step is used, so no support is claimed for legacy browsers.
- NFR-12: Must work when served from any static HTTP(S) origin — no dependency on a specific hosting platform's features.

### 3.4 Maintainability
- NFR-13: No build step, bundler, or `node_modules` should be required to run the app locally (an optional npm/Vite path may exist alongside this, but must not be the only way to run it).
- NFR-14: Every non-obvious security or design decision must be documented in-repo (comments and/or AUTH.md/PROJECT_DOCUMENTATION.md) so a future change doesn't unknowingly reintroduce an already-solved problem.

## 4. Constraints

- C-1: Frontend-only — explicitly no backend, no database, no server-rendered pages.
- C-2: Single Microsoft Entra tenant; sign-in restricted to specific, individually-assigned accounts within it.
- C-3: One specific MSA template's placeholder text is what the replacement logic is tuned to; supporting a materially different template requires a code change, not just a re-upload.
- C-4: OCR/extraction tuned specifically to the IRS W-9 (Rev. March 2024) layout; other revisions or heavily non-standard scans may not extract cleanly.

## 5. Out of scope

- A true PDF↔Word format conversion preserving the original template's exact layout.
- Multi-tenant / multi-organization support (only one Entra tenant is targeted).
- Any server-side storage, audit logging, or multi-device history sync.
- Support for W-9 revisions other than the one the TIN-box OCR crop coordinates were tuned against (extraction may still work via the general label-based parser, just without the specialized digit-box OCR pass).
