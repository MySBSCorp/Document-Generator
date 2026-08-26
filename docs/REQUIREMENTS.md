# Project Requirements

What this product must do, and the constraints it must do it within. See [PROJECT_DOCUMENTATION.md](PROJECT_DOCUMENTATION.md) for how these requirements are actually implemented, and [AUTH.md](../AUTH.md) for the authentication requirements specifically.

## 1. Purpose

Let an internal team turn a client's W-9 into a filled Master Service Agreement (MSA) in a few clicks, without a backend, and without ever sending the client's SSN/EIN data off the user's own browser — usable from a desktop or a phone, by anyone signed in with an approved Microsoft work/school account.

## 2. Functional requirements

### 2.1 Authentication
- FR-1: Users must sign in with a Microsoft Entra ID (Azure AD) account before any part of the app is usable.
- FR-2: Only specifically approved accounts may use the app — enforced primarily by Entra's own "Assignment required" Enterprise Application setting, not by a list inside this app's code (see AUTH.md Section E for why). An optional frontend email allow-list may additionally be configured as a convenience layer, but is never the sole gate.
- FR-3: A returning user with an active session should not have to sign in again on every page load (silent session restore).
- FR-4: The signed-in user's name and email must be visible in the app once authenticated.
- FR-5: A "Sign out" action must be available and must end the session with Entra, not just locally.
- FR-6: If Entra itself refuses a sign-in, the user must see a clear message and be able to retry (return to the sign-in screen), not a broken or blank page.

### 2.2 W-9 data extraction & MSA fill (Step 1)
- FR-7: Accept a W-9 upload — a PDF, or a photo/scan (JPEG, PNG, WEBP, HEIC/HEIF).
- FR-8: Extract company/business name, tax ID (SSN or EIN, with type), and address from the uploaded W-9.
- FR-9: Extraction must work for both a normal, text-layer PDF and a scanned/photographed W-9 with no text layer (OCR fallback) — including scans that are rotated, shifted, low-contrast, or otherwise imperfect.
- FR-10: All extracted fields must be shown to the user and be editable before proceeding — extraction is a starting point, not a silent black box. A visible warning must tell the user to review the extracted values before continuing.
- FR-11: The user must not be able to proceed past Step 1 without at least a tax ID value present (extracted or manually entered).
- FR-12: On successful extraction, the app must automatically insert the extracted data into the bundled MSA template (see Section 2.3) and advance to Step 2 — no separate template-upload or confirmation step is required from the user.

### 2.3 MSA template
- FR-13: The app ships with exactly one MSA template, bundled into the deployed app itself — the user is never asked to upload one. (Historically this was a separate upload-and-verify wizard step; it no longer is — see C-3 for what's required to support a different template.)

### 2.4 Manual details (Step 2)
- FR-14: Collect fields that never appear on a W-9: contractor representative name(s), role, location, start date, billing rate.
- FR-15: Validate these fields before allowing document generation.
- FR-16: Generating the document must replace all occurrences of the previous contractor's company name, tax ID, and address throughout the template — not just the first occurrence, and not leave any of the old values recoverable from the generated file (including from the underlying PDF's own text/content layer, not just what's visible on screen).
- FR-17: The manual-detail values must be inserted next to their corresponding labels in the template.
- FR-18: The user must be able to see a live preview of the document as it will be generated, before committing to it, and a preview of the originally-uploaded W-9 for reference.

### 2.5 Output (Step 3)
- FR-19: Produce a downloadable PDF and a downloadable Word (.docx) version of the filled document — both, always, as full and complete generations of the same content (not a partial/summary fallback for either format).
- FR-20: The user must be able to reset the wizard and generate another document without reloading the page.

### 2.6 History
- FR-21: Every generated document pair (PDF + Word) must be automatically retained locally so the user can re-download without redoing the wizard.
- FR-22: The user must be able to delete an individual history entry.
- FR-23: History must not persist indefinitely or leak between unrelated browser sessions — see NFR-6.

## 3. Non-functional requirements

### 3.1 Security & privacy
- NFR-1: **No backend, database, or server component of any kind.** All processing (PDF parsing, OCR, document generation) happens entirely in the user's browser. No W-9 or generated-document data is ever transmitted to any server this project controls.
- NFR-2: **No secrets in the frontend.** No client secret, API key, password, or database credential may exist anywhere in this codebase, in any form (hardcoded, encoded, "obfuscated," or otherwise) — see AUTH.md Section E for the full reasoning. Only public OAuth identifiers (client ID, tenant ID) may be present.
- NFR-3: **Least-privilege OAuth scopes.** Only the minimum OIDC scopes needed to establish identity (`openid`/`profile`/`email`) — no Microsoft Graph or other API permission.
- NFR-4: **Content Security Policy** restricting script/connect/frame sources to exactly what's needed: `script-src` limited to the app's own origin (plus `wasm-unsafe-eval`, required for Tesseract's WASM OCR engine), `connect-src`/`frame-src` limited to Entra's sign-in host, and no third-party CDN permitted in the policy at all.
- NFR-5: **No stored-XSS surface.** Any value that ultimately comes from outside this codebase — uploaded filenames, extracted W-9 text, the signed-in account's name/email — must be inserted into the DOM only via `textContent` (or equivalent), never `innerHTML`/`insertAdjacentHTML`/`document.write` with untrusted content.
- NFR-6: **Minimize how long sensitive data lingers.** Generated documents (which contain SSN/EIN) stored for the History feature must be scoped to the current browser tab session and purged once that session ends — not retained indefinitely across tabs/devices.
- NFR-7: **The auth gate must be a real block, not merely a visual one.** Content behind the sign-in screen must be unreachable by any interaction method (mouse, keyboard/Tab, or programmatic focus/DOM events), not just hidden from view — see AUTH.md's security-explanation section for the specific fix this requirement drove.
- NFR-8: **No dependency executes without shipping same-origin.** Every third-party library is either vendored into the repo with its source URL and SHA-512 recorded, or bundled into the app's own build output by Vite — none is ever fetched live from a third-party CDN at runtime, so a compromised CDN can't silently inject code that touches SSN/EIN data.

### 3.2 Usability
- NFR-9: The wizard must clearly communicate progress (which step, loading states) and must not let the user proceed past a step with missing/invalid required data.
- NFR-10: Extraction and generation must complete within a reasonable time on typical hardware; long-running operations (OCR, PDF rasterization) must show a loading indicator rather than appearing frozen — including on mobile, where the loading indicator must remain correctly positioned regardless of scroll position.
- NFR-11: **The app must be usable on mobile phones, not just desktop.** Every step of the wizard, the Document Preview panel, History, and all modals must remain fully readable and operable (no clipped or unreachable content, no unusably small tap targets) down to a ~320px-wide viewport — validated against common iPhone and Samsung Galaxy screen sizes.

### 3.3 Compatibility
- NFR-12: Must run in current evergreen browsers (desktop and mobile) supporting native ES modules, `<meta>` CSP, `IndexedDB`, and the `inert` HTML attribute.
- NFR-13: Must work when served from any static HTTP(S) origin after being built — no dependency on a specific hosting platform's features.

### 3.4 Maintainability
- NFR-14: The project is built with Vite; running or deploying it requires `npm install` and either `npm run dev` (local) or `npm run build` (production) — there is no build-step-free fallback path, and none is required.
- NFR-15: Every non-obvious security or design decision must be documented in-repo (comments and/or AUTH.md/PROJECT_DOCUMENTATION.md) so a future change doesn't unknowingly reintroduce an already-solved problem.

## 4. Constraints

- C-1: Frontend-only — explicitly no backend, no database, no server-rendered pages.
- C-2: Single Microsoft Entra tenant; sign-in restricted to specific, individually-assigned accounts within it.
- C-3: Exactly one MSA template is supported, bundled at `public/assets/msa-template.docx`. Supporting a materially different template requires replacing that file **and** updating the `script.js` constants describing its previous placeholder text (company name/tax ID/address, and manual-field label wording) — it is a code change, not a re-upload, and there is no way for an end user to supply their own template.
- C-4: OCR/extraction tuned specifically to the IRS W-9 (Rev. March 2024) layout; other revisions or heavily non-standard scans may not extract cleanly via the specialized digit-box pass (the general label-based parser may still work).
- C-5: Only one uploaded file is used per extraction, even if the file picker allows selecting several — the first PDF/image in the selection is used and any others are ignored.

## 5. Out of scope

- A true PDF↔Word format conversion preserving an arbitrary template's exact layout (moot in practice, since the one supported template is bundled and always produces a full real generation of both formats — see FR-19 — but a *different*, user-supplied template's layout still could not be faithfully reproduced without a backend).
- Multi-tenant / multi-organization support (only one Entra tenant is targeted).
- Any server-side storage, audit logging, or multi-device history sync.
- Support for W-9 revisions other than the one the TIN-box OCR crop coordinates were tuned against (extraction may still work via the general label-based parser, just without the specialized digit-box OCR pass).
- A native mobile app or installable PWA — "mobile support" (NFR-11) means the responsive web app works well in a phone's browser, not a packaged app.
