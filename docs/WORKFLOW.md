# Workflow

How this app is used end-to-end, and how it's developed, tested, and deployed. See [AUTH.md](AUTH.md) for the authentication setup workflow specifically, and [PROJECT_DOCUMENTATION.md](PROJECT_DOCUMENTATION.md) for the architecture behind each step.

## A. End-user workflow

```
Sign in (Entra ID)
        │
        ▼
┌─────────────────────────────────────────────────────────────┐
│ Step 1 — Upload W-9                                         │
│   Upload a W-9 PDF → Extract Data → review/edit the         │
│   extracted company name, tax ID, address                   │
└──────────────────────┬────────────────────────────────────────┘
                       ▼
┌─────────────────────────────────────────────────────────────┐
│ Step 2 — Upload MSA template                                │
│   Upload the MSA template (PDF or Word) → app confirms it   │
│   actually contains the expected placeholder text            │
└──────────────────────┬────────────────────────────────────────┘
                       ▼
┌─────────────────────────────────────────────────────────────┐
│ Step 3 — Manual details                                     │
│   Enter contractor rep, role, location, start date,          │
│   billing rate → Apply Details (this is the step that        │
│   actually edits the template and generates both files)      │
└──────────────────────┬────────────────────────────────────────┘
                       ▼
┌─────────────────────────────────────────────────────────────┐
│ Step 4 — Download                                            │
│   Download as PDF and/or Word — both are always available,  │
│   regardless of which format the template was uploaded in    │
└─────────────────────────────────────────────────────────────┘
                       │
                       ▼
        History (optional) — every generated pair of
        files is auto-saved locally in this browser tab;
        re-download or delete anytime from the History view.
        Gone once the tab closes and any tab is reloaded.
```

1. **Sign in.** The very first thing anyone sees is the sign-in screen (`#authGate`) — nothing else on the page is reachable, by mouse or keyboard, until sign-in succeeds. Only accounts assigned to this app in Entra ID can get past it.
2. **Step 1 — Upload W-9.** Drop or select a W-9 PDF, click "Extract Data." The app reads the PDF's text layer directly if it has one; if it's a scanned/photographed W-9 with no text layer, it falls back to OCR (including a specialized pass just for the boxed SSN/EIN grid). The extracted company name, tax ID, and address show up in editable fields — correct anything the extraction got wrong before continuing.
3. **Step 2 — Upload MSA template.** Upload the actual Master Service Agreement template (PDF or Word) that needs the new contractor's details inserted. The app checks the template really contains the expected placeholder text before letting you continue — if it doesn't recognize the template, it flags that instead of silently producing a broken document.
4. **Step 3 — Manual details.** Fill in the handful of fields that aren't on a W-9 at all (representative name, role, location, start date, billing rate), then click "Apply Details." This is the step where the actual document editing happens — the previous contractor's name/tax ID/address are replaced with the new ones throughout the template, and the manual fields are inserted next to their labels.
5. **Step 4 — Download.** Download both the PDF and the Word version. Whichever format wasn't the uploaded template gets a plain from-scratch summary document instead of a real format conversion (there's no backend available to do a real conversion).
6. **History (anytime).** Switch to the "History" nav tab to see every document generated in the current browser tab, with per-entry re-download and delete. This is intentionally scoped to the current tab session, not saved forever — see [PROJECT_DOCUMENTATION.md](PROJECT_DOCUMENTATION.md#history) for why.
7. **Sign out.** Reloads the page, clearing both the Entra session and any in-memory app state.

## B. Local development workflow

1. Clone the repo.
2. One-time Entra setup: follow [AUTH.md Section A](AUTH.md#a-microsoft-entra-configuration) (App Registration + who's allowed to sign in). This is a manual step in the Entra admin portal — nothing in the repo can do it for you.
3. Point `authConfig.js` (or `.env`, if using the optional Vite path — see AUTH.md Section G) at the real client ID / tenant ID.
4. Serve the folder over real HTTP — not `file://`. Any static server works:
   ```bash
   npx serve .
   # or
   python -m http.server 5500
   ```
5. Register that exact local origin as a redirect URI in Entra (AUTH.md Section A/F) — this is the step that most often needs redoing when the local dev port changes.
6. Open the dev URL, sign in, and walk through all 4 steps with a real (or test) W-9 and MSA template.

## C. Testing workflow

- `tests/e2e.js` (Playwright) — drives the app in a real browser through the wizard flow. Run via the scripts declared in `tests/package.json`.
- **Manual security/auth smoke test** (do this after any change to `auth.js`/`auth-ui.js`/`authConfig.js`/`index.html`'s gate markup): sign in with an assigned account, reload (confirm silent restore), sign out, attempt sign-in with a non-assigned account (confirm Entra itself refuses it), and — specifically — **press Tab repeatedly from the sign-in screen and confirm focus never lands on anything behind it** (nav links, wizard controls). See [AUTH.md Section H](AUTH.md#h-security-checklist-before-deployment) for the full checklist.
- **Manual extraction spot-check**: run a real (or synthetic/dummy) W-9 through Step 1 and confirm company name / tax ID / address come out correctly for both a text-layer PDF and a scanned/OCR'd one.
- **Manual generation spot-check**: confirm the previous contractor's details don't remain visible/selectable anywhere in the generated PDF (this was specifically engineered around by rasterizing any edited PDF page — see PROJECT_DOCUMENTATION.md).

## D. Change/contribution workflow

- **Vendored dependency updates** (`vendor/*.esm.js`, Tesseract's worker/core/lang files): follow the exact steps in [vendor/README.md](vendor/README.md#updating-a-version) — fetch, check for further nested imports, record the new SHA-512, update the version in the importing file's path.
- **New MSA template support**: `script.js`'s `MSA_TEMPLATE_PLACEHOLDERS` is hardcoded to one specific template's previous values. Supporting a different template means updating that constant (and possibly `MANUAL_FIELD_LABELS` if the label wording differs) — re-uploading a different template alone does not work.
- **Auth changes**: read [AUTH.md](AUTH.md) first — it documents the security reasoning behind every non-obvious choice (redirect vs. popup, no allow-list in the frontend, sessionStorage cache, etc.), so a change that looks like an improvement in isolation may reintroduce an already-solved problem.
- **Any change touching the auth gate's DOM** (`#authGate`, `#introOverlay`, `<nav>`, `#historyView`, `#mainStage`): re-run the keyboard Tab-through check above. A prior version of this gate was CSS-only and could be bypassed via keyboard focus — see AUTH.md's security-explanation section for what fixed that (the `inert` attribute) and don't remove it without an equivalent replacement.

## E. Deployment workflow

This is a static site — any static host works (Azure Static Web Apps, Netlify, GitHub Pages, an S3 bucket + CDN, or a plain web server). See [AUTH.md Section G](AUTH.md#g-production-deployment-instructions) for the auth-specific parts (registering the production redirect URI, HTTPS requirement). There is no build step, database, or backend to deploy — deploying is "copy these files to a static host over HTTPS."
