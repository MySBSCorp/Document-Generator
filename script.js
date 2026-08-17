// docx building/editing — no plain-<script> UMD build exists for these
// (pizzip's npm package targets bundlers), so this file is loaded as a
// module and imports jsDelivr's auto-bundled ESM builds directly.
import PizZip from 'https://cdn.jsdelivr.net/npm/pizzip@3.2.0/+esm';
import { Document, Packer, Paragraph, TextRun } from 'https://cdn.jsdelivr.net/npm/docx@8.5.0/+esm';

const generateBtn = document.getElementById('generateBtn');
const uploadZone = document.querySelector('.upload-block');
const fileInput = document.getElementById('fileInput');
const fileList = document.getElementById('fileList');
const status = document.getElementById('status');
const downloadPdfBtn = document.getElementById('downloadPdfBtn');
const downloadWordBtn = document.getElementById('downloadWordBtn');
const brandLogo = document.querySelector('.brand');
const msaTemplateInput = document.getElementById('msaTemplateInput');
const msaFileList = document.getElementById('msaFileList');
const msaUploadZone = msaTemplateInput.closest('.upload-block');
const manualRep = document.getElementById('manualRep');
const manualRole = document.getElementById('manualRole');
const manualLocation = document.getElementById('manualLocation');
const manualStartDate = document.getElementById('manualStartDate');
const manualStartDatePicker = document.getElementById('manualStartDatePicker');
const manualStartDatePickerBtn = document.getElementById('manualStartDatePickerBtn');
const manualBillingRate = document.getElementById('manualBillingRate');
const applyManualBtn = document.getElementById('applyManualBtn');

// Start Date stays a free-text field (so "August 5th, 2026 (Tentative)"
// style values still work) — the calendar button just offers a native date
// picker as a shortcut that formats its result back into that text field.
manualStartDatePickerBtn.addEventListener('click', () => {
  if (manualStartDatePicker.showPicker) {
    manualStartDatePicker.showPicker();
  } else {
    manualStartDatePicker.focus();
    manualStartDatePicker.click();
  }
});

manualStartDatePicker.addEventListener('change', () => {
  if (!manualStartDatePicker.value) return;
  const [year, month, day] = manualStartDatePicker.value.split('-').map(Number);
  const picked = new Date(year, month - 1, day);
  manualStartDate.value = picked.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });
  manualStartDate.dispatchEvent(new Event('input', { bubbles: true }));
});

/* ---------- Wizard: one step's screen visible at a time ---------- */
const wizard = document.getElementById('wizard');
const actionPanel = document.querySelector('.action-panel');
const wizardSteps = Array.from(wizard.querySelectorAll('.wizard-step'));
const wizardLoading = document.getElementById('wizardLoading');
const wizardLoadingText = document.getElementById('wizardLoadingText');
const extractDataBtn = document.getElementById('extractDataBtn');
const step1NextBtn = document.getElementById('step1NextBtn');
const step2BackBtn = document.getElementById('step2BackBtn');
const step2NextBtn = document.getElementById('step2NextBtn');
const step3BackBtn = document.getElementById('step3BackBtn');
const step3NextBtn = document.getElementById('step3NextBtn');
const step4BackBtn = document.getElementById('step4BackBtn');

function showWizardStep(n) {
  wizardSteps.forEach((el) => {
    el.classList.toggle('active', Number(el.dataset.step) === n);
  });
}

// #wizard is a normal, content-sized child of the scrollable .action-panel —
// it can be taller than the panel's visible viewport (that's what lets the
// panel scroll at all). A position:fixed overlay is used instead of an
// absolute one nested inside #wizard so it can't scroll away from under
// itself; this keeps it pinned to exactly the panel's on-screen rect.
function positionWizardLoading() {
  const rect = actionPanel.getBoundingClientRect();
  wizardLoading.style.top = `${rect.top}px`;
  wizardLoading.style.left = `${rect.left}px`;
  wizardLoading.style.width = `${rect.width}px`;
  wizardLoading.style.height = `${rect.height}px`;
}

function showWizardLoading(text) {
  wizardLoadingText.textContent = text;
  positionWizardLoading();
  wizardLoading.hidden = false;
}

function hideWizardLoading() {
  wizardLoading.hidden = true;
}

window.addEventListener('resize', () => {
  if (!wizardLoading.hidden) positionWizardLoading();
});

// The stacking-file loader only appears for the real work buttons (Extract
// Data / Insert Data / Apply Details) — Next/Back just switch which step is
// visible, so they stay instant.
step1NextBtn.addEventListener('click', () => showWizardStep(2));
step2BackBtn.addEventListener('click', () => showWizardStep(1));
step2NextBtn.addEventListener('click', () => showWizardStep(3));
step3BackBtn.addEventListener('click', () => showWizardStep(2));
step3NextBtn.addEventListener('click', () => showWizardStep(4));
step4BackBtn.addEventListener('click', () => showWizardStep(3));

// Each action button's loading state is shown for at least this long, even
// if the underlying work finishes sooner — otherwise a fast text-layer PDF
// extraction would flash the loader for a fraction of a second.
const MIN_LOADING_MS = 3000;

async function waitRemaining(startedAt, minMs) {
  const elapsed = Date.now() - startedAt;
  if (elapsed < minMs) await new Promise((resolve) => setTimeout(resolve, minMs - elapsed));
}

let uploadedFiles = [];
let generatedPdfUrl = null;
let generatedWordUrl = null;
let generatedBaseName = 'MSA';
let extractedW9Data = null;
let msaTemplateFile = null;
let isDocxTemplate = false;

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

// A running CSS animation (even finished, with fill-mode holding its end state)
// always wins the cascade for the property it touches, regardless of selector
// specificity — so the entrance "rise" animation on .eyebrow/etc. would
// otherwise permanently block a later :hover transform from ever applying.
// Clearing animation once it's done frees the property back up for normal rules.
document.querySelectorAll('.eyebrow, .hero h1, .subtitle').forEach((el) => {
  el.addEventListener('animationend', () => { el.style.animation = 'none'; }, { once: true });
});

/* ---------- File upload ---------- */
function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function renderFileList() {
  fileList.innerHTML = '';
  uploadedFiles.forEach((file) => {
    const li = document.createElement('li');
    li.innerHTML = `<span>${file.name}</span><span class="size">${formatSize(file.size)}</span>`;
    fileList.appendChild(li);
  });
}

function addFiles(fileArray) {
  uploadedFiles = uploadedFiles.concat(fileArray);
  renderFileList();
  setStatus(`${uploadedFiles.length} file(s) ready.`);

  const hasW9 = uploadedFiles.some((f) => f.type === 'application/pdf' || /\.pdf$/i.test(f.name));
  extractDataBtn.disabled = !hasW9;
  step1NextBtn.disabled = true;
}

extractDataBtn.addEventListener('click', async () => {
  const w9File = uploadedFiles.find((f) => f.type === 'application/pdf' || /\.pdf$/i.test(f.name));
  if (!w9File) return;

  const startedAt = Date.now();
  extractDataBtn.disabled = true;
  step1NextBtn.disabled = true;
  showWizardLoading('Extracting data...');

  let success = false;
  try {
    await processW9(w9File);
    success = validateW9Fields();
  } finally {
    await waitRemaining(startedAt, MIN_LOADING_MS);
    hideWizardLoading();
    extractDataBtn.disabled = false;
    step1NextBtn.disabled = !success;
    if (success) showWizardStep(2);
  }
});

fileInput.addEventListener('change', () => {
  addFiles(Array.from(fileInput.files));
});

['dragenter', 'dragover'].forEach((evt) =>
  uploadZone.addEventListener(evt, (e) => {
    e.preventDefault();
    uploadZone.classList.add('drag-over');
  })
);

['dragleave', 'drop'].forEach((evt) =>
  uploadZone.addEventListener(evt, (e) => {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
  })
);

uploadZone.addEventListener('drop', (e) => {
  const dropped = Array.from(e.dataTransfer.files || []);
  if (dropped.length) addFiles(dropped);
});

/* ---------- MSA template upload (Step 2) ---------- */
function setMsaTemplate(file) {
  msaTemplateFile = file;
  msaFileList.innerHTML = '';
  const li = document.createElement('li');
  li.innerHTML = `<span>${file.name}</span><span class="size">${formatSize(file.size)}</span>`;
  msaFileList.appendChild(li);
  setStatus(`${file.name} ready to fill.`);
  generateBtn.disabled = false;
  step2NextBtn.disabled = true;
  step3NextBtn.disabled = true;
  downloadPdfBtn.disabled = true;
  downloadWordBtn.disabled = true;
}

msaTemplateInput.addEventListener('change', () => {
  if (msaTemplateInput.files[0]) setMsaTemplate(msaTemplateInput.files[0]);
});

['dragenter', 'dragover'].forEach((evt) =>
  msaUploadZone.addEventListener(evt, (e) => {
    e.preventDefault();
    msaUploadZone.classList.add('drag-over');
  })
);

['dragleave', 'drop'].forEach((evt) =>
  msaUploadZone.addEventListener(evt, (e) => {
    e.preventDefault();
    msaUploadZone.classList.remove('drag-over');
  })
);

msaUploadZone.addEventListener('drop', (e) => {
  const dropped = Array.from(e.dataTransfer.files || []);
  if (dropped[0]) setMsaTemplate(dropped[0]);
});

/* ---------- W-9 extraction ---------- */
const w9Summary = document.getElementById('w9Summary');
const w9CompanyName = document.getElementById('w9CompanyName');
const w9TaxId = document.getElementById('w9TaxId');
const w9TaxIdType = document.getElementById('w9TaxIdType');
const w9Address = document.getElementById('w9Address');
const w9EditHint = document.getElementById('w9EditHint');
const w9Debug = document.getElementById('w9Debug');
const w9RawText = document.getElementById('w9RawText');

// Extraction (OCR especially) is imperfect, so the summary fields are real
// inputs the user can correct — edits write straight back into
// extractedW9Data, which is what Step 3's replacement pass actually reads.
// Required fields the MSA fill depends on — missing ones get a red outline
// (see .field-error in style.css) instead of silently producing a broken
// document. Re-checked after every manual edit so fixing a field by hand
// clears its red mark and re-enables Next without re-running extraction.
function validateW9Fields() {
  // Checked against what's actually shown in the inputs (not the raw
  // extractedW9Data sub-fields) — e.g. the address box can be filled purely
  // from city/state_zip with an empty street_address, which would otherwise
  // read as "missing" even though the field on screen isn't blank.
  const checks = [
    [w9CompanyName, () => !!w9CompanyName.value.trim()],
    [w9TaxId, () => !!w9TaxId.value.trim()],
    [w9Address, () => !!w9Address.value.trim()],
  ];
  let valid = true;
  checks.forEach(([el, isOk]) => {
    const ok = isOk();
    el.classList.toggle('field-error', !ok);
    if (!ok) valid = false;
  });
  return valid;
}

w9CompanyName.addEventListener('input', () => {
  if (extractedW9Data) extractedW9Data.company_name = w9CompanyName.value.trim();
  step1NextBtn.disabled = !validateW9Fields();
});
w9TaxId.addEventListener('input', () => {
  if (extractedW9Data) extractedW9Data.tax_id_number = w9TaxId.value.trim();
  step1NextBtn.disabled = !validateW9Fields();
});
w9TaxIdType.addEventListener('change', () => {
  if (extractedW9Data) extractedW9Data.selected_tax_id_type = w9TaxIdType.value;
});
w9Address.addEventListener('input', () => {
  if (!extractedW9Data) return;
  // The summary shows address as one combined line, so an edit collapses
  // back into a single field — city/state_zip are cleared so the parts
  // used to rebuild the full address (see buildAutoReplacements) don't
  // duplicate what's now typed directly into street_address.
  extractedW9Data.street_address = w9Address.value.trim();
  extractedW9Data.city = '';
  extractedW9Data.state_zip = '';
  step1NextBtn.disabled = !validateW9Fields();
});

async function extractPdfText(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let combined = '';

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    combined += textContent.items.map((item) => item.str).join(' ') + '\n';

    // Fillable W-9s often carry typed values as AcroForm field values rather
    // than drawn page text, so getTextContent() alone would miss them.
    const annotations = await page.getAnnotations();
    annotations.forEach((a) => {
      if (a.fieldValue) combined += ' ' + a.fieldValue;
    });
  }

  return combined;
}

// Falls back to OCR when a PDF has no embedded text layer — a scanned or
// photographed W-9 (e.g. via CamScanner) is just page images to pdf.js, so
// extractPdfText() above finds nothing. Rendering each page to a canvas and
// running Tesseract against the bitmap is the only way to read those.
// Whole-page OCR reliably reads running prose but falls apart on the Part I
// TIN box grid — bordered digit cells confuse Tesseract's page segmentation
// and come back as garbage fragments instead of digits. Cropping just that
// region, upscaling it, and re-running Tesseract with a digit-only whitelist
// (no border/label noise to segment around) reads the boxes far more
// reliably. Coordinates are fractions of the full page tuned to the IRS
// W-9 (Rev. March 2024) standard layout — Part I sits a bit past the page's
// midpoint, with the SSN box row directly above the EIN box row.
const TIN_BOX_REGION = { x0: 0.55, x1: 1.0, y0: 0.4, y1: 0.62 };

// Tight bands for just the digit cells of each row, as fractions *within*
// TIN_BOX_REGION (not the full page) — deliberately excludes the label bar
// ("Social security number" / "Employer identification number") and the
// cell border lines. OCRing the whole grid+labels in one digit-whitelisted
// pass makes Tesseract try to read border pixels as digits/hyphens and
// drowns out the real ones; a tight per-row crop avoids that entirely.
// Measured directly off a rendered W-9 (Rev. March 2024) TIN_BOX_REGION
// crop — the EIN row coordinates were verified to read "88-4163133"
// reliably; the SSN row is the same shape shifted up by one label+box unit.
const EIN_ROW_FRAC = { x0: 0.2985, x1: 0.8645, y0: 0.4811, y1: 0.571 };
const SSN_ROW_FRAC = { x0: 0.2985, x1: 0.8645, y0: 0.24, y1: 0.33 };

function cropCanvasRegion(source, region, upscale) {
  const x0 = Math.floor(source.width * region.x0);
  const x1 = Math.floor(source.width * region.x1);
  const y0 = Math.floor(source.height * region.y0);
  const y1 = Math.floor(source.height * region.y1);
  const w = x1 - x0;
  const h = y1 - y0;
  if (w <= 0 || h <= 0) return null;

  const out = document.createElement('canvas');
  out.width = w * upscale;
  out.height = h * upscale;
  const ctx = out.getContext('2d');
  ctx.imageSmoothingEnabled = false;
  ctx.drawImage(source, x0, y0, w, h, 0, 0, out.width, out.height);
  return out;
}

async function ocrDigitsOnly(cropCanvas) {
  if (!cropCanvas) return '';

  // Tesseract.recognize(image, langs, options)'s third argument only
  // configures worker *creation* — a job parameter like
  // tessedit_char_whitelist has to go through worker.setParameters()
  // instead, or it's silently dropped and the crop OCRs unconstrained.
  // tessedit_pageseg_mode 7 (SINGLE_LINE) suits a tight single row of boxed
  // digits far better than the default automatic paragraph/column detection.
  const worker = await Tesseract.createWorker('eng');
  await worker.setParameters({
    tessedit_char_whitelist: '0123456789-',
    tessedit_pageseg_mode: '7',
  });
  const result = await worker.recognize(cropCanvas);
  await worker.terminate();

  return (result.data.text || '').replace(/\D/g, '');
}

async function ocrTinBoxDigits(canvas) {
  const boxRegionCanvas = cropCanvasRegion(canvas, TIN_BOX_REGION, 2);
  if (!boxRegionCanvas) return { ssn: '', ein: '' };

  const ssnStrip = cropCanvasRegion(boxRegionCanvas, SSN_ROW_FRAC, 2);
  const einStrip = cropCanvasRegion(boxRegionCanvas, EIN_ROW_FRAC, 2);

  const [ssn, ein] = await Promise.all([ocrDigitsOnly(ssnStrip), ocrDigitsOnly(einStrip)]);

  return { ssn, ein };
}

async function ocrPdfText(file, onProgress) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const pagesToScan = Math.min(pdf.numPages, 2); // W-9's fields all live on page 1
  let combined = '';
  let tinDigits = { ssn: '', ein: '' };

  for (let i = 1; i <= pagesToScan; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 3 }); // higher res improves OCR accuracy, esp. on small boxed digits
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;

    const result = await Tesseract.recognize(canvas, 'eng', {
      logger: onProgress ? (m) => onProgress(i, pagesToScan, m) : undefined,
    });
    combined += result.data.text + '\n';

    if (i === 1) {
      tinDigits = await ocrTinBoxDigits(canvas);
    }
  }

  return { text: combined, tinDigits };
}

// Builds a regex that matches a phrase loosely: whitespace between words may
// be missing or doubled and commas/periods/colons may be missing entirely,
// which is common OCR noise (e.g. "List account number" comes back as
// "Listaccount number", "street, and apt." loses its comma). Punctuation
// should be omitted from `phrase` itself — the connector already allows it.
function loosePattern(phrase) {
  return phrase
    .split(/\s+/)
    .map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
    .join('\\s*[,.:]?\\s*');
}

function loose(phrase, flags = 'i') {
  return new RegExp(loosePattern(phrase), flags);
}

// Any of these appearing marks the start of the *next* labeled field, so a
// capture should stop there regardless of which field it started from —
// bounds every capture even when its own "expected" stop label is missing
// or worded slightly differently across W-9 template variants.
const LINE_NUM_PREFIX = '(?:\\b\\d{1,2}[a-z]?\\s+)?';
const NEXT_LABEL_PHRASES = [
  'Business name/disregarded entity name',
  'Federal tax classification',
  'Check the appropriate',
  'Exemptions',
  'Address (number',
  // 'City state and ZIP code' deliberately excluded — the combined address
  // capture below is meant to span straight through that label (it sits
  // between W-9 lines 5 and 6, both part of the same mailing address), not
  // stop at it.
  'Social security number',
  'Employer identification number',
  'Part I',
  'Part II',
  'List account number',
];
const NEXT_LABEL_STOP = new RegExp(
  LINE_NUM_PREFIX + '(?:' + NEXT_LABEL_PHRASES.map(loosePattern).join('|') + ')',
  'i'
);

// Strips boilerplate that can sit between a field's label and its actual
// filled-in value: parenthetical instructions (e.g. "(For a sole proprietor
// ...)"), a stray "See instructions." aside, and the punctuation left
// dangling around either of those.
function cleanCapture(str) {
  let s = str.trim();
  s = s.replace(/^\([^)]*\)\s*/, '');
  s = s.replace(/^[.:\-–]\s*/, '');
  s = s.replace(/^See instructions\.?\s*/i, '');
  s = s.replace(/^[.:\-–]\s*/, '');
  return s.trim();
}

function captureAfterLabel(text, labelRegex, stopRegex) {
  const labelMatch = text.match(labelRegex);
  if (!labelMatch) return '';
  const rest = text.slice(labelMatch.index + labelMatch[0].length);

  const stopCandidates = [stopRegex, NEXT_LABEL_STOP]
    .map((re) => rest.match(re))
    .filter(Boolean)
    .sort((a, b) => a.index - b.index);

  const cutoff = stopCandidates.length ? stopCandidates[0].index : 160;
  return cleanCapture(rest.slice(0, Math.min(cutoff, 160)));
}

// SSN/EIN are entered into individual boxed digit cells on the real form, so
// OCR (and even some text layers) often reads them with stray spaces, dashes,
// or box-border noise between digits — "8 8 - 4 1 6 3 1 3 3" rather than a
// clean "88-4163133". Anchoring to the label and pulling only the digit
// characters out of the text right after it survives that noise; a strict
// contiguous-format match is kept as a fallback for already-clean text.
function digitsAfterLabel(text, labelRegex, windowChars) {
  const m = text.match(labelRegex);
  if (!m) return '';
  const start = m.index + m[0].length;
  return text.slice(start, start + windowChars).replace(/\D/g, '');
}

function parseW9Fields(rawText, tinDigits) {
  const text = rawText.replace(/\s+/g, ' ').trim();

  let selectedType = null;
  let taxId = '';

  // Digits read from a targeted, cropped re-OCR of the TIN box grid (see
  // ocrTinBoxDigits) are far more trustworthy than anything found by
  // searching the whole-page text, since that box grid routinely OCRs as
  // garbage in a full-page pass. Prefer them when available.
  const ssnDigits = (tinDigits && tinDigits.ssn) || digitsAfterLabel(text, loose('Social security number'), 80);
  const einDigits = (tinDigits && tinDigits.ein) || digitsAfterLabel(text, loose('Employer identification number'), 80);

  if (ssnDigits.length >= 9) {
    const d = ssnDigits.slice(0, 9);
    selectedType = 'SSN';
    taxId = `${d.slice(0, 3)}-${d.slice(3, 5)}-${d.slice(5)}`;
  } else if (einDigits.length >= 9) {
    const d = einDigits.slice(0, 9);
    selectedType = 'EIN';
    taxId = `${d.slice(0, 2)}-${d.slice(2)}`;
  } else {
    const ssnMatch = text.match(/\b(\d{3}-\d{2}-\d{4})\b/);
    const einMatch = text.match(/\b(\d{2}-\d{7})\b/);
    if (ssnMatch) {
      selectedType = 'SSN';
      taxId = ssnMatch[1];
    } else if (einMatch) {
      selectedType = 'EIN';
      taxId = einMatch[1];
    }
  }

  const companyName =
    captureAfterLabel(
      text,
      loose('Business name/disregarded entity name if different from above'),
      new RegExp(LINE_NUM_PREFIX + loosePattern('Address (number'), 'i')
    ) ||
    captureAfterLabel(
      text,
      new RegExp(
        loosePattern('Name (as shown on your income tax return)') +
          '|' +
          loosePattern('Name of entity/individual'),
        'i'
      ),
      loose('Business name/disregarded entity name')
    );

  // W-9 lines 5 ("Address") and 6 ("City, state, and ZIP code") are the two
  // halves of one mailing address. Rather than capturing each line
  // independently (which silently drops line 6 whenever its label doesn't
  // OCR/extract cleanly), grab everything from after the line-5 label
  // through to the SSN/EIN section in one pass — that span covers both
  // lines' filled-in values regardless of whether the line-6 label survived
  // in between — then strip that label out if it's present and split the
  // combined block back into street/city/state+ZIP for the individual
  // fields used elsewhere (debug view, MSA replacement building).
  const addressBlock = captureAfterLabel(
    text,
    new RegExp(
      loosePattern('Address (number street and apt or suite no') + '\\.?\\)?\\.?(?:\\s*See\\s*instructions\\.?)?',
      'i'
    ),
    new RegExp(
      LINE_NUM_PREFIX + '(?:' + loosePattern('Social security number') + '|' + loosePattern('Employer identification number') + ')',
      'i'
    )
  );

  // `text` had all whitespace (including the real line break between W-9
  // lines 5 and 6) collapsed to single spaces at the top of this function,
  // so once the "City, state, and ZIP code" label is stripped down to a
  // plain space, nothing distinguishes "...Terrace" (end of line 5) from
  // "Fremont..." (start of line 6) — a regex split on the fused text can't
  // tell them apart and mis-splits (e.g. swallowing both into "city").
  // Marking the exact label position *before* collapsing further, then
  // splitting on that marker, keeps the one boundary that actually matters.
  const ADDRESS_LINE_MARKER = '';
  const withMarker = addressBlock
    // The real form prints "Requester's name and address (optional)" as a
    // second-column header immediately to the right of line 5's label, so
    // it lands right at the start of the captured text (before the actual
    // street value, which is on the next visual line) — strip it out too.
    .replace(new RegExp('^\\s*Requester\'?s?\\s*name\\s*and\\s*address\\s*\\(optional\\)\\.?', 'i'), ' ')
    .replace(new RegExp(LINE_NUM_PREFIX + loosePattern('City state and ZIP code') + '\\.?', 'i'), ADDRESS_LINE_MARKER);

  const [rawStreetPart, rawCityPart = ''] = withMarker.split(ADDRESS_LINE_MARKER);
  let streetAddress = cleanCapture(rawStreetPart.replace(/\s+/g, ' ').trim());
  let city = '';
  let stateZip = '';

  const cityPart = cleanCapture(rawCityPart.replace(/\s+/g, ' ').trim());
  if (cityPart) {
    // Line 6's own content is short and unambiguous once isolated, so a
    // simple "<city>, <ST> <ZIP>" match (comma optional, since OCR often
    // drops it) is reliable here in a way it wasn't across the whole block.
    const cityStateZipMatch = cityPart.match(/^([A-Za-z][A-Za-z .'-]*?),?\s+([A-Za-z]{2}\s*\d{5}(?:-\d{4})?)\s*$/);
    if (cityStateZipMatch) {
      city = cityStateZipMatch[1].trim();
      stateZip = cityStateZipMatch[2].trim();
    } else {
      city = cityPart;
    }
  } else if (!withMarker.includes(ADDRESS_LINE_MARKER)) {
    // Line 6's label never extracted at all, so there was no marker to split
    // on — fall back to a best-effort match against the whole (ambiguous)
    // block rather than losing whatever came after the street. Strip the
    // state+ZIP suffix first, then prefer the LAST comma as the street/city
    // boundary (far less ambiguous than guessing where a multi-word city
    // name starts); only fall back to treating just the final word as the
    // city when OCR dropped that comma too.
    const stateZipSuffix = streetAddress.match(/([A-Za-z]{2}\s*\d{5}(?:-\d{4})?)\s*$/);
    if (stateZipSuffix) {
      stateZip = stateZipSuffix[1].trim();
      const beforeStateZip = streetAddress.slice(0, stateZipSuffix.index).trim().replace(/,\s*$/, '');
      const lastComma = beforeStateZip.lastIndexOf(',');
      if (lastComma !== -1) {
        streetAddress = beforeStateZip.slice(0, lastComma).trim();
        city = beforeStateZip.slice(lastComma + 1).trim();
      } else {
        const words = beforeStateZip.split(' ');
        city = words.pop() || '';
        streetAddress = words.join(' ');
      }
    }
  }

  return {
    company_name: companyName,
    selected_tax_id_type: selectedType,
    tax_id_number: taxId,
    street_address: streetAddress,
    city,
    state_zip: stateZip,
  };
}

function renderW9Summary(data) {
  w9CompanyName.value = data.company_name || '';
  w9TaxId.value = data.tax_id_number || '';
  w9TaxIdType.value = data.selected_tax_id_type || 'EIN';
  const addressParts = [data.street_address, data.city, data.state_zip].filter(Boolean);
  w9Address.value = addressParts.join(', ');
  w9Summary.hidden = false;
  w9EditHint.hidden = false;
}

const HAS_TEXT_THRESHOLD = 20; // chars; below this, treat the PDF as having no usable text layer

function showRawText(text) {
  w9RawText.textContent = text;
  w9Debug.hidden = false;
}

function finishExtraction(text, viaOcr, tinDigits) {
  extractedW9Data = parseW9Fields(text, tinDigits);
  renderW9Summary(extractedW9Data);
  showRawText(text);

  const source = viaOcr ? ' (via OCR)' : '';
  if (!extractedW9Data.selected_tax_id_type) {
    setStatus(`W-9 read${source}, but no SSN or EIN was found — fill it in manually before generating.`, true);
  } else {
    setStatus(
      `W-9 parsed${source}. Tax ID used: ${extractedW9Data.selected_tax_id_type} (${extractedW9Data.tax_id_number}).`
    );
  }
}

async function processW9(file) {
  if (!window.pdfjsLib) {
    setStatus('PDF parser failed to load — check your connection and try again.', true);
    return;
  }

  setStatus(`Reading ${file.name}...`);
  try {
    const text = await extractPdfText(file);
    if (text.replace(/\s+/g, '').length >= HAS_TEXT_THRESHOLD) {
      finishExtraction(text, false);
      return;
    }

    if (!window.Tesseract) {
      extractedW9Data = null;
      setStatus(
        `${file.name} has no readable text layer (likely a scanned image), and the OCR engine failed to load — please fill fields manually.`,
        true
      );
      return;
    }

    setStatus(`${file.name} looks scanned — running OCR (this can take a few seconds)...`);
    const { text: ocrText, tinDigits } = await ocrPdfText(file, (page, totalPages, m) => {
      if (m.status === 'recognizing text') {
        setStatus(`Running OCR on page ${page}/${totalPages}... ${Math.round(m.progress * 100)}%`);
      }
    });

    if (ocrText.replace(/\s+/g, '').length < HAS_TEXT_THRESHOLD) {
      extractedW9Data = null;
      setStatus(`OCR could not read any text from ${file.name} — please fill fields manually.`, true);
      return;
    }

    finishExtraction(ocrText, true, tinDigits);
  } catch (err) {
    extractedW9Data = null;
    setStatus(`Could not read ${file.name} as a W-9 PDF: ${err.message}`, true);
  }
}

/* ---------- Status helper ---------- */
function setStatus(message, isError = false) {
  status.textContent = message;
  status.classList.toggle('error', isError);

  // Restart the settle-in animation on every update instead of just popping the new text in
  status.classList.remove('flash');
  void status.offsetWidth;
  status.classList.add('flash');

  if (!wizardLoading.hidden) {
    wizardLoadingText.textContent = message;
  }
}

/* ---------- Generate: locate literal text in the MSA template and replace it ----------
   Real MSA templates are neither fillable PDF forms nor {{merge-tag}} Word
   docs — they're ordinary contracts where the contractor's name/tax id/
   address appear as plain repeated text across many pages. So instead of
   matching form fields or tags, this locates the *previous* contractor's
   placeholder text (however many times it appears) and redraws/rewrites
   the new value in its place. */

// The standard MSA template reused for every contractor keeps the last
// contractor's info as the literal placeholder text to search for and swap
// out — these are Aptiva Corp's values in the template this was built against.
const MSA_TEMPLATE_PLACEHOLDERS = {
  companyName: 'APTIVA CORP',
  taxId: '26-1282577',
  address: '825 Georges Rd, Suite 1, North Brunswick, NJ, 08902',
};

const MANUAL_FIELD_LABELS = {
  rep: 'Contractor Representative(s):',
  role: 'Role:',
  location: 'Location:',
  startDate: 'Start Date:',
  billingRate: 'Billing Rate:',
};

async function toArrayBuffer(source) {
  if (source instanceof ArrayBuffer) return source;
  if (source instanceof Uint8Array) return source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength);
  return source.arrayBuffer(); // File or Blob
}

function isDocxFile(file) {
  return (
    /\.docx$/i.test(file.name) ||
    file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  );
}

// --- Word (.docx): text lives in <w:t> runs inside the XML, so this can do
// a real find/replace — a match spanning multiple runs (common when a Word
// doc has kept separate runs for formatting reasons) gets spliced across them. ---
function findRunSpan(texts, start, len) {
  let pos = 0;
  let startRun = -1;
  let startOffset = 0;
  let endRun = -1;
  let endOffset = 0;
  for (let i = 0; i < texts.length; i++) {
    const runLen = texts[i].length;
    if (startRun === -1 && start < pos + runLen) {
      startRun = i;
      startOffset = start - pos;
    }
    if (startRun !== -1 && start + len <= pos + runLen) {
      endRun = i;
      endOffset = start + len - pos;
      break;
    }
    pos += runLen;
  }
  return { startRun, startOffset, endRun, endOffset };
}

function replaceAllInRuns(runs, find, value) {
  let count = 0;
  for (;;) {
    const texts = runs.map((r) => r.textContent || '');
    const combined = texts.join('');
    const idx = combined.indexOf(find);
    if (idx === -1) break;

    const { startRun, startOffset, endRun, endOffset } = findRunSpan(texts, idx, find.length);
    if (startRun === -1 || endRun === -1) break;

    if (startRun === endRun) {
      const t = texts[startRun];
      runs[startRun].textContent = t.slice(0, startOffset) + value + t.slice(endOffset);
    } else {
      runs[startRun].textContent = texts[startRun].slice(0, startOffset) + value;
      for (let i = startRun + 1; i < endRun; i++) runs[i].textContent = '';
      runs[endRun].textContent = texts[endRun].slice(endOffset);
    }
    count += 1;
    if (count > 200) break; // safety valve against a runaway loop
  }
  return count;
}

function replaceAfterLabelInParagraphs(paragraphs, label, value) {
  let count = 0;
  paragraphs.forEach((p) => {
    const runs = Array.from(p.getElementsByTagName('w:t'));
    if (!runs.length) return;
    const texts = runs.map((r) => r.textContent || '');
    const combined = texts.join('');
    const idx = combined.indexOf(label);
    if (idx === -1) return;

    let { startRun, startOffset } = findRunSpan(texts, idx + label.length, 0);
    if (startRun === -1) return;

    // Word often keeps a label's trailing space in the same (bold) run as
    // the label itself — e.g. one run holding "Billing Rate: " — so the
    // position right after the label can still land inside that run rather
    // than the plain run the actual value lives in. Splicing the new value
    // in there would make it inherit the label's bold formatting. Leave any
    // such whitespace-only remainder where it is and advance into the next
    // run(s) until landing on one that actually holds value content.
    while (startRun + 1 < runs.length && /^\s*$/.test(texts[startRun].slice(startOffset))) {
      startRun += 1;
      startOffset = 0;
    }

    runs[startRun].textContent = texts[startRun].slice(0, startOffset) + value;
    for (let i = startRun + 1; i < runs.length; i++) runs[i].textContent = '';
    count += 1;
  });
  return count;
}

async function applyDocxReplacements(source, replacements) {
  const arrayBuffer = await toArrayBuffer(source);
  const zip = new PizZip(arrayBuffer);
  const xml = zip.file('word/document.xml').asText();
  const xmlDoc = new DOMParser().parseFromString(xml, 'application/xml');

  const allRuns = Array.from(xmlDoc.getElementsByTagName('w:t'));
  const paragraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));

  let replacedCount = 0;
  replacements.forEach((r) => {
    if (!r.value) return;
    if (r.type === 'value') {
      replacedCount += replaceAllInRuns(allRuns, r.find, r.value);
    } else if (r.type === 'afterLabel') {
      replacedCount += replaceAfterLabelInParagraphs(paragraphs, r.label, r.value);
    }
  });

  zip.file('word/document.xml', new XMLSerializer().serializeToString(xmlDoc));
  const bytes = zip.generate({ type: 'uint8array' });
  return { bytes, replacedCount };
}

// Read-only occurrence count for the Step 2 preview — doesn't touch the
// document, just reports how many "value" placeholders exist so the status
// message has something concrete before the real edit happens in Step 3.
async function countDocxOccurrences(source, replacements) {
  const arrayBuffer = await toArrayBuffer(source);
  const zip = new PizZip(arrayBuffer);
  const xml = zip.file('word/document.xml').asText();
  const xmlDoc = new DOMParser().parseFromString(xml, 'application/xml');
  const combined = Array.from(xmlDoc.getElementsByTagName('w:t'))
    .map((r) => r.textContent || '')
    .join('');

  return replacements.reduce((sum, r) => {
    if (r.type !== 'value' || !r.find) return sum;
    return sum + findAllIndices(combined, r.find).length; // findAllIndices is defined below (hoisted)
  }, 0);
}

// --- PDF: there's no editable text layer to splice, so this locates each
// match's bounding box via pdf.js (which shares PDF's coordinate space with
// pdf-lib), paints over it, and draws the new value in roughly the same spot
// and size. Character-level offsets within a text run are approximated
// proportionally, since pdf.js only reports whole-run boxes. ---
function findAllIndices(haystack, needle) {
  const indices = [];
  if (!needle) return indices;
  let from = 0;
  for (;;) {
    const idx = haystack.indexOf(needle, from);
    if (idx === -1) break;
    indices.push(idx);
    from = idx + needle.length;
  }
  return indices;
}

function coveringRect(items, texts, start, len) {
  let pos = 0;
  let x0 = null;
  let x1 = null;
  let y0 = null;
  let y1 = null;

  for (let i = 0; i < items.length; i++) {
    const t = texts[i];
    const itemStart = pos;
    const itemEnd = pos + t.length;
    pos = itemEnd;
    if (itemEnd <= start || itemStart >= start + len) continue;

    const item = items[i];
    const [a, , , d, e, f] = item.transform;
    const height = item.height || Math.abs(d) || 10;
    const width = item.width || Math.abs(a) * t.length;

    const overlapStart = Math.max(start, itemStart) - itemStart;
    const overlapEnd = Math.min(start + len, itemEnd) - itemStart;
    const fracStart = t.length ? overlapStart / t.length : 0;
    const fracEnd = t.length ? overlapEnd / t.length : 1;

    const itemX0 = e + fracStart * width;
    const itemX1 = e + fracEnd * width;

    if (x0 === null || itemX0 < x0) x0 = itemX0;
    if (x1 === null || itemX1 > x1) x1 = itemX1;
    if (y0 === null || f < y0) y0 = f;
    if (y1 === null || f + height > y1) y1 = f + height;
  }

  if (x0 === null) return null;
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
}

async function locatePdfReplacements(arrayBuffer, replacements) {
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer.slice(0) }).promise;
  const matches = [];

  for (let pageIndex = 0; pageIndex < pdf.numPages; pageIndex++) {
    const page = await pdf.getPage(pageIndex + 1);
    const textContent = await page.getTextContent();
    const items = textContent.items;
    const texts = items.map((it) => it.str);
    const combined = texts.join('');

    replacements.forEach((r) => {
      if (!r.value) return;

      if (r.type === 'value') {
        findAllIndices(combined, r.find).forEach((idx) => {
          const rect = coveringRect(items, texts, idx, r.find.length);
          if (rect) matches.push({ pageIndex, ...rect, value: r.value });
        });
        return;
      }

      // afterLabel: cover the rest of the label's own row (pdf.js has no
      // notion of line breaks, so "same row" is approximated by baseline y).
      const idx = combined.indexOf(r.label);
      if (idx === -1) return;
      const afterStart = idx + r.label.length;

      let rowY = null;
      let rowHeight = 10;
      let pos = 0;
      for (let i = 0; i < items.length; i++) {
        const itemStart = pos;
        pos += texts[i].length;
        if (afterStart >= itemStart && afterStart < pos) {
          rowY = Math.round(items[i].transform[5]);
          rowHeight = items[i].height || 10;
          break;
        }
      }
      if (rowY === null) return;

      // Ordinal suffixes ("5th", "31st") are drawn as superscripts a few
      // points off the row's baseline — well under a full line height —
      // so only a drop of more than one line height counts as "next line".
      // (PDF y increases upward, so a lower line has a *smaller* y.)
      let rowEnd = combined.length;
      pos = 0;
      for (let i = 0; i < items.length; i++) {
        const itemStart = pos;
        pos += texts[i].length;
        if (itemStart < afterStart) continue;
        const itemY = Math.round(items[i].transform[5]);
        if (rowY - itemY > rowHeight * 1.2) {
          rowEnd = itemStart;
          break;
        }
      }

      const rect = coveringRect(items, texts, afterStart, rowEnd - afterStart);
      if (rect) matches.push({ pageIndex, ...rect, value: r.value });
    });
  }

  return matches;
}

// Read-only occurrence count for the Step 2 preview — see countDocxOccurrences.
async function countPdfOccurrences(source, replacements) {
  const arrayBuffer = await toArrayBuffer(source);
  const matches = await locatePdfReplacements(arrayBuffer, replacements);
  return matches.length;
}

async function rasterizePageToPng(pdfjsDoc, pageIndex, scale) {
  const page = await pdfjsDoc.getPage(pageIndex + 1);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;

  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  return blob.arrayBuffer();
}

// Painting a white box over old text and drawing new text on top of it
// covers the OLD text visually, but pdf-lib only ever *adds* content — the
// original text-drawing instructions stay in the page's content stream and
// are still fully recoverable by copy/paste or programmatic text extraction.
// That's a real data leak here specifically, since what's being covered up
// is the previous contractor's tax ID, company name, and personal details.
// The only reliable fix without a backend PDF-editing service is to
// rasterize any page that has a replacement on it — pdf.js renders it to a
// PNG (turning its old text into plain pixels with no text layer at all),
// pdf-lib rebuilds that page as just that image, and only then does the
// white box + new text get drawn on top. Pages with no replacements are
// left completely untouched (still real, selectable text).
async function applyPdfReplacements(source, replacements) {
  const arrayBuffer = await toArrayBuffer(source);
  const matches = await locatePdfReplacements(arrayBuffer, replacements);
  if (!matches.length) {
    return { bytes: new Uint8Array(arrayBuffer), replacedCount: 0 };
  }

  const pagesWithMatches = [...new Set(matches.map((m) => m.pageIndex))];

  const pdfjsDoc = await pdfjsLib.getDocument({ data: arrayBuffer.slice(0) }).promise;
  const RASTER_SCALE = 2; // higher = crisper re-render, bigger file
  const pageImageBytes = new Map();
  for (const pageIndex of pagesWithMatches) {
    pageImageBytes.set(pageIndex, await rasterizePageToPng(pdfjsDoc, pageIndex, RASTER_SCALE));
  }

  const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
  const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);

  for (const pageIndex of pagesWithMatches) {
    const { width, height } = pdfDoc.getPage(pageIndex).getSize();
    const image = await pdfDoc.embedPng(pageImageBytes.get(pageIndex));

    pdfDoc.removePage(pageIndex);
    const newPage = pdfDoc.insertPage(pageIndex, [width, height]);
    newPage.drawImage(image, { x: 0, y: 0, width, height });

    matches
      .filter((m) => m.pageIndex === pageIndex)
      .forEach((m) => {
        newPage.drawRectangle({
          x: m.x - 1,
          y: m.y - 2,
          width: Math.max(m.width, 4) + 4,
          height: m.height + 4,
          color: PDFLib.rgb(1, 1, 1),
        });
        newPage.drawText(m.value, {
          x: m.x,
          y: m.y + m.height * 0.15,
          size: Math.max(7, m.height * 0.78),
          font,
          color: PDFLib.rgb(0, 0, 0),
        });
      });
  }

  const bytes = await pdfDoc.save();
  return { bytes, replacedCount: matches.length };
}

// The "other" format (whichever wasn't uploaded) can't be a real conversion
// of the template without a backend — these build a plain document from the
// same extracted + manually-entered values instead, so a download link
// always has something real to offer, just without the original template's
// layout/design.
function summaryLines(data) {
  const lines = [
    'MASTER SERVICE AGREEMENT',
    '',
    `Company / Business Name: ${data?.company_name || '(not found)'}`,
    `Tax ID Used: ${data?.selected_tax_id_type || '(none found)'}${data?.tax_id_number ? ` (${data.tax_id_number})` : ''}`,
    `Street Address: ${data?.street_address || '(not found)'}`,
    `City: ${data?.city || '(not found)'}`,
    `State / ZIP: ${data?.state_zip || '(not found)'}`,
  ];

  if (data?.manual_rep) lines.push(`Contractor Representative(s): ${data.manual_rep}`);
  if (data?.manual_role) lines.push(`Role: ${data.manual_role}`);
  if (data?.manual_location) lines.push(`Location: ${data.manual_location}`);
  if (data?.manual_start_date) lines.push(`Start Date: ${data.manual_start_date}`);
  if (data?.manual_billing_rate) lines.push(`Billing Rate: ${data.manual_billing_rate}`);

  return lines;
}

async function buildSimplePdf(data) {
  const pdfDoc = await PDFLib.PDFDocument.create();
  const page = pdfDoc.addPage([612, 792]);
  const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
  const boldFont = await pdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);

  let y = 740;
  summaryLines(data).forEach((line, i) => {
    const isTitle = i === 0;
    page.drawText(line, { x: 72, y, size: isTitle ? 16 : 11, font: isTitle ? boldFont : font });
    y -= isTitle ? 30 : 22;
  });

  const bytes = await pdfDoc.save();
  return new Blob([bytes], { type: 'application/pdf' });
}

async function buildSimpleDocx(data) {
  const lines = summaryLines(data);
  const doc = new Document({
    sections: [
      {
        children: lines.map(
          (line, i) =>
            new Paragraph({
              children: [new TextRun({ text: line, bold: i === 0, size: i === 0 ? 32 : 24 })],
            })
        ),
      },
    ],
  });
  return Packer.toBlob(doc);
}

function setGeneratedFiles(pdfBlob, wordBlob, baseName) {
  if (generatedPdfUrl) URL.revokeObjectURL(generatedPdfUrl);
  if (generatedWordUrl) URL.revokeObjectURL(generatedWordUrl);

  generatedPdfUrl = URL.createObjectURL(pdfBlob);
  generatedWordUrl = URL.createObjectURL(wordBlob);
  generatedBaseName = baseName;

  downloadPdfBtn.disabled = false;
  downloadWordBtn.disabled = false;

  saveToHistory(pdfBlob, wordBlob, baseName).catch((err) => {
    console.error('Could not save generated document to history:', err);
  });
}

// Shared between Step 2's preview count and Step 3's real, final edit — both
// need the identical replacement spec so the count the user sees matches
// what actually gets changed.
function buildAutoReplacements(data) {
  return [
    { type: 'value', find: MSA_TEMPLATE_PLACEHOLDERS.companyName, value: data?.company_name },
    { type: 'value', find: MSA_TEMPLATE_PLACEHOLDERS.taxId, value: data?.tax_id_number },
    {
      type: 'value',
      find: MSA_TEMPLATE_PLACEHOLDERS.address,
      value: [data?.street_address, data?.city, data?.state_zip].filter(Boolean).join(', '),
    },
  ];
}

// Step 2 only previews how many placeholder occurrences exist — it doesn't
// edit anything yet. The real edit happens once, in Step 3, combining these
// auto values with the manually-entered ones in a single pass. Doing it in
// two separate passes would break the PDF path: Step 2 rasterizes any page
// it touches (see applyPdfReplacements), so if a manual-field label shares a
// page with the company name (as it does on this template's last page,
// where "APTIVA CORP" appears again in the signature block), that label's
// real text would already be gone — turned into pixels — by the time Step 3
// went looking for it.
generateBtn.addEventListener('click', async () => {
  if (generateBtn.classList.contains('is-loading')) return;

  msaFileList.querySelectorAll('li').forEach((li) => li.classList.remove('field-error'));

  if (!msaTemplateFile) {
    msaUploadZone.classList.add('field-error');
    setStatus('Please upload an MSA template (PDF or Word) before inserting data.', true);
    return;
  }
  msaUploadZone.classList.remove('field-error');

  const startedAt = Date.now();
  step2NextBtn.disabled = true;
  generateBtn.classList.add('is-loading');
  generateBtn.disabled = true;
  generateBtn.querySelector('.btn-label').textContent = 'Scanning...';
  showWizardLoading('Scanning the MSA template for the placeholder text to replace...');

  let success = false;
  try {
    const autoReplacements = buildAutoReplacements(extractedW9Data);
    isDocxTemplate = isDocxFile(msaTemplateFile);

    const occurrences = isDocxTemplate
      ? await countDocxOccurrences(msaTemplateFile, autoReplacements)
      : await countPdfOccurrences(msaTemplateFile, autoReplacements);

    success = occurrences > 0;
    msaFileList.querySelectorAll('li').forEach((li) => li.classList.toggle('field-error', !success));
    setStatus(
      success
        ? `Found ${occurrences} occurrence(s) of the template's placeholder company/tax ID/address. Continuing to manual details.`
        : "This template doesn't contain the expected placeholder text, so nothing would be auto-replaced — check the uploaded file.",
      !success
    );
  } catch (err) {
    msaFileList.querySelectorAll('li').forEach((li) => li.classList.add('field-error'));
    setStatus(`Could not scan the MSA template: ${err.message}`, true);
  } finally {
    await waitRemaining(startedAt, MIN_LOADING_MS);
    hideWizardLoading();
    generateBtn.classList.remove('is-loading');
    generateBtn.disabled = false;
    generateBtn.querySelector('.btn-label').textContent = 'Insert Data';
    step2NextBtn.disabled = !success;
    if (success) showWizardStep(3);
  }
});

// Missing manual fields get the same red-outline treatment as Step 1 —
// cleared as soon as the user fills them in, without needing to re-click.
const MANUAL_FIELD_INPUTS = [manualRep, manualRole, manualLocation, manualStartDate, manualBillingRate];

function validateManualFields() {
  let valid = true;
  MANUAL_FIELD_INPUTS.forEach((el) => {
    const ok = !!el.value.trim();
    el.classList.toggle('field-error', !ok);
    if (!ok) valid = false;
  });
  return valid;
}

MANUAL_FIELD_INPUTS.forEach((el) => {
  el.addEventListener('input', () => {
    if (el.value.trim()) el.classList.remove('field-error');
  });
});

/* ---------- Apply manual details (Step 3): the one real, final edit ---------- */
applyManualBtn.addEventListener('click', async () => {
  const startedAt = Date.now();
  downloadPdfBtn.disabled = true;
  downloadWordBtn.disabled = true;
  step3NextBtn.disabled = true;
  applyManualBtn.disabled = true;
  showWizardLoading('Applying details and finalizing the MSA...');

  let success = false;
  try {
    if (!validateManualFields()) {
      setStatus('Please fill in all manual details before applying.', true);
      return;
    }

    const manualValues = {
      rep: manualRep.value.trim(),
      role: manualRole.value.trim(),
      location: manualLocation.value.trim(),
      startDate: manualStartDate.value.trim(),
      // The field only ever collects the bare number — "$" and "/hr" are
      // fixed UI decoration, not something the user can mistype — so they're
      // added back on here to reconstruct the value actually inserted into
      // the document (e.g. "70" -> "$70/hr").
      billingRate: manualBillingRate.value.trim() ? `$${manualBillingRate.value.trim()}/hr` : '',
    };

    const allReplacements = [
      ...buildAutoReplacements(extractedW9Data),
      { type: 'afterLabel', label: MANUAL_FIELD_LABELS.rep, value: manualValues.rep },
      { type: 'afterLabel', label: MANUAL_FIELD_LABELS.role, value: manualValues.role },
      { type: 'afterLabel', label: MANUAL_FIELD_LABELS.location, value: manualValues.location },
      { type: 'afterLabel', label: MANUAL_FIELD_LABELS.startDate, value: manualValues.startDate },
      { type: 'afterLabel', label: MANUAL_FIELD_LABELS.billingRate, value: manualValues.billingRate },
    ];

    const finalData = {
      ...extractedW9Data,
      manual_rep: manualValues.rep,
      manual_role: manualValues.role,
      manual_location: manualValues.location,
      manual_start_date: manualValues.startDate,
      manual_billing_rate: manualValues.billingRate,
    };

    const baseName = msaTemplateFile.name.replace(/\.(pdf|docx)$/i, '');

    if (isDocxTemplate) {
      const { bytes } = await applyDocxReplacements(msaTemplateFile, allReplacements);
      const wordBlob = new Blob([bytes], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      const pdfBlob = await buildSimplePdf(finalData);
      setGeneratedFiles(pdfBlob, wordBlob, baseName);
    } else {
      const { bytes } = await applyPdfReplacements(msaTemplateFile, allReplacements);
      const pdfBlob = new Blob([bytes], { type: 'application/pdf' });
      const wordBlob = await buildSimpleDocx(finalData);
      setGeneratedFiles(pdfBlob, wordBlob, baseName);
    }

    setStatus('Manual details applied and the MSA is finalized. Ready to download.');
    success = true;
  } catch (err) {
    setStatus(`Could not finalize the MSA: ${err.message}`, true);
  } finally {
    await waitRemaining(startedAt, MIN_LOADING_MS);
    hideWizardLoading();
    applyManualBtn.disabled = false;
    step3NextBtn.disabled = !success;
    if (success) showWizardStep(4);
  }
});

/* ---------- Download ---------- */
function downloadUrl(url, filename) {
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

downloadPdfBtn.addEventListener('click', () => {
  if (!generatedPdfUrl) return;
  downloadUrl(generatedPdfUrl, `${generatedBaseName}.pdf`);
});

downloadWordBtn.addEventListener('click', () => {
  if (!generatedWordUrl) return;
  downloadUrl(generatedWordUrl, `${generatedBaseName}.docx`);
});

/* ---------- History (generated documents, stored locally in this browser) ----------
   IndexedDB rather than localStorage — it can hold the Blobs directly instead of
   needing a base64 round-trip, and isn't capped at localStorage's ~5MB. */
const HISTORY_DB_NAME = 'documentGeneratorHistory';
const HISTORY_STORE = 'generatedDocuments';

function openHistoryDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(HISTORY_DB_NAME, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(HISTORY_STORE, { keyPath: 'id', autoIncrement: true });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveToHistory(pdfBlob, wordBlob, baseName) {
  const db = await openHistoryDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(HISTORY_STORE, 'readwrite');
    tx.objectStore(HISTORY_STORE).add({ baseName, pdfBlob, wordBlob, createdAt: Date.now() });
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function getHistoryEntries() {
  const db = await openHistoryDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(HISTORY_STORE, 'readonly').objectStore(HISTORY_STORE).getAll();
    request.onsuccess = () => resolve(request.result.sort((a, b) => b.createdAt - a.createdAt));
    request.onerror = () => reject(request.error);
  });
}

async function deleteHistoryEntry(id) {
  const db = await openHistoryDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(HISTORY_STORE, 'readwrite');
    tx.objectStore(HISTORY_STORE).delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

function formatHistoryDate(timestamp) {
  return new Date(timestamp).toLocaleString(undefined, {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  downloadUrl(url, filename);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function renderHistory() {
  const entries = await getHistoryEntries();
  historyList.innerHTML = '';

  if (!entries.length) {
    historyList.innerHTML = '<p class="history-empty">No documents generated yet on this device.</p>';
    return;
  }

  entries.forEach((entry) => {
    const item = document.createElement('div');
    item.className = 'history-item';
    item.innerHTML = `
      <div class="history-item-info">
        <span class="history-item-name">${entry.baseName}</span>
        <span class="history-item-date">${formatHistoryDate(entry.createdAt)}</span>
      </div>
      <div class="history-item-actions">
        <button type="button" class="btn btn-secondary" data-action="pdf">PDF</button>
        <button type="button" class="btn btn-secondary" data-action="word">Word</button>
        <button type="button" class="history-delete-btn" data-action="delete" aria-label="Delete this document">✕</button>
      </div>
    `;
    item.querySelector('[data-action="pdf"]').addEventListener('click', () => {
      downloadBlob(entry.pdfBlob, `${entry.baseName}.pdf`);
    });
    item.querySelector('[data-action="word"]').addEventListener('click', () => {
      downloadBlob(entry.wordBlob, `${entry.baseName}.docx`);
    });
    item.querySelector('[data-action="delete"]').addEventListener('click', async () => {
      await deleteHistoryEntry(entry.id);
      renderHistory();
    });
    historyList.appendChild(item);
  });
}

const mainStage = document.getElementById('mainStage');
const historyView = document.getElementById('historyView');
const historyList = document.getElementById('historyList');
const navGeneratorLink = document.getElementById('navGeneratorLink');
const navHistoryLink = document.getElementById('navHistoryLink');

navHistoryLink.addEventListener('click', (e) => {
  e.preventDefault();
  mainStage.hidden = true;
  historyView.hidden = false;
  navHistoryLink.classList.add('active');
  navGeneratorLink.classList.remove('active');
  renderHistory();
});

navGeneratorLink.addEventListener('click', (e) => {
  e.preventDefault();
  historyView.hidden = true;
  mainStage.hidden = false;
  navGeneratorLink.classList.add('active');
  navHistoryLink.classList.remove('active');
});

/* ---------- Intro animation: plays once, then hands off to the real page ---------- */
(function () {
  const overlay = document.getElementById('introOverlay');
  if (!overlay) return;

  const man = document.getElementById('man');
  const elevator = document.getElementById('elevator');

  document.body.style.overflow = 'hidden';

  function runIntroAnimation() {
    // 1. Elevator doors open, revealing the empty lift
    setTimeout(() => {
      elevator.classList.add('doors-open');
    }, 300);

    // 2. Once the doors are open, the man walks in. `walk-in` drives the
    // translateX (and holds position once arrived); `stepping` drives the
    // leg/arm/head motion and is removed the instant he arrives, so he
    // stands naturally still rather than continuing to march in place.
    setTimeout(() => {
      man.classList.add('walk-in', 'stepping');
    }, 1400);
    setTimeout(() => {
      man.classList.remove('stepping');
    }, 3600); // matches the 2.2s walk transition (1400ms + 2200ms)

    // 3. Man is inside — doors close
    setTimeout(() => {
      elevator.classList.remove('doors-open');
    }, 3700);

    // 4. Only after the doors have fully closed does the man disappear
    setTimeout(() => {
      man.classList.add('entered');
    }, 4750);

    // 5. Elevator morphs smoothly into the "GENERATE MSA" button
    setTimeout(() => {
      elevator.classList.add('morphed-button');
    }, 5300);

    // 6. After a beat to let the button register, it flies to its real place automatically
    setTimeout(() => {
      continueToApp();
    }, 6100);
  }

  // The button flies from its on-screen spot to the real Generate MSA button's
  // position first, against the still-opaque backdrop. Only once it has fully
  // landed does the backdrop fade away, revealing the real page underneath —
  // the page never appears mid-flight.
  //
  // Uses a single `transform` (translate + scale) instead of animating
  // left/top/width/height directly. Transform is GPU-composited and doesn't
  // trigger layout on every frame, so there's no risk of the browser batching
  // the "pin" and "fly" writes together and skipping straight to the end state
  // (which is what a left/top jump-to-start-of-transition looks like) — the
  // forced reflow between them guarantees the start state is committed first.
  function continueToApp() {
    const FLY_DURATION = 850;

    const startRect = elevator.getBoundingClientRect();

    // The nav logo is always on-screen (fixed top-left, never needs scrolling
    // into view like the real action buttons further down the page) — it's a
    // stable target to fly to.
    const targetRect = brandLogo.getBoundingClientRect();

    const dx = targetRect.left - startRect.left;
    const dy = targetRect.top - startRect.top;
    const sx = targetRect.width / startRect.width;
    const sy = targetRect.height / startRect.height;

    // Pin the button at exactly its current on-screen position/size, with an
    // identity transform and no transition yet — this is the animation's true start.
    elevator.style.position = 'fixed';
    elevator.style.margin = '0';
    elevator.style.left = startRect.left + 'px';
    elevator.style.top = startRect.top + 'px';
    elevator.style.width = startRect.width + 'px';
    elevator.style.height = startRect.height + 'px';
    elevator.style.transformOrigin = 'top left';
    elevator.style.transition = 'none';
    elevator.style.transform = 'translate(0px, 0px) scale(1, 1)';

    // Force the browser to commit the pinned state above before we change anything else,
    // so the upcoming transition animates FROM here — not from wherever it happened to be.
    void elevator.offsetWidth;

    // border-radius isn't animated here — it lives on the inner .lift-cab
    // (already at its pill shape from the morph), and since transform:scale
    // visually scales everything painted inside the element, its rounded
    // corners scale proportionally with the box automatically.
    elevator.style.transition = `transform ${FLY_DURATION}ms cubic-bezier(0.4, 0, 0.2, 1)`;
    elevator.style.transform = `translate(${dx}px, ${dy}px) scale(${sx}, ${sy})`;

    // Only once the button has actually arrived: reveal the real page and let
    // the traveling button dissolve into the real logo that's now visible underneath
    setTimeout(() => {
      overlay.classList.add('is-hidden');
      document.body.style.overflow = '';
      elevator.style.transition = 'opacity 0.35s ease';
      elevator.style.opacity = '0';

      brandLogo.classList.add('cta-pulse');
      setTimeout(() => brandLogo.classList.remove('cta-pulse'), 1600);
    }, FLY_DURATION);

    // Clean up once the backdrop fade has fully finished
    setTimeout(() => overlay.remove(), FLY_DURATION + 750);
  }

  // Plays once on load, then hands off to the real page automatically — no click needed
  runIntroAnimation();
})();
