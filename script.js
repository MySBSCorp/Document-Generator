import PizZip from './vendor/pizzip-3.2.0.esm.js';

import * as pdfjsLib from 'pdfjs-dist';
import pdfjsWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.js?url';
import Tesseract from 'tesseract.js';
import * as PDFLib from 'pdf-lib';
import { renderAsync as renderDocxAsync } from 'docx-preview';
import html2canvas from 'html2canvas';

if (window.top !== window.self) {
  window.top.location = window.self.location.href;
}

const uploadZone = document.querySelector('.upload-block');
const fileInput = document.getElementById('fileInput');
const fileList = document.getElementById('fileList');
const status = document.getElementById('status');
const downloadPdfBtn = document.getElementById('downloadPdfBtn');
const downloadWordBtn = document.getElementById('downloadWordBtn');
const brandLogo = document.querySelector('.brand');
const manualRep = document.getElementById('manualRep');
const manualRole = document.getElementById('manualRole');
const manualLocation = document.getElementById('manualLocation');
const manualStartDate = document.getElementById('manualStartDate');
const manualStartDatePicker = document.getElementById('manualStartDatePicker');
const manualStartDatePickerBtn = document.getElementById('manualStartDatePickerBtn');
const manualBillingRate = document.getElementById('manualBillingRate');
const applyManualBtn = document.getElementById('applyManualBtn');

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

function syncStartDatePickerFromText() {
  const text = manualStartDate.value.trim().replace(/(\d+)(st|nd|rd|th)\b/gi, '$1');
  const parsed = text ? new Date(text) : NaN;
  if (text && !Number.isNaN(parsed.getTime())) {
    const year = parsed.getFullYear();
    const month = String(parsed.getMonth() + 1).padStart(2, '0');
    const day = String(parsed.getDate()).padStart(2, '0');
    manualStartDatePicker.value = `${year}-${month}-${day}`;
  } else {
    manualStartDatePicker.value = '';
  }
}
manualStartDate.addEventListener('input', syncStartDatePickerFromText);

const wizard = document.getElementById('wizard');
const actionPanel = document.querySelector('.action-panel');
const wizardSteps = Array.from(wizard.querySelectorAll('.wizard-step'));
const wizardLoading = document.getElementById('wizardLoading');
const wizardLoadingText = document.getElementById('wizardLoadingText');
const extractDataBtn = document.getElementById('extractDataBtn');
const step1NextBtn = document.getElementById('step1NextBtn');
const step2BackBtn = document.getElementById('step2BackBtn');
const step3BackBtn = document.getElementById('step3BackBtn');
const startNewBtn = document.getElementById('startNewBtn');

let w9SummaryRevealed = false;
let w9DebugRevealed = false;

// Opt-in developer diagnostics for the W-9 extraction pipeline. Off by
// default (and never turned on automatically) so sensitive tax-document
// content never reaches the console/UI in normal use. Enable locally with:
//   localStorage.setItem('w9_debug', '1')
const W9_DEBUG = (() => {
  try {
    return localStorage.getItem('w9_debug') === '1';
  } catch {
    return false;
  }
})();

// Populated only when W9_DEBUG is on. Holds the anchor/value bboxes the
// dynamic detector used for the current document, purely for the debug
// overlay/trace — never contains extracted digits or address text.
let debugExtractionTrace = null;
function resetDebugTrace() {
  debugExtractionTrace = W9_DEBUG ? { tin: {}, address: {} } : null;
}

function showWizardStep(n) {
  wizardSteps.forEach((el) => {
    el.classList.toggle('active', Number(el.dataset.step) === n);
  });
  const showW9Panel = n < 3;
  w9Summary.hidden = !(showW9Panel && w9SummaryRevealed);
  w9EditHint.hidden = !(showW9Panel && w9SummaryRevealed);
  w9Debug.hidden = !(showW9Panel && w9DebugRevealed);
}

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

step2BackBtn.addEventListener('click', () => showWizardStep(1));
step3BackBtn.addEventListener('click', () => showWizardStep(2));

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
let msaTemplateDocxFile = null;

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorkerUrl;

document.querySelectorAll('.eyebrow, .hero h1, .subtitle').forEach((el) => {
  el.addEventListener('animationend', () => { el.style.animation = 'none'; }, { once: true });
});

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function renderFileList() {
  fileList.innerHTML = '';
  uploadedFiles.forEach((file) => {
    const li = document.createElement('li');
    const nameSpan = document.createElement('span');
    nameSpan.textContent = file.name;
    const sizeSpan = document.createElement('span');
    sizeSpan.className = 'size';
    sizeSpan.textContent = formatSize(file.size);
    li.append(nameSpan, sizeSpan);
    fileList.appendChild(li);
  });
}

function isW9UploadCandidate(f) {
  return f.type === 'application/pdf' || /\.pdf$/i.test(f.name) || isImageFile(f);
}

function addFiles(fileArray) {
  uploadedFiles = fileArray;
  renderFileList();
  setStatus(`${uploadedFiles.length} file(s) ready.`);

  extractedW9Data = null;
  w9SummaryRevealed = false;
  w9DebugRevealed = false;
  w9Summary.hidden = true;
  w9EditHint.hidden = true;
  w9Debug.hidden = true;
  clearDocumentPreview();

  const hasW9 = uploadedFiles.some(isW9UploadCandidate);
  extractDataBtn.disabled = !hasW9;
  step1NextBtn.disabled = true;
}

extractDataBtn.addEventListener('click', async () => {
  const w9File = uploadedFiles.find(isW9UploadCandidate);
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
  }

  if (success) await insertDataAndAdvance();
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

const MSA_TEMPLATE_DOCX_URL = '/assets/msa-template.docx';

async function loadFileAsset(url, name, type) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Could not load ${name}.`);
  const blob = await response.blob();
  return new File([blob], name, { type });
}

async function loadMsaTemplates() {
  if (!msaTemplateDocxFile) {
    msaTemplateDocxFile = await loadFileAsset(
      MSA_TEMPLATE_DOCX_URL,
      'MSA-Template.docx',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    );
  }
}

const w9Summary = document.getElementById('w9Summary');
const w9CompanyName = document.getElementById('w9CompanyName');
const w9TaxId = document.getElementById('w9TaxId');
const w9TaxIdType = document.getElementById('w9TaxIdType');
const w9Address = document.getElementById('w9Address');
const w9EditHint = document.getElementById('w9EditHint');
const w9Debug = document.getElementById('w9Debug');
const w9RawText = document.getElementById('w9RawText');

function validateW9Fields() {
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

function refreshW9ValidationStatus() {
  const valid = validateW9Fields();
  step1NextBtn.disabled = !valid;
  if (valid && status.classList.contains('error')) {
    setStatus('Looks good — you can continue to the next step.');
  }
  return valid;
}

function detectTaxIdType(value) {
  const dashIndex = value.indexOf('-');
  if (dashIndex === 3) return 'SSN';
  if (dashIndex === 2) return 'EIN';
  return null;
}

function formatTaxId(digits, type) {
  if (digits.length !== 9) return digits;
  if (type === 'SSN') return `${digits.slice(0, 3)}-${digits.slice(3, 5)}-${digits.slice(5)}`;
  if (type === 'EIN') return `${digits.slice(0, 2)}-${digits.slice(2)}`;
  return digits;
}

w9CompanyName.addEventListener('input', () => {
  if (extractedW9Data) extractedW9Data.company_name = w9CompanyName.value.trim();
  refreshW9ValidationStatus();
  scheduleLivePreviewUpdate();
});
w9TaxId.addEventListener('input', () => {
  const value = w9TaxId.value.trim();
  if (extractedW9Data) extractedW9Data.tax_id_number = value;

  const detectedType = detectTaxIdType(value);
  if (detectedType) w9TaxIdType.value = detectedType;

  if (extractedW9Data) {
    extractedW9Data.selected_tax_id_type = w9TaxIdType.value;
    if (w9TaxIdType.value === 'SSN') extractedW9Data.ssn_number = value;
    else if (w9TaxIdType.value === 'EIN') extractedW9Data.ein_number = value;
  }

  refreshW9ValidationStatus();
  scheduleLivePreviewUpdate();
});
w9TaxIdType.addEventListener('change', () => {
  const newType = w9TaxIdType.value;
  let newValue = w9TaxId.value.trim();

  if (extractedW9Data) {
    newValue = (newType === 'SSN' ? extractedW9Data.ssn_number : extractedW9Data.ein_number) || '';
  } else {
    newValue = formatTaxId(newValue.replace(/\D/g, ''), newType);
  }

  w9TaxId.value = newValue;
  if (extractedW9Data) {
    extractedW9Data.selected_tax_id_type = newType;
    extractedW9Data.tax_id_number = newValue;
  }
  refreshW9ValidationStatus();
  scheduleLivePreviewUpdate();
});
w9Address.addEventListener('input', () => {
  if (!extractedW9Data) return;
  extractedW9Data.street_address = w9Address.value.trim();
  extractedW9Data.city = '';
  extractedW9Data.state_zip = '';
  refreshW9ValidationStatus();
  scheduleLivePreviewUpdate();
});

const IRS_ACROFORM_FIELD_RE = /\bf1_(\d{2})\[0\]$/;

function extractAcroFormDirectFields(annotations) {
  const byId = {};
  annotations.forEach((a) => {
    const m = a.fieldName && a.fieldName.match(IRS_ACROFORM_FIELD_RE);
    if (m) byId[m[1]] = String(a.fieldValue || '').trim();
  });

  if (byId['01'] === undefined && byId['02'] === undefined) return null;

  const cityStateZipMatch = (byId['08'] || '').match(/([A-Za-z][A-Za-z .'-]*),?\s+([A-Za-z]{2})\s*(\d{5}(?:-\d{4})?)/);

  return {
    company_name: byId['02'] || '',
    street_address: byId['07'] || '',
    city: cityStateZipMatch ? cityStateZipMatch[1].trim() : '',
    state_zip: cityStateZipMatch ? `${cityStateZipMatch[2]} ${cityStateZipMatch[3]}` : '',
    ssnDigits: `${byId['11'] || ''}${byId['12'] || ''}${byId['13'] || ''}`.replace(/\D/g, ''),
    einDigits: `${byId['14'] || ''}${byId['15'] || ''}`.replace(/\D/g, ''),
  };
}

async function extractPdfText(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let combined = '';
  let directFields = null;

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();

    const items = textContent.items.map((item) => ({ str: item.str, y: item.transform[5] }));

    const annotations = await page.getAnnotations();
    if (i === 1) directFields = extractAcroFormDirectFields(annotations);
    annotations.forEach((a) => {
      if (!a.fieldValue) return;
      const value = String(a.fieldValue);
      if (!a.rect) {
        items.push({ str: value, y: -Infinity });
        return;
      }
      const fieldY = a.rect[1];
      const insertAt = items.findIndex((it) => it.y < fieldY);
      if (insertAt === -1) items.push({ str: value, y: fieldY });
      else items.splice(insertAt, 0, { str: value, y: fieldY });
    });

    combined += items.map((it) => it.str).join(' ') + '\n';
  }

  return { text: combined, directFields };
}

const TIN_BOX_REGION = { x0: 0.55, x1: 1.0, y0: 0.4, y1: 0.62 };

const ADDRESS_BOX_REGION = { x0: 0.06, x1: 0.615, y0: 0.335, y1: 0.415 };

// Debug-only: converts a fraction-of-page region (TIN_BOX_REGION-style) to
// pixel bounds on the given canvas, so the debug overlay can draw it with
// the same rectangle shape as a dynamic (already pixel-space) crop region.
function fractionRegionToPixels(source, region) {
  return {
    x0: source.width * region.x0,
    x1: source.width * region.x1,
    y0: source.height * region.y0,
    y1: source.height * region.y1,
  };
}

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

function grayscaleContrastCanvas(source) {
  const out = document.createElement('canvas');
  out.width = source.width;
  out.height = source.height;
  const ctx = out.getContext('2d');
  ctx.drawImage(source, 0, 0);
  const imgData = ctx.getImageData(0, 0, out.width, out.height);
  const d = imgData.data;
  const gray = new Uint8ClampedArray(out.width * out.height);

  let min = 255;
  let max = 0;
  for (let i = 0, p = 0; i < d.length; i += 4, p++) {
    const g = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
    gray[p] = g;
    if (g < min) min = g;
    if (g > max) max = g;
  }

  const range = max - min;
  const scale = range > 10 ? 255 / range : 1;
  for (let i = 0, p = 0; i < d.length; i += 4, p++) {
    const v = range > 10 ? (gray[p] - min) * scale : gray[p];
    d[i] = d[i + 1] = d[i + 2] = v;
  }
  ctx.putImageData(imgData, 0, 0);
  return out;
}

const DESKEW_MAX_ANGLE_DEG = 8;
const DESKEW_STEP_DEG = 1;
const DESKEW_SAMPLE_MAX_DIM = 700;

function downscaleForAnalysis(source) {
  const scale = Math.min(1, DESKEW_SAMPLE_MAX_DIM / Math.max(source.width, source.height));
  if (scale >= 1) return source;
  const out = document.createElement('canvas');
  out.width = Math.round(source.width * scale);
  out.height = Math.round(source.height * scale);
  out.getContext('2d').drawImage(source, 0, 0, out.width, out.height);
  return out;
}

function rowProfileVariance(canvas, angleDeg) {
  const rad = (angleDeg * Math.PI) / 180;
  const w = canvas.width;
  const h = canvas.height;

  const rotated = document.createElement('canvas');
  const diag = Math.ceil(Math.sqrt(w * w + h * h));
  rotated.width = diag;
  rotated.height = diag;
  const ctx = rotated.getContext('2d');
  ctx.translate(diag / 2, diag / 2);
  ctx.rotate(rad);
  ctx.drawImage(canvas, -w / 2, -h / 2);

  const imgData = ctx.getImageData(0, 0, diag, diag);
  const d = imgData.data;
  const rowSums = new Float64Array(diag);
  for (let y = 0; y < diag; y++) {
    let sum = 0;
    const rowStart = y * diag * 4;
    for (let x = 0; x < diag; x++) {
      sum += d[rowStart + x * 4];
    }
    rowSums[y] = sum;
  }

  const mean = rowSums.reduce((a, b) => a + b, 0) / diag;
  const variance = rowSums.reduce((a, b) => a + (b - mean) * (b - mean), 0) / diag;
  return variance;
}

function deskewCanvas(source) {
  const sample = grayscaleContrastCanvas(downscaleForAnalysis(source));

  let bestAngle = 0;
  let bestVariance = rowProfileVariance(sample, 0);
  for (let a = -DESKEW_MAX_ANGLE_DEG; a <= DESKEW_MAX_ANGLE_DEG; a += DESKEW_STEP_DEG) {
    if (a === 0) continue;
    const variance = rowProfileVariance(sample, a);
    if (variance > bestVariance) {
      bestVariance = variance;
      bestAngle = a;
    }
  }

  if (bestAngle === 0) return source;

  const w = source.width;
  const h = source.height;
  const rad = (bestAngle * Math.PI) / 180;
  const out = document.createElement('canvas');
  out.width = w;
  out.height = h;
  const ctx = out.getContext('2d');
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, w, h);
  ctx.translate(w / 2, h / 2);
  ctx.rotate(rad);
  ctx.drawImage(source, -w / 2, -h / 2);
  return out;
}

const TESSERACT_VENDORED_PATHS = {
  workerPath: './vendor/tesseract/worker.min.js',
  corePath: './vendor/tesseract/tesseract-core-simd-lstm.wasm.js',
  langPath: './vendor/tesseract',
};

async function ocrLinesWithBoxes(canvas) {
  const worker = await Tesseract.createWorker('eng', undefined, TESSERACT_VENDORED_PATHS);
  const result = await worker.recognize(canvas, {}, { blocks: true });
  await worker.terminate();
  return linesFromBlocks(result.data.blocks);
}

function linesFromBlocks(blocks) {
  const lines = [];
  (blocks || []).forEach((block) =>
    (block.paragraphs || []).forEach((paragraph) => (paragraph.lines || []).forEach((line) => lines.push(line)))
  );
  lines.sort((a, b) => a.bbox.y0 - b.bbox.y0);
  return lines;
}

function findDigitRowBelow(lines, labelPattern) {
  for (const line of lines) {
    if (!labelPattern.test(line.text)) continue;
    const minWidth = (line.bbox.x1 - line.bbox.x0) * 0.4;
    const row = lines.find((l) => l.bbox.y0 >= line.bbox.y1 - 8 && l.bbox.x1 - l.bbox.x0 >= minWidth);
    if (row) return row.bbox;
  }
  return null;
}

// TIN_BOX_REGION's crop is a fixed fraction of the page, so on a standard W-9
// layout these two box rows land in a consistent y-band within it regardless
// of scan quality — used only when the label text above a row OCRs into
// unrecognizable noise and findDigitRowBelow can't find it that way.
const SSN_ROW_Y_BAND = { min: 400, max: 480 };
const EIN_ROW_Y_BAND = { min: 670, max: 750 };
const DIGIT_ROW_MIN_WIDTH = 400;

function findDigitRowByPosition(lines, yBand) {
  const candidates = lines.filter(
    (l) => l.bbox.y0 >= yBand.min && l.bbox.y0 <= yBand.max && l.bbox.x1 - l.bbox.x0 >= DIGIT_ROW_MIN_WIDTH
  );
  if (!candidates.length) return null;
  return candidates.reduce((best, l) =>
    l.bbox.x1 - l.bbox.x0 > best.bbox.x1 - best.bbox.x0 ? l : best
  ).bbox;
}

// --- Dynamic, full-page label anchoring -----------------------------------
// TIN_BOX_REGION/ADDRESS_BOX_REGION are a fixed fraction of the page, so a
// document with a shifted/cropped/rotated layout that pushes the real label
// outside that fraction never gets a chance. The functions below search the
// FULL page's OCR'd lines for the label instead (tolerating OCR typos), and
// use its position to compute a content-aware CROP REGION — replacing the
// fixed fraction, not the pipeline that follows it. Everything downstream
// (re-OCR the crop at 2x for precise local bboxes, findDigitRowBelow/
// findDigitRowByPosition, ocrTaxIdDigits, parseAddressBoxText) is the same
// already-tested logic that ran on the fixed crop before; only the source of
// the crop region changed. That matters because a first OCR pass over the
// whole page — lower zoom, more noise competing for Tesseract's attention —
// finds labels reliably but doesn't localize a value's box edges as
// precisely as a second, tightly-cropped-and-upscaled OCR pass of just that
// area, which is exactly what the existing fixed-region tier already did.
// The fixed fraction remains the fallback crop region when no label anchor
// is found anywhere on the page.

function normalizeFuzzy(s) {
  return (s || '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function levenshtein(a, b) {
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  let prev = new Array(n + 1);
  let curr = new Array(n + 1);
  for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[n];
}

// Tolerant match between one OCR'd line of text and a label phrase: exact
// (after normalization) first, then edit-distance within 25% of the phrase
// length so things like "Employer identlfication nurnber" still match
// "Employer identification number".
function fuzzyContainsLabel(lineText, phrase) {
  const normLine = normalizeFuzzy(lineText);
  const normPhrase = normalizeFuzzy(phrase);
  if (!normPhrase || !normLine) return false;
  // The length gate has to run before the exact-substring check too, not
  // just the fuzzy one: the W-9's own Part I instructions read "...this is
  // generally your social security number (SSN). However..." — a plain
  // substring search would "find" the real label phrase inside that
  // sentence and anchor on descriptive prose instead of the short label
  // line actually sitting next to the input box. A ratio (rather than a
  // fixed character-count difference) also catches a shorter but still
  // real form of the same problem: Tesseract's line detection sometimes
  // merges a bit of unrelated noise (a stray checkbox glyph, a scan
  // artifact) onto the same reported "line" as the true label — the label
  // must make up most of the line's own content, not just be present in it.
  if (normPhrase.length / normLine.length < 0.82) return false;
  if (normLine.includes(normPhrase)) return true;
  const maxDistance = Math.max(1, Math.floor(normPhrase.length * 0.25));
  return levenshtein(normLine, normPhrase) <= maxDistance;
}

// Finds the first line across the whole page matching any of the given
// label phrase variants (tried in order, most specific first).
// Tesseract's line-level bbox can be wider than the label actually printed
// on the page: a stray mark or an unrelated adjacent word sometimes gets
// grouped into the same reported "line" even when it barely affects the
// line's overall text length (a couple of extra characters can still pass
// the label-match ratio gate above). Re-deriving the bbox from just the
// run of words that actually match the phrase — rather than trusting the
// whole line's bbox — throws out that kind of stray content.
function narrowAnchorBbox(line, phrase) {
  const words = line.words;
  if (!words || !words.length) return line.bbox;
  const normPhrase = normalizeFuzzy(phrase);

  let best = null;
  for (let start = 0; start < words.length; start++) {
    let combined = '';
    for (let end = start; end < words.length; end++) {
      combined = combined ? `${combined} ${words[end].text}` : words[end].text;
      const normCombined = normalizeFuzzy(combined);
      if (normCombined.length > normPhrase.length * 1.4) break;
      if (normCombined.length < normPhrase.length * 0.6) continue;
      if (!fuzzyContainsLabel(combined, phrase)) continue;
      const span = words.slice(start, end + 1);
      const bbox = {
        x0: Math.min(...span.map((w) => w.bbox.x0)),
        x1: Math.max(...span.map((w) => w.bbox.x1)),
        y0: Math.min(...span.map((w) => w.bbox.y0)),
        y1: Math.max(...span.map((w) => w.bbox.y1)),
      };
      if (!best || bbox.x1 - bbox.x0 < best.x1 - best.x0) best = bbox;
    }
  }
  return best || line.bbox;
}

function findAnchorLine(lines, phraseVariants) {
  for (const line of lines) {
    for (const phrase of phraseVariants) {
      // Try the precise word-span match first: it's immune to a whole-line
      // length gate rejecting a real label just because an adjacent
      // column's text got merged onto the same OCR "line" — e.g. "Address
      // (number, street...)" and "Requester's name and address (optional)"
      // sit on the same row and can merge into one reported line even
      // though both are genuine text, not noise. Only fall back to the
      // whole-line check (which is what actually enforces the length-ratio
      // gate) when no word-level data narrowed anything down.
      const narrowedBbox = narrowAnchorBbox(line, phrase);
      if (narrowedBbox !== line.bbox) {
        return { line: { ...line, bbox: narrowedBbox }, matchedPhrase: phrase };
      }
      if (fuzzyContainsLabel(line.text, phrase)) {
        return { line, matchedPhrase: phrase };
      }
    }
  }
  return null;
}

const SSN_LABEL_VARIANTS = ['Social security number'];
const EIN_LABEL_VARIANTS = ['Employer identification number'];
const ADDRESS_LABEL_VARIANTS = ['Address number street and apt or suite no', 'Street address'];

// Crop bounds in the full-page canvas's own pixel space (as opposed to
// cropCanvasRegion's fraction-of-page-size bounds).
function cropRegionPixels(source, region, upscale) {
  const x0 = Math.max(0, Math.floor(region.x0));
  const y0 = Math.max(0, Math.floor(region.y0));
  const x1 = Math.min(source.width, Math.ceil(region.x1));
  const y1 = Math.min(source.height, Math.ceil(region.y1));
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

// Finds where the SSN/EIN labels actually are on the page and returns a
// generous crop region around them (label plus room for the box row(s)
// below), or null if neither anchor was found anywhere on the page.
const TIN_CROP_BELOW_HEIGHT = 140;

function findTinCropRegion(pageLines) {
  const ssnAnchor = findAnchorLine(pageLines, SSN_LABEL_VARIANTS);
  const einAnchor = findAnchorLine(pageLines, EIN_LABEL_VARIANTS);
  if (!ssnAnchor && !einAnchor) return null;

  const anchorBoxes = [ssnAnchor, einAnchor].filter(Boolean).map((a) => a.line.bbox);
  const widestAnchor = anchorBoxes.reduce((a, b) => (b.x1 - b.x0 > a.x1 - a.x0 ? b : a));

  return {
    region: {
      x0: Math.min(...anchorBoxes.map((b) => b.x0)) - 10,
      x1: Math.max(...anchorBoxes.map((b) => b.x1)) + (widestAnchor.x1 - widestAnchor.x0) * 0.6,
      y0: Math.min(...anchorBoxes.map((b) => b.y0)) - 10,
      y1: Math.max(...anchorBoxes.map((b) => b.y1)) + TIN_CROP_BELOW_HEIGHT,
    },
    ssnAnchored: !!ssnAnchor,
    einAnchored: !!einAnchor,
  };
}

// Finds where the address label actually is on the page and returns a crop
// region for its value: below the label, capped in width to roughly the
// label's own span so the crop doesn't reach across into the "Requester's
// name and address" box a real W-9 prints in the same row, further right.
const ADDRESS_CROP_BELOW_HEIGHT = 220;
const ADDRESS_CROP_MAX_X1_FRACTION = 0.65;

function findAddressCropRegion(canvas, pageLines) {
  const anchor = findAnchorLine(pageLines, ADDRESS_LABEL_VARIANTS);
  if (!anchor) return null;

  const b = anchor.line.bbox;
  const labelWidth = b.x1 - b.x0;
  // A noisier scan's label bbox is itself only approximately located, so a
  // fixed few-pixel left margin isn't always enough — it was clipping the
  // first character of the value line below on rougher scans (a street
  // address losing its leading digit, a city losing its first letter).
  // The label's own average character width (its bbox width divided by how
  // many characters it actually matched) is a more direct proxy for "how
  // wide is one character at this scan's resolution" than a fixed guess.
  const matchedLen = normalizeFuzzy(anchor.matchedPhrase).length || 1;
  const avgCharWidth = labelWidth / matchedLen;
  const leftPad = Math.max(25, avgCharWidth * 2);
  return {
    x0: Math.max(0, b.x0 - leftPad),
    x1: Math.min(canvas.width * ADDRESS_CROP_MAX_X1_FRACTION, b.x0 + Math.max(labelWidth, canvas.width * 0.4)),
    y0: Math.max(0, b.y0 - 10),
    y1: Math.min(canvas.height, b.y1 + ADDRESS_CROP_BELOW_HEIGHT),
  };
}

// A truly blank fillable box (common when a filer picked EIN and left the
// SSN box's printed cell dividers empty, or vice versa) can still OCR into a
// digit-shaped false positive under a digit whitelist — the box's own
// border/cell-divider lines get misread as characters. Requiring some real
// ink in the box's interior (away from its own border) before trusting a
// digit reading rejects that case without needing per-document tuning.
const MIN_DIGIT_INK_DENSITY = 0.015;

function regionInkDensity(canvas, bbox, marginRatio = 0.12) {
  const w = bbox.x1 - bbox.x0;
  const h = bbox.y1 - bbox.y0;
  if (w <= 0 || h <= 0) return 0;
  const mx = Math.round(w * marginRatio);
  const my = Math.round(h * marginRatio);
  const x0 = Math.max(0, Math.round(bbox.x0) + mx);
  const y0 = Math.max(0, Math.round(bbox.y0) + my);
  const x1 = Math.min(canvas.width, Math.round(bbox.x1) - mx);
  const y1 = Math.min(canvas.height, Math.round(bbox.y1) - my);
  const w2 = x1 - x0;
  const h2 = y1 - y0;
  if (w2 <= 0 || h2 <= 0) return 0;

  const { data } = canvas.getContext('2d').getImageData(x0, y0, w2, h2);
  let dark = 0;
  for (let i = 0; i < data.length; i += 4) {
    const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
    if (lum < 140) dark++;
  }
  return dark / (w2 * h2);
}

// Takes an already-created, already-parameterized worker rather than
// spinning up its own: a worker's WASM init/teardown is real overhead
// (hundreds of ms), and the multi-crop, multi-scale digit reading above it
// calls this many times per field — up to a couple dozen times per document
// between SSN and EIN. Tesseract.js workers run jobs one at a time
// internally anyway, so sharing one across all those calls costs nothing in
// correctness and saves most of that overhead.
async function ocrDigitRowAtBbox(worker, canvas, bbox, upscale) {
  if (!bbox) return '';
  const pad = 4;
  const x0 = Math.max(0, bbox.x0 - pad);
  const y0 = Math.max(0, bbox.y0 - pad);
  const w = bbox.x1 - bbox.x0 + pad * 2;
  const h = bbox.y1 - bbox.y0 + pad * 2;

  const out = document.createElement('canvas');
  out.width = w * upscale;
  out.height = h * upscale;
  const outCtx = out.getContext('2d');
  outCtx.imageSmoothingEnabled = false;
  outCtx.drawImage(canvas, x0, y0, w, h, 0, 0, out.width, out.height);

  const result = await worker.recognize(out);

  let digits = (result.data.text || '').replace(/\D/g, '');
  if (digits.length === 10) digits = digits.slice(1);
  return digits;
}

// Tesseract's coarse line detection sometimes clips a row's left or right edge,
// dropping a digit. Since we don't know which side (if either) was clipped,
// try both a widened-left and an as-detected-left crop — margined by the
// row's own width, so this works the same whether bbox came from a small
// fixed-region crop or a full-page search — and only trust a reading of the
// expected length when there's no disagreement about what it is; two
// different same-length stories is a sign something's wrong and it's safer
// to report "not found" than guess.
async function ocrTaxIdDigitsAtScale(worker, canvas, bbox, expectedLength, upscale) {
  const rowWidth = bbox.x1 - bbox.x0;
  const wideLeft = Math.max(0, bbox.x0 - rowWidth);
  const wideRight = bbox.x1 + rowWidth * 0.5;
  const [fromWideLeft, fromDetected] = await Promise.all([
    ocrDigitRowAtBbox(worker, canvas, { x0: wideLeft, x1: wideRight, y0: bbox.y0, y1: bbox.y1 }, upscale),
    ocrDigitRowAtBbox(worker, canvas, { x0: bbox.x0, x1: wideRight, y0: bbox.y0, y1: bbox.y1 }, upscale),
  ]);
  const wideLeftOk = fromWideLeft.length === expectedLength;
  const detectedOk = fromDetected.length === expectedLength;
  if (wideLeftOk && detectedOk) return fromWideLeft === fromDetected ? fromWideLeft : '';
  if (wideLeftOk) return fromWideLeft;
  if (detectedOk) return fromDetected;
  return fromWideLeft.length > fromDetected.length ? fromWideLeft : fromDetected;
}

// Different scan qualities favor different amounts of upscaling before
// whitelisted-digit OCR: a faint/low-resolution photo needs more magnification
// to resolve strokes at all, but the same amount of magnification on an
// already-sharp, high-contrast scan (e.g. CamScanner's own enhancement) can
// over-sharpen two adjacent digits into merging or dropping one. Rather than
// pick one fixed scale that helps one case and hurts the other, try both and
// only trust the result when they agree — the same "prefer no answer over a
// guess" principle ocrTaxIdDigitsAtScale already applies to its own crop
// variants, one level up.
// Three scales, not two: with only two, a disagreement is a coin flip with
// no way to break the tie except distrusting both. A third vote lets a real
// majority (2 of 3) win even when one scale's magnification happens to suit
// this particular scan poorly — too little to resolve a faint stroke, or
// (just as real a failure mode) too much, over-sharpening two digits into
// one. Only a genuine three-way split — every scale reading something
// different — falls back to "ambiguous", since there's no way to pick a
// winner from that honestly.
const DIGIT_OCR_SCALES = [2, 3, 4];

async function ocrTaxIdDigits(worker, canvas, bbox, expectedLength) {
  if (!bbox) return '';
  if (regionInkDensity(canvas, bbox) < MIN_DIGIT_INK_DENSITY) return '';

  const results = await Promise.all(
    DIGIT_OCR_SCALES.map((scale) => ocrTaxIdDigitsAtScale(worker, canvas, bbox, expectedLength, scale))
  );
  const valid = results.filter((d) => d.length === expectedLength);
  if (!valid.length) {
    // No scale reached the expected digit count. If most readings came
    // back genuinely empty, that's a blank box — safe to report "not
    // found" and let a fallback try elsewhere. But if multiple scales
    // agree on the SAME non-empty, wrong-length reading, that's evidence
    // the box has real content that's consistently hard to fully resolve
    // (a digit dropped at every magnification) — not nothing. Blocking
    // the fallback here does cost real coverage sometimes: a case was
    // measured where the full-text fallback this blocks would actually
    // have recovered the correct value. But the same shape of signal was
    // also measured letting a wrong value through uncorroborated in a
    // different document. With no way to tell those two apart from this
    // signal alone, and an incorrect tax ID being the worse outcome by
    // far, this stays biased toward flagging ambiguous over guessing.
    const nonEmpty = results.filter((d) => d.length > 0);
    return nonEmpty.length >= 2 ? null : '';
  }

  const counts = new Map();
  valid.forEach((d) => counts.set(d, (counts.get(d) || 0) + 1));
  const [bestValue, bestCount] = [...counts.entries()].reduce((best, cur) => (cur[1] > best[1] ? cur : best));

  // A lone scale's reading, unconfirmed by any other, isn't corroborated —
  // same reasoning as requiring both wideLeft/detected crops to agree one
  // level down. Return null (not '') so callers can tell "found candidates
  // that disagree" apart from "found nothing at all": the former means a
  // less careful fallback (plain regex over noisy OCR text) shouldn't get a
  // turn to guess either, since the more rigorous method already tried and
  // came back genuinely unsure.
  //
  // Requiring full 3-way unanimity instead of a 2-of-3 majority was tried
  // and measured worse: two scales can share one correlated misread while
  // the third gets it right, so majority voting isn't foolproof — but
  // empirically, requiring unanimity turned several genuinely-correct
  // majority reads into "uncertain" for every one confidently-wrong
  // majority it caught. Without another independent signal to break a
  // 2-1 split correctly, majority is the better bet on the whole corpus.
  return bestCount >= 2 ? bestValue : null;
}

async function tinDigitsFromCrop(cropCanvas) {
  if (!cropCanvas) return { ssn: '', ein: '' };
  const lines = await ocrLinesWithBoxes(cropCanvas);
  const ssnBbox =
    findDigitRowBelow(lines, loose('Social security number')) || findDigitRowByPosition(lines, SSN_ROW_Y_BAND);
  const einBbox =
    findDigitRowBelow(lines, loose('Employer identification number')) ||
    findDigitRowByPosition(lines, EIN_ROW_Y_BAND);

  // One worker for every digit-OCR attempt this crop needs (up to a couple
  // dozen, across SSN/EIN × 3 scales × 2 crop variants) instead of one per
  // attempt — see the note on ocrDigitRowAtBbox.
  const digitWorker = await Tesseract.createWorker('eng', undefined, TESSERACT_VENDORED_PATHS);
  await digitWorker.setParameters({
    tessedit_char_whitelist: '0123456789-',
    tessedit_pageseg_mode: '6',
  });
  const [ssn, ein] = await Promise.all([
    ocrTaxIdDigits(digitWorker, cropCanvas, ssnBbox, 9),
    ocrTaxIdDigits(digitWorker, cropCanvas, einBbox, 9),
  ]);
  await digitWorker.terminate();

  return { ssn, ein };
}

// Primary: crop around the SSN/EIN labels wherever findTinCropRegion locates
// them on the full page — same precise local-crop pipeline (re-OCR at 2x for
// bboxes, then targeted digit OCR) the fixed region always used. Fallback:
// re-run that same pipeline on the fixed TIN_BOX_REGION fraction instead,
// for whichever of SSN/EIN is still missing — covers both "no anchor found
// anywhere" and "anchor found, but its crop didn't turn up a usable row"
// (a tighter, content-aware crop is usually more precise, but the fixed
// region's generous bounds are more forgiving when something about this
// document's layout or scan quality throws the anchor off).
async function ocrTinBoxDigits(canvas, pageLines) {
  const dynamic = pageLines && pageLines.length ? findTinCropRegion(pageLines) : null;
  const primaryCanvas = dynamic ? cropRegionPixels(canvas, dynamic.region, 2) : null;
  const primary = await tinDigitsFromCrop(primaryCanvas);

  let ssn = primary.ssn;
  let ein = primary.ein;
  let source = dynamic ? 'dynamic-anchor' : 'fixed-region';
  let fixedCanvas = null;

  // ssn/ein can be a digit string, '' (nothing usable found), or null
  // (found candidates that disagree — genuinely ambiguous, distinct from
  // "not found"). Only ever overwrite with a fallback reading that's a
  // clean, confirmed 9 digits. When neither tier lands a clean 9-digit
  // answer, the merged result must stay null if EITHER tier flagged real
  // disagreement — collapsing that back to '' would erase the one signal
  // that stops extractTaxId's uncross-checked text fallback from guessing
  // where a more careful method already found reason for doubt.
  if ((ssn || '').length !== 9 || (ein || '').length !== 9) {
    fixedCanvas = cropCanvasRegion(canvas, TIN_BOX_REGION, 2);
    const fallback = await tinDigitsFromCrop(fixedCanvas);
    if ((ssn || '').length !== 9) {
      if ((fallback.ssn || '').length === 9) {
        ssn = fallback.ssn;
        source = dynamic ? 'dynamic-anchor+fixed-region' : 'fixed-region';
      } else if (ssn === null || fallback.ssn === null) {
        ssn = null;
      }
    }
    if ((ein || '').length !== 9) {
      if ((fallback.ein || '').length === 9) {
        ein = fallback.ein;
        source = dynamic ? 'dynamic-anchor+fixed-region' : 'fixed-region';
      } else if (ein === null || fallback.ein === null) {
        ein = null;
      }
    }
  }

  if (debugExtractionTrace) {
    debugExtractionTrace.tin.source = source;
    debugExtractionTrace.tin.cropRegion = dynamic ? dynamic.region : fractionRegionToPixels(canvas, TIN_BOX_REGION);
    debugExtractionTrace.tin.cropCanvas = primaryCanvas || fixedCanvas;
  }

  return { ssn, ein, source };
}

async function addressTextFromCrop(cropCanvas) {
  if (!cropCanvas) return '';
  const preprocessed = grayscaleContrastCanvas(cropCanvas);
  const worker = await Tesseract.createWorker('eng', undefined, TESSERACT_VENDORED_PATHS);
  await worker.setParameters({ tessedit_pageseg_mode: '4' });
  const result = await worker.recognize(preprocessed);
  await worker.terminate();
  const text = result.data.text || '';
  return text.replace(/\s+/g, '').length >= 5 ? text : '';
}

// Same two-tier shape as ocrTinBoxDigits: try the content-aware crop around
// wherever the address label was found on the full page first, and fall
// back to the fixed ADDRESS_BOX_REGION fraction if that didn't turn up
// enough text (covers both "no anchor found" and "anchor found, but its
// crop wasn't usable").
async function ocrAddressBoxText(canvas, pageLines) {
  const dynamic = pageLines && pageLines.length ? findAddressCropRegion(canvas, pageLines) : null;
  const primaryCanvas = dynamic ? cropRegionPixels(canvas, dynamic, 2) : null;
  let text = await addressTextFromCrop(primaryCanvas);
  let source = dynamic ? 'dynamic-anchor' : 'fixed-region';
  let fixedCanvas = null;

  if (!text) {
    fixedCanvas = cropCanvasRegion(canvas, ADDRESS_BOX_REGION, 2);
    text = await addressTextFromCrop(fixedCanvas);
    if (text) source = dynamic ? 'dynamic-anchor+fixed-region' : 'fixed-region';
  }

  if (debugExtractionTrace) {
    debugExtractionTrace.address.source = text ? source : 'none';
    debugExtractionTrace.address.cropRegion = dynamic || fractionRegionToPixels(canvas, ADDRESS_BOX_REGION);
    debugExtractionTrace.address.cropCanvas = primaryCanvas || fixedCanvas;
  }

  return text;
}

function isImageFile(file) {
  return /^image\//.test(file.type) || /\.(jpe?g|png|webp|heic|heif|bmp)$/i.test(file.name);
}

async function imageFileToCanvas(file) {
  const bitmap = await createImageBitmap(file);
  const canvas = document.createElement('canvas');
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  canvas.getContext('2d').drawImage(bitmap, 0, 0);
  return canvas;
}

// Full-page layout OCR: a single recognize pass that returns both the plain
// text (existing regex-parsing path) and structured line bboxes (blocks),
// so the dynamic label-anchoring tier can search the whole page for free
// instead of paying for a second full-page OCR pass.
async function ocrPageWithLines(canvas, onProgress) {
  const worker = await Tesseract.createWorker('eng', undefined, {
    logger: onProgress,
    ...TESSERACT_VENDORED_PATHS,
  });
  const result = await worker.recognize(canvas, {}, { text: true, blocks: true });
  await worker.terminate();
  return { text: result.data.text, lines: linesFromBlocks(result.data.blocks) };
}

async function ocrImageFile(file, onProgress) {
  resetDebugTrace();
  const rawCanvas = await imageFileToCanvas(file);
  const canvas = deskewCanvas(rawCanvas);
  const ocrCanvas = grayscaleContrastCanvas(canvas);
  if (debugExtractionTrace) debugExtractionTrace.canvas = canvas;

  const { text, lines } = await ocrPageWithLines(ocrCanvas, onProgress ? (m) => onProgress(1, 1, m) : undefined);

  const [tinDigits, addressBoxText] = await Promise.all([
    ocrTinBoxDigits(canvas, lines),
    ocrAddressBoxText(canvas, lines),
  ]);
  return { text, tinDigits, addressBoxText };
}

async function ocrPdfText(file, onProgress) {
  resetDebugTrace();
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const pagesToScan = Math.min(pdf.numPages, 2);
  let combined = '';
  let tinDigits = { ssn: '', ein: '' };
  let addressBoxText = '';

  for (let i = 1; i <= pagesToScan; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 3 });
    const rawCanvas = document.createElement('canvas');
    rawCanvas.width = viewport.width;
    rawCanvas.height = viewport.height;
    await page.render({ canvasContext: rawCanvas.getContext('2d'), viewport }).promise;

    const canvas = deskewCanvas(rawCanvas);
    const ocrCanvas = grayscaleContrastCanvas(canvas);

    const { text, lines } = await ocrPageWithLines(
      ocrCanvas,
      onProgress ? (m) => onProgress(i, pagesToScan, m) : undefined
    );
    combined += text + '\n';

    if (i === 1) {
      if (debugExtractionTrace) debugExtractionTrace.canvas = canvas;
      [tinDigits, addressBoxText] = await Promise.all([
        ocrTinBoxDigits(canvas, lines),
        ocrAddressBoxText(canvas, lines),
      ]);
    }
  }

  return { text: combined, tinDigits, addressBoxText };
}

function loosePattern(phrase) {
  return phrase
    .split(/\s+/)
    .map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
    .join('\\s*[,.:]?\\s*');
}

function loose(phrase, flags = 'i') {
  return new RegExp(loosePattern(phrase), flags);
}

const LINE_NUM_PREFIX = '(?:\\b\\d{1,2}[a-z]?\\s+)?';
const NEXT_LABEL_PHRASES = [
  'Business name/disregarded entity name',
  'Federal tax classification',
  'Check the appropriate',
  'Exemptions',
  'Address (number',
  'Social security number',
  'Employer identification number',
  'List account number',
];
const PART_I_PATTERN = 'Part\\s*[Il]\\b';
const PART_II_PATTERN = 'Part\\s*[Il]{2}\\b';
const NEXT_LABEL_STOP = new RegExp(
  LINE_NUM_PREFIX +
    '(?:' +
    NEXT_LABEL_PHRASES.map(loosePattern).join('|') +
    '|' +
    PART_I_PATTERN +
    '|' +
    PART_II_PATTERN +
    ')',
  'i'
);

const NEXT_LABEL_STOP_NO_LINE_PREFIX = new RegExp(
  '(?:' + NEXT_LABEL_PHRASES.map(loosePattern).join('|') + '|' + PART_I_PATTERN + '|' + PART_II_PATTERN + ')',
  'i'
);

function cleanCapture(str) {
  let s = str.trim();
  for (let i = 0; i < 5; i++) {
    const before = s;
    s = s.replace(/^\([^)]*\)\s*/, '');
    s = s.replace(/^[.:\-–]\s*/, '');
    s = s.replace(/^See instructions\.?\s*/i, '');
    s = s.replace(/^An entry is required\.?\s*/i, '');
    s = s.trim();
    if (s === before) break;
  }
  return s;
}

const CAPTURE_FALLBACK_WINDOW = 220;

function captureAfterLabel(text, labelRegex, stopRegex) {
  const labelMatch = text.match(labelRegex);
  if (!labelMatch) return '';
  const rest = text.slice(labelMatch.index + labelMatch[0].length);

  const stopCandidates = [stopRegex, NEXT_LABEL_STOP]
    .map((re) => rest.match(re))
    .filter(Boolean)
    .sort((a, b) => a.index - b.index);

  const cutoff = stopCandidates.length ? stopCandidates[0].index : CAPTURE_FALLBACK_WINDOW;
  return cleanCapture(rest.slice(0, cutoff));
}

function lastMatch(text, regex) {
  const global = new RegExp(regex.source, regex.flags.includes('g') ? regex.flags : regex.flags + 'g');
  let found = null;
  let m;
  while ((m = global.exec(text))) {
    found = m;
    if (m[0].length === 0) global.lastIndex += 1;
  }
  return found;
}

const DIGIT_LOOKALIKE = { o: '0', O: '0', l: '1', I: '1', z: '2', Z: '2', s: '5', S: '5', e: '6', g: '9', q: '9', b: '6', B: '8' };
function normalizeBracketedZero(str) {
  return str.replace(/(?<=[[\]\d])([a-zA-Z])(?=[[\]\d])/g, (m, ch) => DIGIT_LOOKALIKE[ch] || ch);
}

function bestDigitCluster(str) {
  const clusters = normalizeBracketedZero(str).match(/[\d[\]\- ]+/g) || [];
  const digitStrings = clusters.map((c) => c.replace(/\D/g, '')).filter(Boolean);
  const exact9 = digitStrings.filter((d) => d.length === 9);
  if (exact9.length) return exact9[exact9.length - 1];
  return digitStrings.reduce((best, d) => (d.length > best.length ? d : best), '');
}

function digitsAfterLabel(text, labelRegex, windowChars) {
  const m = lastMatch(text, labelRegex);
  if (!m) return '';
  const start = m.index + m[0].length;
  const rest = text.slice(start, start + windowChars);
  const stopMatch = rest.match(NEXT_LABEL_STOP_NO_LINE_PREFIX);
  const bounded = stopMatch ? rest.slice(0, stopMatch.index) : rest;
  return bestDigitCluster(bounded);
}

function extractPartISection(text) {
  const headingMatch = text.match(loose('Taxpayer Identification Number (TIN)'));
  const start = headingMatch ? headingMatch.index + headingMatch[0].length : 0;
  const rest = text.slice(start);
  const partIIMatch = rest.match(new RegExp(PART_II_PATTERN, 'i'));
  return partIIMatch ? rest.slice(0, partIIMatch.index) : rest;
}

const ADDRESS_LABEL_STOP_RE = new RegExp(LINE_NUM_PREFIX + loosePattern('Address (number'), 'i');

const STATE_ZIP_CODE_PATTERN = loosePattern('state and ZIP') + '\\s*c[o0a]de';

const BUSINESS_LABEL_TAIL_RE =
  /^.{0,3}?disregarded entity(?:'?s)?\s*name\s*[,.:;-]*\s*(?:if\s*different(?:\s*from\s*above)?\s*[,.:;-]*\s*)?/i;
const IF_DIFFERENT_TAIL_RE = /^.{0,3}?if\s*different(?:\s*from\s*above)?\s*[,.:;-]*\s*/i;

// Fuzzy fallback for the two regexes above: they require exact spelling
// (tolerant only of whitespace/punctuation), so an OCR typo like "enlity"
// for "entity" or "diferent" for "different" — common on rougher scans —
// slips straight through both and leaves the label's own tail text stuck
// in front of the real company name. This strips a fuzzy (edit-distance
// tolerant) match of the same phrases instead of an exact one.
// Longest phrases first, down to short trailing fragments: a first pass can
// leave a partial residual behind (e.g. matching "if different f" against a
// duplicated OCR artifact and stopping there), and the retry loop in
// extractBusinessName needs a shorter phrase to recognize what's left as
// still-leftover label text rather than the start of the real company name.
const BUSINESS_LABEL_TAIL_PHRASES = [
  "disregarded entity's name if different from above",
  'disregarded entity name if different from above',
  'if different from above',
  'from above',
];

function stripFuzzyLeadingPhrase(value, phrases) {
  const trimmed = value.trim();
  for (const phrase of phrases) {
    const normPhrase = normalizeFuzzy(phrase);
    if (!normPhrase) continue;
    const maxCut = Math.min(trimmed.length, Math.ceil(normPhrase.length * 1.4) + 10);
    const minCut = Math.max(1, Math.floor(normPhrase.length * 0.6));
    // Taking the first cut that merely satisfies a fixed distance
    // threshold is unreliable: a cut a few characters past the true
    // boundary can still clear that same threshold (a couple of real typos
    // plus a short run of genuine trailing content can add up to no more
    // edits than the typos alone), so scanning in any fixed direction and
    // stopping early risks landing on a worse match than one just next to
    // it. Evaluating every candidate length and keeping the one with the
    // lowest edit distance finds the actual best-fit boundary instead.
    let bestCut = -1;
    let bestDistance = Infinity;
    for (let cut = minCut; cut <= maxCut; cut++) {
      const normCandidate = normalizeFuzzy(trimmed.slice(0, cut));
      if (Math.abs(normCandidate.length - normPhrase.length) > Math.ceil(normPhrase.length * 0.3)) continue;
      const distance = levenshtein(normCandidate, normPhrase);
      if (distance < bestDistance) {
        bestDistance = distance;
        bestCut = cut;
      }
    }
    if (bestCut > -1 && bestDistance <= Math.ceil(normPhrase.length * 0.25)) {
      return trimmed.slice(bestCut).trim();
    }
  }
  return value;
}

function extractBusinessName(text) {
  const full = captureAfterLabel(
    text,
    loose('Business name/disregarded entity name if different from above'),
    ADDRESS_LABEL_STOP_RE
  );
  if (full) return full;

  let value = captureAfterLabel(text, loose('Business name'), ADDRESS_LABEL_STOP_RE);
  for (let i = 0; i < 5; i++) {
    const before = value;
    value = value.replace(BUSINESS_LABEL_TAIL_RE, '');
    value = value.replace(IF_DIFFERENT_TAIL_RE, '');
    value = stripFuzzyLeadingPhrase(value, BUSINESS_LABEL_TAIL_PHRASES);
    value = cleanCapture(value);
    if (value === before) break;
  }
  return value;
}

const TIN_LINE_REF_RE = /\bfor\s+line\s+\d{1,2}\b/gi;

function extractTaxId(text, tinDigits) {
  const tinSection = extractPartISection(text).replace(TIN_LINE_REF_RE, ' ');

  // tinDigits.ssn/.ein can be a digit string, '' (box OCR found nothing
  // usable), or null (box OCR found candidates that genuinely disagreed —
  // see ocrTaxIdDigits). null is a stronger signal than '': the targeted,
  // cross-checked box reading already tried and came back unsure, so the
  // cruder full-text regex below — which has no cross-checking of its own —
  // shouldn't get a turn to guess either.
  const ssnAmbiguous = !!tinDigits && tinDigits.ssn === null;
  const einAmbiguous = !!tinDigits && tinDigits.ein === null;
  const ssnBoxDigits = tinDigits && tinDigits.ssn && tinDigits.ssn.length >= 9 ? tinDigits.ssn : '';
  const einBoxDigits = tinDigits && tinDigits.ein && tinDigits.ein.length >= 9 ? tinDigits.ein : '';
  const ssnDigits = ssnBoxDigits || (ssnAmbiguous ? '' : digitsAfterLabel(tinSection, loose('Social security number'), 350));
  const einDigits =
    einBoxDigits || (einAmbiguous ? '' : digitsAfterLabel(tinSection, loose('Employer identification number'), 350));

  let ssnNumber = '';
  if (ssnDigits.length >= 9) {
    const d = ssnDigits.slice(0, 9);
    ssnNumber = `${d.slice(0, 3)}-${d.slice(3, 5)}-${d.slice(5)}`;
  } else if (!ssnAmbiguous) {
    const ssnMatch = tinSection.match(/\b(\d{3}-\d{2}-\d{4})\b/);
    if (ssnMatch) ssnNumber = ssnMatch[1];
  }

  let einNumber = '';
  if (einDigits.length >= 9) {
    const d = einDigits.slice(0, 9);
    einNumber = `${d.slice(0, 2)}-${d.slice(2)}`;
  } else if (!einAmbiguous) {
    const einMatch = tinSection.match(/\b(\d{2}-\d{7})\b/);
    if (einMatch) einNumber = einMatch[1];
  }

  let selectedType = null;
  let taxId = '';
  if (ssnNumber) {
    selectedType = 'SSN';
    taxId = ssnNumber;
  } else if (einNumber) {
    selectedType = 'EIN';
    taxId = einNumber;
  }

  return { ssnNumber, einNumber, selectedType, taxId };
}

function parseAddressBoxText(rawText) {
  const text = rawText.replace(/\s+/g, ' ').trim();

  const ADDRESS_LABEL_RE = new RegExp(
    loosePattern('Address (number street and apt or suite no') + '\\.?\\)?\\.?(?:\\s*See\\s*instructions\\.?)?',
    'i'
  );
  // Not the full LINE_NUM_PREFIX here (unlike the otherwise-identical
  // regex in extractAddress below): the real form line-number ("6") that
  // legitimately precedes "City, state, and ZIP code" is always a single
  // bare digit — LINE_NUM_PREFIX's fuller "1-2 digits + optional letter,
  // case-insensitive" shape also matches a genuine unit number like "4B"
  // sitting in that same position, swallowing it into the label match and
  // truncating it out of the street text. A single digit immediately
  // followed by whitespace can't match "4B" (the "B" blocks it) but still
  // catches the real line-number and its common OCR misreads (6 -> 5/3/8).
  const CITY_ZIP_LABEL_RE = new RegExp('(?:\\d\\s+)?(?:city\\s*[,.:]?\\s*)?' + STATE_ZIP_CODE_PATTERN + '\\.?', 'i');

  // Each word required to start uppercase, rather than any letter: real
  // city names are always Title Case, and requiring that rejects a run of
  // OCR noise sitting between the label and the real city (a stray
  // lowercase-led fragment like "I se CR" right before "Atlanta") that the
  // fully permissive [A-Za-z ...] version would otherwise happily swallow
  // into the captured city name as if it were the start of it.
  const CITY_STATE_ZIP_RE = /((?:[A-Z][A-Za-z'.-]*\s*)+?),?\s+([A-Za-z]{2})\s*(\d{5}(?:-\d{4})?)/;

  const addressLabelMatch = text.match(ADDRESS_LABEL_RE);
  const afterAddressLabel = addressLabelMatch ? text.slice(addressLabelMatch.index + addressLabelMatch[0].length) : text;

  const splitMatch = afterAddressLabel.match(CITY_ZIP_LABEL_RE);
  const cityPart = splitMatch ? afterAddressLabel.slice(splitMatch.index + splitMatch[0].length) : afterAddressLabel;
  const cityMatch = cityPart.match(CITY_STATE_ZIP_RE);
  if (!cityMatch) return null;

  const streetRaw = splitMatch ? afterAddressLabel.slice(0, splitMatch.index) : cityPart.slice(0, cityMatch.index);

  // The first digit-led token in the street text is the actual house
  // number — picking the last one instead backfires when a later token
  // (a suite number, a stray OCR character right after it) also happens to
  // look like "digits + letter" and gets mistaken for a second, "more
  // recent" street start, discarding the real street name that came
  // before it.
  const streetStartMatches = [...streetRaw.matchAll(/\d{2,}[\d-]*\s+[A-Za-z]/g)];
  const boundedStreetRaw = streetStartMatches.length ? streetRaw.slice(streetStartMatches[0].index) : streetRaw;

  const streetPart = cleanCapture(boundedStreetRaw.replace(/\s+/g, ' ').trim())
    .replace(/^[|~._,&+-]+\s*/, '')
    .replace(/\s*[|~._,&+-]+$/, '')
    // A stray single OCR-noise character right at the street/city crop
    // boundary is common enough to show up across several unrelated
    // documents — not punctuation (already handled above), just a bare
    // lowercase letter or two. Real unit letters in these forms are always
    // uppercase ("Ste B", "Apt A"), so this can't be mistaken for one.
    .replace(/\s+[a-z]{1,2}$/, '')
    .trim();

  return {
    streetAddress: streetPart,
    city: cityMatch[1].trim(),
    stateZip: `${cityMatch[2]} ${cityMatch[3]}`,
  };
}

function extractAddress(text, tinSection) {
  const ADDRESS_LABEL_RE = new RegExp(
    loosePattern('Address (number street and apt or suite no') + '\\.?\\)?\\.?(?:\\s*See\\s*instructions\\.?)?',
    'i'
  );
  const CITY_ZIP_LABEL_RE = new RegExp(
    LINE_NUM_PREFIX + '(?:city\\s*[,.:]?\\s*)?' + STATE_ZIP_CODE_PATTERN + '\\.?',
    'i'
  );
  const REQUESTER_LABEL_RE = /Requester.{0,3}s?\s*name\s*and\s*address\s*[(\[]?\s*optional\s*[)\]]?\.?/i;
  const CITY_STATE_ZIP_RE = /([A-Za-z][A-Za-z .'-]{1,40}?),?\s+([A-Za-z]{2})\s*(\d{5}(?:-\d{4})?)/;
  const ADDRESS_STOP_RE = new RegExp(
    LINE_NUM_PREFIX +
      '(?:' +
      loosePattern('Social security number') +
      '|' +
      loosePattern('Employer identification number') +
      '|' +
      loosePattern('List account number') +
      '|' +
      loosePattern('Part I') +
      ')',
    'i'
  );
  const ADDRESS_FALLBACK_WINDOW = 400;

  const STREET_ADDRESS_MAX_LEN = 60;
  function capStreetAddress(str) {
    if (str.length <= STREET_ADDRESS_MAX_LEN) return str;
    const truncated = str.slice(0, STREET_ADDRESS_MAX_LEN);
    const lastComma = truncated.lastIndexOf(',');
    if (lastComma > 10) return truncated.slice(0, lastComma).trim();
    const lastSpace = truncated.lastIndexOf(' ');
    return (lastSpace > 10 ? truncated.slice(0, lastSpace) : truncated).trim();
  }

  const STREET_CORE_RE = /\d{2,}[\d-]*\s+[^,]{1,40}/;
  const STREET_UNIT_RE =
    /^,?\s*(?:(?:Suite|Ste\.?|Apt\.?|Unit|Studio|Loft|Rm\.?|Room|Bldg\.?|Building|Floor|Fl\.?|Penthouse|PH|Trlr|Lot|Space|Bay|#)\s*[A-Za-z0-9-]{1,6}|\d{1,3}(?:st|nd|rd|th)\s*(?:Floor|Fl\.?))\b/i;
  function extractStreetValue(str) {
    const coreMatch = str.match(STREET_CORE_RE);
    if (!coreMatch) return  capStreetAddress(str);
    const core = coreMatch[0].trim();
    const rest = str.slice(coreMatch.index + coreMatch[0].length);
    const unitMatch = rest.match(STREET_UNIT_RE);
    return unitMatch ? `${core}, ${unitMatch[0].replace(/^,?\s*/, '').trim()}` : core;
  }

  const STREET_START_RE = /\d{2,}[\d-]*\s+[A-Za-z]/;
  function trimToStreetStart(str) {
    const m = str.match(STREET_START_RE);
    return m && m.index > 0 ? str.slice(m.index) : str;
  }

  const COMPANY_SUFFIX_RE =
    /\s+[A-Z][A-Za-z&.,'-]*(?:\s+[A-Z][A-Za-z&.,'-]*)*\s+(?:Inc\.?|LLC|L\.L\.C\.|Corp\.?|Corporation|Co\.|Company|Ltd\.?|Group|Partners|LP|LLP)(?=\s|$|,).*$/;
  const trimTrailingCompanyName = (str) => {
    const m = str.match(COMPANY_SUFFIX_RE);
    return m && m.index > 8 ? str.slice(0, m.index).trim() : str;
  };

  let streetAddress = '';
  let city = '';
  let stateZip = '';

  const addressLabelMatch = text.match(ADDRESS_LABEL_RE);
  if (addressLabelMatch) {
    const afterLabel = text.slice(addressLabelMatch.index + addressLabelMatch[0].length);
    const stopMatch = afterLabel.match(ADDRESS_STOP_RE);
    const windowText = stopMatch ? afterLabel.slice(0, stopMatch.index) : afterLabel.slice(0, ADDRESS_FALLBACK_WINDOW);

    const cityZipLabelMatch = windowText.match(CITY_ZIP_LABEL_RE);
    const requesterLabelMatch = windowText.match(REQUESTER_LABEL_RE);
    const streetStartMatch = windowText.match(STREET_START_RE);

    let trimmedWindow = windowText;
    if (requesterLabelMatch && streetStartMatch && requesterLabelMatch.index < streetStartMatch.index) {
      trimmedWindow =
        windowText.slice(0, requesterLabelMatch.index) +
        windowText.slice(requesterLabelMatch.index + requesterLabelMatch[0].length);
    } else if (requesterLabelMatch && cityZipLabelMatch && requesterLabelMatch.index < cityZipLabelMatch.index) {
      trimmedWindow = windowText.slice(0, requesterLabelMatch.index) + windowText.slice(cityZipLabelMatch.index);
    } else if (requesterLabelMatch) {
      trimmedWindow =
        windowText.slice(0, requesterLabelMatch.index) +
        windowText.slice(requesterLabelMatch.index + requesterLabelMatch[0].length);
    }

    const splitMatch = trimmedWindow.match(CITY_ZIP_LABEL_RE);
    let streetPart = trimmedWindow;
    let cityPart = '';
    if (splitMatch) {
      streetPart = trimmedWindow.slice(0, splitMatch.index);
      cityPart = trimmedWindow.slice(splitMatch.index + splitMatch[0].length);
    }

    streetAddress = extractStreetValue(
      trimTrailingCompanyName(trimToStreetStart(cleanCapture(streetPart.replace(/\s+/g, ' ').trim())))
    );

    const cityZipValueMatch = cityPart.match(CITY_STATE_ZIP_RE);
    if (cityZipValueMatch) {
      city = cityZipValueMatch[1].trim();
      stateZip = `${cityZipValueMatch[2]} ${cityZipValueMatch[3]}`.trim();
    } else if (cityPart.trim()) {
      city = cleanCapture(cityPart.replace(/\s+/g, ' ').trim());
    } else if (!splitMatch) {
      const wholeMatch = trimmedWindow.match(CITY_STATE_ZIP_RE);
      if (wholeMatch) {
        city = wholeMatch[1].trim();
        stateZip = `${wholeMatch[2]} ${wholeMatch[3]}`.trim();
        streetAddress = extractStreetValue(
          trimTrailingCompanyName(trimToStreetStart(cleanCapture(trimmedWindow.slice(0, wholeMatch.index).replace(/\s+/g, ' ').trim())))
        );
      }
    }
  }

  if (!streetAddress && !city) {
    const tinSectionIdx = text.indexOf(tinSection);
    const searchArea = tinSectionIdx >= 0 ? text.slice(0, tinSectionIdx) : text;
    const fallbackMatch = searchArea.match(CITY_STATE_ZIP_RE);
    if (fallbackMatch) {
      city = fallbackMatch[1].trim();
      stateZip = `${fallbackMatch[2]} ${fallbackMatch[3]}`.trim();
    }
  }

  return { streetAddress, city, stateZip };
}

function parseW9Fields(rawText, tinDigits, directFields, addressBoxText) {
  const text = rawText.replace(/\s+/g, ' ').trim();
  const tinSection = extractPartISection(text);

  const effectiveTinDigits = directFields
    ? {
        ssn: directFields.ssnDigits || (tinDigits && tinDigits.ssn) || '',
        ein: directFields.einDigits || (tinDigits && tinDigits.ein) || '',
      }
    : tinDigits;

  const companyNameSource = directFields && directFields.company_name ? 'acroform' : 'text-regex';
  const companyName = (directFields && directFields.company_name) || extractBusinessName(text);
  const { ssnNumber, einNumber, selectedType, taxId } = extractTaxId(text, effectiveTinDigits);

  let streetAddress;
  let city;
  let stateZip;
  let addressSource;
  const boxParsed = addressBoxText ? parseAddressBoxText(addressBoxText) : null;
  if (directFields && (directFields.street_address || directFields.city)) {
    streetAddress = directFields.street_address;
    city = directFields.city;
    stateZip = directFields.state_zip;
    addressSource = 'acroform';
  } else if (boxParsed) {
    ({ streetAddress, city, stateZip } = boxParsed);
    addressSource = (debugExtractionTrace && debugExtractionTrace.address.source) || 'ocr-box';
  } else {
    ({ streetAddress, city, stateZip } = extractAddress(text, tinSection));
    addressSource = 'text-regex';
  }

  const tinSource = directFields
    ? 'acroform'
    : (tinDigits && tinDigits.source) || (effectiveTinDigits && (effectiveTinDigits.ssn || effectiveTinDigits.ein) ? 'ocr-box' : 'text-regex');

  return {
    company_name: companyName,
    selected_tax_id_type: selectedType,
    tax_id_number: taxId,
    ssn_number: ssnNumber,
    ein_number: einNumber,
    street_address: streetAddress,
    city,
    state_zip: stateZip,
    _confidence: computeFieldConfidence({ companyName, ssnNumber, einNumber, selectedType, streetAddress, city, stateZip }, {
      companyNameSource,
      tinSource,
      addressSource,
    }),
  };
}

// Lightweight, purely computed confidence heuristics — no numeric OCR
// confidence is threaded through (Tesseract's per-word score is not
// currently plumbed to this layer), so this scores format validity and
// which detection tier produced the value: a dynamically-anchored or
// AcroForm-sourced value is trusted more than a bare text-regex fallback.
function computeFieldConfidence(fields, sources) {
  const tierScore = { acroform: 1, 'dynamic-anchor': 0.95, 'ocr-box': 0.85, 'dynamic-anchor+fixed-region': 0.85, 'fixed-region': 0.75, 'text-regex': 0.6 };

  const einValid = /^\d{2}-\d{7}$/.test(fields.einNumber || '');
  const ssnValid = /^\d{3}-\d{2}-\d{4}$/.test(fields.ssnNumber || '');
  const taxIdFormatValid = fields.selectedType === 'SSN' ? ssnValid : fields.selectedType === 'EIN' ? einValid : false;
  const taxId = fields.selectedType ? (taxIdFormatValid ? tierScore[sources.tinSource] ?? 0.5 : 0.2) : 0;

  const address = fields.streetAddress && fields.city && fields.stateZip
    ? tierScore[sources.addressSource] ?? 0.6
    : fields.streetAddress || fields.city
      ? (tierScore[sources.addressSource] ?? 0.6) * 0.6
      : 0;

  const companyName = fields.companyName ? (tierScore[sources.companyNameSource] ?? 0.6) : 0;

  return {
    companyName: Math.round(companyName * 100) / 100,
    taxId: Math.round(taxId * 100) / 100,
    address: Math.round(address * 100) / 100,
    sources,
  };
}

function renderW9Summary(data) {
  w9CompanyName.value = data.company_name || '';
  w9TaxId.value = data.tax_id_number || '';
  w9TaxIdType.value = data.selected_tax_id_type || 'EIN';
  const addressParts = [data.street_address, data.city, data.state_zip].filter(Boolean);
  w9Address.value = addressParts.join(', ');
  w9SummaryRevealed = true;
  w9Summary.hidden = false;
  w9EditHint.hidden = false;
}

const HAS_TEXT_THRESHOLD = 20;

const W9_TITLE_HEAD_CHARS = 300;
function looksLikeW9(text) {
  const head = text.slice(0, W9_TITLE_HEAD_CHARS);
  if (/\bform\s*w[\s=-]*9\b/i.test(head) || /\bw[\s=-]*9\b/i.test(head)) return true;
  // The "Form W-9" logo is rendered far larger than body text, so on a
  // rough scan (photo glare, fax dropout, heavy compression) it's often the
  // single worst-OCR'd thing on the page — sometimes garbled beyond any
  // reasonable typo tolerance, even bleeding noise into the "9" reliably
  // enough to misread it as an "8". The official title text next to it
  // ("Request for Taxpayer Identification Number and Certification") is
  // normal-weight body text and OCRs far more reliably; it's also required,
  // standardized wording unique to this form, so it's a solid fallback
  // signal when the logo itself is unreadable. Checking just its second
  // half tolerates "Request for Taxpayer" specifically being the part that
  // picked up noise bleeding over from the adjacent oversized logo.
  return loose('Identification Number and Certification').test(head);
}

function showRawText(text) {
  w9RawText.textContent = text;
  w9DebugRevealed = true;
  w9Debug.hidden = false;
}

// Debug-only: sanitized console trace (confidence + which detection tier
// won for each field — never the extracted digits/address/name text) plus a
// bbox overlay drawn on the OCR'd page image, appended into the existing
// raw-text debug panel. No-ops entirely unless W9_DEBUG is on.
function logExtractionTrace(data) {
  if (!W9_DEBUG) return;
  console.groupCollapsed('[W9 extraction trace]');
  console.log('company_name:', data._confidence.sources.companyNameSource, 'confidence:', data._confidence.companyName);
  console.log(
    'tax id:',
    data.selected_tax_id_type || 'none',
    'via',
    data._confidence.sources.tinSource,
    'confidence:',
    data._confidence.taxId
  );
  console.log('address:', data._confidence.sources.addressSource, 'confidence:', data._confidence.address);
  if (debugExtractionTrace) console.log('bboxes:', debugExtractionTrace);
  console.groupEnd();
}

function renderDebugOverlay() {
  if (!W9_DEBUG || !debugExtractionTrace || !debugExtractionTrace.canvas) return;
  const src = debugExtractionTrace.canvas;
  let overlay = document.getElementById('w9DebugOverlay');
  if (!overlay) {
    overlay = document.createElement('canvas');
    overlay.id = 'w9DebugOverlay';
    overlay.style.maxWidth = '100%';
    w9Debug.appendChild(overlay);
  }
  overlay.width = src.width;
  overlay.height = src.height;
  const ctx = overlay.getContext('2d');
  ctx.drawImage(src, 0, 0);

  const boxStyles = [
    { bbox: debugExtractionTrace.tin.cropRegion, color: '#22c55e', label: `TIN crop (${debugExtractionTrace.tin.source || '?'})` },
    { bbox: debugExtractionTrace.address.cropRegion, color: '#a855f7', label: `Address crop (${debugExtractionTrace.address.source || '?'})` },
  ];
  ctx.lineWidth = 3;
  ctx.font = '20px sans-serif';
  boxStyles.forEach(({ bbox, color, label }) => {
    if (!bbox) return;
    ctx.strokeStyle = color;
    ctx.strokeRect(bbox.x0, bbox.y0, bbox.x1 - bbox.x0, bbox.y1 - bbox.y0);
    ctx.fillStyle = color;
    ctx.fillText(label, bbox.x0, Math.max(16, bbox.y0 - 6));
  });

  [debugExtractionTrace.tin.cropCanvas, debugExtractionTrace.address.cropCanvas].filter(Boolean).forEach((cropCanvas) => {
    const thumb = document.createElement('canvas');
    thumb.width = cropCanvas.width;
    thumb.height = cropCanvas.height;
    thumb.style.maxWidth = '100%';
    thumb.style.border = '1px solid currentColor';
    thumb.getContext('2d').drawImage(cropCanvas, 0, 0);
    w9Debug.appendChild(thumb);
  });
}

function finishExtraction(text, viaOcr, tinDigits, directFields, addressBoxText) {
  extractedW9Data = parseW9Fields(text, tinDigits, directFields, addressBoxText);
  renderW9Summary(extractedW9Data);
  showRawText(text);
  logExtractionTrace(extractedW9Data);
  renderDebugOverlay();

  const source = viaOcr ? ' (via OCR)' : '';
  if (!extractedW9Data.selected_tax_id_type) {
    setStatus(`W-9 read${source}, but no SSN or EIN was found — fill it in manually before generating.`, true);
  } else if (extractedW9Data.ssn_number && extractedW9Data.ein_number) {
    setStatus(
      `W-9 parsed${source}. Both an SSN and EIN were found — using ${extractedW9Data.selected_tax_id_type} (${extractedW9Data.tax_id_number}). Switch the dropdown to use the other one instead.`
    );
  } else {
    setStatus(
      `W-9 parsed${source}. Tax ID used: ${extractedW9Data.selected_tax_id_type} (${extractedW9Data.tax_id_number}).`
    );
  }
}

function revealManualW9Entry(message) {
  extractedW9Data = {};
  renderW9Summary(extractedW9Data);
  setStatus(message, true);
  refreshW9ValidationStatus();
}

async function processW9(file) {
  if (isImageFile(file)) {
    setStatus(`Reading ${file.name} — running OCR (this can take a few seconds)...`);
    try {
      const { text: ocrText, tinDigits, addressBoxText } = await ocrImageFile(file, (page, totalPages, m) => {
        if (m.status === 'recognizing text') {
          setStatus(`Running OCR... ${Math.round(m.progress * 100)}%`);
        }
      });

      if (ocrText.replace(/\s+/g, '').length < HAS_TEXT_THRESHOLD) {
        showRawText(ocrText);
        revealManualW9Entry(`OCR could not read any text from ${file.name} — fill in the fields below manually.`);
        return;
      }
      if (!looksLikeW9(ocrText)) {
        showRawText(ocrText);
        revealManualW9Entry(
          `${file.name} doesn't look like a W-9 — "W-9" wasn't found at the top of the image. Fill in the fields below manually, or upload a correct W-9.`
        );
        return;
      }
      finishExtraction(ocrText, true, tinDigits, undefined, addressBoxText);
    } catch (err) {
      revealManualW9Entry(`Could not read ${file.name} as a W-9 image: ${err.message} — fill in the fields below manually.`);
    }
    return;
  }

  setStatus(`Reading ${file.name}...`);
  try {
    const { text, directFields } = await extractPdfText(file);
    const hasUsableText = text.replace(/\s+/g, '').length >= HAS_TEXT_THRESHOLD;
    if (hasUsableText && looksLikeW9(text)) {
      finishExtraction(text, false, undefined, directFields);
      return;
    }

    setStatus(`${file.name} looks scanned — running OCR (this can take a few seconds)...`);
    const { text: ocrText, tinDigits, addressBoxText } = await ocrPdfText(file, (page, totalPages, m) => {
      if (m.status === 'recognizing text') {
        setStatus(`Running OCR on page ${page}/${totalPages}... ${Math.round(m.progress * 100)}%`);
      }
    });

    if (ocrText.replace(/\s+/g, '').length < HAS_TEXT_THRESHOLD) {
      showRawText(ocrText);
      revealManualW9Entry(`OCR could not read any text from ${file.name} — fill in the fields below manually.`);
      return;
    }

    if (!looksLikeW9(ocrText)) {
      showRawText(ocrText);
      revealManualW9Entry(
        `${file.name} doesn't look like a W-9 — "W-9" wasn't found at the top of the document. Fill in the fields below manually, or upload a correct W-9 PDF.`
      );
      return;
    }

    finishExtraction(ocrText, true, tinDigits, undefined, addressBoxText);
  } catch (err) {
    revealManualW9Entry(`Could not read ${file.name} as a W-9 PDF: ${err.message} — fill in the fields below manually.`);
  }
}

function setStatus(message, isError = false) {
  status.textContent = message;
  status.classList.toggle('error', isError);

  status.classList.remove('flash');
  void status.offsetWidth;
  status.classList.add('flash');

  if (!wizardLoading.hidden) {
    wizardLoadingText.textContent = message;
  }
}

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
  return source.arrayBuffer();
}

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
    if (count > 200) break;
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

const DOCX_PAGE_PX_TO_PT = 72 / 96;

async function renderDocxToPdf(docxBytes) {
  const container = document.createElement('div');
  container.style.position = 'fixed';
  container.style.left = '-10000px';
  container.style.top = '0';
  container.style.zIndex = '-1';
  document.body.appendChild(container);

  const styleContainer = document.createElement('div');
  container.appendChild(styleContainer);

  try {
    const blob = new Blob([docxBytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await renderDocxAsync(blob, container, styleContainer, {
      inWrapper: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
    });

    const pageEls = Array.from(container.querySelectorAll('section.docx'));
    const pdfDoc = await PDFLib.PDFDocument.create();

    for (const pageEl of pageEls) {
      const canvas = await html2canvas(pageEl, { scale: 2, backgroundColor: '#ffffff' });
      const pngBlob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
      const pngBytes = await pngBlob.arrayBuffer();

      const pdfWidth = pageEl.offsetWidth * DOCX_PAGE_PX_TO_PT;
      const pdfHeight = pageEl.offsetHeight * DOCX_PAGE_PX_TO_PT;

      const image = await pdfDoc.embedPng(pngBytes);
      const page = pdfDoc.addPage([pdfWidth, pdfHeight]);
      page.drawImage(image, { x: 0, y: 0, width: pdfWidth, height: pdfHeight });
    }

    return pdfDoc.save();
  } finally {
    document.body.removeChild(container);
  }
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

async function insertDataAndAdvance() {
  if (step1NextBtn.classList.contains('is-loading')) return false;

  const startedAt = Date.now();
  step1NextBtn.classList.add('is-loading');
  step1NextBtn.disabled = true;
  showWizardLoading('Filling the MSA template with the extracted data...');

  let success = false;
  try {
    await loadMsaTemplates();
    const autoReplacements = buildAutoReplacements(extractedW9Data);
    const { bytes: docxBytes, replacedCount } = await applyDocxReplacements(msaTemplateDocxFile, autoReplacements);
    success = replacedCount > 0;
    setStatus(
      success
        ? 'MSA template filled with the extracted W-9 data. Continuing to manual details.'
        : "The MSA template doesn't contain the expected placeholder text, so nothing was auto-replaced.",
      !success
    );

    if (success) {
      try {
        await setDocumentPreview(docxBytes);
      } catch {
        // Preview is a convenience; the template scan already succeeded, so don't block progression.
      }
    }
  } catch (err) {
    setStatus(`Could not prepare the MSA template: ${err.message}`, true);
  } finally {
    await waitRemaining(startedAt, MIN_LOADING_MS);
    hideWizardLoading();
    step1NextBtn.classList.remove('is-loading');
    refreshW9ValidationStatus();
    if (success) showWizardStep(2);
  }
  return success;
}

step1NextBtn.addEventListener('click', insertDataAndAdvance);

let livePreviewTimer = null;

function scheduleLivePreviewUpdate() {
  if (!msaTemplateDocxFile || !extractedW9Data) return;
  clearTimeout(livePreviewTimer);
  livePreviewTimer = setTimeout(async () => {
    try {
      const replacements = [
        ...buildAutoReplacements(extractedW9Data),
        { type: 'afterLabel', label: MANUAL_FIELD_LABELS.rep, value: manualRep.value.trim() },
        { type: 'afterLabel', label: MANUAL_FIELD_LABELS.role, value: manualRole.value.trim() },
        { type: 'afterLabel', label: MANUAL_FIELD_LABELS.location, value: manualLocation.value.trim() },
        { type: 'afterLabel', label: MANUAL_FIELD_LABELS.startDate, value: manualStartDate.value.trim() },
        {
          type: 'afterLabel',
          label: MANUAL_FIELD_LABELS.billingRate,
          value: manualBillingRate.value.trim() ? `$${manualBillingRate.value.trim()}/hr` : '',
        },
      ];
      const { bytes: docxBytes } = await applyDocxReplacements(msaTemplateDocxFile, replacements);
      await setDocumentPreview(docxBytes);
    } catch {
      // Live preview is best-effort; ignore failures while the user is still typing.
    }
  }, 600);
}

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
    scheduleLivePreviewUpdate();
  });
});

applyManualBtn.addEventListener('click', async () => {
  const startedAt = Date.now();
  downloadPdfBtn.disabled = true;
  downloadWordBtn.disabled = true;
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

    const companyName = (extractedW9Data?.company_name || '').replace(/[\\/:*?"<>|]+/g, '').trim();
    const baseName = companyName ? `MSA and PO - ${companyName}` : 'MSA and PO';

    const { bytes: docxBytes } = await applyDocxReplacements(msaTemplateDocxFile, allReplacements);
    const wordBlob = new Blob([docxBytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const pdfBytes = await renderDocxToPdf(docxBytes);
    const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
    setGeneratedFiles(pdfBlob, wordBlob, baseName);
    await setDocumentPreview(docxBytes);

    setStatus('Manual details applied and the MSA is finalized. Ready to download.');
    success = true;
  } catch (err) {
    setStatus(`Could not finalize the MSA: ${err.message}`, true);
  } finally {
    await waitRemaining(startedAt, MIN_LOADING_MS);
    hideWizardLoading();
    applyManualBtn.disabled = false;
    if (success) showWizardStep(3);
  }
});

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

function resetWizardForNewDocument() {
  uploadedFiles = [];
  fileInput.value = '';
  renderFileList();
  extractDataBtn.disabled = true;
  step1NextBtn.disabled = true;

  extractedW9Data = null;
  w9SummaryRevealed = false;
  w9DebugRevealed = false;
  w9Summary.hidden = true;
  w9EditHint.hidden = true;
  w9Debug.hidden = true;
  w9RawText.textContent = '';
  [w9CompanyName, w9TaxId, w9Address].forEach((el) => {
    el.value = '';
    el.classList.remove('field-error');
  });
  w9TaxIdType.value = 'EIN';

  MANUAL_FIELD_INPUTS.forEach((el) => {
    el.value = '';
    el.classList.remove('field-error');
  });
  manualStartDatePicker.value = '';

  if (generatedPdfUrl) {
    URL.revokeObjectURL(generatedPdfUrl);
    generatedPdfUrl = null;
  }
  if (generatedWordUrl) {
    URL.revokeObjectURL(generatedWordUrl);
    generatedWordUrl = null;
  }
  downloadPdfBtn.disabled = true;
  downloadWordBtn.disabled = true;

  clearDocumentPreview();
  setStatus('Ready for a new document.');
  showWizardStep(1);
}

startNewBtn.addEventListener('click', resetWizardForNewDocument);

const previewToggleBtn = document.getElementById('previewToggleBtn');
const previewToggleLabel = document.getElementById('previewToggleLabel');
const previewPanel = document.getElementById('previewPanel');
const previewEmptyState = document.getElementById('previewEmptyState');
const previewLoading = document.getElementById('previewLoading');
const previewZoomControls = document.getElementById('previewZoomControls');
const previewZoomOutBtn = document.getElementById('previewZoomOutBtn');
const previewZoomInBtn = document.getElementById('previewZoomInBtn');
const previewZoomLevel = document.getElementById('previewZoomLevel');
const previewFrame = document.getElementById('previewFrame');

const PREVIEW_ZOOM_MIN = 0.4;
const PREVIEW_ZOOM_MAX = 2;
const PREVIEW_ZOOM_STEP = 0.1;
let previewZoom = 1;

function applyPreviewZoom(animate) {
  const doc = previewFrame.contentDocument;
  const wrapperEl = doc?.querySelector('.docx-wrapper');
  if (doc?.body && wrapperEl) {
    if (!wrapperEl.dataset.naturalWidth) {
      wrapperEl.dataset.naturalWidth = wrapperEl.scrollWidth;
      wrapperEl.dataset.naturalHeight = wrapperEl.scrollHeight;
    }
    const naturalWidth = Number(wrapperEl.dataset.naturalWidth);
    const naturalHeight = Number(wrapperEl.dataset.naturalHeight);
    wrapperEl.style.transition = animate ? 'transform 0.15s ease' : 'none';
    wrapperEl.style.transform = `scale(${previewZoom})`;
    doc.body.style.width = `${naturalWidth * previewZoom}px`;
    doc.body.style.height = `${naturalHeight * previewZoom}px`;
  }
  previewZoomLevel.textContent = `${Math.round(previewZoom * 100)}%`;
}

function setPreviewZoom(zoom, animate) {
  previewZoom = Math.min(PREVIEW_ZOOM_MAX, Math.max(PREVIEW_ZOOM_MIN, zoom));
  applyPreviewZoom(animate);
}

// Zooms while keeping the document point under the cursor visually fixed in
// the viewport — the way PDF viewers anchor zoom to the cursor instead of
// always scaling from the top-left corner. cursorX/cursorY are relative to
// the preview iframe's own viewport (i.e. straight from a pointer/wheel
// event dispatched inside that iframe's document).
function zoomPreviewAt(zoom, cursorX, cursorY, animate) {
  const doc = previewFrame.contentDocument;
  const scroller = doc?.scrollingElement || doc?.documentElement;
  const oldZoom = previewZoom || 1;

  let docX = 0;
  let docY = 0;
  if (scroller) {
    docX = (scroller.scrollLeft + cursorX) / oldZoom;
    docY = (scroller.scrollTop + cursorY) / oldZoom;
  }

  setPreviewZoom(zoom, animate);

  if (scroller) {
    scroller.scrollLeft = docX * previewZoom - cursorX;
    scroller.scrollTop = docY * previewZoom - cursorY;
  }
}

function zoomPreviewAtCenter(zoom, animate) {
  zoomPreviewAt(zoom, previewFrame.clientWidth / 2, previewFrame.clientHeight / 2, animate);
}

previewZoomOutBtn.addEventListener('click', () => zoomPreviewAtCenter(previewZoom - PREVIEW_ZOOM_STEP, true));
previewZoomInBtn.addEventListener('click', () => zoomPreviewAtCenter(previewZoom + PREVIEW_ZOOM_STEP, true));

function writeIframeShell(iframe) {
  const doc = iframe.contentDocument;
  doc.open();
  doc.write(
    '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
      'html,body{margin:0;background:#fff;} body{cursor:grab;}' +
      '</style></head><body></body></html>'
  );
  doc.close();
  return doc;
}

// docx-preview clears the style container's innerHTML when it renders, so any
// override style has to be appended to <head> *after* rendering, not before.
function applyPreviewZoomStyleOverrides(doc) {
  const style = doc.createElement('style');
  style.textContent =
    'html,body{background:gray !important;} ' +
    '.docx-wrapper{transform-origin:top left !important;align-items:flex-start !important;will-change:transform;min-width:100%;min-height:100%;}';
  doc.head.appendChild(style);
}

function attachDragPan(doc) {
  const scroller = doc.scrollingElement || doc.documentElement;
  let dragging = false;
  let pointerId = null;
  let startX = 0;
  let startY = 0;
  let startLeft = 0;
  let startTop = 0;

  doc.body.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    dragging = true;
    pointerId = e.pointerId;
    startX = e.clientX;
    startY = e.clientY;
    startLeft = scroller.scrollLeft;
    startTop = scroller.scrollTop;
    doc.body.style.cursor = 'grabbing';
    doc.body.style.userSelect = 'none';
    try {
      doc.body.setPointerCapture(pointerId);
    } catch {
      // Pointer capture is a nice-to-have for dragging past the frame edge; ignore if unsupported.
    }
    e.preventDefault();
  });

  doc.body.addEventListener('pointermove', (e) => {
    if (!dragging) return;
    scroller.scrollLeft = startLeft - (e.clientX - startX);
    scroller.scrollTop = startTop - (e.clientY - startY);
  });

  const stopDrag = (e) => {
    if (!dragging) return;
    dragging = false;
    doc.body.style.cursor = 'grab';
    doc.body.style.userSelect = '';
    try {
      if (pointerId !== null) doc.body.releasePointerCapture(pointerId);
    } catch {
      // Already released or unsupported; nothing to do.
    }
  };
  doc.body.addEventListener('pointerup', stopDrag);
  doc.body.addEventListener('pointercancel', stopDrag);
  doc.body.addEventListener('pointerleave', (e) => {
    if (pointerId === null || !doc.body.hasPointerCapture?.(pointerId)) stopDrag(e);
  });
}

function attachWheelZoom(doc) {
  doc.addEventListener(
    'wheel',
    (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      const step = Math.min(0.15, Math.max(0.01, Math.abs(e.deltaY) / 200));
      zoomPreviewAt(previewZoom + (e.deltaY < 0 ? step : -step), e.clientX, e.clientY, false);
    },
    { passive: false }
  );

  doc.addEventListener('dblclick', (e) => {
    e.preventDefault();
    const target = previewZoom >= PREVIEW_ZOOM_MAX - PREVIEW_ZOOM_STEP / 2 ? 1 : previewZoom + PREVIEW_ZOOM_STEP * 2;
    zoomPreviewAt(target, e.clientX, e.clientY, true);
  });
}

async function setDocumentPreview(docxBytes) {
  previewLoading.hidden = false;
  try {
    const blob = new Blob([docxBytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const doc = writeIframeShell(previewFrame);
    previewFrame.hidden = false;
    previewEmptyState.hidden = true;

    await renderDocxAsync(blob, doc.body, doc.head, {
      inWrapper: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
    });

    applyPreviewZoomStyleOverrides(doc);
    applyPreviewZoom(false);
    attachDragPan(doc);
    attachWheelZoom(doc);
    previewZoomControls.hidden = false;
    previewFrame.classList.remove('is-ready');
    requestAnimationFrame(() => previewFrame.classList.add('is-ready'));
  } finally {
    previewLoading.hidden = true;
  }
}

function clearDocumentPreview() {
  if (previewFrame.contentDocument) writeIframeShell(previewFrame);
  previewFrame.hidden = true;
  previewFrame.classList.remove('is-ready');
  previewLoading.hidden = true;
  previewZoomControls.hidden = true;
  previewEmptyState.hidden = false;
}

function setPreviewPanelOpen(open) {
  previewPanel.classList.toggle('is-open', open);
  previewPanel.setAttribute('aria-hidden', String(!open));
  previewToggleBtn.setAttribute('aria-expanded', String(open));
  previewToggleLabel.textContent = open ? 'Close' : 'Preview';
}

previewToggleBtn.addEventListener('click', () => {
  setPreviewPanelOpen(!previewPanel.classList.contains('is-open'));
});

const HISTORY_DB_NAME = 'documentGeneratorHistory';
const HISTORY_STORE = 'generatedDocuments';
const HISTORY_SESSION_KEY = 'docGenHistorySessionId';

function getHistorySessionId() {
  let id = sessionStorage.getItem(HISTORY_SESSION_KEY);
  if (!id) {
    id = crypto.randomUUID();
    sessionStorage.setItem(HISTORY_SESSION_KEY, id);
  }
  return id;
}

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
    tx.objectStore(HISTORY_STORE).add({
      baseName,
      pdfBlob,
      wordBlob,
      createdAt: Date.now(),
      sessionId: getHistorySessionId(),
    });
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function purgeStaleHistoryEntries(db, all, sessionId) {
  const stale = all.filter((entry) => entry.sessionId !== sessionId);
  if (!stale.length) return;
  const tx = db.transaction(HISTORY_STORE, 'readwrite');
  const store = tx.objectStore(HISTORY_STORE);
  stale.forEach((entry) => store.delete(entry.id));
}

async function purgeStaleHistoryOnLoad() {
  try {
    const db = await openHistoryDb();
    const sessionId = getHistorySessionId();
    const all = await new Promise((resolve, reject) => {
      const request = db.transaction(HISTORY_STORE, 'readonly').objectStore(HISTORY_STORE).getAll();
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    await purgeStaleHistoryEntries(db, all, sessionId);
  } catch (err) {
    console.error('History cleanup failed:', err);
  }
}
purgeStaleHistoryOnLoad();

async function getHistoryEntries() {
  const db = await openHistoryDb();
  const sessionId = getHistorySessionId();
  const all = await new Promise((resolve, reject) => {
    const request = db.transaction(HISTORY_STORE, 'readonly').objectStore(HISTORY_STORE).getAll();
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });

  await purgeStaleHistoryEntries(db, all, sessionId);

  return all
    .filter((entry) => entry.sessionId === sessionId)
    .sort((a, b) => b.createdAt - a.createdAt);
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
        <span class="history-item-name"></span>
        <span class="history-item-date"></span>
      </div>
      <div class="history-item-actions">
        <button type="button" class="btn btn-secondary" data-action="pdf">PDF</button>
        <button type="button" class="btn btn-secondary" data-action="word">Word</button>
        <button type="button" class="history-delete-btn" data-action="delete" aria-label="Delete this document">✕</button>
      </div>
    `;
    item.querySelector('.history-item-name').textContent = entry.baseName;
    item.querySelector('.history-item-date').textContent = formatHistoryDate(entry.createdAt);
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

(function () {
  const overlay = document.getElementById('introOverlay');
  if (!overlay) return;

  const man = document.getElementById('man');
  const elevator = document.getElementById('elevator');

  document.body.style.overflow = 'hidden';

  function runIntroAnimation() {
    setTimeout(() => {
      elevator.classList.add('doors-open');
    }, 300);

    setTimeout(() => {
      man.classList.add('walk-in', 'stepping');
    }, 1400);
    setTimeout(() => {
      man.classList.remove('stepping');
    }, 3600);

    setTimeout(() => {
      elevator.classList.remove('doors-open');
    }, 3700);

    setTimeout(() => {
      man.classList.add('entered');
    }, 4750);

    setTimeout(() => {
      elevator.classList.add('morphed-button');
    }, 5300);

    setTimeout(() => {
      continueToApp();
    }, 6100);
  }

  function continueToApp() {
    const FLY_DURATION = 850;

    const startRect = elevator.getBoundingClientRect();

    const targetRect = brandLogo.getBoundingClientRect();

    const dx = targetRect.left - startRect.left;
    const dy = targetRect.top - startRect.top;
    const sx = targetRect.width / startRect.width;
    const sy = targetRect.height / startRect.height;

    elevator.style.position = 'fixed';
    elevator.style.margin = '0';
    elevator.style.left = startRect.left + 'px';
    elevator.style.top = startRect.top + 'px';
    elevator.style.width = startRect.width + 'px';
    elevator.style.height = startRect.height + 'px';
    elevator.style.transformOrigin = 'top left';
    elevator.style.transition = 'none';
    elevator.style.transform = 'translate(0px, 0px) scale(1, 1)';

    void elevator.offsetWidth;

    elevator.style.transition = `transform ${FLY_DURATION}ms cubic-bezier(0.4, 0, 0.2, 1)`;
    elevator.style.transform = `translate(${dx}px, ${dy}px) scale(${sx}, ${sy})`;

    setTimeout(() => {
      overlay.classList.add('is-hidden');
      document.body.style.overflow = '';
      elevator.style.transition = 'opacity 0.35s ease';
      elevator.style.opacity = '0';

      brandLogo.classList.add('cta-pulse');
      setTimeout(() => brandLogo.classList.remove('cta-pulse'), 1600);
    }, FLY_DURATION);

    setTimeout(() => overlay.remove(), FLY_DURATION + 750);
  }

  runIntroAnimation();
})();
