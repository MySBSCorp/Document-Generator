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
const manualFieldsStatus = document.getElementById('manualFieldsStatus');

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
  debugExtractionTrace = W9_DEBUG ? { tin: {}, address: {}, rotationDegrees: 0, orientationConfidence: null } : null;
}

function showWizardStep(n) {
  wizardSteps.forEach((el) => {
    el.classList.toggle('active', Number(el.dataset.step) === n);
  });
  const showW9Panel = n < 3;
  w9Summary.hidden = !(showW9Panel && w9SummaryRevealed);
  w9EditHint.hidden = !(showW9Panel && w9SummaryRevealed);
  w9Debug.hidden = !(showW9Panel && w9DebugRevealed);
  startNewBtn.hidden = n !== 3;

  if (n === 3) {
    status.textContent = '';
    status.classList.remove('error');
  }
}

function positionWizardLoading() {
  // On the mobile layout (see style.css's 900px breakpoint) the action panel
  // is a normal, natural-flow block that can be taller than the viewport and
  // scrolls with the page — so a `position: fixed` overlay sized to its
  // getBoundingClientRect() (captured once, at whatever scroll position the
  // page happened to be at) drifts out of alignment, or off-screen entirely,
  // the moment the user scrolls. The desktop layout doesn't have that
  // problem (the action panel is a fixed-height pane that never scrolls the
  // page itself), so only mobile needs the full-viewport fallback.
  if (window.innerWidth <= 900) {
    wizardLoading.style.top = '0px';
    wizardLoading.style.left = '0px';
    wizardLoading.style.width = '100vw';
    wizardLoading.style.height = '100vh';
    return;
  }
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

  const w9File = uploadedFiles.find(isW9UploadCandidate);
  extractDataBtn.disabled = !w9File;
  step1NextBtn.disabled = true;
  renderW9Preview(w9File);
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

function formatTaxId(digits, type) {
  if (digits.length !== 9) return digits;
  if (type === 'SSN') return `${digits.slice(0, 3)}-${digits.slice(3, 5)}-${digits.slice(5)}`;
  if (type === 'EIN') return `${digits.slice(0, 2)}-${digits.slice(2)}`;
  return digits;
}

// Builds a dash-formatted tax ID from whatever digits have been typed so
// far (not just once all 9 are in), so the hyphen appears live as the user
// types instead of only after the field is complete.
function formatTaxIdPartial(raw, type) {
  const digits = raw.replace(/\D/g, '').slice(0, 9);
  const groups = type === 'SSN' ? [3, 2, 4] : [2, 7];
  let out = '';
  let i = 0;
  for (const len of groups) {
    const part = digits.slice(i, i + len);
    if (!part) break;
    out += (out ? '-' : '') + part;
    i += len;
  }
  return out;
}

w9CompanyName.addEventListener('input', () => {
  if (extractedW9Data) extractedW9Data.company_name = w9CompanyName.value.trim();
  refreshW9ValidationStatus();
  scheduleLivePreviewUpdate();
});
w9TaxId.addEventListener('input', () => {
  const formatted = formatTaxIdPartial(w9TaxId.value, w9TaxIdType.value);
  w9TaxId.value = formatted;
  const value = formatted;
  if (extractedW9Data) extractedW9Data.tax_id_number = value;

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
    newValue = formatTaxIdPartial(newValue, newType);
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
  return downscaleTo(source, DESKEW_SAMPLE_MAX_DIM);
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

// Used exclusively for the SSN/EIN crop's label/row detection (see
// tinDigitsFromCrop). pageSegMode defaults to Tesseract's own default (PSM
// 3, fully-automatic) when omitted.
async function ocrLinesWithBoxes(canvas, pageSegMode) {
  const worker = await Tesseract.createWorker('eng', undefined, TESSERACT_VENDORED_PATHS);
  if (pageSegMode) await worker.setParameters({ tessedit_pageseg_mode: pageSegMode });
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

// --- Page orientation (0/90/180/270) ---------------------------------------
// deskewCanvas above only straightens a slight (<=8deg) scanning skew; it
// can't recover a page that was fed into the scanner sideways or upside
// down. This mirrors the reference pipeline's orientation step: OCR a small
// header/name-band crop at each right-angle rotation and keep whichever
// orientation reads with the highest average word confidence — a page
// OCR'd in the wrong orientation reads as near-random noise, so confidence
// drops sharply, while the correct orientation reads cleanly.
const ORIENTATION_DEGREES = [0, 90, 180, 270];
// Fraction-of-page crop covering the header/entity-name band — present and
// legible near the top of every W-9 regardless of layout drift, and small
// enough that probing it four times stays cheap.
const ORIENTATION_PROBE_REGION = { x0: 0.03, x1: 0.98, y0: 0.02, y1: 0.24 };

function rotateCanvasExpand(source, degrees) {
  if (degrees % 360 === 0) return source;
  const rad = (degrees * Math.PI) / 180;
  const swapDims = degrees % 180 !== 0;
  const w = source.width;
  const h = source.height;
  const out = document.createElement('canvas');
  out.width = swapDims ? h : w;
  out.height = swapDims ? w : h;
  const ctx = out.getContext('2d');
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, out.width, out.height);
  ctx.translate(out.width / 2, out.height / 2);
  ctx.rotate(rad);
  ctx.drawImage(source, -w / 2, -h / 2);
  return out;
}

// PDF pages are rendered at 3x scale before this runs, so the full-size
// canvas is large; rotating and OCR'ing that four times over (once per
// candidate orientation) is expensive enough to create real memory/CPU
// pressure right before the actual extraction OCR passes that follow —
// measured as a regression on at least one low-quality fixture where the
// real OCR pass came back worse purely from that added load, despite the
// orientation choice itself (0deg) being correct. Probing on a small
// downscaled copy instead keeps all four probes cheap; the expensive
// full-resolution rotate then runs at most once, only for the orientation
// actually chosen (and not at all when that's 0deg).
const ORIENTATION_PROBE_MAX_DIM = 1200;

function downscaleTo(source, maxDim) {
  const scale = Math.min(1, maxDim / Math.max(source.width, source.height));
  if (scale >= 1) return source;
  const out = document.createElement('canvas');
  out.width = Math.round(source.width * scale);
  out.height = Math.round(source.height * scale);
  out.getContext('2d').drawImage(source, 0, 0, out.width, out.height);
  return out;
}

// Runs all four orientation probes on a single shared worker (spinning up a
// Tesseract worker, not the recognize call itself, is the expensive part),
// so this stays one extra pass rather than four full ones. Ties keep the
// earliest-tried orientation (0 first), so an ambiguous/blank page is left
// unrotated rather than flipped on a coin toss.
async function pickBestOrientation(rawCanvas) {
  const probeBase = downscaleTo(rawCanvas, ORIENTATION_PROBE_MAX_DIM);
  const worker = await Tesseract.createWorker('eng', undefined, TESSERACT_VENDORED_PATHS);
  let bestDegrees = 0;
  let bestConfidence = -1;

  for (const degrees of ORIENTATION_DEGREES) {
    const candidate = rotateCanvasExpand(probeBase, degrees);
    const probe = cropCanvasRegion(candidate, ORIENTATION_PROBE_REGION, 1);
    if (!probe) continue;
    const preprocessed = grayscaleContrastCanvas(probe);
    const result = await worker.recognize(preprocessed, {}, { blocks: true });
    const lines = linesFromBlocks(result.data.blocks);
    const confidences = lines
      .map((l) => l.confidence)
      .filter((c) => typeof c === 'number' && c >= 0);
    const confidence = confidences.length ? confidences.reduce((a, b) => a + b, 0) / confidences.length : 0;
    if (confidence > bestConfidence) {
      bestConfidence = confidence;
      bestDegrees = degrees;
    }
  }

  await worker.terminate();
  const canvas = bestDegrees === 0 ? rawCanvas : rotateCanvasExpand(rawCanvas, bestDegrees);
  return { canvas, degrees: bestDegrees, confidence: bestConfidence };
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
// Boxed tax-ID forms print each digit inside its own bordered cell, and
// those border strokes sit immediately next to the digit glyphs. Verified
// by eye against several scans where the digit itself was perfectly legible
// but Tesseract still silently dropped it under every page-segmentation
// mode tried — the border stroke was fusing into the neighboring glyph
// during character segmentation. A box border is easy to tell apart from a
// digit stroke on its own terms: it runs near-solid-dark for almost the
// entire cell height, where even the tallest digit strokes ("1", "7") only
// partially cover it. Painting those near-full-height columns white before
// OCR removes the interference without needing to know where any
// individual digit boundary falls.
function whiteOutGridLines(ctx, width, height) {
  const { data } = ctx.getImageData(0, 0, width, height);
  const colDark = new Array(width).fill(0);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const i = (y * width + x) * 4;
      const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      if (lum < 150) colDark[x]++;
    }
  }
  // A perfectly vertical border column is dark for nearly its whole height,
  // but a scan with even a slight rotation renders that same border as a
  // shallow diagonal — its darkness spreads across a few neighboring
  // columns as y increases, so no single column reaches the height
  // threshold on its own. Scoring each column by the best-covered column
  // within a small window around it tolerates that drift (a few degrees of
  // skew over a digit row's height) while still only affecting the narrow
  // band immediately around a real border stroke.
  const DRIFT_RADIUS = 2;
  const threshold = height * 0.7;
  const isBorderCol = new Array(width).fill(false);
  for (let x = 0; x < width; x++) {
    let best = 0;
    for (let dx = -DRIFT_RADIUS; dx <= DRIFT_RADIUS; dx++) {
      const xi = x + dx;
      if (xi >= 0 && xi < width) best = Math.max(best, colDark[xi]);
    }
    if (best >= threshold) isBorderCol[x] = true;
  }
  let changed = false;
  for (let x = 0; x < width; x++) {
    if (!isBorderCol[x]) continue;
    changed = true;
    for (let y = 0; y < height; y++) {
      const i = (y * width + x) * 4;
      data[i] = data[i + 1] = data[i + 2] = 255;
    }
  }
  if (changed) ctx.putImageData(new ImageData(data, width, height), 0, 0);
}

// A photo taken in poor lighting (rather than a flatbed scan) can leave
// printed digits and the box border only a few shades darker than the
// paper — real ink, but too low-contrast for Tesseract's binarization to
// separate from the background at all, let alone for the border-column
// detection above to tell a border from a digit. Stretching the crop's own
// darkest-to-lightest pixel range out to full black-to-white before either
// step fixes both: it costs nothing on an already high-contrast scan
// (its range is already close to 0-255) and can recover an otherwise
// invisible-to-OCR photo.
function stretchContrast(ctx, width, height) {
  const imgData = ctx.getImageData(0, 0, width, height);
  const d = imgData.data;
  let min = 255;
  let max = 0;
  for (let i = 0; i < d.length; i += 4) {
    const g = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
    if (g < min) min = g;
    if (g > max) max = g;
  }
  const range = max - min;
  // Skip entirely once the crop already reaches close to white: measured
  // against a real regression — this same stretch, run unconditionally,
  // pushed a crisp CamScanner-quality digit row's already near-full-range
  // contrast (min 0, max ~254) a little further, and that was enough to
  // turn a legible "9" into something the classifier read as "0". A photo
  // taken in bad lighting doesn't have this problem to begin with: its
  // brightest pixel (the paper background) never gets close to white
  // (measured around 155-175, not 240+), which is exactly the signal that
  // there's real headroom to recover rather than noise to amplify.
  if (range <= 10 || max >= 240) return;
  const scale = 255 / range;
  for (let i = 0; i < d.length; i += 4) {
    const g = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
    const v = Math.max(0, Math.min(255, (g - min) * scale));
    d[i] = d[i + 1] = d[i + 2] = v;
  }
  ctx.putImageData(imgData, 0, 0);
}

async function ocrDigitRowAtBbox(worker, canvas, bbox, upscale, useWhiteout) {
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
  if (useWhiteout) {
    stretchContrast(outCtx, out.width, out.height);
    whiteOutGridLines(outCtx, out.width, out.height);
  }

  const result = await worker.recognize(out);

  // The wideLeft crop variant deliberately reaches further left than the
  // detected row to recover a digit clipped at the box's own left edge —
  // but that same reach can also pull in a NEARBY unrelated line of small
  // print (an instruction referencing "line 1", a footnote number) that
  // happens to sit just outside the box. A stray digit from that text
  // isn't distinguishable from a real one once every digit in the crop's
  // text gets pooled together — verified against a case where exactly
  // this happened: one stray "1" from nearby text plus the box's own last
  // digit landing just outside this crop's right edge together added up
  // to a wrong-but-plausible 9-digit string. The box's own digits are
  // printed contiguously (with only dashes between them), so picking the
  // longest [\d-]+ run instead of pooling every digit in the whole crop
  // keeps a same-line stray digit from being mistaken for a dropped one.
  const runs = (result.data.text || '').match(/[\d-]+/g) || [];
  const longestRun = runs.reduce((best, r) => (r.replace(/-/g, '').length > best.replace(/-/g, '').length ? r : best), '');
  const digits = longestRun.replace(/\D/g, '');
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
async function ocrTaxIdDigitsAtScale(worker, canvas, bbox, expectedLength, upscale, useWhiteout) {
  const rowWidth = bbox.x1 - bbox.x0;
  const wideLeft = Math.max(0, bbox.x0 - rowWidth);
  const wideRight = bbox.x1 + rowWidth * 0.5;
  const [fromWideLeft, fromDetected] = await Promise.all([
    ocrDigitRowAtBbox(worker, canvas, { x0: wideLeft, x1: wideRight, y0: bbox.y0, y1: bbox.y1 }, upscale, useWhiteout),
    ocrDigitRowAtBbox(worker, canvas, { x0: bbox.x0, x1: wideRight, y0: bbox.y0, y1: bbox.y1 }, upscale, useWhiteout),
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

// Returns { value, candidate }. `value` is the same three-state signal as
// before (digit string / '' not-found / null ambiguous) built from this
// bbox's own 3 scale attempts alone. `candidate` additionally surfaces the
// single best-supported reading even when it didn't reach its own 2-of-3
// majority, so a caller can cross-check it against an independently
// sourced reading (e.g. the fixed-region fallback crop's own attempt) —
// corroboration this one tier's internal vote count couldn't provide by
// itself. Verified against a case where the dynamic-anchor crop and the
// fixed-region crop each independently landed on the same correct 9-digit
// value exactly once (1 of their own 3 scales) — neither alone reached a
// majority, so without this the box was left unconfirmed even though two
// differently-sourced crops agreed with each other.
async function ocrTaxIdDigits(worker, canvas, bbox, expectedLength) {
  if (!bbox) return { value: '', candidate: '' };
  const density = regionInkDensity(canvas, bbox);
  if (density < MIN_DIGIT_INK_DENSITY) return { value: '', candidate: '' };

  const whiteoutResults = await Promise.all(
    DIGIT_OCR_SCALES.map((scale) => ocrTaxIdDigitsAtScale(worker, canvas, bbox, expectedLength, scale, true))
  );

  // Whiting out the box's own grid lines (see whiteOutGridLines) fixes
  // real cases where a border stroke was fusing into a digit glyph and
  // dropping it entirely — but the same operation can, on a different
  // scan, nibble into a digit's OWN stroke near the border closely enough
  // to bias its shape (an open "6" reading as a closed "8", verified
  // against a real case; a similarly-shaped confusion isn't hard to
  // imagine for other digit pairs). Re-running the identical 3-scale vote
  // with whiteout OFF and pooling both sets into one majority catches
  // that: a whiteout-only artifact shows up as a 3-3 split against the
  // untouched raw reading rather than a clean majority, and pairwise ties
  // are treated as ambiguous below rather than arbitrarily preferring
  // whichever group happened to run first.
  //
  // This always runs, even when the whiteout group already agrees with
  // itself unanimously — a tempting-looking shortcut that was tried and
  // is actively dangerous: whiteout being wrong AND perfectly consistent
  // about it (all 3 scales biasing the same digit's shape the same way)
  // is exactly the failure mode this second pass exists to catch. "It
  // already looks confident" is not evidence it's uncontaminated.
  const rawResults = await Promise.all(
    DIGIT_OCR_SCALES.map((scale) => ocrTaxIdDigitsAtScale(worker, canvas, bbox, expectedLength, scale, false))
  );
  const results = [...whiteoutResults, ...rawResults];
  const exactLength = results.filter((d) => d.length === expectedLength);
  // A one-digit-too-long reading is usually a real box digit plus exactly
  // one stray extra character — but WHICH end it lands on varies: verified
  // cases exist of it landing at the front (a nearby digit bleeding in
  // before the box) and in the middle (a border/dash fragment misread as a
  // digit between two real ones), so blindly trimming a fixed end (like
  // always dropping the first digit) fixes one shape and silently breaks
  // the other. Only trimming a length+1 reading down when doing so lands
  // on a value some OTHER scale already read cleanly at the right length
  // avoids guessing which end is the stray one — it only ever "rescues" a
  // noisy reading when a genuinely independent scale already corroborates
  // the trimmed result.
  // Every candidate value is tagged with which preprocessing group(s)
  // produced it, so a later tie can be broken by an actual signal (see
  // below) instead of an arbitrary pick.
  const valid = []; // { value, source: 'whiteout' | 'raw' }
  whiteoutResults.forEach((d) => {
    if (d.length === expectedLength) valid.push({ value: d, source: 'whiteout' });
  });
  rawResults.forEach((d) => {
    if (d.length === expectedLength) valid.push({ value: d, source: 'raw' });
  });

  const longerWithSource = [
    ...whiteoutResults.filter((d) => d.length === expectedLength + 1).map((d) => ({ value: d, source: 'whiteout' })),
    ...rawResults.filter((d) => d.length === expectedLength + 1).map((d) => ({ value: d, source: 'raw' })),
  ];
  const longerCounts = new Map();
  longerWithSource.forEach(({ value }) => longerCounts.set(value, (longerCounts.get(value) || 0) + 1));
  for (const [d] of longerCounts.entries()) {
    const dropFirst = d.slice(1);
    const dropLast = d.slice(0, -1);
    const matches = longerWithSource.filter((r) => r.value === d);
    const frontWitnessed = exactLength.includes(dropFirst);
    const backWitnessed = exactLength.includes(dropLast);
    if (frontWitnessed && !backWitnessed) {
      matches.forEach(({ source }) => valid.push({ value: dropFirst, source }));
    } else if (backWitnessed && !frontWitnessed) {
      matches.forEach(({ source }) => valid.push({ value: dropLast, source }));
    } else if (!frontWitnessed && !backWitnessed && matches.length >= 2) {
      // No independent exact-length scale witnessed either trim — but this
      // same length+1 value showed up on its own more than once (e.g. all
      // 3 untouched-raw scales agreeing on it), which is itself a real
      // signal something's there, just not which end is the stray digit.
      // Adding BOTH trims as competing candidates, not picking one, means
      // a genuine front-vs-back ambiguity becomes a tie (safe/ambiguous)
      // below rather than an arbitrary guess — while a case where one
      // direction is ALSO independently reinforced by another group (the
      // whiteout-processed scales, say) still lets that direction win on
      // the numbers, or lets the raw-preference tie-break below decide it.
      matches.forEach(({ source }) => {
        valid.push({ value: dropFirst, source, siblingOf: d, isFrontTrim: true });
        valid.push({ value: dropLast, source, siblingOf: d, isFrontTrim: false });
      });
    }
  }
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
    return { value: nonEmpty.length >= 2 ? null : '', candidate: '' };
  }

  const counts = new Map();
  const sources = new Map(); // value -> Set of sources that voted for it
  const siblingInfo = new Map(); // value -> { siblingOf, isFrontTrim } (only for front/back trim candidates)
  valid.forEach(({ value, source, siblingOf, isFrontTrim }) => {
    counts.set(value, (counts.get(value) || 0) + 1);
    if (!sources.has(value)) sources.set(value, new Set());
    sources.get(value).add(source);
    if (siblingOf !== undefined) siblingInfo.set(value, { siblingOf, isFrontTrim });
  });
  const ranked = [...counts.entries()].sort((a, b) => b[1] - a[1]);
  const [bestValue, bestCount] = ranked[0];
  const runnerUpCount = ranked[1] ? ranked[1][1] : 0;

  // With 6 total votes (3 scales × {whiteout, raw}), a tie between the top
  // two values is exactly the signature of the whiteout preprocessing
  // itself being the source of disagreement — e.g. 3 whiteout-processed
  // scales agreeing on one digit shape against 3 untouched-raw scales
  // agreeing on a different one (verified against a real case where
  // whiteout biased an open "6" into reading as a closed "8"). When
  // that's exactly the shape of the tie — the top value came only from
  // whiteout-processed scales, the runner-up only from untouched-raw ones
  // — it's resolved in favor of raw rather than left ambiguous: whiteout
  // is a heuristic patch layered on top of the actual pixels, and it's
  // been measured distorting a real digit's shape into a different valid
  // digit; raw has no equivalent failure mode measured against it, since
  // its only known weakness is losing information (reading empty or a
  // wrong length) rather than confidently misreading a clean digit as a
  // different one. A tie that ISN'T cleanly whiteout-only vs raw-only
  // (either side mixes sources, or there's a genuine 3-way split) still
  // falls through to ambiguous below — this only ever resolves the one
  // specific failure signature it was measured against.
  let isTie = runnerUpCount > 0 && runnerUpCount === bestCount;
  let winner = bestValue;
  if (isTie) {
    const runnerUpValue = ranked[1][0];
    const bestSources = sources.get(bestValue);
    const runnerUpSources = sources.get(runnerUpValue);
    const bestIsWhiteoutOnly = bestSources.size === 1 && bestSources.has('whiteout');
    const runnerUpIsRawOnly = runnerUpSources.size === 1 && runnerUpSources.has('raw');
    if (bestIsWhiteoutOnly && runnerUpIsRawOnly) {
      winner = runnerUpValue;
      isTie = false;
    } else {
      const runnerUpIsWhiteoutOnly = runnerUpSources.size === 1 && runnerUpSources.has('whiteout');
      const bestIsRawOnly = bestSources.size === 1 && bestSources.has('raw');
      if (runnerUpIsWhiteoutOnly && bestIsRawOnly) {
        isTie = false; // bestValue (raw) already wins as-is
      } else {
        // The other resolvable shape: the tie is between the front-trim
        // and back-trim of the very SAME self-corroborated length+1
        // reading (no whiteout involvement either way — e.g. whiteout
        // found nothing usable at all). The crop geometry itself makes
        // one direction more likely than the other: the wideLeft variant
        // reaches a full row-width further left to recover a digit
        // clipped at the box's own left edge, while the right side is
        // only extended by half a row-width — so a stray character
        // bleeding in from OUTSIDE the box is structurally more likely to
        // land at the front of the reading than the back. Verified
        // against three independent cases, all with the stray digit at
        // the front and none with it at the back.
        const bestSib = siblingInfo.get(bestValue);
        const runnerUpSib = siblingInfo.get(runnerUpValue);
        if (bestSib && runnerUpSib && bestSib.siblingOf === runnerUpSib.siblingOf) {
          winner = bestSib.isFrontTrim ? bestValue : runnerUpValue;
          isTie = false;
        }
      }
    }
  }

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
  return { value: bestCount >= 3 && !isTie ? winner : null, candidate: bestValue };
}


async function tinDigitsFromCrop(cropCanvas) {
  if (!cropCanvas) return { ssn: '', ein: '' };
  let lines = await ocrLinesWithBoxes(cropCanvas);
  let ssnBbox =
    findDigitRowBelow(lines, loose('Social security number')) || findDigitRowByPosition(lines, SSN_ROW_Y_BAND);
  let einBbox =
    findDigitRowBelow(lines, loose('Employer identification number')) ||
    findDigitRowByPosition(lines, EIN_ROW_Y_BAND);

  // Default page segmentation (PSM 3) analyzes the crop's overall layout
  // before deciding what counts as a text line — and a bordered digit-box
  // table sitting right next to a paragraph of prose is exactly the kind of
  // mixed layout that mode can misjudge, silently swallowing the bold
  // "Social security number"/"Employer identification number" header text
  // above the table into whatever it decided the table region was (verified
  // against a real crop where that header text never appeared in the OCR'd
  // lines at all despite being clearly legible). Sparse-text mode (PSM 11)
  // recovers that — but only worth reaching for when the default mode found
  // NEITHER label anywhere: PSM 11 skips layout analysis entirely, which can
  // just as easily fragment an otherwise-clean wide digit row into pieces too
  // narrow to pass findDigitRowBelow's width check on a page where the
  // default mode was already working fine (verified as a real regression on
  // a different document). Trying it only as a fallback, not unconditionally,
  // keeps each mode's failure case from undoing the other's success.
  if (!ssnBbox && !einBbox) {
    const sparseLines = await ocrLinesWithBoxes(cropCanvas, '11');
    const sparseSsnBbox =
      findDigitRowBelow(sparseLines, loose('Social security number')) ||
      findDigitRowByPosition(sparseLines, SSN_ROW_Y_BAND);
    const sparseEinBbox =
      findDigitRowBelow(sparseLines, loose('Employer identification number')) ||
      findDigitRowByPosition(sparseLines, EIN_ROW_Y_BAND);
    if (sparseSsnBbox || sparseEinBbox) {
      lines = sparseLines;
      ssnBbox = sparseSsnBbox;
      einBbox = sparseEinBbox;
    }
  }

  // One worker for every digit-OCR attempt this crop needs (up to a couple
  // dozen, across SSN/EIN × 3 scales × 2 crop variants) instead of one per
  // attempt — see the note on ocrDigitRowAtBbox.
  const digitWorker = await Tesseract.createWorker('eng', undefined, TESSERACT_VENDORED_PATHS);
  await digitWorker.setParameters({
    tessedit_char_whitelist: '0123456789-',
    tessedit_pageseg_mode: '6',
  });
  const [ssnResult, einResult] = await Promise.all([
    ocrTaxIdDigits(digitWorker, cropCanvas, ssnBbox, 9),
    ocrTaxIdDigits(digitWorker, cropCanvas, einBbox, 9),
  ]);
  await digitWorker.terminate();

  return { ssn: ssnResult.value, ein: einResult.value, ssnCandidate: ssnResult.candidate, einCandidate: einResult.candidate };
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
      } else if (
        dynamic &&
        primary.ssnCandidate &&
        primary.ssnCandidate.length === 9 &&
        primary.ssnCandidate === fallback.ssnCandidate
      ) {
        // Neither tier's own 3 scales reached a 2-of-3 majority on its
        // own, but the dynamic-anchor crop and the fixed-region crop are
        // two independently sourced images of the same box — if their
        // single best (otherwise-unconfirmed) guesses agree with each
        // other, that's real corroboration a single tier's internal vote
        // count couldn't provide by itself.
        ssn = primary.ssnCandidate;
        source = 'dynamic-anchor+fixed-region-cross-confirmed';
      } else if (ssn === null || fallback.ssn === null) {
        ssn = null;
      }
    }
    if ((ein || '').length !== 9) {
      if ((fallback.ein || '').length === 9) {
        ein = fallback.ein;
        source = dynamic ? 'dynamic-anchor+fixed-region' : 'fixed-region';
      } else if (
        dynamic &&
        primary.einCandidate &&
        primary.einCandidate.length === 9 &&
        primary.einCandidate === fallback.einCandidate
      ) {
        ein = primary.einCandidate;
        source = 'dynamic-anchor+fixed-region-cross-confirmed';
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
  const loadedCanvas = await imageFileToCanvas(file);
  const { canvas: orientedCanvas, degrees: rotationDegrees, confidence: orientationConfidence } =
    await pickBestOrientation(loadedCanvas);
  const canvas = deskewCanvas(orientedCanvas);
  const ocrCanvas = grayscaleContrastCanvas(canvas);
  if (debugExtractionTrace) {
    debugExtractionTrace.canvas = canvas;
    debugExtractionTrace.rotationDegrees = rotationDegrees;
    debugExtractionTrace.orientationConfidence = orientationConfidence;
  }

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
  let rotationDegrees = 0;
  let orientationConfidence = null;

  for (let i = 1; i <= pagesToScan; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 3 });
    const rawCanvas = document.createElement('canvas');
    rawCanvas.width = viewport.width;
    rawCanvas.height = viewport.height;
    await page.render({ canvasContext: rawCanvas.getContext('2d'), viewport }).promise;

    // Orientation is only probed on page 1 and reused for the rest — a
    // scanned multi-page batch is fed through the scanner the same way
    // every page, so a second/third probe would just repeat the same OCR
    // cost for the same answer.
    let orientedCanvas = rawCanvas;
    if (i === 1) {
      const oriented = await pickBestOrientation(rawCanvas);
      orientedCanvas = oriented.canvas;
      rotationDegrees = oriented.degrees;
      orientationConfidence = oriented.confidence;
    } else if (rotationDegrees) {
      orientedCanvas = rotateCanvasExpand(rawCanvas, rotationDegrees);
    }

    const canvas = deskewCanvas(orientedCanvas);
    const ocrCanvas = grayscaleContrastCanvas(canvas);

    const { text, lines } = await ocrPageWithLines(
      ocrCanvas,
      onProgress ? (m) => onProgress(i, pagesToScan, m) : undefined
    );
    combined += text + '\n';

    if (i === 1) {
      if (debugExtractionTrace) {
        debugExtractionTrace.canvas = canvas;
        debugExtractionTrace.rotationDegrees = rotationDegrees;
        debugExtractionTrace.orientationConfidence = orientationConfidence;
      }
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
  // The digit also needs a negative lookbehind for another digit right
  // before it: without one, this "single bare digit" shape happily matches
  // just the LAST digit of a multi-digit suite number sitting right before
  // this label ("Suite 240" + "City..." lets "0 " alone satisfy the
  // group), truncating the real suite number out of the street text same
  // as the unit-number bug this was written to avoid in the first place.
  const CITY_ZIP_LABEL_RE = new RegExp(
    '(?:(?<!\\d)\\d\\s+)?(?:city\\s*[,.:]?\\s*)?' + STATE_ZIP_CODE_PATTERN + '\\.?',
    'i'
  );

  // Each word required to be Title Case — capital first letter, at least
  // one lowercase letter after — rather than any-case-after-the-first:
  // real city names are always Title Case, and requiring that rejects a
  // run of OCR noise sitting between the label and the real city. A
  // lowercase-led fragment like "I se CR" gets rejected by the
  // leading-capital rule, but an ALL-CAPS noise token like "CR" needs the
  // *-vs-+ distinction to actually get rejected: with `*` (zero or more
  // lowercase), each of "C" and "R" alone still satisfies "one capital
  // plus zero-or-more lowercase", so the two letters can each anchor their
  // own one-letter "word" and chain together as if they were a real
  // two-word city name. Requiring at least one lowercase letter (`+`, not
  // `*`) after the capital blocks that, since no real single-letter word
  // can start a city name here.
  const CITY_STATE_ZIP_RE = /((?:[A-Z][a-z'.-]+\s*)+?),?\s+([A-Za-z]{2})\s*(\d{5}(?:-\d{4})?)/;

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

// Real two-letter US state/territory codes. The state, in both
// extractAddress and parseAddressBoxText, is matched purely by shape
// ("two letters"), so OCR noise that happens to land on two adjacent
// letters (e.g. a stray "IX" or "CR" picked up near the ZIP) can pass
// through as if it were a genuine state. Gating the final value against
// the real list catches that case without touching how the value was
// found in the first place.
const VALID_US_STATES = new Set([
  'AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA', 'HI', 'ID', 'IL', 'IN', 'IA',
  'KS', 'KY', 'LA', 'ME', 'MD', 'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ',
  'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC', 'SD', 'TN', 'TX', 'UT', 'VT',
  'VA', 'WA', 'WV', 'WI', 'WY', 'DC',
  // US territories and military/diplomatic mail codes — a legitimate
  // address here (Puerto Rico, Guam, an APO/FPO address) is otherwise
  // indistinguishable, by this same two-letter shape, from OCR noise.
  'PR', 'GU', 'VI', 'AS', 'MP', 'AA', 'AE', 'AP',
]);

// A captured "state ZIP" string is only trustworthy if its state half is a
// real US state/territory code; otherwise the whole value is dropped
// rather than shown to the user as if it were reliable.
function validateStateZip(stateZip) {
  const match = /^([A-Za-z]{2})\s+(\d{5}(?:-\d{4})?)$/.exec((stateZip || '').trim());
  if (!match) return stateZip;
  return VALID_US_STATES.has(match[1].toUpperCase()) ? stateZip : '';
}

// Federal tax classification (the "W-9 type" checkbox row). The reference
// Python pipeline detects this from a checkbox-region crop with OCR-glyph
// heuristics for the checkmark itself; this pipeline has no such region, so
// rather than translate that line-by-line, this looks for a checkmark-like
// glyph immediately preceding each classification's label anywhere in the
// extracted text, falling back to the tax-ID type the same way the Python
// version does when no checkbox reads cleanly.
const LLC_TAX_CLASSIFICATION_RE = /\b([CSP])\s+LLC,\s*unless\s*it\s*is\s*a\s*disregarded\s*entity/i;
// Checked when the classification-code regex above doesn't match — e.g. the
// filer's handwritten C/S/P code didn't land next to that exact sentence in
// the OCR'd text — so a plain "[x] Limited liability company" mark still
// resolves to LLC instead of silently falling through to the SSN/EIN
// tax-ID-based guess (which would otherwise mislabel an EIN-filing LLC as
// the generic 'Other').
const LLC_CHECKBOX_RE = /[xX✓✔]\s*Limited\s*liability\s*compan/i;
const W9_TYPE_CHECKBOX_LABELS = [
  ['Individual/sole proprietor', 'Individual/sole proprietor'],
  ['C Corporation', 'C corporation'],
  ['S Corporation', 'S corporation'],
  ['Partnership', 'Partnership'],
  ['Trust/estate', 'Trust/estate'],
  ['Other', 'Other'],
];

function extractW9Type(text, tinType) {
  if (LLC_TAX_CLASSIFICATION_RE.test(text)) return 'LLC';
  if (LLC_CHECKBOX_RE.test(text)) return 'LLC';

  for (const [label, result] of W9_TYPE_CHECKBOX_LABELS) {
    if (new RegExp('[xX✓✔]\\s*' + loosePattern(label), 'i').test(text)) return result;
  }

  if (tinType === 'SSN') return 'Individual/sole proprietor';
  if (tinType === 'EIN') return 'Other';
  return '';
}

// Noise phrases that occasionally get grabbed as the company name when the
// real value is missing or unreadable and a label-adjacent fragment slips
// through the regex instead.
const NAME_NOISE_TERMS = [
  'name of entity',
  'business name',
  'requester',
  'print or type',
  'taxpayer identification',
  'part i',
  'part ii',
  'signature',
];

// Trims a trailing run of tokens that contain no letters or digits at all
// (a lone "©", a stray "[" — OCR noise from a nearby checkbox glyph or form
// artifact bleeding into the capture window when the real stop-label after
// the name didn't OCR cleanly enough for the stop regex to catch). Stops at
// the first trailing token that has real alphanumeric content, so a
// legitimate trailing "." or "&" — as in "Smith & Sons, Inc." — is left
// alone; only tokens with NOTHING but symbols get dropped.
function stripTrailingSymbolNoise(value) {
  const tokens = (value || '').split(/\s+/).filter(Boolean);
  while (tokens.length > 1 && !/[A-Za-z0-9]/.test(tokens[tokens.length - 1])) {
    tokens.pop();
  }
  return tokens.join(' ');
}

function isBadCompanyName(value) {
  const name = (value || '').replace(/\s+/g, ' ').trim();
  if (!name) return true;

  const normalized = name.toLowerCase().replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();
  // A threshold of 2 (not 3) so real short names — "3M", "GE", "GM" — survive;
  // only a single stray character/digit (near-certainly OCR noise) is caught.
  if (normalized.length < 2) return true;
  if ((name.replace(/[^A-Za-z]/g, '')).length < 2) return true;
  // Equality, not substring: NAME_NOISE_TERMS exists to catch a captured
  // value that IS leftover label text (e.g. just "Signature", just
  // "Requester"), not to reject a real company name that merely contains
  // one of these words — "Signature Bank" is a real company name.
  if (NAME_NOISE_TERMS.some((term) => normalized === term)) return true;
  if (/^[0-9 ]+$/.test(normalized)) return true;

  return false;
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
  let companyName =
    (directFields && directFields.company_name) || stripTrailingSymbolNoise(extractBusinessName(text));
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

  // Final validation pass: reject values that were captured but aren't
  // plausible, rather than passing OCR/regex noise through to the user.
  // Acroform-sourced values (direct field extraction, not OCR/regex) skip
  // this — they come straight from the PDF's own form fields.
  if (!(directFields && directFields.company_name) && isBadCompanyName(companyName)) {
    companyName = '';
  }
  if (addressSource !== 'acroform') {
    stateZip = validateStateZip(stateZip);
  }
  const entityType = extractW9Type(text, selectedType);

  return {
    company_name: companyName,
    selected_tax_id_type: selectedType,
    tax_id_number: taxId,
    ssn_number: ssnNumber,
    ein_number: einNumber,
    street_address: streetAddress,
    city,
    state_zip: stateZip,
    entity_type: entityType,
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
  console.log('entity_type:', data.entity_type || 'none');
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

// Live OCR progress ("Running OCR... 45%") already updates once per tick
// inside the wizard's own loading overlay (wizardLoadingText, below) — the
// hero's #status line mirroring the exact same fast-ticking text right next
// to it was pure visual duplication, not two different pieces of
// information. Suppressing just this one repeating message pattern from
// #status keeps every other message (the initial "Reading file..." line,
// and the final parsed/error result once extraction finishes) exactly as
// before.
const OCR_PROGRESS_TICK_RE = /^Running OCR(?: on page \d+\/\d+)?\.\.\. \d+%$/;

function setStatus(message, isError = false) {
  if (!OCR_PROGRESS_TICK_RE.test(message)) {
    status.textContent = message;
    status.classList.toggle('error', isError);

    status.classList.remove('flash');
    void status.offsetWidth;
    status.classList.add('flash');
  }

  if (!wizardLoading.hidden) {
    wizardLoadingText.textContent = message;
  }
}

// Messages about generating/finalizing the MSA document surface next to the
// document preview (right side) instead of the left-side status line, since
// that's the panel they actually describe. An error here also opens the
// preview panel — it's easy to miss otherwise, since the panel is closed by
// default and this message can be the only explanation for why the wizard
// didn't advance.
function setPreviewStatus(message, isError = false) {
  previewStatus.textContent = message;
  previewStatus.classList.toggle('error', isError);

  previewStatus.classList.remove('flash');
  void previewStatus.offsetWidth;
  previewStatus.classList.add('flash');

  if (isError) {
    setPreviewPanelOpen(true);
    setActivePreviewTab('generated');
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

// This template's own w:pgSz is 12240x15840 twips — US Letter (8.5in x 11in),
// not A4. Targeting A4 (595.28x841.89pt, narrower AND taller than Letter) forced
// every page to scale down ~2.7% in width to fit, which then left a stray blank
// margin at the bottom of every page (A4's extra height beyond the actual
// 792pt-tall content) and shifted where the content-aware slicer below found a
// "full page" — visible as inconsistent extra whitespace, worst on whichever
// page's real content happened to land closest to that leftover margin. Letter
// at 72pt/in matches the source exactly, so a Word page (already paginated by
// docx-preview to its own true size) maps to one output page with no rescaling
// and, in the ordinary case, no slicing at all — output pagination then matches
// Word's own page breaks one-to-one instead of drifting from them.
const LETTER_WIDTH_PT = 612;
const LETTER_HEIGHT_PT = 792;

// The header positions the logo with a paragraph tab (w:ptab, "jump to the right
// margin") that docx-preview silently drops, so the logo renders at the left instead
// of the right — Word itself (and the .docx download) shows it top-right.
function alignHeaderLogoRight(root) {
  root.querySelectorAll('.docx_header').forEach((p) => {
    p.style.textAlign = 'right';
  });
}

// Every run in this template's document.xml explicitly sets w:rFonts to Times New
// Roman — there is no other font used anywhere in the body. docx-preview still
// renders body text in Calibri (the Normal style's docDefaults font) for most
// runs regardless, only partially honoring the run-level override for a handful
// of runs (e.g. bold Fed-Id values) and even those come out as a font stack
// ("Arial, Times New Roman" on the title) rather than the plain single font Word
// itself shows. Since the template never legitimately uses any other font, it's
// safe to force Times New Roman as an inline style on every element docx-preview
// rendered — an inline style already outranks docx-preview's class-based rules
// without needing !important, which is worth avoiding here: html2canvas parses
// each element's inline style text itself rather than only using the resolved
// computed value, and a literal "!important" suffix in that text confused its
// font-matching enough to silently fall back to a generic sans-serif font.
function forceDocumentFont(root) {
  root.querySelectorAll('section.docx, section.docx *').forEach((el) => {
    el.style.fontFamily = '"Times New Roman", Times, serif';
  });
}

// docx-preview computes every custom Word tab stop's width (tabStopClass, sized
// by updateTabStop/refreshTabStops) from getBoundingClientRect() measurements
// taken while the rendered nodes are still detached from the document —
// refreshTabStops runs at the end of its own render() pass, and that pass
// returns before the caller (renderAsync) ever appends the result into a
// container. Every measurement it reads back is therefore zero, so every
// custom tab in this template renders far narrower than it should: the
// manually-tabbed "(b) Employer's..." clause on page 3 (every other lettered
// item gets its indent from numbering.xml instead, unaffected), and the
// "SBS CORP <tab> [Contractor]" / "Name:"/"Title:"/"Date:" two-column tab
// stops on both signature pages. Once these nodes are actually attached (they
// are by the time this runs), redoing docx-preview's own computation — this
// time against real layout — fixes it without hardcoding pixel guesses that
// would break if the fixed text before a tab (e.g. "SBS CORP ") or the
// template's font ever changed.
const PX_TO_PT = 72 / 96;

function fixTabStopWidth(tabSpan, targetPt) {
  const p = tabSpan.closest('p');
  if (!p) return;
  const pRect = p.getBoundingClientRect();
  const spanRect = tabSpan.getBoundingClientRect();
  const marginLeft = parseFloat(getComputedStyle(p).marginLeft) || 0;
  const leftPt = (spanRect.left - pRect.left - marginLeft) * PX_TO_PT;
  const widthPt = Math.max(0, targetPt - leftPt);
  tabSpan.style.wordSpacing = `${widthPt.toFixed(0)}pt`;
}

// Word's own default tab stop increment (used by any paragraph that doesn't
// define its own custom w:tabs, like the small "SBS <tab> Contractor" label
// under each page's Initial block) — a half inch.
const DEFAULT_TAB_STOP_PT = 36;

// pt targets read straight from this template's own w:tabs (twips / 20): the
// "(b)" clause's own first stop (1440 twips = 72pt); its continuation line
// ("covering bodily injury...", a separate paragraph docx-preview renders as
// one visual line via matching indentation) walks the same four stops in
// sequence (1440/1530/1620/1800 twips = 72/76.5/81/90pt) to reach its final
// 90pt indent; and each signature block's company-name column (4860 twips =
// 243pt) and Name:/Title:/Date: column (5490 twips = 274.5pt). The matching
// "Sign:" line uses literal space characters instead of a tab in the source
// document, so it isn't affected and isn't touched here.
function fixTemplateTabStops(root) {
  root.querySelectorAll('section.docx p').forEach((p) => {
    const tabSpans = Array.from(p.querySelectorAll('.docx-tab-stop'));
    if (!tabSpans.length) return;
    const text = p.textContent.trim();
    if (text.startsWith('(b) Employer')) {
      fixTabStopWidth(tabSpans[0], 72);
    } else if (text.startsWith('covering bodily injury')) {
      const targets = [72, 76.5, 81, 90];
      tabSpans.forEach((span, i) => fixTabStopWidth(span, targets[Math.min(i, targets.length - 1)]));
    } else if (text.startsWith('SBS CORP')) {
      fixTabStopWidth(tabSpans[0], 243);
    } else if (text.startsWith('Name:') || text.startsWith('Title:') || text.startsWith('Date:')) {
      fixTabStopWidth(tabSpans[0], 274.5);
    } else {
      // Every other tab in this template (a handful, all either fully blank
      // or the short "SBS <tab> Contractor" label line) doesn't define its
      // own custom w:tabs, so Word falls back to its own default half-inch
      // stops for it. Replicating that keeps these few unnamed tabs at the
      // small, harmless gap Word shows instead of collapsing to near-zero
      // width now that enabling `experimental` above (to get the named
      // fixes right) hands them to this same, otherwise-broken computation.
      tabSpans.forEach((span) => {
        const pp = span.closest('p');
        if (!pp) return;
        const pRect = pp.getBoundingClientRect();
        const spanRect = span.getBoundingClientRect();
        const marginLeft = parseFloat(getComputedStyle(pp).marginLeft) || 0;
        const leftPt = (spanRect.left - pRect.left - marginLeft) * PX_TO_PT;
        const nextStopPt = Math.ceil((leftPt + 1) / DEFAULT_TAB_STOP_PT) * DEFAULT_TAB_STOP_PT;
        fixTabStopWidth(span, nextStopPt);
      });
    }
  });
}

// Several clauses in this template (the insurance list in "No Power to Act..." /
// Section 3, for one) are authored as several short Word paragraphs in a row
// instead of one paragraph that wraps naturally. Those paragraphs use Word's
// built-in "No Spacing" style, which — in real Word — collapses the gap between
// them to zero even though that's implied by Word's own built-in definition of
// the style rather than spelled out in this template's styles.xml. docx-preview
// only reads what's explicit in styles.xml, so it falls back to the document's
// default paragraph spacing (~10pt) between them instead, visibly gapping out
// what should read as continuous lines of one clause. Forcing the margins to
// zero for this style's class makes the render match what Word actually shows.
function fixNoSpacingParagraphMargins(root) {
  root.querySelectorAll('.docx_nospacing').forEach((p) => {
    p.style.marginTop = '0';
    p.style.marginBottom = '0';
  });
}

// docx-preview can't render a Word feature this template's footer relies on: the
// page-number frame (w:framePr) and its PAGE field, plus the fact the template only
// fills in its "default" footer while leaving the "even page" footer (word/footer1.xml)
// blank/uncached — Word ignores that stray footer because evenAndOddHeaders is off in
// settings.xml, but docx-preview alternates footers by page parity regardless, so
// every other rendered page loses its footer entirely and the pages that keep one show
// a static, unpositioned "1". On top of that, a Word page whose content overflows a
// single sheet gets sliced into several A4 pages below (see renderDocxToPdf) — a
// footer baked into the source screenshot would then land on only whichever slice
// happens to contain those pixels, missing from the rest and never renumbered.
// So: read the address out of whichever docx-preview page actually kept a footer,
// and let the caller draw a real, correctly-numbered footer on every output PDF
// page instead, painted over the source footer (see renderDocxToPdf) rather than
// hiding it here. Mutating the footer's own style (visibility:hidden, opacity:0,
// display:none all reproduce it) before the html2canvas capture below reliably
// makes html2canvas mismeasure this template's justified body text elsewhere on
// the same page and split words mid-letter ("technic" / "al support"), even though
// the footer has nothing to do with that text — so this function only reads text
// now, it doesn't touch the DOM.
function extractFooterAddressText(root) {
  const pageEls = Array.from(root.querySelectorAll('section.docx'));
  let addressText = '';
  for (const pageEl of pageEls) {
    const footer = pageEl.querySelector('footer');
    if (!footer || !footer.textContent.trim()) continue;
    const clone = footer.cloneNode(true);
    const cloneNumberEl = clone.querySelector('.docx_pagenumber');
    if (cloneNumberEl) cloneNumberEl.remove();
    addressText = clone.textContent.replace(/\s+/g, ' ').trim();
    if (addressText) break;
  }
  return addressText;
}

// Blanking out the footer's own text is what actually keeps the source's stale
// address/page-number from ever showing up underneath the fresh one drawn in
// renderDocxToPdf — a size-based "paint a white rectangle over roughly where the
// footer should be" cover can't tell a stale footer apart from a legitimately
// full page whose own last line (an "Initial" line right against the bottom
// margin, say) happens to sit in the same band, and got both wrong on different
// pages (covering real text on one, missing the stale footer on another).
// Clearing individual text nodes (not the footer's own visibility/display/
// opacity) is a far smaller DOM change than hiding the whole block, and unlike
// that doesn't reproduce the html2canvas mid-word-split bug elsewhere on the
// page.
function clearFooterText(root) {
  root.querySelectorAll('footer').forEach((footer) => {
    const walker = document.createTreeWalker(footer, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      node.nodeValue = '';
    }
  });
}

// Cosmetic-only equivalent for the in-app preview iframe, which just displays the
// docx-preview DOM directly rather than compositing a PDF — so instead of hiding the
// footer and redrawing it elsewhere, put the address and a correct page number back
// on every page's existing footer, on one line as Word shows them.
function normalizePreviewFooter(root) {
  const pageEls = Array.from(root.querySelectorAll('section.docx'));
  let addressFooterHtml = null;
  for (const pageEl of pageEls) {
    const footer = pageEl.querySelector('footer');
    if (footer && footer.textContent.trim()) {
      addressFooterHtml = footer.innerHTML;
      break;
    }
  }
  if (!addressFooterHtml) return;

  pageEls.forEach((pageEl, i) => {
    const footer = pageEl.querySelector('footer');
    if (!footer) return;
    footer.innerHTML = addressFooterHtml;
    const pageNumberEl = footer.querySelector('.docx_pagenumber');
    if (!pageNumberEl) return;
    pageNumberEl.textContent = String(i + 1);

    footer.style.display = 'flex';
    footer.style.alignItems = 'baseline';
    footer.style.justifyContent = 'space-between';
    const pageNumberPara = pageNumberEl.closest('p');
    if (pageNumberPara) {
      pageNumberPara.style.margin = '0';
      pageNumberPara.style.order = '1';
    }
    const addressPara = Array.from(footer.children).find((el) => el !== pageNumberPara);
    if (addressPara) addressPara.style.margin = '0';
  });
}

// docx-preview's own predefined stylesheet sets `.docx span { overflow-wrap:
// break-word }`. html2canvas re-lays-out text with its own metrics instead of
// reading the browser's real layout, and every so often that measurement is just
// imprecise enough to conclude a word — or even part of one word-span, see
// bakeInLineBreaks below — barely doesn't fit a line when, with the browser's
// actual precise layout, it always did; this rule then legitimately lets it split
// that word mid-letter to avoid an overflow that was never actually going to
// happen. Neutralizing it removes html2canvas's ability to invoke that fallback
// at all — combined with bakeInLineBreaks, which independently removes html2canvas's
// need to decide *where* a line ends, this closes off both paths to a mid-word
// split it might otherwise take.
function withWordBreakNeutralized(styleContainer, fn) {
  const override = document.createElement('style');
  override.textContent = '.docx span { overflow-wrap: normal !important; word-break: normal !important; }';
  styleContainer.appendChild(override);
  return Promise.resolve(fn()).finally(() => {
    override.remove();
  });
}

// When a numbered list paragraph happens to fall right at a Word-page boundary,
// docx-preview splits it into two <p> elements — an empty stub ending the first
// page and the real text starting the next — but gives BOTH the same numbering
// class, so the CSS counter (which drives the "(a)", "(b)", "(c)" markers)
// increments twice for what is really one list item. The next page then shows
// every remaining item off by one letter (see "(e)" for what should read "(d)").
// The stub carries no visible text, so removing it outright removes its
// phantom counter-increment along with it — the real item right after it then
// gets the correct, single increment and the correct letter.
function removeEmptyNumberingStubs(root) {
  root.querySelectorAll('p[class*="docx-num-"]').forEach((p) => {
    if (!p.textContent.trim()) p.remove();
  });
}

// This template's page is 792pt tall (LETTER_HEIGHT_PT). docx-preview's own
// automatic pagination gives each "section.docx" container a *min-height* of
// that, not a hard cap, so when its own page-break estimate runs long it lets
// a container grow past one true page's worth of content instead of cutting
// it off, showing (and, before this, exporting) one too-tall "page" spanning
// what should be several. This clones the oversized page's own shell (keeping
// its header/footer/margins) for each overflow and moves the overflowing
// blocks into it — cutting only between top-level blocks (each <p>/table
// docx-preview renders directly under <article>), never through one — so
// both the live preview and the exported PDF end up with one true US-Letter
// sheet per page, matching Word's own pagination instead of docx-preview's
// fewer, taller pages. Shared by both, so they can never disagree with each
// other on where a page actually breaks.
function repaginateToLetterPages(root) {
  const pageEls = Array.from(root.querySelectorAll('section.docx'));
  if (!pageEls.length) return;

  // A page's own inline style declares its width as exactly 612pt
  // (LETTER_WIDTH_PT) — but docx-preview doesn't reliably render at a clean
  // 96px/in, so assuming that ratio to convert 792pt into a pixel threshold
  // drifts by a few percent depending on however it actually rendered this
  // run. Measuring the page's own rendered pixel width against its known
  // 612pt and deriving the page-height threshold from that ratio keeps this
  // exact, however docx-preview happened to render it — a fixed 96dpi
  // assumption here previously left a handful of pages a little taller than
  // a true page even after every fix above, because their real content did
  // fit in 792pt but not in the dpi-mismatched pixel count standing in for it.
  const pxPerPt = pageEls[0].getBoundingClientRect().width / LETTER_WIDTH_PT;
  const PAGE_HEIGHT_PX = LETTER_HEIGHT_PT * pxPerPt;

  function splitOversizedPagesOnce() {
    let splitAny = false;
    for (let i = 0; i < pageEls.length; i++) {
      const pageEl = pageEls[i];
      const article = pageEl.querySelector('article');
      if (!article) continue;
      const pageTop = pageEl.getBoundingClientRect().top;
      const children = Array.from(article.children);

      // Find the first block that doesn't fully fit — checking each block's
      // own *bottom* against the page height, not its top, matters for a
      // block several lines tall (a long paragraph, say) that starts before
      // the boundary but wraps past it: its top alone would never cross the
      // threshold, so a top-only check finds no split point at all and lets
      // the whole page run long. Moving that block whole to the next page
      // (never split before the first block — that would produce an empty
      // page) can leave this page a little short of a full 792pt, which is
      // the ordinary, expected way a paragraph that doesn't fit is handled.
      let splitIndex = -1;
      for (let c = 1; c < children.length; c++) {
        const bottom = children[c].getBoundingClientRect().bottom - pageTop;
        if (bottom > PAGE_HEIGHT_PX) {
          splitIndex = c;
          break;
        }
      }
      if (splitIndex === -1) continue;

      const newPageEl = pageEl.cloneNode(true);
      const newArticle = newPageEl.querySelector('article');
      while (newArticle.firstChild) newArticle.removeChild(newArticle.firstChild);
      for (let c = splitIndex; c < children.length; c++) {
        newArticle.appendChild(children[c]);
      }
      pageEl.after(newPageEl);
      // Splice the new page in right after this one so the loop reaches it too —
      // a severely oversized container can need more than one extra page.
      pageEls.splice(i + 1, 0, newPageEl);
      splitAny = true;
    }
    return splitAny;
  }

  // A single pass already re-examines any page it just created (the loop
  // above reaches it via the splice), but repeats here too: moving a whole
  // straddling block onto a fresh page can still leave that new page itself
  // over height if more than one such block was needed, or if a later fixup
  // (moveContentAfterTrailingBlankRun, below) prepends onto a page that was
  // already full.
  function splitOversizedPages() {
    for (let round = 0; round < 10 && splitOversizedPagesOnce(); round++);
  }

  // Many sections in this template end with a run of several blank paragraphs
  // right after their own "Initial ___ / SBS Contractor" sign-off line —
  // padding that, in the source document's own original pagination, existed
  // to push whatever comes next onto a new page. docx-preview's pagination
  // doesn't treat that blank run as an actual page break, so with pages now
  // cut at their own true height, whatever trails it — a major section
  // heading like "4. Fees", or just as often a single lettered clause like
  // "(d)" continuing the very same section after its own sign-off line — can
  // still land on the same page, stranded far below everything else with
  // nothing following it. A run of three or more consecutive blank paragraphs
  // is long enough to tell this deliberate padding apart from the ordinary
  // one-line gaps between paragraphs elsewhere (those never reach that many
  // blank paragraphs in a row) — but that same padding also separates every
  // OTHER section from the next one that just happens to follow normally on
  // the same page, with a full section's worth of real content after it, not
  // a small stranded fragment; moving that much unconditionally turned every
  // one of those into a forced page break too (11 pages became 15). Only
  // stepping in when what follows the run is itself small — under a third of
  // a page — keeps this to rescuing genuinely orphaned fragments like the two
  // above, without reshaping pages that were never actually broken. Returns
  // whether anything moved.
  const MIN_BLANK_RUN = 3;
  const MAX_ORPHAN_HEIGHT_PX = PAGE_HEIGHT_PX / 3;
  function moveContentAfterTrailingBlankRun() {
    let movedAny = false;
    for (let i = 0; i < pageEls.length - 1; i++) {
      const article = pageEls[i].querySelector('article');
      const nextArticle = pageEls[i + 1].querySelector('article');
      if (!article || !nextArticle) continue;
      const pageTop = pageEls[i].getBoundingClientRect().top;
      const kids = Array.from(article.children);
      let splitAt = -1;
      let runLen = 0;
      for (let c = 0; c < kids.length; c++) {
        if (!kids[c].textContent.trim()) {
          runLen++;
        } else {
          // This template's own "Initial ___ / SBS Contractor" sign-off is
          // itself commonly preceded by this same blank-padding pattern, but
          // it belongs at the end of whichever page it naturally falls on —
          // it's the padding *after* it, not before it, that signals a real
          // page break. Treating "Initial" as the orphan here moved the
          // sign-off itself onto the next page instead of leaving it in
          // place, so it's excluded from counting as a split point (the run
          // simply keeps being ignored, not reset past it) while scanning
          // continues forward for the genuine split that follows it.
          if (runLen >= MIN_BLANK_RUN && kids[c].textContent.trim() !== 'Initial') {
            splitAt = c;
          }
          runLen = 0;
        }
      }
      if (splitAt <= 0) continue;
      const lastKid = kids[kids.length - 1];
      const trailingHeight = lastKid.getBoundingClientRect().bottom - pageTop - (kids[splitAt].getBoundingClientRect().top - pageTop);
      if (trailingHeight > MAX_ORPHAN_HEIGHT_PX) continue;
      nextArticle.prepend(...kids.slice(splitAt));
      movedAny = true;
    }
    return movedAny;
  }

  splitOversizedPages();
  // Prepending trailing content onto a page that was already nearly full
  // can itself push that page back over true page height, so keep
  // alternating the two passes until a round moves nothing further — bounded,
  // since this template's pages settle within one or two rounds in practice.
  for (let round = 0; round < 5; round++) {
    if (!moveContentAfterTrailingBlankRun()) break;
    splitOversizedPages();
  }
}

// The main MSA's own "IN WITNESS WHEREOF" closing line and the signature
// block right after it ("SBS CORP" / the contractor's name, plus the
// Sign:/Name:/Title:/Date: columns) are one cohesive unit, but this
// template's own blank-padding pattern between them (the same one
// moveContentAfterTrailingBlankRun corrects elsewhere) can still leave them
// split across two pages, since neither piece alone is large enough to look
// like a genuine orphan and there's a whole "Exhibit A" section coming right
// after to fill the rest of that page anyway. This targets only this one
// specific pair of paragraphs by their own text — deliberately narrower than
// a general rule, so it can't reshape pagination anywhere else in the
// document.
function moveMainSignatureBlockToWitnessPage(root) {
  const paragraphs = Array.from(root.querySelectorAll('section.docx p'));
  const witnessIdx = paragraphs.findIndex((p) => p.textContent.includes('IN WITNESS WHEREOF'));
  if (witnessIdx === -1) return;
  const sigStartIdx = paragraphs.findIndex(
    (p, idx) => idx > witnessIdx && p.textContent.trim().startsWith('SBS CORP')
  );
  if (sigStartIdx === -1) return;
  const witnessPage = paragraphs[witnessIdx].closest('section.docx');
  const sigPage = paragraphs[sigStartIdx].closest('section.docx');
  if (!witnessPage || !sigPage || witnessPage === sigPage) return;
  const witnessArticle = witnessPage.querySelector('article');
  if (!witnessArticle) return;
  const exhibitIdx = paragraphs.findIndex(
    (p, idx) => idx > sigStartIdx && p.textContent.trim().startsWith('Exhibit A')
  );
  const endIdx = exhibitIdx === -1 ? paragraphs.length : exhibitIdx;
  paragraphs
    .slice(sigStartIdx, endIdx)
    .filter((p) => p.closest('section.docx') === sigPage)
    .forEach((p) => witnessArticle.appendChild(p));
}

// This template's own per-section "Initial ___ / SBS Contractor" sign-off
// isn't wanted on the two pages that close out the document — the one
// carrying the main "IN WITNESS WHEREOF" signature block, and the Exhibit A
// page — since the signature blocks already on those pages cover that. Which
// earlier section's own Initial line happens to land there instead shifts
// with how long the substituted company name/address/etc. are, so this
// anchors on the two closing pages by their own text (not a fixed page
// number) and removes only an Initial group that ends up sharing a page with
// either — every other Initial block elsewhere in the document is untouched.
function removeInitialFromClosingPages(root) {
  const pageEls = Array.from(root.querySelectorAll('section.docx'));
  pageEls.forEach((pageEl) => {
    const text = pageEl.textContent;
    if (!text.includes('IN WITNESS WHEREOF') && !text.includes('Exhibit A')) return;
    const article = pageEl.querySelector('article');
    if (!article) return;
    Array.from(article.children).forEach((p) => {
      if (p.textContent.trim() !== 'Initial') return;
      // Remove this paragraph and the rest of its fixed group (an empty
      // spacer line, the underscore signature line, then the "SBS
      // Contractor" labels) — stop once the label line itself is removed,
      // since that's always the last piece of this group in the template.
      let cur = p;
      while (cur) {
        const next = cur.nextElementSibling;
        const isLabelLine = cur.textContent.includes('Contractor');
        cur.remove();
        if (isLabelLine) break;
        cur = next;
      }
    });
  });
}

// Moving the main signature block back onto its "IN WITNESS WHEREOF" page
// above can leave the Exhibit A content that used to share a page with it
// reflowing such that just one trailing field ends up alone starting the
// very last page, with everything else back on the page before it. This
// only ever looks at the last page in the whole document, and only acts
// when it holds exactly one real line — narrower than a general rule, so it
// can't affect pagination anywhere earlier in the document.
function mergeSparseLastPageBack(root) {
  const pageEls = Array.from(root.querySelectorAll('section.docx'));
  if (pageEls.length < 2) return;
  const lastPage = pageEls[pageEls.length - 1];
  const prevPage = pageEls[pageEls.length - 2];
  const lastArticle = lastPage.querySelector('article');
  const prevArticle = prevPage.querySelector('article');
  if (!lastArticle || !prevArticle) return;
  const realChildren = Array.from(lastArticle.children).filter((c) => c.textContent.trim());
  if (realChildren.length !== 1) return;
  Array.from(lastArticle.children).forEach((c) => prevArticle.appendChild(c));
  lastPage.remove();
}

// Pages 2 and 4 each end with a clause that continues onto the next page, so
// this template's own recurring "Initial ___ / SBS Contractor" sign-off
// naturally lands wherever that clause actually finishes instead — leaving
// visible blank room at the bottom of these two pages specifically, with
// nothing removed or moved to put it there. Cloning the group from wherever
// it already occurs and appending a copy is purely additive: it only ever
// touches these two page numbers, nothing else in the document.
function addInitialToPages(root, pageNumbers) {
  const pageEls = Array.from(root.querySelectorAll('section.docx'));
  let template = null;
  for (const pageEl of pageEls) {
    const article = pageEl.querySelector('article');
    if (!article) continue;
    const kids = Array.from(article.children);
    const idx = kids.findIndex((k) => k.textContent.trim() === 'Initial');
    if (idx === -1) continue;
    let end = idx;
    while (end < kids.length && !kids[end].textContent.includes('Contractor')) end++;
    if (end < kids.length) {
      template = kids.slice(idx, end + 1);
      break;
    }
  }
  if (!template) return;
  pageNumbers.forEach((pn) => {
    const pageEl = pageEls[pn - 1];
    if (!pageEl) return;
    const article = pageEl.querySelector('article');
    if (!article) return;
    if (Array.from(article.children).some((k) => k.textContent.trim() === 'Initial')) return;
    template.forEach((node) => article.appendChild(node.cloneNode(true)));
  });
}

// html2canvas re-lays-out text with its own metrics instead of reading the
// browser's real layout, and every so often that measurement is just imprecise
// enough to conclude a word barely doesn't fit a line when, with the browser's
// actual precise layout, it always did — splitting that word mid-letter
// ("technic" / "al support") in the exported PDF only, never in the live preview.
// No amount of CSS coaxing (word-break, overflow-wrap, letter-spacing) alone fixes
// this reliably, because it isn't really about any one of those properties — it's
// html2canvas's own line-fitting decision that's wrong. So: read the browser's own
// (always-correct) line breaks directly off the live, already-laid-out DOM by
// wrapping every word in its own inline span, grouping those spans into lines by
// their rendered top offset, and inserting a real <br> at each line boundary.
// Once every line break already exists as a hard break in the DOM, html2canvas
// has no wrapping decision left to make — it only has to draw each line's fixed
// content, which is exactly what the real browser layout already decided.
function bakeInLineBreaks(root) {
  const paragraphs = root.querySelectorAll('p');
  paragraphs.forEach((p) => wrapWordsInSpans(p));

  // Each new .w2c-word span is unstyled, so docx-preview's own bare `span` rule
  // (not this template's actual Times New Roman) wins until forceDocumentFont
  // reaches it — and it has to reach it here, before the measurement pass below,
  // or every word gets measured under the wrong font's metrics and the <br>s
  // this function bakes in land at the wrong offsets (short orphaned words like
  // "and" or "at" left alone on their own line once the real font is drawn).
  forceDocumentFont(root);

  // Read every word's position first, in one pass with no DOM writes in between —
  // inserting a <br> reflows everything after it, so measuring and mutating in
  // the same pass made each break shift the words after it, which read as new
  // (wrong) line starts and put almost every remaining word on its own line.
  paragraphs.forEach((p) => {
    const words = Array.from(p.querySelectorAll(':scope .w2c-word'));
    if (words.length < 2) return;
    const tops = words.map((w) => w.getBoundingClientRect().top);
    let lineTop = tops[0];
    for (let i = 1; i < words.length; i++) {
      // A single word can be split across sibling spans by a mid-word formatting
      // change in the source document (no whitespace between the fragments) —
      // never insert a break there, real or not, or a word could come out broken
      // across two lines with no hyphen.
      if (words[i].dataset.continuation) continue;
      // Neighboring words with different formatting (bold vs. regular, a different
      // run's slightly different line-height) can land a few px apart in top even
      // while genuinely on the same visual line — much less than the full line
      // height a real new line differs by. Too small an epsilon here misreads that
      // jitter as a new line and inserts a break mid-sentence (once even landing a
      // real, plain word like "located" alone at the top of the next line).
      const SAME_LINE_EPSILON_PX = 6;
      if (tops[i] > lineTop + SAME_LINE_EPSILON_PX) {
        words[i].before(document.createElement('br'));
        lineTop = tops[i];
      }
    }
  });
}

function wrapWordsInSpans(p) {
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  const textNodes = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    // Whitespace-only nodes (a run boundary can put a space in its own separately
    // formatted span) must still be collected, even though they're left untouched
    // below — skipping them here meant the continuation tracking never saw them,
    // so a word right after one incorrectly read as a no-space continuation of
    // whatever word came before the space and never got a line-break candidacy of
    // its own ("Emily" glued to the preceding "and" that way, so a real new line
    // there had nowhere to insert a <br> and "Emily" just fell to the next line
    // wherever html2canvas's own — unreliable — reflow happened to put it).
    if (node.nodeValue) textNodes.push(node);
  }
  // Tracks whether the content immediately before the next word-span, anywhere
  // earlier in the paragraph (possibly in a different original run/span), ended
  // in whitespace — false means that next span is a continuation fragment of the
  // same source word, not the start of a new one.
  let lastEndedWithWhitespace = true;
  textNodes.forEach((node) => {
    const frag = document.createDocumentFragment();
    // Keep each run of whitespace as a plain text node (so inter-word spacing,
    // including justify's stretched spaces, renders exactly as before) and wrap
    // only the non-whitespace runs so each word can be measured individually.
    const parts = node.nodeValue.split(/(\s+)/);
    parts.forEach((part) => {
      if (!part) return;
      if (/^\s+$/.test(part)) {
        frag.appendChild(document.createTextNode(part));
        lastEndedWithWhitespace = true;
      } else {
        const span = document.createElement('span');
        span.className = 'w2c-word';
        span.style.overflowWrap = 'normal';
        span.style.wordBreak = 'normal';
        span.style.whiteSpace = 'nowrap';
        span.textContent = part;
        if (!lastEndedWithWhitespace) span.dataset.continuation = 'true';
        frag.appendChild(span);
        lastEndedWithWhitespace = false;
      }
    });
    node.replaceWith(frag);
  });
}

async function renderDocxToPdf(docxBytes) {
  const container = document.createElement('div');
  container.style.position = 'fixed';
  container.style.left = '-10000px';
  container.style.top = '0';
  document.body.appendChild(container);

  // Must NOT be a descendant of container: docx-preview's renderAsync clears the
  // body container's innerHTML before appending rendered nodes into it, which would
  // detach a nested style container (and the numbering/theme <style> rules it holds)
  // from the live document before html2canvas ever sees it — silently dropping all
  // numbered-list markers ("1.", "(a)", etc.) and other stylesheet-driven formatting
  // from the exported PDF even though inline per-run formatting still renders fine.
  const styleContainer = document.createElement('div');
  document.body.appendChild(styleContainer);

  try {
    const blob = new Blob([docxBytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await renderDocxAsync(blob, container, styleContainer, {
      inWrapper: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
      // Without this, docx-preview renders every Word tab character as a
      // fixed small em-space instead of computing where it should actually
      // land — its own tab-stop system (the class fixTemplateTabStops below
      // corrects) only runs when this flag is on.
      experimental: true,
    });

    alignHeaderLogoRight(container);
    fixNoSpacingParagraphMargins(container);
    forceDocumentFont(container);
    fixTemplateTabStops(container);
    const footerAddressText = extractFooterAddressText(container);
    clearFooterText(container);
    removeEmptyNumberingStubs(container);
    // Splits any docx-preview page taller than its own true US-Letter size
    // into properly-sized ones (and relocates an orphaned section heading to
    // the page it introduces) before anything is measured or rasterized below
    // — the same repagination the live preview uses, so the two can never
    // disagree on where a page actually breaks.
    repaginateToLetterPages(container);
    moveMainSignatureBlockToWitnessPage(container);
    removeInitialFromClosingPages(container);
    mergeSparseLastPageBack(container);
    addInitialToPages(container, [2, 4]);
    bakeInLineBreaks(container);

    const pageEls = Array.from(container.querySelectorAll('section.docx'));
    const pdfDoc = await PDFLib.PDFDocument.create();
    const footerFont = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    const FOOTER_SIZE = 8;
    const FOOTER_MARGIN = 36; // matches the template's 0.5in page margin
    const FOOTER_Y = 20;
    let outputPageNumber = 0;

    for (const pageEl of pageEls) {
      const canvas = await withWordBreakNeutralized(styleContainer, () =>
        html2canvas(pageEl, { scale: 2, backgroundColor: '#ffffff' })
      );

      // Each pageEl is now already sized to this template's true US-Letter
      // page, so it maps directly to one output page — no further slicing.
      const scale = LETTER_WIDTH_PT / canvas.width;
      const drawWidth = LETTER_WIDTH_PT;
      const drawHeight = canvas.height * scale;
      const pageHeight = Math.max(LETTER_HEIGHT_PT, drawHeight);

      const pngBlob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
      const pngBytes = await pngBlob.arrayBuffer();

      const image = await pdfDoc.embedPng(pngBytes);
      const page = pdfDoc.addPage([LETTER_WIDTH_PT, pageHeight]);
      const imageBottomY = pageHeight - drawHeight;
      page.drawImage(image, { x: 0, y: imageBottomY, width: drawWidth, height: drawHeight });

      outputPageNumber += 1;
      if (footerAddressText) {
        page.drawText(footerAddressText, {
          x: FOOTER_MARGIN,
          y: FOOTER_Y,
          size: FOOTER_SIZE,
          font: footerFont,
        });
      }
      const pageNumberText = String(outputPageNumber);
      const pageNumberWidth = footerFont.widthOfTextAtSize(pageNumberText, FOOTER_SIZE);
      page.drawText(pageNumberText, {
        x: LETTER_WIDTH_PT - FOOTER_MARGIN - pageNumberWidth,
        y: FOOTER_Y,
        size: FOOTER_SIZE,
        font: footerFont,
      });
    }

    return pdfDoc.save();
  } finally {
    document.body.removeChild(container);
    document.body.removeChild(styleContainer);
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
    setPreviewStatus(
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
    setPreviewStatus(`Could not prepare the MSA template: ${err.message}`, true);
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

function setManualFieldsStatus(message, isError = false) {
  manualFieldsStatus.textContent = message;
  manualFieldsStatus.classList.toggle('error', isError);

  manualFieldsStatus.classList.remove('flash');
  void manualFieldsStatus.offsetWidth;
  manualFieldsStatus.classList.add('flash');
}

// Representative, Role, and Location are free-typed text, so this is where a
// stray special character would actually come from a keystroke — Start Date
// is populated from the date picker in a fixed "Month D, YYYY" format (via
// syncStartDatePickerFromText, which needs its comma) and Billing Rate is
// already a type="number" input the browser itself restricts, so neither
// belongs in this filter.
const NO_SPECIAL_CHARS_FIELDS = [manualRep, manualRole, manualLocation];
const SPECIAL_CHARS_PATTERN = /[^a-zA-Z0-9 ]/g;
NO_SPECIAL_CHARS_FIELDS.forEach((el) => {
  el.addEventListener('input', () => {
    const original = el.value;
    const sanitized = original.replace(SPECIAL_CHARS_PATTERN, '');
    if (sanitized === original) return;
    // Re-deriving the caret from how much of the text *before* it survived
    // sanitizing (rather than just shifting back by one) keeps this correct
    // when a paste strips several characters at once, not only a single
    // disallowed keystroke.
    const caret = el.selectionStart ?? original.length;
    const newCaret = original.slice(0, caret).replace(SPECIAL_CHARS_PATTERN, '').length;
    el.value = sanitized;
    el.setSelectionRange(newCaret, newCaret);
  });
});

MANUAL_FIELD_INPUTS.forEach((el) => {
  el.addEventListener('input', () => {
    if (el.value.trim()) {
      el.classList.remove('field-error');
      if (MANUAL_FIELD_INPUTS.every((f) => f.value.trim())) setManualFieldsStatus('');
    }
    scheduleLivePreviewUpdate();
  });
});

applyManualBtn.addEventListener('click', async () => {
  // A pending debounced live-preview update (scheduleLivePreviewUpdate, 600ms) can
  // still be queued from the user's last keystroke. If it fires while renderDocxToPdf
  // below is mid-flight, its own concurrent docx-preview render into the preview
  // iframe reproducibly makes html2canvas mismeasure this template's justified body
  // text and split words mid-letter ("technic" / "al support") in the exported PDF.
  clearTimeout(livePreviewTimer);

  const startedAt = Date.now();
  downloadPdfBtn.disabled = true;
  downloadWordBtn.disabled = true;
  applyManualBtn.disabled = true;
  showWizardLoading('Applying details and finalizing the MSA...');

  let success = false;
  try {
    if (!validateManualFields()) {
      setManualFieldsStatus('Please fill in all manual details before applying.', true);
      return;
    }
    setManualFieldsStatus('');

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

    setPreviewStatus('Manual details applied and the MSA is finalized. Ready to download.');
    success = true;
  } catch (err) {
    setPreviewStatus(`Could not finalize the MSA: ${err.message}`, true);
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
  clearW9Preview();
  setActivePreviewTab('generated');
  previewStatus.textContent = '';
  previewStatus.classList.remove('error');
  setManualFieldsStatus('');
  setStatus('Ready for a new document.');
  showWizardStep(1);
}

startNewBtn.addEventListener('click', resetWizardForNewDocument);

const previewToggleBtn = document.getElementById('previewToggleBtn');
const previewToggleLabel = document.getElementById('previewToggleLabel');
const previewPanel = document.getElementById('previewPanel');
const previewStatus = document.getElementById('previewStatus');
const previewEmptyState = document.getElementById('previewEmptyState');
const previewLoading = document.getElementById('previewLoading');
const previewZoomControls = document.getElementById('previewZoomControls');
const previewZoomOutBtn = document.getElementById('previewZoomOutBtn');
const previewZoomInBtn = document.getElementById('previewZoomInBtn');
const previewZoomLevel = document.getElementById('previewZoomLevel');
const previewFrame = document.getElementById('previewFrame');
const previewTabGenerated = document.getElementById('previewTabGenerated');
const previewTabW9 = document.getElementById('previewTabW9');
const previewViewGenerated = document.getElementById('previewViewGenerated');
const previewViewW9 = document.getElementById('previewViewW9');
const w9PreviewEmptyState = document.getElementById('w9PreviewEmptyState');
const w9PreviewLoading = document.getElementById('w9PreviewLoading');
const w9PreviewPages = document.getElementById('w9PreviewPages');
const w9PreviewImage = document.getElementById('w9PreviewImage');

const PREVIEW_ZOOM_MIN = 0.3;
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

previewZoomOutBtn.addEventListener('click', () => {
  if (activePreviewTab === 'w9') zoomW9PreviewAtCenter(w9PreviewZoom - PREVIEW_ZOOM_STEP, true);
  else zoomPreviewAtCenter(previewZoom - PREVIEW_ZOOM_STEP, true);
});
previewZoomInBtn.addEventListener('click', () => {
  if (activePreviewTab === 'w9') zoomW9PreviewAtCenter(w9PreviewZoom + PREVIEW_ZOOM_STEP, true);
  else zoomPreviewAtCenter(previewZoom + PREVIEW_ZOOM_STEP, true);
});

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
      // Without this, docx-preview renders every Word tab character as a
      // fixed small em-space instead of computing where it should actually
      // land — its own tab-stop system (the class fixTemplateTabStops below
      // corrects) only runs when this flag is on.
      experimental: true,
    });

    alignHeaderLogoRight(doc.body);
    fixNoSpacingParagraphMargins(doc.body);
    forceDocumentFont(doc.body);
    fixTemplateTabStops(doc.body);
    removeEmptyNumberingStubs(doc.body);
    repaginateToLetterPages(doc.body);
    moveMainSignatureBlockToWitnessPage(doc.body);
    removeInitialFromClosingPages(doc.body);
    mergeSparseLastPageBack(doc.body);
    addInitialToPages(doc.body, [2, 4]);
    normalizePreviewFooter(doc.body);

    applyPreviewZoomStyleOverrides(doc);
    applyPreviewZoom(false);
    attachDragPan(doc);
    attachWheelZoom(doc);
    previewZoomControls.hidden = activePreviewTab !== 'generated';
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

let activePreviewTab = 'generated';

function refreshPreviewZoomControlsVisibility() {
  previewZoomControls.hidden =
    activePreviewTab === 'generated' ? !previewFrame.classList.contains('is-ready') : w9PreviewPages.hidden;
  previewZoomLevel.textContent = `${Math.round((activePreviewTab === 'w9' ? w9PreviewZoom : previewZoom) * 100)}%`;
}

function setActivePreviewTab(tab) {
  activePreviewTab = tab;
  previewTabGenerated.classList.toggle('is-active', tab === 'generated');
  previewTabGenerated.setAttribute('aria-selected', String(tab === 'generated'));
  previewTabW9.classList.toggle('is-active', tab === 'w9');
  previewTabW9.setAttribute('aria-selected', String(tab === 'w9'));
  previewViewGenerated.hidden = tab !== 'generated';
  previewViewW9.hidden = tab !== 'w9';
  refreshPreviewZoomControlsVisibility();
}

previewTabGenerated.addEventListener('click', () => setActivePreviewTab('generated'));
previewTabW9.addEventListener('click', () => setActivePreviewTab('w9'));

const W9_PREVIEW_ZOOM_MIN = 0.3;
const W9_PREVIEW_ZOOM_MAX = 2;
let w9PreviewZoom = 1;

function getW9ZoomTargets() {
  const canvases = Array.from(w9PreviewPages.querySelectorAll('canvas.w9-preview-page'));
  if (canvases.length) return canvases;
  return w9PreviewImage.hidden ? [] : [w9PreviewImage];
}

// The "natural" (100%-zoom) display size. For a canvas this is NOT its raw
// pixel buffer — that's rendered at a higher resolution than 100% display
// size on purpose (see W9_PDF_RENDER_OVERSAMPLE), so zooming in via CSS still
// has real pixel detail behind it instead of just stretching a blurrier image.
function w9NaturalSize(el) {
  if (el.tagName === 'CANVAS') {
    return { w: Number(el.dataset.naturalWidth) || el.width, h: Number(el.dataset.naturalHeight) || el.height };
  }
  return { w: el.naturalWidth || el.clientWidth, h: el.naturalHeight || el.clientHeight };
}

// Each page/image is resized directly (rather than transform-scaling one
// shared wrapper, the way the MSA docx preview does) since the W-9 preview
// can hold several independently-sized canvases stacked with gaps between
// them — there's no single natural width/height to scale as one unit.
function applyW9PreviewZoom(animate) {
  getW9ZoomTargets().forEach((el) => {
    const { w, h } = w9NaturalSize(el);
    el.style.transition = animate ? 'width 0.15s ease, height 0.15s ease' : 'none';
    el.style.width = `${w * w9PreviewZoom}px`;
    el.style.height = `${h * w9PreviewZoom}px`;
  });
  if (activePreviewTab === 'w9') previewZoomLevel.textContent = `${Math.round(w9PreviewZoom * 100)}%`;
}

function setW9PreviewZoom(zoom, animate) {
  w9PreviewZoom = Math.min(W9_PREVIEW_ZOOM_MAX, Math.max(W9_PREVIEW_ZOOM_MIN, zoom));
  applyW9PreviewZoom(animate);
}

// Approximate cursor-anchored zoom: assumes a roughly uniform scale factor
// across whatever page is currently under the cursor. Good enough for a
// form that's usually one or two pages — attachDragPan lets the user
// correct any drift by hand.
function zoomW9PreviewAt(zoom, cursorX, cursorY, animate) {
  const scroller = w9PreviewPages;
  const oldZoom = w9PreviewZoom || 1;
  const docX = (scroller.scrollLeft + cursorX) / oldZoom;
  const docY = (scroller.scrollTop + cursorY) / oldZoom;

  setW9PreviewZoom(zoom, animate);

  scroller.scrollLeft = docX * w9PreviewZoom - cursorX;
  scroller.scrollTop = docY * w9PreviewZoom - cursorY;
}

function zoomW9PreviewAtCenter(zoom, animate) {
  zoomW9PreviewAt(zoom, w9PreviewPages.clientWidth / 2, w9PreviewPages.clientHeight / 2, animate);
}

function attachW9DragPan() {
  const scroller = w9PreviewPages;
  let dragging = false;
  let pointerId = null;
  let startX = 0;
  let startY = 0;
  let startLeft = 0;
  let startTop = 0;

  scroller.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    dragging = true;
    pointerId = e.pointerId;
    startX = e.clientX;
    startY = e.clientY;
    startLeft = scroller.scrollLeft;
    startTop = scroller.scrollTop;
    scroller.classList.add('is-dragging');
    try {
      scroller.setPointerCapture(pointerId);
    } catch {
      // Pointer capture is a nice-to-have for dragging past the frame edge; ignore if unsupported.
    }
    e.preventDefault();
  });

  scroller.addEventListener('pointermove', (e) => {
    if (!dragging) return;
    scroller.scrollLeft = startLeft - (e.clientX - startX);
    scroller.scrollTop = startTop - (e.clientY - startY);
  });

  const stopDrag = () => {
    if (!dragging) return;
    dragging = false;
    scroller.classList.remove('is-dragging');
    try {
      if (pointerId !== null) scroller.releasePointerCapture(pointerId);
    } catch {
      // Already released or unsupported; nothing to do.
    }
  };
  scroller.addEventListener('pointerup', stopDrag);
  scroller.addEventListener('pointercancel', stopDrag);
  scroller.addEventListener('pointerleave', (e) => {
    if (pointerId === null || !scroller.hasPointerCapture?.(pointerId)) stopDrag(e);
  });

  scroller.addEventListener(
    'wheel',
    (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      const rect = scroller.getBoundingClientRect();
      const step = Math.min(0.15, Math.max(0.01, Math.abs(e.deltaY) / 200));
      zoomW9PreviewAt(w9PreviewZoom + (e.deltaY < 0 ? step : -step), e.clientX - rect.left, e.clientY - rect.top, false);
    },
    { passive: false }
  );

  scroller.addEventListener('dblclick', (e) => {
    const rect = scroller.getBoundingClientRect();
    const target = w9PreviewZoom >= W9_PREVIEW_ZOOM_MAX - 0.1 ? 1 : w9PreviewZoom + 0.2;
    zoomW9PreviewAt(target, e.clientX - rect.left, e.clientY - rect.top, true);
  });
}

attachW9DragPan();

let w9PreviewObjectUrl = null;
let w9PreviewRequestId = 0;

function clearW9Preview() {
  w9PreviewRequestId++;
  if (w9PreviewObjectUrl) {
    URL.revokeObjectURL(w9PreviewObjectUrl);
    w9PreviewObjectUrl = null;
  }
  w9PreviewPages.hidden = true;
  w9PreviewPages.querySelectorAll('canvas.w9-preview-page').forEach((c) => c.remove());
  w9PreviewImage.hidden = true;
  w9PreviewImage.removeAttribute('src');
  w9PreviewImage.style.removeProperty('width');
  w9PreviewImage.style.removeProperty('height');
  w9PreviewLoading.hidden = true;
  w9PreviewEmptyState.hidden = false;
  w9PreviewZoom = 1;
  if (activePreviewTab === 'w9') refreshPreviewZoomControlsVisibility();
}

// Rendered at more pixels than the 100%-zoom display size calls for, so
// zooming in (up to W9_PREVIEW_ZOOM_MAX) via CSS width/height still draws on
// real detail instead of stretching an already-100%-sized bitmap into a
// blurry mess — the zoom used to *report* a bigger percentage without the
// page actually looking any bigger or sharper.
const W9_PDF_RENDER_OVERSAMPLE = W9_PREVIEW_ZOOM_MAX;

// The browser's built-in PDF plugin (what an <iframe src="blob:...pdf">
// relies on) isn't guaranteed to be available or enabled, so it can render
// as a silently blank frame with no error to catch. Rendering PDF pages to
// canvas via pdf.js — already a dependency, already used for W-9 OCR —
// works the same way everywhere instead of depending on that plugin.
async function renderW9PdfPreview(file, requestId) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  if (requestId !== w9PreviewRequestId) return;

  const baseScale = 1.5;
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const displayViewport = page.getViewport({ scale: baseScale });
    const renderViewport = page.getViewport({ scale: baseScale * W9_PDF_RENDER_OVERSAMPLE });
    const canvas = document.createElement('canvas');
    canvas.className = 'w9-preview-page';
    canvas.width = renderViewport.width;
    canvas.height = renderViewport.height;
    canvas.dataset.naturalWidth = displayViewport.width;
    canvas.dataset.naturalHeight = displayViewport.height;
    await page.render({ canvasContext: canvas.getContext('2d'), viewport: renderViewport }).promise;
    if (requestId !== w9PreviewRequestId) return;
    w9PreviewPages.appendChild(canvas);
  }
  applyW9PreviewZoom(false);
}

async function renderW9Preview(file) {
  clearW9Preview();
  if (!file) return;

  const requestId = w9PreviewRequestId;
  w9PreviewEmptyState.hidden = true;

  if (isImageFile(file)) {
    w9PreviewObjectUrl = URL.createObjectURL(file);
    w9PreviewImage.src = w9PreviewObjectUrl;
    w9PreviewImage.hidden = false;
    w9PreviewPages.hidden = false;
    if (activePreviewTab === 'w9') refreshPreviewZoomControlsVisibility();
    return;
  }

  w9PreviewLoading.hidden = false;
  try {
    await renderW9PdfPreview(file, requestId);
    if (requestId !== w9PreviewRequestId) return;
    w9PreviewPages.hidden = false;
    if (activePreviewTab === 'w9') refreshPreviewZoomControlsVisibility();
  } catch {
    if (requestId !== w9PreviewRequestId) return;
    w9PreviewEmptyState.hidden = false;
  } finally {
    if (requestId === w9PreviewRequestId) w9PreviewLoading.hidden = true;
  }
}

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

  const INTRO_RELOADED_FLAG = 'docgen_intro_reloaded';

  if (sessionStorage.getItem(INTRO_RELOADED_FLAG)) {
    sessionStorage.removeItem(INTRO_RELOADED_FLAG);
    overlay.remove();
    return;
  }

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

    setTimeout(() => {
      sessionStorage.setItem(INTRO_RELOADED_FLAG, '1');
      window.location.reload();
    }, FLY_DURATION + 750);
  }

  runIntroAnimation();
})();
