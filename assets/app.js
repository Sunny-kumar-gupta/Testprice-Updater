/* ================================================================
   Department Wise Test Price Updater — Main Script
   File: assets/app.js

   Logic:
   1. User uploads an Excel file (.xlsx / .xls)
   2. User enters department-wise percentage rules
   3. App reads the file, applies rules to "Test Price" column
   4. Generates and downloads the updated Excel file

   Formula: New Price = (Old Price × Percentage) / 100

   Dependencies: SheetJS (xlsx) — loaded via CDN in index.html
   ================================================================ */

'use strict';

// ── State ──────────────────────────────────────────────────────────
let workbook        = null;   // Original SheetJS workbook
let updatedWorkbook = null;   // Modified workbook ready for download

// ── DOM References ─────────────────────────────────────────────────
const fileInput     = document.getElementById('fileInput');
const dropZone      = document.getElementById('dropZone');
const filePill      = document.getElementById('filePill');
const fileNameEl    = document.getElementById('fileName');
const removeFileBtn = document.getElementById('removeFileBtn');
const rulesInput    = document.getElementById('rulesInput');
const processBtn    = document.getElementById('processBtn');
const downloadBtn   = document.getElementById('downloadBtn');
const toastArea     = document.getElementById('toast-area');
const resultCard    = document.getElementById('result-card');


/* ================================================================
   TOAST HELPER
   Displays a temporary status message at the top of the page.
   @param {string}  msg  - Message to show
   @param {'error'|'success'|'info'} type - Visual style
   ================================================================ */
function showToast(msg, type = 'info') {
  const icons = { error: '❌', success: '✅', info: 'ℹ️' };

  const div = document.createElement('div');
  div.className = `toast ${type}`;
  div.innerHTML = `<span class="icon">${icons[type]}</span><span>${msg}</span>`;

  toastArea.innerHTML = '';       // Clear any previous toast
  toastArea.appendChild(div);

  // Auto-remove: errors stay 8 s, others 5 s
  setTimeout(() => div.remove(), type === 'error' ? 8000 : 5000);
}


/* ================================================================
   FILE UPLOAD — via <input> click or drag-and-drop
   ================================================================ */
fileInput.addEventListener('change', handleFileSelect);

// Drag-over highlight
dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('drag-over');
});

// Handle drop
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');

  const files = e.dataTransfer.files;
  if (files.length) {
    fileInput.files = files;    // Assign dropped file to the input
    handleFileSelect();
  }
});

/**
 * Validates the selected file and reads it with FileReader.
 * SheetJS parses the ArrayBuffer into a workbook object.
 */
function handleFileSelect() {
  const file = fileInput.files[0];
  if (!file) return;

  // --- Validate file extension ---
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls'].includes(ext)) {
    showToast('Invalid file type. Please upload an .xlsx or .xls file.', 'error');
    resetFile();
    return;
  }

  // --- Read file as ArrayBuffer ---
  const reader = new FileReader();

  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: 'array' });

      // Show file name pill
      fileNameEl.textContent = file.name;
      filePill.classList.add('show');

      showToast(`File loaded: "${file.name}"`, 'info');

      // Clear any previous result
      hideResult();
      updatedWorkbook = null;

    } catch (err) {
      showToast('Could not read Excel file. Make sure it is not corrupted.', 'error');
      resetFile();
    }
  };

  reader.readAsArrayBuffer(file);
}

/** Remove the selected file and reset state */
removeFileBtn.addEventListener('click', () => {
  resetFile();
  showToast('File removed.', 'info');
});

function resetFile() {
  fileInput.value    = '';
  workbook           = null;
  updatedWorkbook    = null;
  filePill.classList.remove('show');
  fileNameEl.textContent = '—';
  hideResult();
}


/* ================================================================
   RULE PARSER
   Converts the textarea text into a Map for fast lookups.

   Input example:
     CLINICAL PATHOLOGY - 30%
     BIOCHEMISTRY - 50%

   Output: Map { 'CLINICAL PATHOLOGY' => 30, 'BIOCHEMISTRY' => 50 }

   @param  {string}          rawText
   @returns {Map<string,number>}
   ================================================================ */
function parseRules(rawText) {
  const rulesMap = new Map();
  const lines    = rawText.split('\n');

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;   // Skip blank lines

    // Regex: capture department name and percentage number
    // Pattern: <name> - <number>%
    const match = trimmed.match(/^(.+?)\s*-\s*([\d.]+)\s*%\s*$/i);

    if (!match) {
      // Warn in console, but continue parsing other lines
      console.warn('[Rules] Could not parse line:', trimmed);
      continue;
    }

    const deptKey = match[1].trim().toUpperCase();  // Normalize to UPPERCASE
    const pct     = parseFloat(match[2]);

    if (isNaN(pct) || pct < 0) continue;            // Skip invalid percentages

    rulesMap.set(deptKey, pct);
  }

  return rulesMap;
}


/* ================================================================
   PROCESS BUTTON — Validates inputs, then runs price update logic
   ================================================================ */
processBtn.addEventListener('click', processFile);

function processFile() {

  // --- Validation 1: No file uploaded ---
  if (!workbook) {
    showToast('Please upload an Excel file first.', 'error');
    return;
  }

  // --- Validation 2: Empty rules ---
  const rawRules = rulesInput.value.trim();
  if (!rawRules) {
    showToast('Please enter at least one department rule.', 'error');
    return;
  }

  // --- Validation 3: Parse rules ---
  const rulesMap = parseRules(rawRules);
  if (rulesMap.size === 0) {
    showToast(
      'No valid rules found. Use format: DEPARTMENT NAME - 30%',
      'error'
    );
    return;
  }

  // --- Read first sheet from workbook ---
  const sheetName = workbook.SheetNames[0];
  const sheet     = workbook.Sheets[sheetName];

  // Convert sheet to array of row objects (header row → keys)
  const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  if (jsonData.length === 0) {
    showToast('The Excel file appears to be empty.', 'error');
    return;
  }

  // --- Validation 4: Check required columns ---
  const headers      = Object.keys(jsonData[0]).map(h => h.trim());
  const requiredCols = ['Test Id', 'Test Name', 'Test Code', 'Department', 'Test Price'];

  const missingCols = requiredCols.filter(req =>
    !headers.some(h => h.toLowerCase() === req.toLowerCase())
  );

  if (missingCols.length > 0) {
    showToast(
      `Missing required column(s): ${missingCols.join(', ')}. Please check your Excel file.`,
      'error'
    );
    return;
  }

  // --- Show loading state on button ---
  processBtn.disabled = true;
  processBtn.innerHTML = `<span class="spinner"></span> Processing…`;

  // Tiny timeout lets the browser repaint the loading state
  // before the synchronous JS work begins
  setTimeout(() => {
    try {
      applyPriceRules(jsonData, rulesMap, sheetName);
    } catch (err) {
      showToast('An unexpected error occurred: ' + err.message, 'error');
    } finally {
      processBtn.disabled = false;
      processBtn.innerHTML = `⚡ Apply Price Rules`;
    }
  }, 80);
}


/* ================================================================
   CORE LOGIC — applyPriceRules
   Iterates every row, matches Department to rules, and updates
   the Test Price column only.

   Formula: New Price = (Old Price × Percentage) / 100

   @param {Object[]}        rows      - Sheet data as row objects
   @param {Map<string,number>} rulesMap - Parsed dept→pct rules
   @param {string}          sheetName - Name of the sheet to update
   ================================================================ */
function applyPriceRules(rows, rulesMap, sheetName) {
  let totalRows   = rows.length;
  let updatedRows = 0;
  let skippedRows = 0;

  // Track per-department counts for the summary table
  // Structure: Map { 'CLINICAL PATHOLOGY' => { pct: 30, count: 5 } }
  const deptStats = new Map();

  // Find the actual header key strings (preserving original file casing)
  const sampleRow = rows[0];
  const deptKey   = Object.keys(sampleRow).find(k => k.toLowerCase() === 'department');
  const priceKey  = Object.keys(sampleRow).find(k => k.toLowerCase() === 'test price');

  // --- Walk through every data row ---
  for (let i = 0; i < rows.length; i++) {
    const row       = rows[i];
    const deptRaw   = String(row[deptKey] || '').trim();
    const deptUpper = deptRaw.toUpperCase();     // Normalize for lookup
    const oldPrice  = parseFloat(row[priceKey]);

    // Skip if:
    // • Department is not in rules
    // • Test Price is not a valid number
    if (!rulesMap.has(deptUpper) || isNaN(oldPrice)) {
      skippedRows++;
      continue;
    }

    // Apply formula
    const pct      = rulesMap.get(deptUpper);
    const newPrice = parseFloat(((oldPrice * pct) / 100).toFixed(2));

    // ★ Only modify Test Price — all other keys stay untouched
    rows[i][priceKey] = newPrice;
    updatedRows++;

    // Record dept stats
    if (!deptStats.has(deptRaw)) {
      deptStats.set(deptRaw, { pct, count: 0 });
    }
    deptStats.get(deptRaw).count++;
  }

  // --- Rebuild the sheet from modified JSON data ---
  const newSheet = XLSX.utils.json_to_sheet(rows);

  // Clone the workbook (avoid mutating the original)
  updatedWorkbook = {
    ...workbook,
    Sheets: { ...workbook.Sheets, [sheetName]: newSheet }
  };

  // --- Show result summary ---
  showResultCard({ totalRows, updatedRows, skippedRows, deptStats });

  showToast(
    `Done! ${updatedRows} row(s) updated across ${deptStats.size} department(s).`,
    'success'
  );
}


/* ================================================================
   RESULT CARD — Populate and show the summary panel
   ================================================================ */
function showResultCard({ totalRows, updatedRows, skippedRows, deptStats }) {
  // Fill stat boxes
  document.getElementById('stat-total').textContent   = totalRows;
  document.getElementById('stat-updated').textContent = updatedRows;
  document.getElementById('stat-skipped').textContent = skippedRows;
  document.getElementById('stat-depts').textContent   = deptStats.size;

  // Build department breakdown table body
  const tbody = document.querySelector('#deptBreakdown tbody');
  tbody.innerHTML = '';

  if (deptStats.size === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="3" style="color:var(--muted)">
          No departments matched the rules.
        </td>
      </tr>`;
  } else {
    for (const [dept, { pct, count }] of deptStats) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${dept}</td>
        <td><span class="pill-pct">${pct}%</span></td>
        <td>${count} row${count !== 1 ? 's' : ''}</td>
      `;
      tbody.appendChild(tr);
    }
  }

  // Show card and scroll to it
  resultCard.classList.add('show');
  resultCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function hideResult() {
  resultCard.classList.remove('show');
}


/* ================================================================
   DOWNLOAD BUTTON — Triggers SheetJS file write & browser download
   ================================================================ */
downloadBtn.addEventListener('click', () => {
  if (!updatedWorkbook) {
    showToast(
      'No processed file available. Please click "Apply Price Rules" first.',
      'error'
    );
    return;
  }

  // XLSX.writeFile handles the binary conversion and download
  XLSX.writeFile(updatedWorkbook, 'Updated_Test_Prices.xlsx');
  showToast('File downloaded: Updated_Test_Prices.xlsx', 'success');
});
