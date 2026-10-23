/* ==========================================================
   EPS Part Controller - Print Label (all logic runs in the browser)

   Sections:
     1.  Helpers & constants
     2.  App state
     3.  Column detection (Excel/CSV)
     4.  Bulk import (Excel/CSV)
     5.  Manual labels
     6.  Settings (label size, page, fonts)
     7.  QR code & barcode generators
     8.  Preview rendering
     9.  Export helpers (shared by PDF / Word)
     10. Export: PDF
     12. Export: Word (.docx)
     13. Event wiring & first render
   ========================================================== */


/* ==========================================================
   1. HELPERS & CONSTANTS
   ========================================================== */
const $ = id => document.getElementById(id);

// Escape text before putting it into innerHTML
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({
  '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
}[c]));

// Normalize a header name for matching ("Part_Name" -> "part name")
const norm = v => String(v ?? '')
  .trim()
  .toLowerCase()
  .replace(/[_\-.]+/g, ' ')
  .replace(/\s+/g, ' ');

// Copies must be a whole number from 1 to 100
const clampCopies = v => Math.max(1, Math.min(100, parseInt(v) || 1));

// A4 page size in mm: [width, height]
const A4 = { portrait: [210, 297], landscape: [297, 210] };


/* ==========================================================
   2. APP STATE
   ========================================================== */
let workbook    = null;   // parsed Excel workbook
let rawRows     = [];     // current sheet as array of rows
let headers     = [];     // header names of the chosen header row
let bulkItems   = [];     // items imported from Excel/CSV
let manualItems = [];     // items typed in by hand
let bulkActive  = false;  // true after the first successful import (enables live re-import)


/* ==========================================================
   3. COLUMN DETECTION (Excel/CSV)
   Each field lists header names that match exactly (first try)
   and partial names (second try).
   ========================================================== */
const FIELDS = {
  partName: {
    exact: ['part name', 'partname', 'name', 'nama part', 'nama barang', 'item name', 'item', 'description'],
    fuzzy: ['part name', 'nama part', 'nama barang', 'item name', 'description', 'deskripsi']
  },
  partNumber: {
    exact: ['part number', 'part no', 'partnumber', 'pn', 'p n', 'part code', 'kode part', 'number'],
    fuzzy: ['part number', 'part no', 'part code', 'kode part', 'p/n']
  },
  category: {
    exact: ['category', 'kategori', 'type', 'jenis'],
    fuzzy: ['category', 'kategori']
  },
  serialNumber: {
    exact: ['serial number', 'serial no', 'serialnumber', 'serial', 'sn', 's n'],
    fuzzy: ['serial', 's/n']
  },
  copiesCol: {
    exact: ['copies', 'copy', 'qty', 'quantity', 'jumlah', 'qty label', 'labels'],
    fuzzy: ['copies', 'quantity', 'qty', 'jumlah']
  }
};

const ALL_ALIASES = Object.values(FIELDS).flatMap(f => f.exact);

// Finds the row that looks most like a header (most known column names)
function detectHeaderRow(rows) {
  let best = -1, bestScore = 0, firstNonEmpty = -1;

  rows.slice(0, 30).forEach((row, i) => {
    if (firstNonEmpty < 0 && row.some(c => String(c ?? '').trim() !== '')) firstNonEmpty = i;
    const score = row.filter(c => ALL_ALIASES.includes(norm(c))).length;
    if (score > bestScore) { bestScore = score; best = i; }
  });

  return best >= 0 ? best : firstNonEmpty;
}

// Pre-selects each dropdown with the best matching column
function autoMap() {
  const taken = new Set();            // a column can only be used once
  const normalizedHeaders = headers.map(norm);

  Object.keys(FIELDS).forEach(id => {
    const select = $(id);

    let idx = normalizedHeaders.findIndex((h, i) => !taken.has(i) && FIELDS[id].exact.includes(h));
    if (idx < 0) {
      idx = normalizedHeaders.findIndex((h, i) => !taken.has(i) && FIELDS[id].fuzzy.some(a => h.includes(a)));
    }

    if (idx >= 0) {
      select.value = String(idx);
      taken.add(idx);
    }
  });
}


/* ==========================================================
   4. BULK IMPORT (Excel/CSV)
   ========================================================== */
function setBulkStatus(text, kind = '') {
  $('bulkStatus').textContent = text;
  $('bulkStatus').className = 'status ' + kind;
}

// Read the chosen file
$('excelFile').addEventListener('change', async e => {
  const file = e.target.files[0];
  if (!file) return;

  try {
    workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });

    $('fileinfo').textContent = `${file.name} · ${(file.size / 1024).toFixed(1)} KB`;
    $('sheet').innerHTML = workbook.SheetNames.map(s => `<option>${esc(s)}</option>`).join('');
    $('sheet').disabled = false;

    loadSheet(workbook.SheetNames[0]);
  } catch (err) {
    setBulkStatus('Could not read this file. Please check that it is a valid Excel or CSV file.', 'error');
  }
});

$('sheet').addEventListener('change', () => loadSheet($('sheet').value));

// Load one worksheet and guess its header row
function loadSheet(name) {
  rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: false });

  const headerIdx = detectHeaderRow(rawRows);
  if (headerIdx < 0) {
    setBulkStatus('This worksheet appears to be empty.', 'error');
    $('importBtn').disabled = true;
    return;
  }

  // header-row dropdown (first 30 rows, with a short preview of each)
  $('headerRow').innerHTML = rawRows
    .slice(0, Math.min(rawRows.length, 30))
    .map((row, i) => {
      const preview = row.filter(x => String(x).trim()).slice(0, 3).join(' · ').slice(0, 40) || '(blank)';
      return `<option value="${i}">Row ${i + 1}: ${esc(preview)}</option>`;
    })
    .join('');
  $('headerRow').disabled = false;
  $('headerRow').value = String(headerIdx);

  buildMapping();
  setBulkStatus('File loaded. Columns were detected automatically; check them, then import.', 'ok');
}

$('headerRow').addEventListener('change', () => {
  buildMapping();
  if (bulkActive) importBulk();
});

// Fill the column dropdowns from the chosen header row
function buildMapping() {
  headers = (rawRows[Number($('headerRow').value) || 0] || [])
    .map((h, i) => String(h ?? '').trim() || `Column ${i + 1}`);

  Object.keys(FIELDS).forEach(id => {
    const emptyLabel = id === 'copiesCol' ? '— Use "Copies per item" —' : '— None —';
    $(id).innerHTML =
      `<option value="">${emptyLabel}</option>` +
      headers.map((h, i) => `<option value="${i}">${esc(h)}</option>`).join('');
    $(id).disabled = false;
  });

  autoMap();
  $('importBtn').disabled = false;
}

// Once imported, changing any mapping re-imports automatically
['partName', 'partNumber', 'category', 'serialNumber', 'copiesCol', 'bulkCopies'].forEach(id =>
  $(id).addEventListener('change', () => { if (bulkActive) importBulk(); })
);

// Turn the sheet rows into label items
function importBulk() {
  if (!rawRows.length) return;

  // which column index is used for each field (-1 = not mapped)
  const cols = {};
  for (const id of ['partName', 'partNumber', 'category', 'serialNumber', 'copiesCol']) {
    cols[id] = $(id).value === '' ? -1 : Number($(id).value);
  }

  if (['partName', 'partNumber', 'category', 'serialNumber'].every(k => cols[k] < 0)) {
    setBulkStatus('Please map at least one column.', 'error');
    return;
  }

  const get = (row, i) => i < 0 ? '' : String(row[i] ?? '').trim();
  const defaultCopies = clampCopies($('bulkCopies').value);
  const headerIdx = Number($('headerRow').value) || 0;

  bulkItems = rawRows
    .slice(headerIdx + 1)                                            // rows below the header
    .filter(row => row.some(v => String(v ?? '').trim() !== ''))     // skip blank rows
    .map(row => ({
      partName:     get(row, cols.partName),
      partNumber:   get(row, cols.partNumber),
      category:     get(row, cols.category),
      serialNumber: get(row, cols.serialNumber),
      copies:       cols.copiesCol >= 0 ? clampCopies(get(row, cols.copiesCol)) : defaultCopies
    }))
    .filter(r => r.partName || r.partNumber || r.category || r.serialNumber);

  if (!bulkItems.length) {
    setBulkStatus('No data rows found below the header row.', 'error');
    bulkActive = false;
    render();
    return;
  }

  bulkActive = true;
  const missing = cols.serialNumber < 0 ? ' No serial column mapped, so no barcodes will be made.' : '';
  setBulkStatus(`Imported ${bulkItems.length} item(s).${missing}`, missing ? 'error' : 'ok');
  render();
}

$('importBtn').addEventListener('click', importBulk);


/* ==========================================================
   5. MANUAL LABELS
   ========================================================== */
function renderManualList() {
  $('manualList').innerHTML = manualItems.map((m, i) => `
    <div class="mitem">
      <span>
        <b>${esc(m.partName || m.partNumber || m.serialNumber || 'Label')}</b>
        ${m.serialNumber ? ' · ' + esc(m.serialNumber) : ''} · ×${m.copies}
      </span>
      <button data-i="${i}" title="Remove">×</button>
    </div>`).join('');
}

// Remove one manual item
$('manualList').addEventListener('click', e => {
  const btn = e.target.closest('button[data-i]');
  if (!btn) return;
  manualItems.splice(Number(btn.dataset.i), 1);
  renderManualList();
  render();
});

function addManual() {
  const item = {
    partName:     $('mName').value.trim(),
    partNumber:   $('mNumber').value.trim(),
    category:     $('mCategory').value.trim(),
    serialNumber: $('mSerial').value.trim(),
    copies:       clampCopies($('mCopies').value)
  };
  const status = $('manualStatus');

  if (!(item.partName || item.partNumber || item.category || item.serialNumber)) {
    status.textContent = 'Enter at least one field.';
    status.className = 'status error';
    return;
  }

  manualItems.push(item);

  // reset the form
  ['mName', 'mNumber', 'mCategory', 'mSerial'].forEach(id => $(id).value = '');
  $('mCopies').value = 1;
  $('mName').focus();

  status.textContent = `Added. ${manualItems.length} manual item(s) in the preview.`;
  status.className = 'status ok';
  renderManualList();
  render();
}

$('addManual').addEventListener('click', addManual);

// Enter key adds the label
['mName', 'mNumber', 'mCategory', 'mSerial', 'mCopies'].forEach(id =>
  $(id).addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); addManual(); }
  })
);


/* ==========================================================
   6. SETTINGS (label size, page, fonts)
   ========================================================== */

// Label size in mm: [width, height]
function getSize() {
  if ($('size').value === 'custom') {
    const w = Math.max(20, Math.min(200, parseFloat($('cw').value) || 60));
    const h = Math.max(15, Math.min(200, parseFloat($('ch').value) || 40));
    return [w, h];
  }
  return $('size').value.split('x').map(Number);
}

// QR side length in mm (top-right corner of the label)
const qrSize = (w, h, pad) => Math.min((w - 2 * pad) * .34, h * .45);

// Scaled font sizes / padding, used by PDF and Word export
function metrics(s) {
  const u = Math.max(Math.min(s.h / 38, s.w / 63), .6);   // scale factor vs. the 63x38 design
  const pad = Math.max(2 * u, 1.4);
  return {
    u,
    pad,
    qr: qrSize(s.w, s.h, pad),
    brand: Math.max(7 * u, 5.5),
    name: Math.max(10 * u, 8),
    row: Math.max(7.5 * u, 6),
    compact: s.h < 32           // short labels hide the "PART NAME" caption
  };
}

// Read all settings, update CSS variables + @page rule, return the settings object
function applySettings() {
  const [w, h] = getSize();
  const mode   = $('mode').value;
  const orient = $('orient').value;
  const margin = Math.max(0, parseFloat($('margin').value) || 0);
  const gap    = Math.max(0, parseFloat($('gap').value) || 0);

  // Roll mode + portrait = label turned 90° so the page is H×W
  const rotate = mode === 'roll' && orient === 'portrait';

  // label + slot size
  const root = document.documentElement.style;
  root.setProperty('--lw', w + 'mm');
  root.setProperty('--lh', h + 'mm');
  root.setProperty('--sw', (rotate ? h : w) + 'mm');
  root.setProperty('--sh', (rotate ? w : h) + 'mm');
  root.setProperty('--gap', gap + 'mm');

  // fonts, padding and QR scale with the label size
  const u = Math.max(Math.min(h / 38, w / 63), .6);
  root.setProperty('--f-brand', Math.max(7 * u, 5.5).toFixed(2) + 'pt');
  root.setProperty('--f-name',  Math.max(10 * u, 8).toFixed(2) + 'pt');
  root.setProperty('--f-row',   Math.max(7.5 * u, 6).toFixed(2) + 'pt');
  root.setProperty('--pad',     Math.max(2 * u, 1.4).toFixed(2) + 'mm');
  root.setProperty('--qr',      qrSize(w, h, Math.max(2 * u, 1.4)).toFixed(2) + 'mm');

  // real @page rule, injected so it follows the settings
  $('pageStyle').textContent = mode === 'roll'
    ? `@page{size:${rotate ? h : w}mm ${rotate ? w : h}mm;margin:0}`
    : `@page{size:A4 ${orient};margin:${margin}mm}`;

  // show/hide dependent controls
  document.body.classList.toggle('roll', mode === 'roll');
  $('customSize').style.display = $('size').value === 'custom' ? 'grid' : 'none';
  $('a4opts').style.display = mode === 'a4' ? 'grid' : 'none';

  $('orient').title = mode === 'a4'
    ? 'Sets the A4 page orientation'
    : 'Landscape = label printed as designed; Portrait = label turned 90°';

  $('printerNotice').textContent = mode === 'a4'
    ? 'A regular A4 printer needs compatible A4 sticker sheets. Page layout sets the A4 orientation; the label size is set separately above.'
    : 'Roll mode prints one label per page at the exact label size. This needs a printer that supports roll media; a regular A4 printer cannot feed a label roll. Portrait turns each label 90° on the page.';

  $('sizeStat').textContent = `${w} × ${h} mm`;

  return { w, h, mode, orient, margin, gap, compact: h < 32 };
}


/* ==========================================================
   7. QR CODE & BARCODE GENERATORS
   ========================================================== */

/* ---- QR code (encodes the serial number) ---- */
const qrCache = {};

function qrOf(value) {
  if (qrCache[value]) return qrCache[value];
  if (typeof qrcode === 'undefined') throw new Error('lib');

  qrcode.stringToBytes = qrcode.stringToBytesFuncs['UTF-8'];
  const q = qrcode(0, 'M');           // auto version, medium error correction
  q.addData(String(value));
  q.make();
  return qrCache[value] = q;
}

// QR as PNG bytes (used by Word export)
const qrPngCache = {};

function qrPng(value) {
  if (qrPngCache[value]) return qrPngCache[value];

  const q = qrOf(value);
  const n = q.getModuleCount();
  const k = 8;                        // pixels per module
  const canvas = document.createElement('canvas');
  canvas.width = canvas.height = n * k;

  const g = canvas.getContext('2d');
  g.fillStyle = '#fff';
  g.fillRect(0, 0, n * k, n * k);
  g.fillStyle = '#000';
  for (let r = 0; r < n; r++) {
    for (let c = 0; c < n; c++) {
      if (q.isDark(r, c)) g.fillRect(c * k, r * k, k, k);
    }
  }
  return qrPngCache[value] = dataUrlToBytes(canvas.toDataURL('image/png'));
}

/* ---- Code128 barcode ---- */

// Barcode as a string of 1/0 bars (used by PDF export for vector bars)
function barBits(value) {
  const out = {};
  JsBarcode(out, value, { format: 'CODE128' });
  return out.encodings.map(e => e.data).join('');
}

// Barcode as PNG bytes (used by Word export)
const barPngCache = {};

function barPng(value) {
  if (barPngCache[value]) return barPngCache[value];

  const canvas = document.createElement('canvas');
  JsBarcode(canvas, value, { format: 'CODE128', displayValue: false, margin: 0, height: 60, width: 3 });
  return barPngCache[value] = dataUrlToBytes(canvas.toDataURL('image/png'));
}

// "data:image/png;base64,...." -> Uint8Array
function dataUrlToBytes(dataUrl) {
  const bin = atob(dataUrl.split(',')[1]);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}


/* ==========================================================
   8. PREVIEW RENDERING
   ========================================================== */

// HTML for one label (barcode/QR are drawn into the empty boxes afterwards)
function labelHtml(r, i) {
  return `
  <div class="slot">
    <article class="label">
      <div class="top">
        <div class="txt">
          <div class="brand">Part Name</div>
          <div class="partname">${esc(r.partName || '—')}</div>
          <div class="info">
            <span class="key">Part number:</span><span class="val">${esc(r.partNumber || '—')}</span>
            <span class="key">Category:</span><span class="val">${esc(r.category || '—')}</span>
          </div>
        </div>
        ${r.serialNumber ? `<div class="qrbox" data-i="${i}"></div>` : ''}
      </div>
      <div class="serial">Serial number: ${esc(r.serialNumber || '—')}</div>
      <div class="barcodebox" data-i="${i}"></div>
    </article>
  </div>`;
}

// Draw the Code128 barcode (SVG) into a box
function drawBarcode(box, value) {
  if (!value) {
    box.innerHTML = '<div class="nobar">No serial number, barcode skipped</div>';
    return;
  }

  try {
    if (typeof JsBarcode === 'undefined') throw new Error('lib');

    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    JsBarcode(svg, value, { format: 'CODE128', displayValue: false, margin: 0, height: 40, width: 2 });

    // make the SVG scale to its box
    const w = svg.getAttribute('width');
    const h = svg.getAttribute('height');
    if (!svg.getAttribute('viewBox')) svg.setAttribute('viewBox', `0 0 ${parseFloat(w)} ${parseFloat(h)}`);
    svg.removeAttribute('width');
    svg.removeAttribute('height');
    svg.setAttribute('preserveAspectRatio', 'none');
    svg.setAttribute('role', 'img');
    svg.setAttribute('aria-label', 'Barcode for serial number ' + value);

    box.appendChild(svg);
  } catch (e) {
    box.innerHTML = '<div class="nobar" style="color:#b42318">Barcode unavailable for this value</div>';
  }
}

// Draw the QR code (SVG path) into a box
function drawQr(box, value) {
  try {
    const q = qrOf(value);
    const n = q.getModuleCount();

    let path = '';
    for (let r = 0; r < n; r++) {
      for (let c = 0; c < n; c++) {
        if (q.isDark(r, c)) path += `M${c} ${r}h1v1h-1z`;
      }
    }

    box.innerHTML =
      `<svg viewBox="0 0 ${n} ${n}" shape-rendering="crispEdges" role="img" aria-label="QR code for serial number ${esc(value)}">` +
      `<path d="${path}" fill="#000"/></svg>`;
  } catch (e) {
    box.remove();
  }
}

// Redraw stats + the whole preview
function render() {
  const s = applySettings();
  const { items, all } = expand();

  // stats + button states
  $('itemCount').textContent = items.length;
  $('labelCount').textContent = all.length;
  $('print').disabled = $('clear').disabled = !all.length;
  ['savePdf', 'saveDocx'].forEach(id => $(id).disabled = !all.length);

  const preview = $('preview');
  preview.className = (s.compact ? 'compact ' : '') + ($('border').checked ? 'bordered' : '');

  // empty state
  if (!all.length) {
    $('pageStat').textContent = '–';
    preview.innerHTML = `
      <div class="empty">
        <div class="symbol">▤</div>
        <strong>Your labels will appear here</strong>
        <div style="margin-top:5px">Import a spreadsheet or add a label manually.</div>
      </div>`;
    return;
  }

  // page count (A4 mode only)
  if (s.mode === 'a4') {
    const [pw, ph] = A4[s.orient];
    const usableW = pw - 2 * s.margin;
    const usableH = ph - 2 * s.margin;
    const cols = Math.max(1, Math.floor((usableW + s.gap) / (s.w + s.gap)));
    const rows = Math.max(1, Math.floor((usableH + s.gap) / (s.h + s.gap)));
    const fits = s.w <= usableW && s.h <= usableH;
    $('pageStat').textContent = fits ? Math.ceil(all.length / (cols * rows)) : '⚠';
  } else {
    $('pageStat').textContent = 'n/a';
  }

  // labels
  preview.innerHTML = '<div class="labelgrid">' + all.map(labelHtml).join('') + '</div>';
  preview.querySelectorAll('.barcodebox').forEach(b => drawBarcode(b, all[Number(b.dataset.i)].serialNumber));
  preview.querySelectorAll('.qrbox').forEach(b => drawQr(b, all[Number(b.dataset.i)].serialNumber));
}


/* ==========================================================
   9. EXPORT HELPERS (shared by PDF / Word)
   ========================================================== */
const MM = 0.352778;   // 1 pt in mm

// PDF and Word libraries are loaded on demand (first click on Save)
const LIBS = {
  jspdf: {
    url: 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',
    get: () => window.jspdf && window.jspdf.jsPDF
  },
  docx: {
    url: 'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js',
    get: () => window.docx
  }
};
const libPromises = {};

function loadLib(key) {
  if (LIBS[key].get()) return Promise.resolve();

  return libPromises[key] || (libPromises[key] = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = LIBS[key].url;
    script.onload = resolve;
    script.onerror = () => {
      delete libPromises[key];
      reject(new Error('Could not load the ' + key + ' library. Check your internet connection.'));
    };
    document.head.appendChild(script);
  }));
}

function setExp(text, kind = '') {
  $('exportStatus').textContent = text;
  $('exportStatus').className = 'status exportnote no-print ' + kind;
}

// All items (bulk + manual) and the full label list (each item repeated by its copies)
function expand() {
  const items = [...bulkItems, ...manualItems];
  const all = [];
  items.forEach(r => { for (let i = 0; i < r.copies; i++) all.push(r); });
  return { items, all };
}

// "20260121-1530" - used in file names
function stamp() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
}

function saveBlob(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}

// PDF standard fonts only cover Latin-1: swap or replace other characters
const L1 = t => String(t ?? '')
  .replace(/[\u2013\u2014]/g, '-')
  .replace(/[\u2018\u2019]/g, "'")
  .replace(/[\u201C\u201D]/g, '"')
  .replace(/\u2026/g, '...')
  .replace(/[^\x00-\xFF]/g, '?');

// A4 grid (columns/rows per page); null if one label is bigger than the usable page
function pageFit(s) {
  const [pw, ph] = A4[s.orient];
  const uw = pw - 2 * s.margin;
  const uh = ph - 2 * s.margin;
  if (s.w > uw || s.h > uh) return null;

  return {
    pw, ph, uw, uh,
    cols: Math.max(1, Math.floor((uw + s.gap) / (s.w + s.gap))),
    rows: Math.max(1, Math.floor((uh + s.gap) / (s.h + s.gap)))
  };
}

const MSG_TOO_BIG = 'This label size does not fit the A4 page with the current margin. Reduce the margin or choose a smaller label.';
const MSG_ROTATED = 'Portrait roll: PDF turns the label 90°; Word keeps it upright on a landscape page.';


/* ==========================================================
   10. EXPORT: PDF (vector text + vector barcode/QR)
   ========================================================== */

// Draw one label at (x0, y0) mm. `rot` = label turned 90° (roll + portrait).
function pdfLabel(doc, x0, y0, r, s, m, rot) {
  const W = s.w, H = s.h, iw = W - 2 * m.pad;

  // helpers that work in label coordinates and handle rotation
  const T = (text, lx, ly) => {
    const px = rot ? x0 + H - ly : x0 + lx;
    const py = rot ? y0 + lx : y0 + ly;
    rot ? doc.text(text, px, py, { angle: -90 }) : doc.text(text, px, py);
  };
  const R = (lx, ly, w, h, style) =>
    rot ? doc.rect(x0 + H - ly - h, y0 + lx, h, w, style)
        : doc.rect(x0 + lx, y0 + ly, w, h, style);

  // cut text with "..." so it fits a width
  const fit = (text, max) => {
    text = L1(text);
    if (doc.getTextWidth(text) <= max) return text;
    while (text.length > 1 && doc.getTextWidth(text + '...') > max) text = text.slice(0, -1);
    return text + '...';
  };

  let y = m.pad;
  let qrSide = 0;
  const code = r.serialNumber;

  // QR code, top-right corner (adjacent dark modules are merged into one rectangle per row)
  if (code) {
    try {
      const q = qrOf(code);
      const n = q.getModuleCount();
      const md = m.qr / n;
      const qx = m.pad + iw - m.qr;
      doc.setFillColor(0);

      for (let row = 0; row < n; row++) {
        for (let c = 0; c < n;) {
          if (q.isDark(row, c)) {
            let j = c;
            while (j < n && q.isDark(row, j)) j++;
            R(qx + c * md, y + row * md, (j - c) * md, md, 'F');
            c = j;
          } else c++;
        }
      }
      qrSide = m.qr;
    } catch (e) { /* QR skipped */ }
  }

  const textW = qrSide ? iw - qrSide - 1.5 : iw;   // text width beside the QR code
  doc.setFont('helvetica', 'bold');

  // "PART NAME" caption
  if (!m.compact) {
    const sz = m.brand * MM;
    doc.setFontSize(m.brand);
    doc.setTextColor(113, 102, 141);
    T('PART NAME', m.pad, y + sz * .8);
    y += sz * 1.15 + .8;
  }

  // part name (max 2 lines, 1 when compact)
  {
    const sz = m.name * MM;
    const maxLines = m.compact ? 1 : 2;
    doc.setFontSize(m.name);
    doc.setTextColor(0);

    let lines = doc.splitTextToSize(L1(r.partName || '—'), textW);
    if (lines.length > maxLines) {
      lines = lines.slice(0, maxLines);
      lines[maxLines - 1] = fit(lines[maxLines - 1] + '...', textW);
    }
    lines.forEach((line, i) => T(line, m.pad, y + sz * .85 + i * sz * 1.15));
    y += lines.length * sz * 1.15 + 1;
  }

  // part number + category rows, divider line, serial number line
  {
    const sz = m.row * MM;
    doc.setFontSize(m.row);
    doc.setFont('helvetica', 'normal');
    const keyW = Math.max(doc.getTextWidth('Part number:'), doc.getTextWidth('Category:')) + 1.5;

    [['Part number:', r.partNumber], ['Category:', r.category]].forEach(([key, value]) => {
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(85);
      T(key, m.pad, y + sz * .9);

      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0);
      T(fit(value || '—', textW - keyW), m.pad + keyW, y + sz * .9);
      y += sz * 1.3;
    });

    // divider (below the QR code if it is taller than the text)
    y = Math.max(y, m.pad + qrSide) + 1;
    doc.setFillColor(187);
    R(m.pad, y, iw, .2, 'F');
    y += 1;

    // serial number
    const ss = m.name * MM;
    doc.setFontSize(m.name);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(0);
    T(fit('Serial number: ' + (r.serialNumber || '—'), iw), m.pad, y + ss * .85);
    y += ss * 1.15 + .6;
  }

  // barcode fills the remaining height
  const barH = Math.max(2, H - m.pad - y);
  if (code) {
    try {
      const bits = barBits(code);
      const moduleW = iw / bits.length;
      doc.setFillColor(0);

      for (let i = 0; i < bits.length;) {
        if (bits[i] === '1') {
          let j = i;
          while (j < bits.length && bits[j] === '1') j++;
          R(m.pad + i * moduleW, y, (j - i) * moduleW, barH, 'F');
          i = j;
        } else i++;
      }
    } catch (e) {
      doc.setFontSize(6);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(180, 35, 24);
      T('Barcode unavailable for this value', m.pad, y + 3);
    }
  } else {
    doc.setFontSize(6);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(119);
    T('No serial number, barcode skipped', m.pad, y + 3);
  }

  // thin border
  if ($('border').checked) {
    doc.setDrawColor(68);
    doc.setLineWidth(.2);
    R(.1, .1, W - .2, H - .2, 'S');
  }
}

async function exportPDF() {
  const s = applySettings();
  const { all } = expand();
  if (!all.length) return;

  await loadLib('jspdf');
  const { jsPDF } = window.jspdf;
  const m = metrics(s);
  const roll = s.mode === 'roll';
  let doc;

  if (roll) {
    // one label per page, page = label size (swapped when rotated)
    const rot = s.orient === 'portrait';
    const pw = rot ? s.h : s.w;
    const ph = rot ? s.w : s.h;
    const orientation = pw > ph ? 'landscape' : 'portrait';

    doc = new jsPDF({ unit: 'mm', format: [pw, ph], orientation });
    all.forEach((r, i) => {
      if (i) doc.addPage([pw, ph], orientation);
      pdfLabel(doc, 0, 0, r, s, m, rot);
    });
  } else {
    // A4 sticker sheet: labels laid out in a grid
    const fit = pageFit(s);
    if (!fit) throw new Error(MSG_TOO_BIG);

    doc = new jsPDF({ unit: 'mm', format: 'a4', orientation: s.orient });
    const perPage = fit.cols * fit.rows;

    all.forEach((r, i) => {
      if (i && i % perPage === 0) doc.addPage('a4', s.orient);
      const k = i % perPage;
      const x = s.margin + (k % fit.cols) * (s.w + s.gap);
      const y = s.margin + Math.floor(k / fit.cols) * (s.h + s.gap);
      pdfLabel(doc, x, y, r, s, m, false);
    });
  }

  doc.save(`labels-${stamp()}.pdf`);
  setExp(
    `PDF saved: ${all.length} label(s), ${doc.getNumberOfPages()} page(s). ` +
    (roll && s.orient === 'portrait' ? MSG_ROTATED : ''),
    'ok'
  );
}




/* ==========================================================
   12. EXPORT: WORD (.docx)
   ========================================================== */
const TW = mm => Math.round(mm * 56.6929);          // mm -> twips (Word units)
const PX = mm => Math.max(1, Math.round(mm * 3.7795)); // mm -> pixels

// Text measuring (so long text can be cut to fit)
const measureCtx = document.createElement('canvas').getContext('2d');

function textWidth(text, pt, bold) {
  measureCtx.font = `${bold ? 'bold ' : ''}${pt}pt Arial`;
  return measureCtx.measureText(text).width / 3.7795;   // px -> mm
}

function fitWidth(text, pt, bold, max) {
  text = String(text ?? '');
  if (textWidth(text, pt, bold) <= max) return text;
  while (text.length > 1 && textWidth(text + '...', pt, bold) > max) text = text.slice(0, -1);
  return text + '...';
}

// Build the Word paragraphs/tables for one label
function wordLabel(D, r, s, m) {
  const iw = s.w - 2 * m.pad;
  const out = [];     // final content of the label
  const top = [];     // text block (goes beside the QR code)
  let used = 0;       // height used so far (mm)
  let qr = null;

  if (r.serialNumber) {
    try { qr = qrPng(r.serialNumber); } catch (e) { /* QR skipped */ }
  }
  const qrSide = qr ? m.qr : 0;
  const textW = qr ? iw - qrSide - 1.5 : iw;   // text width beside the QR code

  // spacing with an exact line height
  const spacing = (line, before = 0, after = 0) => ({
    line: TW(line),
    lineRule: D.LineRuleType.EXACT,
    before: TW(before),
    after: TW(after)
  });
  const run = (text, pt, opts = {}) => new D.TextRun({ text, font: 'Arial', size: Math.round(pt * 2), ...opts });

  // "PART NAME" caption
  if (!m.compact) {
    const lh = m.brand * MM * 1.25;
    top.push(new D.Paragraph({
      spacing: spacing(lh, 0, .8),
      children: [run('PART NAME', m.brand, { bold: true, color: '71668D' })]
    }));
    used += lh + .8;
  }

  // part name
  {
    const lh = m.name * MM * 1.25;
    const text = String(r.partName || '—');
    const maxLines = m.compact ? 1 : 2;
    const shown = maxLines === 1
      ? fitWidth(text, m.name, true, textW)
      : fitWidth(text, m.name, true, textW * 1.8);
    const lines = Math.min(maxLines, textWidth(shown, m.name, true) > textW * .98 ? 2 : 1);

    top.push(new D.Paragraph({
      spacing: spacing(lh, 0, 1),
      children: [run(shown, m.name, { bold: true })]
    }));
    used += lines * lh + 1;
  }

  // part number + category rows, QR table, serial number line
  {
    const lh = m.row * MM * 1.4;
    const keyW = Math.max(textWidth('Part number:', m.row), textWidth('Category:', m.row)) + 1.5;

    [['Part number:', r.partNumber], ['Category:', r.category]].forEach(([key, value]) => {
      top.push(new D.Paragraph({
        spacing: spacing(lh),
        tabStops: [{ type: D.TabStopType.LEFT, position: TW(keyW) }],
        children: [
          run(key, m.row, { color: '555555' }),
          new D.TextRun({
            children: [new D.Tab(), fitWidth(value || '—', m.row, true, textW - keyW)],
            font: 'Arial',
            size: Math.round(m.row * 2),
            bold: true
          })
        ]
      }));
      used += lh;
    });

    if (qr) {
      // invisible 2-column table: text on the left, QR code in the top-right corner
      const none = { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
      const noBorders = { top: none, bottom: none, left: none, right: none };
      const zero = { top: 0, bottom: 0, left: 0, right: 0 };

      out.push(new D.Table({
        width: { size: TW(iw), type: D.WidthType.DXA },
        columnWidths: [TW(iw - qrSide), TW(qrSide)],
        layout: D.TableLayoutType.FIXED,
        borders: { ...noBorders, insideHorizontal: none, insideVertical: none },
        rows: [new D.TableRow({
          height: { value: TW(qrSide), rule: D.HeightRule.ATLEAST },
          children: [
            new D.TableCell({
              width: { size: TW(iw - qrSide), type: D.WidthType.DXA },
              borders: noBorders,
              margins: zero,
              children: top
            }),
            new D.TableCell({
              width: { size: TW(qrSide), type: D.WidthType.DXA },
              borders: noBorders,
              margins: zero,
              children: [new D.Paragraph({
                spacing: { before: 0, after: 0 },
                children: [new D.ImageRun({ data: qr, transformation: { width: PX(qrSide), height: PX(qrSide) } })]
              })]
            })
          ]
        })]
      }));
      used = Math.max(used, qrSide);
    } else {
      out.push(...top);
    }

    // serial number line with a thin divider above
    const serialLine = m.name * MM * 1.25;
    out.push(new D.Paragraph({
      spacing: spacing(serialLine, 1, 0),
      border: { top: { style: D.BorderStyle.SINGLE, size: 4, color: 'BBBBBB', space: 1 } },
      children: [run(fitWidth('Serial number: ' + (r.serialNumber || '—'), m.name, true, iw), m.name, { bold: true })]
    }));
    used += serialLine + 1 + .7;
  }

  // barcode (image) fills the remaining height
  const barH = Math.max(3, s.h - 2 * m.pad - used - 1.4);
  if (r.serialNumber) {
    try {
      out.push(new D.Paragraph({
        spacing: { before: TW(.6), after: 0 },
        children: [new D.ImageRun({ data: barPng(r.serialNumber), transformation: { width: PX(iw), height: PX(barH) } })]
      }));
      return out;
    } catch (e) { /* fall through to the text message below */ }
  }

  out.push(new D.Paragraph({
    spacing: spacing(m.row * MM * 1.4, .6),
    children: [run(
      r.serialNumber ? 'Barcode unavailable for this value' : 'No serial number, barcode skipped',
      6,
      { color: '777777' }
    )]
  }));
  return out;
}

async function exportDOCX() {
  const s = applySettings();
  const { all } = expand();
  if (!all.length) return;

  await loadLib('docx');
  const D = window.docx;
  const m = metrics(s);
  const roll = s.mode === 'roll';

  // shared border / spacing definitions
  const none = { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const line = { style: D.BorderStyle.SINGLE, size: 4, color: '8D86A0' };
  const noBorders = { top: none, bottom: none, left: none, right: none };
  const lineBorders = { top: line, bottom: line, left: line, right: line };
  const padMargins = { top: TW(m.pad), bottom: TW(m.pad), left: TW(m.pad), right: TW(m.pad) };

  // tiny empty paragraph (spacer; Word needs a paragraph after tables / in empty cells)
  const tiny = () => new D.Paragraph({
    spacing: { line: 20, lineRule: D.LineRuleType.EXACT, before: 0, after: 0 },
    children: [new D.TextRun({ text: '', size: 2 })]
  });

  let sections;

  if (roll) {
    // one label per page; page size = label size
    sections = all.map(r => ({
      properties: {
        page: {
          size: {
            width:  TW(s.h < s.w ? s.h : s.w),
            height: TW(s.h < s.w ? s.w : s.h),
            orientation: s.w > s.h ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT
          },
          margin: { top: TW(m.pad), bottom: TW(m.pad), left: TW(m.pad), right: TW(m.pad), header: 0, footer: 0 }
        }
      },
      children: wordLabel(D, r, s, m)
    }));

  } else {
    // A4 sticker sheet: one table per page, labels in cells
    const fit = pageFit(s);
    if (!fit) throw new Error(MSG_TOO_BIG);

    const rowsPerPage = Math.max(1, Math.floor((fit.uh - 1.5 + s.gap) / (s.h + s.gap)));
    const perPage = fit.cols * rowsPerPage;

    // column widths (label, gap, label, gap, ...)
    const widths = [];
    for (let c = 0; c < fit.cols; c++) {
      if (c && s.gap > 0) widths.push(TW(s.gap));
      widths.push(TW(s.w));
    }

    const children = [];

    for (let p = 0; p < all.length; p += perPage) {
      const chunk = all.slice(p, p + perPage);
      const rows = [];

      for (let rr = 0; rr * fit.cols < chunk.length; rr++) {

        // spacer row between label rows
        if (rr && s.gap > 0) {
          rows.push(new D.TableRow({
            height: { value: TW(s.gap), rule: D.HeightRule.EXACT },
            cantSplit: true,
            children: widths.map(w => new D.TableCell({
              width: { size: w, type: D.WidthType.DXA },
              borders: noBorders,
              margins: { top: 0, bottom: 0, left: 0, right: 0 },
              children: [tiny()]
            }))
          }));
        }

        // label row
        const cells = [];
        for (let c = 0; c < fit.cols; c++) {
          // spacer column between labels
          if (c && s.gap > 0) {
            cells.push(new D.TableCell({
              width: { size: TW(s.gap), type: D.WidthType.DXA },
              borders: noBorders,
              children: [tiny()]
            }));
          }
          const r = chunk[rr * fit.cols + c];
          cells.push(new D.TableCell({
            width: { size: TW(s.w), type: D.WidthType.DXA },
            margins: padMargins,
            borders: r && $('border').checked ? lineBorders : noBorders,
            children: r ? wordLabel(D, r, s, m) : [tiny()]
          }));
        }
        rows.push(new D.TableRow({
          height: { value: TW(s.h), rule: D.HeightRule.EXACT },
          cantSplit: true,
          children: cells
        }));
      }

      // page break before every page except the first
      if (p) {
        children.push(new D.Paragraph({
          pageBreakBefore: true,
          spacing: { line: 20, lineRule: D.LineRuleType.EXACT, before: 0, after: 0 },
          children: [new D.TextRun({ text: '', size: 2 })]
        }));
      }

      children.push(new D.Table({
        width: { size: widths.reduce((a, b) => a + b, 0), type: D.WidthType.DXA },
        columnWidths: widths,
        layout: D.TableLayoutType.FIXED,
        borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
        rows
      }));
    }

    children.push(tiny());

    sections = [{
      properties: {
        page: {
          size: {
            width: TW(210),
            height: TW(297),
            orientation: s.orient === 'landscape' ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT
          },
          margin: { top: TW(s.margin), bottom: TW(s.margin), left: TW(s.margin), right: TW(s.margin), header: 0, footer: 0 }
        }
      },
      children
    }];
  }

  const doc = new D.Document({ creator: 'Part Controller', title: 'Labels', sections });
  saveBlob(await D.Packer.toBlob(doc), `labels-${stamp()}.docx`);

  setExp(
    `Word file saved: ${all.length} label(s). ` +
    (roll && s.orient === 'portrait' ? MSG_ROTATED : ''),
    'ok'
  );
}


/* ==========================================================
   13. EVENT WIRING & FIRST RENDER
   ========================================================== */

// Any settings change redraws the preview
['size', 'cw', 'ch', 'orient', 'mode', 'margin', 'gap', 'border'].forEach(id =>
  $(id).addEventListener('input', render)
);

// Save buttons (PDF / Word): disable while working, show errors in the status line
[
  ['savePdf',  exportPDF],
  ['saveDocx', exportDOCX]
].forEach(([id, fn]) =>
  $(id).addEventListener('click', async () => {
    const btn = $(id);
    btn.disabled = true;
    setExp('Preparing file…');
    try {
      await fn();
    } catch (e) {
      setExp(e.message || 'Could not create the file.', 'error');
    } finally {
      btn.disabled = !expand().all.length;
    }
  })
);

// Print
$('print').addEventListener('click', () => window.print());

// Clear everything
$('clear').addEventListener('click', () => {
  bulkItems = [];
  manualItems = [];
  bulkActive = false;
  renderManualList();
  $('manualStatus').textContent = '';
  setBulkStatus('Preview cleared. Re-import to add the file again.');
  render();
});

// First draw
render();
