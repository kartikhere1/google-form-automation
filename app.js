// ── State ─────────────────────────────────────────────────────
let excelHeaders = [];   // Column headers from Excel row 1
let excelRows    = [];   // Data rows (row 2 onwards)
let columnMap    = {};   // { excelHeader: formQuestionTitle }

// ── DOM ────────────────────────────────────────────────────────
const dropZone      = document.getElementById('drop-zone');
const fileInput     = document.getElementById('file-input');
const fileSuccess   = document.getElementById('file-success');
const statsStrip    = document.getElementById('stats-strip');
const statRows      = document.getElementById('stat-rows');
const statCols      = document.getElementById('stat-cols');
const mapperSection = document.getElementById('mapper-section');
const colMapList    = document.getElementById('col-map-list');
const previewSection= document.getElementById('preview-section');
const previewTbody  = document.getElementById('preview-tbody');
const previewThead  = document.getElementById('preview-thead');
const outputSection = document.getElementById('output-section');
const codeOutput    = document.getElementById('code-output');
const generateBtn   = document.getElementById('generate-btn');

// ── File Upload ───────────────────────────────────────────────
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => { if (fileInput.files[0]) handleFile(fileInput.files[0]); });

function handleFile(file) {
  const name = file.name.toLowerCase();
  if (!name.endsWith('.csv') && !name.endsWith('.xlsx') && !name.endsWith('.xls')) {
    showToast('⚠️ Please upload a .csv or .xlsx file'); return;
  }
  if (name.endsWith('.csv')) {
    const reader = new FileReader();
    reader.onload = e => processRows(parseCSV(e.target.result), file.name);
    reader.readAsText(file);
  } else {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      processRows(XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }), file.name);
    };
    reader.readAsArrayBuffer(file);
  }
}

function parseCSV(text) {
  return text.split('\n').map(line => {
    const result = []; let cur = '', inQ = false;
    for (const c of line) {
      if (c === '"') { inQ = !inQ; }
      else if (c === ',' && !inQ) { result.push(cur.trim()); cur = ''; }
      else cur += c;
    }
    result.push(cur.trim());
    return result;
  }).filter(r => r.some(c => c !== ''));
}

function processRows(rows, filename) {
  if (!rows || rows.length < 2) { showToast('⚠️ File must have headers + at least 1 data row'); return; }

  excelHeaders = rows[0].map(h => String(h).trim()).filter(Boolean);
  excelRows    = rows.slice(1).filter(r => r.some(c => c !== '' && c !== null && c !== undefined));
  columnMap    = {};

  // Show file success
  fileSuccess.textContent = `✅  ${filename}  —  ${excelRows.length} responses · ${excelHeaders.length} columns detected`;
  fileSuccess.classList.add('show');

  // Stats strip
  statRows.textContent = excelRows.length;
  statCols.textContent = excelHeaders.length;
  statsStrip.classList.add('show');

  // Build column mapper
  buildMapper();

  // Show data preview
  buildPreview();

  // Validate and enable
  runValidation();
  showToast(`✅ ${excelRows.length} responses loaded!`);
}

// ── Column Mapper ─────────────────────────────────────────────
function buildMapper() {
  colMapList.innerHTML = '';
  excelHeaders.forEach(header => {
    columnMap[header] = header; // default: same name

    const row = document.createElement('div');
    row.className = 'col-row mapped';
    row.dataset.col = header;

    // Sample values from first 3 rows
    const samples = excelRows.slice(0, 3)
      .map(r => r[excelHeaders.indexOf(header)])
      .filter(v => v !== '' && v !== null && v !== undefined)
      .slice(0, 2).join(', ');

    row.innerHTML = `
      <div class="col-excel">
        <span>${header}</span>
        <small>${samples || '—'}</small>
      </div>
      <div class="col-arrow">→</div>
      <input class="col-input" placeholder="Form question title (leave blank to skip)"
             value="${header}" data-col="${header}" />
    `;
    colMapList.appendChild(row);

    row.querySelector('.col-input').addEventListener('input', e => {
      const val = e.target.value.trim();
      columnMap[header] = val;
      row.classList.toggle('mapped', val !== '');
      row.classList.toggle('skipped', val === '');
      runValidation();
    });
  });

  mapperSection.classList.add('show');
}

// ── Auto-Map ──────────────────────────────────────────────────
document.getElementById('auto-map-btn').addEventListener('click', () => {
  // Auto-map: use exact column header as form question title
  colMapList.querySelectorAll('.col-input').forEach(input => {
    const col = input.dataset.col;
    input.value = col;
    columnMap[col] = col;
    input.closest('.col-row').classList.add('mapped');
    input.closest('.col-row').classList.remove('skipped');
  });
  runValidation();
  showToast('✅ Auto-mapped all columns');
});

// ── Data Preview ──────────────────────────────────────────────
function buildPreview() {
  const previewCols = excelHeaders.slice(0, 5); // show max 5 cols
  const more = excelHeaders.length > 5 ? ` + ${excelHeaders.length - 5} more` : '';

  previewThead.innerHTML = '<tr>' + previewCols.map(h => `<th>${h}</th>`).join('') + '</tr>';
  previewTbody.innerHTML = '';

  excelRows.slice(0, 5).forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = previewCols.map((h, i) => {
      const val = row[excelHeaders.indexOf(h)];
      return `<td>${val !== undefined && val !== '' ? val : '—'}</td>`;
    }).join('');
    previewTbody.appendChild(tr);
  });

  document.getElementById('preview-more').textContent =
    `Showing first 5 of ${excelRows.length} rows · ${excelHeaders.length} columns${more}`;

  previewSection.classList.add('show');
}

// ── Validation ────────────────────────────────────────────────
function runValidation() {
  const mapped    = Object.values(columnMap).filter(v => v !== '').length;
  const skipped   = excelHeaders.length - mapped;
  const hasForm   = document.getElementById('form-url').value.trim() !== '';
  const hasSheet  = document.getElementById('sheet-id').value.trim() !== '';

  document.getElementById('val-responses').textContent  = `${excelRows.length} responses ready to import`;
  document.getElementById('val-mapped').textContent     = `${mapped} of ${excelHeaders.length} columns mapped`;
  document.getElementById('val-skipped').textContent    = skipped > 0 ? `${skipped} column(s) will be skipped` : 'All columns mapped';
  document.getElementById('val-skipped').style.color    = skipped > 0 ? 'var(--orange)' : 'var(--green)';

  document.getElementById('val-form-icon').textContent  = hasForm  ? '✅' : '⚠️';
  document.getElementById('val-sheet-icon').textContent = hasSheet ? '✅' : '⚠️';

  generateBtn.disabled = excelRows.length === 0 || mapped === 0;
}

document.getElementById('form-url').addEventListener('input', runValidation);
document.getElementById('sheet-id').addEventListener('input', runValidation);

// ── Generate Script ───────────────────────────────────────────
generateBtn.addEventListener('click', () => {
  const formUrl   = document.getElementById('form-url').value.trim();
  const sheetId   = document.getElementById('sheet-id').value.trim();
  const sheetName = document.getElementById('sheet-name').value.trim() || 'Sheet1';

  // Build column→question map comment block
  const mapLines = Object.entries(columnMap)
    .filter(([, v]) => v !== '')
    .map(([k, v]) => `  //   "${k}"  →  "${v}"`)
    .join('\n');

  const code = generateScript(formUrl, sheetId, sheetName, mapLines);
  codeOutput.textContent = code;
  outputSection.classList.add('show');
  outputSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  showToast('✅ Script generated! Copy it into Apps Script.');
});

function generateScript(formUrl, sheetId, sheetName, mapComment) {
  const mappedEntries = Object.entries(columnMap).filter(([, v]) => v !== '');
  const colMapJS = mappedEntries
    .map(([k, v]) => `  "${k}": "${v}"`)
    .join(',\n');

  return `// ============================================================
// Code.gs — Bulk Response Importer for Google Forms
// Generated by FormGen · ${new Date().toLocaleString()}
//
// Column → Question mapping:
${mapComment}
// ============================================================

var CONFIG = {
  FORM_URL  : "${formUrl || 'PASTE_YOUR_GOOGLE_FORM_URL_HERE'}",
  SHEET_ID  : "${sheetId || 'PASTE_YOUR_SPREADSHEET_ID_HERE'}",
  SHEET_NAME: "${sheetName}",
};

// Maps Excel column header → exact Google Form question title
var COLUMN_MAP = {
${colMapJS}
};

// ─────────────────────────────────────────────────────────────
// MAIN: importResponses
// Run this once to submit all rows from your sheet as form
// responses. Check the Execution log for progress.
// ─────────────────────────────────────────────────────────────

function importResponses() {
  var form  = FormApp.openByUrl(CONFIG.FORM_URL);
  var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID)
                            .getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    Logger.log("❌ Sheet tab '" + CONFIG.SHEET_NAME + "' not found.");
    return;
  }

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).trim(); });

  // Build a map: form question title → form item
  var items   = form.getItems();
  var itemMap = {};
  items.forEach(function(item) {
    itemMap[item.getTitle()] = item;
  });

  var success = 0, skipped = 0, errors = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // Skip completely empty rows
    if (row.every(function(c){ return c === '' || c === null; })) continue;

    var response = form.createResponse();

    for (var j = 0; j < headers.length; j++) {
      var excelCol   = headers[j];
      var formTitle  = COLUMN_MAP[excelCol];
      var answer     = String(row[j] || '').trim();

      if (!formTitle || answer === '') continue; // skip unmapped or empty

      var item = itemMap[formTitle];
      if (!item) {
        Logger.log("⚠️  Row " + i + ": No form question found for '" + formTitle + "'");
        continue;
      }

      var type = item.getType();

      try {
        if (type === FormApp.ItemType.TEXT) {
          response.withItemResponse(item.asTextItem().createResponse(answer));

        } else if (type === FormApp.ItemType.PARAGRAPH_TEXT) {
          response.withItemResponse(item.asParagraphTextItem().createResponse(answer));

        } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
          response.withItemResponse(item.asMultipleChoiceItem().createResponse(answer));

        } else if (type === FormApp.ItemType.LIST) {
          response.withItemResponse(item.asListItem().createResponse(answer));

        } else if (type === FormApp.ItemType.CHECKBOX) {
          // Checkbox: answer can be comma-separated values
          var choices = answer.split(',').map(function(c){ return c.trim(); });
          response.withItemResponse(item.asCheckboxItem().createResponse(choices));

        } else if (type === FormApp.ItemType.SCALE) {
          response.withItemResponse(item.asScaleItem().createResponse(parseInt(answer)));

        } else if (type === FormApp.ItemType.DATE) {
          var d = new Date(answer);
          response.withItemResponse(item.asDateItem().createResponse(d));

        } else {
          Logger.log("⚠️  Row " + i + ": Unsupported type for '" + formTitle + "'");
        }
      } catch(e) {
        Logger.log("❌ Row " + i + ", col '" + formTitle + "': " + e.message);
        errors++;
      }
    }

    try {
      response.submit();
      success++;
    } catch(e) {
      Logger.log("❌ Row " + i + " submit failed: " + e.message);
      errors++;
      skipped++;
    }
  }

  Logger.log("─────────────────────────────────────────");
  Logger.log("✅ Import complete!");
  Logger.log("   Submitted : " + success);
  Logger.log("   Skipped   : " + skipped);
  Logger.log("   Errors    : " + errors);
  Logger.log("─────────────────────────────────────────");
}`;
}

// ── Tabs ──────────────────────────────────────────────────────
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
  });
});

// ── Copy ──────────────────────────────────────────────────────
document.getElementById('copy-btn').addEventListener('click', () => {
  navigator.clipboard.writeText(codeOutput.textContent)
    .then(() => showToast('✅ Copied to clipboard!'));
});

// ── Toast ─────────────────────────────────────────────────────
function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.style.display = 'block';
  setTimeout(() => { t.style.display = 'none'; }, 3000);
}
