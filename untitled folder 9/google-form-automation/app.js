// ── State ───────────────────────────────────────────────────
let parsedQuestions = [];

// ── DOM refs ─────────────────────────────────────────────────
const dropZone     = document.getElementById('drop-zone');
const fileInput    = document.getElementById('file-input');
const fileInfo     = document.getElementById('file-info');
const generateBtn  = document.getElementById('generate-btn');
const previewEmpty = document.getElementById('preview-empty');
const previewCont  = document.getElementById('preview-container');
const previewList  = document.getElementById('preview-list');
const previewTitle = document.getElementById('preview-title');
const previewDesc  = document.getElementById('preview-desc');
const statsRow     = document.getElementById('stats-row');
const outputSec    = document.getElementById('output-section');
const codeOutput   = document.getElementById('code-output');
const qCountEl     = document.getElementById('q-count');
const typeCountEl  = document.getElementById('type-count');
const reqCountEl   = document.getElementById('req-count');

// ── Toggle switches ──────────────────────────────────────────
document.querySelectorAll('.toggle').forEach(t => {
  t.addEventListener('click', () => t.classList.toggle('on'));
});

// ── File drop / click ────────────────────────────────────────
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});

// ── Handle uploaded file ─────────────────────────────────────
function handleFile(file) {
  const name = file.name.toLowerCase();
  if (!name.endsWith('.csv') && !name.endsWith('.xlsx') && !name.endsWith('.xls')) {
    showToast('⚠️ Please upload a .csv or .xlsx file');
    return;
  }

  if (name.endsWith('.csv')) {
    const reader = new FileReader();
    reader.onload = e => {
      parsedQuestions = parseCSV(e.target.result);
      onFileLoaded(file.name, parsedQuestions.length);
    };
    reader.readAsText(file);
  } else {
    // XLSX using SheetJS
    const reader = new FileReader();
    reader.onload = e => {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      parsedQuestions = parseRows(rows);
      onFileLoaded(file.name, parsedQuestions.length);
    };
    reader.readAsArrayBuffer(file);
  }
}

function onFileLoaded(filename, count) {
  fileInfo.style.display = 'flex';
  fileInfo.innerHTML = `<span>✅</span> <strong>${filename}</strong> — ${count} question${count !== 1 ? 's' : ''} found`;
  generateBtn.disabled = count === 0;
  if (count === 0) {
    showToast('⚠️ No questions found. Check your column headers.');
    return;
  }
  renderPreview();
  showToast(`✅ ${count} questions loaded!`);
}

// ── CSV parser ────────────────────────────────────────────────
function parseCSV(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  if (lines.length < 2) return [];
  const rows = lines.map(line => {
    // Handle quoted commas
    const result = [];
    let cur = '', inQ = false;
    for (let c of line) {
      if (c === '"') { inQ = !inQ; }
      else if (c === ',' && !inQ) { result.push(cur.trim()); cur = ''; }
      else cur += c;
    }
    result.push(cur.trim());
    return result;
  });
  return parseRows(rows);
}

function parseRows(rows) {
  if (rows.length < 2) return [];
  // Skip header row
  const questions = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0] || String(row[0]).trim() === '') continue;
    questions.push({
      question : String(row[0] || '').trim(),
      type     : String(row[1] || 'short_answer').trim().toLowerCase(),
      options  : String(row[2] || '').trim(),
      required : String(row[3] || '').toUpperCase() === 'TRUE',
      helpText : String(row[4] || '').trim(),
    });
  }
  return questions;
}

// ── Render live preview ──────────────────────────────────────
function renderPreview() {
  const title = document.getElementById('form-title').value || 'My Generated Form';
  const desc  = document.getElementById('form-desc').value || '';

  previewTitle.textContent = title;
  previewDesc.textContent  = desc;
  previewList.innerHTML    = '';

  parsedQuestions.forEach((q, i) => {
    const el = document.createElement('div');
    el.className = 'preview-q';
    el.style.animationDelay = (i * 0.05) + 's';
    el.innerHTML = buildQuestionHTML(q, i + 1);
    previewList.appendChild(el);
  });

  previewEmpty.style.display = 'none';
  previewCont.style.display  = 'block';

  // Stats
  const types = new Set(parsedQuestions.map(q => q.type)).size;
  const req   = parsedQuestions.filter(q => q.required).length;
  qCountEl.textContent    = parsedQuestions.length;
  typeCountEl.textContent = types;
  reqCountEl.textContent  = req;
  statsRow.style.cssText = 'display:flex;grid-column:1/-1;gap:16px;';
}

function buildQuestionHTML(q, num) {
  const star = q.required ? '<span class="required-star">*</span>' : '';
  const help  = q.helpText ? `<div class="preview-q-help">${q.helpText}</div>` : '';
  const opts  = q.options ? q.options.split(',').map(o => o.trim()).filter(Boolean) : [];

  let input = '';
  switch (q.type) {
    case 'short_answer':
      input = `<div class="mock-input">Short answer text</div>`;
      break;
    case 'paragraph':
      input = `<div class="mock-input mock-textarea">Long answer text...</div>`;
      break;
    case 'multiple_choice':
      input = opts.map(o => `<div class="mock-option"><div class="mock-radio"></div>${o}</div>`).join('');
      break;
    case 'checkbox':
      input = opts.map(o => `<div class="mock-option"><div class="mock-check"></div>${o}</div>`).join('');
      break;
    case 'dropdown':
      input = `<div class="mock-input">${opts[0] || 'Select an option'} ▾</div>`;
      break;
    case 'linear_scale': {
      const parts = opts;
      const min = parseInt(parts[0]) || 1;
      const max = parseInt(parts[1]) || 5;
      const minL = parts[2] || ''; const maxL = parts[3] || '';
      const btns = Array.from({length: max - min + 1}, (_,i) =>
        `<div class="mock-scale-btn">${min + i}</div>`).join('');
      input = `<div class="mock-scale">${btns}</div>
               <div class="scale-labels"><span>${minL}</span><span>${maxL}</span></div>`;
      break;
    }
    case 'date':
      input = `<div class="mock-input">📅 MM / DD / YYYY</div>`;
      break;
    case 'time':
      input = `<div class="mock-input">🕐 HH : MM</div>`;
      break;
    default:
      input = `<div class="mock-input">Answer...</div>`;
  }

  return `
    <div class="preview-q-type">${q.type.replace(/_/g,' ')}</div>
    <div class="preview-q-text">${num}. ${q.question}${star}</div>
    ${help}${input}`;
}

// ── Generate Script ──────────────────────────────────────────
generateBtn.addEventListener('click', () => {
  const title   = document.getElementById('form-title').value || 'My Generated Form';
  const desc    = document.getElementById('form-desc').value || '';
  const tabName = document.getElementById('sheet-tab').value || 'Questions';
  const email   = document.getElementById('notify-email').value || 'your@email.com';
  const sendEmail = document.getElementById('toggle-email').classList.contains('on');
  const collectEmail = document.getElementById('toggle-collect').classList.contains('on');

  // Re-render preview with current title
  renderPreview();

  const code = generateCode(title, desc, tabName, email, sendEmail, collectEmail);
  codeOutput.textContent = code;

  outputSec.style.display = 'block';
  outputSec.scrollIntoView({ behavior: 'smooth', block: 'start' });
  showToast('✅ Script generated! Copy it into Apps Script.');
});

function generateCode(title, desc, tabName, email, sendEmail, collectEmail) {
  return `// ============================================================
// Code.gs — Generated by Google Form Automation Tool
// Form Title: ${title}
// Generated: ${new Date().toLocaleString()}
// ============================================================

var CONFIG = {
  QUESTIONS_SHEET_ID: "YOUR_SPREADSHEET_ID_HERE",
  QUESTIONS_TAB_NAME: "${tabName}",
  FORM_TITLE: "${title}",
  FORM_DESCRIPTION: "${desc}",
  SEND_EMAIL_ON_COMPLETION: ${sendEmail},
  COLLECT_EMAIL: ${collectEmail},
  NOTIFICATION_EMAIL: "${email}",
};

function generateForm() {
  Logger.log("🚀 Starting form generation...");
  try {
    var questions = readQuestionsFromSheet();
    if (questions.length === 0) { Logger.log("❌ No questions found."); return; }
    Logger.log("✅ Found " + questions.length + " questions.");

    var form = FormApp.create(CONFIG.FORM_TITLE);
    form.setDescription(CONFIG.FORM_DESCRIPTION);
    form.setCollectEmail(CONFIG.COLLECT_EMAIL);

    for (var i = 0; i < questions.length; i++) {
      addQuestionToForm(form, questions[i]);
    }

    var ss = SpreadsheetApp.create(CONFIG.FORM_TITLE + " — Responses");
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

    Logger.log("🎉 DONE!");
    Logger.log("🔗 Share: " + form.getPublishedUrl());
    Logger.log("✏️  Edit:  " + form.getEditUrl());
    Logger.log("📊 Sheet: " + ss.getUrl());

    if (CONFIG.SEND_EMAIL_ON_COMPLETION) {
      GmailApp.sendEmail(CONFIG.NOTIFICATION_EMAIL,
        "✅ Form Ready — " + CONFIG.FORM_TITLE,
        "Form link: " + form.getPublishedUrl() + "\\nResponses: " + ss.getUrl());
    }
  } catch(e) { Logger.log("❌ " + e.message); }
}

function readQuestionsFromSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.QUESTIONS_SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.QUESTIONS_TAB_NAME);
  if (!sheet) throw new Error('Tab "' + CONFIG.QUESTIONS_TAB_NAME + '" not found.');
  var data = sheet.getDataRange().getValues();
  var questions = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0] || String(r[0]).trim() === "") continue;
    questions.push({
      question: String(r[0]).trim(),
      type:     String(r[1]).trim().toLowerCase(),
      options:  r[2] ? String(r[2]).trim() : "",
      required: String(r[3]).toUpperCase() === "TRUE",
      helpText: r[4] ? String(r[4]).trim() : "",
    });
  }
  return questions;
}

function addQuestionToForm(form, q) {
  var item, opts;
  function parseOpts(s) { return s.split(",").map(function(x){return x.trim();}).filter(Boolean); }
  switch(q.type) {
    case "short_answer":
      item = form.addTextItem(); break;
    case "paragraph":
      item = form.addParagraphTextItem(); break;
    case "multiple_choice":
      item = form.addMultipleChoiceItem();
      if (q.options) item.setChoiceValues(parseOpts(q.options));
      break;
    case "checkbox":
      item = form.addCheckboxItem();
      if (q.options) item.setChoiceValues(parseOpts(q.options));
      break;
    case "dropdown":
      item = form.addListItem();
      if (q.options) item.setChoiceValues(parseOpts(q.options));
      break;
    case "linear_scale":
      item = form.addScaleItem();
      opts = parseOpts(q.options);
      item.setBounds(parseInt(opts[0])||1, parseInt(opts[1])||5);
      item.setLabels(opts[2]||"", opts[3]||"");
      break;
    case "date": item = form.addDateItem(); break;
    case "time": item = form.addTimeItem(); break;
    default:     item = form.addTextItem(); break;
  }
  item.setTitle(q.question);
  item.setRequired(q.required);
  if (q.helpText) item.setHelpText(q.helpText);
}`;
}

// ── Tabs ─────────────────────────────────────────────────────
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
  });
});

// ── Copy button ───────────────────────────────────────────────
document.getElementById('copy-btn').addEventListener('click', () => {
  navigator.clipboard.writeText(codeOutput.textContent).then(() => {
    showToast('✅ Copied to clipboard!');
  });
});

// ── Live preview on title/desc change ────────────────────────
document.getElementById('form-title').addEventListener('input', () => {
  if (parsedQuestions.length) renderPreview();
});
document.getElementById('form-desc').addEventListener('input', () => {
  if (parsedQuestions.length) renderPreview();
});

// ── Toast ─────────────────────────────────────────────────────
function showToast(msg) {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.style.display = 'block';
  setTimeout(() => { toast.style.display = 'none'; }, 3000);
}
