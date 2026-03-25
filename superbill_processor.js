#!/usr/bin/env node
/**
 * Superbill Processor (Node.js)
 * --------------------------------
 * Modes:
 *   node superbill_processor.js <input.xlsx> <output.xlsx>        # interactive CLI
 *   node superbill_processor.js <input.xlsx> <output.xlsx> --yes  # batch (no prompts)
 *   node superbill_processor.js --gui                             # browser GUI server
 */

'use strict';

const ExcelJS        = require('exceljs');
const XLSX           = require('xlsx');
const fs             = require('fs');
const http           = require('http');
const { exec }       = require('child_process');
const os             = require('os');
const path           = require('path');
const readline       = require('readline');

// ── Constants ─────────────────────────────────────────────────────────────────

// Row index of the header inside the Superbill file (0-based).
// Row 6 (0-based) means ExcelJS row 7 (1-based) after skipping 6 rows.
const INPUT_HEADER_ROW = 6; // 0-based; data starts at row INPUT_HEADER_ROW+1

const EXPECTED_EMPTY_COLS = new Set([
  'Primary Carrier',
  'Primary Policy',
  'Secondary Carrier',
  'Secondary Policy',
  'Tertiary Carrier',
  'Tertiary Policy',
  'Clinical Trial',
  'Seq No',
  'Comment',
]);

// Input col index (0-based) → Output col index (0-based)
const COL_MAP = new Map([
  [0,  0],  // A  → A
  [1,  2],  // B  → C
  [2,  3],  // C  → D
  [3,  4],  // D  → E
  [4,  5],  // E  → F
  [5,  6],  // F  → G
  [6,  7],  // G  → H
  [14, 8],  // O  → I
  [15, 9],  // P  → J
  [16, 10], // Q  → K
  [17, 11], // R  → L
  [18, 12], // S  → M
  [19, 13], // T  → N
  [20, 14], // U  → O
  [21, 15], // V  → P
  [29, 23], // AD → X
  [30, 24], // AE → Y
  [31, 25], // AF → Z
]);

// Identity columns: input indices [A, B, O], output indices [A, C, I]
const IDENTITY_INPUT_COLS  = [0, 1, 14];
const IDENTITY_OUTPUT_COLS = [0, 2, 8];

const MAX_OUT_COL = Math.max(...COL_MAP.values()) + 1; // number of output columns needed

// ── Helpers ───────────────────────────────────────────────────────────────────

function normaliseCol(name) {
  return cellStr(name).split(/\s+/).join(' ').trim();
}

/**
 * Normalize an ExcelJS cell value to a plain string.
 * Handles: null/undefined, primitives, Date, rich-text objects, formula result objects.
 */
function cellStr(value) {
  if (value === null || value === undefined) return '';
  // Formula cell: { formula: '...', result: <actual value> }
  if (typeof value === 'object' && 'result' in value) return cellStr(value.result);
  // Rich-text cell: { richText: [{ text: '...' }, ...] }
  if (typeof value === 'object' && Array.isArray(value.richText)) {
    return value.richText.map(r => r.text ?? '').join('').trim();
  }
  // Date
  if (value instanceof Date) return value.toLocaleDateString('en-US');
  return String(value).trim();
}

function isBlank(value) {
  return cellStr(value) === '';
}

/** CLI confirm: prompts stdin and resolves true/false. */
function cliConfirm(question) {
  return new Promise(resolve => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(`\n${question}\nProceed? [y/N] `, answer => {
      rl.close();
      resolve(answer.trim().toLowerCase() === 'y');
    });
  });
}

// ── Core processing ───────────────────────────────────────────────────────────

/**
 * Core processing.
 * @param {string}   inputPath
 * @param {string}   outputPath
 * @param {function} logFn     - logFn(msg: string)
 * @param {function} confirmFn - confirmFn(question: string) => Promise<boolean>
 */
async function run(inputPath, outputPath, logFn, confirmFn) {
  const log     = logFn;
  const confirm = confirmFn;

  // ── 1. Read input ──────────────────────────────────────────────────────────
  log(`Reading input: ${path.basename(inputPath)}`);
  const wbIn = new ExcelJS.Workbook();
  try {
    await wbIn.xlsx.readFile(inputPath);
  } catch (e) {
    log(`ERROR reading input file: ${e.message}`);
    return false;
  }

  const wsIn = wbIn.worksheets[0];
  if (!wsIn) {
    log('ERROR: input file has no worksheets.');
    return false;
  }

  // ExcelJS rows are 1-based. Header is at 1-based row INPUT_HEADER_ROW + 1 (= row 7).
  const headerRowNum = INPUT_HEADER_ROW + 1;
  const headerRow = wsIn.getRow(headerRowNum);
  const headers = [];
  headerRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
    headers[colNum - 1] = normaliseCol(cell.value);
  });

  // Collect data rows (everything after the header that is not all-empty)
  const dataRows = []; // each entry: array of raw cell values (0-based col index)
  wsIn.eachRow((row, rowNum) => {
    if (rowNum <= headerRowNum) return;
    const vals = [];
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      vals[colNum - 1] = cell.value;
    });
    // Skip all-empty rows
    if (vals.every(v => isBlank(v))) return;
    dataRows.push({ vals, excelRowNum: rowNum });
  });

  log(`  Input rows (after cleanup): ${dataRows.length}  |  Columns: ${headers.length}`);

  // ── 2. Detect empty columns vs expected list ───────────────────────────────
  const actualEmpty = new Set();
  for (let ci = 0; ci < headers.length; ci++) {
    const col = headers[ci];
    if (!col) continue;
    if (dataRows.every(r => isBlank(r.vals[ci]))) {
      actualEmpty.add(col);
    }
  }

  const missingFromActual = [...EXPECTED_EMPTY_COLS].filter(c => !actualEmpty.has(c));
  const extraActualEmpty  = [...actualEmpty].filter(c => !EXPECTED_EMPTY_COLS.has(c));

  if (missingFromActual.length || extraActualEmpty.length) {
    log('⚠  DISCREPANCY between expected empty columns and actual empty columns:');
    if (missingFromActual.length) {
      log('   Expected to be empty but are NOT empty:');
      missingFromActual.sort().forEach(c => log(`     • ${c}`));
    }
    if (extraActualEmpty.length) {
      log('   Empty but NOT in the expected list:');
      extraActualEmpty.sort().forEach(c => log(`     • ${c}`));
    }
    log('   (Continuing – actually-empty columns will be removed.)');
  } else {
    log('✓  Empty columns match expected list exactly.');
  }

  // ── 3. Load output file ────────────────────────────────────────────────────
  if (!fs.existsSync(outputPath)) {
    log('ERROR: output file does not exist. Please provide an existing output file.');
    return false;
  }

  log(`Loading output file: ${path.basename(outputPath)}`);
  let wbOut, wsOut, outRows;
  try {
    wbOut   = XLSX.readFile(outputPath, { cellDates: true, sheetStubs: true });
    wsOut   = wbOut.Sheets[wbOut.SheetNames[0]];
    if (!wsOut) throw new Error('No worksheets found in output file.');
    outRows = XLSX.utils.sheet_to_json(wsOut, { header: 1, defval: null });
  } catch (e) {
    log(`ERROR reading output file: ${e.message}`);
    return false;
  }

  // Build existing identity key set (skip header row 0)
  const existingKeys = new Set();
  for (let i = 1; i < outRows.length; i++) {
    const row = outRows[i];
    if (!row || row.every(v => isBlank(v))) continue;
    const key = IDENTITY_OUTPUT_COLS.map(ci => cellStr(row[ci])).join('\x00');
    existingKeys.add(key);
  }

  // ── 4. Duplicate detection ─────────────────────────────────────────────────
  const duplicates  = [];
  const rowsToAdd   = [];

  for (const row of dataRows) {
    const key = IDENTITY_INPUT_COLS.map(ci => cellStr(row.vals[ci]));
    // Skip rows with no identifying information
    if (key.every(v => v === '')) continue;
    const keyStr = key.join('\x00');
    if (existingKeys.has(keyStr)) {
      duplicates.push({ excelRow: row.excelRowNum, key });
    } else {
      rowsToAdd.push(row);
    }
  }

  if (duplicates.length) {
    log('');
    log(`⚠  ${duplicates.length} duplicate row(s) already exist in the output:`);
    log(`    ${'Row'.padEnd(5)}  ${'Date of Service'.padEnd(15)}  ${'Patient Name'.padEnd(30)}  Billing Code`);
    log(`    ${'-'.repeat(5)}  ${'-'.repeat(15)}  ${'-'.repeat(30)}  ${'-'.repeat(15)}`);
    for (const d of duplicates) {
      log(`    ${String(d.excelRow).padEnd(5)}  ${d.key[0].padEnd(15)}  ${d.key[1].padEnd(30)}  ${d.key[2]}`);
    }
    log('');

    if (rowsToAdd.length === 0) {
      log('ℹ  Nothing new to add – all input rows are already in the output.');
      return true;
    }

    log(`  ${rowsToAdd.length} new row(s) would be appended.`);

    const proceed = await confirm(
      `${duplicates.length} duplicate(s) found. Append ${rowsToAdd.length} new row(s)?`
    );
    if (!proceed) {
      log('⛔  Aborted by user. Output file was NOT modified.');
      return false;
    }
  } else {
    log(`✓  No duplicates found. ${rowsToAdd.length} row(s) will be appended.`);
  }

  if (rowsToAdd.length === 0) {
    log('ℹ  Nothing new to add – all input rows are already in the output.');
    return true;
  }

  // ── 5. Backup ──────────────────────────────────────────────────────────────
  const backupPath = path.join(os.tmpdir(), `superbill_backup_${Date.now()}.xlsx`);
  try {
    fs.copyFileSync(outputPath, backupPath);
  } catch (e) {
    log(`ERROR creating backup: ${e.message}`);
    return false;
  }

  // ── 6. Find last non-empty row and append ──────────────────────────────────
  // Find last non-empty row (1-based, for logging)
  let lastNonEmpty = 0;
  for (let i = 0; i < outRows.length; i++) {
    if (outRows[i] && outRows[i].some(v => !isBlank(v))) lastNonEmpty = i + 1;
  }

  // Trim trailing empty rows in the sheet by shrinking !ref
  if (wsOut['!ref']) {
    const range = XLSX.utils.decode_range(wsOut['!ref']);
    if (lastNonEmpty > 0) {
      range.e.r = lastNonEmpty - 1; // 0-based
      wsOut['!ref'] = XLSX.utils.encode_range(range);
    }
  }

  log(`  Appending after row ${lastNonEmpty} (last non-empty row in output).`);

  const appendedKeys = [];
  const newRows = [];
  for (const row of rowsToAdd) {
    const outVals = new Array(MAX_OUT_COL).fill(null);
    for (const [inIdx, outIdx] of COL_MAP.entries()) {
      const v = row.vals[inIdx];
      outVals[outIdx] = isBlank(v) ? null : v;
    }
    newRows.push(outVals);
    const key = IDENTITY_INPUT_COLS.map(ci => cellStr(row.vals[ci])).join('\x00');
    appendedKeys.push(key);
  }
  XLSX.utils.sheet_add_aoa(wsOut, newRows, { origin: -1, cellDates: true });

  // ── 7. Save ────────────────────────────────────────────────────────────────
  try {
    XLSX.writeFile(wbOut, outputPath);
  } catch (e) {
    log(`ERROR saving output file: ${e.message}`);
    log('  Restoring backup…');
    fs.copyFileSync(backupPath, outputPath);
    fs.unlinkSync(backupPath);
    return false;
  }

  log('  Saved. Verifying written data…');

  // ── 8. Verify ──────────────────────────────────────────────────────────────
  const verifiedKeys = new Set();
  try {
    const wbVerify  = XLSX.readFile(outputPath, { cellDates: true });
    const wsVerify  = wbVerify.Sheets[wbVerify.SheetNames[0]];
    const verifyRows = XLSX.utils.sheet_to_json(wsVerify, { header: 1, defval: null });
    for (const r of verifyRows) {
      verifiedKeys.add(IDENTITY_OUTPUT_COLS.map(ci => cellStr(r ? r[ci] : null)).join('\x00'));
    }
  } catch (e) {
    log(`ERROR during verification read: ${e.message}`);
  }

  const missingAfterWrite = appendedKeys.filter(k => !verifiedKeys.has(k));

  if (missingAfterWrite.length) {
    log('');
    log(`⚠  VERIFICATION FAILED: ${missingAfterWrite.length} row(s) not found in output after save.`);

    const keep = await confirm(
      `Verification found ${missingAfterWrite.length} missing row(s).\nKeep the output anyway?`
    );
    if (!keep) {
      log('⛔  Aborted – restoring original output file from backup.');
      fs.copyFileSync(backupPath, outputPath);
      fs.unlinkSync(backupPath);
      return false;
    }
    log('⚠  User chose to keep output despite verification discrepancy.');
  } else {
    log(`✓  Verification passed – all ${appendedKeys.length} row(s) confirmed in output.`);
  }

  fs.unlinkSync(backupPath);
  log('');
  log(`✅  Done.  ${appendedKeys.length} row(s) added to ${path.basename(outputPath)}.`);
  return true;
}

// ── GUI server ────────────────────────────────────────────────────────────────

const HTML_UI = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Superbill Processor</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;background:#f0f2f5;color:#1e1e1e;padding:20px;max-width:860px;margin:auto}
h1{font-size:1.35rem;margin-bottom:16px;color:#0078d4}
.card{background:#fff;border-radius:8px;padding:16px;margin-bottom:14px;box-shadow:0 1px 4px rgba(0,0,0,.1)}
label{display:block;font-weight:600;font-size:.85rem;margin-bottom:4px}
input[type=text]{width:100%;padding:8px 10px;border:1px solid #ccc;border-radius:4px;font-size:.9rem;font-family:inherit}
input[type=text]:focus{outline:none;border-color:#0078d4;box-shadow:0 0 0 2px rgba(0,120,212,.2)}
.row{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
button{padding:8px 20px;border:none;border-radius:4px;cursor:pointer;font-size:.88rem;font-family:inherit}
#processBtn{background:#0078d4;color:#fff}
#processBtn:disabled{background:#aaa;cursor:not-allowed}
#confirmPanel{background:#fff8e1;border:1px solid #f0ad4e;border-radius:8px;padding:14px;margin-bottom:14px;display:none}
#confirmMsg{white-space:pre-wrap;margin-bottom:10px;font-size:.9rem}
#proceedBtn{background:#28a745;color:#fff}
#abortBtn{background:#dc3545;color:#fff}
.log-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
.log-header span{font-weight:600;font-size:.85rem}
#clearBtn{background:#6c757d;color:#fff;font-size:.8rem;padding:4px 12px}
#logBox{font-family:Consolas,monospace;font-size:.82rem;background:#1e1e1e;color:#d4d4d4;padding:12px;border-radius:6px;height:320px;overflow-y:auto;white-space:pre-wrap;word-break:break-all}
#status{font-size:.82rem;color:#555}
.dot{display:inline-block;width:8px;height:8px;border-radius:50%;background:#ccc;margin-right:5px;vertical-align:middle}
.dot.ok{background:#28a745}.dot.busy{background:#ffc107}.dot.err{background:#dc3545}
.browse-btn{background:#e0e0e0;color:#1e1e1e;white-space:nowrap}
.browse-btn:hover{background:#c8c8c8}
</style>
</head>
<body>
<h1>Superbill Processor</h1>

<div class="card">
  <div style="margin-bottom:12px">
    <label for="inputPath">Input Superbill file &mdash; full path</label>
    <div class="row" style="gap:6px">
      <input type="text" id="inputPath" placeholder="e.g. C:\\Users\\...\\superbill.xlsx" style="flex:1">
      <input type="file" id="inputFilePicker" accept=".xlsx,.xls" style="display:none">
      <button class="browse-btn" id="inputBrowseBtn">Browse&hellip;</button>
    </div>
  </div>
  <div>
    <label for="outputPath">Output file &mdash; full path</label>
    <div class="row" style="gap:6px">
      <input type="text" id="outputPath" placeholder="e.g. C:\\Users\\...\\output.xlsx" style="flex:1">
      <input type="file" id="outputFilePicker" accept=".xlsx,.xls" style="display:none">
      <button class="browse-btn" id="outputBrowseBtn">Browse&hellip;</button>
    </div>
  </div>
</div>

<div class="row" style="margin-bottom:14px">
  <button id="processBtn">&#9654;&#xFE0E;&nbsp; Process</button>
  <span id="status"><span class="dot" id="dot"></span><span id="statusText">Connecting&hellip;</span></span>
</div>

<div id="confirmPanel">
  <div id="confirmMsg"></div>
  <div class="row">
    <button id="proceedBtn">&#10004;&nbsp; Proceed</button>
    <button id="abortBtn">&#10008;&nbsp; Abort</button>
  </div>
</div>

<div class="card">
  <div class="log-header"><span>Messages</span><button id="clearBtn">Clear log</button></div>
  <div id="logBox"></div>
</div>

<script>
fetch('/config').then(r=>r.json()).then(cfg=>{
  if(cfg.inputPath)  document.getElementById('inputPath').value  = cfg.inputPath;
  if(cfg.outputPath) document.getElementById('outputPath').value = cfg.outputPath;
});
const processBtn  = document.getElementById('processBtn');
const confirmPanel= document.getElementById('confirmPanel');
const confirmMsg  = document.getElementById('confirmMsg');
const proceedBtn  = document.getElementById('proceedBtn');
const abortBtn    = document.getElementById('abortBtn');
const logBox      = document.getElementById('logBox');
const clearBtn    = document.getElementById('clearBtn');
const dot         = document.getElementById('dot');
const statusText  = document.getElementById('statusText');

function appendLog(msg){
  logBox.textContent += msg + '\\n';
  logBox.scrollTop = logBox.scrollHeight;
}
function setStatus(state, text){
  dot.className = 'dot ' + state;
  statusText.textContent = text;
}
function showConfirm(msg){
  confirmMsg.textContent = msg;
  confirmPanel.style.display = 'block';
  proceedBtn.disabled = false;
  abortBtn.disabled   = false;
}
function hideConfirm(){ confirmPanel.style.display = 'none'; }

const evtSource = new EventSource('/events');
evtSource.onerror = () => setStatus('err','Disconnected \u2013 reload to reconnect');
evtSource.onmessage = e => {
  const d = JSON.parse(e.data);
  if      (d.type === 'ready')   setStatus('ok','Ready');
  else if (d.type === 'log')     appendLog(d.msg);
  else if (d.type === 'confirm') showConfirm(d.msg);
  else if (d.type === 'done') {
    processBtn.disabled = false;
    setStatus('ok','Ready');
    hideConfirm();
  }
};

processBtn.addEventListener('click', async () => {
  const inputPath  = document.getElementById('inputPath').value.trim();
  const outputPath = document.getElementById('outputPath').value.trim();
  if (!inputPath || !outputPath){ alert('Please fill in both file paths.'); return; }
  processBtn.disabled = true;
  setStatus('busy','Processing\u2026');
  appendLog('='.repeat(60));
  const res = await fetch('/process',{
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body: JSON.stringify({ inputPath, outputPath })
  });
  if (!res.ok){
    const err = await res.json();
    appendLog('ERROR: ' + err.error);
    processBtn.disabled = false;
    setStatus('ok','Ready');
  }
});

function setupPicker(pickerId, textId, btnId) {
  const picker = document.getElementById(pickerId);
  const btn    = document.getElementById(btnId);
  btn.addEventListener('click', () => picker.click());
  picker.addEventListener('change', () => {
    const file = picker.files && picker.files[0];
    if (!file) return;
    // file.path is available in some environments (e.g. nw.js); browsers omit it for security
    if (file.path) { document.getElementById(textId).value = file.path; return; }
    // Reconstruct: keep directory from existing value, replace filename
    const existing = document.getElementById(textId).value.trim();
    const sep = existing.includes('/') ? '/' : '\\\\';
    const lastSep = existing.lastIndexOf(sep);
    document.getElementById(textId).value =
      lastSep >= 0 ? existing.substring(0, lastSep + 1) + file.name : file.name;
  });
}
setupPicker('inputFilePicker',  'inputPath',  'inputBrowseBtn');
setupPicker('outputFilePicker', 'outputPath', 'outputBrowseBtn');
async function sendConfirm(proceed){
  proceedBtn.disabled = true;
  abortBtn.disabled   = true;
  await fetch('/confirm',{
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body: JSON.stringify({ proceed })
  });
}
proceedBtn.addEventListener('click', () => sendConfirm(true));
abortBtn.addEventListener  ('click', () => sendConfirm(false));
clearBtn.addEventListener  ('click', () => { logBox.textContent = ''; });
</script>
</body>
</html>`;

async function startGuiServer(initialInput = '', initialOutput = '') {
  let sseClients    = [];
  let pendingConfirm = null;
  let processing     = false;

  function sseLog(msg) {
    const payload = JSON.stringify({ type: 'log', msg });
    sseClients.forEach(r => r.write(`data: ${payload}\n\n`));
  }

  function sseSend(event) {
    const payload = JSON.stringify(event);
    sseClients.forEach(r => r.write(`data: ${payload}\n\n`));
  }

  function sseConfirm(question) {
    return new Promise(resolve => {
      pendingConfirm = { resolve };
      sseSend({ type: 'confirm', msg: question });
    });
  }

  const server = http.createServer((req, res) => {
    // ── GET / ──────────────────────────────────────────────────────────────
    if (req.method === 'GET' && req.url === '/') {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(HTML_UI);

    // ── GET /config ────────────────────────────────────────────────────────
    } else if (req.method === 'GET' && req.url === '/config') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ inputPath: initialInput, outputPath: initialOutput }));

    // ── GET /events (SSE) ──────────────────────────────────────────────────
    } else if (req.method === 'GET' && req.url === '/events') {
      res.writeHead(200, {
        'Content-Type':  'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection':    'keep-alive',
      });
      res.write(': connected\n\n');
      res.write('data: {"type":"ready"}\n\n');
      sseClients.push(res);
      req.on('close', () => { sseClients = sseClients.filter(c => c !== res); });

    // ── POST /process ──────────────────────────────────────────────────────
    } else if (req.method === 'POST' && req.url === '/process') {
      let body = '';
      req.on('data', chunk => { body += chunk; });
      req.on('end', () => {
        if (processing) {
          res.writeHead(409, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Already processing – wait for the current run to finish.' }));
          return;
        }
        let parsed;
        try { parsed = JSON.parse(body); } catch (e) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Invalid JSON.' }));
          return;
        }
        const { inputPath, outputPath } = parsed;
        if (!inputPath || !outputPath) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'inputPath and outputPath are required.' }));
          return;
        }
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true }));
        processing = true;
        run(inputPath, outputPath, sseLog, sseConfirm)
          .then(ok  => { sseSend({ type: 'done', ok });                   processing = false; })
          .catch(err => { sseSend({ type: 'done', ok: false, error: err.message }); processing = false; });
      });

    // ── POST /confirm ──────────────────────────────────────────────────────
    } else if (req.method === 'POST' && req.url === '/confirm') {
      let body = '';
      req.on('data', chunk => { body += chunk; });
      req.on('end', () => {
        let parsed;
        try { parsed = JSON.parse(body); } catch (e) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Invalid JSON.' }));
          return;
        }
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true }));
        if (pendingConfirm) {
          const resolve = pendingConfirm.resolve;
          pendingConfirm = null;
          resolve(parsed.proceed === true);
        }
      });

    } else {
      res.writeHead(404);
      res.end();
    }
  });

  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  return server.address().port;
}

// ── Entry point ───────────────────────────────────────────────────────────────

(async () => {
  const argv    = process.argv.slice(2);
  const guiMode = argv.includes('--gui');
  const autoYes = argv.includes('--yes');
  const args    = argv.filter(a => a !== '--gui' && a !== '--yes');

  if (guiMode) {
    const port = await startGuiServer(args[0] || '', args[1] || '');
    const url  = `http://127.0.0.1:${port}`;
    console.log(`Superbill Processor GUI → ${url}`);
    console.log('Press Ctrl+C to stop the server.');
    // Auto-open browser on Windows
    exec(`start ${url}`, err => { if (err) console.error('Could not open browser:', err.message); });
    // Keep process alive
    process.stdin.resume();
    return;
  }

  if (args.length < 2) {
    console.error('Usage:');
    console.error('  node superbill_processor.js <input.xlsx> <output.xlsx> [--yes]');
    console.error('  node superbill_processor.js --gui');
    process.exit(1);
  }

  const [inputPath, outputPath] = args;
  if (!fs.existsSync(inputPath)) {
    console.error(`Input file not found: ${inputPath}`);
    process.exit(1);
  }

  const confirmFn = autoYes ? () => Promise.resolve(true) : cliConfirm;
  const ok = await run(inputPath, outputPath, console.log, confirmFn);
  process.exit(ok ? 0 : 1);
})();
