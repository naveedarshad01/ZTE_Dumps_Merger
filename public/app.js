'use strict';
const $ = id => document.getElementById(id);
const fmt = n => Number(n).toLocaleString();
const state = { files: [], summary: null, audit: null, worker: null, pending: new Map(), seq: 0, busy: false, selected: new Set(), coveragePage: 0, auditPage: 0, pageResult: null, pageRequest: 0 };
function el(tag, text, className) { const e = document.createElement(tag); if (text !== undefined) e.textContent = text; if (className) e.className = className; return e; }
function showNotice(message, success = false) { $('notice').replaceChildren(el('div', message)); $('notice').className = 'notice' + (success ? ' success' : ''); $('notice').hidden = false; }
function showError(error) { showNotice(error.message || String(error)); }
function toast(text) { $('toast').textContent = text; $('toast').hidden = false; setTimeout(() => { $('toast').hidden = true; }, 5000); }
function resetWorker() {
  state.worker?.terminate();
  for (const { reject } of state.pending.values()) reject(new Error('Operation cancelled. Inspect the files again to continue.'));
  state.pending.clear(); state.worker = null;
}
function worker() {
  if (state.worker) return state.worker;
  if (!window.Worker) throw new Error('This browser does not support background processing. Open the app in a recent Chrome, Edge or Firefox browser.');
  const w = new Worker('worker.js?v=1.2.0');
  w.onmessage = ({ data }) => {
    if (data.type === 'progress') { $('progress-label').textContent = data.value.message; if (data.value.fraction > 0) $('progress').value = data.value.fraction; else $('progress').removeAttribute('value'); return; }
    const entry = state.pending.get(data.id); if (!entry) return; state.pending.delete(data.id); data.error ? entry.reject(new Error(data.error)) : entry.resolve(data.value);
  };
  w.onerror = () => { for (const p of state.pending.values()) p.reject(new Error('The processing worker stopped. Try a smaller batch, or reopen the app in Chrome or Edge.')); state.pending.clear(); w.terminate(); state.worker = null; };
  state.worker = w; return w;
}
function call(action, payload, transfers = []) { return new Promise((resolve, reject) => { const id = ++state.seq; try { const w = worker(); state.pending.set(id, { resolve, reject }); w.postMessage({ id, action, payload }, transfers); } catch (e) { state.pending.delete(id); reject(e); } }); }
function controls() {
  $('inspect').disabled = state.busy || state.files.length < 2;
  $('files').disabled = state.busy; $('clear').disabled = state.busy;
  $('run').disabled = state.busy || !state.summary;
  $('download-merge').disabled = state.busy || !state.summary || !!state.summary.issues.length;
  $('download-audit').disabled = state.busy || !state.audit;
  for (const e of document.querySelectorAll('.remove-file,#settings input,#settings select,#select-all,#select-none')) e.disabled = state.busy;
}
async function task(message, action) {
  if (state.busy) return;
  state.busy = true; document.body.classList.add('busy'); controls(); $('progress-panel').hidden = false; $('progress-label').textContent = message; $('progress').removeAttribute('value');
  try { return await action(); } catch (e) { showError(e); } finally { state.busy = false; document.body.classList.remove('busy'); $('progress-panel').hidden = true; controls(); }
}
function invalidate() {
  resetWorker(); state.summary = null; state.audit = null; state.pageRequest++; $('settings').hidden = true; $('results').hidden = true; $('notice').hidden = true; controls();
}
function invalidateAudit() {
  if (!state.audit) return;
  state.audit = null; state.pageRequest++; $('audit-tab').disabled = true; $('stat-differences').textContent = '—'; $('stat-missing').textContent = '—'; switchTab('coverage'); controls(); toast('Audit settings changed. Run the audit again.');
}
function drawFiles() {
  const list = $('file-list'); list.replaceChildren();
  state.files.forEach((f, i) => {
    const row = el('div', undefined, 'file-row'), meta = el('div', undefined, 'file-meta'), title = el('span', f.name, 'file-name'); title.title = f.name;
    meta.append(title, el('span', (f.size / 1048576).toFixed(2) + ' MiB' + (i === 0 ? ' · base workbook' : ''), 'file-size'));
    const remove = el('button', '×', 'remove-file'); remove.setAttribute('aria-label', 'Remove ' + f.name); remove.onclick = () => { if (!state.busy) { state.files.splice(i, 1); invalidate(); drawFiles(); } };
    row.append(el('span', 'File-' + (i + 1), 'file-label'), meta, remove); list.append(row);
  });
  $('file-count').textContent = state.files.length ? state.files.length + ' workbooks · ' + (state.files.reduce((n, f) => n + f.size, 0) / 1048576).toFixed(1) + ' MiB' : 'No files selected';
  $('clear').hidden = !state.files.length; controls();
}
function addFiles(files) {
  if (state.busy) return;
  const incoming = Array.from(files), invalid = incoming.find(f => !/\.xlsx$/i.test(f.name));
  if (invalid) return showNotice('Only .xlsx files are supported. Please convert ' + invalid.name + ' to a standard Excel workbook.');
  const result = [...state.files];
  for (const f of incoming) if (!result.some(x => x.name === f.name && x.size === f.size && x.lastModified === f.lastModified)) result.push(f);
  if (result.length > 20 || result.reduce((n, f) => n + f.size, 0) > 150 * 1048576) return showNotice('Choose up to 20 files with a combined size of 150 MiB or less.');
  state.files = result; invalidate(); drawFiles();
}
$('files').onchange = e => { addFiles(e.target.files); e.target.value = ''; };
for (const type of ['dragenter', 'dragover']) $('dropzone').addEventListener(type, e => { e.preventDefault(); if (!state.busy) $('dropzone').classList.add('drag'); });
for (const type of ['dragleave', 'drop']) $('dropzone').addEventListener(type, e => { e.preventDefault(); $('dropzone').classList.remove('drag'); if (type === 'drop') addFiles(e.dataTransfer.files); });
$('clear').onclick = () => { state.files = []; invalidate(); drawFiles(); };
$('cancel').onclick = () => { invalidate(); showNotice('Processing cancelled. Your selected files are ready to inspect again.'); };
$('inspect').onclick = () => task('Opening workbooks…', async () => {
  resetWorker(); state.audit = null; $('notice').hidden = true;
  const data = []; for (const f of state.files) data.push({ name: f.name, bytes: await f.arrayBuffer() });
  state.summary = await call('inspect', data, data.map(f => f.bytes)); state.selected = new Set(state.summary.sheets.map(s => s.name)); state.coveragePage = 0;
  $('settings').hidden = false; $('results').hidden = false; $('stat-sheets').textContent = fmt(state.summary.sheets.length); $('stat-records').textContent = fmt(state.summary.totalRecords);
  $('stat-differences').textContent = '—'; $('stat-missing').textContent = '—'; $('audit-tab').disabled = true; switchTab('coverage'); renderPicker(); renderCoverage();
  if (state.summary.issues.length) {
    showNotice(state.summary.issues.length + ' template compatibility issue(s) prevent merging. You can still compare parameter values.');
    const ul = el('ul'); for (const i of state.summary.issues.slice(0, 25)) ul.append(el('li', i.sheet + ' · ' + i.file + ': ' + i.message)); $('notice').append(ul);
    if (state.summary.issues.length > 25) $('notice').append(el('p', 'The audit export includes the complete issue list.'));
  } else { showNotice(fmt(state.summary.totalRecords) + ' configuration records ready to merge. Each MO uses its widest sheet as the reference. Values match by parameter name; missing parameters stay blank.', !state.summary.warnings.length); }
  if (state.summary.warnings.length) {
    const details = el('details'), heading = el('summary', state.summary.warnings.length + ' template notice(s) — merging is ' + (state.summary.issues.length ? 'blocked by the issues above' : 'available'));
    const ul = el('ul'); for (const warning of state.summary.warnings.slice(0, 25)) ul.append(el('li', warning));
    details.append(heading, ul); $('notice').append(details);
    if (!state.summary.issues.length) details.open = true;
  }
  if (state.summary.templateDetails?.length) $('notice').append(el('p', 'The audit workbook includes a Template details sheet with the header cells, original values and actions.'));
  $('settings').scrollIntoView({ behavior: 'smooth', block: 'start' });
});
function settings() { return { mode: $('match-mode').value, normalizeRoots: $('normalize-roots').checked, includeIdentity: $('include-identity').checked, ...($('scope').value === 'selected' ? { selected: [...state.selected] } : {}) }; }
$('match-mode').onchange = () => { $('mode-help').textContent = $('match-mode').value === 'strict' ? 'Match all Primary Key columns from row 5, including site identity. Use for repeated exports of the same sites.' : 'Match the same local object identifier across sites. Duplicate keys are flagged for review.'; invalidateAudit(); };
$('include-identity').onchange = invalidateAudit;
$('normalize-roots').onchange = invalidateAudit;
$('scope').onchange = () => { $('sheet-picker').hidden = $('scope').value !== 'selected'; invalidateAudit(); };
function visibleSheets() { const term = $('sheet-search').value.toLowerCase(); return (state.summary?.sheets || []).filter(s => s.name.toLowerCase().includes(term)); }
function renderPicker() {
  const holder = $('sheet-choices'); holder.replaceChildren();
  for (const s of visibleSheets()) { const label = el('label', undefined, 'sheet-choice'), cb = el('input'); cb.type = 'checkbox'; cb.checked = state.selected.has(s.name); cb.onchange = () => { cb.checked ? state.selected.add(s.name) : state.selected.delete(s.name); $('selected-count').textContent = fmt(state.selected.size) + ' selected'; invalidateAudit(); }; label.append(cb, el('span', s.name)); holder.append(label); }
  $('selected-count').textContent = fmt(state.selected.size) + ' selected';
}
$('sheet-search').oninput = renderPicker;
$('select-all').onclick = () => { visibleSheets().forEach(s => state.selected.add(s.name)); renderPicker(); invalidateAudit(); };
$('select-none').onclick = () => { state.selected.clear(); renderPicker(); invalidateAudit(); };
$('run').onclick = () => task('Comparing object parameters…', async () => {
  const opts = settings(); if (opts.selected && !opts.selected.length) throw new Error('Select at least one sheet to audit.');
  state.audit = await call('audit', opts); state.auditPage = 0; $('audit-tab').disabled = false;
  $('stat-differences').textContent = fmt(state.audit.counts.Different || 0);
  $('stat-missing').textContent = fmt(Object.entries(state.audit.counts).reduce((n, [k, v]) => n + (!['Same', 'Different'].includes(k) ? v : 0), 0));
  const select = $('audit-sheet'); select.replaceChildren(new Option('All audited sheets', '')); for (const name of Object.keys(state.audit.bySheet)) select.append(new Option(name, name));
  $('result-caption').textContent = fmt(state.audit.total) + ' parameter comparisons'; switchTab('audit'); await renderAudit();
  $('results').scrollIntoView({ behavior: 'smooth', block: 'start' }); toast(state.summary.issues.length ? 'Audit complete. Review template issues before merging.' : 'Audit complete. Both workbooks are ready to download.');
});
function switchTab(name) {
  for (const n of ['coverage', 'audit']) { $(n + '-view').hidden = n !== name; $(n + '-tab').setAttribute('aria-selected', String(n === name)); }
}
$('coverage-tab').onclick = () => switchTab('coverage'); $('audit-tab').onclick = () => switchTab('audit');
for (const id of ['coverage-tab', 'audit-tab']) $(id).onkeydown = e => { if (['ArrowLeft', 'ArrowRight'].includes(e.key) && !$('audit-tab').disabled) { e.preventDefault(); const target = id === 'coverage-tab' ? 'audit' : 'coverage'; switchTab(target); $(target + '-tab').focus(); } };
function tableHead(table, labels) { const tr = el('tr'); labels.forEach(t => { const th = el('th', t); th.scope = 'col'; tr.append(th); }); table.querySelector('thead').replaceChildren(tr); }
function badge(text, type) { return el('span', text, 'status ' + type); }
function renderCoverage() {
  if (!state.summary) return;
  const search = $('coverage-search').value.toLowerCase(), rows = state.summary.sheets.filter(s => (!search || s.name.toLowerCase().includes(search)) && (!$('populated-only').checked || s.total));
  state.coveragePage = Math.min(state.coveragePage, Math.max(0, Math.ceil(rows.length / 40) - 1));
  const start = state.coveragePage * 40, table = $('coverage-table'), body = table.querySelector('tbody');
  tableHead(table, ['Configuration sheet', ...state.summary.files.map(f => f.label + ' rows / columns'), 'Merged rows', 'Reference', 'Template']); body.replaceChildren();
  for (const s of rows.slice(start, start + 40)) {
    const tr = el('tr'); tr.append(el('td', s.name)); s.counts.forEach((n, i) => tr.append(el('td', n === null ? 'Missing sheet' : fmt(n) + ' / ' + fmt(s.columnCounts[i]), 'number')));
    tr.append(el('td', fmt(s.total), 'number'), el('td', s.reference + ' · ' + fmt(s.columns) + ' columns')); const status = el('td'); status.append(badge(s.compatible ? 'Compatible' : 'Review needed', s.compatible ? 'compatible' : 'missing')); tr.append(status); body.append(tr);
  }
  if (!rows.length) emptyRow(body, state.summary.files.length + 4, 'No sheets match this filter.');
  $('coverage-count').textContent = rows.length ? fmt(start + 1) + '–' + fmt(Math.min(start + 40, rows.length)) + ' of ' + fmt(rows.length) + ' sheets' : '0 sheets';
  $('coverage-prev').disabled = start === 0; $('coverage-next').disabled = start + 40 >= rows.length;
}
function emptyRow(body, columns, text) { const tr = el('tr'), td = el('td', text, 'empty-row'); td.colSpan = columns; tr.append(td); body.append(tr); }
for (const id of ['coverage-search', 'populated-only']) $(id).addEventListener('input', () => { state.coveragePage = 0; renderCoverage(); });
$('coverage-prev').onclick = () => { state.coveragePage--; renderCoverage(); }; $('coverage-next').onclick = () => { state.coveragePage++; renderCoverage(); };
async function renderAudit() {
  if (!state.audit) return;
  const token = ++state.pageRequest; $('audit-count').textContent = 'Loading comparisons…';
  $('audit-prev').disabled = true; $('audit-next').disabled = true;
  try {
    const page = await call('page', { offset: state.auditPage * 40, limit: 40, status: $('audit-status').value, sheet: $('audit-sheet').value, search: $('audit-search').value });
    if (token !== state.pageRequest) return; state.pageResult = page;
    const table = $('audit-table'), body = table.querySelector('tbody'); tableHead(table, ['Sheet', 'Comparison key', 'Parameter', 'Status', ...state.summary.files.flatMap(f => [f.label + ' original LDN', f.label + ' value'])]); body.replaceChildren();
    for (const r of page.rows) {
      const tr = el('tr'); tr.append(el('td', r.sheet), el('td', r.object, 'object'), el('td', r.parameter));
      const status = el('td', undefined, 'status-cell'); status.title = r.reason; status.append(badge(r.status, r.status === 'Different' ? 'different' : r.status.startsWith('Missing') ? 'missing' : ['No key', 'Ambiguous key'].includes(r.status) ? 'ambiguous' : '')); tr.append(status);
      r.values.forEach((v, i) => { tr.append(el('td', r.ldns[i], 'object')); const td = el('td', v); if (r.contexts[i]) td.append(el('small', r.contexts[i], 'cell-source')); tr.append(td); }); body.append(tr);
    }
    if (!page.rows.length) emptyRow(body, state.summary.files.length * 2 + 4, 'No parameter comparisons match these filters.');
    $('audit-count').textContent = page.total ? fmt(page.offset + 1) + '–' + fmt(Math.min(page.offset + page.limit, page.total)) + ' of ' + fmt(page.total) + ' comparisons' : '0 comparisons';
    $('audit-prev').disabled = state.auditPage === 0; $('audit-next').disabled = page.offset + page.limit >= page.total;
  } catch (e) { if (token === state.pageRequest) showError(e); }
}
let searchTimer;
$('audit-search').oninput = () => { clearTimeout(searchTimer); searchTimer = setTimeout(() => { state.auditPage = 0; renderAudit(); }, 400); };
for (const id of ['audit-sheet', 'audit-status']) $(id).onchange = () => { state.auditPage = 0; renderAudit(); };
$('audit-prev').onclick = () => { if (state.auditPage > 0) state.auditPage--; renderAudit(); }; $('audit-next').onclick = () => { state.auditPage++; renderAudit(); };
function download(bytes, filename) { const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), url = URL.createObjectURL(blob), a = el('a'); a.href = url; a.download = filename; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 60000); toast(filename + ' is ready.'); }
$('download-merge').onclick = () => task('Merging the original templates…', async () => download(await call('merge', {}), 'ITBBU_Merged.xlsx'));
$('download-audit').onclick = () => task('Building the parameter audit workbook…', async () => download(await call('export', { issuesOnly: $('export-issues').checked }), 'ITBBU_Parameter_Audit.xlsx'));
if (location.protocol === 'file:') showNotice('Open this app through a local web server or deploy it to Vercel. Background processing cannot run from a double-clicked HTML file. See README.md.');
drawFiles();
