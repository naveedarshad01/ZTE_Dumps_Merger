/* ITBBU Config Studio. OOXML is patched in place to preserve the supplied template. */
(function (root) {
  'use strict';
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const IGNORED = new Set(['templateinfo', 'index']);
  const IDENTITY = new Set(['modind', 'managedelementtype', 'subnetwork', 'managedelement', 'ne_name', 'ldn', 'moid']);
  const MiB = 1024 * 1024;
  const limits = { files: 20, inputBytes: 150 * MiB, expandedBytes: 600 * MiB, entryBytes: 96 * MiB, auditRows: 2000000 };
  const xml = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  const unxml = s => String(s ?? '').replace(/&#x([\da-f]+);|&#(\d+);|&(amp|lt|gt|quot|apos);/gi, (_, h, d, n) => h ? String.fromCodePoint(parseInt(h, 16)) : d ? String.fromCodePoint(Number(d)) : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[n.toLowerCase()]));
  function attrs(s) { const a = {}; for (const m of s.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) a[m[1]] = unxml(m[2] ?? m[3]); return a; }
  function colNum(s) { let n = 0; for (const c of s.toUpperCase()) n = n * 26 + c.charCodeAt(0) - 64; return n - 1; }
  function colName(n) { let s = ''; for (n++; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; }
  function textRuns(s) { return [...s.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(m => unxml(m[1])).join(''); }
  function setAttr(tag, key, value) {
    const re = new RegExp('(\\s' + key + '=)(?:"[^"]*"|\'[^\']*\')');
    return re.test(tag) ? tag.replace(re, '$1"' + xml(value) + '"') : tag.replace(/\/?>(?=$)/, m => ' ' + key + '="' + xml(value) + '"' + m);
  }
  const normal = s => String(s ?? '').toLowerCase().trim();
  function resolvePart(base, target) {
    const parts = (target.startsWith('/') ? target.slice(1) : base.slice(0, base.lastIndexOf('/') + 1) + target).split('/');
    const result = []; for (const p of parts) { if (p === '..') result.pop(); else if (p && p !== '.') result.push(p); }
    return result.join('/');
  }
  function zipSafety(bytes) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let end = -1;
    for (let p = bytes.length - 22; p >= Math.max(0, bytes.length - 65557); p--) if (view.getUint32(p, true) === 0x06054b50 && p + 22 + view.getUint16(p + 20, true) === bytes.length) { end = p; break; }
    if (end < 0) throw new Error('This is not a valid .xlsx ZIP package. Encrypted workbooks are not supported.');
    const count = view.getUint16(end + 10, true); let at = view.getUint32(end + 16, true), expanded = 0;
    if (count === 65535 || at === 0xffffffff || view.getUint16(end + 4, true) !== 0) throw new Error('ZIP64 and multi-part archives are not supported.');
    const names = new Set();
    for (let n = 0; n < count; n++) {
      if (at + 46 > end || view.getUint32(at, true) !== 0x02014b50) throw new Error('Damaged ZIP directory.');
      if (view.getUint16(at + 8, true) & 1) throw new Error('Please remove the workbook password before uploading.');
      const size = view.getUint32(at + 24, true), length = view.getUint16(at + 28, true), extra = view.getUint16(at + 30, true), comment = view.getUint16(at + 32, true);
      if (size > limits.entryBytes) throw new Error('A worksheet exceeds the 96 MiB expanded size limit. Split the source workbook.');
      const name = new TextDecoder().decode(bytes.subarray(at + 46, at + 46 + length));
      if (names.has(name) || name.split('/').includes('..') || name.startsWith('/')) throw new Error('Invalid or duplicate ZIP entry.');
      names.add(name); expanded += size; at += 46 + length + extra + comment;
    }
    if (expanded > limits.expandedBytes) throw new Error('Expanded workbook is too large for browser processing.');
    return expanded;
  }
  function parseCells(row, strings) {
    const cells = [];
    for (const m of row.matchAll(/<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/g)) {
      const tag = m[0].match(/^<c\b[^>]*>/)[0], a = attrs(tag), ref = /^([A-Z]+)(\d+)$/i.exec(a.r || '');
      if (!ref) throw new Error('A worksheet cell has no valid A1 reference.');
      const value = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/.exec(m[0]);
      let val = value ? unxml(value[1]) : '', kind = a.t || 'n';
      if (kind === 's') { const index = Number(val); if (!Number.isInteger(index) || !strings[index]) throw new Error('Invalid shared-string reference at ' + a.r); val = strings[index].text; kind = 's'; }
      else if (kind === 'inlineStr') { val = textRuns(m[0]); kind = 's'; }
      else if (kind === 'str') kind = 's';
      const f = /<f(?:\s[^>]*)?>([\s\S]*?)<\/f>|<f\b[^>]*\/>/.exec(m[0]);
      if (f) { val = '=' + unxml(f[1] || '[shared formula]'); kind = 'f'; }
      cells[colNum(ref[1])] = { text: val, kind: val === '' ? 'blank' : kind };
    }
    return cells;
  }
  const values = c => Array.from(c, v => v?.text ?? '');
  async function readWorkbook(input, index, Zip, progress) {
    const bytes = input.bytes instanceof Uint8Array ? input.bytes : new Uint8Array(input.bytes);
    const expanded = zipSafety(bytes), zip = await Zip.loadAsync(bytes);
    const read = async path => { if (!zip.file(path)) throw new Error('Missing workbook part: ' + path); const s = await zip.file(path).async('string'); if (/<!DOCTYPE|<!ENTITY/i.test(s)) throw new Error('XML entity declarations are not supported.'); return s; };
    const content = await read('[Content_Types].xml');
    if (/macroEnabled|vbaProject/i.test(content)) throw new Error('Macro-enabled workbooks are not supported. Save a macro-free .xlsx copy.');
    const book = await read('xl/workbook.xml'), rels = await read('xl/_rels/workbook.xml.rels');
    if (!book.includes(NS)) throw new Error('Please save this file as a standard Excel .xlsx workbook.');
    const links = {};
    for (const m of rels.matchAll(/<Relationship\b[^>]*\/>/g)) { const a = attrs(m[0]); if (a.TargetMode !== 'External') links[a.Id] = resolvePart('xl/workbook.xml', a.Target); }
    const sharedPath = Object.values(links).find(p => /sharedStrings\.xml$/.test(p));
    const strings = sharedPath && zip.file(sharedPath) ? [...(await read(sharedPath)).matchAll(/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)].map(m => ({ text: textRuns(m[1]), inner: m[1] })) : [];
    const entries = [...book.matchAll(/<sheet\b[^>]*\/>/g)].map(m => ({ ...attrs(m[0]), tag: m[0] }));
    const w = { index, label: 'File-' + (index + 1), name: input.name, bytes: bytes.length, expanded, zip, read, book, rels, content, strings, entries, sheets: new Map(), helpers: new Map(), styles: await read('xl/styles.xml'), records: 0, sites: new Set() };
    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i], path = links[entry['r:id']];
      if (!path) throw new Error('Unresolved worksheet: ' + entry.name);
      entry.path = path;
      if (IGNORED.has(normal(entry.name))) continue;
      const raw = await read(path);
      const body = /<sheetData(?:\s[^>]*)?>([\s\S]*?)<\/sheetData>/.exec(raw);
      if (!body) throw new Error('Worksheet has no standard sheetData: ' + entry.name);
      // Read only the first five rows initially: lookup sheets have a different structure.
      const headers = []; let isConfig = false;
      for (const m of body[1].matchAll(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/g)) {
        const n = Number(attrs(m[0].match(/^<row\b[^>]*>/)[0]).r); if (n > 5) break;
        headers[n - 1] = values(parseCells(m[0], strings));
      }
      isConfig = normal(headers[0]?.[0]) === 'modind';
      if (!isConfig) { w.helpers.set(entry.name, { entry, raw }); continue; }
      const names = headers[0] || [], seen = new Set();
      for (const name of names) { if (!name || seen.has(name)) throw new Error(entry.name + ': blank or duplicate parameter name in row 1.'); seen.add(name); }
      if (headers.length !== 5 || headers.some(h => !h)) throw new Error(entry.name + ': expected five template rows.');
      const rows = [], mergeIssues = [];
      for (const m of body[1].matchAll(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/g)) {
        const n = Number(attrs(m[0].match(/^<row\b[^>]*>/)[0]).r); if (n < 6) continue;
        const cells = parseCells(m[0], strings); if (!cells.some(c => c && c.text !== '')) continue;
        if (cells.length > names.length) throw new Error(entry.name + ': data extends beyond the parameter header at row ' + n + '.');
        if (cells.some(c => c?.kind === 'f')) mergeIssues.push('Formula in data row ' + n + '. Convert configuration formulas to values before merging.');
        const row = { number: n, cells }; rows.push(row);
        const site = cells[names.findIndex(h => normal(h) === 'ne_name')]?.text; if (site) w.sites.add(site);
      }
      for (const m of raw.matchAll(/<mergeCell\b[^>]*\/>/g)) if (Number((attrs(m[0]).ref || '').match(/\d+$/)?.[0]) >= 6) mergeIssues.push('Merged cells occur in the data area.');
      if (/<tableParts\b|<drawing\b|<hyperlinks\b|<legacyDrawing\b/.test(raw)) mergeIssues.push('This sheet contains a table, drawing, hyperlink or comment that needs manual preservation.');
      const keys = names.filter((_, c) => /primary\s*key/i.test(headers[4][c] || ''));
      const sheet = { name: entry.name, entry, names, headers, rows, keys, mergeIssues };
      w.sheets.set(entry.name, sheet); w.records += rows.length;
      if (i % 60 === 0) progress?.({ stage: 'read', message: w.label + ': reading ' + entry.name, file: index, fraction: (i + 1) / entries.length });
    }
    return w;
  }
  async function inspect(inputs, Zip, progress) {
    if (inputs.length < 2 || inputs.length > limits.files) throw new Error('Select between 2 and 20 .xlsx workbooks.');
    if (inputs.reduce((n, f) => n + f.bytes.byteLength, 0) > limits.inputBytes) throw new Error('Select at most 150 MiB of workbooks per batch.');
    const books = []; let expanded = 0;
    for (let i = 0; i < inputs.length; i++) { if (!/\.xlsx$/i.test(inputs[i].name)) throw new Error('Only .xlsx workbooks are supported.'); const w = await readWorkbook(inputs[i], i, Zip, progress); books.push(w); expanded += w.expanded; if (expanded > limits.expandedBytes) throw new Error('Combined expanded workbooks exceed 600 MiB. Use a smaller batch.'); }
    const names = [...new Set(books.flatMap(w => [...w.sheets.keys()]))], issues = [], warnings = [];
    const sheets = names.map(name => {
      const present = books.map(w => w.sheets.get(name)), donor = present.find(Boolean), counts = present.map(s => s?.rows.length ?? null);
      let compatible = true;
      present.forEach((s, i) => {
        if (!s && books[i].entries.some(e => e.name === name)) { compatible = false; issues.push({ sheet: name, file: books[i].label, message: 'This named sheet exists but does not contain the expected row-6 ZTE configuration template.' }); }
        if (s && JSON.stringify(s.headers) !== JSON.stringify(donor.headers)) { compatible = false; issues.push({ sheet: name, file: books[i].label, message: 'Rows 1–5 differ from the first occurrence. Use matching template versions.' }); }
        for (const message of s?.mergeIssues || []) { compatible = false; issues.push({ sheet: name, file: books[i].label, message }); }
      });
      if (!present[0]) {
        const donorBook = books.find(w => w.sheets.has(name));
        if (donorBook.styles !== books[0].styles) { compatible = false; issues.push({ sheet: name, file: donorBook.label, message: 'A sheet absent from File-1 requires the same style table to preserve its template.' }); }
      }
      if (counts.reduce((n, c) => n + (c || 0), 0) > 1048571) { compatible = false; issues.push({ sheet: name, file: 'All', message: 'Merged data exceeds the Excel worksheet row limit.' }); }
      return { name, counts, total: counts.reduce((n, c) => n + (c || 0), 0), columns: donor.names.length, keys: donor.keys, compatible };
    });
    const caseNames = new Map();
    for (const name of names) { const key = normal(name); if (caseNames.has(key) && caseNames.get(key) !== name) issues.push({ sheet: name, file: 'All', message: 'Sheet names differ only in letter case. Rename them consistently before merging.' }); caseNames.set(key, name); }
    // Hidden enumeration tables supply the existing validation formulas; they are never appended.
    const helperNames = [...new Set(books.flatMap(w => [...w.helpers.keys()]))];
    for (const name of helperNames) {
      const helpers = books.map(w => w.helpers.get(name)), first = helpers.find(Boolean);
      if (!books[0].helpers.has(name) || helpers.some(h => h && h.raw !== first.raw)) issues.push({ sheet: name, file: 'All', message: 'Support/lookup sheet differs or is absent from File-1. Use a common template before merging.' });
      warnings.push(name + ': preserved unchanged from File-1 as template support; not a row-6 configuration table.');
    }
    if (!names.length) throw new Error('No ZTE configuration sheets found. Row 1 must begin with MODIND and data must start at row 6.');
    const session = { books, sheets, issues, warnings, Zip, audit: null };
    return session;
  }
  function publicSummary(session) {
    return { files: session.books.map(w => ({ label: w.label, name: w.name, size: w.bytes, records: w.records, sheets: w.sheets.size, sites: [...w.sites] })), sheets: session.sheets, issues: session.issues, warnings: session.warnings, totalRecords: session.books.reduce((n, w) => n + w.records, 0) };
  }
  function identityKey(sheet, row, mode, normalizeRoots = true) {
    const get = name => row.cells[sheet.names.indexOf(name)]?.text ?? '';
    let keys;
    if (mode === 'strict') keys = sheet.keys.length ? sheet.keys : sheet.names.filter(n => ['managedelementtype', 'subnetwork', 'managedelement', 'ldn'].includes(normal(n)));
    else {
      const ldn = sheet.names.find(n => normal(n) === 'ldn');
      if (ldn) keys = [ldn];
      else {
        keys = sheet.keys.filter(n => !IDENTITY.has(normal(n)));
        if (!keys.length) { const id = sheet.names.find(n => normal(n) === 'moid'); if (id) keys = [id]; }
        if (!keys.length && normal(sheet.name) === 'managedelement') return { key: 'singleton', label: '(site object)' };
      }
    }
    if (!keys.length || keys.some(k => get(k) === '')) return null;
    const pairs = keys.map(k => {
      let value = get(k);
      if (mode !== 'strict' && normalizeRoots && normal(k) === 'ldn') {
        const me = sheet.names.find(n => normal(n) === 'managedelement'), siteId = me ? get(me) : '';
        // Normalize only the known site-derived function ROOT, never a cell/neighbor ID.
        value = value.replace(/^((?:ENB|GNB)(?:CUCP|CUUP|DU)Function=)([^,]+)(?=,|$)/, (all, prefix, id) => siteId && id.endsWith('_' + siteId) ? prefix + id.slice(0, -siteId.length) + '{site}' : all);
      }
      return [k, value];
    });
    return { key: JSON.stringify(pairs), label: pairs.map(([k, v]) => k + '=' + v).join(' | ') };
  }
  function recordContext(sheet, row) {
    const get = n => row.cells[sheet.names.findIndex(h => normal(h) === n)]?.text ?? '';
    return (get('ne_name') || get('managedelement') || 'unnamed site') + ' / row ' + row.number;
  }
  const cellToken = c => c && c.text !== '' ? c.kind + '\u0000' + c.text : 'blank';
  function displayCell(c) { if (!c || c.text === '') return '(blank)'; return c.kind === 's' ? (/^\((?:blank|missing object|missing sheet|missing parameter)\)$/.test(c.text) ? '[text] ' + c.text : c.text) : c.text + (c.kind === 'n' ? ' [number]' : c.kind === 'b' ? ' [boolean]' : c.kind === 'e' ? ' [Excel error]' : ' [formula]'); }
  function *auditRows(session, options = {}) {
    const selected = options.selected ? new Set(options.selected) : null;
    const mode = options.mode === 'strict' ? 'strict' : 'local';
    for (const coverage of session.sheets) {
      if (selected && !selected.has(coverage.name)) continue;
      const sheets = session.books.map(w => w.sheets.get(coverage.name));
      const params = [...new Set(sheets.filter(Boolean).flatMap(s => s.names))].filter(n => options.includeIdentity || !IDENTITY.has(normal(n)));
      const groups = new Map();
      sheets.forEach((sheet, f) => {
        for (const row of sheet?.rows || []) {
          const id = identityKey(sheet, row, mode, options.normalizeRoots !== false), key = id?.key || '__unmatched_' + f + '_' + row.number;
          if (!groups.has(key)) groups.set(key, { label: id?.label || '(no usable key: ' + session.books[f].label + ', row ' + row.number + ')', unkeyed: !id, rows: session.books.map(() => []) });
          groups.get(key).rows[f].push(row);
        }
      });
      for (const group of groups.values()) for (const parameter of params) {
        let status = 'Same', reason = '', hasMissingSheet = false, hasMissingObject = false, hasMissingParam = false, hasDuplicate = false;
        const tokens = [], vals = [], contexts = [], ldns = [];
        for (let f = 0; f < sheets.length; f++) {
          const sheet = sheets[f], rows = group.rows[f], column = sheet?.names.indexOf(parameter) ?? -1;
          const ldnCol = sheet?.names.findIndex(n => normal(n) === 'ldn') ?? -1;
          ldns.push(!sheet ? '(missing sheet)' : !rows.length ? '(missing object)' : ldnCol < 0 ? '(no LDN column)' : rows.map(r => r.cells[ldnCol]?.text || '(blank LDN)').join('\n'));
          if (!sheet) { hasMissingSheet = true; vals.push('(missing sheet)'); contexts.push(''); }
          else if (!rows.length) { hasMissingObject = true; vals.push('(missing object)'); contexts.push(''); }
          else if (column < 0) { hasMissingParam = true; vals.push('(missing parameter)'); contexts.push(rows.map(r => recordContext(sheet, r)).join('\n')); }
          else if (rows.length > 1) { hasDuplicate = true; vals.push(rows.map(r => 'row ' + r.number + ': ' + displayCell(r.cells[column])).join('\n')); contexts.push(rows.map(r => recordContext(sheet, r)).join('\n')); }
          else { vals.push(displayCell(rows[0].cells[column])); contexts.push(recordContext(sheet, rows[0])); tokens.push(cellToken(rows[0].cells[column])); }
        }
        if (group.unkeyed) { status = 'No key'; reason = 'No reliable object key. This row was not paired by position.'; }
        else if (hasDuplicate) { status = 'Ambiguous key'; reason = 'Multiple rows share the comparison key in at least one file. No arbitrary pairing.'; }
        else if (hasMissingSheet) { status = 'Missing sheet'; reason = 'Worksheet is absent in at least one file.'; }
        else if (hasMissingObject) { status = 'Missing object'; reason = 'Object key is absent in at least one file.'; }
        else if (hasMissingParam) { status = 'Missing parameter'; reason = 'Parameter column is absent in at least one file.'; }
        else if (new Set(tokens).size > 1) { status = 'Different'; reason = 'Exact value or stored cell type differs.'; }
        yield { sheet: coverage.name, object: group.label, parameter, status, ldns, values: vals, contexts, reason };
      }
    }
  }
  function summarizeAudit(session, options = {}, progress) {
    let total = 0; const counts = {}, bySheet = {}, sample = [];
    for (const r of auditRows(session, options)) {
      if (++total > limits.auditRows) throw new Error('Audit exceeds 2 million comparisons. Select fewer sheets.');
      counts[r.status] = (counts[r.status] || 0) + 1;
      const s = bySheet[r.sheet] ||= { total: 0, different: 0, missing: 0, ambiguous: 0 };
      s.total++; if (r.status === 'Different') s.different++; else if (r.status.startsWith('Missing')) s.missing++; else if (['Ambiguous key', 'No key'].includes(r.status)) s.ambiguous++;
      if (sample.length < 60 && r.status !== 'Same') sample.push(r);
      if (total % 30000 === 0) progress?.({ stage: 'audit', message: 'Comparing parameters: ' + total.toLocaleString(), fraction: 0 });
    }
    session.audit = { total, counts, bySheet, options, sample };
    return session.audit;
  }
  function auditPage(session, query = {}) {
    const offset = Math.max(0, query.offset || 0), limit = Math.min(100, query.limit || 50), rows = [], search = normal(query.search);
    let total = 0;
    const opts = { ...(session.audit?.options || {}) };
    if (query.sheet) opts.selected = [query.sheet];
    for (const r of auditRows(session, opts)) {
      if (query.status === 'issues' && r.status === 'Same') continue;
      if (query.status && !['all', 'issues'].includes(query.status) && r.status !== query.status) continue;
      if (search && !normal([r.sheet, r.object, r.parameter, ...r.ldns, ...r.values].join(' ')).includes(search)) continue;
      if (total >= offset && rows.length < limit) rows.push(r); total++;
    }
    return { rows, total, offset, limit };
  }
  function transformRow(row, newNumber, source, base, baseStyle) {
    let out = row.replace(/^<row\b[^>]*>/, tag => setAttr(tag, 'r', newNumber));
    out = out.replace(/<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/g, cell => {
      let tag = cell.match(/^<c\b[^>]*>/)[0], a = attrs(tag); const col = (a.r || '').match(/^[A-Z]+/i)?.[0];
      tag = setAttr(tag, 'r', col + newNumber);
      if (source.styles !== base.styles) { const style = baseStyle[col]; if (style == null) throw new Error('Data-cell style cannot be preserved for column ' + col + '. Use matching templates.'); tag = setAttr(tag, 's', style); }
      let body = cell.slice(cell.indexOf('>') + 1);
      if (a.t === 's') { const v = /<v>(\d+)<\/v>/.exec(body), str = v && source.strings[Number(v[1])]; if (!str) throw new Error('Invalid shared string.'); tag = setAttr(tag, 't', 'inlineStr'); body = '<is>' + str.inner + '</is></c>'; }
      return tag + (tag.endsWith('/>') ? '' : body);
    });
    return out;
  }
  async function merge(session, progress) {
    if (session.issues.length) throw new Error('Resolve template compatibility issues before merging. The audit remains available.');
    const base = session.books[0], output = new session.Zip();
    // Reuse compressed ZIP objects for untouched entries, avoiding a lossy workbook round trip.
    for (const [name, file] of Object.entries(base.zip.files)) output.files[name] = file;
    let book = base.book, rels = base.rels, content = base.content, added = 0;
    let id = Math.max(0, ...base.entries.map(e => Number(e.sheetId)));
    const usedPaths = new Set(Object.keys(output.files));
    for (let i = 0; i < session.sheets.length; i++) {
      const coverage = session.sheets[i], name = coverage.name, donor = session.books.find(w => w.sheets.has(name)), target = donor.sheets.get(name);
      let raw = await donor.read(target.entry.path);
      const originalBody = /<sheetData(?:\s[^>]*)?>([\s\S]*?)<\/sheetData>/.exec(raw);
      const header = [...originalBody[1].matchAll(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/g)].filter(m => Number(attrs(m[0].match(/^<row\b[^>]*>/)[0]).r) < 6).map(m => m[0]).join('\n');
      const style = {};
      for (const m of raw.matchAll(/<col\b[^>]*\/>/g)) { const a = attrs(m[0]); for (let c = Number(a.min) - 1; c < Math.min(Number(a.max), target.names.length); c++) style[colName(c)] = a.style || '0'; }
      const prototype = [...originalBody[1].matchAll(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/g)].find(m => Number(attrs(m[0].match(/^<row\b[^>]*>/)[0]).r) >= 6)?.[0];
      for (const m of (prototype || '').matchAll(/<c\b[^>]*>/g)) { const a = attrs(m[0]); style[(a.r || '').replace(/\d+/g, '')] = a.s || '0'; }
      const data = []; let rowNumber = 6;
      for (const w of session.books) {
        const sheet = w.sheets.get(name); if (!sheet || !sheet.rows.length) continue;
        if (donor !== base && w.styles !== base.styles) throw new Error(name + ': unique sheet uses a different style table. Use a common complete template.');
        const source = w === donor ? raw : await w.read(sheet.entry.path), body = /<sheetData(?:\s[^>]*)?>([\s\S]*?)<\/sheetData>/.exec(source)[1], keep = new Set(sheet.rows.map(r => r.number));
        for (const m of body.matchAll(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/g)) {
          const n = Number(attrs(m[0].match(/^<row\b[^>]*>/)[0]).r);
          if (keep.has(n)) data.push(transformRow(m[0], rowNumber++, w, base, style));
        }
      }
      if (donor !== base) {
        if (donor.styles !== base.styles) throw new Error(name + ': additional sheet requires the same style table as File-1.');
        // Source headers with shared strings must be translated into inline strings as well.
      }
      const safeHeader = donor === base ? header : header.replace(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/g, r => transformRow(r, Number(attrs(r.match(/^<row\b[^>]*>/)[0]).r), donor, base, {}));
      raw = raw.replace(originalBody[0], '<sheetData>\n' + safeHeader + '\n' + data.join('\n') + '\n</sheetData>');
      raw = raw.replace(/<dimension\b[^>]*\/>/, '<dimension ref="A1:' + colName(target.names.length - 1) + Math.max(5, rowNumber - 1) + '"/>');
      // Extend existing validation / conditional-format ranges only where appended rows exceed them.
      raw = raw.replace(/<(?:dataValidation|conditionalFormatting)\b[^>]*>/g, tag => {
        const a = attrs(tag); if (!a.sqref) return tag;
        const refs = a.sqref.split(' ').map(r => r.replace(/^(\$?[A-Z]+\$?)(\d+):(\$?[A-Z]+\$?)(\d+)$/, (_, ac, ar, bc, br) => Number(ar) >= 6 && Number(br) < rowNumber - 1 ? ac + ar + ':' + bc + (rowNumber - 1) : _));
        return setAttr(tag, 'sqref', refs.join(' '));
      });
      let path = target.entry.path;
      if (donor !== base) {
        do { path = 'xl/worksheets/merged' + (++added) + '.xml'; } while (usedPaths.has(path)); usedPaths.add(path);
        const rid = 'rIdITBBU' + added;
        book = book.replace('</sheets>', '<sheet name="' + xml(name) + '" sheetId="' + (++id) + '" r:id="' + rid + '"' + (target.entry.state ? ' state="' + xml(target.entry.state) + '"' : '') + '/></sheets>');
        rels = rels.replace('</Relationships>', '<Relationship Id="' + rid + '" Type="' + REL + '/worksheet" Target="' + path.slice(3) + '"/></Relationships>');
        content = content.replace('</Types>', '<Override PartName="/' + path + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
      }
      output.file(path, raw);
      if (i % 35 === 0) progress?.({ stage: 'merge', message: 'Merging ' + name, fraction: i / session.sheets.length });
    }
    output.file('xl/workbook.xml', book); output.file('xl/_rels/workbook.xml.rels', rels); output.file('[Content_Types].xml', content);
    return output.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } }, m => progress?.({ stage: 'compress', message: 'Preparing merged workbook', fraction: m.percent / 100 }));
  }
  function auditHeaders(session) { return ['Sheet', 'Comparison key', 'Parameter', 'Status', ...session.books.flatMap(w => [w.label + ' original LDN', w.label + ' value', w.label + ' source']), 'Reason']; }
  function auditValues(r) { return [r.sheet, r.object, r.parameter, r.status, ...r.values.flatMap((v, i) => [r.ldns[i], v, r.contexts[i]]), r.reason]; }
  const auditStyles = '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="' + NS + '"><fonts count="3"><font><sz val="11"/><name val="Arial"/><color rgb="FF17243B"/></font><font><b/><sz val="11"/><name val="Arial"/><color rgb="FFFFFFFF"/></font><font><b/><sz val="15"/><name val="Arial"/><color rgb="FF17243B"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF17243B"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  function xlsxCell(v, c, row, style = 0) {
    if (String(v ?? '').length > 32767) throw new Error('A report cell exceeds Excel’s 32,767-character limit. Select fewer records or use strict keys.');
    const ref = colName(c) + row;
    return typeof v === 'number' ? '<c r="' + ref + '" s="' + style + '"><v>' + v + '</v></c>' : '<c r="' + ref + '" s="' + style + '" t="inlineStr"><is><t xml:space="preserve">' + xml(v) + '</t></is></c>';
  }
  function sheetXml(rows, widths, filter = true) {
    return '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="' + NS + '"><dimension ref="A1:' + colName(widths.length - 1) + rows.length + '"/><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>' + widths.map((w, c) => '<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols><sheetData>' + rows.map((r, i) => '<row r="' + (i + 1) + '" ht="' + (i ? 32 : 30) + '" customHeight="1">' + r.map((v, c) => xlsxCell(v, c, i + 1, i === 0 ? 1 : 0)).join('') + '</row>').join('') + '</sheetData>' + (filter ? '<autoFilter ref="A1:' + colName(widths.length - 1) + rows.length + '"/>' : '') + '</worksheet>';
  }
  async function exportAudit(session, options = {}, progress) {
    const zip = new session.Zip(), names = [], headers = auditHeaders(session), widths = [28, 64, 32, 22, ...session.books.flatMap(() => [68,32,32]), 56];
    const add = (name, data, widths, filter = true) => { names.push(name); zip.file('xl/worksheets/sheet' + names.length + '.xml', sheetXml(data, widths, filter)); };
    const a = session.audit || summarizeAudit(session, options);
    add('Summary', [['Metric', 'Value'], ['Matching mode', a.options.mode === 'strict' ? 'Same site: full primary key' : 'Across sites: local LDN / MO key'], ['Normalize site roots', a.options.mode !== 'strict' && a.options.normalizeRoots !== false ? 'ENB/GNB CUCP, CUUP and DU root suffix equal to ManagedElement becomes {site}. Other IDs remain exact.' : 'No'], ['Identity parameters included', a.options.includeIdentity ? 'Yes' : 'No'], ['Export scope', options.issuesOnly ? 'Differences, missing and ambiguous only' : 'All comparisons'], ['Merged configuration records', session.books.reduce((n, w) => n + w.records, 0)], ['Audited parameter comparisons', a.total], ...Object.entries(a.counts), ['Blank value', '(blank) means a present, empty cell; it is not a missing object. Literal marker text is prefixed [text].'], ['Value types', 'Text is exact and case-sensitive. Non-text values have [number], [boolean], [Excel error] or [formula] labels.'], ['Object matching', 'Rows are never paired by position. Duplicate local keys are marked Ambiguous key.'], ['Comparison assumption', 'Equal local IDs are assumed comparable. Hardware and neighbor IDs may differ in meaning across sites. Variation does not establish incorrect configuration.'], ['Lookup sheets', 'TemplateInfo and Index are not merged/audited. Non-configuration lookup sheets remain unchanged.'], ['Compatibility issues', session.issues.length]], [38, 110], false);
    add('Files', [['Label', 'Filename', 'Sites', 'Configuration rows'], ...session.books.map(w => [w.label, w.name, [...w.sites].join(', '), w.records])], [14, 100, 45, 24]);
    add('Sheet coverage', [['Sheet', ...session.books.map(w => w.label + ' rows'), 'Merged rows', 'Parameters', 'Audit comparisons', 'Different', 'Missing', 'Ambiguous / no key'], ...session.sheets.map(s => { const k = a.bySheet[s.name] || {}; return [s.name, ...s.counts.map(n => n === null ? 'Missing sheet' : n), s.total, s.columns, k.total || 0, k.different || 0, k.missing || 0, k.ambiguous || 0]; })], [34, ...session.books.map(() => 20), 20, 18, 24, 18, 18, 26]);
    // Shared strings and chunked XML keep a large audit below JavaScript's string-size limit.
    const stringIds = new Map(), shared = [], encoder = new TextEncoder();
    const sid = value => { const s = String(value ?? ''); if (s.length > 32767) throw new Error('A report cell exceeds Excel’s 32,767-character limit. Use stricter object keys.'); if (!stringIds.has(s)) { stringIds.set(s, shared.length); shared.push(s); } return stringIds.get(s); };
    let chunks = [], pending = '', rowCount = 0, part = 1, total = 0;
    const flush = () => { if (pending) { chunks.push(encoder.encode(pending)); pending = ''; } };
    const writeRow = vals => { rowCount++; pending += '<row r="' + rowCount + '"' + (rowCount === 1 ? ' ht="30" customHeight="1"' : '') + '>' + vals.map((v, c) => '<c r="' + colName(c) + rowCount + '" s="' + (rowCount === 1 ? 1 : 0) + '" t="s"><v>' + sid(v) + '</v></c>').join('') + '</row>'; if (rowCount % 2000 === 0) flush(); };
    const startPart = () => { chunks = []; rowCount = 0; pending = '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="' + NS + '"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane xSplit="3" ySplit="1" topLeftCell="D2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="30"/><cols>' + widths.map((w, c) => '<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols><sheetData>'; writeRow(headers); };
    const finishPart = () => { pending += '</sheetData><autoFilter ref="A1:' + colName(headers.length - 1) + rowCount + '"/></worksheet>'; flush(); const length = chunks.reduce((n, a) => n + a.length, 0), bytes = new Uint8Array(length); let offset = 0; for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; } names.push(part === 1 ? 'Parameter audit' : 'Parameter audit ' + part); zip.file('xl/worksheets/sheet' + names.length + '.xml', bytes); chunks = []; };
    startPart();
    for (const r of auditRows(session, a.options)) {
      if (options.issuesOnly && r.status === 'Same') continue;
      if (rowCount === 1000001) { finishPart(); part++; startPart(); }
      writeRow(auditValues(r)); total++;
      if (total % 40000 === 0) progress?.({ stage: 'export', message: 'Writing audit: ' + total.toLocaleString() + ' rows', fraction: total / a.total });
    }
    finishPart();
    zip.file('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8"?><sst xmlns="' + NS + '" uniqueCount="' + shared.length + '">' + shared.map(s => '<si><t xml:space="preserve">' + xml(s) + '</t></si>').join('') + '</sst>');
    if (session.issues.length) add('Template issues', [['Sheet', 'File', 'Issue'], ...session.issues.map(i => [i.sheet, i.file, i.message])], [34, 16, 100]);
    zip.file('xl/styles.xml', auditStyles);
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="' + NS + '" xmlns:r="' + REL + '"><sheets>' + names.map((n, i) => '<sheet name="' + xml(n) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + names.map((_, i) => '<Relationship Id="rId' + (i + 1) + '" Type="' + REL + '/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') + '<Relationship Id="rIdStrings" Type="' + REL + '/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rIdStyles" Type="' + REL + '/styles" Target="styles.xml"/></Relationships>');
    zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + '/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' + names.map((_, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') + '</Types>');
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } }, m => progress?.({ stage: 'compress', message: 'Preparing audit workbook', fraction: m.percent / 100 }));
  }
  const api = { inspect, publicSummary, summarizeAudit, auditRows, auditPage, merge, exportAudit, auditHeaders, auditValues, limits, _internals: { parseCells, identityKey, zipSafety, sheetXml, xml, colName, attrs } };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.ITBBU = api;
})(typeof self !== 'undefined' ? self : globalThis);
