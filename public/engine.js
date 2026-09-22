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
    for (const m of row.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)) {
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
    const themePath = Object.values(links).find(p => /theme\/[^/]+\.xml$/.test(p));
    const w = { index, label: 'File-' + (index + 1), name: input.name, bytes: bytes.length, expanded, zip, read, book, rels, content, strings, entries, sheets: new Map(), helpers: new Map(), styles: await read('xl/styles.xml'), theme: themePath ? await read(themePath) : '', records: 0, sites: new Set() };
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
      for (const m of body[1].matchAll(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)) {
        const n = Number(attrs(m[0].match(/^<row\b[^>]*>/)[0]).r); if (n > 5) break;
        headers[n - 1] = values(parseCells(m[0], strings));
      }
      isConfig = normal(headers[0]?.[0]) === 'modind';
      if (!isConfig) { w.helpers.set(entry.name, { entry, raw }); continue; }
      const names = headers[0] || [], seen = new Set();
      for (const name of names) { if (!name || seen.has(name)) throw new Error(entry.name + ': blank or duplicate parameter name in row 1.'); seen.add(name); }
      if (headers.length !== 5 || headers.some(h => !h)) throw new Error(entry.name + ': expected five template rows.');
      const rows = [], mergeIssues = [];
      for (const m of body[1].matchAll(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)) {
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
      const sheet = { name: entry.name, entry, names, headers, rows, keys, mergeIssues, hasTemplateRules: /<dataValidation\b|<conditionalFormatting\b|<f\b/.test(raw), externalTemplateParts: /\br:id\s*=|<extLst\b/.test(raw), dynamicTemplateRefs: [...raw.matchAll(/<(?:formula1|formula2|formula|f)\b[^>]*>([\s\S]*?)<\//g)].some(m => /\bINDIRECT\s*\(|\[[^\]]+\]/i.test(m[1])) };
      w.sheets.set(entry.name, sheet); w.records += rows.length;
      if (i % 60 === 0) progress?.({ stage: 'read', message: w.label + ': reading ' + entry.name, file: index, fraction: (i + 1) / entries.length });
    }
    return w;
  }
  function helperSignature(helper, book) {
    if (!helper) return null;
    if (helper.signature !== undefined) return helper.signature;
    const records = [];
    for (const m of helper.raw.matchAll(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)) {
      const row = Number(attrs(m[0].match(/^<row\b[^>]*>/)[0]).r);
      parseCells(m[0], book.strings).forEach((cell, col) => {
        if (cell.text !== '') records.push([row, col, cell.kind, cell.text]);
      });
    }
    records.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    return helper.signature = JSON.stringify(records);
  }
  function definedNames(book) {
    const result = new Map();
    for (const m of book.book.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
      const a = attrs(m[1]);
      result.set(a.name + '\u0000' + (a.localSheetId ?? ''), unxml(m[2]));
    }
    return result;
  }
  function compareTemplate(donor, source, donorLabel, fileLabel, issues, warnings, details) {
    const positions = new Map(donor.names.map((name, i) => [name, i]));
    source.columnMap = source.names.map(name => positions.has(name) ? positions.get(name) : -1);
    let compatible = true;
    const note = (severity, parameter, row, dc, sc, before, after, action) => {
      details.push({ severity, sheet: donor.name, file: fileLabel, parameter, template: donorLabel,
        templateCell: dc < 0 ? '' : colName(dc) + row, sourceCell: sc < 0 ? '' : colName(sc) + row,
        before, after, action });
    };
    const issue = message => { compatible = false; issues.push({ sheet: donor.name, file: fileLabel, message }); };
    for (let c = 0; c < source.names.length; c++) if (source.columnMap[c] < 0) {
      const parameter = source.names[c];
      const action = 'Source parameter is absent from ' + donorLabel + '. No source column will be discarded.';
      issue('Row 1: parameter "' + parameter + '" at ' + colName(c) + '1 is absent from the widest reference (' + donorLabel + '). Use a reference containing all source parameters; no values will be dropped.');
      note('Blocked', parameter, 1, -1, c, '(missing column)', parameter, action);
    }
    const sourcePositions = new Map(source.names.map((name, i) => [name, i]));
    let missingColumns = 0;
    for (let c = 0; c < donor.names.length; c++) if (!sourcePositions.has(donor.names[c])) {
      const parameter = donor.names[c];
      missingColumns++;
      note('Notice', parameter, 1, c, -1, parameter, '(missing column)', 'Leave this parameter blank for every appended ' + fileLabel + ' row. The audit reports Missing parameter.');
    }
    if (missingColumns) warnings.push(donor.name + ' · ' + fileLabel + ': ' + missingColumns + ' parameter column(s) absent. Matching values go into ' + donorLabel + ' columns; absent parameters stay blank.');
    let metadataChanges = 0;
    for (let sc = 0; sc < source.names.length; sc++) {
      const dc = source.columnMap[sc]; if (dc < 0) continue;
      for (let h = 1; h < 5; h++) {
        const before = donor.headers[h][dc] ?? '', after = source.headers[h][sc] ?? '';
        if (before === after) continue;
        const tidy = s => String(s).trim().replace(/\r\n?/g, '\n');
        const sameRule = h === 4 ? normal(before).replace(/\s+/g, ' ') === normal(after).replace(/\s+/g, ' ') : tidy(before) === tidy(after);
        const blocked = (h === 2 || h === 4) && !sameRule;
        const parameter = source.names[sc], row = h + 1;
        const action = blocked ? (h === 2 ? 'Parameter type/range/default definitions must agree.' : 'Primary-key and requirement rules must agree.') : 'Retain ' + donorLabel + ' header text; append values by parameter name.';
        note(blocked ? 'Blocked' : 'Warning', parameter, row, dc, sc, before, after, action);
        if (blocked) {
          const short = s => JSON.stringify(s.length > 140 ? s.slice(0, 140) + '…' : s);
          issue('Parameter "' + parameter + '", row ' + row + ': ' + donorLabel + ' ' + colName(dc) + row + ' = ' + short(before) + '; ' + fileLabel + ' ' + colName(sc) + row + ' = ' + short(after) + '. ' + action);
        } else metadataChanges++;
      }
    }
    if (metadataChanges) warnings.push(donor.name + ' · ' + fileLabel + ': ' + metadataChanges + ' descriptive/formatting header difference(s). Rows 1–5 are retained from ' + donorLabel + '.');
    if (compatible && source.columnMap.some((col, i) => col !== i)) {
      const action = 'Columns are reordered by exact parameter name into ' + donorLabel + ' column order.';
      warnings.push(donor.name + ' · ' + fileLabel + ': ' + action);
      details.push({ severity: 'Warning', sheet: donor.name, file: fileLabel, template: donorLabel, action });
    }
    return compatible;
  }
  async function inspect(inputs, Zip, progress) {
    if (inputs.length < 2 || inputs.length > limits.files) throw new Error('Select between 2 and 20 .xlsx workbooks.');
    if (inputs.reduce((n, f) => n + f.bytes.byteLength, 0) > limits.inputBytes) throw new Error('Select at most 150 MiB of workbooks per batch.');
    const books = []; let expanded = 0;
    for (let i = 0; i < inputs.length; i++) { if (!/\.xlsx$/i.test(inputs[i].name)) throw new Error('Only .xlsx workbooks are supported.'); const w = await readWorkbook(inputs[i], i, Zip, progress); books.push(w); expanded += w.expanded; if (expanded > limits.expandedBytes) throw new Error('Combined expanded workbooks exceed 600 MiB. Use a smaller batch.'); }
    const names = [...new Set(books.flatMap(w => [...w.sheets.keys()]))], issues = [], warnings = [], templateDetails = [];
    const sheets = names.map(name => {
      const present = books.map(w => w.sheets.get(name)), counts = present.map(s => s?.rows.length ?? null);
      // Decide once, per MO: widest parameter template, with ties resolved in upload order.
      const referenceIndex = present.reduce((best, sheet, i) => sheet && (best < 0 || sheet.names.length > present[best].names.length) ? i : best, -1);
      const donor = present[referenceIndex], donorBook = books[referenceIndex];
      let compatible = true;
      present.forEach((s, i) => {
        if (!s && books[i].entries.some(e => e.name === name)) { compatible = false; issues.push({ sheet: name, file: books[i].label, message: 'This named sheet exists but does not contain the expected row-6 ZTE configuration template.' }); }
        if (s && !compareTemplate(donor, s, donorBook.label, books[i].label, issues, warnings, templateDetails)) compatible = false;
        for (const message of s?.mergeIssues || []) { compatible = false; issues.push({ sheet: name, file: books[i].label, message }); }
      });
      if (referenceIndex !== 0) {
        const action = 'Use ' + donorBook.label + ' as the reference with ' + donor.names.length + ' parameter columns. Preserve its five template rows and append data in file order.';
        warnings.push(name + ': ' + action);
        templateDetails.push({ severity: 'Notice', sheet: name, file: donorBook.label, template: donorBook.label, action });
        if (donor.externalTemplateParts) { compatible = false; issues.push({ sheet: name, file: donorBook.label, message: 'The selected reference contains external worksheet parts or extension rules that cannot be imported automatically.' }); }
        if (appearanceConverter(donorBook, books[0]) !== unchangedXml) {
          const action = 'Resolve the reference theme fonts and colours to explicit values when importing styles. Keep File-1\'s theme and palette unchanged.';
          warnings.push(name + ' · ' + donorBook.label + ': ' + action);
          templateDetails.push({ severity: 'Notice', sheet: name, file: donorBook.label, template: donorBook.label, action });
        }
        if (donor.hasTemplateRules) for (const [helperName, helper] of donorBook.helpers) {
          if (/<tableParts\b|<drawing\b|<hyperlinks\b|<legacyDrawing\b|\br:id\s*=|<extLst\b|<f\b/.test(helper.raw) && helperSignature(helper, donorBook) !== helperSignature(books[0].helpers.get(helperName), books[0])) { compatible = false; issues.push({ sheet: name, file: donorBook.label, message: 'Reference lookup sheet ' + helperName + ' contains formulas or linked objects that need manual preservation.' }); }
        }
        if (donor.dynamicTemplateRefs || (donor.hasTemplateRules && [...definedNames(donorBook).values()].some(formula => /\bINDIRECT\s*\(|\[[^\]]+\]/i.test(formula)))) { compatible = false; issues.push({ sheet: name, file: donorBook.label, message: 'The selected reference uses dynamic or external formula references that cannot be relocated automatically.' }); }
      }
      if (counts.reduce((n, c) => n + (c || 0), 0) > 1048571) { compatible = false; issues.push({ sheet: name, file: 'All', message: 'Merged data exceeds the Excel worksheet row limit.' }); }
      return { name, counts, columnCounts: present.map(s => s?.names.length ?? null), referenceIndex, reference: donorBook.label, total: counts.reduce((n, c) => n + (c || 0), 0), columns: donor.names.length, keys: donor.keys, compatible };
    });
    const caseNames = new Map();
    for (const name of names) { const key = normal(name); if (caseNames.has(key) && caseNames.get(key) !== name) issues.push({ sheet: name, file: 'All', message: 'Sheet names differ only in letter case. Rename them consistently before merging.' }); caseNames.set(key, name); }
    // File-1's support sheets stay intact. A later reference brings its own hidden lookup dependencies.
    const helperNames = [...new Set(books.flatMap(w => [...w.helpers.keys()]))];
    for (const name of helperNames) {
      const baseHelper = books[0].helpers.get(name);
      for (const source of books.slice(1)) {
        const sourceHelper = source.helpers.get(name);
        if (helperSignature(baseHelper, books[0]) === helperSignature(sourceHelper, source)) continue;
        const isReference = sheets.some(s => s.referenceIndex === source.index && source.sheets.get(s.name).hasTemplateRules);
        const action = isReference ? 'Retain File-1 lookup sheets; import the selected reference lookup under a separate hidden name when needed. Values are unchanged; enum codes are not translated.' : 'Retain File-1 lookup sheet and the selected MO reference rules. Review appended values against reference options; enum codes are not translated.';
        warnings.push(name + ' · ' + source.label + ': lookup content differs or is missing. ' + action);
        templateDetails.push({ severity: 'Warning', sheet: name, file: source.label, template: 'File-1', before: baseHelper ? '(present in File-1)' : '(absent from File-1)', after: sourceHelper ? '(present in source)' : '(absent from source)', action });
      }
    }
    if (!names.length) throw new Error('No ZTE configuration sheets found. Row 1 must begin with MODIND and data must start at row 6.');
    const session = { books, sheets, issues, warnings, templateDetails, Zip, audit: null };
    return session;
  }
  function publicSummary(session) {
    return { files: session.books.map(w => ({ label: w.label, name: w.name, size: w.bytes, records: w.records, sheets: w.sheets.size, sites: [...w.sites] })), sheets: session.sheets, issues: session.issues, warnings: session.warnings, templateDetails: session.templateDetails, totalRecords: session.books.reduce((n, w) => n + w.records, 0) };
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
    const pairs = [...keys].sort().map(k => {
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
      const reference = sheets[coverage.referenceIndex];
      const params = [...new Set([...(reference?.names || []), ...sheets.filter(Boolean).flatMap(s => s.names)])].filter(n => options.includeIdentity || !IDENTITY.has(normal(n)));
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
          if (rows.length > 1) hasDuplicate = true;
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
  function xmlCollection(raw, name) { return new RegExp('<' + name + '\\b[^>]*?(?:\\/>|>[\\s\\S]*?<\\/' + name + '>)').exec(raw)?.[0] || ''; }
  function styleItems(raw, collection, item) { return [...xmlCollection(raw, collection).matchAll(new RegExp('<' + item + '\\b[^>]*?(?:\\/>|>[\\s\\S]*?<\\/' + item + '>)', 'g'))].map(m => m[0]); }
  const unchangedXml = raw => raw;
  const sameAppearance = (a, b) => a.styles === b.styles && a.theme === b.theme;
  const DEFAULT_PALETTE = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
    '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 ' +
    '9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF ' +
    '00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 ' +
    '003366 339966 003300 333300 993300 993366 333399 333333').split(' ').map(c => 'FF' + c);
  function themeElement(raw, name) {
    const prefix = '(?:[\\w.-]+:)?';
    return new RegExp('<' + prefix + name + '\\b[^>]*?(?:\\/>|>[\\s\\S]*?<\\/' + prefix + name + '>)').exec(raw)?.[0] || '';
  }
  function appearanceInfo(book) {
    if (book.appearanceInfo) return book.appearanceInfo;
    // SpreadsheetML's slots are light1, dark1, light2, dark2, accents, then hyperlinks.
    const slots = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    const defaults = ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646', '0000FF', '800080'];
    const scheme = themeElement(book.theme, 'clrScheme');
    const colours = slots.map((slot, i) => {
      const raw = themeElement(scheme, slot), srgb = themeElement(raw, 'srgbClr'), system = themeElement(raw, 'sysClr');
      const colour = srgb ? attrs(srgb).val : system ? attrs(system).lastClr : null;
      return 'FF' + (/^[\da-f]{6}$/i.test(colour || '') ? colour.toUpperCase() : defaults[i]);
    });
    const fonts = {};
    for (const kind of ['major', 'minor']) {
      const raw = themeElement(book.theme, kind + 'Font'), scripts = {};
      for (const m of raw.matchAll(/<(?:[\w.-]+:)?font\b[^>]*\/>/g)) { const a = attrs(m[0]); scripts[a.script] = a.typeface; }
      fonts[kind] = { latin: attrs(themeElement(raw, 'latin')).typeface, scripts };
    }
    const custom = [...xmlCollection(book.styles, 'indexedColors').matchAll(/<rgbColor\b[^>]*\/>/g)].map(m => attrs(m[0]).rgb);
    const palette = DEFAULT_PALETTE.map((value, i) => custom[i] ? 'FF' + custom[i].slice(-6).toUpperCase() : value);
    return book.appearanceInfo = { colours, fonts, palette };
  }
  function appearanceConverter(source, base) {
    const themeChanged = source.theme !== base.theme, from = appearanceInfo(source), to = appearanceInfo(base);
    const paletteChanged = from.palette.some((c, i) => c !== to.palette[i]);
    if (!themeChanged && !paletteChanged) return unchangedXml;
    return raw => {
      // Keep tint on the explicit RGB colour, so Excel applies its original HLS adjustment.
      raw = raw.replace(/<(?:color|fgColor|bgColor|tabColor)\b[^>]*\/>/g, tag => {
        const a = attrs(tag); let rgb;
        if (themeChanged && a.theme !== undefined) rgb = from.colours[Number(a.theme)];
        else if (paletteChanged && a.indexed !== undefined && Number(a.indexed) < 64) rgb = from.palette[Number(a.indexed)];
        if (!rgb) return tag;
        tag = tag.replace(/\s(?:theme|indexed|auto|rgb)=(?:"[^"]*"|'[^']*')/g, '');
        return setAttr(tag, 'rgb', rgb);
      });
      if (!themeChanged) return raw;
      return raw.replace(/<(font|rPr)\b[^>]*>[\s\S]*?<\/\1>/g, font => {
        const scheme = /<scheme\b[^>]*\/>/.exec(font)?.[0], kind = scheme && attrs(scheme).val;
        if (!['major', 'minor'].includes(kind)) return font;
        const nameTag = /<(name|rFont)\b[^>]*\/>/.exec(font), stored = nameTag && attrs(nameTag[0]).val;
        const charset = attrs(/<charset\b[^>]*\/>/.exec(font)?.[0] || '').val;
        const script = { 128: 'Jpan', 129: 'Hang', 130: 'Hang', 134: 'Hans', 136: 'Hant', 163: 'Viet', 177: 'Hebr', 178: 'Arab', 222: 'Thai' }[charset];
        const face = (script && from.fonts[kind].scripts[script]) || (charset && charset !== '0' ? stored : from.fonts[kind].latin) || stored || (kind === 'major' ? 'Cambria' : 'Calibri');
        font = font.replace(scheme, '');
        if (nameTag) return font.replace(nameTag[0], () => setAttr(nameTag[0], 'val', face));
        const tag = font.startsWith('<rPr') ? 'rFont' : 'name';
        return font.replace(/^(<(?:font|rPr)\b[^>]*>)/, (_, start) => start + '<' + tag + ' val="' + xml(face) + '"/>');
      });
    };
  }
  function styleImporter(base) {
    let raw = base.styles;
    const styleKey = source => source.styles + '\u0000' + source.theme;
    const cache = new Map([[styleKey(base), { cells: null, dxfs: null, format: unchangedXml }]]), converters = new Map([[base, unchangedXml]]);
    const formatFor = source => { if (!converters.has(source)) converters.set(source, appearanceConverter(source, base)); return converters.get(source); };
    const order = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
    const write = (collection, items) => {
      const existing = xmlCollection(raw, collection), value = '<' + collection + ' count="' + items.length + '">' + items.join('') + '</' + collection + '>';
      if (existing) raw = raw.replace(existing, () => value);
      else {
        const next = order.slice(order.indexOf(collection) + 1).map(n => xmlCollection(raw, n)).find(Boolean);
        raw = next ? raw.replace(next, () => value + next) : raw.replace('</styleSheet>', () => value + '</styleSheet>');
      }
    };
    const append = (source, collection, item, change = s => s) => {
      const items = styleItems(raw, collection, item), ids = new Map(items.map((value, i) => [value, i]));
      const mapping = styleItems(source.styles, collection, item).map(value => {
        value = change(value);
        if (!ids.has(value)) { ids.set(value, items.length); items.push(value); }
        return ids.get(value);
      });
      if (mapping.length) write(collection, items);
      return mapping;
    };
    const importBook = source => {
      const key = styleKey(source); if (cache.has(key)) return cache.get(key);
      const format = formatFor(source), from = appearanceInfo(source), to = appearanceInfo(base);
      const formats = styleItems(raw, 'numFmts', 'numFmt'), byCode = new Map(formats.map(s => { const a = attrs(s); return [a.formatCode, Number(a.numFmtId)]; })), numFormats = new Map();
      let nextFormat = Math.max(163, ...formats.map(s => Number(attrs(s).numFmtId)));
      for (const f of styleItems(source.styles, 'numFmts', 'numFmt')) {
        const a = attrs(f);
        const code = a.formatCode.replace(/"(?:[^"]|"")*"|\\.|\[Color(\d+)\]/gi, (token, index) => {
          if (index === undefined || from.palette[Number(index) + 7] === to.palette[Number(index) + 7]) return token;
          const destination = to.palette.findIndex((colour, i) => i >= 8 && colour === from.palette[Number(index) + 7]);
          if (destination < 0) throw new Error(source.label + ': custom number format uses a palette colour unavailable in the base workbook.');
          return '[Color' + (destination - 7) + ']';
        });
        if (!byCode.has(code)) { byCode.set(code, ++nextFormat); formats.push(setAttr(setAttr(f, 'formatCode', code), 'numFmtId', nextFormat)); }
        numFormats.set(Number(a.numFmtId), byCode.get(code));
      }
      if (numFormats.size) write('numFmts', formats);
      const fonts = append(source, 'fonts', 'font', format), fills = append(source, 'fills', 'fill', format), borders = append(source, 'borders', 'border', format);
      const remap = (value, mappings) => value.replace(/^<\w+\b[^>]*>/, tag => {
        const a = attrs(tag);
        for (const [key, map] of Object.entries(mappings)) if (a[key] !== undefined) {
          const id = Number(a[key]), mapped = map instanceof Map ? (map.get(id) ?? id) : (map[id] ?? id);
          tag = setAttr(tag, key, mapped);
        }
        return tag;
      });
      const mappings = { fontId: fonts, fillId: fills, borderId: borders, numFmtId: numFormats };
      const xfs = append(source, 'cellStyleXfs', 'xf', value => remap(value, mappings));
      const cells = append(source, 'cellXfs', 'xf', value => remap(value, { ...mappings, xfId: xfs }));
      const dxfs = append(source, 'dxfs', 'dxf', value => format(value.replace(/<numFmt\b[^>]*\/>/g, tag => remap(tag, { numFmtId: numFormats }))));
      const result = { cells, dxfs, format }; cache.set(key, result); return result;
    };
    return { importBook, formatFor, value: () => raw };
  }
  function mapStyle(id, mapping) {
    if (!mapping) return id;
    const mapped = mapping[Number(id)];
    if (mapped === undefined) throw new Error('A reference cell uses an invalid style index: ' + id + '.');
    return String(mapped);
  }
  function mapSheetStyles(raw, maps) {
    return maps.format(raw).replace(/<(?:c|row|col|cfRule)\b[^>]*>/g, tag => {
      const a = attrs(tag), key = tag.startsWith('<col ') ? 'style' : tag.startsWith('<cfRule ') ? 'dxfId' : 's';
      return a[key] === undefined ? tag : setAttr(tag, key, mapStyle(a[key], key === 'dxfId' ? maps.dxfs : maps.cells));
    });
  }
  function rewriteFormula(formula, sheets, names) {
    // Only reference tokens change. Quoted Excel text, cell addresses and user values stay literal.
    return formula.replace(/"(?:[^"]|"")*"|'(?:[^']|'')+'!|[\p{L}\p{N}_.$]+!|[\p{L}_\\][\p{L}\p{N}_.\\]*/gu, token => {
      if (token.startsWith('"')) return token;
      if (token.endsWith('!')) {
        const name = token.slice(0, -1).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
        return sheets.has(normal(name)) ? "'" + sheets.get(normal(name)).replace(/'/g, "''") + "'!" : token;
      }
      return names.get(normal(token)) || token;
    });
  }
  function rewriteSheetFormulas(raw, sheets, names) {
    return raw.replace(/<(formula1|formula2|formula|f)\b([^>]*)>([\s\S]*?)<\/\1>/g, (_, tag, a, text) => '<' + tag + a + '>' + xml(rewriteFormula(unxml(text), sheets, names)) + '</' + tag + '>');
  }
  function scopeRowNamespaces(row, sourceRoot) {
    // Excel adds extension attributes such as x14ac:dyDescent to ordinary rows.
    // Keep their original namespace bindings when moving into an older template.
    const used = new Set();
    for (const m of row.matchAll(/<\/?[\w:.-]+\b[^>]*>/g)) {
      const tag = m[0], prefix = /^<\/?([\w.-]+):/.exec(tag)?.[1];
      if (prefix) used.add(prefix);
      const a = attrs(tag);
      for (const key of Object.keys(a)) if (key.includes(':') && !key.startsWith('xmlns:')) used.add(key.split(':')[0]);
      for (const p of (a['mc:Ignorable'] || '').split(/\s+/).filter(Boolean)) used.add(p);
    }
    if (!used.size) return row;
    let opening = row.match(/^<row\b[^>]*>/)[0];
    const original = opening, local = attrs(opening);
    const ignored = (sourceRoot['mc:Ignorable'] || '').split(/\s+/).filter(p => used.has(p));
    if (ignored.length) used.add('mc');
    for (const prefix of used) {
      const key = 'xmlns:' + prefix;
      if (prefix !== 'xml' && sourceRoot[key] && !local[key]) opening = setAttr(opening, key, sourceRoot[key]);
    }
    if (ignored.length) opening = setAttr(opening, 'mc:Ignorable', [...new Set([...(local['mc:Ignorable'] || '').split(/\s+/).filter(Boolean), ...ignored])].join(' '));
    return opening + row.slice(original.length);
  }
  function transformRow(row, newNumber, source, base, baseStyle, columnMap, styleMap = null, width = 0, format = unchangedXml) {
    const originalTag = row.match(/^<row\b[^>]*>/)[0];
    let rowTag = setAttr(originalTag, 'r', newNumber);
    if (rowTag.endsWith('/>')) return rowTag;
    if (columnMap && attrs(rowTag).spans) rowTag = setAttr(rowTag, 'spans', '1:' + (width || Math.max(...columnMap) + 1));
    if (attrs(rowTag).s !== undefined) rowTag = setAttr(rowTag, 's', sameAppearance(source, base) ? mapStyle(attrs(rowTag).s, styleMap) : '0');
    const cells = [];
    const remainder = row.slice(originalTag.length).replace(/<\/row>$/, '').replace(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, cell => {
      let tag = cell.match(/^<c\b[^>]*>/)[0], a = attrs(tag);
      const sourceCol = (a.r || '').match(/^[A-Z]+/i)?.[0];
      const destination = columnMap ? columnMap[colNum(sourceCol)] : colNum(sourceCol);
      if (!Number.isInteger(destination) || destination < 0) throw new Error('No destination parameter column for ' + a.r + '.');
      const col = colName(destination);
      tag = setAttr(tag, 'r', col + newNumber);
      if (!sameAppearance(source, base)) { const style = baseStyle[col]; if (style == null) throw new Error('Data-cell style cannot be preserved for column ' + col + '. Use matching templates.'); tag = setAttr(tag, 's', style); }
      else if (a.s !== undefined || styleMap) tag = setAttr(tag, 's', mapStyle(a.s || '0', styleMap));
      let body = cell.slice(cell.indexOf('>') + 1);
      if (a.t === 's') { const v = /<v>(\d+)<\/v>/.exec(body), str = v && source.strings[Number(v[1])]; if (!str) throw new Error('Invalid shared string.'); tag = setAttr(tag, 't', 'inlineStr'); body = '<is>' + str.inner + '</is></c>'; }
      body = format(body);
      cells.push({ column: destination, xml: tag + (tag.endsWith('/>') ? '' : body) });
      return '';
    });
    if (width) {
      const present = new Set(cells.map(c => c.column));
      for (let c = 0; c < width; c++) if (!present.has(c)) cells.push({ column: c, xml: '<c r="' + colName(c) + newNumber + '" s="' + (baseStyle[colName(c)] || '0') + '"/>' });
    }
    cells.sort((a, b) => a.column - b.column);
    return rowTag + cells.map(c => c.xml).join('') + remainder + '</row>';
  }
  async function merge(session, progress) {
    if (session.issues.length) throw new Error('Resolve template compatibility issues before merging. The audit remains available.');
    const base = session.books[0], output = new session.Zip();
    // Reuse compressed ZIP objects for untouched entries, avoiding a lossy workbook round trip.
    for (const [name, file] of Object.entries(base.zip.files)) output.files[name] = file;
    let book = base.book, rels = base.rels, content = base.content, added = 0;
    let id = Math.max(0, ...base.entries.map(e => Number(e.sheetId)));
    const usedPaths = new Set(Object.keys(output.files));
    const usedNames = new Set([...base.entries.map(e => normal(e.name)), ...session.sheets.map(s => normal(s.name))]);
    const sheetOrder = base.entries.map(e => e.name), styles = styleImporter(base), dependencies = new Map();
    const usedDefinedNames = new Set([...definedNames(base).keys()].map(k => normal(k.split('\u0000')[0])));
    const uniqueName = (stem, used, limit) => {
      let name = stem.slice(0, limit), suffix = 1;
      while (used.has(normal(name))) { const ending = '_' + (++suffix); name = stem.slice(0, limit - ending.length) + ending; }
      used.add(normal(name)); return name;
    };
    const addSheet = (name, state = '') => {
      let path, rid;
      do { path = 'xl/worksheets/merged' + (++added) + '.xml'; rid = 'rIdITBBU' + added; } while (usedPaths.has(path) || rels.includes('Id="' + rid + '"'));
      usedPaths.add(path);
      book = book.replace('</sheets>', () => '<sheet name="' + xml(name) + '" sheetId="' + (++id) + '" r:id="' + rid + '"' + (state ? ' state="' + xml(state) + '"' : '') + '/></sheets>');
      rels = rels.replace('</Relationships>', () => '<Relationship Id="' + rid + '" Type="' + REL + '/worksheet" Target="' + path.slice(3) + '"/></Relationships>');
      content = content.replace('</Types>', () => '<Override PartName="/' + path + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
      sheetOrder.push(name); return { path, index: sheetOrder.length - 1 };
    };
    const addDefinedNames = tags => {
      if (!tags.length) return;
      if (/<definedNames\b[^>]*\/>/.test(book)) book = book.replace(/<definedNames\b[^>]*\/>/, () => '<definedNames>' + tags.join('') + '</definedNames>');
      else if (book.includes('</definedNames>')) book = book.replace('</definedNames>', () => tags.join('') + '</definedNames>');
      else book = book.replace('</sheets>', () => '</sheets><definedNames>' + tags.join('') + '</definedNames>');
    };
    const sourceNames = donor => [...donor.book.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)].map(m => ({ tag: '<definedName' + m[1] + '>', a: attrs(m[1]), formula: unxml(m[2]) }));
    const importLocalNames = (donor, sourceName, outputIndex, resources) => {
      const sourceIndex = donor.entries.findIndex(e => e.name === sourceName);
      // A replacement template also owns its scoped print areas and local names.
      book = book.replace(/<definedName\b[^>]*>[\s\S]*?<\/definedName>/g, tag => attrs(tag.match(/^<definedName\b[^>]*>/)[0]).localSheetId === String(outputIndex) ? '' : tag);
      const local = sourceNames(donor).filter(n => n.a.localSheetId === String(sourceIndex));
      const nameMap = new Map(resources.names); for (const n of local) nameMap.set(normal(n.a.name), n.a.name);
      addDefinedNames(local.map(n => setAttr(n.tag, 'localSheetId', outputIndex) + xml(rewriteFormula(n.formula, resources.sheets, nameMap)) + '</definedName>'));
      return nameMap;
    };
    const importDependencies = async donor => {
      if (dependencies.has(donor.index)) return dependencies.get(donor.index);
      const resources = { sheets: new Map(), names: new Map() }, copies = [], maps = styles.importBook(donor);
      for (const [name, helper] of donor.helpers) {
        const retained = base.helpers.get(name);
        if (retained && helperSignature(helper, donor) === helperSignature(retained, base)) continue;
        const renamed = uniqueName('__ITBBU_F' + (donor.index + 1) + '_' + name, usedNames, 31);
        resources.sheets.set(normal(name), renamed);
        copies.push({ helper, name, ...addSheet(renamed, 'hidden') });
      }
      const globals = sourceNames(donor).filter(n => n.a.localSheetId === undefined);
      for (const n of globals) resources.names.set(normal(n.a.name), uniqueName('ITBBU_F' + (donor.index + 1) + '_' + n.a.name, usedDefinedNames, 255));
      addDefinedNames(globals.map(n => setAttr(n.tag, 'name', resources.names.get(normal(n.a.name))) + xml(rewriteFormula(n.formula, resources.sheets, resources.names)) + '</definedName>'));
      for (const copy of copies) {
        const localNames = importLocalNames(donor, copy.name, copy.index, resources);
        let raw = copy.helper.raw.replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, row => transformRow(row, Number(attrs(row.match(/^<row\b[^>]*>/)[0]).r), donor, donor, {}));
        raw = rewriteSheetFormulas(mapSheetStyles(raw, maps), resources.sheets, localNames);
        output.file(copy.path, raw);
      }
      dependencies.set(donor.index, resources); return resources;
    };
    for (let i = 0; i < session.sheets.length; i++) {
      const coverage = session.sheets[i], name = coverage.name, donor = session.books[coverage.referenceIndex], target = donor.sheets.get(name);
      const maps = styles.importBook(donor), existing = base.sheets.get(name);
      const destination = existing ? { path: existing.entry.path, index: base.entries.findIndex(e => e.name === name) } : addSheet(name, target.entry.state);
      let resources = { sheets: new Map(), names: new Map() };
      if (donor !== base) {
        if (target.hasTemplateRules) resources = await importDependencies(donor);
        resources = { sheets: resources.sheets, names: importLocalNames(donor, name, destination.index, resources) };
      }
      let raw = await donor.read(target.entry.path);
      const originalBody = /<sheetData(?:\s[^>]*)?>([\s\S]*?)<\/sheetData>/.exec(raw);
      const header = [...originalBody[1].matchAll(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)].filter(m => Number(attrs(m[0].match(/^<row\b[^>]*>/)[0]).r) < 6).map(m => m[0]).join('\n');
      const style = Object.fromEntries(target.names.map((_, c) => [colName(c), '0']));
      for (const m of raw.matchAll(/<col\b[^>]*\/>/g)) { const a = attrs(m[0]); for (let c = Number(a.min) - 1; c < Math.min(Number(a.max), target.names.length); c++) style[colName(c)] = a.style || '0'; }
      const prototype = [...originalBody[1].matchAll(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)].find(m => Number(attrs(m[0].match(/^<row\b[^>]*>/)[0]).r) >= 6)?.[0];
      for (const m of (prototype || '').matchAll(/<c\b[^>]*>/g)) { const a = attrs(m[0]); style[(a.r || '').replace(/\d+/g, '')] = a.s || '0'; }
      for (const col of Object.keys(style)) style[col] = mapStyle(style[col], maps.cells);
      const data = []; let rowNumber = 6;
      for (const w of session.books) {
        const sheet = w.sheets.get(name); if (!sheet || !sheet.rows.length) continue;
        const source = w === donor ? raw : await w.read(sheet.entry.path), body = /<sheetData(?:\s[^>]*)?>([\s\S]*?)<\/sheetData>/.exec(source)[1], keep = new Set(sheet.rows.map(r => r.number));
        const sourceRoot = w === donor ? null : attrs(source.match(/<worksheet\b[^>]*>/)[0]);
        for (const m of body.matchAll(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)) {
          const n = Number(attrs(m[0].match(/^<row\b[^>]*>/)[0]).r);
          if (keep.has(n)) {
            const row = transformRow(m[0], rowNumber++, w, donor, style, sheet.columnMap, maps.cells, target.names.length, styles.formatFor(w));
            data.push(sourceRoot ? scopeRowNamespaces(row, sourceRoot) : row);
          }
        }
      }
      const safeHeader = donor === base ? header : header.replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, r => transformRow(r, Number(attrs(r.match(/^<row\b[^>]*>/)[0]).r), donor, donor, {}, null, maps.cells, 0, maps.format));
      raw = mapSheetStyles(raw.slice(0, originalBody.index), maps) + '<sheetData>\n' + safeHeader + '\n' + data.join('\n') + '\n</sheetData>' + mapSheetStyles(raw.slice(originalBody.index + originalBody[0].length), maps);
      if (donor !== base) raw = rewriteSheetFormulas(raw, resources.sheets, resources.names);
      raw = raw.replace(/<dimension\b[^>]*\/>/, '<dimension ref="A1:' + colName(target.names.length - 1) + Math.max(5, rowNumber - 1) + '"/>');
      // Extend existing validation / conditional-format ranges only where appended rows exceed them.
      raw = raw.replace(/<(?:dataValidation|conditionalFormatting)\b[^>]*>/g, tag => {
        const a = attrs(tag); if (!a.sqref) return tag;
        const refs = a.sqref.split(' ').map(r => r.replace(/^(\$?[A-Z]+\$?)(\d+):(\$?[A-Z]+\$?)(\d+)$/, (_, ac, ar, bc, br) => Number(ar) >= 6 && Number(br) < rowNumber - 1 ? ac + ar + ':' + bc + (rowNumber - 1) : _));
        return setAttr(tag, 'sqref', refs.join(' '));
      });
      output.file(destination.path, raw);
      if (i % 35 === 0) progress?.({ stage: 'merge', message: 'Merging ' + name, fraction: i / session.sheets.length });
    }
    if (styles.value() !== base.styles) output.file('xl/styles.xml', styles.value());
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
    add('Sheet coverage', [['Sheet', ...session.books.flatMap(w => [w.label + ' rows', w.label + ' columns']), 'Merged rows', 'Reference template', 'Parameters', 'Audit comparisons', 'Different', 'Missing', 'Ambiguous / no key'], ...session.sheets.map(s => { const k = a.bySheet[s.name] || {}; return [s.name, ...s.counts.flatMap((n, i) => [n === null ? 'Missing sheet' : n, s.columnCounts[i] ?? 'Missing sheet']), s.total, s.reference, s.columns, k.total || 0, k.different || 0, k.missing || 0, k.ambiguous || 0]; })], [34, ...session.books.flatMap(() => [20, 20]), 20, 22, 18, 24, 18, 18, 26]);
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
    if (session.templateDetails?.length) add('Template details', [['Severity', 'Sheet', 'Source file', 'Parameter', 'Template file', 'Template cell', 'Source cell', 'Template value', 'Source value', 'Action'], ...session.templateDetails.map(d => [d.severity, d.sheet, d.file, d.parameter || '', d.template || '', d.templateCell || '', d.sourceCell || '', d.before || '', d.after || '', d.action])], [16, 32, 16, 30, 18, 18, 18, 65, 65, 100]);
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
