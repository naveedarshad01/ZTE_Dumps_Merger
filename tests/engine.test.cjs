const { test } = require('node:test');
const assert = require('node:assert/strict');
const E = require('../public/engine.js');
const Zip = require('../public/vendor/jszip.min.js');
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const names = ['MODIND', 'ManagedElementType', 'SubNetwork', 'ManagedElement', 'NE_Name', 'ldn', 'moId', 'power'];
const template = [names, ['Operation','NE type','Subnetwork','NE ID','Site','LDN','MO ID','Power'], ['string','string','string','string','string','string','string','string'], ['','','','','','','','Exact parameter'], ['','Primary Key','Primary Key','Primary Key','R','Primary Key','M','--']];
const row = (site, ldn, power, id = '1') => ['', 'ITBBU', '1', site, 'SITE-' + site, ldn, id, power];
async function fixture(name, sheets, opts = {}) {
 const z = new Zip(), keys = [...Object.keys(sheets), ...Object.keys(opts.helpers || {})];
 z.file('[Content_Types].xml',`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${keys.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}${opts.shared ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' : ''}</Types>`);
 z.file('_rels/.rels',`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="xl/workbook.xml" Type="${R}/officeDocument"/></Relationships>`);
 z.file('xl/styles.xml',opts.styles || '<styleSheet xmlns="'+NS+'"><fonts count="1"><font><name val="Arial"/><sz val="11"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
 z.file('xl/workbook.xml',`<workbook xmlns="${NS}" xmlns:r="${R}"><sheets>${keys.map((n,i)=>`<sheet name="${n}" sheetId="${i+1}" r:id="rId${i+1}"${opts.helpers?.[n] ? ' state="hidden"' : ''}/>`).join('')}</sheets>${opts.definedNames ? '<definedNames>'+opts.definedNames+'</definedNames>' : ''}</workbook>`);
 z.file('xl/_rels/workbook.xml.rels',`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${keys.map((n,i)=>`<Relationship Id="rId${i+1}" Target="worksheets/sheet${i+1}.xml" Type="${R}/worksheet"/>`).join('')}<Relationship Id="rIdStyles" Target="styles.xml" Type="${R}/styles"/>${opts.shared ? `<Relationship Id="rIdS" Target="sharedStrings.xml" Type="${R}/sharedStrings"/>` : ''}</Relationships>`);
 keys.forEach((key,i)=>{
   if (opts.helpers?.[key]) { z.file(`xl/worksheets/sheet${i+1}.xml`,opts.helpers[key]); return; }
   let grid = [...(opts.template || template).map(r=>[...r]),...sheets[key]];
   for (const [rowIndex, cells] of Object.entries(opts.headerCells || {})) for (const [column, value] of Object.entries(cells)) grid[Number(rowIndex)-1][Number(column)] = value;
   const order = Array.isArray(opts.order) ? opts.order : opts.order?.[key];
   if(order) grid=grid.map(r=>order.map(c=>r[c] ?? ''));
   let xml = E._internals.sheetXml(grid,grid[0].map(()=>20),false);
   if(opts.headerMismatch && i===0) xml=xml.replace('Exact parameter','Changed metadata');
   if(opts.shared) xml=xml.replace(/<c r="([A-Z]+6)" s="0" t="inlineStr"><is><t xml:space="preserve">shared value<\/t><\/is><\/c>/,'<c r="$1" s="0" t="s"><v>0</v></c>');
   if(opts.formula) xml=xml.replace('<c r="H6" s="0" t="inlineStr"><is><t xml:space="preserve">formula</t></is></c>','<c r="H6" s="0"><f>1+1</f><v>2</v></c>');
   if(opts.validation) xml=xml.replace('</worksheet>',opts.validation+'</worksheet>');
   if(opts.mutateXml) xml=opts.mutateXml(xml,key);
   z.file(`xl/worksheets/sheet${i+1}.xml`,xml);
 });
 if(opts.shared) z.file('xl/sharedStrings.xml',`<sst xmlns="${NS}">${(opts.strings || ['shared value']).map(s=>'<si><t>'+E._internals.xml(s)+'</t></si>').join('')}</sst>`);
 return {name:name+'.xlsx',bytes:await z.generateAsync({type:'uint8array',compression:'DEFLATE'})};
}
const inspect = async files => E.inspect(files,Zip);

test('three files merge in upload order, preserve headers and match reordered rows by key',async()=>{
 const a=await fixture('A',{Equipment:[row('100','Equipment=1','1'),row('100','Equipment=2','0')]}),b=await fixture('B',{Equipment:[row('200','Equipment=2','0'),row('200','Equipment=1','2')]}),c=await fixture('C',{Equipment:[row('300','Equipment=1','3'),row('300','Equipment=2','0')]});
 const s=await inspect([a,b,c]); assert.equal(s.issues.length,0);
 const rows=[...E.auditRows(s)];assert.equal(rows.length,2);assert.deepEqual(rows[0].values,['1','2','3']);assert.equal(rows[0].status,'Different');assert.equal(rows[1].status,'Same');
 const merged=await Zip.loadAsync(await E.merge(s)),orig=await Zip.loadAsync(a.bytes),raw=await merged.file('xl/worksheets/sheet1.xml').async('string'),before=await orig.file('xl/worksheets/sheet1.xml').async('string');
 for(let i=1;i<=5;i++)assert.equal(raw.match(new RegExp(`<row r="${i}"[^>]*>[\\s\\S]*?<\\/row>`))[0],before.match(new RegExp(`<row r="${i}"[^>]*>[\\s\\S]*?<\\/row>`))[0]);
 assert.equal((raw.match(/<row /g)||[]).length,11);assert.match(raw,/<dimension ref="A1:H11"/);
 const parsed=[...raw.matchAll(/<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g)].filter(m=>+m[1]>=6).map(m=>E._internals.parseCells(m[0],[])[3].text);
 assert.deepEqual(parsed,['100','100','200','200','300','300']);
});
test('blank differs from zero, literal markers, cell types and missing records',async()=>{
 const s=await inspect([await fixture('A',{Equipment:[row('1','Equipment=1',''),row('1','Equipment=2','(blank)'),row('1','Equipment=3',1)]}),await fixture('B',{Equipment:[row('2','Equipment=1','0'),row('2','Equipment=2',''),row('2','Equipment=3','1'),row('2','Equipment=4','5')]})]);
 const rows=[...E.auditRows(s)];assert.deepEqual(rows.map(r=>r.status),['Different','Different','Different','Missing object']);assert.equal(rows[0].values[0],'(blank)');assert.equal(rows[1].values[0],'[text] (blank)');assert.equal(rows[2].values[0],'1 [number]');
});
test('duplicate and missing keys are explicit; strict keys distinguish sites',async()=>{
 const s=await inspect([await fixture('A',{Equipment:[row('1','Equipment=1','1'),row('9','Equipment=1','2'),row('1','','3')]}),await fixture('B',{Equipment:[row('2','Equipment=1','9')]})]);
 const rows=[...E.auditRows(s)];assert.equal(rows[0].status,'Ambiguous key');assert.equal(rows[1].status,'No key');assert.match(rows[0].values[0],/row 6: 1\nrow 7: 2/);
 assert.equal([...E.auditRows(s,{mode:'strict'})].filter(r=>r.status==='Missing object').length,3);
});
test('site-root normalization changes only the known function root suffix',async()=>{
 const s=await inspect([await fixture('A',{Cell:[row('100','ENBCUCPFunction=621-20_100,Cell=100','1')]}),await fixture('B',{Cell:[row('200','ENBCUCPFunction=621-20_200,Cell=100','2')]})]);
 assert.equal([...E.auditRows(s)].length,1);assert.equal([...E.auditRows(s)][0].status,'Different');assert.match([...E.auditRows(s)][0].object,/\{site\},Cell=100/);
 assert.deepEqual([...E.auditRows(s)][0].ldns,['ENBCUCPFunction=621-20_100,Cell=100','ENBCUCPFunction=621-20_200,Cell=100']);
 assert.equal([...E.auditRows(s,{normalizeRoots:false})].length,2);
});
test('shared strings resolve correctly when copied into the destination',async()=>{
 const s=await inspect([await fixture('A',{Equipment:[row('1','Equipment=1','start')]}),await fixture('B',{Equipment:[row('2','Equipment=1','shared value')]},{shared:true})]);
 assert.equal([...E.auditRows(s)][0].values[1],'shared value');
 const z=await Zip.loadAsync(await E.merge(s));assert.match(await z.file('xl/worksheets/sheet1.xml').async('string'),/<c r="H7"[^>]*t="inlineStr"><is><t>shared value<\/t>/);
});
test('missing sheets are included in the merge union and reported in the audit',async()=>{
 const s=await inspect([await fixture('A',{Equipment:[row('1','Equipment=1','1')]}),await fixture('B',{Equipment:[],Extra:[row('2','Extra=1','2')]})]);
 assert.equal(s.sheets.length,2);assert.equal([...E.auditRows(s)].find(r=>r.sheet==='Extra').status,'Missing sheet');
 const z=await Zip.loadAsync(await E.merge(s));assert.match(await z.file('xl/workbook.xml').async('string'),/name="Extra"/);assert.ok(z.file('xl/worksheets/merged1.xml'));
});
test('description differences merge with File-1 headers intact and appear in the audit details',async()=>{
 const a=await fixture('A',{Equipment:[row('1','Equipment=1','1')]}),b=await fixture('B',{Equipment:[row('2','Equipment=1','2')]},{headerMismatch:true});
 const s=await inspect([a,b]);assert.equal(s.issues.length,0);assert.equal(s.warnings.length,1);assert.equal([...E.auditRows(s)].length,1);
 const merged=await Zip.loadAsync(await E.merge(s)),base=await Zip.loadAsync(a.bytes);
 const raw=await merged.file('xl/worksheets/sheet1.xml').async('string'),before=await base.file('xl/worksheets/sheet1.xml').async('string');
 for(let i=1;i<=5;i++)assert.equal(raw.match(new RegExp(`<row r="${i}"[^>]*>[\\s\\S]*?<\\/row>`))[0],before.match(new RegExp(`<row r="${i}"[^>]*>[\\s\\S]*?<\\/row>`))[0]);
 assert.equal((raw.match(/<row /g)||[]).length,7);assert.equal(s.templateDetails[0].parameter,'power');assert.equal(s.templateDetails[0].sourceCell,'H4');
 const audit=await Zip.loadAsync(await E.exportAudit(s));assert.match(await audit.file('xl/workbook.xml').async('string'),/name="Template details"/);
 const details=await audit.file('xl/worksheets/sheet5.xml').async('string');assert.match(details,/Changed metadata/);assert.match(details,/Exact parameter/);
});
test('different parameter sets, type/range and key rules still block and identify the exact cells',async()=>{
 const a=await fixture('A',{Equipment:[row('1','Equipment=1','1')]});
 for (const [headerRow,value] of [[1,'newPower'],[3,'integer:0..10'],[5,'Primary Key']]) {
   const s=await inspect([a,await fixture('B',{Equipment:[row('2','Equipment=1','2')]},{headerCells:{[headerRow]:{7:value}}})]);
   assert.ok(s.issues.length);assert.match(s.issues[0].message,/H[135]/);assert.ok(s.templateDetails.some(d=>d.sourceCell==='H'+headerRow && d.after===value && d.severity==='Blocked'));
   assert.ok([...E.auditRows(s)].length);await assert.rejects(()=>E.merge(s),/compatibility/);
 }
 const s=await inspect([a,await fixture('F',{Equipment:[row('2','Equipment=1','formula')]},{formula:true})]);assert.match(s.issues[0].message,/Formula/);await assert.rejects(()=>E.merge(s),/compatibility/);
});
test('three differently ordered column layouts map all values and strict keys by parameter name',async()=>{
 const a=await fixture('A',{Equipment:[row('100','Equipment=1','base')]}),b=await fixture('B',{Equipment:[row('100','Equipment=1','shared value')]},{order:[0,1,7,5,4,6,3,2],shared:true}),c=await fixture('C',{Equipment:[row('100','Equipment=1',0)]},{order:[0,7,3,6,5,2,1,4]});
 const s=await inspect([a,b,c]);assert.equal(s.issues.length,0);assert.equal(s.warnings.length,2);
 const audited=[...E.auditRows(s,{mode:'strict'})];assert.equal(audited.length,1);assert.deepEqual(audited[0].values,['base','shared value','0 [number]']);
 const z=await Zip.loadAsync(await E.merge(s));const raw=await z.file('xl/worksheets/sheet1.xml').async('string');
 const rows=[...raw.matchAll(/<row r="([678])"[^>]*>[\s\S]*?<\/row>/g)];
 assert.deepEqual(rows.map(m=>[...m[0].matchAll(/<c r="([A-Z]+)\d+"/g)].map(c=>c[1])),Array(3).fill(['A','B','C','D','E','F','G','H']));
 const vals=rows.map(m=>E._internals.parseCells(m[0],[]).map(c=>c?.text));
 assert.deepEqual(vals,[row('100','Equipment=1','base'),row('100','Equipment=1','shared value'),row('100','Equipment=1','0')]);
});
test('reordered data uses the destination column style when source styles differ',async()=>{
 const a=await fixture('A',{Equipment:[row('1','Equipment=1','base')]},{styles:'<styleSheet xmlns="'+NS+'"><cellXfs count="2"><xf/><xf numFmtId="49"/></cellXfs></styleSheet>',mutateXml:raw=>raw.replace('<c r="H6" s="0"','<c r="H6" s="1"')});
 const b=await fixture('B',{Equipment:[row('2','Equipment=1','copied')]},{order:[0,1,7,5,4,6,3,2]});
 const s=await inspect([a,b]);assert.equal(s.issues.length,0);
 const z=await Zip.loadAsync(await E.merge(s));assert.match(await z.file('xl/worksheets/sheet1.xml').async('string'),/<c r="H7" s="1"[^>]*><is><t[^>]*>copied<\/t>/);
});
const helperXml = value => E._internals.sheetXml([[value]], [20], false);
const validation = '<dataValidations count="1"><dataValidation type="list" sqref="H6:H6"><formula1>PowerOptions</formula1></dataValidation></dataValidations>';
const powerName = '<definedName name="PowerOptions">HideEnum!$A$1</definedName>';
test('later lookup changes retain File-1 helper, named ranges and validation definitions',async()=>{
 const a=await fixture('A',{Equipment:[row('1','Equipment=1','1')]},{helpers:{HideEnum:helperXml('1')},definedNames:powerName,validation});
 const b=await fixture('B',{Equipment:[row('2','Equipment=1','2')]},{helpers:{HideEnum:helperXml('2')},definedNames:powerName.replace('$A$1','$B$1'),validation});
 const s=await inspect([a,b]);assert.equal(s.issues.length,0);assert.match(s.warnings[0],/HideEnum/);assert.match(s.warnings[0],/enum codes are not translated/);
 const z=await Zip.loadAsync(await E.merge(s)),base=await Zip.loadAsync(a.bytes);
 assert.equal(await z.file('xl/worksheets/sheet2.xml').async('string'),await base.file('xl/worksheets/sheet2.xml').async('string'));
 assert.equal(await z.file('xl/workbook.xml').async('string'),await base.file('xl/workbook.xml').async('string'));
 const raw=await z.file('xl/worksheets/sheet1.xml').async('string');assert.match(raw,/<formula1>PowerOptions<\/formula1>/);assert.match(raw,/sqref="H6:H7"/);assert.match(raw,/<c r="H7"[^>]*><is><t[^>]*>2<\/t>/);
});
test('lookup comparison resolves shared strings and ignores XML/style-only differences',async()=>{
 const a=await fixture('A',{Equipment:[]},{helpers:{HideEnum:helperXml('shared value')}});
 const shared='<worksheet xmlns="'+NS+'"><sheetData><row r="1"><c r="A1" s="0" t="s"><v>0</v></c></row></sheetData></worksheet>';
 const b=await fixture('B',{Equipment:[]},{shared:true,helpers:{HideEnum:shared}});
 let s=await inspect([a,b]);assert.equal(s.issues.length,0);assert.equal(s.warnings.length,0);
 const c=await fixture('C',{Equipment:[]},{shared:true,strings:['different value'],helpers:{HideEnum:shared}});
 s=await inspect([b,c]);assert.equal(s.issues.length,0);assert.equal(s.warnings.length,1);
});
test('missing later-file or File-1 lookup sheets do not block shared-sheet append',async()=>{
 const a=await fixture('A',{Equipment:[row('1','Equipment=1','1')]}),b=await fixture('B',{Equipment:[row('2','Equipment=1','2')]},{helpers:{HideEnum:helperXml('2')}});
 for(const inputs of [[a,b],[b,a]]) {
   const s=await inspect(inputs);assert.equal(s.issues.length,0);assert.equal(s.warnings.length,1);
   const z=await Zip.loadAsync(await E.merge(s));assert.equal((await z.file('xl/workbook.xml').async('string')).includes('HideEnum'),inputs[0]===b);
 }
});
test('additional reference sheets import their own lookups and names without changing File-1',async()=>{
 const a=await fixture('A',{Equipment:[]},{helpers:{HideEnum:helperXml('1')},definedNames:powerName});
 const b=await fixture('B',{Extra:[row('2','Extra=1','2')]},{helpers:{HideEnum:helperXml('2')},validation,definedNames:powerName});
 let s=await inspect([a,b]);assert.equal(s.issues.length,0);
 let z=await Zip.loadAsync(await E.merge(s));assert.match(await z.file('xl/worksheets/merged1.xml').async('string'),/<formula1>ITBBU_F2_PowerOptions<\/formula1>/);
 assert.match(await z.file('xl/workbook.xml').async('string'),/name="__ITBBU_F2_HideEnum"[^>]*state="hidden"/);
 assert.match(await z.file('xl/workbook.xml').async('string'),/name="ITBBU_F2_PowerOptions">&apos;__ITBBU_F2_HideEnum&apos;!\$A\$1/);
 const original=await Zip.loadAsync(a.bytes);assert.equal(await z.file('xl/worksheets/sheet2.xml').async('string'),await original.file('xl/worksheets/sheet2.xml').async('string'));
 const c=await fixture('C',{Extra:[row('2','Extra=1','2')]},{helpers:{HideEnum:helperXml('2')}});
 s=await inspect([a,c]);assert.equal(s.issues.length,0);assert.ok((await Zip.loadAsync(await E.merge(s))).file('xl/worksheets/merged1.xml'));
 const d=await fixture('D',{Extra:[row('2','Extra=1','1')]},{helpers:{HideEnum:helperXml('1').replace('showGridLines="0"','showGridLines="1"')},validation,definedNames:powerName});
 s=await inspect([a,d]);assert.equal(s.issues.length,0);assert.ok((await Zip.loadAsync(await E.merge(s))).file('xl/worksheets/merged1.xml'));
 const e=await fixture('E',{Extra:[row('2','Extra=1','1')]},{helpers:{HideEnum:helperXml('1')},validation,definedNames:powerName.replace('$A$1','$A$2')});
 s=await inspect([a,e]);assert.equal(s.issues.length,0);z=await Zip.loadAsync(await E.merge(s));
 assert.match(await z.file('xl/workbook.xml').async('string'),/name="ITBBU_F2_PowerOptions">HideEnum!\$A\$2/);
});
test('audit export includes all source labels, literal formulas, complete rows and safe typed strings',async()=>{
 const s=await inspect([await fixture('A',{Equipment:[row('1','Equipment=1','=1+1')]}),await fixture('B',{Equipment:[row('2','Equipment=1','=2+2')]})]);
 E.summarizeAudit(s);const z=await Zip.loadAsync(await E.exportAudit(s));assert.ok(z.file('xl/sharedStrings.xml'));const strings=await z.file('xl/sharedStrings.xml').async('string');assert.match(strings,/File-1 value/);assert.match(strings,/File-2 value/);assert.match(strings,/=1\+1/);
 const body=await z.file('xl/worksheets/sheet4.xml').async('string');assert.equal((body.match(/<row /g)||[]).length,2);assert.equal(/<f>/.test(body),false);assert.match(body,/<autoFilter ref="A1:K2"/);assert.match(strings,/File-1 original LDN/);assert.match(strings,/File-2 original LDN/);
});
test('pagination and selected-sheet audit remain scoped',async()=>{
 const s=await inspect([await fixture('A',{Equipment:[row('1','Equipment=1','1')],Other:[row('1','Other=1','1')]}),await fixture('B',{Equipment:[row('2','Equipment=1','1')],Other:[row('2','Other=1','2')]})]);
 E.summarizeAudit(s,{selected:['Other']});const p=E.auditPage(s,{status:'issues',limit:1});assert.equal(p.total,1);assert.equal(p.rows[0].sheet,'Other');assert.equal(E.auditPage(s,{status:'Same'}).total,0);
});

const wideNames = [...names, ...Array.from({length:27}, (_, i) => 'parameter' + (i + 9))];
const wideTemplate = template.map((r, i) => [...r, ...wideNames.slice(8).map(name => i === 0 ? name : i === 1 ? 'Description ' + name : i === 2 ? 'string' : i === 3 ? 'Guidance' : '--')]);
const wideRow = (site, ldn, power) => [...row(site, ldn, power), ...wideNames.slice(8).map((_, i) => site + '-' + (i + 9))];
const allColumns = wideNames.map((_, i) => i);
const dataRows = raw => [...raw.matchAll(/<row\b[^>]*>[\s\S]*?<\/row>/g)].filter(m => Number(E._internals.attrs(m[0].match(/^<row[^>]*>/)[0]).r) >= 6).map(m => E._internals.parseCells(m[0], []));
const headerRows = raw => [...raw.matchAll(/<row\b[^>]*>[\s\S]*?<\/row>/g)].filter(m => Number(E._internals.attrs(m[0].match(/^<row[^>]*>/)[0]).r) < 6).map(m => m[0]);

test('23, 35 and 30 columns use the widest File-2 reference and preserve every value in upload order',async()=>{
 const orders = [allColumns.slice(0,23), [0,...allColumns.slice(1).reverse()], [0,...allColumns.slice(12).reverse(),...allColumns.slice(1,7)]];
 const inputs = [[wideRow('100','Object=1','01'),wideRow('100','Object=2',0)], [wideRow('200','Object=1',0)], [wideRow('300','Object=1','')]];
 const files = await Promise.all(inputs.map((rows, i) => fixture('site'+i, {NRDcRelPSCellA2:rows}, {template:wideTemplate, order:orders[i]})));
 const s=await inspect(files);assert.equal(s.issues.length,0);const c=s.sheets[0];
 assert.deepEqual(c.columnCounts,[23,35,30]);assert.equal(c.reference,'File-2');assert.equal(c.referenceIndex,1);assert.equal(c.columns,35);
 const z=await Zip.loadAsync(await E.merge(s)), reference=await Zip.loadAsync(files[1].bytes);
 const raw=await z.file('xl/worksheets/sheet1.xml').async('string');
 assert.equal((await z.file('xl/workbook.xml').async('string')).match(/name="NRDcRelPSCellA2"/g).length,1);
 assert.deepEqual(headerRows(raw),headerRows(await reference.file('xl/worksheets/sheet1.xml').async('string')));
 const merged=dataRows(raw);let targetRow=0;
 inputs.forEach((records, f) => records.forEach(source => {
   orders[1].forEach((sourceColumn, destinationColumn) => {
     const cell=merged[targetRow][destinationColumn];
     assert.equal(cell.text,orders[f].includes(sourceColumn) ? String(source[sourceColumn]) : '', `file ${f}, row ${targetRow}, ${wideNames[sourceColumn]}`);
     if (orders[f].includes(sourceColumn) && typeof source[sourceColumn] === 'number') assert.equal(cell.kind,'n');
   });targetRow++;
 }));
 assert.equal(targetRow,4);assert.match(raw,/<dimension ref="A1:AI9"/);
 const audit=[...E.auditRows(s,{includeIdentity:true})];
 const missing=audit.find(r=>r.parameter==='parameter35' && r.ldns[0]==='Object=1');
 assert.equal(missing.status,'Missing parameter');assert.deepEqual(missing.values,['(missing parameter)','200-35','300-35']);
 assert.equal(missing.ldns[0],'Object=1');assert.equal(missing.ldns[1],'Object=1');
 const report=await Zip.loadAsync(await E.exportAudit(s));
 const coverage=await report.file('xl/worksheets/sheet3.xml').async('string');assert.match(coverage,/Reference template/);assert.match(coverage,/File-2/);assert.match(coverage,/>35<\/v>/);
 assert.ok(s.templateDetails.some(d=>d.parameter==='parameter35' && d.file==='File-1' && /blank/.test(d.action)));
});

test('each MO selects its own widest file, including File-3, and empty wide templates are eligible',async()=>{
 const moNames=['First','Second','Third','EmptyReference'];
 const files=[];
 for (let f=0;f<3;f++) {
   const sheets={}, order={};
   for (let m=0;m<moNames.length;m++) {
     const name=moNames[m], isReference=f===(m%3);
     sheets[name]=name==='EmptyReference' && isReference ? [] : [wideRow(String(f),name+'=1',String(m))];
     order[name]=isReference ? allColumns : allColumns.slice(0,23);
   }
   files.push(await fixture('file'+f,sheets,{template:wideTemplate,order}));
 }
 const s=await inspect(files);assert.equal(s.issues.length,0);assert.deepEqual(s.sheets.map(s=>s.reference),['File-1','File-2','File-3','File-1']);
 const z=await Zip.loadAsync(await E.merge(s));
 for(let m=0;m<moNames.length;m++) {
   const raw=await z.file(`xl/worksheets/sheet${m+1}.xml`).async('string');
   assert.equal(E._internals.parseCells(headerRows(raw)[0],[]).length,35);
   assert.equal(dataRows(raw).length,m===3 ? 2 : 3);
 }
});

test('ties keep the earliest template and different non-subset parameter sets cannot lose values',async()=>{
 const a=await fixture('A',{MO:[row('1','Object=1','base')]});
 const b=await fixture('B',{MO:[row('2','Object=1','later')]},{order:[0,7,6,5,4,3,2,1],headerMismatch:true});
 let s=await inspect([a,b]);assert.equal(s.sheets[0].reference,'File-1');assert.equal(s.issues.length,0);
 s=await inspect([a,await fixture('C',{MO:[row('2','Object=1','must keep')]},{headerCells:{1:{7:'otherParameter'}}})]);
 assert.ok(s.issues.some(i=>/otherParameter/.test(i.message) && /widest reference/.test(i.message)));await assert.rejects(()=>E.merge(s),/compatibility/);
});

test('missing parameter, present blank, zero and duplicate LDNs remain distinct in the audit',async()=>{
 const narrow=await fixture('A',{MO:[row('1','Object=1','unused')]},{order:[0,1,2,3,4,5,6]});
 const blank=await fixture('B',{MO:[row('2','Object=1','')]});
 const zero=await fixture('C',{MO:[row('3','Object=1',0)]});
 let s=await inspect([narrow,blank,zero]);assert.equal(s.issues.length,0);
 const audit=[...E.auditRows(s)];assert.equal(audit.length,1);assert.equal(audit[0].status,'Missing parameter');assert.deepEqual(audit[0].values,['(missing parameter)','(blank)','0 [number]']);
 const duplicate=await fixture('D',{MO:[row('1','Object=1','unused'),row('1','Object=1','unused')]},{order:[0,1,2,3,4,5,6]});
 s=await inspect([duplicate,blank]);assert.equal([...E.auditRows(s)][0].status,'Ambiguous key');assert.equal([...E.auditRows(s)][0].values[0],'(missing parameter)');
});

test('a later wider reference retains dropdown dependencies, scoped names and File-1 support sheets',async()=>{
 const a=await fixture('A',{MO:[row('1','Object=1','')]},{order:[0,1,2,3,4,5,6],helpers:{HideEnum:helperXml('base choice')},definedNames:powerName+'<definedName name="_xlnm.Print_Area" localSheetId="0">MO!$A$1:$G$6</definedName>'});
 const b=await fixture('B',{MO:[row('2','Object=1','reference choice')]},{helpers:{HideEnum:helperXml('reference choice')},definedNames:powerName+'<definedName name="_xlnm.Print_Area" localSheetId="0">MO!$A$1:$H$6</definedName>',validation,shared:true});
 const s=await inspect([a,b]);assert.equal(s.issues.length,0);assert.equal(s.sheets[0].reference,'File-2');
 const z=await Zip.loadAsync(await E.merge(s)), original=await Zip.loadAsync(a.bytes);
 const raw=await z.file('xl/worksheets/sheet1.xml').async('string'),book=await z.file('xl/workbook.xml').async('string');
 assert.match(raw,/<formula1>ITBBU_F2_PowerOptions<\/formula1>/);assert.match(raw,/sqref="H6:H7"/);
 assert.equal((book.match(/name="MO"/g)||[]).length,1);assert.equal((book.match(/name="_xlnm.Print_Area"/g)||[]).length,1);
 assert.match(book,/name="_xlnm.Print_Area" localSheetId="0">MO!\$A\$1:\$H\$6/);
 assert.match(book,/name="ITBBU_F2_PowerOptions">&apos;__ITBBU_F2_HideEnum&apos;!\$A\$1/);
 assert.match(book,/name="__ITBBU_F2_HideEnum"[^>]*state="hidden"/);
 assert.equal(await z.file('xl/worksheets/sheet2.xml').async('string'),await original.file('xl/worksheets/sheet2.xml').async('string'));
 assert.match(await z.file('xl/worksheets/merged1.xml').async('string'),/reference choice/);
 assert.deepEqual(dataRows(raw).map(c=>c[7].text),['','reference choice']);
});

const importedStyles='<styleSheet xmlns="'+NS+'"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.0000"/></numFmts><fonts count="2"><font><name val="Courier New"/><sz val="12"/></font><font><name val="Calibri"/><b/><color rgb="FF123456"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFB788E0"/></patternFill></fill></fills><borders count="1"><border><bottom style="thin"><color rgb="FF445566"/></bottom></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>';
test('later reference fonts, fills, borders, formats and conditional styles are imported with valid indices',async()=>{
 const a=await fixture('A',{MO:[row('1','Object=1','')]},{order:[0,1,2,3,4,5,6]});
 const b=await fixture('B',{MO:[row('2','Object=1',1.25)]},{styles:importedStyles,mutateXml:raw=>raw.replace('<c r="H6" s="0"','<c r="H6" s="2"').replace('max="8" width','max="8" style="2" width').replace('</worksheet>','<conditionalFormatting sqref="H6:H6"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting></worksheet>')});
 const s=await inspect([a,b]);assert.equal(s.issues.length,0);
 const z=await Zip.loadAsync(await E.merge(s)), raw=await z.file('xl/worksheets/sheet1.xml').async('string'), styles=await z.file('xl/styles.xml').async('string');
 const items=(collection,tag)=>[...styles.match(new RegExp('<'+collection+'\\b[^>]*>[\\s\\S]*?</'+collection+'>'))[0].matchAll(new RegExp('<'+tag+'\\b[^>]*?(?:/>|>[\\s\\S]*?</'+tag+'>)','g'))].map(m=>m[0]);
 const xfs=items('cellXfs','xf'), fonts=items('fonts','font'), fills=items('fills','fill'), borders=items('borders','border');
 const styleAt=ref=>E._internals.attrs(raw.match(new RegExp('<c r="'+ref+'"[^>]*>'))[0]).s;
 const header=E._internals.attrs(xfs[styleAt('H1')]), missing=E._internals.attrs(xfs[styleAt('H6')]), value=E._internals.attrs(xfs[styleAt('H7')]);
 assert.match(fonts[header.fontId],/Calibri/);assert.match(fills[header.fillId],/FFB788E0/);assert.match(borders[header.borderId],/thin/);
 assert.equal(missing.numFmtId,value.numFmtId);assert.match(fonts[value.fontId],/Courier New/);assert.match(styles,/formatCode="0.0000"/);
 assert.match(raw,/conditionalFormatting sqref="H6:H7"/);assert.match(raw,/dxfId="0"/);
 assert.deepEqual(dataRows(raw).map(c=>c[7].text),['','1.25']);
});

test('self-closing blank cells in source files never consume the following parameter value',async()=>{
 const a=await fixture('A',{MO:[row('1','Object=1','01')]},{mutateXml:raw=>raw.replace(/<c r="A6"[^>]*>[\s\S]*?<\/c>/,'<c r="A6" s="0"/>')});
 const b=await fixture('B',{MO:[row('2','Object=1',0)]},{order:[0,7,6,5,4,3,2,1]});
 const s=await inspect([a,b]);assert.equal(s.issues.length,0);assert.equal(s.books[0].sheets.get('MO').rows[0].cells[1].text,'ITBBU');
 const z=await Zip.loadAsync(await E.merge(s)), rows=dataRows(await z.file('xl/worksheets/sheet1.xml').async('string'));
 assert.deepEqual(rows.map(r=>r.map(c=>c.text)),[row('1','Object=1','01'),row('2','Object=1','0')]);
});

test('references with dynamic or external dropdown dependencies stop before a broken workbook is produced',async()=>{
 const a=await fixture('A',{MO:[row('1','Object=1','')]},{order:[0,1,2,3,4,5,6]});
 for(const opts of [
   {validation:validation.replace('PowerOptions','INDIRECT(&quot;HideEnum!A1&quot;)')},
   {validation,definedNames:powerName.replace('HideEnum!','[lookup.xlsx]HideEnum!')}
 ]) {
   const b=await fixture('B',{MO:[row('2','Object=1','1')]},opts), s=await inspect([a,b]);
   assert.ok(s.issues.some(i=>/dynamic or external/.test(i.message)));await assert.rejects(()=>E.merge(s),/compatibility/);
 }
});
