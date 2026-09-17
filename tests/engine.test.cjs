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
 const z = new Zip(), keys = Object.keys(sheets);
 z.file('[Content_Types].xml',`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`);
 z.file('xl/styles.xml','<styleSheet xmlns="'+NS+'"><cellXfs count="1"><xf/></cellXfs></styleSheet>');
 z.file('xl/workbook.xml',`<workbook xmlns="${NS}" xmlns:r="${R}"><sheets>${keys.map((n,i)=>`<sheet name="${n}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets></workbook>`);
 z.file('xl/_rels/workbook.xml.rels',`<Relationships>${keys.map((n,i)=>`<Relationship Id="rId${i+1}" Target="worksheets/sheet${i+1}.xml" Type="${R}/worksheet"/>`).join('')}${opts.shared ? `<Relationship Id="rIdS" Target="sharedStrings.xml" Type="${R}/sharedStrings"/>` : ''}</Relationships>`);
 keys.forEach((key,i)=>{
   let xml = E._internals.sheetXml([...template.map(r=>[...r]),...sheets[key]],names.map(()=>20),false);
   if(opts.headerMismatch && i===0) xml=xml.replace('Exact parameter','Changed metadata');
   if(opts.shared) xml=xml.replace('<c r="H6" s="0" t="inlineStr"><is><t xml:space="preserve">shared value</t></is></c>','<c r="H6" s="0" t="s"><v>0</v></c>');
   if(opts.formula) xml=xml.replace('<c r="H6" s="0" t="inlineStr"><is><t xml:space="preserve">formula</t></is></c>','<c r="H6" s="0"><f>1+1</f><v>2</v></c>');
   z.file(`xl/worksheets/sheet${i+1}.xml`,xml);
 });
 if(opts.shared) z.file('xl/sharedStrings.xml',`<sst xmlns="${NS}"><si><t>shared value</t></si></sst>`);
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
test('incompatible metadata and formula cells block merge, while audit stays available',async()=>{
 const a=await fixture('A',{Equipment:[row('1','Equipment=1','1')]}),b=await fixture('B',{Equipment:[row('2','Equipment=1','2')]},{headerMismatch:true});
 let s=await inspect([a,b]);assert.equal(s.issues.length,1);assert.equal([...E.auditRows(s)].length,1);await assert.rejects(()=>E.merge(s),/compatibility/);
 s=await inspect([a,await fixture('F',{Equipment:[row('2','Equipment=1','formula')]},{formula:true})]);assert.match(s.issues[0].message,/Formula/);
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
