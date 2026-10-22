import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import JSZip from 'jszip';
import {analyzeArticle,articleDefaults} from './article.js';
// Small OOXML fixtures make inheritance, split runs and exact units testable.
const run=(text,pr='')=>`<w:r><w:rPr>${pr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para=(text,pr='',rp='')=>`<w:p><w:pPr>${pr}</w:pPr>${run(text,rp)}</w:p>`;
const style='<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="微軟正黑體" w:eastAsia="微軟正黑體"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:jc w:val="both"/><w:spacing w:line="288" w:lineRule="auto"/><w:ind w:firstLineChars="200"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Body" w:default="1"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="H1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="H3"><w:name w:val="heading 3"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style></w:styles>';
const title=para('精密量測技術展望','','<w:sz w:val="64"/>');
const start=title+para('王小明、陳小華')+para('摘要')+para('本研究以實驗分析技術差異。');
const end=para('結論')+para('研究顯示方法具有穩定性。')+para('參考文獻')+para('王小明，精密量測研究，2026。')+para('作者簡介')+para('王小明／研究所／研究員');
async function document(body,extra={}){const z=new JSZip();z.file('word/document.xml',`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${body}</w:body></w:document>`);z.file('word/styles.xml',extra.styles||style);for(const[k,v]of Object.entries(extra.files||{}))z.file(k,v);return z.generateAsync({type:'nodebuffer'});}
const has=(r,field,status)=>r.checks.some(c=>c.field===field&&c.status===status);
test('繼承字體與段落格式：有效文章自動項目通過，文獻仍需人工確認',async()=>{const r=await analyzeArticle(await document(start+para('研究方法','<w:pStyle w:val="H1"/>')+para('實驗重複進行以確認結果。')+end));assert.equal(r.summary.failed,0);assert.equal(r.summary.score,100);assert.ok(has(r,'文獻正確性','manual'));assert.equal(r.summary.complete,false);assert.ok(has(r,'行距','passed'))});
test('檢出局部錯誤字體／字級／段落並指出段落與建議',async()=>{const bad=para('此段格式有誤','<w:jc w:val="left"/><w:spacing w:line="240"/><w:ind w:firstLineChars="0"/>','<w:rFonts w:eastAsia="新細明體"/><w:sz w:val="20"/>');const r=await analyzeArticle(await document(start+bad+end));for(const field of ['字體','字級','左右對齊','行距','首行縮排'])assert.ok(has(r,field,'failed'),field);const issues=r.issues.filter(i=>i.paragraph===5);assert.ok(issues.every(i=>i.suggestion&&i.excerpt==='此段格式有誤'))});
test('標題跳級與 Heading 3 下的首行不縮排',async()=>{const r=await analyzeArticle(await document(start+para('研究方法','<w:pStyle w:val="H1"/>')+para('條列方法','<w:pStyle w:val="H3"/>')+para('條列下內文','<w:ind w:firstLineChars="0"/>')+end));assert.ok(has(r,'標題層級','warning'));assert.ok(r.checks.some(c=>c.paragraph===7&&c.field==='首行縮排'&&c.status==='passed'))});
test('拆分文字仍檢查圖表引用加粗及編號對應',async()=>{const content='<w:p>'+run('見 ')+run('圖','<w:b/>')+run('一')+run(' 的比較結果。')+'</w:p>';const r=await analyzeArticle(await document(start+content+para('圖一 實驗結果')+end));assert.ok(has(r,'引用加粗','failed'));assert.ok(has(r,'文中引用','passed'));assert.ok(has(r,'引用對應圖表','passed'))});
test('錯誤圖號不會以部分字串誤認引用',async()=>{const r=await analyzeArticle(await document(start+para('見圖10 的實驗。')+para('圖1 實驗結果')+end));assert.ok(has(r,'文中引用','warning'));assert.ok(has(r,'引用對應圖表','warning'))});
test('停用規則不執行，也不會產生虛假的百分之百',async()=>{const r=await analyzeArticle(await document(start+end),Object.fromEntries(Object.keys(articleDefaults).map(k=>[k,false])));assert.equal(r.checks.length,0);assert.equal(r.summary.score,null);assert.equal(r.summary.disabled.length,6);assert.equal(r.summary.complete,false)});
test('未填章節及範例文字不會被當成合格文章',async()=>{const r=await analyzeArticle(await document(para('標題(32pt)')+para('作者1')+para('這是摘要範例。')));assert.ok(r.issues.some(i=>i.field==='結論'));assert.ok(r.issues.some(i=>i.field==='範例文字'));assert.ok(r.summary.failed>0)});
test('實際上傳範本可解析，保留來源版本及雜湊，不假設範本本身全部合格',async()=>{const r=await analyzeArticle(readFileSync('templates/article-format.docx'));assert.ok(r.summary.paragraphs>20);assert.equal(r.policy.sha256.length,64);assert.ok(r.issues.some(i=>i.field==='範例文字'));assert.equal(r.issues.find(i=>i.field==='摘要').paragraph,3)});
test('圖片以像素及置入尺寸計算 DPI，圖說在下方',async()=>{const drawing='<w:p><w:r><w:drawing><wp:inline><wp:extent cx="3600000" cy="3600000"/><a:graphic><a:blip r:embed="rId1"/></a:graphic></wp:inline></w:drawing></w:r></w:p>';const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jV9kAAAAASUVORK5CYII=','base64');const r=await analyzeArticle(await document(start+drawing+para('圖一 測試')+end,{files:{'word/_rels/document.xml.rels':'<Relationships><Relationship Id="rId1" Target="media/test.png"/></Relationships>','word/media/test.png':png}}));assert.ok(has(r,'圖片有效解析度','failed'));assert.ok(has(r,'圖說位置','passed'));assert.ok(has(r,'線條圖解析度','manual'));assert.ok(has(r,'圖片寬度','passed'))});
test('非 Word 與無文字文件明確拒絕',async()=>{await assert.rejects(analyzeArticle(Buffer.from('invalid')),/有效且未加密/);await assert.rejects(analyzeArticle(await document('<w:p/>')),/可檢核的文章文字/)});

test('補充規範：作者頓號、圖說置中、編號層級與章節順序',async()=>{
 const bad=await analyzeArticle(await document(para('技術研究')+para('王小明，陳小華')+para('摘要內容。')+para('一、研究方法','<w:pStyle w:val="H3"/>')+para('圖一 實驗結果','<w:jc w:val="left"/>')+para('作者簡介')+para('參考文獻')+para('結論')));
 for(const field of ['作者分隔符號','圖說置中','章節編號格式','章節順序'])assert.ok(has(bad,field,'failed'),field);
 const good=await analyzeArticle(await document(para('技術研究')+para('王小明、陳小華')+para('摘要內容。')+para('圖一 實驗結果','<w:jc w:val="center"/>')+end));
 assert.ok(has(good,'作者分隔符號','passed'));assert.ok(has(good,'圖說置中','passed'));
});
