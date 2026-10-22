import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import JSZip from 'jszip';
import {load} from 'cheerio/slim';
import {analyzeArticle,articleDefaults} from './article.js';
import {createReportExcel} from './report-excel.js';
import {reviseWord} from './revise-word.js';
const original=readFileSync('templates/article-format.docx');
const makeReport=r=>({id:7,kind:'article',name:'文章測試.docx',period:'2026-10',created:'2026-10-01T01:00:00Z',total:r.summary.paragraphs,issues:r.issues,rows:r.paragraphs,rules:articleDefaults,metadata:{summary:r.summary,policy:r.policy,checks:r.checks}});

test('真正 XLSX 包含完整錯誤、逐項結果與規則快照；公式字串以純文字存放',async()=>{
 const result=await analyzeArticle(original);const report=makeReport(result);report.name='=HYPERLINK("https://example.invalid")';
 const bytes=await createReportExcel(report);const zip=await JSZip.loadAsync(bytes,{checkCRC32:true});
 const workbook=load(await zip.file('xl/workbook.xml').async('string'),{xml:true});
 assert.deepEqual(workbook('sheet').toArray().map(n=>workbook(n).attr('name')),['檢核摘要','錯誤與修改建議','全部檢核項目','當次規則快照','文章文字摘錄']);
 const summary=load(await zip.file('xl/worksheets/sheet1.xml').async('string'),{xml:true});const nameCell=summary('row').eq(1).find('c').eq(1);assert.equal(nameCell.attr('t'),'inlineStr');assert.equal(nameCell.find('t').text(),report.name);assert.equal(summary('f').length,0);
 const issues=load(await zip.file('xl/worksheets/sheet2.xml').async('string'),{xml:true});assert.equal(issues('row').length,result.issues.length+1);assert.equal(issues('row').last().find('c').last().find('t').text(),result.issues.at(-1).suggestion);
 const checks=load(await zip.file('xl/worksheets/sheet3.xml').async('string'),{xml:true});assert.equal(checks('row').length,result.checks.length+1);assert.equal(checks('pane').attr('state'),'frozen');assert.equal(checks('autoFilter').length,1);
});

test('修訂 Word 修正可確定格式，保留原文與全部非本文封裝檔案',async()=>{
 const beforeBytes=Buffer.from(original);const r=await reviseWord(original);assert.ok(r.changes.length>0);assert.ok(r.after.summary.failed<r.before.summary.failed);assert.deepEqual(r.before.paragraphs,r.after.paragraphs);assert.deepEqual(original,beforeBytes);
 assert.ok(!r.after.checks.some(c=>['字體','字級','左右對齊','行距','首行縮排'].includes(c.field)&&c.status==='failed'));
 assert.ok(r.after.issues.some(i=>i.status==='manual'));assert.ok(r.after.issues.some(i=>i.field==='圖片有效解析度'));
 const a=await JSZip.loadAsync(original);const b=await JSZip.loadAsync(r.bytes,{checkCRC32:true});assert.deepEqual(Object.keys(a.files).sort(),Object.keys(b.files).sort());
 for(const name of Object.keys(a.files).filter(n=>!a.files[n].dir&&n!=='word/document.xml'))assert.deepEqual(await a.file(name).async('uint8array'),await b.file(name).async('uint8array'),name);
 const second=await reviseWord(r.bytes);assert.equal(second.changes.length,0);
});

test('停用規則不自動修訂；全停用時直接保留原始位元組',async()=>{
 const r=await reviseWord(original,Object.fromEntries(Object.keys(articleDefaults).map(k=>[k,false])));assert.equal(r.changes.length,0);assert.deepEqual(Buffer.from(r.bytes),original);
 const partial=await reviseWord(original,{...articleDefaults,typography:false});assert.ok(!partial.changes.some(c=>['字體','字級'].includes(c.field)));
});

test('Excel 修訂紀錄及剩餘問題與 Word 再檢核結果一致',async()=>{
 const revision=await reviseWord(original);const bytes=await createReportExcel(makeReport(revision.before),revision);const zip=await JSZip.loadAsync(bytes);
 const changes=load(await zip.file('xl/worksheets/sheet6.xml').async('string'),{xml:true});assert.equal(changes('row').length,revision.changes.length+1);
 const remaining=load(await zip.file('xl/worksheets/sheet7.xml').async('string'),{xml:true});assert.equal(remaining('row').length,revision.after.issues.length+1);assert.ok(remaining('row').toArray().some(n=>remaining(n).text().includes('圖片有效解析度')));
});
