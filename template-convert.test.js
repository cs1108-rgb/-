import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import {load} from 'cheerio/slim';
import {applyTemplateWord} from './template-word.js';
import {articleDefaults} from './article.js';
const source=readFileSync('templates/article-format.docx');
const template=readFileSync('templates/article-format-a4.docx');
const xml=s=>load(s,{xml:true});
const local=n=>n.name?.split(':').pop();
async function textList(bytes){const z=await JSZip.loadAsync(bytes);const $=xml(await z.file('word/document.xml').async('string'));return $('*').filter((_,n)=>local(n)==='t').toArray().map(n=>$(n).text());}

test('不同版型 Word 在檢核全停用時仍統一 A4 與附件樣式，原文和媒體保留',async()=>{
 const z=await JSZip.loadAsync(source);const $=xml(await z.file('word/document.xml').async('string'));
 $('w\\:pgSz').attr('w:w','18000').attr('w:h','10000').attr('w:orient','landscape');
 $('w\\:pgMar').attr('w:top','300').attr('w:left','200');$('w\\:rFonts').attr('w:eastAsia','Arial');$('w\\:sz').attr('w:val','80');
 z.file('word/document.xml',$.xml(),{createFolders:false});const input=await z.generateAsync({type:'uint8array'});
 const r=await applyTemplateWord(input,template,Object.fromEntries(Object.keys(articleDefaults).map(k=>[k,false])));assert.equal(r.mode,'template');assert.ok(r.changes.length>20);assert.equal(r.template.page,'A4');assert.deepEqual(await textList(input),await textList(r.bytes));
 const result=await JSZip.loadAsync(r.bytes);const doc=xml(await result.file('word/document.xml').async('string'));
 for(const pg of doc('w\\:pgSz').toArray()){assert.equal(doc(pg).attr('w:w'),'11906');assert.equal(doc(pg).attr('w:h'),'16838');assert.equal(doc(pg).attr('w:orient'),undefined);}
 const expected=xml(await(await JSZip.loadAsync(template)).file('word/document.xml').async('string'));
 for(const k of['top','bottom','left','right','header','footer'])assert.equal(doc('w\\:pgMar').last().attr('w:'+k),expected('w\\:pgMar').last().attr('w:'+k));
 for(const name of Object.keys(z.files).filter(n=>n.startsWith('word/media/')))assert.deepEqual(await z.file(name).async('uint8array'),await result.file(name).async('uint8array'));
 assert.ok(doc('w\\:pStyle').toArray().some(n=>doc(n).attr('w:val').endsWith('Role_Title')));assert.ok(r.warnings.length>0);
});

test('套版後全部封裝關聯可解析，頁首頁尾採範本；重複套版不累積樣式或資源',async()=>{
 const a=await applyTemplateWord(source,template);const zip=await JSZip.loadAsync(a.bytes,{checkCRC32:true});
 for(const name of Object.keys(zip.files).filter(n=>n.endsWith('.rels'))){const rel=xml(await zip.file(name).async('string'));const owner=name==='_rels/.rels'?'':name.replace('/_rels/','/').replace(/\.rels$/,'');for(const r of rel('Relationship').toArray()){if(rel(r).attr('TargetMode')==='External')continue;const target=rel(r).attr('Target');const resolved=target.startsWith('/')?target.slice(1):path.posix.normalize(path.posix.join(path.posix.dirname(owner),target));assert.ok(zip.file(resolved),`${name} -> ${resolved}`);}}
 const document=xml(await zip.file('word/document.xml').async('string'));const rel=xml(await zip.file('word/_rels/document.xml.rels').async('string'));const headerId=document('w\\:headerReference').last().attr('r:id');assert.ok(rel('Relationship').toArray().some(n=>rel(n).attr('Id')===headerId&&rel(n).attr('Target').includes('mios-template-')));
 const b=await applyTemplateWord(a.bytes,template);const twice=await JSZip.loadAsync(b.bytes);assert.equal(Object.keys(twice.files).length,Object.keys(zip.files).length);
 const styles=xml(await twice.file('word/styles.xml').async('string'));const ids=styles('w\\:style').toArray().map(n=>styles(n).attr('w:styleId'));assert.equal(new Set(ids).size,ids.length);
 assert.deepEqual(await textList(a.bytes),await textList(b.bytes));
});

test('沒有的章節只提示，不把範例作者、內文或圖片搬入文章',async()=>{
 const zip=await JSZip.loadAsync(source);const $=xml(await zip.file('word/document.xml').async('string'));$('w\\:body').empty().append('<w:p><w:r><w:t>新型感測技術</w:t></w:r></w:p><w:p><w:r><w:t>陳大明</w:t></w:r></w:p><w:p><w:r><w:t>本文探討感測技術的發展與效益。</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="16000" w:h="9000"/></w:sectPr>');zip.file('word/document.xml',$.xml(),{createFolders:false});
 const input=await zip.generateAsync({type:'uint8array'});const result=await applyTemplateWord(input,template);assert.deepEqual(await textList(result.bytes),['新型感測技術','陳大明','本文探討感測技術的發展與效益。']);assert.ok(result.warnings.some(w=>w.includes('結論')));assert.ok(result.warnings.some(w=>w.includes('參考文獻')));assert.ok(result.warnings.some(w=>w.includes('作者簡介')));
});
