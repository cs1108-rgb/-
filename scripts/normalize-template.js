import JSZip from 'jszip';
import {load} from 'cheerio/slim';
import {readFile,writeFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
const zip=await JSZip.loadAsync(await readFile('templates/article-format.docx'));
const part=zip.file('word/document.xml');const $=load(await part.async('string'),{xml:true});
$('*').filter((_,n)=>n.name?.split(':').pop()==='pgSz').each((_,n)=>{
 for(const k of Object.keys(n.attribs||{}))if(['w','h','orient','code'].includes(k.split(':').pop()))$(n).removeAttr(k);
 $(n).attr('w:w','11906').attr('w:h','16838');
});
zip.file('word/document.xml',$.xml(),{createFolders:false,date:part.date});
const bytes=await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});
await writeFile('templates/article-format-a4.docx',bytes);
console.log('A4 output template SHA-256:',createHash('sha256').update(bytes).digest('hex'));
