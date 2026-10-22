import JSZip from 'jszip';
import {load} from 'cheerio/slim';
import {analyzeArticle,articleDefaults,policy} from './article.js';
const WORD_NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const local=n=>n?.name?.split(':').pop();
const child=($,n,k)=>$(n).children().toArray().find(c=>local(c)===k);
const pOrder='pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'.split(' ');
const rOrder='rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange'.split(' ');
export async function reviseWord(buffer,rules=articleDefaults){
 // Always re-analyze the retained source so old report locations cannot edit the wrong paragraphs.
 const before=await analyzeArticle(buffer,rules);
 const zip=await JSZip.loadAsync(buffer);const $=load(await zip.file('word/document.xml').async('string'),{xml:true});
 const root=$('*').filter((_,n)=>local(n)==='document').get(0);
 let prefix=Object.entries(root.attribs||{}).find(([k,v])=>k.startsWith('xmlns:')&&v===WORD_NS)?.[0].slice(6);
 if(!prefix){prefix='moW';$(root).attr('xmlns:moW',WORD_NS);}
 const setAttr=(node,k,v)=>{const old=Object.keys(node.attribs||{}).find(a=>a.split(':').pop()===k);$(node).attr(old||`${prefix}:${k}`,String(v));};
 const removeAttr=(node,...keys)=>{for(const k of Object.keys(node.attribs||{}))if(keys.includes(k.split(':').pop()))$(node).removeAttr(k);};
 function ensure(parent,name,order){let n=child($,parent,name);if(n)return n;const markup=`<${prefix}:${name}/>`;const target=$(parent).children().toArray().find(c=>order.indexOf(local(c))>order.indexOf(name));if(target)$(target).before(markup);else $(parent).append(markup);return child($,parent,name);}
 function pr(parent,name){let n=child($,parent,name);if(!n){$(parent).prepend(`<${prefix}:${name}/>`);n=child($,parent,name);}return n;}
 function textRuns(p){return $(p).find('*').toArray().filter(n=>local(n)==='r'&&$(n).parents().toArray().find(a=>local(a)==='p')===p&&!$(n).parents().toArray().some(a=>['del','moveFrom'].includes(local(a)))&&$(n).find('*').toArray().some(t=>local(t)==='t'&&$(t).text().length));}
 const body=$('*').filter((_,n)=>local(n)==='body').get(0);
 const nodes=$(body).find('*').toArray().filter(n=>local(n)==='p'&&!$(n).parents().toArray().some(a=>['del','moveFrom'].includes(local(a))||local(a)==='Fallback'&&$(a).parent().children().toArray().some(c=>local(c)==='Choice')));
 const changes=[];
 for(const c of before.checks){
  if(!['failed','manual'].includes(c.status)||!c.paragraph||!['字體','字級','左右對齊','行距','首行縮排'].includes(c.field))continue;
  const p=nodes[c.paragraph-1];if(!p)throw Error('無法對應原稿段落，為避免錯誤修改，未產生 Word 副本。');
  const pPr=pr(p,'pPr');
  if(c.field==='字體'||c.field==='字級'){
   const runs=textRuns(p);if(!runs.length)continue;
   const rProperties=[...runs.map(r=>pr(r,'rPr')),ensure(pPr,'rPr',pOrder)];
   for(const rp of rProperties){if(c.field==='字體'){const fonts=ensure(rp,'rFonts',rOrder);removeAttr(fonts,'asciiTheme','hAnsiTheme','eastAsiaTheme','cstheme','csTheme');for(const k of ['ascii','hAnsi','eastAsia','cs'])setAttr(fonts,k,policy.font);}else{const pt=Number(String(c.expected).replace('pt',''));if(![12,32].includes(pt))continue;setAttr(ensure(rp,'sz',rOrder),'val',pt*2);setAttr(ensure(rp,'szCs',rOrder),'val',pt*2);}}
  }else if(c.field==='左右對齊')setAttr(ensure(pPr,'jc',pOrder),'val','both');
  else if(c.field==='行距'){const spacing=ensure(pPr,'spacing',pOrder);setAttr(spacing,'line',288);setAttr(spacing,'lineRule','auto');setAttr(ensure(pPr,'snapToGrid',pOrder),'val','0');}
  else if(c.field==='首行縮排'){const expected=String(c.expected).startsWith('2')?2:0;const ind=ensure(pPr,'ind',pOrder);removeAttr(ind,'hanging','hangingChars');setAttr(ind,'firstLineChars',expected*100);setAttr(ind,'firstLine',expected*240);}
  changes.push({paragraph:c.paragraph,location:c.location,excerpt:c.excerpt,field:c.field,before:c.actual??'未能解析',after:c.expected});
 }
 let bytes;
 if(changes.length){zip.file('word/document.xml',$.xml(),{createFolders:false});bytes=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});}else bytes=buffer instanceof ArrayBuffer?new Uint8Array(buffer):new Uint8Array(buffer);
 const after=await analyzeArticle(bytes,rules);
 return {bytes,changes,before,after,limitations:['僅修正當次啟用規則中可確定的字體、字級、對齊、行距及首行縮排。','未補寫文章、移動圖表、重編章節或更改引用文字；未知項目仍須人工確認。','保留原稿的文字、圖表、其他封裝檔案與未處理的格式；此為另存副本，非 Word 追蹤修訂模式。']};
}
