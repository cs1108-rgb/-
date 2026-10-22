import JSZip from 'jszip';
import {load} from 'cheerio/slim';
import {analyzeArticle,articleDefaults,policy} from './article.js';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR='http://schemas.openxmlformats.org/package/2006/relationships';
const local=n=>n?.name?.split(':').pop();
const elements=($,n,k)=>$(n).children().toArray().filter(c=>local(c)===k);
const first=($,n,k)=>elements($,n,k)[0];
const val=(n,k='val')=>Object.entries(n?.attribs||{}).find(([a])=>a.split(':').pop()===k)?.[1];
const all=($,n,k)=>$(n).find('*').toArray().filter(c=>local(c)===k);
const parse=s=>load(s,{xml:true});
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
function normalize(s){const out=[];for(const v of s.split('/')){if(v==='..')out.pop();else if(v&&v!=='.')out.push(v);}return out.join('/');}
const dir=s=>s.split('/').slice(0,-1).join('/');
const resolve=(from,to)=>normalize(to.startsWith('/')?to:dir(from)+'/'+to);
function relative(from,to){const a=dir(from).split('/').filter(Boolean),b=to.split('/');while(a.length&&a[0]===b[0]){a.shift();b.shift();}return [...a.map(()=> '..'),...b].join('/');}
const relPath=s=>`${dir(s)}/_rels/${s.split('/').pop()}.rels`;
function preferred($,p){return !$(p).parents().toArray().some(a=>['del','moveFrom'].includes(local(a))||local(a)==='Fallback'&&$(a).parent().children().toArray().some(c=>local(c)==='Choice'));}
function text($,p){return all($,p,'r').filter(r=>$(r).parents().toArray().find(a=>local(a)==='p')===p&&preferred($,r)).map(r=>all($,r,'t').map(t=>$(t).text()).join('')).join('');}
function remove($,n,k){for(const e of elements($,n,k))$(e).remove();}
function put($,n,k,v){const attr=Object.keys(n.attribs||{}).find(a=>a.split(':').pop()===k);$(n).attr(attr||`w:${k}`,String(v));}
function ns($,target,source){for(const[k,v]of Object.entries(source?.attribs||{})){if(!k.startsWith('xmlns:'))continue;if(target.attribs[k]&&target.attribs[k]!==v)throw Error('文件的 XML 命名空間與範本衝突，請先以 Word 另存為 .docx 後重試。');$(target).attr(k,v);}$(target).attr('xmlns:w',W);}
const names={Title:'文章標題',Author:'作者列',Abstract:'摘要',H1:'第一層標題',H2:'第二層標題',H3:'第三層標題',H4:'第四層標題',Body1:'一般內文',Body2:'小節內文',Body3:'條列內文',Body4:'次條列內文',Conclusion:'結論內文',Reference:'參考文獻',Bio:'作者簡介',Caption:'圖表標題',Cell:'表格文字'};

export async function applyTemplateWord(buffer,templateBuffer,enabled=articleDefaults){
 const before=await analyzeArticle(buffer,enabled);
 const zip=await JSZip.loadAsync(buffer);const tpl=await JSZip.loadAsync(templateBuffer);
 const $=parse(await zip.file('word/document.xml').async('string'));
 const t$=parse(await tpl.file('word/document.xml').async('string'));
 const root=$('*').filter((_,n)=>local(n)==='document').get(0),tRoot=t$('*').filter((_,n)=>local(n)==='document').get(0);
 const body=first($,root,'body'),tBody=first(t$,tRoot,'body');if(!body||!tBody)throw Error('無法取得原稿或範本本文。');
 ns($,root,tRoot);$(root).attr('xmlns:r',R);
 const tParas=elements(t$,tBody,'p');
 const base={Title:0,Author:1,Abstract:2,H1:4,Body1:5,H2:6,Body2:7,H3:8,Body3:9,H4:10,Body4:11,Conclusion:24,Reference:27,Bio:Math.max(0,tParas.findLastIndex(p=>/作者.*[/／].*[/／]/.test(text(t$,p)))),Caption:5,Cell:9};
 const prefix=`MIOS_${policy.sha256.slice(0,8)}_`;
 const tStyles=parse(await tpl.file('word/styles.xml').async('string'));
 const styles=parse(await zip.file('word/styles.xml')?.async('string')||`<w:styles xmlns:w="${W}"/>`);
 const sRoot=styles('*').filter((_,n)=>local(n)==='styles').get(0),tsRoot=tStyles('*').filter((_,n)=>local(n)==='styles').get(0);ns(styles,sRoot,tsRoot);ns(styles,sRoot,tRoot);
 const number=parse(await zip.file('word/numbering.xml')?.async('string')||`<w:numbering xmlns:w="${W}"/>`);
 const tNum=parse(await tpl.file('word/numbering.xml').async('string'));
 const nRoot=number('*').filter((_,n)=>local(n)==='numbering').get(0),tnRoot=tNum('*').filter((_,n)=>local(n)==='numbering').get(0);ns(number,nRoot,tnRoot);
 const numMap=new Map();let abstractNext=Math.max(0,...elements(number,nRoot,'abstractNum').map(n=>Number(val(n,'abstractNumId'))||0))+1;let numNext=Math.max(0,...elements(number,nRoot,'num').map(n=>Number(val(n,'numId'))||0))+1;
 for(const num of elements(tNum,tnRoot,'num')){
  const oldId=val(num,'numId'),marker=prefix+'Number'+oldId;
  let abstract=elements(number,nRoot,'abstractNum').find(a=>val(first(number,a,'name'))===marker);
  if(!abstract){const aid=val(first(tNum,num,'abstractNumId'));const original=elements(tNum,tnRoot,'abstractNum').find(a=>val(a,'abstractNumId')===aid);if(!original)continue;const clone=number(tNum.xml(original)).get(0);put(number,clone,'abstractNumId',abstractNext++);remove(number,clone,'name');const lvl=first(number,clone,'lvl');const markerXML=`<w:name w:val="${marker}"/>`;if(lvl)number(lvl).before(markerXML);else number(clone).append(markerXML);const firstNum=first(number,nRoot,'num');if(firstNum)number(firstNum).before(clone);else number(nRoot).append(clone);abstract=clone;}
  const abstractId=val(abstract,'abstractNumId');let target=elements(number,nRoot,'num').find(n=>val(first(number,n,'abstractNumId'))===abstractId);
  if(!target){target=number(tNum.xml(num)).get(0);put(number,target,'numId',numNext++);put(number,first(number,target,'abstractNumId'),'val',abstractId);const cleanup=first(number,nRoot,'numIdMacAtCleanup');if(cleanup)number(cleanup).before(target);else number(nRoot).append(target);}
  numMap.set(oldId,val(target,'numId'));
 }
 function remap($x,node){for(const n of [node,...$x(node).find('*').toArray()]){if(['pStyle','rStyle','tblStyle','basedOn','next','link','styleLink','numStyleLink'].includes(local(n))&&val(n))put($x,n,'val',prefix+'Imported_'+val(n));if(local(n)==='numId'&&numMap.has(val(n)))put($x,n,'val',numMap.get(val(n)));}return node;}
 for(const s of elements(tStyles,tsRoot,'style')){const id=prefix+'Imported_'+val(s,'styleId');for(const old of elements(styles,sRoot,'style').filter(n=>val(n,'styleId')===id))styles(old).remove();const clone=styles(tStyles.xml(s)).get(0);put(styles,clone,'styleId',id);for(const a of Object.keys(clone.attribs||{}))if(a.split(':').pop()==='default')styles(clone).removeAttr(a);remap(styles,clone);styles(sRoot).append(clone);}
 const rels=parse(await zip.file('word/_rels/document.xml.rels')?.async('string')||`<Relationships xmlns="${PR}"/>`);const relRoot=rels('*').filter((_,n)=>local(n)==='Relationships').get(0);
 const types=parse(await zip.file('[Content_Types].xml').async('string'));const typeRoot=types('*').filter((_,n)=>local(n)==='Types').get(0);const tTypes=parse(await tpl.file('[Content_Types].xml').async('string'));const ttRoot=tTypes('*').filter((_,n)=>local(n)==='Types').get(0);
 function addType(part,type){if(!elements(types,typeRoot,'Override').some(n=>val(n,'PartName')==='/'+part))types(typeRoot).append(`<Override PartName="/${esc(part)}" ContentType="${esc(type)}"/>`);}
 function mainRel(type,target){const existing=elements(rels,relRoot,'Relationship').find(n=>val(n,'Type')===type&&val(n,'Target')===target);if(existing)return val(existing,'Id');let i=1;while(elements(rels,relRoot,'Relationship').some(n=>val(n,'Id')===`rIdMIOS${i}`))i++;const id=`rIdMIOS${i}`;rels(relRoot).append(`<Relationship Id="${id}" Type="${type}" Target="${esc(target)}"/>`);return id;}
 function ensureMainPart(type,target){if(!elements(rels,relRoot,'Relationship').some(n=>val(n,'Type')===type))mainRel(type,target);}
 const copied=new Map();
 async function importPart(part){if(copied.has(part))return copied.get(part);if(!tpl.file(part))throw Error('範本資源遺失：'+part);const destination=`mios-template-${policy.sha256.slice(0,8)}/${part}`;copied.set(part,destination);let bytes=await tpl.file(part).async('uint8array');
  if(part.endsWith('.xml')){const doc=parse(new TextDecoder().decode(bytes));const docRoot=doc.root().children().toArray().find(n=>n.type==='tag');if(docRoot)remap(doc,docRoot);bytes=new TextEncoder().encode(doc.xml());}
  zip.file(destination,bytes,{createFolders:false});
  const override=elements(tTypes,ttRoot,'Override').find(n=>val(n,'PartName')==='/'+part);const ext=part.split('.').pop();const defaultType=elements(tTypes,ttRoot,'Default').find(n=>val(n,'Extension')===ext);addType(destination,val(override,'ContentType')||val(defaultType,'ContentType')||'application/octet-stream');
  const rp=relPath(part);if(tpl.file(rp)){const r$=parse(await tpl.file(rp).async('string'));for(const rel of r$('*').filter((_,n)=>local(n)==='Relationship').toArray()){if(val(rel,'TargetMode')==='External')continue;const imported=await importPart(resolve(part,val(rel,'Target')));r$(rel).attr('Target',relative(destination,imported));}zip.file(relPath(destination),r$.xml(),{createFolders:false});}
  return destination;
 }
 const tRels=parse(await tpl.file('word/_rels/document.xml.rels').async('string'));
 const sectionProto=elements(t$,tBody,'sectPr').at(-1);if(!sectionProto)throw Error('範本沒有頁面設定。');
 let sectionXML=t$.xml(sectionProto);
 const sectionClone=$(sectionXML).get(0);
 const pageSize=first($,sectionClone,'pgSz');if(!pageSize)throw Error('範本缺少頁面尺寸。');for(const a of Object.keys(pageSize.attribs))if(['orient','code'].includes(a.split(':').pop()))$(pageSize).removeAttr(a);put($,pageSize,'w',11906);put($,pageSize,'h',16838);
 for(const ref of ['headerReference','footerReference'])for(const n of elements($,sectionClone,ref)){const target=tRels('*').filter((_,r)=>local(r)==='Relationship'&&val(r,'Id')===val(n,'id')).get(0);if(!target)throw Error('範本頁首頁尾關聯遺失。');const imported=await importPart(resolve('word/document.xml',val(target,'Target')));const id=mainRel(val(target,'Type'),relative('word/document.xml',imported));const old=Object.keys(n.attribs).find(a=>a.split(':').pop()==='id');$(n).attr(old||'r:id',id);}
 sectionXML=$.xml(sectionClone);
 const changes=[],warnings=[];const nodes=all($,body,'p').filter(p=>preferred($,p));
 const originalText=nodes.map(p=>text($,p));const styleInfo=new Map(before.paragraphs.map(p=>[p.number,p]));
 const top=nodes.map((p,i)=>({p,i,text:text($,p)})).filter(x=>x.text.trim()&&!$(x.p).parents().toArray().some(a=>['tc','txbxContent'].includes(local(a))));
 const tOf=x=>x?.text.replace(/\s/g,'')||'';
 const findSection=word=>top.find(x=>new RegExp('^(?:[一二三四五六七八九十0-9、.．（）()]+)?'+word+'[：:]?$').test(tOf(x)));
 const graphics=findSection('圖表規範');
 const refs=findSection('參考文獻'),conclusion=findSection('結論'),bio=findSection('作者簡介'),abstract=findSection('摘要');
 const hasChineseHeadings=top.some(x=>/^[一二三四五六七八九十]+[、．]/.test(x.text));
 const heading=(info,s)=>{if(info?.level&&info.level<=4&&!/這是.*範例/.test(s))return info.level;const marker=s.match(/Heading\s*([1-4])/i);if(marker)return Number(marker[1]);if(/^[一二三四五六七八九十]+[、．]/.test(s))return 1;if(/^[（(][一二三四五六七八九十]+[）)]/.test(s))return 2;const decimal=s.match(/^(\d+(?:\.\d+)+)[.、\s]+[^\d]/);if(decimal)return Math.min(4,decimal[1].split('.').length);if(/^\d+[、.．]\s*[^\d]/.test(s))return hasChineseHeadings?3:1;if(/^[（(]\d+[）)]/.test(s))return hasChineseHeadings?4:2;return null;};
 const abstractEnd=abstract?top.find(x=>x.i>abstract.i&&heading(styleInfo.get(x.i+1),x.text))?.i??Infinity:0;
 const author=top[1]&&top[1].text.length<110&&!heading(styleInfo.get(top[1].i+1),top[1].text)&&!/[。！？]/.test(top[1].text)?top[1]:null;
 if(author&&/^[\u3400-\u9fff]{2,5}(?:\s*[,，;；]\s*[\u3400-\u9fff]{2,5})+$/.test(author.text.trim())){
  for(const t of all($,author.p,'t'))$(t).text($(t).text().replace(/[,，;；]/g,'、'));
  originalText[author.i]=text($,author.p);
  changes.push({paragraph:author.i+1,location:`第 ${author.i+1} 段`,excerpt:author.text,field:'作者分隔符號',before:author.text,after:text($,author.p)});
 }else if(author&&/[，,；;／/＆&]/.test(author.text))warnings.push('作者列含姓名與單位等混合資訊，無法安全區分分隔符號；請人工將作者姓名間改用頓號「、」。');
 if(!author)warnings.push('無法確定作者列；未補寫作者，請人工確認標題後的作者位置。');
 if(!abstract)warnings.push('未找到獨立摘要標題，開頭的一般文字保留並套用摘要／內文格式，請確認摘要範圍。');
 for(const[name,item]of [['結論',conclusion],['參考文獻',refs],['作者簡介',bio]])if(!item)warnings.push(`未辨識到「${name}」獨立章節，保留原文，不新增虛構內容。`);
 let active=1;
 function prototype(role){const p=tParas[base[role]??base.Body1];const pPr=t$.xml(first(t$,p,'pPr'))||'<w:pPr/>';const result=$(pPr).get(0);for(const k of ['pStyle','rPr','numPr','sectPr'])remove($,result,k);let ind=first($,result,'ind');if(!ind){const nid={H1:'9',H2:'10',H3:'11',H4:'12',Reference:'8'}[role];const num=nid&&elements(tNum,tnRoot,'num').find(n=>val(n,'numId')===nid);const aid=num&&val(first(tNum,num,'abstractNumId'));const abs=aid&&elements(tNum,tnRoot,'abstractNum').find(n=>val(n,'abstractNumId')===aid);const ni=abs&&first(tNum,first(tNum,first(tNum,abs,'lvl'),'pPr'),'ind');$(result).append(ni?tNum.xml(ni):'<w:ind/>');ind=first($,result,'ind');}for(const k of ['left','right'])if(val(ind,k)===undefined)put($,ind,k,0);for(const k of ['leftChars','rightChars'])if(val(ind,k)===undefined&&val(ind,k==='leftChars'?'left':'right')==='0')put($,ind,k,0);
  if(['Body3','Body4','Caption','Cell','Author','Bio'].includes(role)){for(const k of ['firstLine','firstLineChars'])put($,ind,k,0);for(const a of Object.keys(ind.attribs))if(['hanging','hangingChars'].includes(a.split(':').pop()))$(ind).removeAttr(a);}
  if(['Body1','Body2','Abstract','Conclusion'].includes(role)){put($,ind,'firstLineChars',200);put($,ind,'firstLine',480);}
  if(role==='Cell'||role==='Caption')put($,ind,'left',0);
  let spacing=first($,result,'spacing');if(!spacing){$(result).append('<w:spacing/>');spacing=first($,result,'spacing');}
  for(const[k,v]of [['before','0'],['after','0'],['line','288'],['lineRule','auto']])if(val(spacing,k)===undefined)put($,spacing,k,v);
  if(['Author','Title','H1','H2','H3','H4','Reference','Bio'].includes(role)&&!first(t$,first(t$,tParas[base[role]],'pPr'),'spacing'))put($,spacing,'line',240);
  if(role==='Caption'){remove($,result,'jc');$(result).append('<w:jc w:val="center"/>');}
  if(['Body1','Body2','Body3','Body4','Abstract','Conclusion'].includes(role)){remove($,result,'jc');$(result).append('<w:jc w:val="both"/>');put($,spacing,'line',288);put($,spacing,'lineRule','auto');}
  if(!first($,result,'jc'))$(result).append(`<w:jc w:val="${role==='Author'||role==='Title'?'center':role==='Caption'?'center':role.startsWith('H')||role==='Reference'||role==='Bio'?'left':'both'}"/>`);
  $(result).append('<w:snapToGrid w:val="0"/>');return result;
 }
 const roleStyle=new Map();
 const ppOrder='pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'.split(' ');
 function sortProps(node){const children=$(node).children().toArray();children.sort((a,b)=>{const ai=ppOrder.indexOf(local(a)),bi=ppOrder.indexOf(local(b));return (ai<0?100:ai)-(bi<0?100:bi)});for(const c of children)$(node).append(c);return node;}
 function styleFor(role){if(roleStyle.has(role))return roleStyle.get(role);const id=prefix+'Role_'+role;for(const old of elements(styles,sRoot,'style').filter(n=>val(n,'styleId')===id))styles(old).remove();const p=prototype(role);sortProps(p);const pt=role==='Title'?32:role==='H1'?14:12;const bold=role==='H1';styles(sRoot).append(`<w:style w:type="paragraph" w:styleId="${id}" w:customStyle="1"><w:name w:val="${names[role]} · 量測資訊範本"/><w:qFormat/>${$.xml(p)}<w:rPr><w:rFonts w:ascii="微軟正黑體" w:hAnsi="微軟正黑體" w:eastAsia="微軟正黑體" w:cs="微軟正黑體"/><w:b w:val="${bold?1:0}"/><w:bCs w:val="${bold?1:0}"/><w:color w:val="000000"/><w:sz w:val="${pt*2}"/><w:szCs w:val="${pt*2}"/></w:rPr></w:style>`);roleStyle.set(role,id);return id;}
 for(const [index,p]of nodes.entries()){
  const drawings=all($,p,'drawing');const legacy=all($,p,'pict');
  if(drawings.length||legacy.length){
   let props=first($,p,'pPr');if(!props){$(p).prepend('<w:pPr/>');props=first($,p,'pPr');}
   remove($,props,'jc');$(props).append('<w:jc w:val="center"/>');remove($,props,'ind');$(props).append('<w:ind w:left="0" w:right="0" w:firstLine="0" w:firstLineChars="0"/>');sortProps(props);
   for(const anchor of drawings.flatMap(d=>all($,d,'anchor'))){let pos=first($,anchor,'positionH');if(!pos){$(anchor).children().first().after('<wp:positionH xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" relativeFrom="column"><wp:align>center</wp:align></wp:positionH>');pos=first($,anchor,'positionH');}$(pos).attr('relativeFrom','column');$(pos).empty().append('<wp:align xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">center</wp:align>');}
   if(legacy.length)warnings.push(`第 ${index+1} 段含舊式圖形，已置中段落，圖形絕對位置請於 Word 人工確認。`);
   changes.push({paragraph:index+1,location:`第 ${index+1} 段`,excerpt:text($,p).slice(0,120),field:'圖片置中',before:'原稿圖片位置',after:'圖片段落置中；浮動圖片於欄內水平置中'});
  }
  let content=text($,p);if(!content.trim())continue;
  const figurePrefix=content.match(/^(圖\s*[一二三四五六七八九十百\d]+)[\s、，,.．：:]+(?=\S)/);
  if(figurePrefix){
   const replacement=figurePrefix[1].replace(/\s/g,'')+'  ';let remaining=figurePrefix[0].length,inserted=false;
   for(const t of all($,p,'t').filter(t=>preferred($,t))){if(!remaining)break;const str=$(t).text();if(!str.length)continue;const consumed=Math.min(str.length,remaining);$(t).text((!inserted?replacement:'')+str.slice(consumed)).attr('xml:space','preserve');if(str.length)inserted=true;remaining-=consumed;}
   const updated=text($,p);if(updated!==content)changes.push({paragraph:index+1,location:`第 ${index+1} 段`,excerpt:content.slice(0,120),field:'圖說編號分隔',before:content,after:updated});
   content=updated;originalText[index]=content;
  }

  const parents=$(p).parents().toArray();if(parents.some(a=>local(a)==='txbxContent')){warnings.push(`第 ${index+1} 段為文字方塊，位置需人工確認。`);continue;}
  const info=styleInfo.get(index+1);const level=heading(info,content);let role='Body'+active;
  if(/^圖\s*[一二三四五六七八九十百\d]+[\s、.．：:]/.test(content))role='Caption';
  else if(parents.some(a=>local(a)==='tc'))role='Cell';
  else if(top[0]?.p===p)role='Title';else if(author?.p===p)role='Author';
  else if([refs?.p,conclusion?.p,bio?.p,abstract?.p,graphics?.p].includes(p)){role='H1';active=1;}
  else if(/^[^/／]{1,40}\s*[/／]\s*[^/／]{1,100}\s*[/／]\s*[^/／]{1,50}$/.test(content))role='Bio';
  else if(conclusion&&index>conclusion.i&&(!refs||index<refs.i))role='Conclusion';
  else if(bio&&index>bio.i)role='Bio';else if(refs&&index>refs.i)role='Reference';
  else if(abstract&&index>abstract.i&&index<abstractEnd)role='Abstract';
  else if(level){role='H'+level;active=level;}
  else if(/^[圖表]\s*[一二三四五六七八九十百\d]+[\s、.．：:]/.test(content))role='Caption';
  else if(!abstract&&author&&index>author.i&&index<(top.find(x=>x.i>author.i&&heading(styleInfo.get(x.i+1),x.text))?.i??Infinity))role='Abstract';
  if(!base[role]&&base[role]!==0)role='Body1';
  const oldPr=first($,p,'pPr');const preservedSection=oldPr&&first($,oldPr,'sectPr');const oldNum=oldPr&&first($,oldPr,'numPr');const pp=prototype(role);$(pp).prepend(`<w:pStyle w:val="${styleFor(role)}"/>`);
  const manualNumber=/^(?:[一二三四五六七八九十]+[、．]|[（(][一二三四五六七八九十\d]+[）)]|\d+[、.．]|\[\d+\])/.test(content);
  if(manualNumber&&role.startsWith('H'))warnings.push('原稿以文字輸入的章節編號保留，已統一層級與版型；自動編號的樣式仍請人工確認。');
  const numSource=role.startsWith('H')?{H1:'9',H2:'10',H3:'11',H4:'12'}[role]:role==='Reference'?'8':null;
  const sectionHeading=[refs?.p,conclusion?.p,bio?.p,abstract?.p,graphics?.p].includes(p);
  if(numSource&&!manualNumber&&!sectionHeading)$(pp).append(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numMap.get(numSource)}"/></w:numPr>`);else if(oldNum&&!numSource&&role!=='Caption')$(pp).append($.xml(oldNum));
  if(role==='Caption')$(pp).append('<w:numPr><w:numId w:val="0"/></w:numPr>');
  if(drawings.length||legacy.length){remove($,pp,'jc');$(pp).append('<w:jc w:val="center"/>');remove($,pp,'ind');$(pp).append('<w:ind w:left="0" w:right="0" w:firstLine="0" w:firstLineChars="0"/>');}
  if(role.startsWith('H'))$(pp).append(`<w:outlineLvl w:val="${Number(role.slice(1))-1}"/>`);else $(pp).append('<w:outlineLvl w:val="9"/>');
  if(preservedSection)$(pp).append($.xml(preservedSection));sortProps(pp);if(oldPr)$(oldPr).replaceWith(pp);else $(p).prepend(pp);
  for(const run of all($,p,'r').filter(r=>$(r).parents().toArray().find(a=>local(a)==='p')===p&&preferred($,r))){if(!all($,run,'t').length)continue;let rp=first($,run,'rPr');if(!rp){$(run).prepend('<w:rPr/>');rp=first($,run,'rPr');}
   for(const k of ['rFonts','sz','szCs','color','spacing','w','position','kern'])remove($,rp,k);
   const pt=role==='Title'?32:role==='H1'?14:12;$(rp).append(`<w:rFonts w:ascii="微軟正黑體" w:hAnsi="微軟正黑體" w:eastAsia="微軟正黑體" w:cs="微軟正黑體"/><w:color w:val="000000"/><w:sz w:val="${pt*2}"/><w:szCs w:val="${pt*2}"/>`);
   if(role==='Title'||role==='Author'||role.startsWith('H')){remove($,rp,'b');remove($,rp,'bCs');remove($,rp,'i');remove($,rp,'iCs');$(rp).append(`<w:b w:val="${role==='H1'?1:0}"/><w:bCs w:val="${role==='H1'?1:0}"/><w:i w:val="0"/><w:iCs w:val="0"/>`);}
   const ro='rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange'.split(' ');const kids=$(rp).children().toArray().sort((a,b)=>{const ai=ro.indexOf(local(a)),bi=ro.indexOf(local(b));return(ai<0?100:ai)-(bi<0?100:bi)});for(const k of kids)$(rp).append(k);
  }
  changes.push({paragraph:index+1,location:`第 ${index+1} 段`,excerpt:content.slice(0,120),field:'範本套版 · '+names[role],before:info?.style||'原稿格式',after:role==='Title'?'附件標題版型・32pt':role==='H1'?'附件章節標題・14pt':'附件'+names[role]+'版型・12pt'});
 }
 const sections=all($,body,'sectPr');if(!sections.length)$(body).append(sectionXML);else for(const section of sections){const originalType=first($,section,'type');const clone=$(sectionXML).get(0);if(originalType){remove($,clone,'type');const beforePage=first($,clone,'pgSz');if(beforePage)$(beforePage).before($.xml(originalType));else $(clone).append($.xml(originalType));}$(section).replaceWith(clone);}
 changes.push({paragraph:null,location:'整份文件',excerpt:'',field:'頁面與頁首頁尾',before:'原稿頁面設定',after:'附件實際頁面：A4 21 × 29.7 cm；採用附件邊界、頁首及頁尾'});
 if(before.summary.tables)warnings.push('範本未提供實際表格樣式；保留原稿表格結構並統一文字字體，表格寬度與分頁請人工確認。');
 if(before.summary.images)warnings.push('圖表資源及順序保留；浮動圖片、解析度及套版後的位置仍需人工確認。');
 const finalText=all($,body,'p').filter(p=>preferred($,p)).map(p=>text($,p));if(JSON.stringify(originalText)!==JSON.stringify(finalText))throw Error('內容完整性檢查未通過，未產生套版檔。原稿未變更。');
 zip.file('word/document.xml',$.xml(),{createFolders:false});zip.file('word/styles.xml',styles.xml(),{createFolders:false});zip.file('word/numbering.xml',number.xml(),{createFolders:false});
 ensureMainPart(R+'/styles','styles.xml');ensureMainPart(R+'/numbering','numbering.xml');addType('word/styles.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');addType('word/numbering.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');zip.file('word/_rels/document.xml.rels',rels.xml(),{createFolders:false});zip.file('[Content_Types].xml',types.xml(),{createFolders:false});
 const bytes=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});const after=await analyzeArticle(bytes,enabled);
 return {bytes,changes,before,after,mode:'template',warnings:[...new Set(warnings)],template:{source:policy.source,sha256:policy.sha256,page:'A4',widthTwips:11906,heightTwips:16838},limitations:['輸出統一採 A4 版面與附件、段落角色及字型設定；檢核開關不影響套版標準。','文章文字順序與圖表資源保留，頁首頁尾改採附件；不插入範例作者、照片或範例內文。','缺少章節、無法判定的作者／摘要、複雜圖表及文字方塊仍需人工複核，實際分頁需以 Word 確認。']};
}
