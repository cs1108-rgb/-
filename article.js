import JSZip from 'jszip';
import {load} from 'cheerio/slim';
import {imageSize} from 'image-size';
import {policy} from './policy.js';
export {policy} from './policy.js';


export const articleDefaults={structure:true,typography:true,paragraph:true,headings:true,figures:true,references:true};
export const ruleDefinitions=[
 ['structure','文章結構與範例文字','標題、作者、摘要、結論、參考文獻、作者簡介；提示未替換的範例內容。'],
 ['typography','字體與字級','微軟正黑體；文章標題 32pt，作者、摘要、內文 12pt。未明訂的章節標題字級不強制判定。'],
 ['paragraph','段落排版','A4 直式與指定範本邊界；摘要與內文左右對齊、1.2 倍行距；一般內文首行縮排兩字元，Heading 3／4 下內文不縮排。'],
 ['headings','標題層級','辨識 Heading 1–4 與大綱層級，提示跳級或以一般段落模擬的章節標題。'],
 ['figures','圖表與文中引用','圖說在下、表題在上、編號唯一、文中引用加粗；檢查圖片格式、尺寸與可讀取的有效 DPI。'],
 ['references','文獻與人工複核','提示空白文獻、網址句號與不完整作者簡介；內容正確性、線條圖類型及文獻完整性須人工確認。']
].map(([key,title,description])=>({key,title,description}));

function packageTarget(target){const parts=[];for(const segment of (target.startsWith('/')?target:'word/'+target).split('/')){if(segment==='..')parts.pop();else if(segment&&segment!=='.')parts.push(segment);}return parts.join('/');}
const local=n=>n?.name?.split(':').pop();
const direct=($,n,name)=>$(n).children().toArray().find(c=>local(c)===name);
const descendants=($,n,name)=>$(n).find('*').toArray().filter(c=>local(c)===name);
const attr=(n,k)=>n?.attribs?.[k]??Object.entries(n?.attribs||{}).find(([a])=>a.split(':').pop()===k)?.[1];
const value=($,n,k)=>attr(direct($,n,k),'val');
const truth=v=>v!==undefined&&!['0','false','off'].includes(v);
const xml=s=>load(s,{xml:true});
function props($,n){
 if(!n)return {};
 const r={};const fonts=direct($,n,'rFonts');
 for(const k of ['ascii','hAnsi','eastAsia','asciiTheme','hAnsiTheme','eastAsiaTheme'])if(attr(fonts,k)!==undefined)r[k]=attr(fonts,k);
 if(value($,n,'sz')!==undefined)r.size=Number(value($,n,'sz'))/2;
 const bold=direct($,n,'b');if(bold)r.bold=truth(attr(bold,'val')??'1');
 if(value($,n,'jc')!==undefined)r.align=value($,n,'jc');
 const spacing=direct($,n,'spacing');for(const k of ['line','lineRule'])if(attr(spacing,k)!==undefined)r[k]=attr(spacing,k);
 const ind=direct($,n,'ind');for(const k of ['firstLine','firstLineChars','hanging','hangingChars'])if(attr(ind,k)!==undefined)r[k]=Number(attr(ind,k));
 if(value($,n,'outlineLvl')!==undefined)r.outline=Number(value($,n,'outlineLvl'));
 return r;
}
function merge(...layers){const o={};for(const l of layers){if(l?.firstLine!==undefined&&l.firstLineChars===undefined)delete o.firstLineChars;if(l?.hanging!==undefined&&l.hangingChars===undefined)delete o.hangingChars;for(const k of ['ascii','hAnsi','eastAsia']){if(l?.[k+'Theme']!==undefined&&l?.[k]===undefined)delete o[k];if(l?.[k]!==undefined)delete o[k+'Theme'];}Object.assign(o,l||{});}return o;}
export async function analyzeArticle(buffer,enabled=articleDefaults){
 let zip,$,style$,theme$;
 try{
  zip=await JSZip.loadAsync(buffer);const files=Object.values(zip.files);if(files.length>4000||files.reduce((s,f)=>s+(f._data?.uncompressedSize||0),0)>80*1024*1024)throw Error('文件解壓縮後過大，請壓縮圖片或拆分文件。');
  if(!zip.file('word/document.xml'))throw Error('不是有效的 Word 文件');
  $=xml(await zip.file('word/document.xml').async('string'));
  style$=xml(await zip.file('word/styles.xml')?.async('string')||'<styles/>');
  theme$=xml(await zip.file('word/theme/theme1.xml')?.async('string')||'<theme/>');
 }catch(e){if(e.message.includes('解壓縮'))throw e;throw Error('無法讀取 Word 文件，請上傳有效且未加密的 .docx 檔案；舊版 .doc 請先另存為 .docx。');}
 // Word stores both a modern and a legacy copy of shapes; inspect the preferred copy once.
 $('*').filter((_,n)=>local(n)==='AlternateContent').each((_,n)=>{const choice=direct($,n,'Choice');if(choice)$(n).children().filter((_,c)=>c!==choice).remove();});
 const styleNodes=style$('*').filter((_,n)=>local(n)==='style').toArray();
 const styles=new Map(styleNodes.map(n=>[attr(n,'styleId'),n]));
 const defaultParagraph=styleNodes.find(n=>attr(n,'type')==='paragraph'&&attr(n,'default')==='1');
 const dd=style$('*').filter((_,n)=>local(n)==='docDefaults').get(0);
 const baseR=props(style$,direct(style$,direct(style$,dd,'rPrDefault'),'rPr'));
 const baseP=props(style$,direct(style$,direct(style$,dd,'pPrDefault'),'pPr'));
 function chain(id,seen=new Set()){
  if(!id||seen.has(id)||!styles.has(id))return [];
  seen.add(id);const s=styles.get(id);return [...chain(value(style$,s,'basedOn'),seen),s];
 }
 const themeFont=(token,cjk)=>{
  if(!token)return undefined;const scheme=theme$('*').filter((_,n)=>local(n)===(token.startsWith('major')?'majorFont':'minorFont')).get(0);
  if(cjk){const zh=descendants(theme$,scheme,'font').find(n=>attr(n,'script')==='Hant');return attr(direct(theme$,scheme,'ea'),'typeface')||attr(zh,'typeface');}
  return attr(direct(theme$,scheme,'latin'),'typeface');
 };
 function fontFor(r,cjk){const k=cjk?'eastAsia':'ascii';return r[k]||themeFont(r[k+'Theme'],cjk)||r.hAnsi||themeFont(r.hAnsiTheme,false);}
 const body=$('*').filter((_,n)=>local(n)==='body').get(0);if(!body)throw Error('文件缺少可讀取的本文。');
 const nodes=descendants($,body,'p').filter(n=>!$(n).parents().toArray().some(a=>['del','moveFrom'].includes(local(a))));
 if(nodes.length>10000)throw Error('文件最多支援 10,000 個段落，請拆分後上傳。');
 const paragraphs=nodes.map((n,i)=>{
  const pPr=direct($,n,'pPr');const id=value($,pPr,'pStyle')||attr(defaultParagraph,'styleId');const stack=chain(id);
  const p=merge(baseP,...stack.map(s=>props(style$,direct(style$,s,'pPr'))),props($,pPr));
  const rbase=merge(baseR,...stack.map(s=>props(style$,direct(style$,s,'rPr'))));
  const runs=descendants($,n,'r').filter(r=>$(r).closest('w\\:p')[0]===n||$(r).parents().toArray().find(a=>local(a)==='p')===n).filter(r=>!$(r).parents().toArray().some(a=>['del','moveFrom'].includes(local(a)))).map(r=>{
   const rp=direct($,r,'rPr');const charStyles=chain(value($,rp,'rStyle'));const effective=merge(rbase,...charStyles.map(s=>props(style$,direct(style$,s,'rPr'))),props($,rp));
   const text=descendants($,r,'t').map(t=>$(t).text()).join('');return {text,size:effective.size,bold:effective.bold??false,font:fontFor(effective,/[\u3400-\u9fff]/.test(text))};
  }).filter(r=>r.text.length);
  const text=runs.map(r=>r.text).join('');const styleName=value(style$,styles.get(id),'name')||id||'';
  const match=styleName.match(/heading\s*([1-9])/i)||text.match(/(?:章節標題|小節標題|條列標題|次條列標題).*Heading\s*([1-4])/i);
  const level=p.outline!==undefined&&p.outline<9?p.outline+1:match?Number(match[1]):null;
  const inTable=$(n).parents().toArray().some(a=>local(a)==='tc');const inTextbox=$(n).parents().toArray().some(a=>local(a)==='txbxContent');
  const drawings=descendants($,n,'drawing').filter(d=>$(d).parents().toArray().find(a=>local(a)==='p')===n);
  return {number:i+1,text,runs,format:p,style:styleName,level,inTable,inTextbox,drawings,node:n};
 });
 const textParas=paragraphs.filter(p=>p.text.trim()&&!p.inTable&&!p.inTextbox);
 if(!textParas.length)throw Error('找不到可檢核的文章文字。掃描圖片請先轉成可編輯的 Word 文字。');
 const checks=[];const issues=[];
 function record(group,p,field,ok,message,suggestion,actual,expected,warning=false){
  if(enabled[group]===false)return;
  const status=ok===null?'manual':ok?'passed':warning?'warning':'failed';
  const c={group,field,status,paragraph:p?.number??null,location:p?`第 ${p.number} 段`:'整份文件',excerpt:p?.text?.slice(0,120)||'',actual:actual??null,expected:expected??null};checks.push(c);
  if(status!=='passed')issues.push({...c,rule:group,level:status==='failed'?'error':'warning',message,suggestion});
 }
 const sections=descendants($,body,'sectPr');
 if(!sections.length)record('paragraph',null,'文件頁面設定',null,'文件未儲存明確頁面設定','請在 Word 設定 A4 直式，或使用「依範本匯出 Word」統一套版。');
 for(const[si,section]of sections.entries()){
  const pg=direct($,section,'pgSz');const width=Number(attr(pg,'w')),height=Number(attr(pg,'h'));
  record('paragraph',null,`第 ${si+1} 節 A4 頁面`,pg?Math.abs(width-11906)<=5&&Math.abs(height-16838)<=5:null,'頁面尚未設定為 A4 直式','請選擇 A4（21 × 29.7 公分），或使用「依範本匯出 Word」自動統一。',pg?`${(width/1440*2.54).toFixed(2)} × ${(height/1440*2.54).toFixed(2)} cm`:'未設定','A4 21 × 29.7 cm');
  const margins=direct($,section,'pgMar');const known=margins&&Object.keys(policy.outputMarginsTwips).every(k=>attr(margins,k)!==undefined);
  record('paragraph',null,`第 ${si+1} 節頁面邊界`,known?Object.entries(policy.outputMarginsTwips).every(([k,v])=>Math.abs(Number(attr(margins,k))-v)<=5):null,'頁面邊界與指定範本不同','請使用「依範本匯出 Word」套用指定範本的邊界與頁首頁尾。',known?Object.entries(policy.outputMarginsTwips).map(([k])=>`${k}=${attr(margins,k)}`).join(', '):'未設定完整邊界',Object.entries(policy.outputMarginsTwips).map(([k,v])=>`${k}=${v}`).join(', '));
 }
 const first=textParas[0],author=textParas[1];const section=t=>textParas.find(p=>p.text.replace(/\s+/g,'').replace(/^[一二三四五六七八九十0-9、.．（）()]+/,'').startsWith(t));
 const refs=section('參考文獻'),conclusion=section('結論'),bio=section('作者簡介');
 const headingLike=p=>p.level||/^(?:[一二三四五六七八九十]+、|[（(][一二三四五六七八九十]+[）)]|\d+(?:\.\d+)*[.、．\s])/.test(p.text)||['結論','參考文獻','圖表規範','作者簡介'].some(t=>p.text.replace(/\s/g,'').startsWith(t));
 const abstractExplicit=textParas.find(p=>/^摘要\s*[:：]?$/.test(p.text.trim()));
 const abstract=abstractExplicit||textParas[2];
 record('structure',first,'文章標題',!!first,'缺少文章標題','請在文件最前方填寫文章標題。');
 record('structure',author,'作者',!!author&&!headingLike(author),'未能辨識標題後的作者列','請將作者姓名獨立放在標題後的下一個非空白段落。');
 record('structure',abstract,'摘要',abstractExplicit?true:null,'摘要位置需確認','範本允許摘要不加標題；目前以作者後的第一段一般文字推定，請確認該段確為摘要。');
 for(const [title,p]of [['結論',conclusion],['參考文獻',refs],['作者簡介',bio]])record('structure',p,title,!!p,`未找到「${title}」段落`,`請加入「${title}」標題並填寫對應內容。`);
 for(const p of textParas){if(/這是(?:摘要|內文|結論內文)範例|作者\d|OOO研究室|單位名稱|標題[（(]32pt|姓名，出版資訊完整列出/.test(p.text))record('structure',p,'範例文字',false,'疑似尚未替換的範本內容','請將範例文字替換為本期實際文章、作者及所屬單位。',p.text.slice(0,80),'實際投稿內容',true);}
 let activeLevel=1,previousLevel=0;
 for(const p of textParas){
  const special=[first,author,refs,bio,conclusion,abstractExplicit].includes(p);const caption=/^[圖表]\s*[一二三四五六七八九十百\d]+(?:[\s、.．：:]|$)/.test(p.text);
  if(p===conclusion)activeLevel=1;
  if(p.level&&p!==abstract&&p!==first&&p!==author){if(previousLevel&&p.level>previousLevel+1)record('headings',p,'標題層級',false,`標題由 Heading ${previousLevel} 跳至 Heading ${p.level}`,'請調整 Word 標題樣式，依 Heading 1、2、3、4 依序使用。',p.level,previousLevel+1,true);else record('headings',p,'標題層級',p.level<=4,'使用超出範本的標題層級','請使用 Heading 1–4，或人工確認此層級是否必要。',p.level,'1–4',true);activeLevel=p.level;previousLevel=p.level;}
  else if(!special&&p!==abstract&&headingLike(p)){record('headings',p,'標題樣式',null,'疑似章節標題未使用可辨識的大綱層級','請於 Word 套用 Heading 1–4；若此段為一般文字，可忽略此提示。');}
  const isHeading=!!p.level||headingLike(p);const isBody=!special&&!isHeading&&!caption&&p!==first&&p!==author&&(!refs||p.number<refs.number)&&(!bio||p.number<bio.number);
  const isAbstract=p===abstract&&!abstractExplicit||abstractExplicit&&p.number===textParas[textParas.indexOf(abstractExplicit)+1]?.number;
  const knownSize=p===first?32:(p===author||isBody||isAbstract)?12:null;
  const visible=p.runs.filter(r=>r.text.trim());const fonts=[...new Set(visible.map(r=>r.font||'未能解析的主題字體'))];
  const knownFonts=visible.filter(r=>r.font);const wrongFonts=knownFonts.filter(r=>!['微軟正黑體','Microsoft JhengHei','Microsoft JhengHei UI'].includes(r.font));
  record('typography',p,'字體',wrongFonts.length?false:knownFonts.length===visible.length?true:null,wrongFonts.length?'字體與範本不符':'部分字體無法可靠解析','請選取此段文字，將中英文字體設定為「微軟正黑體」。',fonts.join('、'),'微軟正黑體');
  if(knownSize){const sizes=[...new Set(visible.map(r=>r.size??'未知'))];record('typography',p,'字級',visible.some(r=>r.size!==undefined&&r.size!==knownSize)?false:visible.every(r=>r.size!==undefined)?true:null,`字級應為 ${knownSize}pt`,'請選取本段文字，於 Word 字型設定修改字級。',sizes.map(s=>s+'pt').join('、'),knownSize+'pt');}
  if(isBody||isAbstract){
   record('paragraph',p,'左右對齊',(p.format.align||'left')==='both','段落未設定為左右對齊','在 Word「段落」將對齊方式設為「左右對齊」（Ctrl+J）。',p.format.align||'left','both');
   const line=p.format.line?Number(p.format.line):240;const lineRule=p.format.lineRule||'auto';record('paragraph',p,'行距',lineRule==='auto'&&line===288,'段落行距不是 1.2 倍','在 Word「段落 → 行距」選多行，設定為 1.2 倍。',lineRule==='auto'?line/240+' 倍':`${lineRule} ${line}`,'1.2 倍');
   const chars=p.format.firstLineChars;const twips=p.format.firstLine??0;const hanging=p.format.hanging||p.format.hangingChars||0;const expected=activeLevel>=3&&!isAbstract?0:2;
   const indentOk=!hanging&&(chars!==undefined?chars===expected*100:expected===0?twips===0:Math.abs(twips-480)<=12);
   record('paragraph',p,'首行縮排',indentOk,expected?'首行應縮排兩字元':'此層級內文首行不應縮排',`在 Word「段落 → 特殊格式」${expected?'設定首行縮排 2 字元':'設定為「無」，取消首行與凸排'}。`,chars!==undefined?chars/100+' 字元':twips/20+' pt',expected+' 字元');
  }
 }
 if(enabled.headings!==false&&!textParas.some(p=>p.level))record('headings',null,'標題辨識',null,'未辨識到 Word 標題層級','請為章節套用 Heading 1–4；範本中的部分標題僅有外觀格式，系統無法由外觀確定層級。');
 if(conclusion){const next=textParas[textParas.indexOf(conclusion)+1];const explicitSection=next&&['參考文獻','作者簡介','結論'].some(t=>next.text.replace(/\s/g,'')===t);const uncertain=next&&!explicitSection&&headingLike(next);record('structure',conclusion,'結論內容',!next||explicitSection?false:uncertain?null:true,uncertain?'結論後段落使用標題樣式，內容角色需確認':'結論段落沒有內文',uncertain?'請確認結論標題後是否已有內文；若是內文，請改用一般內文樣式。':'請於結論標題下補上研究結果與重點結論。');}
 if(refs){const content=textParas.filter(p=>p.number>refs.number&&(!bio||p.number<bio.number));record('references',refs,'文獻內容',content.length>0,'參考文獻未填寫內容','請在參考文獻標題下列出作者、出版資料或網址。');for(const p of content)if(/https?:\/\//i.test(p.text))record('references',p,'網址標點',!/[。．.]\s*$/.test(p.text),'網址結尾不需句號','請移除網址末尾額外加上的句號；網址本身的必要符號請保留。');record('references',refs,'文獻正確性',null,'文獻出版資訊與引用關係須人工確認','請確認作者、篇名、出版資訊及網址完整正確，且與內文引用一致。');}
 if(bio){const content=textParas.filter(p=>p.number>=bio.number);record('references',bio,'作者簡介內容',content.some(p=>/[^/]+[/／][^/]+[/／][^/]+/.test(p.text)), '作者簡介缺少姓名／單位／職稱','請依「作者姓名／單位名稱／職稱」補齊每位作者的資料。',null,'姓名／單位／職稱',true);}
 // Match captions and references; heading/body excerpt offsets are preserved for split bold runs.
 const captions=textParas.filter(p=>/^[圖表]\s*[一二三四五六七八九十百\d]+(?:[\s、.．：:]|$)/.test(p.text));
 const normalize=s=>s.replace(/\s/g,'');const keyOf=p=>normalize(p.text.match(/^[圖表]\s*[一二三四五六七八九十百\d]+/)[0]);
 const capKeys=new Set();
 for(const p of captions){const key=keyOf(p);record('figures',p,'圖表編號',!capKeys.has(key),`圖表編號「${key}」重複`,'請重新編排圖表編號，並同步更新文中引用。');capKeys.add(key);
  const matches=textParas.filter(q=>!captions.includes(q)&&[...q.text.matchAll(/[圖表]\s*[一二三四五六七八九十百\d]+/g)].some(m=>normalize(m[0])===key));record('figures',p,'文中引用',matches.length>0,`找不到「${key}」的文中引用`,`請在內文提及「${key}」，並將引用文字加粗。`,null,key,true);
 }
 for(const p of textParas.filter(p=>!captions.includes(p))){const re=/[圖表]\s*[一二三四五六七八九十百\d]+/g;for(const m of p.text.matchAll(re)){
  let offset=0;const covered=p.runs.filter(r=>{const start=offset;offset+=r.text.length;return offset>m.index&&start<m.index+m[0].length});
  record('figures',p,'引用加粗',covered.every(r=>r.bold),`「${m[0]}」引用文字未完整加粗`,'請選取該圖表引用文字，套用粗體（Ctrl+B）。');
  record('figures',p,'引用對應圖表',capKeys.has(normalize(m[0])),`「${m[0]}」未找到對應圖說或表題`,'請新增對應的圖說／表題，或修正引用編號。',null,null,true);
 }}
 // Embedded raster dimensions and placement use OOXML display size, not metadata-only DPI.
 const rel$=xml(await zip.file('word/_rels/document.xml.rels')?.async('string')||'<Relationships/>');const rels=new Map(rel$('*').filter((_,n)=>local(n)==='Relationship').toArray().map(n=>[attr(n,'Id'),{target:attr(n,'Target'),external:attr(n,'TargetMode')==='External'}]));
 const blocks=$(body).children().toArray();const blockOf=p=>{let n=p.node;while(n?.parent&&n.parent!==body)n=n.parent;return n;};
 function neighbour(block,direction){const index=blocks.indexOf(block);for(let j=index+direction;j>=0&&j<blocks.length;j+=direction){const pp=paragraphs.find(p=>blockOf(p)===blocks[j]&&p.text.trim());if(pp)return pp;if(local(blocks[j])==='tbl')return null;}return null;}
 let images=0;
 for(const p of paragraphs){for(const drawing of p.drawings){const blips=descendants($,drawing,'blip');for(const b of blips){images++;const rel=rels.get(attr(b,'embed')||attr(b,'link'));const target=rel&&!rel.external?packageTarget(rel.target||''):'';const file=zip.file(target);const ext=(target.split('/').pop().split('.').slice(1).pop()||'').toLowerCase();
   record('figures',p,'圖片格式',file?['jpg','jpeg','png','tif','tiff'].includes(ext):null,file?'圖片格式不符合 JPG／PNG／TIFF 規範':'外部圖片無法讀取','請嵌入原始 JPG、PNG 或 TIFF 圖片，不要使用低解析度 BMP。',ext||'外部或未知格式','JPG／PNG／TIFF');
   const extent=descendants($,drawing,'extent')[0];const cx=Number(attr(extent,'cx')),cy=Number(attr(extent,'cy'));const width=cx/360000;
   record('figures',p,'圖片寬度',cx?width>=8&&width<=12:null,'圖片寬度不在建議 8–12 cm 範圍','請依 A4 版面檢視圖片寬度，建議設為 8–12 cm；特殊版型可人工確認後保留。',cx?width.toFixed(2)+' cm':'無法解析','建議 8–12 cm',true);
   let dims;try{if(file)dims=imageSize(await file.async('uint8array'));}catch{}
   const dpi=dims&&cx&&cy?Math.min(dims.width/(cx/914400),dims.height/(cy/914400)):null;
   record('figures',p,'圖片有效解析度',dpi!==null?dpi>=300:null,dpi!==null?'圖片依置入尺寸計算低於 300 dpi':'無法取得圖片像素或顯示尺寸','請以高解析度原圖替換，或縮小置入尺寸；一般彩色圖片至少 300 dpi，線條圖至少 600 dpi。',dpi!==null?Math.round(dpi)+' dpi':'未知','彩色 ≥300 dpi');
   record('figures',p,'線條圖解析度',null,dpi!==null?`圖片類型須確認（有效解析度約 ${Math.round(dpi)} dpi）`:'圖片類型與解析度須人工確認',dpi!==null&&dpi>=600?'若為線條圖，解析度已達 600 dpi；請確認圖像類型及清晰度。':'若為線條圖，需至少 600 dpi；請人工確認類型並依需求替換原圖。');
   const next=neighbour(blockOf(p),1);const sameCaption=/^圖\s*[一二三四五六七八九十百\d]+/.test(p.text);const inline=descendants($,drawing,'inline').length>0;
   record('figures',p,'圖說位置',inline&&!sameCaption?!!next&&/^圖\s*[一二三四五六七八九十百\d]+/.test(next.text):null,inline?'圖片下方未找到圖說':'浮動圖片或同行圖說需人工確認位置','請將圖說置於圖片下方，使用「圖一 圖說」格式；浮動圖片需以 Word 實際版面確認。',next?.text?.slice(0,50), '圖說在圖片下方',true);
  }}if(descendants($,p.node,'pict').length)record('figures',p,'舊式圖形',null,'發現舊式圖形或文字方塊','此類圖形無法可靠計算解析度及版面位置，請在 Word 人工確認。');}
 const tables=blocks.filter(b=>local(b)==='tbl');for(const [i,t]of tables.entries()){const prev=neighbour(t,-1);record('figures',prev,`第 ${i+1} 個表格標題`,!!prev&&/^表\s*[一二三四五六七八九十百\d]+/.test(prev.text),'表格上方未找到表題','請在表格上方新增「表一 表格標題」，並依序編號。',prev?.text?.slice(0,60),'表題在表格上方',true);}
 if(descendants($,body,'ins').length||descendants($,body,'del').length)record('references',null,'追蹤修訂',null,'文件含尚未接受的追蹤修訂','本次以保留新增、排除刪除的文字檢查；請在 Word 確認修訂後再次上傳。');
 const textboxes=paragraphs.filter(p=>p.inTextbox);if(textboxes.length)record('references',null,'文字方塊',null,'文件包含文字方塊，閱讀順序及版面需人工確認','請在 Word 檢閱文字方塊（例如作者簡介）的內容、字體與位置；這些段落不列入一般內文排版判定。');
 const counts={passed:0,failed:0,warning:0,manual:0};for(const c of checks)counts[c.status]++;
 const disabled=Object.keys(articleDefaults).filter(k=>enabled[k]===false);
 const auto=counts.passed+counts.failed;
 return {issues,checks,paragraphs:paragraphs.filter(p=>p.text.trim()).map(({number,text,style,level,inTable,inTextbox})=>({number,text,style,level,inTable,inTextbox})),summary:{...counts,total:checks.length,automated:auto,score:auto?Math.round(counts.passed/auto*100):null,paragraphs:paragraphs.filter(p=>p.text.trim()).length,images,tables:tables.length,disabled,complete:disabled.length===0&&counts.manual===0&&counts.warning===0&&counts.failed===0},policy};
}
