import {analyzeArticle,articleDefaults,ruleDefinitions,policy} from './article.js';

const DB_NAME='measureos-article-offline-v1';
let db=null;
const memory={reports:new Map(),rules:{...articleDefaults},next:1};
const originalURLs=new Map();
const uid=()=>globalThis.crypto?.randomUUID?.()||`${Date.now()}-${Math.random().toString(36).slice(2)}`;
const normalizeRules=r=>Object.fromEntries(Object.keys(articleDefaults).map(k=>[k,r?.[k]!==false]));
const errorText='瀏覽器儲存失敗，可能空間不足或被禁止。本次資料尚未保存；請先備份歷史紀錄、釋放空間後重試。';
function request(req){return new Promise((resolve,reject)=>{req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(Error(errorText));});}
function done(tx){return new Promise((resolve,reject)=>{tx.oncomplete=resolve;tx.onerror=tx.onabort=()=>reject(Error(errorText));});}
async function getAll(){if(!db)return [...memory.reports.values()];return request(db.transaction('reports').objectStore('reports').getAll());}
async function getOne(id){if(!db)return memory.reports.get(id);return request(db.transaction('reports').objectStore('reports').get(id));}
async function rules(){if(!db)return memory.rules;const r=await request(db.transaction('settings').objectStore('settings').get('rules'));return normalizeRules(r?.value);}
async function setRules(value){value=normalizeRules(value);if(!db){memory.rules=value;return value;}const tx=db.transaction('settings','readwrite');const finished=done(tx);tx.objectStore('settings').put({key:'rules',value});await finished;return value;}
async function addReport(report){if(!db){report.id=memory.next++;memory.reports.set(report.id,report);return report.id;}const tx=db.transaction('reports','readwrite');const finished=done(tx);const [id]=await Promise.all([request(tx.objectStore('reports').add(report)),finished]);return id;}
function summary(r){const{original,rows,...rest}=r;return rest;}
function fromBase64(s){const raw=atob(s);const b=new Uint8Array(raw.length);for(let i=0;i<raw.length;i++)b[i]=raw.charCodeAt(i);return b;}
async function toBase64(blob){const bytes=new Uint8Array(await blob.arrayBuffer());let result='';for(let i=0;i<bytes.length;i+=32768)result+=String.fromCharCode(...bytes.subarray(i,i+32768));return btoa(result);}
export function saveBlob(blob,name){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),3000);}
export async function exportBackup(){
 const reports=await getAll();const entries=[];
 for(const r of reports){const{original,...rest}=r;entries.push({...rest,originalBase64:original?await toBase64(original):null});}
 saveBlob(new Blob([JSON.stringify({format:'measureos-backup-v1',created:new Date().toISOString(),rules:await rules(),reports:entries})],{type:'application/json'}),`量測資訊歷史備份-${new Date().toLocaleDateString('sv-SE')}.json`);
 return entries.length;
}
function validateReport(r){
 const fail=()=>{throw Error('備份內容格式不正確，未匯入任何資料。請選擇此單檔版匯出的完整備份。');};
 if(!r||typeof r.uid!=='string'||r.uid.length>200||r.kind!=='article'||typeof r.name!=='string'||r.name.length>500||!/^\d{4}-(0[1-9]|1[0-2])$/.test(r.period)||!Number.isFinite(Date.parse(r.created)))fail();
 if(!Number.isInteger(r.total)||r.total<0||r.total>10000||!Array.isArray(r.issues)||!Array.isArray(r.rows)||!r.metadata?.summary||!r.metadata?.policy||!Array.isArray(r.metadata.checks))fail();
 for(const k of ['source','version','sha256'])if(typeof r.metadata.policy[k]!=='string')fail();
 const s=r.metadata.summary;for(const k of ['passed','failed','warning','manual','total','automated','paragraphs','images','tables'])if(!Number.isInteger(s[k])||s[k]<0)fail();
 if(s.score!==null&&(!Number.isFinite(s.score)||s.score<0||s.score>100)||!Array.isArray(s.disabled)||s.disabled.some(k=>!Object.hasOwn(articleDefaults,k)))fail();
 if(r.issues.length>100000||r.rows.length>10000||r.metadata.checks.length>100000)fail();
 for(const i of r.issues){for(const k of ['field','message','suggestion','location','excerpt'])if(typeof i[k]!=='string')fail();if(!['error','warning'].includes(i.level))fail();for(const k of ['actual','expected'])if(i[k]!==null&&i[k]!==undefined&&typeof i[k]!=='string'&&typeof i[k]!=='number')fail();}
 for(const p of r.rows)if(!Number.isInteger(p.number)||typeof p.text!=='string')fail();
 if(r.originalBase64!==null&&(typeof r.originalBase64!=='string'||r.originalBase64.length>14*1024*1024||!/^[A-Za-z0-9+/]*={0,2}$/.test(r.originalBase64)))fail();
 const original=r.originalBase64?new Blob([fromBase64(r.originalBase64)],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}):null;
 return {uid:r.uid,kind:'article',name:r.name,period:r.period,created:r.created,total:r.total,issues:r.issues,rows:r.rows,metadata:r.metadata,rules:normalizeRules(r.rules),original};
}
export async function importBackup(file){
 if(file.size>150*1024*1024)throw Error('備份檔案不可超過 150 MB。');
 let data;try{data=JSON.parse(await file.text())}catch{throw Error('無法讀取備份 JSON，請確認檔案沒有損毀。');}
 if(data?.format!=='measureos-backup-v1'||!Array.isArray(data.reports)||data.reports.length>1000)throw Error('請選擇此單檔版匯出的完整歷史備份，不是單份報告。');
 // Validate everything before mutating history. The import is additive and deduplicates stable IDs.
 const validated=data.reports.map(validateReport);const existing=new Set((await getAll()).map(r=>r.uid));const incoming=[];
 for(const r of validated)if(!existing.has(r.uid)){incoming.push(r);existing.add(r.uid);}
 if(db){const tx=db.transaction('reports','readwrite');const finished=done(tx);for(const r of incoming)tx.objectStore('reports').add(r);await finished;}else for(const r of incoming)await addReport(r);
 return {added:incoming.length,skipped:validated.length-incoming.length};
}
export async function localApi(url,options={}){
 const method=(options.method||'GET').toUpperCase();
 if(url==='/api/policy')return {...policy,rules:ruleDefinitions};
 if(url==='/api/rules')return method==='PUT'?setRules(JSON.parse(options.body)):rules();
 if(url==='/api/reports')return (await getAll()).sort((a,b)=>b.id-a.id).map(summary);
 const match=url.match(/^\/api\/reports\/(\d+)$/);if(match){const r=await getOne(Number(match[1]));if(!r)throw Error('找不到此紀錄');if(originalURLs.has(r.id))URL.revokeObjectURL(originalURLs.get(r.id));const originalURL=r.original?URL.createObjectURL(r.original):null;if(originalURL)originalURLs.set(r.id,originalURL);return {...summary(r),rows:r.rows,hasOriginal:!!r.original,originalURL};}
 if(url==='/api/upload'&&method==='POST'){
  const file=options.body.get('file'),period=options.body.get('period');
  if(!(file instanceof Blob)||! /\.docx$/i.test(file.name))throw Error('請選擇 Word（.docx）文章。');
  if(file.size>10*1024*1024)throw Error('檔案不可超過 10 MB');
  if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(period||''))throw Error('請選擇有效的刊期月份');
  const snapshot=await rules();const report=await analyzeArticle(await file.arrayBuffer(),snapshot);
  const r={uid:uid(),kind:'article',name:file.name,period,created:new Date().toISOString(),total:report.summary.paragraphs,issues:report.issues,rows:report.paragraphs,metadata:{summary:report.summary,policy:report.policy,checks:report.checks},rules:snapshot,original:file};
  const id=await addReport(r);return {id,total:r.total,issues:r.issues,metadata:r.metadata};
 }
 throw Error('單檔版不支援這項操作。');
}
export async function installBrowserStore(){
 try{db=await new Promise((resolve,reject)=>{const req=indexedDB.open(DB_NAME,1);const timer=setTimeout(()=>reject(Error('讀取瀏覽器資料庫逾時')),5000);req.onupgradeneeded=()=>{const d=req.result;d.createObjectStore('reports',{keyPath:'id',autoIncrement:true});d.createObjectStore('settings',{keyPath:'key'});};req.onsuccess=()=>{clearTimeout(timer);resolve(req.result)};req.onerror=req.onblocked=()=>{clearTimeout(timer);reject(req.error)}});}
 catch{db=null;}
 globalThis.measureosStorageMode=db?'browser':'memory';
 globalThis.measureosLocalAPI=localApi;
 globalThis.measureosGetOriginal=async id=>{const r=await getOne(Number(id));if(!r?.original)throw Error('此紀錄沒有原始 Word，請重新上傳文章後再產生格式修訂副本。');return r.original;};
 globalThis.measureosExportBackup=exportBackup;
 globalThis.measureosImportBackup=importBackup;
 const templateBlob=new Blob([fromBase64(globalThis.MEASUREOS_TEMPLATE_BASE64)],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
 globalThis.measureosGetTemplate=async()=>templateBlob;
 globalThis.measureosTemplateURL=URL.createObjectURL(templateBlob);
}
