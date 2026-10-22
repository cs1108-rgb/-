import test from 'node:test';
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {mkdtemp,rm,readFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {createServer} from 'node:net';
import {DatabaseSync} from 'node:sqlite';
import {setTimeout as delay} from 'node:timers/promises';

test('文章 API：舊資料遷移、原稿留存、規則快照及重啟後歷史皆可驗證',async()=>{
 const dir=await mkdtemp(join(tmpdir(),'measure-article-'));
 const db=new DatabaseSync(join(dir,'measurements.sqlite'));
 db.exec('CREATE TABLE reports(id INTEGER PRIMARY KEY,name TEXT,period TEXT,created TEXT,total INTEGER,issues TEXT,rows TEXT)');
 db.prepare('INSERT INTO reports VALUES(1,?,?,?,?,?,?)').run('legacy.csv','2026-01','2026-01-01T00:00:00Z',1,'[]','[]');db.close();
 const portFinder=createServer();await new Promise(r=>portFinder.listen(0,'127.0.0.1',r));const port=portFinder.address().port;await new Promise(r=>portFinder.close(r));const base=`http://127.0.0.1:${port}`;
 let child;
 async function start(){child=spawn(process.execPath,['server.js'],{cwd:process.cwd(),env:{...process.env,PORT:String(port),NODE_ENV:'production',MEASURE_DATA_DIR:dir},stdio:'pipe'});for(let i=0;i<70;i++){try{const r=await fetch(base+'/api/rules');if(r.ok)return;}catch{}await delay(100);}throw Error('Server did not become ready');}
 async function stop(){if(!child||child.exitCode!==null)return;await new Promise(r=>{child.once('exit',r);child.kill()});}
 async function json(url,opts){const r=await fetch(base+url,opts);assert.equal(r.status,200);return r.json();}
 async function upload(bytes,name='文章.docx'){const f=new FormData();f.append('file',new Blob([bytes]),name);f.append('period','2026-10');return fetch(base+'/api/upload',{method:'POST',body:f});}
 try{
  await start();const old=await json('/api/reports');assert.equal(old[0].kind,'measurement');
  const template=await readFile('templates/article-format.docx');const downloadable=Buffer.from(await(await fetch(base+'/api/template/word')).arrayBuffer());assert.deepEqual(downloadable,template);
  const response=await upload(template);assert.equal(response.status,200);const result=await response.json();assert.ok(result.metadata.summary.failed>0);
  const original=Buffer.from(await(await fetch(base+`/api/reports/${result.id}/original`)).arrayBuffer());assert.deepEqual(original,template);
  const first=await json('/api/reports/'+result.id);assert.equal(first.name,'文章.docx');assert.equal(first.rules.typography,true);assert.ok(first.rows.length>0);assert.ok(first.metadata.checks.length>0);
  await json('/api/rules',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({typography:false})});
  const second=await(await upload(template)).json();assert.ok(second.metadata.summary.disabled.includes('typography'));assert.equal((await json('/api/reports/'+result.id)).rules.typography,true);
  assert.equal((await upload(Buffer.from('invalid'))).status,400);assert.equal((await upload(template,'legacy.doc')).status,400);
  await stop();await start();const history=await json('/api/reports');assert.equal(history.length,3);assert.equal((await json('/api/rules')).typography,false);assert.equal((await json('/api/reports/'+result.id)).metadata.policy.version,'article-format-v1');
 }finally{await stop();await rm(dir,{recursive:true,force:true});}
});
