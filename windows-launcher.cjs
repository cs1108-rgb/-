const fs=require('node:fs');
const path=require('node:path');
const net=require('node:net');
const {spawn,spawnSync}=require('node:child_process');
const {setTimeout:delay}=require('node:timers/promises');
const root=__dirname;
process.chdir(root);
function fail(message){console.error('\n'+message);process.exitCode=1;}
function npm(command){
 // All command strings are constants; no user input is passed to cmd.exe.
 const result=process.platform==='win32'
  ?spawnSync(process.env.ComSpec||'cmd.exe',['/d','/s','/c',`npm.cmd ${command}`],{cwd:root,stdio:'inherit'})
  :spawnSync('npm',command.split(' '),{cwd:root,stdio:'inherit'});
 if(result.error||result.status!==0)throw Error(`npm ${command} failed. Check your network and the error above, then run SETUP-WINDOWS.cmd again.`);
}
async function main(){
 if(Number(process.versions.node.split('.')[0])<24)throw Error('Node.js 24 LTS or newer is required. Install Node.js from https://nodejs.org and reopen this window.');
 const setupOnly=process.argv.includes('--setup');
 if(setupOnly||!fs.existsSync(path.join(root,'node_modules','express'))||!fs.existsSync(path.join(root,'dist','index.html'))){
  console.log('Preparing MeasureInformationOS. First setup requires Internet access. Your data folder will be preserved.');
  npm('ci --no-audit --no-fund');npm('run build');
  console.log('\nSetup complete. Use START-WINDOWS.cmd to open MeasureInformationOS.');
 }
 if(setupOnly)return;
 const port=Number(process.env.PORT||3000);
 if(!Number.isInteger(port)||port<1024||port>65535)throw Error('PORT must be a number between 1024 and 65535.');
 await new Promise((resolve,reject)=>{const probe=net.createServer();probe.once('error',()=>reject(Error(`Port ${port} is already in use. Close the other MeasureInformationOS window or choose another PORT.`)));probe.listen(port,'127.0.0.1',()=>probe.close(resolve));});
 const url=`http://127.0.0.1:${port}`;
 const child=spawn(process.execPath,['start-production.js'],{cwd:root,stdio:'inherit',env:{...process.env,MEASURE_HOST:'127.0.0.1',PORT:String(port)}});
 let childError;
 child.on('error',e=>{childError=e;});
 const stop=()=>{if(child.exitCode===null)child.kill();};
 process.once('exit',stop);
 process.once('SIGINT',()=>{stop();process.exit(0)});
 process.once('SIGTERM',()=>{stop();process.exit(0)});
 let ready=false;
 for(let i=0;i<100;i++){
  if(childError)throw childError;
  if(child.exitCode!==null)throw Error('The local service stopped. Read the error above.');
  try{const r=await fetch(url+'/api/policy',{signal:AbortSignal.timeout(500)});if(r.ok&&(await r.json()).version?.startsWith('article-format-')){ready=true;break}}catch{}
  await delay(100);
 }
 if(!ready){stop();throw Error('The service did not become ready. Check the error above.');}
 console.log(`\nMeasureInformationOS is ready: ${url}\nOnly this computer can connect. Keep this window open.\nUse Ctrl+C to stop. Word files and reports remain in the data folder.\n`);
 if(process.platform==='win32'&&!process.argv.includes('--no-browser')){
  const browser=spawn(process.env.ComSpec||'cmd.exe',['/d','/s','/c',`start "" "${url}"`],{stdio:'ignore'});
  browser.on('error',()=>console.log(`Open your browser manually: ${url}`));
 }
 await new Promise((resolve,reject)=>child.once('exit',(code)=>code===0||code===null?resolve():reject(Error(`The service stopped (exit ${code}).`))));
}
main().catch(e=>fail(e.message));
