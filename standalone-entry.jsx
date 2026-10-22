import {installBrowserStore} from './browser-store.js';
installBrowserStore().then(()=>import('./src.jsx')).catch(e=>{
 document.getElementById('root').textContent='無法啟動單檔版：'+e.message+'。請使用最新版 Edge 或 Chrome 開啟。';
});
