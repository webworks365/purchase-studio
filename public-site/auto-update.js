export function createUpdateMonitor({currentVersion,fetchVersion,prepare,refresh}){
 let pending=null,checking=false,refreshing=false;
 const valid=value=>typeof value==='string'&&/^[a-f0-9]{7,64}$/.test(value);
 return {async check(){
  if(checking||refreshing)return;checking=true;
  try{const version=pending||await fetchVersion();if(!valid(version)||version===currentVersion)return;pending=version;
   if(!await prepare())return;refreshing=true;refresh(version);
  }catch{/* Keep the current page and local drafts if network or storage is unavailable. */}
  finally{checking=false;}
 }};
}
if(typeof document!=='undefined'){
 const currentVersion=document.querySelector('meta[name="app-build"]')?.content;
 if(/^[a-f0-9]{7,64}$/.test(currentVersion||'')){
  const monitor=createUpdateMonitor({currentVersion,
   fetchVersion:async()=>{const url=new URL('./version.json',import.meta.url);url.searchParams.set('_',Date.now());const r=await fetch(url,{cache:'no-store'});return r.ok?(await r.json()).version:null;},
   prepare:async()=>typeof window.prepareForSiteUpdate==='function'&&await window.prepareForSiteUpdate(),
   refresh:version=>{window.__siteUpdating=true;const url=new URL(location.href);url.searchParams.set('_release',version);location.replace(url.href);}
  });
  setInterval(()=>monitor.check(),30000);
  document.addEventListener('visibilitychange',()=>{if(!document.hidden)monitor.check();});
  window.addEventListener('online',()=>monitor.check());window.addEventListener('focus',()=>monitor.check());
 }
}
