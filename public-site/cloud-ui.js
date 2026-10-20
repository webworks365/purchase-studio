import {CloudSync,workspaceIdentity} from './cloud-sync.js';
export function installCloud({readLocal,writeLocal,getState,replaceState,isEditing,toast}){
 const identity=workspaceIdentity(location,localStorage),cacheKey='cloud:'+identity.key;
 let starting=false,activated=false,recovery=false,lastType='working',lastText='正在準備雲端…',persistChain=Promise.resolve(),timeout;
 const bar=document.createElement('div');bar.className='cloud-bar';bar.setAttribute('role','status');document.querySelector('main').prepend(bar);
 const persist=(workspace,pending)=>{const snapshot={workspace:structuredClone(workspace),pending,version:sync.version};persistChain=persistChain.catch(()=>{}).then(()=>writeLocal(cacheKey,snapshot));return persistChain;};
 function draw(){
  bar.className='cloud-bar '+lastType;bar.replaceChildren();
  const text=document.createElement('span');text.textContent=lastText;bar.append(text);
  const actions=document.createElement('div');actions.className='actions';bar.append(actions);
  function button(text,action){const b=document.createElement('button');b.className='small';b.textContent=text;b.onclick=async()=>{b.disabled=true;try{await action();}catch(error){toast(error.message);}finally{b.disabled=false;}};actions.append(b);}
  if(sync.ready)button('複製專屬連結',async()=>{await navigator.clipboard.writeText(privateUrl());toast('已複製，換電腦請使用這個連結；請勿公開分享');});
  if(lastType==='conflict')button('載入雲端版本',async()=>{if(!confirm('將載入雲端版本。本機目前資料會保留在此裝置的復原備份。'))return;await writeLocal('cloud-recovery:'+identity.key,{workspace:structuredClone(getState()),savedAt:new Date().toISOString()});recovery=true;await sync.reload();});
  if(recovery)button('下載復原備份',async()=>{const record=await readLocal('cloud-recovery:'+identity.key);if(!record)throw Error('找不到復原備份');const blob=new Blob([JSON.stringify({format:'purchase-studio-backup-v1',workspace:record.workspace,exportedAt:record.savedAt})],{type:'application/json'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='採購工作室_本機復原備份.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),10000);});
  if(lastType==='error')button('重試同步',retry);
 }
 function privateUrl(){const url=new URL(location.href);url.hash='workspace='+identity.key;return url.href;}
 const sync=new CloudSync({identity,onStatus:({type,text})=>{lastType=type;lastText=text;draw();if(type==='ok'&&activated)persist(getState(),!!sync.pending).catch(()=>toast('此裝置快取失敗，請匯出備份'));},onRemote:async(workspace,force=false)=>{
  if(!force&&isEditing())return false;await replaceState(workspace);await persist(workspace,false);return true;
 }});
 async function start(){
  if(starting)return;starting=true;
  try{
   const connecting=sync.connect(getState(),async()=>{await persistChain.catch(()=>{});return await readLocal(cacheKey)||{workspace:getState(),pending:false,version:0};});
   const workspace=await Promise.race([connecting,new Promise((_,reject)=>{timeout=setTimeout(()=>{sync.connection++;sync.ready=false;sync.status('error','雲端連線逾時；本機資料已保留，可稍後重試');reject(Error('連線逾時'));},20000);})]);clearTimeout(timeout);
   if(!sync.conflict&&isEditing()){sync.conflict=true;sync.pending=structuredClone(getState());sync.status('conflict','本機內容正在編輯，請稍後載入雲端版本');}
   else if(workspace!==getState())await replaceState(workspace);
   activated=true;history.replaceState(null,'',privateUrl());await persist(getState(),!!sync.pending);await sync.flush();
  }catch{/* A persistent status explains the error. Local work remains available. */}
  finally{clearTimeout(timeout);starting=false;}
 }
 async function retry(){if(!sync.ready)await start();else await sync.flush();}
 window.addEventListener('online',retry);
 draw();
 readLocal('cloud-recovery:'+identity.key).then(v=>{recovery=!!v;draw();}).catch(()=>{});
 return {
  start,
  async read(){const cached=await readLocal(cacheKey);if(cached){sync.version=cached.version||0;return cached.workspace;}return identity.explicit?undefined:readLocal('workspace');},
  async save(workspace,{initial=false}={}){const pending=!initial||!!(await readLocal(cacheKey))?.pending;await persist(workspace,pending);if(!initial)sync.queue(workspace);},
  get unsynced(){return !!sync.pending||sync.running;},
  get cacheKey(){return cacheKey;}
 };
}
