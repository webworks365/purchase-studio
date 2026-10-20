import {workspaceIdentity,packWorkspace,unpackWorkspace,validHead} from './cloud-codec.js';
import {firebaseStore} from './firebase-store.js';
export function cloudMessage(error){
 const code=String(error?.code||''),message=String(error?.message||'');
 if(code.includes('operation-not-allowed')||message.includes('ADMIN_ONLY_OPERATION'))return 'Firebase 尚未啟用匿名驗證';
 if(code.includes('permission-denied'))return 'Firebase 存取規則尚未設定，請發布專案提供的規則';
 if(code.includes('failed-precondition')||code.includes('not-found'))return '請先在 Firebase 建立 Firestore 資料庫';
 if(code==='cloud/conflict')return message;
 if(/工作區超過|附件過大|校驗失敗|格式不正確/.test(message))return message;
 return '雲端暫時無法連線；修改已留在這台裝置，可稍後重試';
}
export class CloudSync {
 constructor({identity,onStatus,onRemote,makeStore=firebaseStore}){Object.assign(this,{identity,onStatus,onRemote,makeStore});this.version=0;this.pending=null;this.running=false;this.conflict=false;this.ready=false;this.loading=false;this.inFlight=null;}
 status(type,text){this.onStatus({type,text});}
 async connect(local,getLocal=async()=>({workspace:local,pending:false,version:0})){
  const connection=(this.connection||0)+1;this.connection=connection;
  const active=()=>{if(this.connection!==connection)throw Object.assign(Error('連線已取消'),{code:'cloud/cancelled'});};
  this.status('working','正在連接雲端…');
  try{
   const store=await this.makeStore(this.identity.key);active();this.store=store;const result=await this.pull();active();
   if(!result&&this.identity.explicit)throw Object.assign(Error('這個專屬連結找不到工作區，請確認連結完整'),{code:'cloud/missing'});
   const latest=await getLocal();active();this.ready=true;this.conflict=false;
   let workspace=result||latest.workspace||local;
   if(latest.pending){workspace=latest.workspace;if(result&&this.version!==latest.version){this.version=latest.version;this.conflict=true;}this.pending=workspace;}
   else if(!result)this.pending=workspace;
   this.unsubscribe?.();this.unsubscribe=this.store.listen(h=>this.changed(h),e=>this.status('error',cloudMessage(e)));
   this.status(this.conflict?'conflict':this.pending?'working':'ok',this.conflict?'雲端有較新版本；本機修改已保留':this.pending?'正在移轉此裝置資料…':'雲端已連線');
   return workspace;
  }catch(error){if(this.connection===connection){this.ready=false;this.status('error',error.code==='cloud/missing'?error.message:cloudMessage(error));}throw error;}
 }
 async pull(){
  for(let attempt=0;attempt<3;attempt++){
   const h=await this.store.head();if(!h)return null;if(!validHead(h))throw Error('雲端工作區格式不正確');
   try{const workspace=await unpackWorkspace(h,await this.store.readChunks(h));this.version=h.version;return workspace;}catch(error){if(attempt===2)throw error;}
  }
 }
 queue(workspace){this.pending=structuredClone(workspace);this.status(this.conflict?'conflict':this.ready?'working':'error',this.conflict?'雲端有較新版本；本機修改已保留':this.ready?'正在儲存雲端…':'修改已保留在本機；雲端尚未連線，請重試同步');return this.flush();}
 async flush(){
  if(this.running||this.loading||this.conflict||!this.ready||!this.pending)return;
  this.running=true;
  try{while(this.pending&&!this.conflict){
   const current=this.pending;this.pending=null;const revision=crypto.randomUUID();let uploaded;
   try{
    const packed=await packWorkspace(current);uploaded={format:1,version:this.version+1,revision,chunks:packed.chunks.length,checksum:packed.checksum};
    this.inFlight=revision;await this.store.upload(revision,packed.chunks);const previous=await this.store.commit(this.version,uploaded);this.version=uploaded.version;this.inFlight=null;
    if(previous)this.store.discard(previous).catch(()=>{});
   }catch(error){this.pending??=current;this.inFlight=null;if(uploaded)this.store.discard(uploaded).catch(()=>{});if(error.code==='cloud/conflict'){this.conflict=true;this.status('conflict',cloudMessage(error));}else this.status('error',cloudMessage(error));return;}
  }if(!this.conflict)this.status('ok','已儲存雲端');}
  finally{this.running=false;}
 }
 async changed(h){
  if(!h||h.version<=this.version||h.revision===this.inFlight||this.loading||this.conflict)return;
  if(this.pending||this.running){this.conflict=true;this.status('conflict','另一台裝置有更新；本機修改已保留');return;}
  this.loading=true;const before=this.version;
  try{const workspace=await this.pull();if(this.pending||this.running||!await this.onRemote(workspace)){this.version=before;this.conflict=true;this.status('conflict','另一台裝置有更新；請先載入雲端版本');}else this.status('ok','已載入雲端更新');}
  catch(error){this.version=before;this.status('error',cloudMessage(error));}finally{this.loading=false;}
 }
 async reload(){
  if(this.running)throw Error('正在同步，請稍候');this.loading=true;
  try{const workspace=await this.pull();if(!workspace)throw Error('找不到雲端工作區');if(!await this.onRemote(workspace,true))throw Error('目前操作尚未完成');this.pending=null;this.conflict=false;this.ready=true;this.status('ok','已載入雲端版本');}finally{this.loading=false;}
 }
}
export {workspaceIdentity};
