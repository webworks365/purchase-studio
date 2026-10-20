import {firebaseConfig} from './firebase-config.js';
import {validHead} from './cloud-codec.js';
// No analytics: the URL fragment is a bearer credential, never telemetry.
const SDK='https://www.gstatic.com/firebasejs/13.0.0/';
export async function firebaseStore(key){
 const [appSDK,authSDK,f]=await Promise.all([import(SDK+'firebase-app.js'),import(SDK+'firebase-auth.js'),import(SDK+'firebase-firestore.js')]);
 const app=appSDK.getApps().length?appSDK.getApp():appSDK.initializeApp(firebaseConfig),auth=authSDK.getAuth(app);
 await auth.authStateReady();if(!auth.currentUser)await authSDK.signInAnonymously(auth);
 const db=f.getFirestore(app),root=f.doc(db,'purchaseWorkspaces',key);
 const chunk=(revision,index)=>f.doc(root,'versions',revision,'chunks',String(index).padStart(4,'0'));
 async function head(){const s=await f.getDocFromServer(root);return s.exists()?s.data():null;}
 return {
  head,
  async readChunks(h){return Promise.all(Array.from({length:h.chunks},async(_,i)=>{const s=await f.getDocFromServer(chunk(h.revision,i));if(!s.exists())throw Error('雲端資料版本已更新，正在重新讀取');return s.data().data;}));},
  async upload(revision,chunks){for(let start=0;start<chunks.length;start+=20){const batch=f.writeBatch(db);chunks.slice(start,start+20).forEach((data,i)=>batch.set(chunk(revision,start+i),{data}));await batch.commit();}},
  async commit(expected,h){return f.runTransaction(db,async tx=>{const snapshot=await tx.get(root),current=snapshot.exists()?snapshot.data():null;if((current?.version||0)!==expected){const e=Error('另一台裝置已更新資料，請先載入雲端版本');e.code='cloud/conflict';throw e;}tx.set(root,{...h,updatedAt:f.serverTimestamp()});return current;});},
  async discard(h){if(!h||!validHead(h))return;for(let start=0;start<h.chunks;start+=20){const batch=f.writeBatch(db);for(let i=start;i<Math.min(start+20,h.chunks);i++)batch.delete(chunk(h.revision,i));await batch.commit();}},
  listen(next,error){return f.onSnapshot(root,{includeMetadataChanges:true},s=>{if(!s.metadata.fromCache&&!s.metadata.hasPendingWrites)next(s.exists()?s.data():null);},error);}
 };
}
