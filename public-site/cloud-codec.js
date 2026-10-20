export const CHUNK_SIZE=500000,MAX_CHUNKS=280;
export const validKey=value=>/^[A-Za-z0-9_-]{43}$/.test(value||'');
export function newKey(){return btoa(String.fromCharCode(...crypto.getRandomValues(new Uint8Array(32)))).replaceAll('+','-').replaceAll('/','_').replace(/=+$/,'');}
export function workspaceIdentity(location,storage){
 const params=new URLSearchParams(location.hash.slice(1)),explicit=params.has('workspace'),saved=storage.getItem('purchase-cloud-key');
 const key=explicit?params.get('workspace'):validKey(saved)?saved:newKey();
 if(!validKey(key))throw Error('專屬連結不完整，請使用原本完整的連結');
 storage.setItem('purchase-cloud-key',key);return {key,explicit};
}
export function validateWorkspace(w){
 if(!w||!Array.isArray(w.suppliers)||!Array.isArray(w.orders)||!w.drafts||typeof w.drafts!=='object'||!w.company||typeof w.company!=='object')throw Error('雲端工作區格式不正確');
 return w;
}
export const digest=async text=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(text))),n=>n.toString(16).padStart(2,'0')).join('');
export async function packWorkspace(workspace){
 validateWorkspace(workspace);const text=JSON.stringify(workspace);if(new TextEncoder().encode(text).length>100*1024*1024)throw Error('工作區超過 100 MB，請先整理舊需求附件');
 const zip=new JSZip();zip.file('workspace.json',text);const data=await zip.generateAsync({type:'base64',compression:'DEFLATE'}),chunks=[];
 for(let i=0;i<data.length;i+=CHUNK_SIZE)chunks.push(data.slice(i,i+CHUNK_SIZE));
 if(chunks.length>MAX_CHUNKS)throw Error('工作區附件過大');return {chunks,checksum:await digest(data)};
}
export function validHead(head){return head&&head.format===1&&Number.isSafeInteger(head.version)&&head.version>0&&/^[a-f0-9-]{36}$/.test(head.revision)&&Number.isInteger(head.chunks)&&head.chunks>0&&head.chunks<=MAX_CHUNKS&&/^[a-f0-9]{64}$/.test(head.checksum);}
export async function unpackWorkspace(head,chunks){
 if(!validHead(head)||chunks.length!==head.chunks||chunks.some(c=>typeof c!=='string'||c.length>CHUNK_SIZE))throw Error('雲端資料尚未完整，請稍後重試');
 const data=chunks.join('');if(await digest(data)!==head.checksum)throw Error('雲端資料校驗失敗，已保留本機內容');
 const zip=await JSZip.loadAsync(data,{base64:true}),entry=zip.file('workspace.json');if(!entry)throw Error('雲端資料檔案不完整');
 const text=await entry.async('string');if(text.length>100*1024*1024)throw Error('雲端工作區過大');return validateWorkspace(JSON.parse(text));
}
