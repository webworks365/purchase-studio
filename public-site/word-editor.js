import {W,parseXml,textOf,exportDocx} from './core.js';
import {quantityWithDue} from './due-date.js';
const escape=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
function separatorPattern(text){return new RegExp(escape(text.trim()).replace(/\s+/g,'\\s*'));}
function cellText(cell){const paragraphs=Array.from(cell.children).filter(n=>n.tagName==='P');return (paragraphs.length?paragraphs.map(p=>p.innerText).join('\n'):cell.innerText).replace(/\u00a0/g,' ').trim();}
function cellValues(fields,original,text){
 if(!fields.some(f=>f.segment))return {[fields[0].id]:text};
 const sorted=[...fields].sort((a,b)=>a.segment.start-b.segment.start),out={};let cursor=0;
 const prefix=original.slice(0,sorted[0].segment.start).trim();if(prefix){const match=separatorPattern(prefix).exec(text);if(!match)throw Error('請保留同一格的欄位名稱');cursor=match.index+match[0].length;}
 for(let i=0;i<sorted.length;i++){const f=sorted[i],next=sorted[i+1],suffix=next?original.slice(f.segment.end,next.segment.start):original.slice(f.segment.end);if(suffix.trim()){const match=separatorPattern(suffix).exec(text.slice(cursor));if(!match)throw Error('請保留同一格的欄位名稱');out[f.id]=text.slice(cursor,cursor+match.index).trim();cursor+=match.index+match[0].length;}else out[f.id]=text.slice(cursor).trim();}
 return out;
}
export function createWordEditor(host,template,{getValues,getOrderNo,onChange,onError}){
 const shadow=host.attachShadow({mode:'open'}),style=document.createElement('style'),body=document.createElement('div'),styles=document.createElement('div');
 style.textContent=':host{display:block;min-width:0}.worddoc-wrapper{background:#c7cdd5!important;padding:20px!important;overflow:auto;align-items:center!important}section.worddoc{flex-shrink:0;box-shadow:0 2px 8px #17273833}td[contenteditable=true]{outline:0;cursor:text}td[contenteditable=true]:hover{box-shadow:inset 0 0 0 1px #6585a5}td[contenteditable=true]:focus{box-shadow:inset 0 0 0 2px #315d88}td[contenteditable=false]{cursor:default}a{pointer-events:none}';
 shadow.append(style,styles,body);let revision=0,disposed=false,timer;
 const originalCells=JSZip.loadAsync(template.data,{base64:true}).then(async zip=>Array.from(parseXml(await zip.file('word/document.xml').async('string')).getElementsByTagNameNS(W,'tc')).map(textOf));
 async function refresh(){
  const current=++revision;clearTimeout(timer);const values=structuredClone(getValues()),orderNo=getOrderNo();host.dataset.ready='false';host.setAttribute('aria-busy','true');
  try{
   const originals=await originalCells,blob=await exportDocx(template,values,orderNo),zip=await JSZip.loadAsync(await blob.arrayBuffer()),xml=parseXml(await zip.file('word/document.xml').async('string')),cells=Array.from(xml.getElementsByTagNameNS(W,'tc'));
   // Anchors identify the actual original cells without changing the exported file.
   cells.forEach((cell,index)=>{const p=Array.from(cell.children).find(n=>n.localName==='p');if(!p)return;const mark=xml.createElementNS(W,'w:bookmarkStart');mark.setAttributeNS(W,'w:id',String(100000+index));mark.setAttributeNS(W,'w:name','PurchaseCell'+index);p.insertBefore(mark,p.firstChild);});zip.file('word/document.xml',new XMLSerializer().serializeToString(xml));
   const renderBody=document.createElement('div'),renderStyles=document.createElement('div');
   await docx.renderAsync(await zip.generateAsync({type:'arraybuffer'}),renderBody,renderStyles,{className:'worddoc',inWrapper:true,ignoreWidth:false,ignoreHeight:false,ignoreFonts:false,breakPages:true,ignoreLastRenderedPageBreak:false,renderHeaders:true,renderFooters:true,renderFootnotes:true,renderEndnotes:true,renderAltChunks:false,useBase64URL:true,experimental:true});
   if(disposed||current!==revision||!host.isConnected)return;
   body.replaceChildren(...renderBody.childNodes);styles.replaceChildren(...renderStyles.childNodes);
   for(const index of new Set(template.fields.map(f=>f.cell))){const cell=body.querySelector('#PurchaseCell'+index)?.closest('td');if(!cell)continue;const fields=template.fields.filter(f=>f.cell===index),editable=fields.some(f=>!['date','delivery'].includes(f.key));cell.contentEditable=String(editable);cell.spellcheck=false;cell.dataset.cell=String(index);cell.setAttribute('aria-label',fields.map(f=>f.label).join('／'));cell.setAttribute('role','textbox');cell.setAttribute('aria-multiline','true');const initial=cellText(cell);let valid=true;
    cell.addEventListener('input',()=>{try{const changes=cellValues(fields,originals[index],cellText(cell));for(const f of fields){if(['date','delivery'].includes(f.key)){delete changes[f.id];continue;}if(f.key==='quantity')changes[f.id]=quantityWithDue(changes[f.id],getValues().__dueDate,getValues().__dueRaw);}valid=true;onChange(changes);}catch{valid=false;}});
    cell.addEventListener('paste',event=>{event.preventDefault();const text=event.clipboardData.getData('text/plain');document.execCommand('insertText',false,text);});
    cell.addEventListener('blur',()=>{if(!valid)onError('請保留同格的欄位名稱，內容已還原');if(cellText(cell)!==initial)refresh();});
   }
   host.dataset.ready='true';host.setAttribute('aria-busy','false');
  }catch(error){if(current!==revision||disposed)return;host.setAttribute('aria-busy','false');onError('Word 版面載入失敗：'+error.message);}
 }
 refresh();return {refresh(){clearTimeout(timer);timer=setTimeout(refresh,180);},dispose(){disposed=true;revision++;clearTimeout(timer);}};
}
