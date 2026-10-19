import {normalizeDueDate,splitQuantityDue} from './due-date.js';
export const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const labels=[['item',/^項目/],['date',/^訂購時間|^訂購日期/],['customerPart',/^客戶料號/],['part',/料號$/],['quantity',/^訂購數量|^數量/],['delivery',/^交貨地點/],['payment',/^付款方式/],['supplier',/^廠商名稱|^供應商/],['supplierAddress',/^廠商地址/],['contact',/^聯絡人/],['phone',/^電話/],['email',/^EM$|^EMAIL$/i],['production',/^製作內容/],['material',/^材料板厚|^材料規格/],['copper',/^銅厚/],['mask',/^防焊/],['finish',/^表面處理/],['inspection',/^檢驗標準/],['ul',/^UL\s*MARK/i],['notes',/^備註/],['maker',/^製表者/]];
export const transactional=new Set(['date','customerPart','part','quantity','price']);
export const importable=new Set(['item','date','customerPart','part','quantity','price','production','material','copper','mask','finish','inspection','ul']);
export const textOf=n=>Array.from(n.getElementsByTagNameNS(W,'p')).map(p=>Array.from(p.getElementsByTagNameNS(W,'t')).map(t=>t.textContent).join('')).join('\n').trim();
export function parseXml(xml){const x=new DOMParser().parseFromString(xml,'application/xml');if(x.getElementsByTagName('parsererror').length)throw Error('Word XML 格式無法解析');return x;}
export function parseDocument(xml){
 const x=parseXml(xml),cells=Array.from(x.getElementsByTagNameNS(W,'tc')),rows=[],fields=[];
 const add=(cell,key,label)=>{const index=cells.indexOf(cell);let id=key;let n=2;while(fields.some(f=>f.id===id))id=key+(n++);const f={id,key,label,cell:index,value:textOf(cell)};fields.push(f);return f;};
 for(const row of Array.from(x.getElementsByTagNameNS(W,'tr'))){
  const cs=Array.from(row.children).filter(n=>n.localName==='tc'),out=[];
  for(let i=0;i<cs.length;i++){
   const value=textOf(cs[i]);const clean=value.replace(/[:：\s]/g,'');const match=labels.find(([,re])=>re.test(clean));
   if(match&&i+1<cs.length){const f=add(cs[++i],match[0],value.replace(/[:：]\s*$/,'').trim());out.push({label:f.label,id:f.id});}
   else if(value&&/^(?:單價|總|Usd|USD|NT|Nt|US\$|\$)/.test(value)){const f=add(cs[i],'price','單價／金額與費用');out.push({label:f.label,id:f.id});}
   else if(cs.length===1&&value){const signature=value.length<60&&/核\s*准|廠商回簽/.test(value);const f=add(cs[i],signature?'signature':'notesTail',signature?'核准／廠商回簽':'補充條款／包裝與附件');out.push({label:f.label,id:f.id});}
   else if(value){out.push({text:value});}
  }
  if(out.length)rows.push(out);
 }
 const full=textOf(x),orderNo=full.match(/訂單號碼\s*[:：]\s*([^\s\n]+)/)?.[1]||'';
 return {fields,rows,orderNo,text:full,name:fields.find(f=>f.key==='supplier')?.value||''};
}
export async function parseDocx(data){const zip=await JSZip.loadAsync(data);const e=zip.file('word/document.xml');if(!e)throw Error('檔案不是有效的 DOCX，請使用 Word .docx 格式');const xml=await e.async('string');return {...parseDocument(xml),xml};}
function setCellText(cell,value){
 // Keep the original paragraphs, runs, style properties, tabs and pictures.
 // Only text nodes change; additional text lines use breaks within an existing run.
 const doc=cell.ownerDocument,paragraphs=Array.from(cell.children).filter(n=>n.localName==='p'),lines=String(value).split('\n');
 if(!paragraphs.length)return;
 const nearestP=n=>{let p=n.parentNode;while(p&&!(p.namespaceURI===W&&p.localName==='p'))p=p.parentNode;return p;};
 for(let i=0;i<paragraphs.length;i++){
  const p=paragraphs[i],text=i===paragraphs.length-1?lines.slice(i).join('\n'):(lines[i]||'');
  let nodes=Array.from(p.getElementsByTagNameNS(W,'t')).filter(t=>nearestP(t)===p);
  if(!nodes.length){if(!text)continue;let r=Array.from(p.children).find(n=>n.localName==='r');if(!r){r=doc.createElementNS(W,'w:r');p.appendChild(r);}const t=doc.createElementNS(W,'w:t');r.appendChild(t);nodes=[t];}
  let remaining=Array.from(text);
  for(let j=0;j<nodes.length;j++){
   const t=nodes[j],length=j===nodes.length-1?remaining.length:Array.from(t.textContent).length,chunk=remaining.splice(0,length).join(''),pieces=chunk.split('\n');
   t.textContent=pieces[0];t.setAttribute('xml:space','preserve');let cursor=t;
   for(const piece of pieces.slice(1)){const br=doc.createElementNS(W,'w:br'),next=doc.createElementNS(W,'w:t');next.setAttribute('xml:space','preserve');next.textContent=piece;cursor.parentNode.insertBefore(br,cursor.nextSibling);br.parentNode.insertBefore(next,br.nextSibling);cursor=next;}
  }
 }
}
function originalLayoutValues(template,values){
 const out={...values},quantity=template.fields.filter(f=>f.key==='quantity');
 const raw=values.__dueRaw||quantity.map(f=>splitQuantityDue(values[f.id]).rawDue).find(Boolean)||'';
 const date=Object.hasOwn(values,'__dueDate')?values.__dueDate:normalizeDueDate(raw),due=date?'交期：'+date.replaceAll('-','/'):(raw?(/^交期/.test(raw)?raw:'交期：'+raw):'');
 for(const f of quantity)if(values[f.id]!==undefined)out[f.id]=splitQuantityDue(values[f.id]).quantity;
 for(const f of template.fields.filter(f=>f.key==='delivery'))out[f.id]='銘星科技有限公司';
 const target=template.fields.find(f=>['quantity','delivery'].includes(f.key)&&/交期/.test(f.value))||quantity[0];
 if(target&&due){const base=out[target.id]??target.value,separator=target.value.match(/([\s]+)交期/)?.[1]||'\n';out[target.id]=/^\s*交期/.test(target.value)?due+'\n'+base:base+separator+due;}
 return out;
}
export async function exportDocx(template,values,orderNo){
 const zip=await JSZip.loadAsync(template.data,{base64:true});const x=parseXml(await zip.file('word/document.xml').async('string'));
 const cells=Array.from(x.getElementsByTagNameNS(W,'tc'));
 const filled=originalLayoutValues(template,values);
 for(const f of template.fields){if(cells[f.cell]&&filled[f.id]!==undefined&&textOf(cells[f.cell])!==filled[f.id])setCellText(cells[f.cell],filled[f.id]);}
 // Order number may span Word runs. Rebuild only its paragraph when changed.
 for(const p of Array.from(x.getElementsByTagNameNS(W,'p'))){const ts=Array.from(p.getElementsByTagNameNS(W,'t')),s=ts.map(t=>t.textContent).join('');if(/訂單號碼\s*[:：]/.test(s)&&ts.length){ts[0].textContent=s.replace(/(訂單號碼\s*[:：]\s*)[^\s\n]+/,(_,label)=>label+orderNo);for(const t of ts.slice(1))t.textContent='';}}
 zip.file('word/document.xml',new XMLSerializer().serializeToString(x));
 return zip.generateAsync({type:'blob',compression:'DEFLATE',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}
export function extractLoose(text){
 const out={},patterns={date:/(?:訂購時間|訂購日期|日期)\s*[:：]?\s*(\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2})/,customerPart:/(?:客戶料號)\s*[:：]?\s*([^\s\n]+)/,part:/(?:^|\n)\s*(?:銘星料號|本公司料號|品號|產品料號|料號)\s*[:：]?\s*([^\s\n]+)/,quantity:/(?:訂購數量|數量)\s*[:：]?\s*([^\n]+)/,material:/(?:材料板厚|材料規格)\s*[:：]?\s*([^\n]+)/,copper:/銅厚\s*[:：]?\s*([^\n]+)/,mask:/防焊\s*[:：]?\s*([^\n]+)/,finish:/表面處理\s*[:：]?\s*([^\n]+)/,production:/製作內容\s*[:：]?\s*([^\n]+)/};
 for(const [k,re]of Object.entries(patterns)){const v=text.match(re)?.[1];if(v)out[k]=v.trim();}return out;
}
export const safeFile=s=>String(s).replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').slice(0,100)||'採購單';

