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
 let p=Array.from(cell.children).find(n=>n.localName==='p');if(!p){p=cell.ownerDocument.createElementNS(W,'w:p');cell.appendChild(p);}
 const pr=p.getElementsByTagNameNS(W,'pPr')[0]?.cloneNode(true),rp=p.getElementsByTagNameNS(W,'rPr')[0]?.cloneNode(true);
 for(const n of Array.from(cell.children))if(n.localName!=='tcPr')cell.removeChild(n);
 const doc=cell.ownerDocument;for(const line of String(value).split('\n')){const np=doc.createElementNS(W,'w:p');if(pr)np.appendChild(pr.cloneNode(true));const r=doc.createElementNS(W,'w:r');if(rp)r.appendChild(rp.cloneNode(true));const t=doc.createElementNS(W,'w:t');t.setAttribute('xml:space','preserve');t.textContent=line;r.appendChild(t);np.appendChild(r);cell.appendChild(np);}
}
export async function exportDocx(template,values,orderNo){
 const zip=await JSZip.loadAsync(template.data,{base64:true});const x=parseXml(await zip.file('word/document.xml').async('string'));
 const cells=Array.from(x.getElementsByTagNameNS(W,'tc'));
 for(const f of template.fields){if(cells[f.cell]&&values[f.id]!==undefined&&textOf(cells[f.cell])!==values[f.id])setCellText(cells[f.cell],values[f.id]);}
 // Order number may span Word runs. Rebuild only its paragraph when changed.
 for(const p of Array.from(x.getElementsByTagNameNS(W,'p'))){const ts=Array.from(p.getElementsByTagNameNS(W,'t')),s=ts.map(t=>t.textContent).join('');if(/訂單號碼\s*[:：]/.test(s)&&ts.length){ts[0].textContent=s.replace(/(訂單號碼\s*[:：]\s*)[^\s\n]+/,(_,label)=>label+orderNo);for(const t of ts.slice(1))t.textContent='';}}
 if(Array.isArray(values.__requirementAppendix)&&values.__requirementAppendix.length){
  const body=x.getElementsByTagNameNS(W,'body')[0],end=Array.from(body.children).find(n=>n.localName==='sectPr');
  const addP=(text,{pageBreak=false,bold=false}={})=>{const p=x.createElementNS(W,'w:p'),r=x.createElementNS(W,'w:r');if(bold){const pr=x.createElementNS(W,'w:rPr');pr.appendChild(x.createElementNS(W,'w:b'));r.appendChild(pr);}if(pageBreak){const br=x.createElementNS(W,'w:br');br.setAttributeNS(W,'w:type','page');r.appendChild(br);}const t=x.createElementNS(W,'w:t');t.textContent=text;r.appendChild(t);p.appendChild(r);body.insertBefore(p,end||null);};
  addP('客戶製作要求（訂單附件）',{pageBreak:true,bold:true});addP('對應訂單：'+orderNo);
  for(const row of values.__requirementAppendix){addP(row.label,{bold:true});for(const line of row.value.split('\n'))addP(line);addP('確認來源：'+row.source);}
 }
 zip.file('word/document.xml',new XMLSerializer().serializeToString(x));
 // Remove comments, tracked revision authors and document personal properties from generated copy.
 if(zip.file('docProps/core.xml')){const meta=parseXml(await zip.file('docProps/core.xml').async('string'));for(const n of Array.from(meta.getElementsByTagName('*')))if(['creator','lastModifiedBy'].includes(n.localName))n.textContent='';zip.file('docProps/core.xml',new XMLSerializer().serializeToString(meta));}
 return zip.generateAsync({type:'blob',compression:'DEFLATE',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}
export function extractLoose(text){
 const out={},patterns={date:/(?:訂購時間|訂購日期|日期)\s*[:：]?\s*(\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2})/,customerPart:/(?:客戶料號)\s*[:：]?\s*([^\s\n]+)/,part:/(?:^|\n)\s*(?:銘星料號|本公司料號|品號|產品料號|料號)\s*[:：]?\s*([^\s\n]+)/,quantity:/(?:訂購數量|數量)\s*[:：]?\s*([^\n]+)/,material:/(?:材料板厚|材料規格)\s*[:：]?\s*([^\n]+)/,copper:/銅厚\s*[:：]?\s*([^\n]+)/,mask:/防焊\s*[:：]?\s*([^\n]+)/,finish:/表面處理\s*[:：]?\s*([^\n]+)/,production:/製作內容\s*[:：]?\s*([^\n]+)/};
 for(const [k,re]of Object.entries(patterns)){const v=text.match(re)?.[1];if(v)out[k]=v.trim();}return out;
}
export const safeFile=s=>String(s).replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').slice(0,100)||'採購單';

