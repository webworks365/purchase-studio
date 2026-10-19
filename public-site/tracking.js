import {splitQuantityDue} from './due-date.js';
import {dueTarget} from './inline-due.js';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const parse=s=>{const x=new DOMParser().parseFromString(s,'application/xml');if(x.querySelector('parsererror'))throw Error('Excel XML 格式不正確');return x;};
const all=(n,name)=>Array.from(n.getElementsByTagNameNS(S,name));
const col=n=>{let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;};
const colNum=s=>Array.from(s).reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
const clean=s=>String(s||'').replace(/\s/g,'');
const b64=bytes=>{let s='';for(let i=0;i<bytes.length;i+=16384)s+=String.fromCharCode(...bytes.subarray(i,i+16384));return btoa(s);};
export const companies=[{id:'sp',name:'興普',header:6,start:7,block:1,last:'AU',stages:Array.from({length:36},(_,i)=>col(i+11))},{id:'yj',name:'優技',header:1,start:2,block:1,last:'AF',stages:Array.from({length:21},(_,i)=>col(i+11))},{id:'cx',name:'采鑫',header:1,start:2,block:2,last:'K',stages:[]}];
export function initTracking(state){state.tracking??={templates:{},entries:[]};state.tracking.templates??={};state.tracking.entries??=[];return state.tracking;}
export function remaining(entry){return entry.quantity!==''&&entry.shipped!==''&&entry.quantity!=null&&entry.shipped!=null?Number(entry.quantity)-Number(entry.shipped):null;}
export function orderInfo(order){
 const template=order.templateSnapshot,value=key=>order.values[template.fields.find(f=>f.key===key)?.id]||'',quantityText=splitQuantityDue(value('quantity')).quantity,main=quantityText.split(/備品|\+/)[0],matches=[...main.matchAll(/(\d+(?:\.\d+)?)\s*(PCS|SET|PNL|PANEL)\b/gi)],target=dueTarget(template),raw=splitQuantityDue(order.values[target?.id]||'').rawDue;
 const unambiguous=matches.length===1&&[...quantityText.matchAll(/(?:^|\n)\s*\d+(?:\.\d+)?\s*(?:PCS|SET|PNL|PANEL)\b/gi)].length<=1;
 return {part:value('part').trim(),externalPart:value('customerPart').trim(),orderDate:value('date'),due:raw.replace(/^交期\s*[:：\->]*\s*/i,''),quantityText,quantity:unambiguous?matches[0][1]:'',unit:unambiguous?matches[0][2].toUpperCase():'',spare:quantityText.match(/備品\s*([^\n]+)/)?.[1]||''};
}
export function newTrackingEntry(order,company){return {id:crypto.randomUUID(),orderId:order.id,orderNo:order.orderNo,supplierName:order.supplierName,company,batch:'1',status:'open',customerPo:'',updatedDue:'',salesDue:'',shipped:'',rest:'',rework:'',currentStage:'',notes:'',stage1:'',stageDate1:'',stage2:'',stageDate2:'',stages:{},...orderInfo(order),createdAt:new Date().toISOString(),updatedAt:new Date().toISOString()};}
export function validateEntry(entry,entries=[]){
 if(!companies.some(c=>c.id===entry.company))throw Error('請選擇進度表公司');if(!entry.batch?.trim())throw Error('請填寫批次');
 for(const [key,label]of [['quantity','訂購數量'],['shipped','已出貨數量'],['rest','尾數']])if(entry[key]!==''&&entry[key]!=null&&(!Number.isFinite(Number(entry[key]))||Number(entry[key])<0))throw Error(label+'請填寫非負數字');
 if(entries.some(e=>e.id!==entry.id&&e.orderId===entry.orderId&&e.company===entry.company&&e.batch===entry.batch))throw Error('這筆訂單與批次已加入此表，請編輯原追蹤單或使用其他批次');
}
async function workbook(zip){
 const x=parse(await zip.file('xl/workbook.xml').async('string')),rels=parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
 return all(x,'sheet').map(s=>{const target=Array.from(rels.documentElement.children).find(r=>r.getAttribute('Id')===s.getAttributeNS(R,'id'))?.getAttribute('Target');if(!target)throw Error('工作表連結缺失');return {name:s.getAttribute('name'),path:target.startsWith('/')?target.slice(1):'xl/'+target};});
}
function valueOf(cell,strings){if(!cell)return '';return cell.getAttribute('t')==='s'?strings[Number(all(cell,'v')[0]?.textContent)]||'':cell.getAttribute('t')==='inlineStr'?all(cell,'t').map(n=>n.textContent).join(''):all(cell,'v')[0]?.textContent||'';}
function erase(cell){for(const child of Array.from(cell.children))if(['v','is','f'].includes(child.localName))child.remove();cell.removeAttribute('t');}
function write(cell,value){erase(cell);if(value===''||value==null)return;const doc=cell.ownerDocument;if(typeof value==='number'&&Number.isFinite(value)){const v=doc.createElementNS(S,'v');v.textContent=String(value);cell.appendChild(v);}else{cell.setAttribute('t','inlineStr');const inline=doc.createElementNS(S,'is'),t=doc.createElementNS(S,'t');t.setAttribute('xml:space','preserve');t.textContent=String(value);inline.appendChild(t);cell.appendChild(inline);}}
export async function importTrackingTemplate(file){
 if(!/\.xlsx$/i.test(file.name)||file.size>15*1024*1024)throw Error('請選擇 15 MB 以內的 XLSX 檔案');const zip=await JSZip.loadAsync(await file.arrayBuffer());if(!zip.file('xl/workbook.xml'))throw Error('不是有效的 Excel 活頁簿');
 const strings=zip.file('xl/sharedStrings.xml')?all(parse(await zip.file('xl/sharedStrings.xml').async('string')),'si').map(s=>all(s,'t').map(t=>t.textContent).join('')):[],sheets=await workbook(zip);let config,main,stageLabels={};
 for(const sheet of sheets){const xml=parse(await zip.file(sheet.path).async('string')),cells=all(xml,'c'),at=addr=>valueOf(cells.find(c=>c.getAttribute('r')===addr),strings);
  const found=clean(at('A6')).includes('SPP/N')&&clean(at('C6')).includes('採購單號')?companies[0]:clean(at('B1'))==='優技料號'?companies[1]:clean(at('B1'))==='采鑫料號'?companies[2]:null;
  if(found){if(main)throw Error('檔案包含多張可用進度表，請分開匯入');config=found;main=sheet;for(const column of found.stages)stageLabels[column]=at(column+found.header).replace(/\s+/g,' ').trim();}
 }
 if(!main)throw Error('尚不支援這個版型，請匯入提供的興普、優技或采鑫進度表');
 // Retain formatting only. Never carry source orders or hidden shared strings.
 for(const sheet of sheets){const xml=parse(await zip.file(sheet.path).async('string'));for(const cell of all(xml,'c')){const row=Number(cell.getAttribute('r').match(/\d+/)[0]);if(row>(sheet.path===main.path?config.header:1))erase(cell);else if(cell.getAttribute('t')==='s')write(cell,valueOf(cell,strings));}zip.file(sheet.path,new XMLSerializer().serializeToString(xml));}
 zip.remove('xl/sharedStrings.xml');const rels=parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));for(const r of Array.from(rels.documentElement.children))if(/\/sharedStrings$/.test(r.getAttribute('Type')))r.remove();zip.file('xl/_rels/workbook.xml.rels',new XMLSerializer().serializeToString(rels));
 const types=parse(await zip.file('[Content_Types].xml').async('string'));for(const r of Array.from(types.documentElement.children))if(r.getAttribute('PartName')==='/xl/sharedStrings.xml')r.remove();zip.file('[Content_Types].xml',new XMLSerializer().serializeToString(types));
 return {company:config.id,name:config.name,filename:file.name,sheetName:main.name,sheetPath:main.path,stageLabels,data:await zip.generateAsync({type:'base64',compression:'DEFLATE'}),importedAt:new Date().toISOString()};
}
function rowValues(entry,config,index){
 const number=value=>value===''||value==null?'':Number(value),r=remaining(entry),stages=Object.fromEntries(config.stages.map(c=>[c,entry.stages[c]??'']));
 if(config.id==='sp')return {A:entry.externalPart,B:entry.part,C:entry.customerPo,D:number(entry.quantity),E:number(entry.rest),F:entry.rework,G:entry.updatedDue,H:entry.orderDate,I:entry.due,J:entry.salesDue,...stages,AU:entry.notes};
 if(config.id==='yj')return {A:index+1,B:entry.externalPart,C:entry.part,D:entry.orderDate,E:entry.due,F:entry.updatedDue,G:number(entry.quantity),H:number(entry.shipped),I:r,J:entry.customerPo,...stages,AF:entry.notes};
 return {A:index+1,B:entry.externalPart,C:entry.part,D:entry.orderDate,E:entry.due,F:entry.updatedDue,G:number(entry.quantity),H:entry.customerPo,I:entry.stage1,J:entry.stage2,K:entry.notes};
}
export async function validateTracking(state){
 if(state.tracking==null)return;const tracking=state.tracking;if(!tracking.templates||typeof tracking.templates!=='object'||!Array.isArray(tracking.entries))throw Error('追蹤備份格式不正確');
 for(const [id,t]of Object.entries(tracking.templates)){if(!companies.some(c=>c.id===id)||t.company!==id||typeof t.data!=='string')throw Error('追蹤格式資料不完整');const checked=await importTrackingTemplate(new File([Uint8Array.from(atob(t.data),c=>c.charCodeAt(0))],'template.xlsx'));if(checked.company!==id)throw Error('進度表公司不符');}
 const ids=new Set();for(const entry of tracking.entries){if(typeof entry.id!=='string'||ids.has(entry.id)||!state.orders.some(o=>o.id===entry.orderId)||!['open','closed'].includes(entry.status)||!entry.stages||typeof entry.stages!=='object')throw Error('追蹤單資料不完整');ids.add(entry.id);validateEntry(entry,tracking.entries);}
}
export async function exportTracking(template,entries){
 const config=companies.find(c=>c.id===template.company);if(!config)throw Error('進度表格式不正確');if(!entries.length)throw Error('此公司目前沒有要匯出的追蹤單');
 for(const entry of entries){if(entry.company!==config.id)throw Error('匯出資料包含不同公司');if(!entry.part?.trim())throw Error(entry.orderNo+' 缺少銘星料號');if(entry.quantity===''||entry.unit!=='PCS'||!Number.isFinite(Number(entry.quantity)))throw Error(entry.orderNo+'：原表使用 PCS，請先確認此批訂購數量與單位');}
 const zip=await JSZip.loadAsync(template.data,{base64:true}),xml=parse(await zip.file(template.sheetPath).async('string')),data=all(xml,'sheetData')[0],rows=Array.from(data.children),samples=Array.from({length:config.block},(_,i)=>rows.find(r=>Number(r.getAttribute('r'))===config.start+i));if(samples.some(r=>!r))throw Error('模板缺少資料列格式');
 for(const row of rows)if(Number(row.getAttribute('r'))>=config.start)row.remove();
 const lastRow=config.start+entries.length*config.block-1;
 for(const [index,entry]of entries.entries())for(let block=0;block<config.block;block++){const r=config.start+index*config.block+block,row=samples[block].cloneNode(true);row.setAttribute('r',String(r));for(const cell of Array.from(row.children).filter(n=>n.localName==='c')){cell.setAttribute('r',cell.getAttribute('r').replace(/\d+$/,String(r)));erase(cell);}const values=block===0?rowValues(entry,config,index):{I:entry.stageDate1,J:entry.stageDate2};
  for(const [column,value]of Object.entries(values)){if(!/^[A-Z]+$/.test(column)||colNum(column)>colNum(config.last))continue;let cell=Array.from(row.children).find(c=>c.getAttribute('r')===column+r);if(!cell){cell=xml.createElementNS(S,'c');cell.setAttribute('r',column+r);row.appendChild(cell);}write(cell,value);}
  const cells=Array.from(row.children).filter(c=>c.localName==='c').sort((a,b)=>colNum(a.getAttribute('r').match(/^[A-Z]+/)[0])-colNum(b.getAttribute('r').match(/^[A-Z]+/)[0]));for(const cell of cells)row.appendChild(cell);data.appendChild(row);
 }
 const merges=all(xml,'mergeCells')[0];if(merges){for(const merge of Array.from(merges.children))if(Number(merge.getAttribute('ref').match(/\d+/)[0])>=config.start)merge.remove();}
 if(config.id==='cx'){let group=merges;if(!group){group=xml.createElementNS(S,'mergeCells');data.parentNode.insertBefore(group,data.nextSibling);}for(let i=0;i<entries.length;i++){const row=config.start+i*2;for(const c of ['A','B','C','D','E','F','G','H','K']){const merge=xml.createElementNS(S,'mergeCell');merge.setAttribute('ref',c+row+':'+c+(row+1));group.appendChild(merge);}}group.setAttribute('count',String(group.children.length));}
 else if(merges){merges.setAttribute('count',String(merges.children.length));if(!merges.children.length)merges.remove();}
 all(xml,'dimension')[0]?.setAttribute('ref','A1:'+config.last+lastRow);all(xml,'autoFilter')[0]?.setAttribute('ref','A'+config.header+':'+config.last+lastRow);
 zip.file(template.sheetPath,new XMLSerializer().serializeToString(xml));const wb=parse(await zip.file('xl/workbook.xml').async('string'));for(const name of all(wb,'definedName'))if(name.getAttribute('name')==='_xlnm.Print_Area'&&name.textContent.includes(template.sheetName))name.textContent="'"+template.sheetName.replaceAll("'","''")+"'!$A$1:$"+config.last+'$'+lastRow;zip.file('xl/workbook.xml',new XMLSerializer().serializeToString(wb));
 return zip.generateAsync({type:'blob',compression:'DEFLATE',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
