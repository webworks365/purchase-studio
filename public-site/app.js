import {installCloud} from './cloud-ui.js';
import {renderTracking} from './tracking-ui.js';
import {initTracking,validateTracking} from './tracking.js';
import {dueTarget,transferDue} from './inline-due.js';
import {createWordEditor} from './word-editor.js';
import {imageFile,combineImages,isAllowedImage} from './requirement-images.js';
import {upgradeWorkspace,replaceSupplierTemplate} from './template-update.js';
import {DELIVERY,creationDate,supplierDisplay,formRows,supplementalRows,normalizeDraft,intakeValues,intakeReport} from './order-layout.js';
import {readRequirementDocx,readRequirementXlsx,readRequirementText,requirementMapping} from './requirements.js';
import {normalizeDueDate,splitQuantityDue,quantityWithDue} from './due-date.js';
import {orderPrefix,formatOrderNumber,validOrderNumber} from './order-number.js';
import {parseDocx,exportDocx,prepareInlineDue,extractLoose,transactional,importable,safeFile} from './core.js';
const $=s=>document.querySelector(s),esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const uid=()=>crypto.randomUUID(),today=()=>new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Taipei'}).format(new Date());
let db,state,page='create',selected='',draft=null,busy=false,scan=null,dirty=false,activeReq=null;
let cloud=null,autosaveTimer,editCounter=0;
function scheduleAutoSave(){editCounter++;clearTimeout(autosaveTimer);autosaveTimer=setTimeout(async()=>{if(!dirty||!draft)return;if(busy){scheduleAutoSave();return;}const current=draft,revision=editCounter;state.drafts[selected]=structuredClone(draft);try{await save();if(current===draft&&revision===editCounter)dirty=false;}catch(error){toast('自動儲存失敗：'+error.message);}},900);}
let wordEditor=null;const app=$('#app'),modal=$('#modal');
function toast(s){$('#toast').textContent=s;$('#toast').classList.add('show');clearTimeout(toast.timer);toast.timer=setTimeout(()=>$('#toast').classList.remove('show'),5000);}
async function openDB(){return new Promise((res,rej)=>{const r=indexedDB.open('purchase-studio-v1',1);r.onupgradeneeded=()=>r.result.createObjectStore('data');r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error);});}
async function readLocal(key){return new Promise((res,rej)=>{const r=db.transaction('data').objectStore('data').get(key);r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error);});}
async function writeLocal(key,value){return new Promise((res,rej)=>{const t=db.transaction('data','readwrite');t.objectStore('data').put(value,key);t.oncomplete=()=>res();t.onerror=()=>rej(t.error);t.onabort=()=>rej(t.error||Error('本機儲存失敗'));});}
async function save(options){if(cloud)return cloud.save(state,options);return writeLocal('workspace',state);}
function cloudState(){if(draft&&dirty){const result=structuredClone(state);result.drafts[selected]=structuredClone(draft);return result;}return state;}
async function replaceCloudState(workspace){wordEditor?.dispose();wordEditor=null;clearTimeout(autosaveTimer);state=workspace;state.requirements??=[];initTracking(state);await upgradeWorkspace(state);activate(state.suppliers.find(s=>s.id===state.lastSupplierId)?.id||state.suppliers[0]?.id||'');render();}
function download(blob,name){const a=document.createElement('a'),url=URL.createObjectURL(blob);a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),10000);}
function bytes64(data){let s='';const a=new Uint8Array(data);for(let i=0;i<a.length;i+=16384)s+=String.fromCharCode(...a.subarray(i,i+16384));return btoa(s);}
function template(){return state.suppliers.find(s=>s.id===selected);}
function newDraft(s){const values={};for(const f of s.fields)values[f.id]=f.value;const result={id:uid(),supplierId:s.id,orderSuffix:'',orderNo:orderPrefix(),dueDate:today(),dueRaw:'',defaultsVersion:3,values,createdAt:new Date().toISOString()};normalizeDraft(s,result);result.values.__dueDate=result.dueDate;result.values.__dueRaw='';result.values=prepareInlineDue(s,result.values);return result;}
function refreshOrderNumber(){if(!draft)return;draft.orderSuffix??=draft.orderNo?.match(/^MH\d{2}A-\d{2}-(\d{3})$/)?.[1]||'';const previous=draft.orderNo;draft.orderNo=formatOrderNumber(draft.orderSuffix);if(previous!==draft.orderNo)wordEditor?.refresh();const prefix=$('#order-prefix');if(prefix)prefix.textContent=orderPrefix();const display=$('#order-full');if(display)display.textContent=draft.orderNo;}

function activate(id){selected=id;state.lastSupplierId=id;const s=template();draft=s?structuredClone(state.drafts[id]||newDraft(s)):null;scan=null;dirty=false;activeReq=draft?.requirementId||null;if(draft&&s)for(const f of s.fields.filter(f=>['part','price'].includes(f.key)))if(!draft.values[f.id]?.trim())draft.values[f.id]=f.value;if(draft&&s)normalizeDraft(s,draft);initializeDueDate();refreshOrderNumber();}
function switchPage(p){wordEditor?.dispose();wordEditor=null;page=p;for(const b of document.querySelectorAll('nav button'))b.classList.toggle('active',b.dataset.page===p);$('#breadcrumb').textContent=({create:'建立採購單',suppliers:'供應商與模板',history:'進度追蹤',settings:'開單資料與備份'})[p];render();}
function render(){if(page==='create')renderCreate();if(page==='suppliers')renderSuppliers();if(page==='history')renderHistory();if(page==='settings')renderSettings();$('#supplier-count').textContent=state.suppliers.length;}
function title(eyebrow,h,sub,action=''){return `<div class="title-line"><div><div class="eyebrow">${eyebrow}</div><h1>${h}</h1><p class="sub">${sub}</p></div>${action}</div>`;}
function card(n,title,body,tag=''){return `<div class="card"><div class="card-head"><h2><span class="number">${n}</span>${title}</h2>${tag?`<span class="pill">${tag}</span>`:''}</div><div class="card-body">${body}</div></div>`;}
function renderCreate(){
 wordEditor?.dispose();wordEditor=null;const s=template();refreshOrderNumber();
 app.innerHTML=title('','採購單','','<div class="actions order-actions"><button id="new-order">＋ 新增</button><button id="save-draft" '+(s?'':'disabled')+'>存草稿</button><button id="export-word" class="primary" '+(s?'':'disabled')+'>下載 Word</button></div>')+
 `<div class="order-workspace"><div class="card"><div class="card-body">
 ${!s?'<div class="message onboarding">先新增供應商模板，或到「設定與備份」還原資料。</div>':''}
 <div class="order-setup"><div class="entry-top"><label class="field">供應商<select id="supplier-select" aria-label="供應商">${state.suppliers.map(t=>'<option value="'+t.id+'" '+(t.id===selected?'selected':'')+'>'+esc(supplierDisplay(t))+'</option>').join('')}</select></label><button class="text small" id="manage-templates">模板設定</button></div><label class="field order-field">訂單號碼<div class="order-number"><span id="order-prefix">${orderPrefix()}</span><input id="order-suffix" inputmode="numeric" maxlength="3" aria-label="自訂三位數編號" placeholder="035" value="${esc(draft?.orderSuffix||'')}" ${s?'':'disabled'}></div></label></div>
 <label class="drop" id="drop-zone"><input id="source-file" type="file" aria-label="上傳客戶需求" accept=".docx,.xlsx,.pdf,.png,.jpg,.jpeg,.webp,.txt" ${s?'':'disabled'}><strong>＋ 上傳需求</strong><small>Word、Excel、PDF、圖片</small></label><div id="scan-status" class="scan-info" role="status">${scan?esc(scan.message):''}</div>
 </div></div>
 ${s?'<div id="word-original-editor" aria-label="原稿採購單"></div>':''}
 <div id="order-attachments"></div>
 </div>`;
 bindCreate();renderAttachmentFields();if(s){const current=draft;wordEditor=createWordEditor($('#word-original-editor'),s,{getValues:()=>current.values,getOrderNo:()=>current.orderNo,onChange:changes=>{if(draft!==current)return;Object.assign(current.values,changes);dirty=true;scheduleAutoSave();},onError:toast});}
}
function renderAttachmentFields(){const el=$('#order-attachments');if(el)el.innerHTML='';}

function bindCreate(){
 $('#source-file').onchange=e=>processFiles(e.target.files);
 const drop=$('#drop-zone');drop.ondragover=e=>{e.preventDefault();drop.classList.add('drag');};drop.ondragleave=()=>drop.classList.remove('drag');drop.ondrop=e=>{e.preventDefault();drop.classList.remove('drag');processFiles(e.dataTransfer.files);};
 $('#supplier-select').onchange=async e=>{if(busy){e.target.value=selected;return;}const id=e.target.value;if(id===selected)return;const sourceDue=splitQuantityDue(draft?.values[dueTarget(template())?.id]||'').rawDue;const source=scan,r=state.requirements?.find(r=>r.id===activeReq),suffix=draft?.orderSuffix||'',due=draft?.dueDate||'',dueRaw=draft?.dueRaw||'';if(draft&&dirty){state.drafts[selected]=structuredClone(draft);try{await save();}catch(err){toast('儲存失敗：'+err.message);e.target.value=selected;return;}}activate(id);scan=source;if(r){applySimpleRequirement(r,suffix);transferDue(template(),draft.values,sourceDue);}renderCreate();};
 $('#manage-templates').onclick=()=>switchPage('suppliers');
 $('#new-order').onclick=()=>{if(busy)return;if(dirty&&!confirm('目前修改尚未儲存，確定建立新單？'))return;const s=template();if(s){draft=newDraft(s);activeReq=null;scan=null;dirty=false;renderCreate();}};
 for(const input of document.querySelectorAll('[data-field]'))input.oninput=()=>{const f=template().fields.find(f=>f.id===input.dataset.field);draft.values[input.dataset.field]=f?.key==='quantity'?quantityWithDue(input.value,draft.dueDate,draft.dueRaw):input.value;dirty=true;scheduleAutoSave();expandOrderText();};

 $('#order-suffix')?.addEventListener('input',e=>{if(!draft)return;e.target.value=e.target.value.replace(/[^0-9]/g,'').slice(0,3);draft.orderSuffix=e.target.value;dirty=true;scheduleAutoSave();refreshOrderNumber();});
 $('#order-suffix')?.addEventListener('blur',e=>{if(!draft)return;if(e.target.value)e.target.value=e.target.value.padStart(3,'0');draft.orderSuffix=e.target.value;refreshOrderNumber();});

 $('#save-draft').onclick=async()=>{if(!draft)return;refreshOrderNumber();state.drafts[selected]=structuredClone(draft);try{await save();dirty=false;toast('草稿已儲存');}catch(e){toast('儲存失敗：'+e.message);}};
 $('#export-word').onclick=()=>generate();
}
function applySimpleRequirement(r,suffix=draft?.orderSuffix||''){
 const s=template();if(!s)return;const previousDraft=draft,previous=draft?.values||{},createdAt=draft?.createdAt||new Date().toISOString();
 draft=newDraft(s);if(previousDraft)draft.id=previousDraft.id;draft.createdAt=createdAt;draft.orderSuffix=suffix;Object.assign(draft.values,intakeValues(r,s,previous));draft.values.__images=combineImages(draft.values.__images||[],r.images||[]);draft.originalLayoutVersion=1;
 normalizeDraft(s,draft);initializeDueDate();draft.requirementId=r.id;activeReq=r.id;refreshOrderNumber();dirty=true;scheduleAutoSave();
}

function applyExtracted(extracted){if(!draft)return;let n=0;for(const f of template().fields)if(importable.has(f.key)&&extracted[f.key]){draft.values[f.id]=extracted[f.key];n++;}draft.reviewed=false;dirty=true;scheduleAutoSave();return n;}
function status(s){const el=$('#scan-status');if(el)el.textContent=s;}
async function processFiles(files){
 if(busy)return;
 for(const file of Array.from(files||[]))await processFile(file);const input=$('#source-file');if(input)input.value='';
}
async function processFile(file){
 if(!file||busy)return;if(!template()){toast('請先新增供應商模板');return;}if(file.size>20*1024*1024){toast('檔案超過 20 MB，請分割後上傳');return;}
 busy=true;document.body.classList.add('busy');status('正在分析 '+file.name+'…');
 try{const ext=file.name.split('.').pop().toLowerCase();let r,text='';const bytes=await file.arrayBuffer();
  if(ext==='docx')r=await readRequirementDocx(bytes,file.name);
  else if(ext==='xlsx')r=await readRequirementXlsx(bytes,file.name);
  else if(ext==='txt')r=readRequirementText(await file.text(),file.name);
  else if(['jpg','jpeg','png','webp'].includes(ext)){r=readRequirementText('',file.name);const image=await imageFile(file);r.images=isAllowedImage(image)?[image]:[];if(!r.images.length){status('已略過 Logo／簽名圖片');return;}}
  else if(ext==='pdf'){
   const pdfjs=await import('./vendor/pdf.min.mjs');pdfjs.GlobalWorkerOptions.workerSrc='./vendor/pdf.worker.min.mjs';const pdf=await pdfjs.getDocument({data:bytes.slice(0)}).promise;
   if(pdf.numPages>15){await pdf.destroy();throw Error('PDF 超過 15 頁，請先擷取與需求有關的頁面。');}
   try{for(let i=1;i<=pdf.numPages;i++){status('讀取 PDF 第 '+i+' / '+pdf.numPages+' 頁…');const p=await pdf.getPage(i),tc=await p.getTextContent();let pt=tc.items.map(t=>t.str+(t.hasEOL?'\n':' ')).join('');if(pt.trim().length<20)pt='';text+=pt+'\n';}}finally{await pdf.destroy();}const images=r?.images||[];r=readRequirementText(text,file.name);r.images=images;
  }else throw Error('支援 DOCX、XLSX、PDF、圖片及 TXT。');
  if(!r.text.trim()&&!r.images?.length)throw Error('沒有取得可用內容，請檢查檔案。');
  const before=structuredClone(draft.values),candidate=intakeValues(r,template(),before),initial=intakeReport(r,template(),before,candidate);
  if(!initial.recognized&&!r.images?.length){const message=initial.blocked.length?'未套用：'+initial.blocked.join('、')+'不明確，保留目前內容':'已讀到文字，但未辨識到採購單對應欄位；目前內容未變更';status(message);toast(message);return;}
  r.sourceDigest=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes)),b=>b.toString(16).padStart(2,'0')).join('');
  r.sourceData=bytes64(bytes);const duplicate=state.requirements.findIndex(q=>q.sourceDigest===r.sourceDigest);if(duplicate>=0){r.id=state.requirements[duplicate].id;r.createdAt=state.requirements[duplicate].createdAt;state.requirements[duplicate]=r;}else state.requirements.push(r);
  applySimpleRequirement(r);const report=intakeReport(r,template(),before,draft.values),updated=[...new Set(report.changed)];
  let message=updated.length?'已更新：'+updated.join('、'):report.recognized?'已讀取，對應欄位與目前內容相同':'已讀取需求圖片';
  if(r.images?.length)message+='；保留 '+r.images.length+' 張原圖';
  if(report.blocked.length)message+='；'+report.blocked.join('、')+'不明確，保留目前值';
  if(r.type==='Excel 工程流程單'&&!r.fields.quantity)message+='；工單未列訂購量，保留目前值';
  scan={text:r.text,message:file.name+'｜'+message};state.drafts[selected]=structuredClone(draft);await save();dirty=false;renderCreate();toast(message);

 }catch(e){status('分析未完成：'+e.message);toast('分析未完成：'+e.message);}finally{busy=false;document.body.classList.remove('busy');}
}

function validate(){if(!draft)return false;normalizeDraft(template(),draft);syncDueDate();refreshOrderNumber();if(!validOrderNumber(draft.orderNo)){toast('請填寫最後三碼編號，例如 035');$('#order-suffix')?.focus();return false;}if(state.orders.some(o=>o.id!==draft.id&&o.orderNo===draft.orderNo)){toast('此訂單號碼已存在，請修改最後三碼');return false;}return true;}

async function generate(){if(busy||!validate())return;busy=true;$('#export-word').disabled=true;try{const s=template();const blob=await exportDocx(s,draft.values,draft.orderNo);const record={...structuredClone(draft),savedAt:new Date().toISOString(),supplierName:s.name,templateSnapshot:structuredClone(s)};state.orders=state.orders.filter(o=>o.id!==draft.id);state.orders.unshift(record);state.drafts[selected]=structuredClone(draft);await save();download(blob,safeFile(draft.orderNo+'_'+s.short)+'採購單.docx');dirty=false;toast('Word 採購單已匯出，並儲存採購紀錄');}catch(e){toast('匯出失敗：'+e.message);}finally{busy=false;const b=$('#export-word');if(b)b.disabled=false;}}
function renderSuppliers(){app.innerHTML=title('SUPPLIERS & TEMPLATES','供應商與模板','每家供應商擁有自己的 Word 版型、付款條件與採購紀錄。','<button id="update-supplier">更新供應商模板</button><input id="update-template-file" type="file" accept=".docx" multiple hidden><button class="primary" id="add-supplier">＋ 新增供應商模板</button>')+`<div class="message" style="margin-bottom:22px">請上傳各供應商的 Word 模板，或從「開單資料與備份」還原原本備份。模板與採購紀錄只保存在目前瀏覽器。</div><div class="supplier-grid">${state.suppliers.map(s=>`<article class="supplier-card"><div style="display:flex;justify-content:space-between"><div class="avatar">${esc(s.short[0])}</div><span class="badge">專屬 Word 模板</span></div><h3>${esc(s.short)}</h3><p>${esc(s.name)}<br>${esc(s.filename)}</p><div class="actions"><button class="primary" data-use="${s.id}">建立採購單</button><button data-edit="${s.id}">編輯模板</button></div><div style="display:flex;justify-content:space-between;margin-top:12px"><button class="text small" data-original="${s.id}">下載模板原檔 ↗</button><button class="text small danger" data-delete="${s.id}">刪除</button></div></article>`).join('')}</div>`;
 $('#update-supplier').onclick=()=>$('#update-template-file').click();$('#update-template-file').onchange=async e=>{const files=Array.from(e.target.files||[]);if(!files.length)return;const previous=structuredClone(state);try{if(draft&&dirty)state.drafts[selected]=structuredClone(draft);let next;for(const file of files)next=await replaceSupplierTemplate(state,file);await save();activate(next.id);renderSuppliers();document.querySelector('#supplier-count').textContent=state.suppliers.length;toast('已更新 '+files.length+' 份供應商模板');}catch(err){state=previous;toast('更新失敗：'+err.message);}};$('#add-supplier').onclick=()=>editTemplate();for(const b of document.querySelectorAll('[data-use]'))b.onclick=()=>{activate(b.dataset.use);switchPage('create');};for(const b of document.querySelectorAll('[data-edit]'))b.onclick=()=>editTemplate(b.dataset.edit);for(const b of document.querySelectorAll('[data-original]'))b.onclick=async()=>{const s=state.suppliers.find(s=>s.id===b.dataset.original);const z=await JSZip.loadAsync(s.data,{base64:true});download(await z.generateAsync({type:'blob'}),s.filename);};for(const b of document.querySelectorAll('[data-delete]'))b.onclick=async()=>{const id=b.dataset.delete;if(!confirm('刪除此供應商與目前草稿？已匯出採購紀錄會保留。'))return;state.suppliers=state.suppliers.filter(s=>s.id!==id);delete state.drafts[id];await save();if(selected===id)activate(state.suppliers[0]?.id||'');renderSuppliers();$('#supplier-count').textContent=state.suppliers.length;};}
function editTemplate(id){let s=id?structuredClone(state.suppliers.find(t=>t.id===id)):null;
 $('#modal-body').innerHTML=`<h2>${id?'編輯供應商模板':'新增供應商模板'}</h2><p class="hint">上傳表格式的 Word 採購單。系統會找出可編輯欄位；原有頁首、圖片、表格與字型保留於匯出的 Word。</p><div class="form-grid"><label class="field">供應商簡稱<input id="tpl-short" value="${esc(s?.short||'')}" placeholder="例如：比斯捷"></label><label class="field">完整名稱<input id="tpl-name" value="${esc(s?.name||'')}" placeholder="供應商完整名稱"></label><label class="field full">${id?'替換 Word 版型（選填）':'上傳 Word 模板（必填）'}<input type="file" id="tpl-file" accept=".docx"></label></div><div id="tpl-status" class="scan-info"></div><div id="tpl-fields"></div><div class="actions"><button id="cancel-modal">取消</button><button id="save-template" class="primary">儲存模板</button></div>`;
 const renderFields=()=>{if(!s)return;$('#tpl-fields').innerHTML=`<div class="section-label">預設欄位與條款（${s.fields.length} 個欄位）</div><p class="hint">新單保留原稿預設內容，訂購日期與交期預設當天；上傳需求後只更新辨識到的項目。</p><div class="form-grid">${s.fields.map(f=>`<label class="field full">${esc(f.label)}<textarea data-tpl="${esc(f.id)}" rows="${f.value.length>170?4:2}">${esc(f.value)}</textarea></label>`).join('')}</div>`;};renderFields();modal.showModal();$('#cancel-modal').onclick=()=>modal.close();let loading=false;
 $('#tpl-file').onchange=async e=>{const f=e.target.files[0];if(!f)return;if(f.size>5*1024*1024){toast('模板上限 5 MB');return;}loading=true;$('#save-template').disabled=true;try{const data=await f.arrayBuffer(),p=await parseDocx(data);if(!p.orderNo||!['date','part','quantity','supplier'].every(k=>p.fields.some(f=>f.key===k)))throw Error('模板須有「訂單號碼」及表格欄位「訂購時間、料號、訂購數量、廠商名稱」。請先在 Word 補上標籤。');s={id:id||uid(),short:$('#tpl-short').value,name:$('#tpl-name').value||p.name,filename:f.name,data:bytes64(data),fields:p.fields,rows:p.rows};if(!$('#tpl-name').value)$('#tpl-name').value=p.name;$('#tpl-status').textContent=`辨識到 ${p.fields.length} 個欄位。請核對預設內容。`;renderFields();}catch(err){$('#tpl-status').textContent=err.message;}finally{loading=false;$('#save-template').disabled=false;}};
 $('#save-template').onclick=async()=>{if(loading)return;if(!s){toast('請先上傳 Word 模板');return;}s.short=$('#tpl-short').value.trim();s.name=$('#tpl-name').value.trim();if(!s.short||!s.name){toast('請填寫供應商簡稱與完整名稱');return;}for(const input of document.querySelectorAll('[data-tpl]'))s.fields.find(f=>f.id===input.dataset.tpl).value=input.value;const supplier=s.fields.find(f=>f.key==='supplier');if(supplier)supplier.value=s.name;const next=state.suppliers.filter(t=>t.id!==s.id);next.push(s);const old=state.suppliers;state.suppliers=next;try{await save();if(selected===s.id){delete state.drafts[s.id];activate(s.id);await save();}modal.close();renderSuppliers();$('#supplier-count').textContent=next.length;toast('供應商模板已儲存');}catch(e){state.suppliers=old;toast('儲存失敗：'+e.message);}};
}
function renderHistory(){renderTracking(app,state,{save,toast,download,word:async o=>{try{download(await exportDocx(o.templateSnapshot,o.values,o.orderNo),safeFile(o.orderNo+'_'+o.templateSnapshot.short)+'採購單.docx');}catch(error){toast(error.message);}},removeOrder:async id=>{if(!confirm('確定刪除此採購紀錄？'))return;const old=state.orders;state.orders=old.filter(o=>o.id!==id);try{await save();}catch(error){state.orders=old;toast(error.message);}}});}
function renderSettings(){const c=state.company;app.innerHTML=title('WORKSPACE SETTINGS','開單資料與備份','管理內容預覽的開單資訊，並備份各供應商模板和採購紀錄。')+`<div class="settings">${card('01','開單公司資訊',`<div class="form-grid">${[['name','公司名稱'],['taxId','統一編號'],['phone','聯絡電話'],['address','公司地址']].map(([k,l])=>`<label class="field ${k==='address'?'full':''}">${l}<input data-company="${k}" value="${esc(c[k])}"></label>`).join('')}</div><p class="message warning">此處更新內容預覽。Word 的公司抬頭、Logo 與地址沿用各供應商原稿；更換開單公司時，請先在 Word 修改模板頁首，再到「供應商與模板」上傳替換。</p><button id="save-company" class="primary">儲存開單資訊</button>`)}${card('02','資料備份',`<p class="sub">修改會先保留在這台裝置，再自動同步雲端。換電腦請使用上方的「專屬連結」；同步狀態會顯示在頁面上方。</p><div class="actions" style="margin-top:20px"><button id="backup" class="primary">↓ 匯出完整備份</button><button id="restore">↑ 還原備份</button></div><input id="restore-file" type="file" accept=".json" hidden><p class="hint">備份包含 Word 模板、草稿、採購紀錄、追蹤資料與 Excel 格式。還原會取代目前工作區。</p>`)}${card('03','使用方式',`<p class="sub">1. 選擇供應商，填入自訂三碼單號。<br>2. 上傳客戶需求，系統直接套入採購單。<br>3. 需要時修改內容，直接下載 Word。<br>年份與月份依台灣時間更新，已儲存的歷史訂單不會變更。</p><p class="hint">只套用採購單對應欄位；圖片保留原圖，僅帶入需求單明確標示的備註。未找到的資料保留預設。</p>`)}</div>`;$('#save-company').onclick=async()=>{for(const el of document.querySelectorAll('[data-company]'))state.company[el.dataset.company]=el.value;await save();toast('開單資訊已儲存');};$('#backup').onclick=backup;$('#restore').onclick=()=>$('#restore-file').click();$('#restore-file').onchange=async e=>{try{const f=e.target.files[0];if(!f)return;if(f.size>100*1024*1024)throw Error('備份超過 100 MB');const v=JSON.parse(await f.text());await validateBackup(v);if(!confirm('還原會取代目前所有供應商、草稿和紀錄。請確認已備份目前資料。'))return;const previous=state;state=v.workspace;state.requirements??=[];initTracking(state);try{await upgradeWorkspace(state);await save();}catch(err){state=previous;throw err;}activate(state.suppliers[0]?.id||'');toast('備份已還原');renderSettings();$('#supplier-count').textContent=state.suppliers.length;}catch(err){toast('無法還原：'+err.message);}};}
async function validateBackup(v){if(v?.format!=='purchase-studio-backup-v1'||!Array.isArray(v.workspace?.suppliers)||!Array.isArray(v.workspace?.orders)||typeof v.workspace?.drafts!=='object'||!v.workspace?.company)throw Error('備份格式不正確');const ids=new Set();const checkTemplate=async s=>{if(!s||typeof s.id!=='string'||typeof s.data!=='string'||!Array.isArray(s.fields)||!Array.isArray(s.rows)||typeof s.name!=='string'||typeof s.short!=='string'||!s.short)throw Error('模板資料不完整');const p=await parseDocx(Uint8Array.from(atob(s.data),c=>c.charCodeAt(0)));if(s.fields.some(f=>typeof f.id!=='string'||!Number.isInteger(f.cell)||f.cell<0||!p.fields.some(x=>x.cell===f.cell)))throw Error('模板欄位不正確');};for(const s of v.workspace.suppliers){if(ids.has(s.id))throw Error('供應商編號重複');ids.add(s.id);await checkTemplate(s);}for(const o of v.workspace.orders){if(!o.values||!o.orderNo)throw Error('採購紀錄格式不正確');await checkTemplate(o.templateSnapshot);}await validateTracking(v.workspace);}
async function backup(){try{const exportState=structuredClone(state);if(draft)exportState.drafts[selected]=structuredClone(draft);download(new Blob([JSON.stringify({format:'purchase-studio-backup-v1',exportedAt:new Date().toISOString(),workspace:exportState})],{type:'application/json'}),'採購工作室備份_'+today()+'.json');toast('完整備份已匯出');}catch(e){toast('備份失敗：'+e.message);}}
for(const b of document.querySelectorAll('nav button'))b.onclick=()=>{if(busy){toast('正在處理文件，請稍候');return;}switchPage(b.dataset.page);};$('#quick-backup').onclick=backup;
window.addEventListener('beforeunload',e=>{if(!window.__siteUpdating&&(dirty||cloud?.unsynced)){e.preventDefault();e.returnValue='';}});
window.prepareForSiteUpdate=async()=>{
 if(!state||!cloud||busy||modal.open)return false;
 if([...document.querySelectorAll('[data-company]')].some(input=>input.value!==String(state.company[input.dataset.company]||'')))return false;
 const revision=editCounter;
 if(dirty&&draft){state.drafts[selected]=structuredClone(draft);try{await save();}catch{return false;}if(revision!==editCounter)return false;dirty=false;}
 for(let i=0;i<20&&cloud.unsynced;i++){await new Promise(r=>setTimeout(r,150));if(dirty||busy||modal.open)return false;}
 return !cloud.unsynced&&!dirty&&!busy&&!modal.open&&revision===editCounter;
};
async function init(){app.innerHTML='<div class="empty">正在準備供應商專屬模板…</div>';try{db=await openDB();cloud=installCloud({readLocal,writeLocal,getState:cloudState,replaceState:replaceCloudState,isEditing:()=>dirty||busy||modal.open,toast});state=await cloud.read();if(!state){const r=await fetch('samples.json');if(!r.ok)throw Error('無法讀取初始模板');const samples=await r.json();state={suppliers:[],orders:[],drafts:{},company:{name:'',address:'',phone:'',taxId:''}};for(const f of samples.files){const data=Uint8Array.from(atob(f.data),c=>c.charCodeAt(0));const p=await parseDocx(data);state.suppliers.push({id:uid(),short:f.name.match(/[（(](.+?)[）)]/)?.[1]||'威爾高',name:p.name,filename:f.name,data:f.data,fields:p.fields,rows:p.rows});}await save({initial:true});}state.requirements??=[];initTracking(state);if(await upgradeWorkspace(state))await save({initial:true});activate(state.suppliers.find(s=>s.id===state.lastSupplierId)?.id||state.suppliers[0]?.id||'');render();cloud.start();}catch(e){app.innerHTML=`<div class="card"><div class="empty"><b>無法開啟工作區</b>${esc(e.message)}<p>請允許瀏覽器儲存資料，並確認使用 HTTP 網址開啟網站。</p><button onclick="location.reload()">重試</button></div></div>`;}}
init();

setInterval(()=>{if(page==='create')refreshOrderNumber();},30000);
document.addEventListener('visibilitychange',()=>{if(!document.hidden&&page==='create')refreshOrderNumber();});


function initializeDueDate(){
 if(!draft||!template()||draft.values.__inlineDue)return;
 if(!Object.hasOwn(draft.values,'__dueDate'))draft.values.__dueDate=draft.dueDate??creationDate(draft.createdAt).replaceAll('/','-');
 if(!Object.hasOwn(draft.values,'__dueRaw'))draft.values.__dueRaw=draft.dueRaw||'';
 draft.values=prepareInlineDue(template(),draft.values);
}
function syncDueDate(manual=false){
 if(!draft||!template()||draft.values.__inlineDue)return;
 draft.values.__dueDate=draft.dueDate||'';draft.values.__dueRaw=draft.dueRaw||'';
 for(const f of template().fields.filter(f=>f.key==='quantity'))draft.values[f.id]=quantityWithDue(draft.values[f.id],draft.dueDate,draft.dueRaw);
 if(manual){
  const rows=draft.values.__requirementAppendix||[],filtered=rows.filter(r=>!['需求交期','交期'].includes(r.label));
  if(draft.dueDate)filtered.push({label:'交期',value:draft.dueDate.replaceAll('-','/'),source:'採購指定'});
  draft.values.__requirementAppendix=filtered;
 }
}



function expandOrderText(){for(const el of document.querySelectorAll('textarea[data-field],textarea[data-annex]')){el.style.height='auto';el.style.height=(el.scrollHeight+2)+'px';}}
window.addEventListener('resize',()=>requestAnimationFrame(expandOrderText));

