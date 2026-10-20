import {applyDueRequirement,fixedDelivery,dueTarget} from './inline-due.js';
import {replaceOriginalDue} from './due-date.js';
import {bisSpareDefault} from './note-format.js';
import {requirementFields,requirementMapping,requirementTargets} from './requirements.js';
export const DELIVERY='銘星科技有限公司';
export function creationDate(iso){return new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Taipei'}).format(new Date(iso)).replaceAll('-','/');}
export function supplierDisplay(s){const company=s.fields.find(f=>f.key==='supplier')?.value||s.name||s.short;const raw=s.fields.find(f=>f.key==='contact')?.value||'';const contact=raw.split(/\s*(?:聯絡電話|電話|TEL|手機|E-?MAIL|EM|電子郵件)\s*[:：]/i)[0].trim();return [company,contact].filter(Boolean).join('｜');}
export function formRows(s){
 const groups=[['delivery','payment'],['supplier','contact'],['phone','email'],['item','date'],['customerPart','part']];
 const consumed=new Set(),result=[];
 for(const row of s.rows){let pending=[];const flush=()=>{if(pending.length){result.push(pending);pending=[];}};
  for(const c of row){const f=s.fields.find(f=>f.id===c.id);if(!f){if(c.text)pending.push(c);continue;}if(consumed.has(f.id))continue;
   const group=groups.find(g=>g.includes(f.key));if(group){flush();const fields=group.flatMap(k=>s.fields.filter(t=>t.key===k&&!consumed.has(t.id)));for(const t of fields)consumed.add(t.id);result.push(fields.map(t=>({id:t.id})));}
   else{consumed.add(f.id);pending.push(c);}
  }flush();
 }return result;
}
const keyOf=row=>row.key||Object.keys(requirementFields).find(k=>requirementFields[k]===row.label);
const compact=s=>String(s||'').replace(/\s/g,'').toLowerCase();
const containsValue=(text,value)=>String(value||'').split('\n').filter(s=>s.trim()).every(line=>compact(text).includes(compact(line)));
function addUnique(base,label,value){let text=String(base||'').trim();for(const line of String(value||'').split('\n').map(s=>s.trim()).filter(Boolean)){if(!compact(text).includes(compact(line)))text+=(text?'\n':'')+(label?label+'：':'')+line;label='';}return text;}
const mapped={customerPart:['customerPart'],internalPart:['part'],quantity:['quantity'],layers:['material'],material:['material'],thickness:['material'],copper:['copper'],mask:['mask'],legend:['mask'],finish:['finish','mask'],inspection:['inspection'],ul:['ul'],delivery:['delivery'],due:['quantity','delivery']};
export function supplementalRows(){return [];}
export function mergeSupplements(){}
function removeGeneratedNotes(s,values){
 const copies=(values.__requirementAppendix||[]).filter(row=>keyOf(row)!=='notes').flatMap(row=>String(row.value||'').split('\n').flatMap(line=>[line.trim(),row.label+'：'+line.trim()]));
 for(const f of s.fields.filter(f=>['notes','notesTail'].includes(f.key))){const original=new Set(String(f.value).split('\n').map(v=>v.trim()));values[f.id]=String(values[f.id]??f.value).split('\n').filter(line=>original.has(line.trim())||!copies.includes(line.trim())).join('\n');}
}
export function normalizeDraft(s,draft){
 draft.createdAt??=new Date().toISOString();
 const generated=new Set(['依本採購單「客戶製作要求」附件及已確認之製作檔案執行。','依本採購單之客戶製作要求附件，及本次經確認的供應商交付條款。','依本採購單客戶需求附件及客戶提供之製作檔案執行。','依本採購單客戶需求附件及本次交付條款。']);
 for(const f of s.fields){if(f.key==='date')draft.values[f.id]=creationDate(draft.createdAt);if(f.key==='delivery')draft.values[f.id]=draft.values.__inlineDue?fixedDelivery(draft.values[f.id]):DELIVERY;
  if(['production','notes','notesTail'].includes(f.key)&&generated.has(draft.values[f.id]))draft.values[f.id]=f.value;
 }
 if(draft.bisSpareVersion!==1){bisSpareDefault(s,draft.values);draft.bisSpareVersion=1;}
 if(draft.notesPolicyVersion!==2){removeGeneratedNotes(s,draft.values);draft.notesPolicyVersion=2;}
}
function patchCombined(base,key,value){
 if(key==='thickness'&&compact(base).includes(compact(value)))return base;
 const patterns={layers:/(?:PCB\s*)?層數\s*[:：]\s*\d+\s*(?:L(?:ayers?)?|層)?|\b\d+\s*(?:L(?:ayers?)?\b|層)/i,thickness:/(?:成品板厚(?:與公差)?|板厚|厚度|\bT)\s*[:：]?\s*\d+(?:\.\d+)?(?:\s*(?:±|\+\/-)\s*\d+(?:\.\d+)?)?\s*T?\s*(?:mm|mil|inch)?|\b\d+(?:\.\d+)?(?:\s*(?:±|\+\/-)\s*\d+(?:\.\d+)?)?\s*(?:mm|mil|inch)/i};
 if(key==='material'){if(/^FR-?4$/i.test(value.trim())&&/^FR-?4(?:\b|TG)/i.test(base.trim()))return base;if(/[／/]/.test(value))return value;const first=base.search(/[／/]/);if(first>=0)return value+base.slice(first);const at=base.search(/\b\d+\s*L\b|(?:板厚|厚度|\bT)\s*\d/i);return at>0?value+' '+base.slice(at):value;}
 if(patterns[key]){const v=key==='layers'&&!/層|L(?:ayers?)?$/i.test(value)?value+'L':value;return patterns[key].test(base)?base.replace(patterns[key],()=>v):[base,requirementFields[key]+'：'+value].filter(Boolean).join('／');}
 return value;
}
export function intakeValues(r,s,previous={}){
 const cleanPrevious={...previous},old=previous.__requirementApplied,part=r.fields.customerPart?.value?.trim(),changedCustomer=old?.customerPart&&part&&compact(old.customerPart)!==compact(part);
 if(changedCustomer){for(const f of s.fields)if(old.values?.[f.id]!==undefined&&previous[f.id]===old.values[f.id])cleanPrevious[f.id]=previous.__dueDate?replaceOriginalDue(f.value,previous.__dueDate):f.value;cleanPrevious.__images=[];}
 removeGeneratedNotes(s,cleanPrevious);const mapped=requirementMapping(r,s),values={...cleanPrevious,...mapped};
 for(const f of s.fields)if(!Object.hasOwn(values,f.id))values[f.id]=f.value;
 for(const f of s.fields){
  const items=Object.entries(r.fields).filter(([key,e])=>!e.blocked&&e.value?.trim()&&requirementTargets(s,key).some(t=>t.id===f.id));
  if(f.key==='material'&&items.length){const complete=['material','layers','thickness'].every(key=>items.some(([k])=>k===key));let base=complete?f.value:cleanPrevious[f.id]??f.value;for(const key of ['material','layers','thickness']){const e=items.find(([k])=>k===key);if(e)base=patchCombined(base,key,e[1].value);}values[f.id]=base;}
  if(['production','notes','notesTail'].includes(f.key)&&mapped[f.id]!==undefined)values[f.id]=addUnique(cleanPrevious[f.id]??f.value,'',mapped[f.id]);
  if(f.key==='delivery')values[f.id]=values.__inlineDue?fixedDelivery(values[f.id]):DELIVERY;
 }
 applyDueRequirement(s,values,cleanPrevious,r);mergeSupplements(s,values);
 const applied=changedCustomer?{}:{...old?.values};for(const [key,entry]of Object.entries(r.fields))if(!entry.blocked&&entry.value?.trim())for(const f of requirementTargets(s,key))if(!['date','delivery'].includes(f.key))applied[f.id]=values[f.id];
 if(r.fields.due?.value&&!r.fields.due.blocked){const target=dueTarget(s);if(target)applied[target.id]=values[target.id];}
 values.__requirementApplied={customerPart:part||old?.customerPart||'',values:applied};return values;
}
export function intakeReport(r,s,before,after){
 const targets=new Set(Object.entries(r.fields).filter(([,e])=>!e.blocked&&e.value?.trim()).flatMap(([key])=>requirementTargets(s,key)).filter(f=>!['date','delivery'].includes(f.key)).map(f=>f.id));
 const same=(a,b)=>String(a??'').replace(/\s+/g,' ').trim()===String(b??'').replace(/\s+/g,' ').trim();
 const changed=s.fields.filter(f=>targets.has(f.id)&&!same(before[f.id]??f.value,after[f.id]));
 const unchanged=s.fields.filter(f=>targets.has(f.id)&&same(before[f.id]??f.value,after[f.id]));
 const blocked=Object.values(r.fields).filter(e=>e.blocked).map(e=>e.label);
 return {recognized:targets.size,changed:changed.map(f=>f.label),unchanged:unchanged.map(f=>f.label),blocked};
}
