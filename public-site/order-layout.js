import {requirementFields,requirementMapping,requirementTargets} from './requirements.js';
export const DELIVERY='銘星科技有限公司';
export function creationDate(iso){return new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Taipei'}).format(new Date(iso)).replaceAll('-','/');}
export function supplierDisplay(s){const company=s.fields.find(f=>f.key==='supplier')?.value||s.name||s.short;const raw=s.fields.find(f=>f.key==='contact')?.value||'';const contact=raw.split(/\s*(?:聯絡電話|電話|TEL|手機|E-?MAIL|EM|電子郵件)\s*[:：]/i)[0].trim();return [company,contact].filter(Boolean).join('｜');}
export function formRows(s){
 const groups=[['delivery','payment'],['supplier','contact'],['phone','email'],['item','date'],['customerPart','part'],['material','copper'],['mask','finish']];
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
export function supplementalRows(s,values){return (values.__requirementAppendix||[]).filter(row=>{const key=keyOf(row);if(requirementTargets(s,key).length)return false;return !s.fields.filter(f=>['notes','notesTail','production'].includes(f.key)).some(f=>containsValue(values[f.id],row.value));});}
export function mergeSupplements(s,values){
 const target=s.fields.find(f=>f.key==='notesTail'&&!/核\s*准|廠商回簽/.test(f.value))||s.fields.find(f=>f.key==='notes')||s.fields.find(f=>f.key==='production');
 if(!target)return;
 for(const row of supplementalRows(s,values))values[target.id]=addUnique(values[target.id],row.label,row.value);
}
export function normalizeDraft(s,draft){
 draft.createdAt??=new Date().toISOString();
 const generated=new Set(['依本採購單「客戶製作要求」附件及已確認之製作檔案執行。','依本採購單之客戶製作要求附件，及本次經確認的供應商交付條款。','依本採購單客戶需求附件及客戶提供之製作檔案執行。','依本採購單客戶需求附件及本次交付條款。']);
 for(const f of s.fields){if(f.key==='date')draft.values[f.id]=creationDate(draft.createdAt);if(f.key==='delivery')draft.values[f.id]=DELIVERY;
  if(['production','notes','notesTail'].includes(f.key)&&generated.has(draft.values[f.id]))draft.values[f.id]=f.value;
 }
 if(draft.originalLayoutVersion!==1){mergeSupplements(s,draft.values);draft.originalLayoutVersion=1;}
}
export function intakeValues(r,s,previous={}){
 const values=requirementMapping(r,s);
 for(const f of s.fields){
  if(['production','notes','notesTail'].includes(f.key))values[f.id]=values[f.id]===undefined?f.value:addUnique(f.value,'',values[f.id]);
  if(['part','price'].includes(f.key)&&!(f.key==='part'&&r.fields.internalPart?.value))values[f.id]=previous[f.id]||f.value;
  if(f.key==='delivery')values[f.id]=DELIVERY;
 }
 mergeSupplements(s,values);return values;
}
