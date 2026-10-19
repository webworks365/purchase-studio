import {parseDocx} from './core.js';
import {requirementTargets,requirementMapping,requirementFields} from './requirements.js';
const compact=s=>String(s||'').replace(/\s/g,'');
export function remapDraft(old,next,draft){
 if(!draft)return;
 const values={...draft.values};
 for(const f of next.fields){const before=old.fields.find(t=>t.id===f.id&&t.key===f.key)||old.fields.find(t=>t.key===f.key);let value=before?draft.values[before.id]:undefined;
  if(value===undefined||value===before?.value)value=f.value;
  if(f.segment&&before&&!before.segment&&before.cell===f.cell){const label=f.key==='delivery'?/\s*付款方式\s*[:：]/:f.key==='mask'?/\s*表面處理\s*[:：]/:['inspection','production'].includes(f.key)?/\s*UL\s*MARK\s*[:：]/i:null;if(label)value=value.split(label)[0].trimEnd();}
  values[f.id]=value;
 }
 // Move only exact generated copies into matching original table fields.
 const entries=(values.__requirementAppendix||[]).map(e=>({...e,key:e.key||Object.keys(requirementFields).find(k=>requirementFields[k]===e.label)}));
 const mapped=requirementMapping({fields:Object.fromEntries(entries.filter(e=>e.key).map(e=>[e.key,e]))},next);
 for(const entry of entries){const targets=requirementTargets(next,entry.key);if(!targets.length||['delivery','due'].includes(entry.key))continue;
  for(const f of targets){const before=old.fields.find(t=>t.id===f.id);if(!values[f.id]||values[f.id]===f.value||values[f.id]===before?.value)values[f.id]=mapped[f.id]??values[f.id];}
  for(const note of next.fields.filter(f=>['notes','notesTail'].includes(f.key)&&!targets.includes(f))){const exact=new Set([entry.value,entry.label+'：'+entry.value]);values[note.id]=String(values[note.id]||'').split('\n').filter(line=>!exact.has(line.trim())).join('\n');}
 }
 draft.values=values;delete draft.originalLayoutVersion;
}
export async function upgradeWorkspace(state){
 if(state.tableMappingVersion===2)return false;
 for(let i=0;i<state.suppliers.length;i++){const old=state.suppliers[i],p=await parseDocx(Uint8Array.from(atob(old.data),c=>c.charCodeAt(0))),next={...old,fields:p.fields,rows:p.rows};
  for(const f of next.fields){const before=old.fields.find(t=>t.id===f.id&&t.cell===f.cell);if(before&&!f.segment)f.value=before.value;}
  remapDraft(old,next,state.drafts[old.id]);state.suppliers[i]=next;
 }state.tableMappingVersion=2;return true;
}
export async function replaceSupplierTemplate(state,file){
 if(!/\.docx$/i.test(file.name)||file.size>5*1024*1024)throw Error('請選擇 5 MB 以內的 Word 模板');
 const bytes=await file.arrayBuffer(),p=await parseDocx(bytes);if(!p.orderNo||!['date','part','quantity','supplier'].every(k=>p.fields.some(f=>f.key===k)))throw Error('無法辨識採購單必要欄位');
 const old=state.suppliers.find(s=>compact(s.name)===compact(p.name)||compact(s.fields.find(f=>f.key==='supplier')?.value)===compact(p.name));if(!old)throw Error('找不到相同公司，請使用「新增供應商模板」');
 let data='';const a=new Uint8Array(bytes);for(let i=0;i<a.length;i+=16384)data+=String.fromCharCode(...a.subarray(i,i+16384));
 const next={...old,name:p.name,filename:file.name,data:btoa(data),fields:p.fields,rows:p.rows};remapDraft(old,next,state.drafts[old.id]);state.suppliers[state.suppliers.indexOf(old)]=next;return next;
}
