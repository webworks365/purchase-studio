import {splitQuantityDue} from './due-date.js';
const unique=values=>[...new Set(values.map(v=>v.trim()).filter(Boolean))];
function derived(r,key,label,value,source){if(!r.fields[key])r.fields[key]={...source,key,label,value,ambiguity:'',blocked:false};}
function block(r,key,detail){const f=r.fields[key];if(!f)return;f.blocked=true;f.ambiguity=detail;if(!r.issues.some(i=>i.id==='unsafe-'+key))r.issues.push({id:'unsafe-'+key,severity:'block',field:key,title:f.label+'無法直接套用',detail,source:r.name});}
// Split a known combined field once. Never infer a manufacturing unit or resolve contradictory specs.
export function normalizeIntake(r){
 const q=r.fields.quantity;if(q){const parts=splitQuantityDue(q.value);if(parts.rawDue){q.value=parts.quantity;if(!q.blocked)derived(r,'due','需求交期',parts.rawDue.replace(/^交期\s*[:：]?\s*/,''),q);}}
 if(q&&!q.blocked&&!/(?:PCS|PIECES?|SETS?|PNL|PANELS?)\b|片|套|個|个/i.test(q.value))block(r,'quantity','數量未明示單位，保留目前值，不自行換算');
 const material=r.fields.material;
 if(material&&/(?:PCB\s*層數|成品板厚與公差)\s*[:：]/.test(material.value)){
  const raw=material.value,layerValues=unique([...raw.matchAll(/(?:PCB\s*)?層數\s*[:：]\s*(\d+)\s*(?:L(?:ayers?)?|層)?/gi)].map(m=>m[1]));
  const thicknessValues=unique([...raw.matchAll(/成品板厚(?:與公差)?\s*[:：]\s*([^／\n]+)/g)].map(m=>m[1]));
  if(layerValues.length){derived(r,'layers','PCB 層數',layerValues[0],material);if(layerValues.length>1)block(r,'layers','文件同時列出 '+layerValues.join('、')+' 層，保留目前值');}
  if(thicknessValues.length){derived(r,'thickness','成品板厚與公差',thicknessValues[0],material);if(thicknessValues.length>1||thicknessValues.some(v=>(v.match(/±/g)||[]).length>1))block(r,'thickness','同一材料欄有多組或重複板厚，保留目前值');}
  material.value=raw.split(/[／\n](?=\s*(?:PCB\s*層數|成品板厚))/)[0].trim();
 }
 const copper=r.fields.copper;if(copper){const match=/原始基銅\s*[:：]\s*([\s\S]+)/.exec(copper.value);if(match){derived(r,'baseCopper','原始基銅',match[1].trim(),copper);copper.value=copper.value.slice(0,match.index).trim();}}
 const thickness=r.fields.thickness;if(thickness&&!thickness.blocked&&/\d/.test(thickness.value)&&!/(?:mm|mil|inch|英吋|英寸|毫米|[μµ]m|["”″])/i.test(thickness.value))block(r,'thickness','原文未標示板厚單位，保留目前值，不自行當成 inch 或 mm');
 const finish=r.fields.finish;
 if(finish){const options=finish.value.replace(/[■□☑☒☐]/g,'').split(/[\s；;,，/／]+/).filter(Boolean);if(options.length&&options.every(s=>/^(?:單|雙|白|黑|黃|無|單面|雙面|白色|黑色|黃色)$/.test(s)))block(r,'finish','這些內容是文字面別／顏色，不是表面處理，保留目前值');if(/流程列不代表已啟用/.test(finish.ambiguity||''))block(r,'finish',finish.ambiguity);}
 if(r.fields.process&&r.issues.some(i=>i.id==='visual-process'))block(r,'process','工單含劃除線，無法確認哪些流程啟用，保留目前製作內容');
 return r;
}
