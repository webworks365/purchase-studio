import {splitQuantityDue} from './due-date.js';
export const dueTarget=template=>template.fields.find(f=>['quantity','delivery'].includes(f.key)&&/交期/.test(f.value))||template.fields.find(f=>f.key==='quantity');
export function fixedDelivery(text){const value=String(text||''),at=value.indexOf('交期');return at<0?'銘星科技有限公司':'銘星科技有限公司'+(value.slice(0,at).match(/\s+$/)?.[0]||'\n')+value.slice(at);}
export function transferDue(template,values,raw){const target=dueTarget(template);if(!target)return;if(target.key==='delivery')values[target.id]=fixedDelivery(raw);else{const base=splitQuantityDue(values[target.id]).quantity;values[target.id]=/^\s*交期/.test(target.value)?[raw,base].filter(Boolean).join('\n'):[base,raw].filter(Boolean).join('\n');}}
export function applyDueRequirement(template,values,previous,requirement){
 if(!previous.__inlineDue)return;
 const target=dueTarget(template);if(!target)return;
 const due=requirement.fields.due?.value?.trim(),quantity=requirement.fields.quantity?.value?.trim();
 const old=String(previous[target.id]??target.value),parts=splitQuantityDue(old);
 if(!due&&!quantity)return;
 if(target.key==='delivery'){
  if(due)values[target.id]=fixedDelivery('交期：'+due);
  if(quantity){const f=template.fields.find(f=>f.key==='quantity');if(f)values[f.id]=quantity;}
  else for(const f of template.fields.filter(f=>f.key==='quantity'))values[f.id]=previous[f.id]??f.value;
  return;
 }
 const base=quantity||parts.quantity;
 if(quantity&&/交期/.test(quantity)&&!due){values[target.id]=quantity;return;}
 const clause=due?'交期：'+due:parts.rawDue;
 values[target.id]=/^\s*交期/.test(old)?[clause,base].filter(Boolean).join('\n'):[base,clause].filter(Boolean).join('\n');
}
