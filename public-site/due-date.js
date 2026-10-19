export function normalizeDueDate(value){
 const text=String(value||'');let y,m,d,match=text.match(/(\d{4})[\/.-]+(\d{1,2})[\/.-]+(\d{1,2})/);
 if(match)[,y,m,d]=match;else{match=text.match(/(\d{1,2})[\/.-]+(\d{1,2})[\/.-]+(\d{4})/);if(match)[,m,d,y]=match;}
 if(!y)return '';const date=new Date(Date.UTC(+y,+m-1,+d));if(date.getUTCFullYear()!==+y||date.getUTCMonth()!==+m-1||date.getUTCDate()!==+d)return '';
 return [y,String(m).padStart(2,'0'),String(d).padStart(2,'0')].join('-');
}
export function splitQuantityDue(value){
 const text=String(value||''),re=/交期[^\n]*?(?=(?:[-－—]+\s*)?\d[\d,]*\s*(?:PCS|PNL|SET|PANELS?)\b|\n|$)/gi,all=[...text.matchAll(re)];
 if(!all.length)return {quantity:text,rawDue:''};
 const quantity=text.replace(re,'').split('\n').map(line=>line.replace(/^[-－—]+\s*(?=\d)/,'').trimEnd()).join('\n').trim();
 return {quantity,rawDue:all.map(m=>m[0].trim()).join('\n')};
}
export function replaceOriginalDue(text,date){
 return String(text).replace(/交期[^\n]*?(?=(?:[-－—]+\s*)?\d[\d,]*\s*(?:PCS|PNL|SET|PANELS?)\b|\n|$)/gi,clause=>{
  const formatted=date.replaceAll('-','/'),pattern=/(?:\d{4}[\/.-]+\d{1,2}[\/.-]+\d{1,2}|\d{1,2}[\/.-]+\d{1,2}(?:[\/….-]+\d{4})?)/;
  return pattern.test(clause)?clause.replace(pattern,()=>formatted):'交期：'+formatted;
 });
}
export function quantityWithDue(quantity,date,rawDue=''){
 const base=splitQuantityDue(quantity).quantity.trim(),suffix=date?'交期：'+date.replaceAll('-','/'):(rawDue?(/^交期/.test(rawDue)?rawDue:'交期：'+rawDue):'');
 return [base,suffix].filter(Boolean).join('\n');
}
