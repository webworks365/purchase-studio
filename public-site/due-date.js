export function normalizeDueDate(value){
 const text=String(value||'');let y,m,d,match=text.match(/(\d{4})[\/.-]+(\d{1,2})[\/.-]+(\d{1,2})/);
 if(match)[,y,m,d]=match;else{match=text.match(/(\d{1,2})[\/.-]+(\d{1,2})[\/.-]+(\d{4})/);if(match)[,m,d,y]=match;}
 if(!y)return '';const date=new Date(Date.UTC(+y,+m-1,+d));if(date.getUTCFullYear()!==+y||date.getUTCMonth()!==+m-1||date.getUTCDate()!==+d)return '';
 return [y,String(m).padStart(2,'0'),String(d).padStart(2,'0')].join('-');
}
export function splitQuantityDue(value){
 const text=String(value||'');const match=text.match(/交期[\s\S]*?(?=(?:[-－—]+\s*)?\d[\d,]*\s*(?:PCS|PNL|SET|PANELS?)\b|$)/i);
 if(!match)return {quantity:text,rawDue:''};
 const quantity=(text.slice(0,match.index)+text.slice(match.index+match[0].length)).trim().replace(/^[-－—]+\s*(?=\d)/,'');
 return {quantity,rawDue:match[0].trim()};
}
export function quantityWithDue(quantity,date,rawDue=''){
 const base=splitQuantityDue(quantity).quantity.trim(),suffix=date?'交期：'+date.replaceAll('-','/'):(rawDue?(/^交期/.test(rawDue)?rawDue:'交期：'+rawDue):'');
 return [base,suffix].filter(Boolean).join('\n');
}
