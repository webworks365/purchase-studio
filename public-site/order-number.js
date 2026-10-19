export function orderPrefix(date=new Date()){
 const parts=new Intl.DateTimeFormat('en-US',{timeZone:'Asia/Taipei',year:'numeric',month:'2-digit'}).formatToParts(date);
 const get=k=>parts.find(p=>p.type===k).value;
 return 'MH'+get('year').slice(-2)+'A-'+get('month')+'-';
}
export function formatOrderNumber(suffix,date=new Date()){
 const text=String(suffix??'').trim();
 return orderPrefix(date)+(/^\d{1,3}$/.test(text)?text.padStart(3,'0'):text);
}
export function validOrderNumber(value){return /^MH\d{2}A-(?:0[1-9]|1[0-2])-\d{3}$/.test(value);}
