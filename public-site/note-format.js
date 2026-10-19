const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const SPARE='備品 : 此筆訂單客戶要求備品2set';
const text=p=>Array.from(p.getElementsByTagNameNS(W,'t')).map(t=>t.textContent).join('');
export function bisSpareDefault(template,values){
 if(!/比斯[捷傑]/.test(template.name||template.fields.find(f=>f.key==='supplier')?.value||''))return;
 const notes=template.fields.filter(f=>['notes','notesTail'].includes(f.key));if(notes.some(f=>String(values[f.id]??f.value).includes(SPARE)))return;
 const target=notes.find(f=>/ROHS|REACH/i.test(values[f.id]??f.value));
 if(target){const lines=String(values[target.id]??target.value).split('\n'),at=lines.findLastIndex(line=>/ROHS|REACH/i.test(line));lines[at]=lines[at].replace(/\s*備品\s*[:：]\s*此筆訂單客戶要求備品\s*\d+\s*set\s*/ig,'').trimEnd()+' '+SPARE;values[target.id]=lines.join('\n');}
 else{const f=notes.find(f=>f.key==='notesTail')||notes[0];if(f)values[f.id]=[values[f.id]??f.value,'須遵守ROHS. REACH 相關法規 '+SPARE].filter(Boolean).join('\n');}
}
export function extendNoteParagraphs(cell,value){
 const paragraphs=Array.from(cell.children).filter(n=>n.localName==='p'),sample=paragraphs.findLast(p=>p.getElementsByTagNameNS(W,'numPr').length);
 if(!sample)return value;
 const lines=String(value).split('\n').map(line=>line.replace(/^\s*\d{1,3}[.、)]\s*/,''));
 while(paragraphs.length<lines.length){const p=sample.cloneNode(true);for(const child of Array.from(p.children))if(child.localName!=='pPr')child.remove();const run=cell.ownerDocument.createElementNS(W,'w:r'),props=sample.getElementsByTagNameNS(W,'rPr')[0];if(props)run.appendChild(props.cloneNode(true));run.appendChild(cell.ownerDocument.createElementNS(W,'w:t'));p.appendChild(run);cell.appendChild(p);paragraphs.push(p);}
 return lines.join('\n');
}
export async function formatNoteLists(zip,doc,template){
 const part=zip.file('word/numbering.xml');if(!part)return;
 const numbers=new DOMParser().parseFromString(await part.async('string'),'application/xml');let id=Math.max(0,...Array.from(numbers.getElementsByTagNameNS(W,'num')).map(n=>Number(n.getAttributeNS(W,'numId'))||0));let changed=false;
 for(const index of new Set(template.fields.filter(f=>['notes','notesTail'].includes(f.key)).map(f=>f.cell))){const cell=doc.getElementsByTagNameNS(W,'tc')[index],paragraphs=Array.from(cell.children).filter(n=>n.localName==='p'),numbered=paragraphs.filter(p=>p.getElementsByTagNameNS(W,'numPr').length),first=numbered[0];if(!first)continue;
  const originalId=first.getElementsByTagNameNS(W,'numId')[0]?.getAttributeNS(W,'val'),original=Array.from(numbers.getElementsByTagNameNS(W,'num')).find(n=>n.getAttributeNS(W,'numId')===originalId);if(!original)continue;
  const num=original.cloneNode(true),nextId=String(++id);num.setAttributeNS(W,'w:numId',nextId);for(const old of Array.from(num.children).filter(n=>n.localName==='lvlOverride'))old.remove();const levels=new Set(numbered.map(p=>p.getElementsByTagNameNS(W,'ilvl')[0]?.getAttributeNS(W,'val')||'0'));
  for(const level of levels){const override=numbers.createElementNS(W,'w:lvlOverride');override.setAttributeNS(W,'w:ilvl',level);const start=numbers.createElementNS(W,'w:startOverride');start.setAttributeNS(W,'w:val','1');override.appendChild(start);num.appendChild(override);}numbers.documentElement.appendChild(num);
  for(const p of numbered){const props=p.getElementsByTagNameNS(W,'numPr')[0];if(!text(p).trim()){props.remove();continue;}let ref=props.getElementsByTagNameNS(W,'numId')[0];if(!ref){ref=doc.createElementNS(W,'w:numId');props.appendChild(ref);}ref.setAttributeNS(W,'w:val',nextId);}changed=true;
 }
 if(changed)zip.file('word/numbering.xml',new XMLSerializer().serializeToString(numbers));
}
export function highlightSpare(doc){
 for(const p of Array.from(doc.getElementsByTagNameNS(W,'p'))){const nodes=Array.from(p.getElementsByTagNameNS(W,'t')),joined=nodes.map(t=>t.textContent).join(''),start=joined.indexOf(SPARE);if(start<0)continue;let offset=0;
  for(const t of nodes){const value=t.textContent,end=offset+value.length,from=Math.max(0,start-offset),to=Math.min(value.length,start+SPARE.length-offset);offset=end;if(to<=from)continue;const run=t.parentNode;if(run.localName!=='r')continue;const make=(value,yellow)=>{const r=doc.createElementNS(W,'w:r'),existing=Array.from(run.children).find(n=>n.localName==='rPr');const props=existing?existing.cloneNode(true):doc.createElementNS(W,'w:rPr');if(yellow){for(const n of Array.from(props.children).filter(n=>n.localName==='highlight'))n.remove();const h=doc.createElementNS(W,'w:highlight');h.setAttributeNS(W,'w:val','yellow');props.appendChild(h);}r.appendChild(props);const node=doc.createElementNS(W,'w:t');node.setAttribute('xml:space','preserve');node.textContent=value;r.appendChild(node);return r;};
   const selected=make(value.slice(from,to),true),after=value.slice(to);t.textContent=value.slice(0,from);run.parentNode.insertBefore(selected,run.nextSibling);if(after)run.parentNode.insertBefore(make(after,false),selected.nextSibling);
  }
 }
}
