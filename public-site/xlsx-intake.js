const compact=s=>String(s||'').replace(/[\s:：()（）]/g,'').toLowerCase();
const column=s=>[...s].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
const position=c=>({row:Number(c.addr.match(/\d+$/)[0]),col:column(c.addr.match(/^[A-Z]+/)[0])});
const hasUnit=s=>/(?:mm|mil|inch|["”″])/i.test(s);
export function readManufacturingCells(sheet,put,selectedOptions){
 const cells=sheet.cells.filter(c=>!c.reference),at=c=>position(c),sameRow=c=>cells.filter(t=>at(t).row===at(c).row).sort((a,b)=>at(a).col-at(b).col);
 const next=c=>sameRow(c).find(t=>at(t).col>at(c).col),find=(label,filter=()=>true)=>cells.find(c=>label.test(compact(c.value))&&filter(c));
 const apply=(key,c,value=c?.value,ambiguity='')=>{if(c&&value?.trim())put(key,value,c.addr,ambiguity);};
 const heading=find(/層板工作流程表/),front=!!heading;
 if(front){
  const layers=compact(heading.value).match(/(\d+)層板工作流程表/);if(layers)apply('layers',heading,layers[1]);
  const part=find(/^(?:客戶料號|產品料號|料號)$/);if(part)apply('customerPart',next(part));
  const material=find(/^(?:材質|基材材質|pcb原料材質)$/);if(material)apply('material',next(material));
  const finished=find(/^成品(?:規格)?$/),finishedColumn=finished&&at(finished).col;
  if(finished){
   for(const [key,label]of [['size',/^(?:尺寸|成品尺寸)$/],['thickness',/^(?:厚度|板厚|成品板厚)$/],['copper',/^(?:銅厚|成品銅厚)$/]]){
    const labelCell=find(label,c=>at(c).col>=finishedColumn&&at(c).row>=at(finished).row&&at(c).row<=at(finished).row+3),value=labelCell&&next(labelCell);
    if(value)apply(key,value,value.value,['size','thickness'].includes(key)&&!hasUnit(value.value)?'原文未標示長度單位，不自行當成 inch 或 mm':'');
   }
  }
  const base=[];for(const [label,title]of [[/^內層材料$/,'內層'],[/^外層材料$/,'外層']]){const group=find(label);if(!group)continue;const left=at(group).col,right=finishedColumn||Infinity;const labelCell=find(/^銅厚$/,c=>at(c).col>=left&&at(c).col<right&&at(c).row>=at(group).row&&at(c).row<=at(group).row+3);const value=labelCell&&next(labelCell);if(value&&/[\dH]/i.test(value.value))base.push({cell:value,text:title+'：'+value.value});}
  if(base.length)apply('baseCopper',base[0].cell,base.map(v=>v.text).join('；'));
  const panel=find(/^排版$/);if(panel){const c=next(panel);if(c){const lower=cells.find(t=>at(t).col===at(c).col&&at(t).row===at(c).row+1);apply('panel',c,[c.value,lower?.value].filter(Boolean).join(' / '),'排版用量／PCS 每板資訊，不代表訂購數量');}}
  const processLabel=find(/^流程$/,c=>at(c).row>=9);if(processLabel){const processes=cells.filter(c=>at(c).col===at(processLabel).col&&at(c).row>at(processLabel).row&&at(c).row<=40);if(processes.length)apply('process',processes[0],processes.map(c=>c.value).join('；'));}
 }
 // Bind manufacturing labels to neighboring values, rather than fixed Excel coordinates.
 const aliases={customerPart:['客戶料號','產品料號','料號'],internalPart:['銘星料號','本公司料號'],quantity:['訂購數量','訂單數量','pcb需求量'],material:['材料規格','材料板厚'],thickness:['成品板厚'],copper:['成品銅厚'],mask:['防焊','防焊顏色'],finish:['表面處理'],legend:['文字顏色'],due:['交期','需求日期'],notes:['備註','重要備註','特別備註']};
 for(const cell of cells)for(const [key,names]of Object.entries(aliases)){
  if(names.includes(compact(cell.value))){const value=next(cell);if(value)apply(key,value,selectedOptions(value.value));}
  else{const inline=cell.value.match(/^\s*([^:：\n]+)\s*[:：]\s*([^\n]+)$/);if(inline&&names.includes(compact(inline[1])))apply(key,cell,selectedOptions(inline[2]));}
 }
 if(/反面|背面/.test(sheet.name)){
  const mask=cells.find(c=>/防焊材料|防焊顏色/.test(c.value));if(mask){if(/[■☑█]/.test(mask.value))apply('mask',mask,selectedOptions(mask.value));else{const row=sameRow(mask),color=row.find((c,i)=>i>0&&/^[■☑█]$/.test(row[i-1].value.trim())&&/綠|绿|白|黑|黃|黄|藍|蓝|紅|红|green|white|black|yellow|blue|red/i.test(c.value));if(color)apply('mask',color,color.value);}}
  const legend=cells.find(c=>/[■☑█]\s*(?:單|雙)/.test(c.value)&&/[■☑█]\s*(?:白|黃|黑)/.test(c.value));
  if(legend){const selected=selectedOptions(legend.value).split('；'),side=selected.find(v=>/^(?:單|雙)/.test(v)),color=selected.find(v=>/^(?:白|黑|黃)/.test(v));apply('legend',legend,[color,side?side+'面':''].filter(Boolean).join('／'));}
  else{const selected=cells.filter(c=>{const row=sameRow(c),i=row.indexOf(c);return i>0&&/^[■☑█]$/.test(row[i-1].value.trim());}),side=selected.find(c=>/^(?:單|雙)$/.test(c.value.trim())),color=selected.find(c=>/^(?:白|黃|黑)$/.test(c.value.trim()));if(color)apply('legend',color,[color.value,side?side.value+'面':''].filter(Boolean).join('／'));}
  const processLabel=find(/^流程$/),processCol=processLabel?at(processLabel).col:3;
  const finishes=cells.filter(c=>at(c).col===processCol&&/化金|鍍金|浸金|沉金|噴(?:純|無鉛|有鉛)?錫|化銀|沉銀|浸銀|\b(?:OSP|ENTEK|ENIG|ENEPIG|HASL)\b/i.test(c.value));
  if(finishes.length)apply('finish',finishes[0],finishes.map(c=>c.value).join('／'),finishes.length>1||sheet.lines?'流程列不代表已啟用製程，需依劃除標記確認本次表面處理':'');
  const inspection=cells.find(c=>/SPEC\s*依據|檢驗標準/i.test(c.value));if(inspection)apply('inspection',inspection,inspection.value.split(/[□☐\uE000-\uF8FF]/)[0].trim());
  const reports=cells.find(c=>/附切片|附試錫板|檢驗報告/.test(c.value));if(reports)apply('reports',reports,selectedOptions(reports.value));
  const via=cells.find(c=>/[■☑█]\s*需要塞孔/.test(c.value));if(via){const continuation=cells.find(c=>at(c).col===at(via).col&&at(c).row===at(via).row+1);apply('via',via,[selectedOptions(via.value).split(';')[0],continuation?.value].filter(Boolean).join('；'));}
 }
 return front;
}
