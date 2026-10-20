// Shared vocabulary for spreadsheet cells, Word tables and extracted text.
// Coordinates describe reading relationships only; no company-specific cell addresses.
const tidy=s=>String(s??'').normalize('NFKC').trim();
const labelText=s=>tidy(s).toLowerCase().replace(/[\s:：()（）_\-./\\]/g,'');
const labels={
 customerPart:['客戶料號','客户料号','客戶產品料號','產品料號','产品料号','料號','料号','品號','customer part no','customer p/n','customer part number','part number','part no','p/n'],
 internalPart:['銘星料號','铭星料号','本公司料號','公司料號','內部料號','内部料号','internal part no','internal p/n','mingxing p/n'],
 quantity:['訂購數量','订购数量','訂單數量','订单数量','採購數量','采购数量','需求數量','需求数量','PCB需求量','訂購量','order qty','order quantity','purchase qty','required quantity','required qty','quantity required'],
 material:['材質','材质','基材材質','PCB原料材質','板材','材料規格','材料规格','材料板厚','material','laminate','base material'],
 layers:['PCB層數','PCB层数','層數','层数','layer count','number of layers','layers'],
 thickness:['成品板厚','成品厚度','成品板厚與公差','PCB厚度','板厚','finished thickness','finished board thickness','board thickness','pcb thickness'],
 copper:['成品銅厚','成品铜厚','成品外層銅厚','finished copper','finished copper weight'],
 baseCopper:['原始基銅','原始基铜','基銅','base copper','starting copper'],
 finish:['表面處理','表面处理','PCB電鍍型態','surface finish','surface treatment'],
 mask:['防焊','防焊顏色','防焊颜色','阻焊顏色','阻焊颜色','solder mask','solder mask color','soldermask color'],
 legend:['文字顏色','文字颜色','絲印顏色','丝印颜色','silkscreen','silkscreen color','legend color'],
 due:['交期','需求日期','交貨日期','交货日期','交付日期','due date','delivery date','required date'],
 inspection:['檢驗標準','检验标准','檢驗依據','inspection standard','acceptance standard'],
 size:['成品尺寸','PCB尺寸大小','finished size','board size','board dimensions'],
 panel:['拼板','連板','连板','排版','排版成型','panelization','pcs/panel','pcs per panel'],
 via:['導通孔處理','塞孔','via treatment'],
 revision:['版本','版次','revision','rev'],
 notes:['備註','备注','重要備註','特別備註','remarks','notes']
};
const lookup=new Map(Object.entries(labels).flatMap(([key,names])=>names.map(name=>[labelText(name),key])));
export const semanticLabel=s=>lookup.get(labelText(s));
const unit=/^(?:PCS|PIECES?|SETS?|PNL|PANELS?|片|套|個|个)$/i;
const quantityPattern=/(\d[\d,]*(?:\.\d+)?)\s*(PCS|PIECES?|SETS?|PNL|PANELS?|片|套|個|个)\b/gi;
const panelContext=/排版|拼板|連板|连板|每板|每拼|每套|每panel|pcs\s*[/／]\s*(?:panel|pnl|set)|per\s*(?:panel|pnl|set)|panelization|裁板|片數\/張|片数\/张/i;
function inlineLabel(line){
 const match=tidy(line).match(/^([^:：\n]+)[:：]\s*([\s\S]+)$/);
 if(match){const key=semanticLabel(match[1]);if(key)return {key,value:match[2].trim()};}
 // Space-delimited labels support forms such as "Order Qty 100 PCS".
 const words=tidy(line).split(/\s+/);for(let i=words.length-1;i>0;i--){const key=semanticLabel(words.slice(0,i).join(' '));if(key)return {key,value:words.slice(i).join(' ')};}
 return null;
}
function splitLabeledLine(line){
 // Semicolon/tab-separated named values remain separate, while a value's units/tolerance stay intact.
 return String(line).split(/\t|[；;](?=\s*[^；;:：]{1,40}[:：])/).map(s=>s.trim()).filter(Boolean);
}
function quantityCandidates(text,locator){
 if(panelContext.test(text))return [];
 const matches=[...String(text).matchAll(quantityPattern)];if(!matches.length)return [];
 // A bracketed conversion is not another order. Keep the complete original expression and units.
 if(matches.length>1&&!/\([^)]*\)|（[^）]*）|備品|备品|\+/.test(text))return matches.map(m=>({key:'quantity',value:m[0].trim(),locator}));
 if(matches.length===1||/\([^)]*\)|（[^）]*）|備品|备品|\+/.test(text))return [{key:'quantity',value:String(text).slice(matches[0].index).trim(),locator}];
 return [];
}
export function semanticText(text,{locator='文字段落',inferQuantity=true}={}){
 const lines=String(text||'').split(/\r?\n/).flatMap(splitLabeledLine),out=[],quantities=[];
 for(let i=0;i<lines.length;i++){
  const line=lines[i],inline=inlineLabel(line);
  if(inline){out.push({...inline,locator:locator+' '+(i+1)});continue;}
  const key=semanticLabel(line);if(key){let value=lines[i+1];if(value&&!semanticLabel(value)&&!inlineLabel(value)){i++;if(key==='quantity'&&/^\d[\d,.]*$/.test(value)&&unit.test(lines[i+1]||''))value+=' '+lines[++i];out.push({key,value,locator:locator+' '+(i+1)});}continue;}
  if(inferQuantity&&( /^(?:\d|數量|数量|qty|quantity)/i.test(line)||/訂購|訂單|採購|需求數量|order\s*(?:qty|quantity)/i.test(line)))quantities.push(...quantityCandidates(line,locator+' '+(i+1)));
 }
 if(!out.some(c=>c.key==='quantity')){const unique=new Map(quantities.map(c=>[c.value.replace(/\s/g,'').toUpperCase(),c]));if(unique.size===1)out.push(...unique.values());}
 return out;
}
const pos=c=>({row:c.row??Number(c.addr?.match(/\d+$/)?.[0]),col:c.col??[...(c.addr?.match(/^[A-Z]+/)?.[0]||'')].reduce((n,x)=>n*26+x.charCodeAt(0)-64,0)});
function range(ref){const [a,b=a]=ref.split(':'),start=pos({addr:a}),end=pos({addr:b});return {start,end};}
export function semanticGrid(cells,{merges=[],inferQuantity=true}={}){
 cells=cells.filter(c=>!c.reference&&String(c.value??c.text??'').trim()).map(c=>({...c,value:String(c.value??c.text),...pos(c)}));
 const merged=merges.map(range),out=[];
 for(const cell of cells){
  const locator=cell.addr||`表格 ${cell.row+1}/${cell.col+1}`,inline=semanticText(cell.value,{locator,inferQuantity:false});out.push(...inline);
  const key=semanticLabel(cell.value);if(!key)continue;
  const span=merged.find(m=>m.start.row===cell.row&&m.start.col===cell.col),right=span?.end.col??cell.col,bottom=span?.end.row??cell.row;
  const row=cells.filter(c=>c.row===cell.row&&c.col>right).sort((a,b)=>a.col-b.col),below=cells.filter(c=>c.row>bottom&&c.col>=cell.col&&c.col<=right).sort((a,b)=>a.row-b.row||a.col-b.col);
  let valueCell=row[0];
  if(!valueCell||semanticLabel(valueCell.value)||inlineLabel(valueCell.value))valueCell=below[0];
  if(!valueCell||semanticLabel(valueCell.value)||inlineLabel(valueCell.value))continue;
  let value=valueCell.value;
  if(['quantity','thickness','copper','size'].includes(key)&&/^\d[\d,.]*$/.test(value)){
   const following=cells.filter(c=>c.row===valueCell.row&&c.col>valueCell.col).sort((a,b)=>a.col-b.col)[0];
   if(following&&/^(?:PCS|PIECES?|SETS?|PNL|PANELS?|mm|mil|inch|oz|片|套|個|个)$/i.test(following.value.trim()))value+=' '+following.value;
  }
  out.push({key,value,locator:locator+' → '+(valueCell.addr||`${valueCell.row+1}/${valueCell.col+1}`)});
 }
 if(inferQuantity&&!out.some(c=>c.key==='quantity')){
  const unlabelled=[];
  for(const cell of cells){if(!/\d\s*(?:PCS|PIECES?|SETS?|PNL|PANELS?)\b/i.test(cell.value))continue;
   const preceding=cells.filter(c=>c.col===cell.col&&c.row<cell.row).sort((a,b)=>b.row-a.row)[0],left=cells.filter(c=>c.row===cell.row&&c.col<cell.col).sort((a,b)=>b.col-a.col)[0];
   if(panelContext.test([cell.value,preceding?.value,left?.value].join(' '))||/\b(?:PNL|PANELS?)\b/i.test(preceding?.value||''))continue;
   if(/^(?:\d|數量|数量|qty|quantity|訂購|訂單|採購)/i.test(cell.value.trim()))unlabelled.push(...quantityCandidates(cell.value,cell.addr||'數量文字'));
  }
  // Multiple unrelated quantities cannot be resolved from the PCS token alone.
  const unique=new Map(unlabelled.map(c=>[c.value.replace(/\s/g,'').toUpperCase(),c]));if(unique.size===1)out.push(...unique.values());
 }
 return out;
}
