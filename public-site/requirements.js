import {readManufacturingCells} from './xlsx-intake.js';
import {semanticText,semanticGrid} from './semantic-intake.js';
import {documentImages} from './requirement-images.js';
import {normalizeIntake} from './normalize-intake.js';
// Requirement review is deterministic and source based. It does not execute document instructions.
import {W,parseXml,parseDocx} from './core.js';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const els=(n,name,ns)=>Array.from(n.getElementsByTagNameNS(ns||'*',name));
const clean=s=>String(s||'').replace(/[\s:：()（）]/g,'').toLowerCase();
const trim=s=>String(s||'').replace(/\r/g,'').trim();
const plain=n=>els(n,'t').map(t=>t.textContent).join('');
export const requirementFields={customerPart:'客戶產品料號',internalPart:'本公司料號',revision:'版本',project:'專案編號',quantity:'訂購數量（含 PCS／PNL 單位）',layers:'PCB 層數',material:'基材材質',thickness:'成品板厚與公差',temperature:'文件所列溫度條件',copper:'成品內外層銅厚',baseCopper:'原始基銅',finish:'表面處理與厚度單位',mask:'防焊顏色／材料',legend:'文字顏色／面別',size:'成品尺寸',panel:'拼板／成型',stackup:'層疊順序',via:'導通孔處理',ul:'UL 標記',origin:'產地文字',goldfinger:'金手指',reports:'測試／隨貨文件',impedance:'阻抗／線寬線距／公差',notes:'備註',special:'特殊製作與交付要求',stencil:'鋼板需求',gerber:'Gerber 版本與交付資訊',due:'需求交期',delivery:'交貨地點',inspection:'檢驗依據',marking:'LOGO／週期標記',process:'製程資訊'};
const mandatory=['customerPart','quantity','layers','material','thickness','copper','finish','mask','due'];
function field(r,key,value,evidence,ambiguity=''){
 value=trim(value);if(!value)return;
 if(r.fields[key]&&r.fields[key].value!==value){r.fields[key].blocked=true;r.fields[key].ambiguity='文件同欄位有不同內容，保留目前值';r.issues.push({id:'duplicate-'+key+'-'+r.issues.length,severity:'block',field:key,title:requirementFields[key]+'出現不同內容',detail:r.fields[key].value+' ／ '+value,source:evidence.locator});return;}
 if(r.fields[key]){r.fields[key].evidence.push(evidence);if(ambiguity&&!r.fields[key].ambiguity)r.fields[key].ambiguity=ambiguity;return;}
 r.fields[key]={key,label:requirementFields[key],value,evidence:[evidence],origin:'source',ambiguity,note:''};
}
function evidence(source,locator,raw){return {file:source,locator,raw};}
export function selectedOptions(value){
 // Read only explicitly selected options. Absence of a tick is not a negative answer.
 value=String(value||'').replace(/[\uE000-\uF8FF]/g,'□');const matches=[...value.matchAll(/[■☑☒█]\s*([^□■☑☒☐█]+)/g)];
 if(matches.length)return matches.map(m=>m[1].trim()).filter(Boolean).join('；');
 return /[□☐]/.test(value)?'':trim(value);
}
function paragraphText(p){
 return Array.from(p.getElementsByTagName('*')).filter(n=>['t','sym','br','tab'].includes(n.localName)).map(n=>{
  if(n.localName==='t')return n.textContent;
  if(n.localName==='sym'){const f=n.getAttributeNS(W,'font'),c=n.getAttributeNS(W,'char');if(f==='Wingdings 2'&&c==='F0CD')return ' × ';return `[未解析符號 ${f}/${c}]`;}
  return n.localName==='br'?'\n':' ';
 }).join('');
}
function fresh(name,type,text){return {id:crypto.randomUUID(),name,type,text,fields:{},issues:[],resolutions:{},approved:false,appliedSupplier:null,attachments:[],createdAt:new Date().toISOString()};}
function fromText(r){
 const semantic=semanticText(r.text).filter(c=>c.key!=='notes'),missing=new Set(semantic.filter(c=>!r.fields[c.key]).map(c=>c.key));
 for(const candidate of semantic)if(missing.has(candidate.key)){const value=selectedOptions(candidate.value);if(value)field(r,candidate.key,value,evidence(r.name,candidate.locator,candidate.value));}
 const rules={customerPart:/(?:^|\n)\s*(?:客戶料號|產品料號|料號)\s*[:：]\s*([^\s\n]+)/,internalPart:/(?:^|\n)\s*(?:銘星料號|本公司料號)\s*[:：]\s*([^\s\n]+)/,layers:/(?:^|\n)\s*(?:PCB層數|層數)\s*[:：]\s*(\d+\s*(?:Layer|層)?)/i,quantity:/(?:^|\n)\s*(?:PCB需求量|訂購數量)\s*[:：]\s*([^\n]+)/,material:/(?:^|\n)\s*(?:PCB原料材質|材料板厚|材料規格|材質)\s*[:：]\s*([^\n]+)/,thickness:/(?:^|\n)\s*(?:PCB厚度|成品板厚|板厚)\s*[:：]\s*([^\n]+)/,copper:/(?:^|\n)\s*(?:成品銅厚|銅厚)\s*[:：]\s*([^\n]+)/,finish:/(?:^|\n)\s*(?:PCB電鍍型態|表面處理)\s*[:：]\s*([^\n]+)/,mask:/(?:^|\n)\s*防焊\s*[:：]\s*([^\n]+)/,due:/(?:^|\n)\s*(?:需求日期|交期)\s*[:：]\s*([^\n]+)/};
 const extra={notes:'(?:重要|特別)?備註',legend:'文字(?:顏色|面別)?',inspection:'檢驗(?:標準|依據)',ul:'UL\\s*(?:MARK|文字|標記)',size:'(?:PCB尺寸大小|成品尺寸)',panel:'(?:排版成型|拼板|連板)',stackup:'(?:板層面順序|層疊順序|疊構)',via:'(?:導通孔處理|塞孔)',baseCopper:'(?:原始基銅|基銅)',gerber:'Gerber(?:\\s*Files|版本)?',marking:'(?:LOGO|週期標記)',impedance:'阻抗',reports:'(?:測試報告|隨貨文件)',revision:'版本',project:'專案編號',origin:'(?:產品產地文字|產地文字)',special:'(?:特殊需求附加說明|特殊需求)'};
 for(const [key,label]of Object.entries(extra))rules[key]=new RegExp('(?:^|\\n)\\s*(?:'+label+')\\s*[:：]\\s*([^\\n]+)','i');
 delete rules.notes;
 if(!r.fields.notes){const lines=r.text.split('\n'),at=lines.findIndex(line=>/^\s*(?:重要|特別)?備註\s*[:：]/.test(line));if(at>=0){const note=[lines[at].replace(/^\s*(?:重要|特別)?備註\s*[:：]\s*/,'')];for(const line of lines.slice(at+1)){if(!line.trim()||/^\s*[^：:\n]{1,32}[:：]/.test(line)&&!/^\s*(?:打叉方式|測試章|包裝方式)\s*[:：]/.test(line))break;note.push(line);}if(note.join('\n').trim())field(r,'notes',note.join('\n'),evidence(r.name,'備註段落',note.join('\n')));}}
 for(const[k,re]of Object.entries(rules))if(!r.fields[k])for(const m of r.text.matchAll(new RegExp(re.source,re.flags+'g'))){const v=selectedOptions(m[1]);if(v)field(r,k,v,evidence(r.name,'文字段落',m[0]));}
}
export async function readRequirementDocx(data,name){
 const p=await parseDocx(data),x=parseXml(p.xml),paras=els(x,'p',W).map(paragraphText),text=paras.join('\n');
 const checklist=/PCB\s*Layout\s*檢查表/i.test(text),r=fresh(name,checklist?'Layout 檢查表':/PCB製作申請單/.test(text)?'PCB 製作申請單':'Word 需求文件',text);
 const archive=await JSZip.loadAsync(data);r.images=await documentImages(archive,x);
 if(archive.file('word/comments.xml')){const cx=parseXml(await archive.file('word/comments.xml').async('string'));if(els(cx,'comment',W).length){r.issues.push({id:'document-comments',severity:'block',title:'Word 含有批註，需確認是否有額外需求',detail:els(cx,'comment',W).map(c=>els(c,'t',W).map(t=>t.textContent).join('')).join('\n'),source:'Word 批註'});}}
 if(els(x,'ins',W).length||els(x,'del',W).length)r.issues.push({id:'tracked-revisions',severity:'block',title:'Word 含有追蹤修訂',detail:'需確認接受哪些新增／刪除內容後，才能指定生產版本。',source:'Word 修訂標記'});
 const rows=els(x,'tr',W).map((row,i)=>({locator:`Word 表格列 ${i+1}`,cells:Array.from(row.children).filter(n=>n.localName==='tc').map(c=>els(c,'p',W).map(paragraphText).join('\n').trim())}));
 const nativeKeys=new Set();
 if(!checklist)for(const f of p.fields){const map={notes:'notes',customerPart:'customerPart',part:'internalPart',quantity:'quantity',material:'material',layers:'layers',thickness:'thickness',copper:'copper',baseCopper:'baseCopper',mask:'mask',legend:'legend',finish:'finish',inspection:'inspection',ul:'ul',size:'size',panel:'panel',stackup:'stackup',via:'via',impedance:'impedance',reports:'reports',revision:'revision',project:'project'};const key=f.key==='part'&&!/銘星|本公司|內部/.test(f.label)?null:map[f.key];if(key&&f.value){field(r,key,['mask','legend','finish'].includes(key)?selectedOptions(f.value):f.value,evidence(name,'Word 儲存格 '+(f.cell+1),f.value));nativeKeys.add(key);}}
 const aliases={notes:['備註','重要備註','特別備註'],customerPart:['產品料號','客戶料號'],internalPart:['銘星料號','本公司料號'],revision:['版本Version','版本'],project:['專案編號'],quantity:['PCB需求量','訂購數量'],layers:['PCB層數'],material:['PCB原料材質'],thickness:['PCB厚度'],temperature:['PCB溫度'],copper:['內、外層成品銅厚','銅厚'],finish:['PCB電鍍型態','表面處理'],mask:['防焊'],legend:['文字'],size:['PCB尺寸大小'],panel:['排版成型'],stackup:['板層面順序'],via:['PCBVIAHOLD'],ul:['UL文字','ULMARK'],origin:['產品產地文字'],goldfinger:['PCB含金手指'],reports:['文字需求'],special:['特殊需求附加說明'],stencil:['SMT加工鋼模'],gerber:['GerberFiles'],due:['需求日期'],delivery:['交貨地點'],inspection:['檢驗標準']};
 for(const row of rows){for(let i=0;i<row.cells.length-1;i++)for(const[key,names]of Object.entries(aliases))if(!nativeKeys.has(key)&&names.some(n=>clean(n)===clean(row.cells[i]))){const raw=row.cells[i+1],value=selectedOptions(raw);if(value)field(r,key,value,evidence(name,row.locator+`／儲存格 ${i+2}`,raw));}}
 const semanticCells=rows.flatMap((row,i)=>row.cells.map((value,j)=>({value,row:i,col:j}))),generic=semanticGrid(semanticCells),unmapped=new Set(generic.filter(c=>!r.fields[c.key]).map(c=>c.key));
 for(const candidate of generic)if(unmapped.has(candidate.key)){const value=selectedOptions(candidate.value);if(value)field(r,candidate.key,value,evidence(name,candidate.locator,candidate.value));}
 fromText(r);
 if(checklist){
  r.issues.push({id:'checklist-only',severity:'block',title:'Layout 檢查表不是完整製造規格',detail:'只提取明確的料號與層數。檢查項目的完成符號、設計通則與問句不視為本次製作要求；請補齊製作申請單／Gerber 規格。',source:'文件標題'});
  for(const key of Object.keys(r.fields))if(!['customerPart','internalPart','layers','project','revision'].includes(key))delete r.fields[key];
 }
 const special=r.fields.special;
 if(special){const lines=special.value.split('\n');const imp=lines.filter(s=>/阻抗|impedance|Differential|ohm/i.test(s));if(imp.length)field(r,'impedance',imp.join('\n'),special.evidence[0]);const delivery=special.value.match(/送貨到([^\n]+)/)?.[1];if(delivery&&!r.fields.delivery)field(r,'delivery',delivery,special.evidence[0]);}
 if(r.fields.panel){const n=r.fields.panel.value.match(/(?:連板[:：]?\s*)?(\d+)\s*連板/);if(n)r.panelCount=+n[1];}
 if(r.fields.quantity){const q=r.fields.quantity.value.match(/(\d+)\s*PCS[^\n]*?(\d+)\s*panel/i);if(q&&r.panelCount){r.quantityCheck={pcs:+q[1],panels:+q[2],perPanel:r.panelCount};if(+q[1]!==+q[2]*r.panelCount)r.issues.push({id:'panel-count',severity:'block',title:'PCS 與 Panel 數量不一致',detail:`${q[2]} Panel × ${r.panelCount} PCS/Panel ≠ ${q[1]} PCS`,source:r.fields.quantity.evidence[0].locator});}}
 if(r.fields.finish&&/\d\s*u\b/i.test(r.fields.finish.value)&&!/[μµ]|u[inm]|["”]/i.test(r.fields.finish.value))r.fields.finish.ambiguity='原文 u 未明示 μin 或 μm，請向客戶確認厚度單位。';
 if(r.fields.temperature)r.issues.push({id:'temperature-meaning',severity:'block',title:'溫度條件的意義需要確認',detail:'原標籤是 PCB溫度，不直接等同 Tg。請確認材料 Tg 或適用溫度要求，避免自行改寫。',source:r.fields.temperature.evidence[0].locator});
 if(r.fields.gerber)r.issues.push({id:'gerber-file',severity:'block',title:'Gerber 交付資訊不等於已核對 Gerber',detail:'目前只讀取需求文件，沒有驗證實際 Gerber／鑽孔檔與版次。請確認檔案包、料號及版本一致。',source:r.fields.gerber.evidence[0].locator});
 if(r.fields.impedance)r.issues.push({id:'impedance-stackup',severity:'block',title:'阻抗與疊構需工程確認',detail:'已保留阻抗值、公差、層別與線寬／線距原文；尚未計算疊構，也未證明板廠可依此製造。請取得疊構／阻抗確認。',source:r.fields.impedance.evidence[0].locator});
 if(r.fields.panel&&/附件|PDF/i.test(r.fields.panel.value))r.issues.push({id:'panel-attachment',severity:'block',title:'排版附件尚未核對',detail:'需求單引用排版圖，且可能同時含未勾選選項旁的檔名。請確認本次適用的排版圖與 V-Cut／Routing 工法。',source:r.fields.panel.evidence[0].locator});
 return normalizeIntake(r);
}
const colIndex=s=>[...s].reduce((v,c)=>v*26+c.charCodeAt(0)-64,0);
function resolvePath(base,target){const parts=(target.startsWith('/')?target.slice(1):base+'/'+target).split('/'),out=[];for(const p of parts){if(p==='..')out.pop();else if(p&&p!=='.')out.push(p);}return out.join('/');}
async function relationships(zip,path){const f=zip.file(path);if(!f)return {};const x=parseXml(await f.async('string'));return Object.fromEntries(els(x,'Relationship').map(n=>[n.getAttribute('Id'),n.getAttribute('Target')]));}
export async function readRequirementXlsx(data,name){
 const zip=await JSZip.loadAsync(data),wb=zip.file('xl/workbook.xml');if(!wb)throw Error('不是有效的 XLSX 活頁簿');const book=parseXml(await wb.async('string')),rels=await relationships(zip,'xl/_rels/workbook.xml.rels');
 const ss=zip.file('xl/sharedStrings.xml'),strings=ss?els(parseXml(await ss.async('string')),'si',S).map(plain):[];
 const r=fresh(name,'Excel 工程流程單','');r.sheets=[];const areas=els(book,'definedName',S).filter(n=>n.getAttribute('name')==='_xlnm.Print_Area');
 for(const [si,s]of els(book,'sheet',S).entries()){
  const title=s.getAttribute('name'),path=resolvePath('xl',rels[s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')]||'');if(!zip.file(path))continue;
  const x=parseXml(await zip.file(path).async('string')),reference=/加註|符號|範例|備忘錄|參考|^(?:其它|其他|圖庫)$/.test(title)||s.getAttribute('state')==='hidden';
  const area=areas.find(a=>+a.getAttribute('localSheetId')===si)?.textContent,maxColumn=area?.match(/:\$?([A-Z]+)\$?\d+/)?.[1];
  const cells=els(x,'c',S).map(c=>{const addr=c.getAttribute('r'),t=c.getAttribute('t'),v=els(c,'v',S)[0]?.textContent;let value=t==='s'?strings[+v]:t==='inlineStr'?plain(c):v||'';return {addr,value:trim(value),formula:els(c,'f',S)[0]?.textContent,reference:reference||(maxColumn&&colIndex(addr.match(/^[A-Z]+/)[0])>colIndex(maxColumn))};}).filter(c=>c.value||c.formula);
  const byAddr=Object.fromEntries(cells.map(c=>[c.addr,c]));const sheet={name:title,path,reference,cells,merges:els(x,'mergeCell',S).map(m=>m.getAttribute('ref')),shapes:[],lines:0};
  if(!reference)for(const c of cells.filter(c=>!c.reference)){const inline=c.value.match(/^(?:重要|特別)?備註\s*[:：]\s*([\s\S]+)$/);if(inline)field(r,'notes',inline[1],evidence(name,title+'!'+c.addr,c.value));else if(/^(?:重要|特別)?備註\s*[:：]?$/.test(c.value)){const row=c.addr.match(/\d+$/)[0],col=colIndex(c.addr.match(/^[A-Z]+/)[0]),next=cells.filter(n=>!n.reference&&n.addr.match(/\d+$/)[0]===row&&colIndex(n.addr.match(/^[A-Z]+/)[0])>col).sort((a,b)=>colIndex(a.addr.match(/^[A-Z]+/)[0])-colIndex(b.addr.match(/^[A-Z]+/)[0]))[0];if(next)field(r,'notes',next.value,evidence(name,title+'!'+c.addr+':'+next.addr,next.value));}}
  const prefix=path.slice(0,path.lastIndexOf('/')),filename=path.slice(path.lastIndexOf('/')+1),sr=await relationships(zip,prefix+'/_rels/'+filename+'.rels');
  for(const d of els(x,'drawing',S)){const target=sr[d.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')],dp=target&&resolvePath(prefix,target);if(!dp||!zip.file(dp))continue;const dx=parseXml(await zip.file(dp).async('string'));for(const anchor of Array.from(dx.documentElement.children)){const tx=els(anchor,'t',A).map(n=>n.textContent).join(''),from=els(anchor,'from')[0],row=Number(els(from||anchor,'row')[0]?.textContent)+1,col=Number(els(from||anchor,'col')[0]?.textContent)+1;const kind=els(anchor,'prstGeom')[0]?.getAttribute('prst');if(kind==='line')sheet.lines++;if(tx)sheet.shapes.push({text:tx,row,col,reference:reference||(maxColumn&&col>colIndex(maxColumn))});}}
  r.sheets.push(sheet);
  if(reference)continue;
  for(const c of cells.filter(c=>!c.reference&&c.formula))r.issues.push({id:'formula-'+si+'-'+c.addr,severity:'block',title:'公式快取尚未重算',detail:'此處僅讀取 Excel 儲存值，不執行公式；請以 Excel 重算後確認。',source:title+'!'+c.addr});
  readManufacturingCells(sheet,(key,value,addr,ambiguity='')=>field(r,key,value,evidence(name,title+'!'+addr,value),ambiguity),selectedOptions);
  const generic=semanticGrid(cells,{merges:sheet.merges}),missing=new Set(generic.filter(c=>!r.fields[c.key]).map(c=>c.key));
  for(const candidate of generic)if(missing.has(candidate.key)){const value=selectedOptions(candidate.value);if(value)field(r,candidate.key,value,evidence(name,title+'!'+candidate.locator,candidate.value));}
  for(const shape of sheet.shapes.filter(s=>!s.reference))for(const candidate of semanticText(shape.text,{inferQuantity:false}))if(!r.fields[candidate.key])field(r,candidate.key,candidate.value,evidence(name,title+'／文字方塊',shape.text));
  if(/流程表反面/.test(title)){
   const put=(key,value,addr,ambiguity='')=>{if(value)field(r,key,value,evidence(name,title+'!'+addr,value),ambiguity);};
   const markings=cells.filter(c=>/^([A-K])(3[2-6])$/.test(c.addr)).map(c=>c.value).join(' ');put('marking',markings,'A32:K36','LOGO、週期與文字有圖形覆蓋／標記，需確認適用項目。');
   const specials=['Y10','O15','O17','O20','R20','W23'].filter(a=>byAddr[a]?.value).map(a=>a+'：'+byAddr[a].value).join('\n');put('special',specials,'Y10,O15,O17,O20,R20,W23');
   if(sheet.lines)r.issues.push({id:'visual-process',severity:'block',title:'工單含圖形線段／劃除標記',detail:`${sheet.lines} 個線段可能表示停用、修改或流程標記；讀取儲存格無法判定意義。請開啟原 Excel 確認表面處理與 LOGO。`,source:title+'／繪圖物件'});
  }
 }
 r.text=r.sheets.map(s=>`【${s.name}${s.reference?'／參考區，不直接帶入':''}】\n`+s.cells.map(c=>`${c.addr}${c.reference?' [參考]':''}：${c.value}`).join('\n')+'\n'+s.shapes.map(sh=>`圖形(${sh.row},${sh.col})${sh.reference?' [參考]':''}：${sh.text}`).join('\n')).join('\n\n');
 r.issues.push({id:'xlsx-reference',severity:'info',title:'參考區與生產需求分開',detail:'符號／疊構範例頁及列印範圍外的油墨表只保留供核對，不自動指定為本次材料、層數或顏色。',source:'工作表與列印範圍'});
 if(!r.fields.quantity)r.issues.push({id:'xlsx-layout-quantity',severity:'info',title:'工單未列訂購數量，保留目前值',detail:'排版欄位 '+(r.fields.panel?.value||'未明示')+' 是裁板／每板資訊，不代表本次訂購數量。',source:'排版欄位'});
 return normalizeIntake(r);
}
export function readRequirementText(text,name='貼上文字',ocr=false){const r=fresh(name,ocr?'OCR／PDF 需求文字':'文字需求',text);fromText(r);if(ocr)r.issues.push({id:'ocr-verify',severity:'block',title:'需核對 OCR／PDF 的選項與數字',detail:'文字擷取可能遺失勾選框、圖形與版面，請對照原圖確認層數、單位、尺寸與選項。',source:name});return normalizeIntake(r);}
export function checkRequirement(r,template,now=new Date()){
 const findings=r.issues.filter(i=>i.id!=='panel-count');
 const q=r.fields.quantity?.value.match(/(\d+)\s*PCS[^\n]*?(\d+)\s*panel/i),pn=r.fields.panel?.value.match(/(?:連板[:：]?\s*)?(\d+)\s*連板/);
 r.quantityCheck=q&&pn?{pcs:+q[1],panels:+q[2],perPanel:+pn[1]}:null;
 if(r.quantityCheck&&+q[1]!==+q[2]*+pn[1])findings.push({id:'panel-count',severity:'block',title:'PCS 與 Panel 數量不一致',detail:`${q[2]} Panel × ${pn[1]} PCS/Panel ≠ ${q[1]} PCS`,field:'quantity',source:r.name});
 for(const key of mandatory)if(!r.fields[key]?.value?.trim())findings.push({id:'missing-'+key,severity:'block',title:'缺少'+requirementFields[key],detail:'請補入客戶確認值，並註明確認依據；不沿用供應商歷史規格。',field:key,source:r.name});
 for(const f of Object.values(r.fields)){
  if(f.ambiguity)findings.push({id:'ambiguity-'+f.key,severity:'block',title:f.label+'需要釐清',detail:f.ambiguity,field:f.key,source:f.evidence[0]?.locator||r.name});
  if(f.origin==='manual'&&!f.note?.trim())findings.push({id:'manual-'+f.key,severity:'block',title:f.label+'缺少修改依據',detail:'請填寫客戶回覆、工程確認或原文件位置。',field:f.key,source:'人工補充'});
 }
 if(r.fields.quantity?.value&&!/\b(?:PCS|PNL|PANEL|PANELS|SET)\b|片|套|個/i.test(r.fields.quantity.value))findings.push({id:'quantity-unit',severity:'block',title:'訂購數量沒有明確單位',detail:'請明示 PCS、Panel、PNL 或 SET，不自動換算。',field:'quantity',source:r.name});
 if(r.fields.thickness?.value&&!/(?:mm|mil|inch|in\b|英吋|吋|毫米|英寸|µm|μm|um)/i.test(r.fields.thickness.value))findings.push({id:'thickness-unit',severity:'block',title:'成品板厚缺少長度單位',detail:'請在規格值內明示 mm、mil 或 inch。',field:'thickness',source:r.name});
 if(r.fields.finish?.value&&/\d\s*u\b/i.test(r.fields.finish.value)&&!/[μµ]|u[inm]|["”]/i.test(r.fields.finish.value))findings.push({id:'finish-unit',severity:'block',title:'表面處理厚度單位不明',detail:'u 不能自行當成 μin 或 μm，請把確認單位寫入規格。',field:'finish',source:r.name});
 const due=r.fields.due?.value;if(due){const m=due.match(/(?:(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})|(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4}))/);if(m){const yy=+(m[1]||m[6]),mm=+(m[2]||m[4]),dd=+(m[3]||m[5]);if(new Date(yy,mm-1,dd)<new Date(now.getFullYear(),now.getMonth(),now.getDate()))findings.push({id:'historical-due',severity:'block',title:'需求交期已是歷史日期',detail:'原需求日期為 '+due+'，請確認這次實際交期，不能用今天日期自動代替。',field:'due',source:r.fields.due.evidence[0]?.locator});}else findings.push({id:'unclear-due',severity:'block',title:'交期缺少完整日期或格式不明',detail:'請確認包含年份的交貨日期與到貨／出貨定義。',field:'due',source:r.name});}
 if(template){
  const norm=s=>String(s).replace(/\s/g,'').toLowerCase();
  const compare=[['material',['material']],['copper',['copper']],['mask',['mask']],['finish',['finish','mask']],['inspection',['inspection']]];
  for(const [key,targets]of compare){const f=r.fields[key];if(!f?.value)continue;const old=template.fields.filter(t=>targets.includes(t.key)).map(t=>t.value).join('；');if(old&&norm(old)!==norm(f.value))findings.push({id:'template-'+template.id+'-'+key,severity:'block',title:f.label+'與舊模板內容不同',detail:`客戶：${f.value}\n模板：${old}\n請確認以本次客戶需求取代歷史規格。這是文字差異，尚不代表工程上的矛盾。`,source:template.filename});}
  findings.push({id:'template-terms-'+template.id,severity:'block',title:'確認模板中原有備品、包裝與製作條款',detail:'套入需求時會清空舊單的技術規格、備註與備品條款，再以已核對需求及採購人員填寫內容取代。請補齊本次需保留的供應商包裝、附件、備品與交期條款。',source:template.filename});
 }
 // A missing critical value cannot be waived by ticking a resolution.
 return findings.map(f=>({...f,resolution:r.resolutions[f.id],resolved:f.severity==='info'||(!f.id.startsWith('missing-')&&!f.id.startsWith('manual-')&&!!r.resolutions[f.id]?.note?.trim())}));
}
export function requirementTargets(template,key){
 const direct=key==='internalPart'?'part':key;
 const exact=template.fields.filter(t=>t.key===direct);if(exact.length)return exact;
 const combined={layers:['material'],thickness:['material'],legend:['mask'],finish:['mask'],baseCopper:['copper'],due:['quantity'],ul:['inspection','production'],gerber:['production'],process:['production'],origin:['production','inspection']};
 if(['ul','origin'].includes(key)){const re=key==='ul'?/UL\b/i:/Made\s+In|產地/i;return template.fields.filter(t=>(combined[key]||[]).includes(t.key)&&re.test(t.value));}
 return template.fields.filter(t=>(combined[key]||[]).includes(t.key));
}
export function requirementSummary(r){return Object.values(r.fields).filter(f=>f.value?.trim()).map(f=>({key:f.key,label:f.label,value:f.value,source:f.origin==='manual'?'人工確認：'+f.note:(f.evidence||[]).map(e=>e.file+'／'+e.locator).join('；')}));}
export function requirementMapping(r,template){
 const values={},groups=new Map();
 for(const [key,entry]of Object.entries(r.fields)){if(entry.blocked||!entry.value?.trim()||['delivery','internalPart'].includes(key)&&!entry.value)continue;for(const t of requirementTargets(template,key)){if(['date','delivery'].includes(t.key))continue;const items=groups.get(t.id)||[];items.push({key,value:entry.value,label:requirementFields[key]||entry.label});groups.set(t.id,items);}}
 for(const t of template.fields){const items=groups.get(t.id);if(!items)continue;values[t.id]=items.map(e=>{if(e.key===t.key||e.key==='internalPart'&&t.key==='part')return e.value;if(e.key==='layers')return /層|L(?:ayer)?s?$/i.test(e.value)?e.value:e.value+' 層';if(e.key==='due')return '交期：'+e.value;return e.label+'：'+e.value;}).join('\n');}
 values.__requirementAppendix=requirementSummary(r);return values;
}
