const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',REL='http://schemas.openxmlformats.org/package/2006/relationships';
const xml=s=>new DOMParser().parseFromString(s,'application/xml');
const mime={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',bmp:'image/bmp',tif:'image/tiff',tiff:'image/tiff',emf:'image/x-emf',wmf:'image/x-wmf'};
const hash=async bytes=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))).map(v=>v.toString(16).padStart(2,'0')).join('');
const identity=/logo|signature|sign[-_ ]?off|簽名|签名|簽核|签核|簽章|签章|印章|用印|核准|審核|审核|公司標誌|公司标志|商標|商标/i;
export const isAllowedImage=image=>!!image&&!image.excluded&&(image.origin==='upload'||image.kind==='requirement'&&image.policyVersion===2)&&!identity.test(image.description||'');
const ancestor=(node,name)=>{while(node&&!(node.namespaceURI===W&&node.localName===name))node=node.parentNode;return node;};
const words=node=>node?Array.from(node.getElementsByTagNameNS(W,'t')).map(t=>t.textContent).join(''):'';
const signing=/簽名|签名|簽核|签核|簽章|签章|簽字|签字|核准|核準|審核|审核|主管|處理人員|处理人员|製表者|制表者|廠商回簽|approvedby|checkedby|preparedby|signature|signoff/i;
function base64(bytes){let text='';for(let i=0;i<bytes.length;i+=16384)text+=String.fromCharCode(...bytes.subarray(i,i+16384));return btoa(text);}
export async function imageFile(file){
 let bytes=new Uint8Array(await file.arrayBuffer()),ext=file.name.split('.').pop().toLowerCase(),width,height;
 const bitmap=await createImageBitmap(file);width=bitmap.width;height=bitmap.height;
 if(ext==='webp'){const canvas=document.createElement('canvas');canvas.width=width;canvas.height=height;canvas.getContext('2d').drawImage(bitmap,0,0);const blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));bytes=new Uint8Array(await blob.arrayBuffer());ext='png';}bitmap.close();
 return {id:await hash(bytes),data:base64(bytes),ext,mime:mime[ext],cx:width*9525,cy:height*9525,origin:'upload',description:file.name,excluded:identity.test(file.name)};
}
export async function documentImages(zip,document){
 const part=zip.file('word/_rels/document.xml.rels');if(!part)return [];
 const rels=Array.from(xml(await part.async('string')).getElementsByTagNameNS(REL,'Relationship')),images=[];
 const nodes=Array.from(document.getElementsByTagName('*')).filter(n=>n.localName==='blip'||n.localName==='imagedata');
 for(const node of nodes){const id=node.getAttributeNS(R,'embed')||node.getAttributeNS(R,'id'),rel=rels.find(r=>r.getAttribute('Id')===id);if(!rel||rel.getAttribute('TargetMode')==='External')continue;
  const target=rel.getAttribute('Target');const parts=(target.startsWith('/')?target.slice(1):'word/'+target).split('/'),path=[];for(const p of parts){if(p==='..')path.pop();else if(p&&p!=='.')path.push(p);}const file=zip.file(path.join('/'));if(!file)continue;
  const ext=target.split('.').pop().toLowerCase();if(!mime[ext])continue;const bytes=await file.async('uint8array'),key=await hash(bytes);if(images.some(i=>i.id===key))continue;
  let parent=node;while(parent&&parent.localName!=='drawing'&&parent.localName!=='pict')parent=parent.parentNode;const extent=parent?.getElementsByTagNameNS('*','extent')[0];
  const cell=ancestor(node,'tc'),para=ancestor(node,'p'),row=ancestor(node,'tr');
  const metadata=Array.from(parent?.getElementsByTagName('*')||[]).filter(n=>['docPr','cNvPr'].includes(n.localName)).map(n=>['name','descr','title'].map(k=>n.getAttribute(k)||'').join(' ')).join(' ');
  // Read the enclosing Word paragraph for inline, floating and VML images.
  // A preceding table is not a caption: it can contain the whole specification.
  const previous=para?.previousElementSibling,caption=previous?.namespaceURI===W&&previous.localName==='p'?words(previous):'';
  const local=[words(para),words(cell),words(row)].join(' '),context=local+' '+caption,description=target+' '+metadata;
  if(signing.test(local.replace(/\s/g,''))||identity.test(description)||identity.test(context)||!/規格|製作|制作|板厚|銅厚|铜厚|阻抗|拼板|排版|成型|品質標準|品质标准|檢驗重點|检验重点|出貨附件|出货附件|注意事項|注意事项|包裝|包装|specification|stackup|impedance|panel|fabrication/i.test(context+' '+metadata))continue;
  images.push({id:key,data:base64(bytes),ext,mime:mime[ext],cx:Number(extent?.getAttribute('cx'))||5000000,cy:Number(extent?.getAttribute('cy'))||2500000,origin:'document',kind:'requirement',policyVersion:2,description});
 }return images;
}
export function combineImages(...groups){return [...new Map(groups.flat().filter(isAllowedImage).map(i=>[i.id,i])).values()];}
export async function upgradeRequirementImages(state){
 if(state.requirementImagePolicyVersion===2)return false;
 const verified=new Map();
 for(const r of state.requirements||[]){
  if(r.sourceData&&/\.docx$/i.test(r.name||'')){
   try{const zip=await JSZip.loadAsync(r.sourceData,{base64:true}),part=zip.file('word/document.xml');r.images=part?await documentImages(zip,xml(await part.async('string'))):[];}catch{r.images=combineImages(r.images||[]);}
  }else r.images=combineImages(r.images||[]);
  for(const image of r.images)verified.set(image.id,image);
 }
 for(const draft of [...Object.values(state.drafts||{}),...(state.orders||[])])if(draft.values?.__images)draft.values.__images=combineImages(draft.values.__images.map(i=>verified.get(i.id)||i));
 state.requirementImagePolicyVersion=2;return true;
}
export async function insertImages(zip,doc,template,images){
 if(!images?.length)return;
 const existing=new Set();for(const file of Object.values(zip.files).filter(f=>/^word\/media\//.test(f.name)&&!f.dir))existing.add(await hash(await file.async('uint8array')));
 const pending=images.filter(i=>{if(!isAllowedImage(i)||existing.has(i.id))return false;existing.add(i.id);return true;});if(!pending.length)return;
 const field=template.fields.find(f=>f.key==='notesTail')||template.fields.find(f=>f.key==='notes');if(!field)throw Error('此模板沒有可放置原圖的備註格，請先在 Word 模板保留圖片位置');
 const cell=doc.getElementsByTagNameNS(W,'tc')[field.cell],paragraphs=Array.from(cell.children).filter(n=>n.localName==='p'),p=paragraphs.at(-1);if(!p)throw Error('找不到原圖放置位置');
 const path='word/_rels/document.xml.rels',rels=xml(zip.file(path)?await zip.file(path).async('string'):'<Relationships xmlns="'+REL+'"/>'),types=xml(await zip.file('[Content_Types].xml').async('string'));
 const width=Number(cell.getElementsByTagNameNS(W,'tcW')[0]?.getAttributeNS(W,'w'))*635||5200000;
 let id=Math.max(0,...Array.from(doc.getElementsByTagNameNS('*','docPr')).map(n=>Number(n.getAttribute('id'))||0));
 for(const image of pending){const stem='purchase-image-'+image.id,filename=stem+'.'+image.ext,rid='rIdPurchaseImage'+(++id);const rel=rels.createElementNS(REL,'Relationship');rel.setAttribute('Id',rid);rel.setAttribute('Type',R+'/image');rel.setAttribute('Target','media/'+filename);rels.documentElement.appendChild(rel);zip.file('word/media/'+filename,image.data,{base64:true});
  if(!Array.from(types.documentElement.children).some(n=>n.localName==='Default'&&n.getAttribute('Extension')===image.ext)){const entry=types.createElementNS(types.documentElement.namespaceURI,'Default');entry.setAttribute('Extension',image.ext);entry.setAttribute('ContentType',image.mime);types.documentElement.appendChild(entry);}
  const scale=Math.min(1,width/image.cx),cx=Math.round(image.cx*scale),cy=Math.round(image.cy*scale);
  const fragment=xml('<w:r xmlns:w="'+W+'" xmlns:r="'+R+'" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:br/><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="'+cx+'" cy="'+cy+'"/><wp:docPr id="'+id+'" name="Picture '+id+'"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="'+id+'" name="Picture '+id+'"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="'+rid+'"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+cx+'" cy="'+cy+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>');p.appendChild(doc.importNode(fragment.documentElement,true));
 }
 zip.file(path,new XMLSerializer().serializeToString(rels));zip.file('[Content_Types].xml',new XMLSerializer().serializeToString(types));
}
