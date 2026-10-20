// Run in a browser with JSZip loaded and public-site as the document base URL.
(async()=>{
const core=await import('./core.js'),pics=await import('./requirement-images.js'),req=await import('./requirements.js'),layout=await import('./order-layout.js');
const assert=(ok,message)=>{if(!ok)throw Error(message)},W=core.W,R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aT1cAAAAASUVORK5CYII=';
const picture=id=>`<w:r><w:drawing><wp:inline><wp:extent cx="5000000" cy="1000000"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData><a:blip r:embed="${id}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const wrap=body=>`<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><w:body>${body}</w:body></w:document>`;
const zip=new JSZip();zip.file('word/media/image1.png',png,{base64:true});zip.file('word/_rels/document.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="media/image1.png" Type="'+R+'/image"/></Relationships>');
zip.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/></Types>');
// A whole preceding table used to falsely qualify signatures as specification images.
const signature=core.parseXml(wrap('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>PCB製作規格 板厚 銅厚</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>研發主管審核： 專案工程師：</w:t></w:r>'+picture('rId1')+'</w:p>'));
assert((await pics.documentImages(zip,signature)).length===0,'signature after specification table must not be imported');
const specXml=wrap('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>品質標準與出貨附件</w:t></w:r>'+picture('rId1')+'</w:p></w:tc></w:tr></w:tbl>'),spec=core.parseXml(specXml),images=await pics.documentImages(zip,spec);
assert(images.length===1&&images[0].data===png&&pics.isAllowedImage(images[0]),'actual specification image bytes retained');
const legacy={...images[0]};delete legacy.policyVersion;
assert(!pics.isAllowedImage(legacy),'unverified old document images must not leak into exports');
zip.file('word/document.xml',specXml);
const state={requirements:[{name:'spec.docx',sourceData:await zip.generateAsync({type:'base64'}),images:[legacy]}],drafts:{a:{values:{__images:[legacy,{...legacy,id:'signature'}]}}},orders:[{values:{__images:[{...legacy,id:'signature'}]}}]};
assert(await pics.upgradeRequirementImages(state),'migration should run once');assert(state.drafts.a.values.__images.length===1&&state.orders[0].values.__images.length===0,'migration rechecks original sources and removes old signatures');assert(!await pics.upgradeRequirementImages(state),'migration is stable on repeat');
const xml=wrap('<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:line="288"/></w:pPr><w:r><w:rPr><w:sz w:val="26"/></w:rPr><w:t>old</w:t><w:br/><w:t>line</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');zip.file('word/document.xml',xml);
const template={data:await zip.generateAsync({type:'base64'}),fields:[{id:'copper',key:'copper',cell:0,value:'oldline'}]},before=core.parseXml(xml);
let blob=await core.exportDocx(template,{copper:'1 OZ\n原始基銅：內層 1 / 1 OZ'},'MH26A-10-001'),out=await JSZip.loadAsync(await blob.arrayBuffer()),doc=core.parseXml(await out.file('word/document.xml').async('string'));
assert(doc.getElementsByTagNameNS(W,'br').length===1,'new line breaks replace former breaks instead of accumulating blank lines');
for(const tag of ['tblPr','tblGrid','trPr','tcPr','pPr','rPr','sectPr'])assert(JSON.stringify([...before.getElementsByTagNameNS(W,tag)].map(n=>n.outerHTML))===JSON.stringify([...doc.getElementsByTagNameNS(W,tag)].map(n=>n.outerHTML)),'original '+tag+' unchanged');
const mapped=req.requirementMapping({fields:{copper:{value:'1        OZ'},baseCopper:{value:'內層：1   /   1 OZ；外層：H   /  H OZ'}}},template);
assert(mapped.copper==='1 OZ\n原始基銅：內層：1 / 1 OZ；外層：H / H OZ','Excel padding must not become Word spacing');
const draft={values:{copper:mapped.copper.replace('1 OZ','1       OZ')},createdAt:new Date().toISOString()};layout.normalizeDraft(template,draft);assert(!/ {2}/.test(draft.values.copper),'stored manufacturing values also normalized');
return {signatureContextExcluded:true,originalSpecificationImagesRetained:true,legacyDraftsAndRecordsCleaned:true,lineBreaksDoNotAccumulate:true,sourcePaddingRemoved:true,originalWordPropertiesPreserved:true};
})()
