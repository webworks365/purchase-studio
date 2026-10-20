import assert from 'node:assert/strict';
import {readRequirementText,selectedOptions} from '../public-site/requirements.js';
import {readManufacturingCells} from '../public-site/xlsx-intake.js';
import {intakeValues,intakeReport} from '../public-site/order-layout.js';
const template={fields:[
 {id:'customerPart',key:'customerPart',label:'客戶料號',value:'CUSTOMER-DEFAULT'},
 {id:'part',key:'part',label:'銘星料號',value:'M-DEFAULT'},
 {id:'quantity',key:'quantity',label:'訂購數量',value:'20 PCS\n交期：2026/11/01'},
 {id:'material',key:'material',label:'材料板厚',value:'FR4 TG150 / 6L / 板厚1.6mm'},
 {id:'copper',key:'copper',label:'銅厚',value:'1 oz'},
 {id:'finish',key:'finish',label:'表面處理',value:'化金'},
 {id:'mask',key:'mask',label:'防焊',value:'綠色'},
 {id:'legend',key:'legend',label:'文字',value:'白色'}
]};
const base=()=>({...Object.fromEntries(template.fields.map(f=>[f.id,f.value])),__inlineDue:true,__dueDate:'2026-10-20'});
const r=readRequirementText('客戶料號：CLIENT-NEW\n銘星料號：M-NEW\n材料板厚：FR4 / 4L / 板厚1.6mm');
assert.equal(r.fields.thickness,undefined);assert.equal(intakeValues(r,template,base()).part,'M-NEW');
assert.equal(intakeValues(readRequirementText('客戶料號：CLIENT-ONLY'),template,base()).part,'M-DEFAULT');
console.log('PASS material label boundaries and customer/internal part distinction');
let current=base();const tolerances=readRequirementText('層數：4\n板厚：1.6±0.127mm');current=intakeValues(tolerances,template,current);const first=current.material;for(let i=0;i<10;i++)current=intakeValues(tolerances,template,current);assert.equal(current.material,first);assert.equal((current.material.match(/±/g)||[]).length,1);
const inch=readRequirementText('板厚：0.063" ± 0.006');current=intakeValues(inch,template,current);assert.equal(current.material,intakeValues(inch,template,current).material);
assert.equal(readRequirementText('板厚：0.063').fields.thickness.blocked,true);
console.log('PASS tolerances remain stable on repeat uploads and units are never inferred');
const corrupted={...base(),material:'FR4／PCB 層數：6／成品板厚與公差：1.6±1.6±0.062／PCB 層數：2'};
const complete=readRequirementText('材質：FR4\n層數：4\n板厚：0.063" ± 0.006');const repaired=intakeValues(complete,template,corrupted).material;assert.ok(repaired.includes('4L'));assert.ok(repaired.includes('0.063"'));assert.ok(!repaired.includes('PCB 層數'));assert.equal((repaired.match(/±/g)||[]).length,1);
console.log('PASS a complete source replaces previously accumulated material specifications');
const mixed=readRequirementText('材料板厚：FR4／PCB 層數：6／PCB 層數：2／成品板厚與公差：1.6±1.6±0.062');assert.equal(mixed.fields.layers.blocked,true);assert.equal(mixed.fields.thickness.blocked,true);const ambiguous=readRequirementText('銘星料號：M-FIRST\n銘星料號：M-SECOND');assert.equal(ambiguous.fields.internalPart.blocked,true);assert.equal(intakeValues(ambiguous,template,base()).part,'M-DEFAULT');
console.log('PASS conflicting source values preserve current fields');
let previous=intakeValues(readRequirementText('客戶料號：FIRST\n銅厚：2oz\n訂購數量：100 PCS\n交期：2026/12/01'),template,base());const next=intakeValues(readRequirementText('客戶料號：SECOND'),template,previous);assert.equal(next.copper,'1 oz');assert.ok(next.quantity.includes('2026/10/20'));assert.ok(!next.quantity.includes('2026/12/01'));
previous.quantity='100 PCS\n交期：手動指定月底';assert.ok(intakeValues(readRequirementText('客戶料號：SECOND'),template,previous).quantity.includes('手動指定月底'));
assert.equal(intakeReport(readRequirementText('這是一般說明'),template,base(),base()).recognized,0);
console.log('PASS new customer clears prior auto-filled values while retaining manual due edits');
const col=n=>{let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;};
for(const offset of [0,3,8]){
 const entries=[[4,1,'Company ( 4 ) 層 板 工 作 流 程 表'],[2,3,'料號'],[4,3,'CLIENT-WORK-001'],[15,4,'材質'],[17,4,'FR4'],[16,6,'成品'],[17,6,'尺寸'],[19,6,'4.8" * 12.5"'],[17,7,'厚度'],[19,7,'0.063" ± 0.006'],[17,8,'銅厚'],[19,8,'1 OZ'],[15,2,'排版'],[16,2,'X PNL'],[16,3,'4 PCS'],[5,9,'數量']];
 const sheet={name:'工單正面',lines:0,cells:entries.map(([c,r,value])=>({addr:col(c+offset)+r,value}))},fields={};readManufacturingCells(sheet,(key,value)=>fields[key]=value,selectedOptions);assert.equal(fields.customerPart,'CLIENT-WORK-001');assert.equal(fields.layers,'4');assert.equal(fields.thickness,'0.063" ± 0.006');assert.equal(fields.quantity,undefined);
}
assert.equal(selectedOptions('顏色：■ 綠 \uE007 其他：'),'綠');
console.log('PASS worksheet fields survive column shifts and panel PCS never becomes order quantity');
