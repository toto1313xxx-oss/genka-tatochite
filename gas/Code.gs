// r3 server candidate. Deploy only after client migration and staging tests.
const VERSION = '2026.11.26-r3';
const HEADERS = {
 '品目マスタ':['品目名','単価(千円)','単位'],
 '下見見積':['ID','日時','会場','ロット名','全体見積(円)','ステータス','落札額(円)','内訳JSON'],
 '仕入データ':['日時','会場','ロット名','品目名','単位','数量','総買値(円)','全体見積(円)','買値比率','個別見積(円)','確定原価(円)','符牒','ロットID','レコードID','取消日時'],
 '操作履歴':['日時','操作','対象ID','内容']
};
function sheet(name,create){
 const ss=SpreadsheetApp.getActiveSpreadsheet();
 let s=ss.getSheetByName(name);
 if(!s&&create)s=ss.insertSheet(name);
 if(s&&create){const h=HEADERS[name];if(s.getLastRow()===0)s.appendRow(h);else if(name==='仕入データ'){for(let i=12;i<h.length;i++){const c=s.getRange(1,i+1);if(!c.getValue())c.setValue(h[i]);else if(c.getValue()!==h[i])throw Error('Header mismatch');}}}
 return s;
}
function rows(s){return s&&s.getLastRow()>1?s.getDataRange().getValues().slice(1):[];}
function requireOk(v,m){if(!v)throw Error(m);}
function str(v){return String(v==null?'':v);}
function positive(v){return typeof v==='number'&&Number.isFinite(v)&&v>0;}
function yen(v){return positive(v)&&Number.isSafeInteger(v);}
function id(v){requireOk(typeof v==='string'&&/^[A-Za-z0-9_-]{1,120}$/.test(v),'Invalid ID');return v;}
function safeCell(v){const s=str(v);return /^[=+\-@]/.test(s)?"'"+s:s;}
function json(v){try{return JSON.parse(v);}catch(e){return [];}}
function response(v){return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON);}
function auth(p){const secret=PropertiesService.getScriptProperties().getProperty('API_TOKEN');requireOk(secret&&secret.length>=32&&p.token===secret,'Unauthorized');}
function audit(action,key,data){sheet('操作履歴',true).appendRow([new Date(),action,safeCell(key),JSON.stringify(data)]);}
function doGet(){return response({status:'error',message:'Use authenticated POST',version:VERSION});}
function doPost(e){
 const lock=LockService.getScriptLock();let held=false;
 try{
 const p=JSON.parse(e.postData.contents);auth(p);lock.waitLock(10000);held=true;
 const result=dispatch(p);SpreadsheetApp.flush();return response(Object.assign({status:'success'},result));
 }catch(e){return response({status:'error',message:str(e.message||e)});}finally{if(held)lock.releaseLock();}
}
function master(){return rows(sheet('品目マスタ',false)).filter(r=>r[0]).map(r=>({name:str(r[0]),rate:Number(r[1]),unit:['個','pcs'].includes(str(r[2]))?'pcs':'kg'}));}
function revision(items){return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(items)));}
function estimates(){return rows(sheet('下見見積',false)).filter(r=>r[0]).map(r=>({id:str(r[0]),date:str(r[1]),auction:str(r[2]),lotName:str(r[3]),totalEstYen:Number(r[4]),status:str(r[5])||'見積済',buyPriceYen:Number(r[6])||0,items:json(r[7])}));}
function records(lot){return rows(sheet('仕入データ',false)).filter(r=>str(r[12])===lot&&!r[14]).map(r=>({recordId:str(r[13]),date:str(r[0]),itemName:str(r[3]),unitType:r[4]==='kg'?'g':str(r[4]),amount:Number(r[5]),actualYen:Number(r[10]),cipher:str(r[11])}));}
function subtotal(i){requireOk(['kg','pcs','flat'].includes(i.unit)&&positive(i.rate)&&positive(i.amount)&&str(i.itemName).trim(),'Invalid item');requireOk(i.unit!=='pcs'||Number.isInteger(i.amount),'Invalid quantity');return i.rate*(i.unit==='kg'?i.amount/1000:i.unit==='pcs'?i.amount:1);}
function dispatch(p){
 switch(p.action){
 case 'getMaster':{const items=master();return {items,revision:revision(items)};}
 case 'getEstimates':return {estimates:estimates()};
 case 'getLotProgress':return {records:records(id(p.lotId))};
 case 'updateMaster':{
 const old=master();requireOk(p.revision===revision(old),'Master changed; reload');
 requireOk(Array.isArray(p.items)&&p.items.length>0,'Empty master');
 const names=new Set();const values=p.items.map(i=>{const n=str(i.name).trim();requireOk(n&&!names.has(n)&&positive(i.rate)&&['kg','pcs'].includes(i.unit),'Invalid master');names.add(n);return [safeCell(n),i.rate,i.unit==='pcs'?'個':'kg'];});
 const s=sheet('品目マスタ',true);audit(p.action,'master',old);const count=Math.max(s.getLastRow()-1,values.length);while(values.length<count)values.push(['','','']);s.getRange(2,1,count,3).setValues(values);return {revision:revision(master())};}
 case 'saveEstimate':{
 const key=id(p.id),s=sheet('下見見積',true);const old=rows(s).find(r=>str(r[0])===key);
 requireOk(Array.isArray(p.items)&&p.items.length&&str(p.lotName).trim()&&str(p.auction).trim(),'Invalid estimate');
 const items=p.items.map(i=>({...i,subtotalThousand:subtotal(i)}));const total=Math.round(items.reduce((a,i)=>a+i.subtotalThousand,0)*1000);
 requireOk(yen(total)&&total===p.totalEstYen,'Estimate mismatch');
 if(old){requireOk(Number(old[4])===total&&str(old[2])===p.auction&&str(old[3])===p.lotName&&JSON.stringify(json(old[7]))===JSON.stringify(items),'ID conflict');return {id:key,duplicate:true};}
 s.appendRow([key,new Date(),safeCell(p.auction),safeCell(p.lotName),total,'見積済',0,JSON.stringify(items)]);return {id:key};}
 case 'saveCostRecord':{
 const key=id(p.recordId),s=sheet('仕入データ',true);
 requireOk(yen(p.buyPriceYen)&&yen(p.lotEstimateYen)&&yen(p.indivEstYen)&&positive(p.amount)&&str(p.itemName).trim()&&str(p.auction).trim()&&str(p.lotName).trim(),'Invalid cost');
 requireOk(['kg','個','一式'].includes(p.unitType)&&(p.unitType!=='個'||Number.isInteger(p.amount)),'Invalid unit');
 const actual=Math.round(p.indivEstYen*p.buyPriceYen/p.lotEstimateYen);requireOk(actual===p.actualYen,'Cost mismatch');
 const old=rows(s).find(r=>str(r[13])===key);
 if(old){requireOk(!old[14]&&str(old[12])===str(p.lotId)&&str(old[3])===p.itemName&&Number(old[5])===p.amount&&Number(old[10])===actual,'ID conflict');return {recordId:key,duplicate:true};}
 if(p.lotId){const lot=estimates().find(i=>i.id===id(p.lotId));requireOk(lot&&lot.status==='落札'&&lot.buyPriceYen===p.buyPriceYen&&lot.totalEstYen===p.lotEstimateYen&&lot.lotName===p.lotName&&lot.auction===p.auction,'Lot changed');}
 s.appendRow([new Date(),safeCell(p.auction),safeCell(p.lotName),safeCell(p.itemName),p.unitType,p.amount,p.buyPriceYen,p.lotEstimateYen,(p.buyPriceYen/p.lotEstimateYen*100).toFixed(1)+'%',p.indivEstYen,actual,safeCell(p.cipher),str(p.lotId),key,'']);return {recordId:key};}
 case 'deleteCostRecord':{
 const key=id(p.recordId),s=sheet('仕入データ',true),a=rows(s),i=a.findIndex(r=>str(r[13])===key);requireOk(i>=0,'Not found');if(!a[i][14]){audit(p.action,key,a[i]);s.getRange(i+2,15).setValue(new Date());}return {recordId:key};}
 case 'updateEstimateStatus':{
 const key=id(p.id),s=sheet('下見見積',false),a=rows(s),i=a.findIndex(r=>str(r[0])===key);requireOk(i>=0,'Not found');const r=a[i],old=str(r[5])||'見積済',next=p.status;
 const allowed={'見積済':['落札','不落'],'落札':['完了','不落'],'不落':['落札','削除'],'完了':['落札','削除']};
 requireOk(old===next||(allowed[old]||[]).includes(next),'Invalid transition');
 const buy=next==='落札'?(p.buyPriceYen===undefined?Number(r[6]):p.buyPriceYen):Number(r[6]);
 if(next==='落札')requireOk(yen(buy),'Invalid buy price');
 if(records(key).length)requireOk(next!=='不落'&&buy===Number(r[6]),'Existing cost records');
 if(next==='完了')requireOk(records(key).length>0,'No cost records');
 audit(p.action,key,{before:old,after:next,buy});s.getRange(i+2,6,1,2).setValues([[next,buy]]);return {id:key,status:next,buyPriceYen:buy};}
 default:throw Error('Unknown action');
 }
}
