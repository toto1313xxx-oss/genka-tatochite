/**
 * Fushimidou r3 / Apps Script V8
 * Append-only journal. No existing sheet is overwritten.
 * Run setupFushimidou once in a spreadsheet-bound Apps Script.
 */
const FS = {sheet:'伏見堂_r3_操作台帳', version:'2026.11.27-r3', maxPayload:40000};
const HEAD = ['サーバー日時','操作ID','操作','担当区分','対象ID','山ID','会場','山名','品目','数量','単位','金額(円)','状態','処理内容JSON'];

function setupFushimidou() {
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    if(!ss) throw Error('スプレッドシートに紐付いたApps Scriptで実行してください');
    const props=PropertiesService.getScriptProperties();
    const existing=props.getProperty('FS_SPREADSHEET_ID');
    if(existing && existing!==ss.getId()) throw Error('既存接続先と異なります');
    let sh=ss.getSheetByName(FS.sheet);
    if(!sh) {
      sh=ss.insertSheet(FS.sheet);
      sh.getRange(1,1,1,HEAD.length).setValues([HEAD]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,HEAD.length).setFontWeight('bold').setBackground('#263238').setFontColor('#ffffff');
    }
    checkHeader_(sh);
    props.setProperty('FS_SPREADSHEET_ID',ss.getId());
    if(!props.getProperty('FS_MANAGER_TOKEN')) props.setProperty('FS_MANAGER_TOKEN',makeToken_());
    if(!props.getProperty('FS_STAFF_TOKEN')) props.setProperty('FS_STAFF_TOKEN',makeToken_());
    return '設定完了。プロジェクトの設定 → スクリプト プロパティで各トークンを確認してください。';
  } finally {lock.releaseLock();}
}
function makeToken_(){return Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');}
function checkHeader_(sh){
  if(JSON.stringify(sh.getRange(1,1,1,HEAD.length).getValues()[0])!==JSON.stringify(HEAD))throw Error('台帳の列構成が変更されています');
}
function journal_(){
  const id=PropertiesService.getScriptProperties().getProperty('FS_SPREADSHEET_ID');
  if(!id)throw Error('初期設定が必要です');
  const sh=SpreadsheetApp.openById(id).getSheetByName(FS.sheet);
  if(!sh)throw Error('台帳が見つかりません');
  checkHeader_(sh);return sh;
}
function equalToken_(a,b){
  if(typeof a!=='string'||typeof b!=='string'||a.length!==b.length||a.length<32)return false;
  let diff=0;for(let i=0;i<a.length;i++)diff|=a.charCodeAt(i)^b.charCodeAt(i);return diff===0;
}
function role_(token){
  const p=PropertiesService.getScriptProperties();
  if(equalToken_(token,p.getProperty('FS_MANAGER_TOKEN')))return 'manager';
  if(equalToken_(token,p.getProperty('FS_STAFF_TOKEN')))return 'staff';
  throw Error('認証に失敗しました');
}
function output_(data,callback){
  const json=JSON.stringify(data).replace(/[\u2028\u2029]/g,c=>'\\u'+c.charCodeAt(0).toString(16));
  if(callback) {
    if(!/^gas_cb_[A-Za-z0-9_]{1,100}$/.test(callback))return ContentService.createTextOutput('/* invalid callback */').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function withLock_(fn){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))throw Error('混雑中です。結果を再確認してください');
  try{return fn();}finally{lock.releaseLock();}
}
function doGet(e){
  const p=e&&e.parameter||{};
  try{
    const role=role_(p.token);
    const result=withLock_(()=>read_(load_(journal_()),p,role));
    return output_(Object.assign({status:'success'},result),p.callback);
  }catch(err){return output_({status:'error',message:String(err.message||err)},p.callback);}
}
function doPost(e){
  try{
    const body=e&&e.postData&&e.postData.contents||'';
    if(body.length>FS.maxPayload+2000)throw Error('送信内容が大きすぎます');
    const p=JSON.parse(body);const role=role_(p.token);delete p.token;
    const result=withLock_(()=>{
      const sh=journal_();
      try{return write_(sh,p,role);}
      catch(err){
        // Record validation failures only when the operation is demonstrably absent.
        const op=id_(p.operationId,'操作ID');
        const state=load_(sh);
        if(state.ops.has(op))throw err;
        const at=new Date().toISOString();
        const event={operationId:op,action:'rejected',at,data:{action:p.action,message:String(err.message||err).slice(0,500)},request:''};
        sh.appendRow([at,op,'rejected',role,'','','','','','','','','拒否',JSON.stringify(event)]);
        SpreadsheetApp.flush();
        return {operationId:op,rejected:true};
      }
    });
    return output_(Object.assign({status:'success'},result));
  }catch(err){return output_({status:'error',message:String(err.message||err)});}
}
function load_(sh){
  const s={est:new Map(),cost:new Map(),master:null,masterRevision:'',ops:new Map()};
  const n=sh.getLastRow();if(n<2)return s;
  for(const row of sh.getRange(2,1,n-1,HEAD.length).getValues()){
    if(!row[1]||!row[13])throw Error('台帳に不完全な行があります');
    const event=JSON.parse(row[13]);const op=String(row[1]);
    if(s.ops.has(op))throw Error('台帳に重複操作IDがあります');
    apply_(s,event);s.ops.set(op,{action:event.action,role:String(row[3]),event,request:event.request});
  }
  return s;
}
function apply_(s,e){
  const d=e.data;
  switch(e.action){
    case 'saveEstimate': s.est.set(d.id,d);break;
    case 'saveCostRecord':s.cost.set(d.recordId,d);break;
    case 'deleteCostRecord':s.cost.set(d.recordId,Object.assign({},s.cost.get(d.recordId),{deleted:true,deletedAt:e.at}));break;
    case 'updateEstimateStatus':s.est.set(d.id,Object.assign({},s.est.get(d.id),d));break;
    case 'updateMaster':s.master=d.items;s.masterRevision=e.operationId;break;
    case 'rejected':break;
    default:throw Error('不明な台帳操作');
  }
}
function read_(s,p,role){
  switch(p.action){
    case 'getMaster':return {items:s.master||[],initialized:s.master!==null,revision:s.masterRevision,role,protocol:'fushimidou-r3',version:FS.version};
    case 'getEstimates':return {estimates:Array.from(s.est.values()),role};
    case 'getLotProgress':return {records:Array.from(s.cost.values()).filter(r=>!r.deleted&&r.lotId===String(p.lotId||'')),role};
    case 'getOperation':{
      const op=s.ops.get(String(p.operationId||''));
      const own=op&&op.role===role;
      if(own&&op.action==='rejected')return {committed:false,rejected:true,action:op.event.data.action,message:op.event.data.message};
      return {committed:!!own,action:own?op.action:null};
    }
    default:throw Error('未対応の取得操作');
  }
}
function text_(v,label,max=200,empty=false){
  if(typeof v!=='string')throw Error(label+'が不正です');
  const t=v.trim();if((!empty&&!t)||t.length>max)throw Error(label+'が不正です');return t;
}
function id_(v,label='ID'){
  const t=text_(v,label,100);if(!/^[A-Za-z0-9_-]+$/.test(t))throw Error(label+'が不正です');return t;
}
function num_(v,label,integer=false,zero=false){
  if(typeof v!=='number'||!Number.isFinite(v)||(zero?v<0:v<=0)||v>1e12||(integer&&!Number.isSafeInteger(v)))throw Error(label+'が不正です');return v;
}
function manager_(role){if(role!=='manager')throw Error('責任者権限が必要です');}
function subtotal_(rate,amount,unit){
  if(!['kg','pcs','flat'].includes(unit))throw Error('単位が不正です');
  num_(rate,'単価');num_(amount,'数量',unit==='pcs'||unit==='flat');
  if(unit==='flat'&&amount!==1)throw Error('一式数量が不正です');
  return unit==='kg'?rate*(amount/1000):unit==='pcs'?rate*amount:rate;
}
function liveCost_(s,lotId){return Array.from(s.cost.values()).filter(r=>!r.deleted&&r.lotId===lotId);}
function canonical_(v){
  if(Array.isArray(v))return '['+v.map(canonical_).join(',')+']';
  if(v&&typeof v==='object')return '{'+Object.keys(v).sort().map(k=>JSON.stringify(k)+':'+canonical_(v[k])).join(',')+'}';
  return JSON.stringify(v);
}
function write_(sh,p,role){
  const op=id_(p.operationId,'操作ID');
  const s=load_(sh);
  const request=canonical_(p);
  if(s.ops.has(op)){
    const old=s.ops.get(op);
    if(old.role!==role||old.request!==request)throw Error('操作IDが競合しています');
    return {operationId:op,duplicate:true};
  }
  const at=new Date().toISOString();
  let d;
  switch(p.action){
    case 'saveEstimate':{
      const id=id_(p.id);
      if(s.est.has(id))throw Error('見積IDが重複しています');
      if(!Array.isArray(p.items)||!p.items.length||p.items.length>300)throw Error('見積内訳を確認してください');
      const items=p.items.map(i=>{
        const unit=i.unit,amount=num_(i.amount,'数量'),rate=num_(i.rate,'単価');
        return {itemName:text_(i.itemName,'品目'),unit,amount,rate,subtotalThousand:subtotal_(rate,amount,unit)};
      });
      const totalEstYen=Math.round(items.reduce((a,i)=>a+i.subtotalThousand,0)*1000);
      num_(totalEstYen,'見積総額',true);
      if(totalEstYen!==p.totalEstYen)throw Error('見積総額が一致しません');
      const limitRate=p.limitRate==null?null:num_(p.limitRate,'掛け率');
      if(limitRate!==null&&limitRate>10)throw Error('掛け率が範囲外です');
      d={id,date:at,auction:text_(p.auction,'会場'),lotName:text_(p.lotName,'山名'),totalEstYen,items,limitRate,limitYen:limitRate===null?null:Math.round(totalEstYen*limitRate),status:'下見中',buyPriceYen:0};
      break;
    }
    case 'saveCostRecord':{
      const recordId=id_(p.recordId);
      if(s.cost.has(recordId))throw Error('原価IDが重複しています');
      const lotId=p.lotId?id_(p.lotId):'';
      const auction=text_(p.auction,'会場'),lotName=text_(p.lotName,'山名'),itemName=text_(p.itemName,'品目');
      const buyPriceYen=num_(p.buyPriceYen,'総買値',true),lotEstimateYen=num_(p.lotEstimateYen,'全体見積',true);
      if(lotId){
        const lot=s.est.get(lotId);
        if(!lot||lot.status!=='落札')throw Error('落札済みの山だけ登録できます');
        if(lot.buyPriceYen!==buyPriceYen||lot.totalEstYen!==lotEstimateYen||lot.auction!==auction||lot.lotName!==lotName)throw Error('山情報が更新されました。再呼出してください');
      }
      const unitType=p.unitType,unit=({'kg':'kg','個':'pcs','一式':'flat'})[unitType];
      const amount=num_(p.amount,'数量'),rate=num_(p.rate,'単価');
      const indiv=subtotal_(rate,amount,unit);
      const indivEstYen=Math.round(indiv*1000);
      const actualYen=Math.round(indiv*(buyPriceYen/lotEstimateYen)*1000);
      if(indivEstYen!==p.indivEstYen||actualYen!==p.actualYen)throw Error('原価計算が一致しません');
      num_(actualYen,'原価',true);if(actualYen>9999999)throw Error('原価範囲外');
      const cipher=text_(p.cipher,'符牒',4);
      if(actualYen>=1000&&decode_(cipher)<actualYen)throw Error('符牒の金額が原価未満です');
      if(actualYen<1000&&cipher!=='----')throw Error('符牒が不正です');
      d={recordId,date:at,auction,lotName,itemName,unitType,amount,rate,buyPriceYen,lotEstimateYen,ratioPercent:(buyPriceYen/lotEstimateYen*100).toFixed(1)+'%',indivEstYen,actualYen,cipher,lotId,deleted:false};break;
    }
    case 'deleteCostRecord':{
      manager_(role);
      const r=s.cost.get(id_(p.recordId));
      if(!r||r.deleted)throw Error('有効な原価登録がありません');
      if(r.lotId&&s.est.get(r.lotId)?.status!=='落札')throw Error('完了済みの山は差戻し後に取消してください');
      d={recordId:r.recordId,lotId:r.lotId};break;
    }
    case 'updateEstimateStatus':{
      const id=id_(p.id),lot=s.est.get(id);
      if(!lot)throw Error('山がありません');
      const status=p.status;
      if(status==='完了'){
        if(lot.status!=='落札'||!liveCost_(s,id).length)throw Error('落札済みの登録原価が必要です');
        const remaining=lot.buyPriceYen-liveCost_(s,id).reduce((a,r)=>a+r.actualYen,0);
        if(p.expectedRemaining!==remaining)throw Error('残額が変更されました。再確認してください');
        d={id,status,remainingYen:remaining,completedAt:at};
      }else{
        manager_(role);
        const allowed={'下見中':['落札','不落'],'落札':['落札','不落'],'不落':['落札','削除'],'完了':['落札','削除']};
        if(!(allowed[lot.status]||[]).includes(status)||p.expectedStatus!==lot.status)throw Error('状態が変更されました。再取得してください');
        if(status==='不落'&&liveCost_(s,id).length)throw Error('原価登録のある山は不落にできません');
        d={id,status};
        if(status==='落札'){
          d.buyPriceYen=num_(p.buyPriceYen,'落札額',true);
          if(liveCost_(s,id).length&&d.buyPriceYen!==lot.buyPriceYen)throw Error('原価登録後の買値変更はできません');
        }
      }break;
    }
    case 'updateMaster':{
      manager_(role);
      if(p.expectedRevision!==s.masterRevision)throw Error('マスタが変更されました。再取得してください');
      if(!Array.isArray(p.items)||!p.items.length||p.items.length>300)throw Error('マスタ件数が不正です');
      const items=p.items.map(i=>{
        if(!['kg','pcs'].includes(i.unit))throw Error('マスタ単位が不正です');
        return {name:text_(i.name,'品目'),rate:num_(i.rate,'単価',false,true),unit:i.unit};
      });
      if(new Set(items.map(i=>i.name)).size!==items.length)throw Error('品目名が重複しています');
      d={items};break;
    }
    default:throw Error('未対応の書込操作');
  }
  const event={operationId:op,action:p.action,at,data:d,request};
  const json=JSON.stringify(event);if(json.length>FS.maxPayload)throw Error('内訳が大きすぎます');
  // One append-only row is both the committed data and the audit trail.
  // Never store a token or interpret user strings as spreadsheet formulas.
  const oldLot=s.est.get(d.id)||{}, oldCost=s.cost.get(d.recordId)||{};
  const row=[at,op,p.action,role,d.id||d.recordId||'',d.lotId||oldCost.lotId||d.id||'',d.auction||oldLot.auction||oldCost.auction||'',d.lotName||oldLot.lotName||oldCost.lotName||'',d.itemName||oldCost.itemName||'',d.amount??oldCost.amount??'',d.unitType||oldCost.unitType||'',d.actualYen??d.totalEstYen??d.buyPriceYen??oldCost.actualYen??'',d.status||'',json];
  sh.appendRow(row.map(cell=>typeof cell==='string'&&/^[=+\-@\t\r]/.test(cell)?"'"+cell:cell));
  SpreadsheetApp.flush();
  return {operationId:op,duplicate:false};
}
function decode_(s){
  const len={'テ':7,'チ':6,'ト':5,'タ':4}[s[0]];
  const map={'ヒ':'1','コ':'2','ト':'3','ニ':'4','サ':'5','カ':'6','エ':'7','ル':'8','ウ':'9','チ':'0'};
  if(!len||!s.slice(1)||s.length>4)return -1;
  const body=s.slice(1);
  if(body.length===2&&body[1]==='オ'&&map[body[0]])return Number(map[body[0]].repeat(len));
  let out='';
  for(const c of body){if(c==='オ'&&out)out+=out.slice(-1);else if(map[c])out+=map[c];else return -1;}
  if(out.length>len)return -1;
  return Number(out.padEnd(len,'0'));
}
