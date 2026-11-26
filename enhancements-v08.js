/* v0.8: Result breakdown, real source zone allocation and arrival-date controls.
   Runs entirely in the browser; never sends private data to a server. */
let priorityDraft=null;
const weekdays={wed:3,thu:4,fri:5,sun:0};
const weekdayZh=['日','一','二','三','四','五','六'];
const restockZoneCatalog={
 信義:['1F倉庫'],西門:['4F倉庫'],ATT:['B1F倉庫','1F陳列'],
 文心:['1F倉庫','1F陳列','2F倉庫','2F陳列'],
 南西:['1F倉庫'],光復:['2F倉庫','3F倉庫'],民族:['1F倉庫','1F陳列'],
 板橋:['1F倉庫','1F陳列'],大隆:['2F倉庫']
};
const defaultWarehouse={信義:'1F倉庫',西門:'4F倉庫',ATT:'B1F倉庫',文心:'1F倉庫',南西:'1F倉庫',光復:'2F倉庫',民族:'1F倉庫',板橋:'1F倉庫',大隆:'2F倉庫'};
const newZonePairs={信義:['1F倉庫'],西門:['4F倉庫','2F陳列'],ATT:['B1F倉庫','1F陳列'],
 文心:['2F倉庫','1F陳列'],南西:['1F倉庫','1F陳列'],光復:['2F倉庫'],
 民族:['1F倉庫','1F陳列'],板橋:['1F倉庫','1F陳列'],大隆:['2F倉庫']};
function isoLocal(d){return [d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-')}
function localToday(){return isoLocal(new Date())}
function nextBatchISO(batch){const today=new Date();today.setHours(12,0,0,0);today.setDate(today.getDate()+(weekdays[batch]-today.getDay()+7)%7);return isoLocal(today)}
function validISO(value){if(!/^\d{4}-\d\d-\d\d$/.test(value||''))return false;const d=new Date(value+'T12:00:00');return Number.isFinite(d.getTime())&&isoLocal(d)===value}
function parsePeriodISO(v){
 const value=String(v||'').trim();const numeric=value.replace(/[-/]/g,'');
 let y,m,d;
 if(/^\d{8}$/.test(numeric)){y=Number(numeric.slice(0,4));m=Number(numeric.slice(4,6));d=Number(numeric.slice(6,8))}
 else if(/^\d{4}$/.test(numeric)){y=new Date().getFullYear();m=Number(numeric.slice(0,2));d=Number(numeric.slice(2,4))}
 else return '';
 const date=[y,String(m).padStart(2,'0'),String(d).padStart(2,'0')].join('-');
 return validISO(date)?date:'';
}
function currentPeriodISO(){
 const dates=[...new Set(rows.map(r=>parsePeriodISO(r.period)).filter(Boolean))];
 if(dates.length>1)throw Error('包含不同新品檔期日期，請分開匯出');
 if(dates.length===1)return dates[0];
 if(allocationTestOnly)return nextBatchISO('sun');
 throw Error('新品正式輸出需要有效檔期日期（例如 20261130）');
}
function zoneKinds(name){return name.includes('陳列')?'display':'warehouse'}
function demandZoneKeys(store,r){
 if(r.zoneDemand&&r.zoneDemand[store])return Object.keys(r.zoneDemand[store]);
 // 沒有來源分區時，絕不自行新增陳列。只有多倉庫的店家需要再輸入分區量。
 return restockZoneCatalog[store].filter(z=>z.includes('倉庫'));
}
function zoneForSKU(r,store){
 const q=Number(r.qty[stores.indexOf(store)]||0);
 if(q===0)return {values:{},verified:true,names:[]};
 if(mode==='new'){
  const names=newZonePairs[store];
  if(names.length===1)return {values:{[names[0]]:q},names,verified:true};
  const size=String(r.size||(/[A-Z]+$/.exec(r.sku||'')||[])[0]||'').toUpperCase();
  if(!size)return {values:{},names,verified:false};
  const show=Math.min(q,store==='南西'?(size==='L'?1:2):1);
  return {values:{[names[0]]:q-show,[names[1]]:show},names,verified:true};
 }
 if(allocationTestOnly)return {values:{[defaultWarehouse[store]]:q},names:[defaultWarehouse[store]],verified:true};
 const names=demandZoneKeys(store,r);
 const manual=r.zoneManual?.[store],given=r.zoneDemand?.[store];
 if(manual){
  const total=Object.values(manual).reduce((a,b)=>a+Number(b||0),0);
  const ok=total===q&&Object.values(manual).every(v=>Number.isInteger(Number(v))&&Number(v)>=0)
   &&Object.keys(manual).every(k=>names.includes(k));
  return {values:manual,names,verified:ok};
 }
 if(given){
  const count=Object.values(given).reduce((a,b)=>a+b,0);
  return {values:given,names,verified:count===q};
 }
 if(names.length===1)return {values:{[names[0]]:q},names,verified:true};
 return {values:{},names,verified:false};
}
function allZonesForStore(store){
 const out=new Set();
 for(const r of rows)if(r.qty[stores.indexOf(store)]>0)for(const zone of zoneForSKU(r,store).names)out.add(zone);
 return [...out];
}
function summarizeZone(store){
 const i=stores.indexOf(store),total=rows.reduce((sum,r)=>sum+(r.qty[i]||0),0);
 let warehouse=0,display=0,unknown=false;
 for(const r of rows){
  const info=zoneForSKU(r,store);if(!info.verified){unknown=true;continue}
  for(const [name,value] of Object.entries(info.values)){
   if(zoneKinds(name)==='display')display+=Number(value);else warehouse+=Number(value);
  }
 }
 return {store,total,warehouse:unknown?null:warehouse,display:unknown?null:display,unknown};
}
function setZoneQty(i,store,zone,v){
 const r=rows[i],allowed=demandZoneKeys(store,r),value=Number(v);
 if(!allowed.includes(zone)||!Number.isInteger(value)||value<0){alert('請輸入有效的非負整數');renderBreakdown();return}
 if(!r.zoneManual)r.zoneManual={};
 if(!r.zoneManual[store]){
  const info=zoneForSKU(r,store);
  r.zoneManual[store]=Object.fromEntries(allowed.map(z=>[z,Number(info.values[z]||0)]));
 }
 r.zoneManual[store][zone]=value;
 renderBreakdown();
 renderLineNotices();
}
function renderBreakdown(){
 if(step!=='result')return;
 const sums=stores.map(summarizeZone);
 const host=document.getElementById('zoneSummary');
 host.innerHTML='<table class="breakdown-table"><thead><tr><th>門市</th><th>總數</th><th>倉庫</th><th>陳列</th></tr></thead><tbody>'+
 sums.map((x,i)=>'<tr><td style="background:'+colors[i]+'88">'+x.store+'</td><td><b>'+x.total+'</b></td><td class="'+(x.unknown?'error':'')+'">'+(x.warehouse??'待確認')+'</td><td class="'+(x.unknown?'error':'')+'">'+(x.display??'待確認')+'</td></tr>').join('')+'</tbody></table>';
 const pending=[];
 if(mode==='restock'){
  rows.forEach((r,i)=>stores.forEach(store=>{
   const q=r.qty[stores.indexOf(store)]||0;
   if(!q)return;
   const info=zoneForSKU(r,store);
   if(info.verified)return;
   pending.push({r,i,store,q,info});
  }));
 }
 const edit=document.getElementById('zoneEditor');
 if(!pending.length){edit.innerHTML='';return}
 edit.innerHTML='<details open><summary><b>補貨分區待確認：'+pending.length+' 筆 SKU／門市</b>（輸入各倉庫實際數量後才允許匯出）</summary><div class="tablebox" style="max-height:400px"><table class="zone-edit-table"><thead><tr><th>門市</th><th>品號</th><th>配貨總數</th><th>倉庫／陳列分區</th><th>已填合計</th></tr></thead><tbody>'+
 pending.map(({r,i,store,q,info})=>'<tr><td>'+store+'</td><td>'+esc(r.sku)+'</td><td>'+q+'</td><td>'+info.names.map(z=>'<label style="display:inline-flex;gap:5px;align-items:center;margin:4px">'+z+' <input type="number" min="0" max="'+q+'" value="'+(info.values[z]??0)+'" onchange="setZoneQty('+i+',decodeURIComponent(\''+encodeURIComponent(store)+'\'),decodeURIComponent(\''+encodeURIComponent(z)+'\'),this.value)"></label>').join('')+'</td><td class="error">'+Object.values(info.values).reduce((a,b)=>a+Number(b||0),0)+'／'+q+'</td></tr>').join('')+'</tbody></table></div></details>';
}
const originalApproveGogo=approve;
approve=function(){
 originalApproveGogo();
 if(step!=='result')return;
 renderBreakdown();
 setupArrivalSelector();
};
const originalRestoreGogo=restore;
restore=function(){originalRestoreGogo();rows.forEach(r=>delete r.zoneManual)};
function initBatchUI(){
 const selector=document.getElementById('lineBatch');
 for(const o of selector.options)o.hidden=mode==='new'&&o.value!=='sun';
 selector.value=mode==='new'?'sun':'fri';
}
function setupArrivalSelector(){
 document.getElementById('lineDayWrap').style.display='flex';
 initBatchUI();
 document.getElementById('lineDate').value=nextBatchISO(document.getElementById('lineBatch').value);
 renderLineNotices();
}
function onLineBatchChanged(){
 const selector=document.getElementById('lineBatch');
 if(mode==='new'&&selector.value!=='sun')selector.value='sun';
 document.getElementById('lineDate').value=nextBatchISO(selector.value);
 renderLineNotices();
}
function displayDateAndTime(timeText){
 const iso=document.getElementById('lineDate').value;
 if(!validISO(iso))return '（日期待確認）';
 const d=new Date(iso+'T12:00:00'),md=(d.getMonth()+1)+'/'+d.getDate();
 const day=weekdayZh[d.getDay()];
 const m=String(timeText||'').match(/^\([^)]+\)\s*(.*)$/);
 const t=(m?m[1]:timeText||'').trim();
 return md+'('+day+')'+(t&&t!=='（時間待確認）'?' '+t:'');
}
renderLineNotices=function(){
 const host=document.getElementById('lineAll');if(!host)return;
 const batch=mode==='new'?'sun':document.getElementById('lineBatch').value;
 const period=rows.map(r=>String(r.period||'')).find(Boolean)||'';
 const isoPeriod=parsePeriodISO(period),label=isoPeriod?isoPeriod.slice(5,7)+isoPeriod.slice(8):'（檔期待確認）';
 const sections=[];
 for(const store of stores){
  const defs=lineDefs[store],config=defs[batch]||null;
  const time=config?.time || '('+({wed:'三',thu:'四'}[batch]||'日')+')（時間待確認）';
  const data=summarizeZone(store),info=new Map(),i=stores.indexOf(store);
  let unresolved=false;
  for(const r of rows)if(r.qty[i]){
   const res=zoneForSKU(r,store);
   if(!res.verified)unresolved=true;
   for(const [zone,n] of Object.entries(res.values))info.set(zone,(info.get(zone)||0)+Number(n));
   if(!res.verified)for(const zone of res.names)if(!info.has(zone))info.set(zone,0);
  }
  // No display line unless source explicitly had a display section or new-product display is required.
  const lines=[...info].filter(([z,n])=>n>0 || (unresolved&&allZonesForStore(store).includes(z)))
    .map(([z,n])=>'PA'+(store==='ATT'?' ATT':store)+'店_'+z+'_'+(mode==='new'?label+'新品':'補貨')+'(共'+(unresolved?'待確認':n)+'件)');
  if(!lines.length&&data.total===0){
    // The user requested the whole supplied store-format notice. A zero-quantity branch is still explicit.
    const fallback=mode==='new'?newZonePairs[store]:[defaultWarehouse[store]];
    for(const z of fallback.slice(0,1))lines.push('PA'+(store==='ATT'?' ATT':store)+'店_'+z+'_'+(mode==='new'?label+'新品':'補貨')+'(共0件)');
  }
  sections.push(lines.join('\n')+'\n'+displayDateAndTime(time));
 }
 host.value=sections.join('\n\n-\n\n');
 const warn=document.getElementById('lineWarning');if(warn){warn.textContent='';warn.style.display='none'}
};
function savePriority(){
 if(!priorityDraft)return;
 saved[mode].priority=[...priorityDraft];
 localStorage.setItem('pz-test-settings',JSON.stringify(saved));
 const t=document.getElementById('prioritySaveState');if(t)t.textContent='排序已儲存';
}
function renderPriority(){
 const target=document.getElementById('priority');
 target.innerHTML='<div class="priority-strip" id="priorityStrip">'+priorityDraft.map((s,i)=>'<div class="priority-chip" draggable="true" data-store="'+s+'" title="拖曳變更優先順序">'+(i+1)+'　'+s+'</div>').join('')+'</div>';
 const strip=document.getElementById('priorityStrip');let source=null;
 function move(a,b){
  if(a===b||a===null||b===null)return;
  const start=priorityDraft.indexOf(a),end=priorityDraft.indexOf(b);
  if(start<0||end<0)return;
  priorityDraft.splice(start,1);priorityDraft.splice(end,0,a);renderPriority();
  document.getElementById('prioritySaveState').textContent='排序尚未儲存';
 }
 for(const chip of strip.children){
  chip.addEventListener('dragstart',e=>{source=chip.dataset.store;chip.classList.add('dragging');e.dataTransfer.effectAllowed='move'});
  chip.addEventListener('dragover',e=>{e.preventDefault();chip.classList.add('over')});
  chip.addEventListener('dragleave',()=>chip.classList.remove('over'));
  chip.addEventListener('drop',e=>{e.preventDefault();move(source,chip.dataset.store)});
  chip.addEventListener('dragend',()=>chip.classList.remove('dragging'));
  let touchStart=null;
  chip.addEventListener('touchstart',e=>{touchStart=chip.dataset.store},{passive:true});
  chip.addEventListener('touchend',e=>{
   if(!touchStart)return;
   const t=e.changedTouches[0],dest=document.elementFromPoint(t.clientX,t.clientY)?.closest('.priority-chip');
   if(dest)move(touchStart,dest.dataset.store);touchStart=null;
  },{passive:true});
 }
}
renderSettings=function(){
 const d=saved[mode];
 priorityDraft=[...d.priority];
 document.getElementById('settingsTitle').textContent=(mode==='new'?'新品':'補貨')+'｜獨立設定';
 document.getElementById('settingFields').innerHTML=mode==='restock'
  ?'<label>官網安全庫存（配貨門檻）<input id="safe" type="number" min="0" step="1" value="'+d.safe+'"></label><button onclick="saveSettings()">儲存安全庫存</button>'
  :'<span class="pill">新品：現貨最多配出一半，奇數官網多保留 1 件</span>';
 document.getElementById('priority').innerHTML='';
 renderPriority();
 const buttons=document.querySelector('#settings .row:last-child');
 buttons.innerHTML='<button class="primary" onclick="savePriority()">儲存排序</button><span id="prioritySaveState" class="sub"></span><button onclick="setStep(\'home\')">返回</button>';
};
saveSettings=function(){
 if(mode==='restock'){
  const n=Number(document.getElementById('safe').value);
  if(!Number.isInteger(n)||n<0){alert('安全庫存請輸入非負整數');return}
  saved.restock.safe=n;
  delete saved.restock.threshold;
 }
 localStorage.setItem('pz-test-settings',JSON.stringify(saved));
};
cap=function(r){return mode==='new'?Math.floor(r.stock/2):Math.max(0,r.stock-Number(saved.restock.safe||0))};
function exportDateISO(){return mode==='restock'?localToday():currentPeriodISO()}
function exportDateYMD(){return exportDateISO().replace(/-/g,'')}
function exportDateMD(){return exportDateYMD().slice(4)}
const originalDownloadResultGogo=downloadResult;
downloadResult=function(){
 const date=exportDateYMD();
 download('PAZZO_'+(mode==='new'?'新品':'補貨')+'_'+date+'.csv',makeCsv([['SKU','SPU','庫存',...stores,'庫存剩餘'],...rows.map(r=>[r.sku,r.spu,r.stock,...r.qty,r.stock-r.qty.reduce((a,b)=>a+b,0)])]));
};
const originalImportGogo=importCsv;
importCsv=function(event){
 const file=event.target.files?.[0];if(!file)return;
 allocationTestOnly=false;
 const reader=new FileReader();
 reader.onload=()=>{
  try{
   const values=csvParse(String(reader.result).replace(/^\uFEFF/,'')),head=values.shift().map(s=>String(s).trim());
   for(const col of ['SKU','SPU','庫存',...stores])if(!head.includes(col))throw Error('缺少欄位：'+col);
   const val=(v,name)=>head.includes(name)?String(v[head.indexOf(name)]??'').trim():'';
   const parsed=values.map(v=>{
    const row={sku:val(v,'SKU'),spu:val(v,'SPU'),stock:Number(val(v,'庫存')),demand:stores.map(s=>Number(val(v,s))),qty:[],
      period:val(v,'檔期'),name:val(v,'品名'),color:val(v,'顏色'),size:val(v,'尺寸'),blocked:false,why:'',zoneDemand:{}};
    if(!row.sku&&!row.spu)return null;
    if(!row.sku||!row.spu||!Number.isInteger(row.stock)||row.stock<0||row.demand.some(x=>!Number.isInteger(x)||x<0))throw Error('品號／SPU／庫存／門市需求格式錯誤：'+row.sku);
    for(const store of stores){
     const found={};
     for(const zone of restockZoneCatalog[store]){
      const headers=[store+'_'+zone,store+'｜'+zone];
      const col=headers.find(h=>head.includes(h));
      if(!col)continue;
      const text=val(v,col);
      if(text==='')continue;
      const n=Number(text);
      if(!Number.isInteger(n)||n<0)throw Error(row.sku+' '+col+' 必須是非負整數');
      found[zone]=n;
     }
     if(Object.keys(found).length)row.zoneDemand[store]=found;
    }
    return row;
   }).filter(Boolean);
   if(!parsed.length)throw Error('沒有商品資料');
   if(new Set(parsed.map(r=>r.sku)).size!==parsed.length)throw Error('SKU 重複，停止匯入');
   rows=parsed;allocate();setStep('work');
  }catch(err){alert('CSV 匯入失敗：'+err.message)}
 };
 reader.readAsText(file,'UTF-8');
};
