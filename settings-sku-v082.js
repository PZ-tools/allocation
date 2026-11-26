
/* v0.8.2: shared settings and SKU-based new-arrival caps */
const newDefaultStores=['信義','西門','ATT','文心','南西','光復','民族','板橋'];
let settingsSection='common', newConfirmed=false, includeDalong=false, zeroReasons={};
const commonKey='pz-common-priority-v1', capKey='pz-new-sku-cap-v1';
let commonPriority=saved.restock.priority.slice(),skuCap={kind:'percent',value:50};
try{const v=JSON.parse(localStorage.getItem(commonKey)||'null');if(Array.isArray(v)&&v.length===stores.length&&stores.every(s=>v.includes(s)))commonPriority=v}catch(e){}
try{const v=JSON.parse(localStorage.getItem(capKey)||'null');if(v&&['percent','count'].includes(v.kind)&&Number.isInteger(v.value)&&v.value>=0&&(v.kind!=='percent'||v.value<=100))skuCap=v}catch(e){}
function saveCommonOrder(){
 commonPriority=priorityDraft.slice();
 saved.restock.priority=commonPriority.slice();saved.new.priority=commonPriority.slice();
 localStorage.setItem(commonKey,JSON.stringify(commonPriority));localStorage.setItem('pz-test-settings',JSON.stringify(saved));
 const e=document.getElementById('prioritySaveState');if(e)e.textContent='排序已儲存';
}
savePriority=saveCommonOrder;
cap=function(r){
 const stock=Math.max(0,Math.floor(Number(r.stock)||0));
 if(mode==='restock')return Math.max(0,stock-Number(saved.restock.safe||0));
 return skuCap.kind==='percent'?Math.floor(stock*skuCap.value/100):Math.min(stock,skuCap.value);
};
allocate=function(){
 for(const r of rows){
  if(r.blocked){r.qty=stores.map(()=>0);continue}
  const available=cap(r);
  r.qty=stores.map((s,i)=>mode==='new'&&s==='大隆'&&!includeDalong?0:r.demand[i]||0);
  let used=r.qty.reduce((a,b)=>a+b,0);
  for(const store of [...commonPriority].reverse()){
   if(used<=available)break;
   const i=stores.indexOf(store),cut=Math.min(r.qty[i],used-available);
   r.qty[i]-=cut;used-=cut;
  }
  if(used>available)r.qty=stores.map(()=>0);
 }
 suggested=rows.map(r=>r.qty.slice());render();
};
const baseSetMode082=setMode;
setMode=function(m){baseSetMode082(m);newConfirmed=false;includeDalong=false;zeroReasons={};renderNewEntry()};
function renderNewEntry(){
 document.getElementById('newEntry')?.remove();
 const home=document.getElementById('home');
 if(mode!=='new'){
  home.querySelectorAll('button').forEach(b=>b.disabled=false);
  document.getElementById('upload').disabled=false;
  return;
 }
 const bar=document.createElement('div');bar.id='newEntry';bar.className='box';bar.style.margin='12px 0';
 bar.innerHTML='<div class="row" style="gap:16px"><b>本次新品配貨門市</b><span>預設門市 8 間</span>'+
 '<label><input type="checkbox" '+(includeDalong?'checked':'')+' onchange="includeDalong=this.checked;newConfirmed=false;renderNewEntry()"> 加入大隆</label>'+
 '<span class="sub">本次共 '+(includeDalong?9:8)+' 間</span><button class="primary" onclick="confirmNewEntry()">確認門市</button>'+
 '<span class="sub">'+(newConfirmed?'已確認':'請先確認再匯入資料')+'</span></div>';
 home.insertBefore(bar,home.querySelector('.row'));
 for(const b of home.querySelectorAll('button')){
  if(b.closest('#newEntry'))continue;
  if(/sample\(|sampleAllStores\(/.test(b.getAttribute('onclick')||''))b.disabled=!newConfirmed;
 }
 document.getElementById('upload').disabled=!newConfirmed;
}
function confirmNewEntry(){newConfirmed=true;renderNewEntry()}
const originalSample082=sample;
sample=function(){if(mode==='new'&&!newConfirmed){alert('請先確認新品配貨門市');return}originalSample082()};
const originalAll082=sampleAllStores;
sampleAllStores=function(){if(mode==='new'&&!newConfirmed){alert('請先確認新品配貨門市');return}originalAll082()};
const originalImport082=importCsv;
importCsv=function(e){if(mode==='new'&&!newConfirmed){alert('請先確認新品配貨門市');if(e?.target)e.target.value='';return}return originalImport082(e)};
const originalChange082=changeQty;
changeQty=function(i,j,v){if(mode==='new'&&stores[j]==='大隆'&&!includeDalong){alert('本次未選大隆');render();return}originalChange082(i,j,v)};
const originalRender082=render;
render=function(){
 originalRender082();
 if(step!=='work'||mode!=='new')return;
 const groups=[...new Set(rows.map(r=>r.spu))];
 document.getElementById('newControls').innerHTML='<div class="box"><h3>當次新品歸零勾選</h3>'+
 '<div style="overflow-x:auto"><table class="zero-table" style="min-width:550px"><thead><tr><th>貨號（9碼）</th><th>原因</th><th>尺寸</th></tr></thead><tbody>'+
 groups.map(spu=>{
  const skuRows=rows.map((r,i)=>({r,i})).filter(x=>x.r.spu===spu);
  const reason=zeroReasons[spu]||'';
  return '<tr><td><label><input type="checkbox" '+(skuRows.every(x=>x.r.blocked)?'checked':'')+
   ' onchange="updateZeroGroup(\''+encodeURIComponent(spu)+'\',this.checked)"> <b>'+esc(spu)+'</b></label></td>'+
   '<td><select onchange="zeroReasons[decodeURIComponent(\''+encodeURIComponent(spu)+'\')]=this.value">'+
   '<option value="">選擇原因</option>'+['尚未入庫','官網預購','分批到貨'].map(v=>'<option '+(v===reason?'selected':'')+'>'+v+'</option>').join('')+
   '</select></td><td>'+skuRows.map(({r,i})=>'<label class="pill"><input type="checkbox" '+(r.blocked?'checked':'')+
   ' onchange="updateZeroSku('+i+',this.checked)"> '+esc(r.size||'尺寸待庫存核對')+'</label>').join(' ')+'</td></tr>'
 }).join('')+'</tbody></table></div></div>';
};
function updateZeroGroup(encoded,checked){
 const spu=decodeURIComponent(encoded);
 if(checked&&!zeroReasons[spu]){alert('請先選擇歸零原因');render();return}
 rows.filter(r=>r.spu===spu).forEach(r=>r.blocked=checked);allocate();
}
function updateZeroSku(i,checked){
 if(checked&&!zeroReasons[rows[i].spu]){alert('請先選擇歸零原因');render();return}
 rows[i].blocked=checked;allocate();
}
const previousApprove082=approve;
approve=function(){
 if(mode==='new'&&!newConfirmed){alert('請先確認新品配貨門市');return}
 previousApprove082();
};
function settingsTab(tab){settingsSection=tab;renderSettings()}
function saveSafety082(){
 const n=Number(document.getElementById('safe')?.value);
 if(!Number.isInteger(n)||n<0){alert('安全庫存必須是非負整數');return}
 saved.restock.safe=n;localStorage.setItem('pz-test-settings',JSON.stringify(saved));
 document.getElementById('settingsFeedback').textContent='安全庫存已儲存';
}
function saveSkuCap082(){
 const kind=document.getElementById('capKind')?.value,n=Number(document.getElementById('capNumber')?.value);
 if(!Number.isInteger(n)||n<0||(kind==='percent'&&n>100)){alert('上限必須是有效整數，百分比範圍 0～100');return}
 skuCap={kind,value:n};localStorage.setItem(capKey,JSON.stringify(skuCap));
 document.getElementById('settingsFeedback').textContent='新品 SKU 上限已儲存';
}
function showStoreDraft082(){document.getElementById('storeDraft').hidden=false}
function checkStoreDraft082(){
 const name=document.getElementById('storeName')?.value.trim();
 if(!name){alert('請填門市名稱');return}
 if(stores.includes(name)){alert('門市已存在');return}
 const zones=document.getElementById('storeZones')?.value.trim(),line=document.getElementById('storeLine')?.value.trim();
 document.getElementById('storeDraftStatus').textContent=name+'：'+(!zones?'待確認倉庫／陳列；':'')+(!line?'待確認 LINE 通知及時間；':'')+'開店位置與配貨優先序也需確認。未正式加入配貨。';
}
renderSettings=function(){
 document.getElementById('settingsTitle').textContent='配貨設定';
 const info=document.querySelector('#settings > p.sub');
 if(info)info.textContent='設定保存在目前瀏覽器。';
 const fields=document.getElementById('settingFields'),prio=document.getElementById('priority');
 const footer=document.querySelector('#settings .row:last-child');
 const tabs='<div class="row" style="width:100%">'+[['common','共通'],['restock','補貨'],['new','新品']].map(([id,label])=>
   '<button class="'+(settingsSection===id?'active':'')+'" onclick="settingsTab(\''+id+'\')">'+label+'</button>').join('')+'</div>';
 if(settingsSection==='common'){
  fields.innerHTML=tabs+'<div class="box" style="width:100%"><h3>門市管理</h3><p class="sub">'+stores.map(esc).join('、')+'</p>'+
  '<button onclick="showStoreDraft082()">＋ 新增門市</button><div id="storeDraft" hidden style="margin-top:12px">'+
  '<div class="row"><label>門市 <input id="storeName" placeholder="例如 中清"></label><label>倉庫／陳列 <input id="storeZones" placeholder="例如 1F倉庫"></label>'+
  '<label>公版 <select id="storeTemplate"><option>美學</option><option>實業</option></select></label></div>'+
  '<label>LINE 到貨資訊 <input id="storeLine" placeholder="到貨星期、時間及文字格式"></label> '+
  '<button onclick="checkStoreDraft082()">檢查新增條件</button><p id="storeDraftStatus" class="sub"></p></div></div>'+
  '<div style="width:100%"><h3>共用配貨優先順序</h3><p class="sub">補貨由左至右優先；新品不足時由右至左扣減。拖曳後儲存。</p></div>';
  priorityDraft=commonPriority.slice();prio.innerHTML='';renderPriority();
  footer.innerHTML='<button class="primary" onclick="saveCommonOrder()">儲存排序</button><span class="sub" id="prioritySaveState"></span><button onclick="setStep(\'home\')">返回</button>';
 }else if(settingsSection==='restock'){
  fields.innerHTML=tabs+'<div class="box" style="width:100%"><h3>補貨</h3><label>官網安全庫存 <input type="number" id="safe" min="0" step="1" value="'+saved.restock.safe+'"></label><p class="sub">安全庫存即配貨門檻。</p></div>';
  prio.innerHTML='';footer.innerHTML='<button class="primary" onclick="saveSafety082()">儲存安全庫存</button><span class="sub" id="settingsFeedback"></span><button onclick="setStep(\'home\')">返回</button>';
 }else{
  fields.innerHTML=tabs+'<div class="box" style="width:100%"><h3>新品</h3><p class="sub">每個 SKU 獨立設定配貨上限。</p>'+
  '<div class="row"><label>計算方式 <select id="capKind"><option value="percent" '+(skuCap.kind==='percent'?'selected':'')+'>百分比 %</option><option value="count" '+(skuCap.kind==='count'?'selected':'')+'>固定件數</option></select></label>'+
  '<label>SKU 上限 <input type="number" min="0" step="1" id="capNumber" value="'+skuCap.value+'"></label></div></div>';
  prio.innerHTML='';footer.innerHTML='<button class="primary" onclick="saveSkuCap082()">儲存新品上限</button><span class="sub" id="settingsFeedback"></span><button onclick="setStep(\'home\')">返回</button>';
 }
};
