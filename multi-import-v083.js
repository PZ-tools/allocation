
/* v0.8.3: independent multi-file import for demand and stock.
   Only local files are read; cloud Sheets connection is not claimed. */
const importState={demand:[],stock:[],errors:[],items:[],ready:false};
const importAlias={
 sku:['SKU','商品SKU','貨號','品號','商品編號','條碼'],
 spu:['SPU','9碼貨號','九碼貨號','主貨號','商品款號'],
 stock:['庫存','現貨庫存','可用庫存','庫存數量','現貨數量','可配庫存'],
 qty:['數量','需求','叫貨數量','叫貨量','需求數量','補貨數量','營運確認數','【2】實體門市 叫貨 營運確認數'],
 store:['門市','店名','門市名稱'],
 zone:['區域','倉庫/陳列','位置','分區'],
 period:['檔期','檔期日期','批次'],
 size:['尺寸','規格'],name:['品名','商品名稱'],color:['顏色']
};
function impHead(h){return String(h??'').replace(/\uFEFF/g,'').replace(/[\s　]/g,'').replace(/[（）()]/g,'').toUpperCase()}
function impIndex(headers,kind){const synonyms=importAlias[kind].map(impHead);return headers.findIndex(h=>synonyms.includes(impHead(h)))}
function impField(row,headers,kind){const i=impIndex(headers,kind);return i<0?'':String(row[i]??'').trim()}
function impWhole(x,label){const value=String(x??'').trim();if(!/^\d+$/.test(value))throw Error(label+' 必須為非負整數');return Number(value)}
function impSku(row,head){return impField(row,head,'sku')}
function impSpu(row,head,sku){
 const explicit=impField(row,head,'spu');if(explicit)return explicit;
 const suffix=/(?:[-_])(XXL|XL|XS|S|M|L|F)$/i.exec(sku||'');
 return suffix?sku.slice(0,-suffix[0].length):sku;
}
function impSize(row,head,sku){
 const explicit=impField(row,head,'size');
 return explicit||((/(?:[-_])(XXL|XL|XS|S|M|L|F)$/i.exec(sku||'')||[])[1]||'');
}
function impGuessedStore(label){
 const found=stores.filter(store=>String(label||'').includes(store));
 return found.length===1?found[0]:'';
}
function impGuessedZone(source,store){
 const zones=[...(restockZoneCatalog[store]||[]),...(newZonePairs[store]||[])];
 const matched=[...new Set(zones.filter(z=>String(source||'').includes(z)))];
 return matched.length===1?matched[0]:'';
}
function impCellValues(raw){
 const rows=csvParse(raw.replace(/^\uFEFF/,''));
 return rows.map(r=>r.map(v=>String(v??'').trim())).filter(r=>r.some(Boolean));
}
async function impXlsxSheets(file){
 const entries=zipRead(new Uint8Array(await file.arrayBuffer()));
 const find=name=>entries.find(x=>x.name===name);
 const wb=find('xl/workbook.xml');if(!wb)throw Error('找不到 XLSX 活頁簿');
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 const decodeXml=async entry=>{
  const xml=new DOMParser().parseFromString(dec.decode(await uncompressed(entry)),'application/xml');
  if(xml.querySelector('parsererror'))throw Error('Excel XML 格式無法辨識');
  return xml;
 };
 const book=await decodeXml(wb);
 let shared=[];
 const ss=find('xl/sharedStrings.xml');
 if(ss){const xml=await decodeXml(ss);shared=Array.from(xml.getElementsByTagNameNS(ns,'si')).map(n=>n.textContent||'')}
 const tabs=Array.from(book.getElementsByTagNameNS(ns,'sheet')).map((n,i)=>({name:n.getAttribute('name')||('工作表'+(i+1)),id:n.getAttribute('r:id')||n.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')}));
 const rel=find('xl/_rels/workbook.xml.rels'),relMap=new Map();
 if(rel){const xml=await decodeXml(rel);for(const n of Array.from(xml.getElementsByTagName('Relationship')))relMap.set(n.getAttribute('Id'),n.getAttribute('Target'))}
 const results=[];
 for(const [i,tab] of tabs.entries()){
  let target=relMap.get(tab.id)||('worksheets/sheet'+(i+1)+'.xml');
  target=target.replace(/^\/?xl\//,'').replace(/^\/+/, '');
  const entry=find('xl/'+target);
  if(!entry)continue;
  const doc=await decodeXml(entry);
  const data=[];
  for(const row of Array.from(doc.getElementsByTagNameNS(ns,'sheetData')[0]?.getElementsByTagNameNS(ns,'row')||[])){
   const values=[];
   for(const cell of Array.from(row.getElementsByTagNameNS(ns,'c'))){
    const letters=/^[A-Z]+/.exec(cell.getAttribute('r')||'')?.[0];if(!letters)continue;
    let index=0;for(const ch of letters)index=index*26+(ch.charCodeAt(0)-64);index--;
    const type=cell.getAttribute('t'),v=cell.getElementsByTagNameNS(ns,'v')[0]?.textContent||'';
    values[index]=(type==='s'?shared[Number(v)]||'':type==='inlineStr'?cell.getElementsByTagNameNS(ns,'is')[0]?.textContent||'':v).trim();
   }
   if(values.some(Boolean))data.push(values);
  }
  if(data.length)results.push({name:tab.name,rows:data});
 }
 if(!results.length)throw Error('找不到有資料的工作表');
 return results;
}
async function impRead(file){
 const name=file.name.toLowerCase();
 if(name.endsWith('.csv'))return [{name:file.name,rows:impCellValues(await file.text())}];
 if(name.endsWith('.xlsx'))return impXlsxSheets(file);
 throw Error('請選擇 CSV 或 XLSX');
}
function impHeaderSheet(data,kind,source){
 const pick=data.findIndex((r,i)=>i<25&&impIndex(r,'sku')>=0&&(kind==='stock'?impIndex(r,'stock')>=0:(
  impIndex(r,'qty')>=0||impIndex(r,'store')>=0||stores.some(s=>r.some(c=>impHead(c)===impHead(s)||impHead(c).startsWith(impHead(s)+'_'))))));
 if(pick<0)throw Error(source+'：找不到 '+(kind==='stock'?'SKU／庫存':'SKU／門市需求')+' 表頭，無法辨識');
 return {head:data[pick].map(v=>String(v||'').trim()),records:data.slice(pick+1)};
}
function impDemandTable(data,source){
 const {head,records}=impHeaderSheet(data,'demand',source);
 const skuIdx=impIndex(head,'sku'),qtyIdx=impIndex(head,'qty'),storeIdx=impIndex(head,'store'),zoneIdx=impIndex(head,'zone');
 const sourceStore=impGuessedStore(source);
 const shopCols=[];
 for(let c=0;c<head.length;c++){
  const col=head[c].replace(/[｜|]/g,'_').replace(/\s+/g,'');
  for(const store of stores){
   if(col===store||col.startsWith(store+'_')){
    const zone=col===store?'':col.slice(store.length+1);
    shopCols.push({c,store,zone});
    break;
   }
  }
 }
 if(shopCols.length===0&&storeIdx<0&&!sourceStore)throw Error(source+'：無法確認需求所屬門市，請使用門市欄或在檔名加入門市名稱');
 if(!shopCols.length&&qtyIdx<0)throw Error(source+'：找不到需求數量欄');
 const out=[];
 for(const [j,row] of records.entries()){
  const sku=String(row[skuIdx]||'').trim();if(!sku)continue;
  const item={sku,spu:impSpu(row,head,sku),size:impSize(row,head,sku),name:impField(row,head,'name'),color:impField(row,head,'color'),period:impField(row,head,'period')};
  const entries=shopCols.length?shopCols.map(c=>({store:c.store,zone:c.zone,qty:row[c.c]})): [{store:storeIdx<0?sourceStore:String(row[storeIdx]||'').trim(),zone:zoneIdx<0?impGuessedZone(source,sourceStore):String(row[zoneIdx]||'').trim(),qty:row[qtyIdx]}];
  for(const e of entries){
   if(!stores.includes(e.store))throw Error(source+'：第 '+(j+2)+' 列門市無效：'+e.store);
   if(e.qty===''||e.qty==null)continue;
   const qty=impWhole(e.qty,source+' '+sku+' '+e.store);
   out.push({...item,store:e.store,zone:e.zone,qty,source});
  }
 }
 if(!out.length)throw Error(source+'：未找到有效需求數量');
 return out;
}
function impStockTable(data,source){
 const {head,records}=impHeaderSheet(data,'stock',source),i=impIndex(head,'sku'),c=impIndex(head,'stock');
 const out=[];
 for(const row of records){
  const sku=String(row[i]||'').trim();if(!sku)continue;
  const val=String(row[c]??'').trim();if(!val)throw Error(source+'：'+sku+' 庫存空白');
  out.push({sku,spu:impSpu(row,head,sku),size:impSize(row,head,sku),name:impField(row,head,'name'),color:impField(row,head,'color'),stock:impWhole(val,source+' '+sku),source});
 }
 if(!out.length)throw Error(source+'：找不到有效庫存資料');
 return out;
}
async function importFiles(kind,fileList){
 if(!fileList?.length)return;
 const container=importState[kind],errors=[];
 for(const file of Array.from(fileList)){
  if(container.some(x=>x.name===file.name&&x.size===file.size&&x.modified===file.lastModified)){errors.push(file.name+'：已匯入相同檔案');continue}
  try{
   const sheets=await impRead(file);
   const extracted=[];
   for(const sheet of sheets){
    const source=file.name+' / '+sheet.name;
    try{extracted.push(...(kind==='demand'?impDemandTable(sheet.rows,source):impStockTable(sheet.rows,source)))}catch(err){
     // Additional blank or unrelated worksheets do not block a valid worksheet in the same XLSX.
     if(sheets.length===1)throw err;
     errors.push(err.message);
    }
   }
   if(!extracted.length)throw Error('沒有可辨識的'+(kind==='demand'?'需求':'庫存')+'資料');
   container.push({name:file.name,size:file.size,modified:file.lastModified,entries:extracted});
  }catch(err){errors.push(file.name+'：'+err.message)}
 }
 importState.errors=errors;importState.ready=false;renderImportState();
}
function removeImported(kind,index){importState[kind].splice(index,1);importState.ready=false;renderImportState()}
function impTotals(){
 const demand=importState.demand.flatMap(f=>f.entries),stock=importState.stock.flatMap(f=>f.entries);
 const warnings=[],demandKeys=new Set(),stockKeys=new Set(),skuSet=new Set(),meta=new Map();
 const demands=new Map(),stocks=new Map();
 for(const r of demand){
  const key=[r.sku,r.store,r.zone].join('|');
  if(demandKeys.has(key)){warnings.push('需求重複：'+r.sku+'／'+r.store+(r.zone?'／'+r.zone:''));continue}
  demandKeys.add(key);skuSet.add(r.sku);
  if(!demands.has(r.sku))demands.set(r.sku,[]);
  demands.get(r.sku).push(r);
  if(!meta.has(r.sku))meta.set(r.sku,r);
 }
 for(const r of stock){
  if(stockKeys.has(r.sku)){warnings.push('庫存貨號重複：'+r.sku);continue}
  stockKeys.add(r.sku);stocks.set(r.sku,r);
 }
 for(const sku of skuSet)if(!stocks.has(sku))warnings.push('缺少庫存：'+sku);
 for(const sku of stocks.keys())if(!skuSet.has(sku))warnings.push('庫存未對應需求：'+sku);
 return {demand,stock,warnings,demands,stocks,skuSet,meta};
}
function makeImportRows(){
 const t=impTotals();
 if(t.warnings.length)throw Error('請先處理資料異常：'+t.warnings.slice(0,5).join('；')+(t.warnings.length>5?'…':''));
 if(!t.demand.length||!t.stock.length)throw Error('需求及庫存都必須匯入');
 return [...t.demands].map(([sku,entries])=>{
  const stock=t.stocks.get(sku),meta=t.meta.get(sku);
  if(stock.spu&&meta.spu&&stock.spu!==meta.spu)throw Error('SKU 與 SPU 不一致：'+sku);
  const demand=stores.map(store=>entries.filter(x=>x.store===store).reduce((a,b)=>a+b.qty,0)),zoneDemand={};
  for(const r of entries){
   if(!r.zone)continue;
   zoneDemand[r.store]??={};zoneDemand[r.store][r.zone]=(zoneDemand[r.store][r.zone]||0)+r.qty;
  }
  return {sku,spu:stock.spu||meta.spu,stock:stock.stock,demand,qty:[],blocked:false,why:'',
   period:meta.period||'',name:stock.name||meta.name||'',color:stock.color||meta.color||'',size:stock.size||meta.size||'',zoneDemand};
 });
}
function downloadImportDemo(kind){
 if(kind==='demand')download('PAZZO_需求測試.csv',makeCsv([
  ['SKU','SPU','檔期','品名','尺寸',...stores],
  ['PZ-DEMO-001-S','PZ-DEMO-001','20261130','模擬商品','S',2,3,2,1,3,1,1,1,0],
  ['PZ-DEMO-001-M','PZ-DEMO-001','20261130','模擬商品','M',2,3,2,1,3,1,1,1,0]
 ]));
 else download('PAZZO_庫存測試.csv',makeCsv([
  ['SKU','SPU','庫存','尺寸','品名'],
  ['PZ-DEMO-001-S','PZ-DEMO-001',50,'S','模擬商品'],
  ['PZ-DEMO-001-M','PZ-DEMO-001',50,'M','模擬商品']
 ]));
}
function downloadImportSkus(){
 const t=impTotals();
 if(!t.demand.length){alert('請先匯入需求');return}
 download('PAZZO_'+(mode==='new'?'新品':'補貨')+'_待查庫存貨號.csv',makeCsv([['SKU','SPU'],...[...t.meta].map(([sku,r])=>[sku,r.spu])]));
}
function startFromFiles(){
 try{
  if(mode==='new'&&!newConfirmed)throw Error('請先確認本次新品配貨門市');
  const parsed=makeImportRows();
  if(mode==='new'&&!includeDalong&&parsed.some(r=>r.demand[stores.indexOf('大隆')]>0)){
   // Explicitly ignore Dalong after warning, never silently allocate.
   document.getElementById('importStatus').textContent='本次未選大隆，已保留來源需求但不分配大隆';
  }
  allocationTestOnly=false;rows=parsed;allocate();setStep('work');
 }catch(err){alert('無法開始配貨：'+err.message);renderImportState()}
}
function renderImportState(){
 const el=document.getElementById('importStatus');if(!el)return;
 const t=impTotals(),parts=[];
 for(const kind of ['demand','stock']){
  const label=kind==='demand'?'需求':'庫存',files=importState[kind];
  const cell=document.getElementById('import'+(kind==='demand'?'Demand':'Stock')+'List');
  if(cell)cell.innerHTML=files.map((f,i)=>'<div class="row" style="margin:3px 0;gap:8px"><span class="pill">'+esc(f.name)+'｜'+f.entries.length+' 筆</span><button type="button" onclick="removeImported(\''+kind+'\','+i+')" style="padding:4px 8px">移除</button></div>').join('');
  parts.push(label+' '+files.length+' 檔');
 }
 parts.push('需求 '+t.demand.length+' 筆','庫存 '+t.stock.length+' 筆','貨號 '+t.skuSet.size+' 個');
 if(importState.errors.length||t.warnings.length)parts.push('⚠ '+[...importState.errors,...t.warnings].slice(0,6).join('；'));
 el.textContent=parts.join('｜');
 const next=document.getElementById('importProceed');if(next)next.disabled=!t.demand.length||!t.stock.length||!!t.warnings.length||!!importState.errors.length;
}
function toggleImportPanel(){
 const panel=document.getElementById('importPanel');
 panel.hidden=!panel.hidden;
}
function clearImportState(){
 importState.demand=[];importState.stock=[];importState.errors=[];importState.ready=false;
 renderImportState();
}
const baseSetModeImport=setMode;
setMode=function(m){baseSetModeImport(m);clearImportState()};
function newImportPanel(){
 const home=document.getElementById('home');
 const legacy=document.getElementById('upload');
 legacy.disabled=true;
 const oldRow=legacy.closest('.row');if(oldRow)oldRow.hidden=true;
 const oldNote=oldRow?.nextElementSibling;if(oldNote?.classList.contains('sub'))oldNote.hidden=true;
 const btnBar=home.querySelector('.row');
 const wrapper=document.createElement('div');wrapper.className='box';wrapper.style.margin='12px 0';
 wrapper.innerHTML='<button type="button" class="primary" onclick="toggleImportPanel()">匯入 ▾</button>'+
 '<div id="importPanel" hidden style="margin-top:12px">'+
 '<div class="row"><b>需求</b><input type="file" accept=".xlsx,.csv" multiple onchange="importFiles(\'demand\',this.files);this.value=\'\'"></div><div id="importDemandList"></div>'+
 '<div class="row"><b>庫存</b><input type="file" accept=".xlsx,.csv" multiple onchange="importFiles(\'stock\',this.files);this.value=\'\'"></div><div id="importStockList"></div>'+
 '<div class="row"><button type="button" onclick="downloadImportSkus()">下載待查庫存貨號</button>'+
 '<button type="button" id="importProceed" class="primary" onclick="startFromFiles()" disabled>開始配貨</button></div>'+
 '<div class="row"><button type="button" onclick="downloadImportDemo(\'demand\')">需求測試範本</button><button type="button" onclick="downloadImportDemo(\'stock\')">庫存測試範本</button></div>'+ 
 '<div id="importStatus" class="sub"></div></div>';
 home.insertBefore(wrapper,btnBar);
 renderImportState();
}
newImportPanel();
