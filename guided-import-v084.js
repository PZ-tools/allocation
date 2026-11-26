/* PAZZO 0.8.4: guided replenishment / new import and 9-code backend lookup. */
const guided={stage:1,batch:0,busy:false,error:'',confirmed:false};
function guideCode(n){return /^[A-Z0-9]{9}$/i.test(String(n||'').trim())?String(n).trim().toUpperCase():''}
function guide9(row,head,barcode){
 const idx=head.findIndex(h=>['貨號','9碼貨號','九碼貨號','SPU','商品貨號'].includes(String(h||'').trim()));
 const first=idx<0?'':guideCode(row[idx]);
 if(first)return first;
 const v=String(barcode||'').trim().toUpperCase();
 return /^[A-Z0-9]{13}$/.test(v)?v.slice(0,9):'';
}
async function guideOpenWorkbook(file){
 const entries=zipRead(new Uint8Array(await file.arrayBuffer()));
 const find=name=>entries.find(e=>e.name===name);
 const workbook=find('xl/workbook.xml');if(!workbook)throw Error('XLSX 缺少活頁簿');
 const parse=async e=>{
  const doc=new DOMParser().parseFromString(dec.decode(await uncompressed(e)),'application/xml');
  if(doc.querySelector('parsererror'))throw Error('Excel XML 解析失敗');return doc;
 };
 const wb=await parse(workbook);
 const rns='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
 const sheetElements=Array.from(wb.getElementsByTagNameNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main','sheet'));
 const rels=new Map(), relFile=find('xl/_rels/workbook.xml.rels');
 if(relFile){const doc=await parse(relFile);for(const e of Array.from(doc.getElementsByTagName('Relationship')))rels.set(e.getAttribute('Id'),e.getAttribute('Target'))}
 const tabs=sheetElements.map((e,i)=>{
  let target=rels.get(e.getAttributeNS(rns,'id'))||('worksheets/sheet'+(i+1)+'.xml');
  target=target.replace(/^\/?xl\//,'').replace(/^\/+/,'');
  return {name:e.getAttribute('name')||('工作表'+(i+1)),entry:find('xl/'+target)};
 }).filter(x=>x.entry);
 return {entries,tabs,shared:null,find,parse};
}
async function guideReadSheet(file,index){
 const book=file.book,tab=book.tabs[index];if(!tab)throw Error('找不到選定工作表');
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 if(book.shared===null){
  const e=book.find('xl/sharedStrings.xml');
  book.shared=e?Array.from((await book.parse(e)).getElementsByTagNameNS(ns,'si')).map(x=>x.textContent||''):[];
 }
 const doc=await book.parse(tab.entry),data=[];
 const sheetData=doc.getElementsByTagNameNS(ns,'sheetData')[0];
 if(!sheetData)throw Error('選定工作表沒有資料');
 for(const row of Array.from(sheetData.getElementsByTagNameNS(ns,'row'))){
  const cells=[];
  for(const cell of Array.from(row.getElementsByTagNameNS(ns,'c'))){
   const a=/^[A-Z]+/.exec(cell.getAttribute('r')||'')?.[0];if(!a)continue;
   let i=0;for(const ch of a)i=i*26+ch.charCodeAt(0)-64;i--;
   const v=cell.getElementsByTagNameNS(ns,'v')[0]?.textContent||'';
   const t=cell.getAttribute('t');
   cells[i]=t==='s'?(book.shared[Number(v)]||''):t==='inlineStr'?(cell.getElementsByTagNameNS(ns,'is')[0]?.textContent||''):String(v);
  }
  if(cells.some(x=>String(x||'').trim()))data.push(cells);
 }
 return file.kind==='demand'?impDemandTable(data,file.name+' / '+tab.name):impStockTable(data,file.name+' / '+tab.name);
}
function guideHeader(data,kind,source){
 const found=data.findIndex((row,i)=>i<35&&row.some(v=>/^(條碼|SKU|品號|貨號|\*貨號\(13碼\))$/i.test(String(v||'').trim()))&&
 (kind==='stock'?row.some(v=>/^(庫存|現貨庫存|可用庫存|庫存數量|現貨數量|可配庫存)$/i.test(String(v||'').trim())):
 row.some(v=>/補貨需求|營運確認數|叫貨數量|^數量$|^需求$|^(信義|西門|ATT|文心|南西|光復|民族|板橋|大隆)/i.test(String(v||'').trim()))));
 if(found<0)throw Error(source+'：缺少可辨識的貨號與'+(kind==='stock'?'庫存':'需求')+'表頭');
 return {head:data[found].map(x=>String(x??'').trim()),records:data.slice(found+1)};
}
function guideIndex(head,pattern){return head.findIndex(v=>pattern.test(v.replace(/\s+/g,'')))}
impDemandTable=function(data,source){
 const {head,records}=guideHeader(data,'demand',source);
 const barcodeIdx=guideIndex(head,/^(條碼|SKU|品號)$/i),nineIdx=guideIndex(head,/^(貨號|9碼貨號|九碼貨號|SPU)$/i);
 const skuIdx=barcodeIdx>=0?barcodeIdx:nineIdx;
 const storeIdx=guideIndex(head,/^(店名|門市|門市名稱)$/),zoneIdx=guideIndex(head,/^(區域|倉庫\/陳列|分區)$/),floorIdx=guideIndex(head,/^樓層$/);
 const sourceStore=impGuessedStore(source),qtyIdx=guideIndex(head,/^(門市)?補貨需求(數|數量)?$|營運確認數|^【2】實體門市叫貨營運確認數$|^(需求|數量|叫貨數量)$/);
 const storeCols=[];
 for(let i=0;i<head.length;i++)for(const store of stores){
  const col=head[i].replace(/\s+/g,'');
  if(col===store||col.startsWith(store+'_'))storeCols.push({i,store,zone:col===store?'':col.slice(store.length+1)});
 }
 if(skuIdx<0||(!storeCols.length&&qtyIdx<0))throw Error(source+'：找不到條碼或補貨需求欄位');
 if(!storeCols.length&&storeIdx<0&&!sourceStore)throw Error(source+'：無法確認門市');
 const out=[];let noCode=0;
 for(const [j,row] of records.entries()){
  const sku=String(row[skuIdx]||'').trim();if(!sku)continue;
  const choices=storeCols.length?storeCols.map(c=>({store:c.store,zone:c.zone,value:row[c.i]})):[{
   store:storeIdx<0?sourceStore:String(row[storeIdx]||'').trim(),
   zone:zoneIdx<0?impGuessedZone(source,sourceStore):String(row[zoneIdx]||'').trim(),
   value:row[qtyIdx]}];
  for(const c of choices){
   if(c.value==null||String(c.value).trim()==='')continue;
   const qty=impWhole(c.value,'需求 '+sku);
   if(qty===0)continue;
   if(!stores.includes(c.store))throw Error(source+'：未知門市 '+c.store);
   let zone=c.zone;
   if(mode==='restock'&&c.store==='光復'&&floorIdx>=0){
    const floor=String(row[floorIdx]??'').trim();
    if(floor==='2'||floor==='3')zone=floor+'F倉庫';
    else if(!zone)throw Error(source+'：光復 '+sku+' 有補貨需求，但樓層不是 2／3');
   }
   const explicitNine=guide9(row,head,sku);
   if(!explicitNine){noCode++;continue}
   const barcode=barcodeIdx<0?'':String(row[barcodeIdx]||'').trim();
   const longSku=/^[A-Z0-9]{13}$/i.test(barcode)?barcode:sku;
   out.push({sku:longSku,spu:explicitNine,lookup9:explicitNine,store:c.store,zone,
     size:impField(row,head,'size')||(head.includes('尺碼')?String(row[head.indexOf('尺碼')]||''):''),
     color:impField(row,head,'color'),name:impField(row,head,'name'),
     period:impField(row,head,'period'),qty,source});
  }
 }
 if(noCode)throw Error(source+'：'+noCode+' 筆需求缺少有效 9 碼貨號或 13 碼條碼');
 if(!out.length)throw Error(source+'：沒有補貨需求大於 0 的商品');
 return out;
};
impStockTable=function(data,source){
 const {head,records}=guideHeader(data,'stock',source);
 const skuCol=guideIndex(head,/^(條碼|SKU|品號|商品SKU|商品編號|\*貨號\(13碼\))$/i);
 const alternative=guideIndex(head,/^(貨號|9碼貨號)$/i);
 const stockCol=guideIndex(head,/^(庫存|現貨庫存|可用庫存|庫存數量|現貨數量|可配庫存)$/i);
 const skuIndex=skuCol>=0?skuCol:alternative;
 if(skuIndex<0||stockCol<0)throw Error(source+'：找不到庫存貨號或數量欄');
 const out=[];
 for(const row of records){
  const sku=String(row[skuIndex]??'').trim();if(!sku)continue;
  const val=String(row[stockCol]??'').trim();
  if(!val)continue;
  const qty=impWhole(val,'庫存 '+sku);
  const nine=guide9(row,head,sku);
  if(!nine)throw Error(source+'：庫存 '+sku+' 缺少 9 碼貨號及有效條碼');
  out.push({sku,spu:nine,stock:qty,size:impField(row,head,'size'),
    name:impField(row,head,'name'),color:impField(row,head,'color'),source});
 }
 if(!out.length)throw Error(source+'：沒有可用庫存');
 return out;
};
const guideBaseTotals=impTotals;
impTotals=function(){
 const t=guideBaseTotals();
 // 後台查詢完整九碼貨號，可能包含本次不需配貨的尺寸/顏色。
 t.warnings=t.warnings.filter(w=>!w.startsWith('庫存未對應需求：'));
 return t;
};
function guideList(){
 const demand=importState.demand.flatMap(f=>f.entries);
 const seen=new Set(),codes=[];
 for(const r of demand){const code=guideCode(r.lookup9||r.spu);if(code&&!seen.has(code)){seen.add(code);codes.push(code)}}
 return codes;
}
function guideQuery(batch){return '貨號:'+guideList().slice(batch*100,batch*100+100).join(' ')}
function guideCopy(){const text=guideQuery(guided.batch);if(!text||text==='貨號:')return;
 if(navigator.clipboard&&window.isSecureContext){navigator.clipboard.writeText(text).then(()=>guideAlert('已複製第 '+(guided.batch+1)+' 批')).catch(()=>guideFallbackCopy(text))}
 else guideFallbackCopy(text);
}
function guideFallbackCopy(text){
 const box=document.createElement('textarea');box.value=text;document.body.appendChild(box);box.select();
 try{document.execCommand('copy');guideAlert('已複製')}catch(e){guideAlert('請手動選取複製')}
 box.remove();
}
function guideAlert(text){const el=document.getElementById('guideNotice');if(el)el.textContent=text}
function guideDemandIssues(){
 const files=importState.demand;
 const selected=files.flatMap(f=>f.entries);
 const warnings=[],keys=new Set();
 for(const f of files)if(f.sheets?.length>1&&f.selected==='')warnings.push(f.name+' 尚未選擇工作表');
 for(const r of selected){const k=[r.sku,r.store,r.zone].join('|');if(keys.has(k))warnings.push('重複需求：'+r.sku+'／'+r.store+(r.zone?'／'+r.zone:''));keys.add(k)}
 return warnings;
}
function guideNext(target){
 if(target===2){
  if(mode==='new'&&!newConfirmed){guideAlert('請先確認本次新品配貨門市');return}
  if(!importState.demand.some(f=>f.entries.length)||guideDemandIssues().length){guideAlert('請完成工作表選擇及需求核對');return}
 }
 if(target===3&&!guideList().length){guideAlert('尚無可查詢的 9 碼貨號');return}
 if(target===3&&mode==='new'&&!newConfirmed){guideAlert('請先確認本次新品配貨門市');return}
 if(target===4){
  const t=impTotals();if(!t.demand.length||!t.stock.length||t.warnings.length||importState.errors.length){guideAlert('請先完成需求與庫存比對');return}
 }
 guided.stage=target;guided.batch=0;guideRender();
 document.getElementById('guidedArea')?.scrollIntoView({behavior:'smooth',block:'start'});
}
function guideBack(){guided.stage=Math.max(1,guided.stage-1);guideRender()}
function guideFileMarkup(kind){
 const files=importState[kind];
 return files.map((f,i)=>'<div class="guide-file"><span>'+esc(f.name)+'</span>'+
 (f.sheets?.length>1?'<select aria-label="工作表" onchange="selectImportedSheet(\''+kind+'\','+i+',this.value)"><option value="">選擇本次工作表</option>'+
 f.sheets.map((sh,j)=>'<option value="'+j+'" '+(f.selected===String(j)?'selected':'')+'>'+esc(sh.name)+'</option>').join('')+
 (kind==='stock'?'<option value="all" '+(f.selected==='all'?'selected':'')+'>全部工作表</option>':'')+
 '</select>':f.sheets?.length===1?'<span class="sub">'+esc(f.sheets[0].name)+'</span>':'')+
 '<span class="sub">'+f.entries.length+' 筆</span><button onclick="removeImported(\''+kind+'\','+i+')">移除</button></div>').join('');
}
function guideRender(){
 const root=document.getElementById('guidedArea');if(!root)return;
 const demand=importState.demand.flatMap(f=>f.entries),stock=importState.stock.flatMap(f=>f.entries);
 const codes=guideList(),count=Math.ceil(codes.length/100),errors=guideDemandIssues(), t=impTotals();
 const stage=guided.stage,labels=['需求','貨號','庫存','配貨'];
 let content='';
 if(stage===1)content='<h3>01｜匯入門市需求</h3><p class="sub">選擇補貨 Excel，可多檔匯入；每份檔案選擇本次日期工作表。</p>'+
 '<label class="guide-upload">選擇需求檔案 <input aria-label="選擇需求檔案" type="file" multiple accept=".xlsx,.csv" onchange="importFiles(\'demand\',this.files);this.value=\'\'"></label>'+
 guideFileMarkup('demand')+
 '<p class="sub">'+(demand.length?'已辨識 '+demand.length+' 筆需求':'尚未匯入需求')+'</p>'+
 (errors.length?'<p class="warn">'+esc(errors.slice(0,3).join('；'))+'</p>':'')+
 '<button class="primary" onclick="guideNext(2)" '+(!demand.length||errors.length||(mode==='new'&&!newConfirmed)?'disabled':'')+'>確認需求並整理貨號 →</button>';
 if(stage===2){
  guided.batch=Math.min(guided.batch,Math.max(0,count-1));
  const visible=codes.slice(guided.batch*100,(guided.batch+1)*100);
  content='<h3>02｜後台貨號查詢</h3><p class="sub">不重複 9 碼貨號 '+codes.length+' 個／'+count+' 批，每批最多 100 個。</p>'+
  '<div class="guide-batches">'+Array.from({length:count},(_,i)=>'<button class="'+(i===guided.batch?'primary':'')+'" onclick="guided.batch='+i+';guideRender()">第 '+(i+1)+' 批（'+Math.min(100,codes.length-i*100)+'）</button>').join('')+'</div>'+
  '<div class="guide-copy">'+esc('貨號:'+visible.join(' '))+'</div>'+
  '<button class="primary" onclick="guideCopy()">複製第 '+(guided.batch+1)+' 批</button>'+
  '<p class="sub">複製後到後台查庫存，查完再上傳庫存檔案。</p>'+
  '<div class="guide-nav"><button onclick="guideBack()">← 返回需求</button><button class="primary" onclick="guideNext(3)">下一步：匯入庫存 →</button></div>';
 }
 if(stage===3)content='<h3>03｜匯入後台庫存</h3><p class="sub">查完所有批次貨號後，上傳後台庫存，可多檔、分次補充。</p>'+
 '<label class="guide-upload">選擇庫存檔案 <input aria-label="選擇庫存檔案" type="file" multiple accept=".xlsx,.csv" onchange="importFiles(\'stock\',this.files);this.value=\'\'"></label>'+
 guideFileMarkup('stock')+
 '<p class="sub">需求 '+t.demand.length+' 筆｜庫存 '+stock.length+' 筆｜待對應 '+t.warnings.filter(x=>x.startsWith('缺少庫存')).length+' 筆</p>'+
 (t.warnings.length?'<p class="warn">'+esc(t.warnings.slice(0,5).join('；'))+'</p>':'')+
 '<div class="guide-nav"><button onclick="guideBack()">← 返回貨號</button><button class="primary" onclick="guideNext(4)" '+(!stock.length||t.warnings.length?'disabled':'')+'>核對庫存 →</button></div>';
 if(stage===4)content='<h3>04｜確認並開始配貨</h3><p>需求 '+t.demand.length+' 筆、庫存 '+stock.length+' 筆</p>'+
 (t.warnings.length?'<p class="warn">'+esc(t.warnings.slice(0,5).join('；'))+'</p>':'<p class="sub">已完成貨號及庫存比對，可進入配貨預覽。</p>')+
 '<div class="guide-nav"><button onclick="guideBack()">← 返回庫存</button><button class="primary" onclick="startFromFiles()" '+(t.warnings.length?'disabled':'')+'>開始配貨 →</button></div>';
 root.innerHTML='<div class="guide-steps">'+labels.map((l,i)=>'<div class="'+(i+1===stage?'is-active':i+1<stage?'is-done':'')+'">'+(i+1)+' '+l+'</div>').join('')+'</div>'+
 content+'<div id="guideNotice" class="sub">'+esc(guided.error||'')+'</div>';
}
const guideOldImport=importFiles;
importFiles=async function(kind,files){
 if(!files?.length)return;
 guided.error='';
 const array=Array.from(files);
 for(const file of array){
  if(importState[kind].some(f=>f.name===file.name&&f.size===file.size&&f.modified===file.lastModified)){
   guided.error=file.name+'：已匯入相同檔案';continue;
  }
  try{
   if(/\.xlsx$/i.test(file.name)){
    const book=await guideOpenWorkbook(file);
    if(!book.tabs.length)throw Error('找不到工作表');
    const selected=book.tabs.length===1?'0':'';
    const f={name:file.name,size:file.size,modified:file.lastModified,sheets:book.tabs.map(t=>({name:t.name,entries:[]})),
      selected,entries:[],book,kind};
    if(selected==='0')f.entries=await guideReadSheet(f,0);
    importState[kind].push(f);
   }else{
    const tabs=await impRead(file);
    const list=tabs.flatMap(sheet=>kind==='demand'?impDemandTable(sheet.rows,file.name+' / '+sheet.name):impStockTable(sheet.rows,file.name+' / '+sheet.name));
    importState[kind].push({name:file.name,size:file.size,modified:file.lastModified,entries:list,sheets:[{name:file.name,entries:list}],selected:'0'});
   }
  }catch(err){guided.error=file.name+'：'+err.message}
 }
 importState.errors=guided.error?[guided.error]:[];
 guideRender();
};
const guideOldSelect=selectImportedSheet;
selectImportedSheet=async function(kind,index,value){
 const file=importState[kind][index];if(!file)return;
 const previous=file.selected;
 file.selected=value;file.entries=[];
 try{
  if(file.book){
   if(value==='all'&&kind==='stock'){
    for(let i=0;i<file.book.tabs.length;i++)file.entries.push(...await guideReadSheet(file,i));
   }else if(value!=='')file.entries=await guideReadSheet(file,Number(value));
  }else{file.entries=value==='all'?file.sheets.flatMap(s=>s.entries):value===''?[]:(file.sheets[Number(value)]?.entries||[])}
  guided.error='';importState.errors=[];
 }catch(err){file.selected='';file.entries=[];guided.error=file.name+'：'+err.message;importState.errors=[guided.error]}
 guideRender();
};
const guideOldRemove=removeImported;
removeImported=function(kind,index){guideOldRemove(kind,index);guided.error='';guideRender()};
const guideOldReset=clearImportState;
clearImportState=function(){guideOldReset();guided.stage=1;guided.batch=0;guided.error='';guideRender()};
const guideOldSetMode=setMode;
setMode=function(m){guideOldSetMode(m);guided.stage=1;guideRender()};
const guideOldConfirmed=confirmNewEntry;
confirmNewEntry=function(){guideOldConfirmed();guideRender()};
renderNewEntry=function(){
 document.getElementById('newEntry')?.remove();
 if(mode!=='new')return;
 const home=document.getElementById('home'),bar=document.createElement('div');
 bar.id='newEntry';bar.className='box';bar.style.margin='12px 0';
 bar.innerHTML='<div class="row" style="gap:12px;align-items:center"><b>本次新品配貨門市</b><span>預設 8 間</span>'+
 '<label><input type="checkbox" '+(includeDalong?'checked':'')+' onchange="includeDalong=this.checked;newConfirmed=false;renderNewEntry();guideRender()"> 加入大隆</label>'+
 '<span>共 '+(includeDalong?9:8)+' 間</span>'+
 '<button class="primary" onclick="confirmNewEntry()">確認門市</button>'+
 '<span class="sub">'+(newConfirmed?'已確認':'尚未確認')+'</span></div>';
 const before=document.getElementById('guidedArea')||home.querySelector('.guide-tools')||home.querySelector('.row');
 if(before?.parentElement===home)home.insertBefore(bar,before);else home.appendChild(bar);
};

function guideMount(){
 const home=document.getElementById('home'),panel=document.getElementById('importPanel');
 if(!home||!panel)return;
 const wrapper=panel.parentElement;wrapper.style.display='none';
 const first=Array.from(home.querySelectorAll('.row')).find(x=>!x.closest('#newEntry')&&!x.closest('#importPanel'));
 if(first&&first.closest('#home')===home){
  const tools=document.createElement('details');tools.className='guide-tools';
  tools.innerHTML='<summary>測試工具</summary>';
  tools.appendChild(first);home.appendChild(tools);
 }
 const area=document.createElement('div');area.id='guidedArea';area.className='guide-main';home.insertBefore(area,wrapper);
 const css=document.createElement('style');
 css.textContent='.guide-main{padding:16px 0}.guide-main h3{margin:16px 0 9px}.guide-steps{display:grid;grid-template-columns:repeat(4,1fr);gap:7px;margin:12px 0 19px}.guide-steps>div{border-radius:8px;text-align:center;background:#f3f3f3;padding:9px 3px;font-size:13px;color:#888}.guide-steps .is-active{background:#202020;color:white}.guide-steps .is-done{color:#21704b;background:#e7f4ed}.guide-upload{border:1px solid #ccc;padding:12px;border-radius:9px;display:block;margin:10px 0}.guide-upload input{display:block;margin-top:8px;max-width:100%}.guide-file{display:flex;align-items:center;flex-wrap:wrap;gap:8px;padding:8px 0;border-bottom:1px solid #eee;font-size:13px}.guide-file span:first-child{flex:1;min-width:190px;overflow-wrap:anywhere}.guide-file select{max-width:200px}.guide-batches{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:10px}.guide-copy{background:#f5f5f5;border:1px solid #dedede;border-radius:8px;padding:12px;max-height:150px;overflow:auto;white-space:normal;overflow-wrap:anywhere;line-height:1.8;font-size:13px;margin:12px 0}.guide-nav{display:flex;flex-wrap:wrap;gap:9px;justify-content:space-between;margin:16px 0}.guide-tools{margin-top:25px;font-size:12px;opacity:.8}.guide-main button:disabled{opacity:.4;cursor:not-allowed}';
 document.head.appendChild(css);
 guideRender();
}
guideMount();
