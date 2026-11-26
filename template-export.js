/* 原始範本只儲存在本裝置瀏覽器的 IndexedDB，沒有任何上傳。 */
const originalTemplateFiles = new Map();
const templateDbName = 'pz-allocation-local-templates-v1';
let templateReadyPromise;
function openTemplateDb(){
 return new Promise((resolve,reject)=>{
  if(!('indexedDB' in window)){reject(new Error('此瀏覽器不支援本機範本保存'));return}
  const r=indexedDB.open(templateDbName,1);
  r.onupgradeneeded=()=>{const db=r.result;if(!db.objectStoreNames.contains('templates'))db.createObjectStore('templates')};
  r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);
 });
}
function localTemplateOperation(type,key,value){
 return openTemplateDb().then(db=>new Promise((resolve,reject)=>{
  const tx=db.transaction('templates',type==='get'?'readonly':'readwrite');
  const st=tx.objectStore('templates');
  const req=type==='put'?st.put(value,key):type==='clear'?st.clear():st.get(key);
  req.onsuccess=()=>resolve(req.result);
  req.onerror=()=>reject(req.error);
  tx.oncomplete=()=>db.close();
  tx.onerror=()=>{db.close();reject(tx.error)};
 }));
}
function displayTemplateStatus(note=''){
 const e=document.getElementById('templateInfo');if(!e)return;
 const labels=[['實業','實業轉撥單模版'],['美學','美學平台資料模版']];
 e.innerHTML='<b>本機範本狀態：</b>'+labels.map(([key,name])=>{
  const f=originalTemplateFiles.get(key);return '<span class="pill">'+name+'：'+(f?'✓ 已保存':'尚未設定')+'</span>'
 }).join('')+'<div class="sub">'+(note||'兩份範本各設定一次，之後於相同瀏覽器自動沿用；換裝置或清除網站資料時須重選。')+'</div>';
}
async function restoreOriginalTemplates(){
 try{
  for(const key of ['實業','美學']){const v=await localTemplateOperation('get','format:'+key);if(v instanceof Blob)originalTemplateFiles.set(key,v)}
  displayTemplateStatus();
 }catch(err){displayTemplateStatus('無法讀取本機範本：本次可手動選取，後續可能需要重選。')}
}
templateReadyPromise=restoreOriginalTemplates();
async function inspectOriginalTemplates(event){
 await templateReadyPromise;
 const input=event.target,selected=Array.from(input.files||[]);
 if(!selected.length)return;
 const failed=[];let added=0;
 for(const file of selected){
  if(!/\\.xlsx$/i.test(file.name)){failed.push(file.name+'（不是 xlsx）');continue}
  try{
   const type=await identifyOriginalWorkbook(file);
   if(!type){failed.push(file.name+'（欄位不符合實業或美學正式範本）');continue}
   originalTemplateFiles.set(type,file);added++;
   try{await localTemplateOperation('put','format:'+type,file)}catch(e){failed.push(file.name+'（可本次使用，但保存至本機失敗）')}
  }catch(err){failed.push(file.name+'（'+err.message+'）')}
 }
 displayTemplateStatus((added?'已設定 '+added+' 種原始範本。':'沒有新增可用範本。')+(failed.length?' '+failed.join('；'):''));
 input.value='';
}
async function clearLocalTemplates(){
 if(!confirm('確定清除本裝置已保存的原始 Excel 範本？'))return;
 await templateReadyPromise;
 try{await localTemplateOperation('clear');originalTemplateFiles.clear();displayTemplateStatus('已清除本機範本。')}
 catch(e){alert('清除失敗：'+e.message)}
}
function u16(d,o){return d[o]|d[o+1]<<8}
function u32(d,o){return (d[o]|d[o+1]<<8|d[o+2]<<16|d[o+3]<<24)>>>0}
function put16(a,v){a.push(v&255,(v>>>8)&255)}
function put32(a,v){put16(a,v);put16(a,v>>>16)}
const enc=new TextEncoder(),dec=new TextDecoder();
function zipRead(bytes){
 let end=-1;
 for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--){if(u32(bytes,i)===0x06054b50){end=i;break}}
 if(end<0)throw Error('非有效 XLSX ZIP 結構');
 const count=u16(bytes,end+10),start=u32(bytes,end+16),items=[];
 let at=start;
 for(let i=0;i<count;i++){
  if(u32(bytes,at)!==0x02014b50)throw Error('工作簿壓縮目錄異常');
  const method=u16(bytes,at+10),crc=u32(bytes,at+16),csize=u32(bytes,at+20),usize=u32(bytes,at+24);
  const nlen=u16(bytes,at+28),elen=u16(bytes,at+30),clen=u16(bytes,at+32),offset=u32(bytes,at+42);
  const nameBytes=bytes.slice(at+46,at+46+nlen),name=dec.decode(nameBytes);
  if(u32(bytes,offset)!==0x04034b50)throw Error('範本壓縮檔頭異常：'+name);
  const dataAt=offset+30+u16(bytes,offset+26)+u16(bytes,offset+28);
  items.push({name,nameBytes,method,crc,csize,usize,data:bytes.slice(dataAt,dataAt+csize)});
  at+=46+nlen+elen+clen;
 }
 return items;
}
async function uncompressed(e){
 if(e.method===0)return e.data;
 if(e.method!==8)throw Error('範本使用不支援的壓縮方式');
 if(typeof DecompressionStream==='undefined')throw Error('瀏覽器不支援直接讀取原始 Excel；請改用新版 Safari 或 Chrome');
 let stream;
 try{stream=new Blob([e.data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))}
 catch(err){throw Error('此瀏覽器無法解壓 XLSX 範本，請改用最新版 Chrome')}
 return new Uint8Array(await new Response(stream).arrayBuffer());
}
function zipRepack(entries){
 const local=[],central=[];let offset=0;
 for(const e of entries){
  const n=e.nameBytes,h=[];put32(h,0x04034b50);put16(h,20);put16(h,0x800);put16(h,e.method);
  put16(h,0);put16(h,0);put32(h,e.crc);put32(h,e.csize);put32(h,e.usize);put16(h,n.length);put16(h,0);
  local.push(new Uint8Array(h),n,e.data);
  const c=[];put32(c,0x02014b50);put16(c,20);put16(c,20);put16(c,0x800);put16(c,e.method);
  put16(c,0);put16(c,0);put32(c,e.crc);put32(c,e.csize);put32(c,e.usize);
  put16(c,n.length);put16(c,0);put16(c,0);put16(c,0);put16(c,0);put32(c,0);put32(c,offset);
  central.push(new Uint8Array(c),n);
  offset+=h.length+n.length+e.data.length;
 }
 const csize=central.reduce((a,x)=>a+x.length,0),tail=[];
 put32(tail,0x06054b50);put16(tail,0);put16(tail,0);put16(tail,entries.length);put16(tail,entries.length);
 put32(tail,csize);put32(tail,offset);put16(tail,0);
 return new Blob([...local,...central,new Uint8Array(tail)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
const crcTable=Array.from({length:256},(_,i)=>{let c=i;for(let n=0;n<8;n++)c=c&1?(c>>>1)^0xEDB88320:c>>>1;return c>>>0});
function crc(bytes){let c=0xFFFFFFFF;for(const b of bytes)c=crcTable[(c^b)&255]^(c>>>8);return (c^0xFFFFFFFF)>>>0}
function alterCell(doc,row,col,value){
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 let cell=Array.from(row.getElementsByTagNameNS(ns,'c')).find(c=>/^([A-Z]+)/.exec(c.getAttribute('r')||'')?.[1]===col);
 if(!cell){cell=doc.createElementNS(ns,'c');row.appendChild(cell)}
 cell.setAttribute('r',col+row.getAttribute('r'));
 for(const node of Array.from(cell.childNodes))if(node.nodeType===1)cell.removeChild(node);
 if(typeof value==='number'){cell.removeAttribute('t');const v=doc.createElementNS(ns,'v');v.textContent=String(value);cell.appendChild(v)}
 else{cell.setAttribute('t','inlineStr');const is=doc.createElementNS(ns,'is'),t=doc.createElementNS(ns,'t');t.textContent=String(value??'');is.appendChild(t);cell.appendChild(is)}
}
async function outputFromOriginal(file,records,type){
 const bytes=new Uint8Array(await file.arrayBuffer()),entries=zipRead(bytes);
 const target=entries.find(x=>x.name==='xl/worksheets/sheet1.xml');
 if(!target)throw Error('範本沒有第一張工作表：'+file.name);
 const raw=dec.decode(await uncompressed(target)),doc=new DOMParser().parseFromString(raw,'application/xml');
 if(doc.querySelector('parsererror'))throw Error('原始工作表 XML 不可讀');
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 const data=doc.getElementsByTagNameNS(ns,'sheetData')[0];
 if(!data)throw Error('工作表缺少 sheetData');
 const original=Array.from(data.getElementsByTagNameNS(ns,'row'));
 if(original.length<1)throw Error('原始工作表缺少表頭');
 const header=original[0],first=original[1]||doc.createElementNS(ns,'row');
 const sharedEntry=entries.find(x=>x.name==='xl/sharedStrings.xml');
 let shared=[];
 if(sharedEntry){
  const ssdoc=new DOMParser().parseFromString(dec.decode(await uncompressed(sharedEntry)),'application/xml');
  shared=Array.from(ssdoc.getElementsByTagNameNS(ns,'si')).map(si=>si.textContent||'');
 }
 const headerVals=Array.from(header.getElementsByTagNameNS(ns,'c')).map(c=>{
  const v=c.getElementsByTagNameNS(ns,'v')[0]?.textContent||'';
  if(c.getAttribute('t')==='inlineStr')return c.getElementsByTagNameNS(ns,'is')[0]?.textContent||'';
  return c.getAttribute('t')==='s'?shared[Number(v)]:v;
 });
 const expected=type==='信義'?['單別','單號','轉出廠別','單據日期','單別','單號','序號','品號','品名','規格','數量','轉出庫','轉入庫','確認碼']:['品號','數量','是否為贈品(若是請寫Y)','單價(微日、賣鬧適用)'];
 if(expected.some((x,i)=>headerVals[i]!==x))throw Error('範本欄位與原始匯入規格不一致：'+file.name);
 for(const row of original.slice(1))data.removeChild(row);
 records.forEach((r,i)=>{
  const clone=first.cloneNode(true);clone.setAttribute('r',String(i+2));
  const cells=Array.from(clone.getElementsByTagNameNS(ns,'c'));
  for(const c of cells){const a=/^([A-Z]+)/.exec(c.getAttribute('r')||'')?.[1];if(a)c.setAttribute('r',a+(i+2))}
  if(type==='信義'){
   if(!r.name||!r.size)throw Error('信義原始範本需要品名與規格，SKU '+r.sku+' 缺少資料');
   alterCell(doc,clone,'G',String(i+1).padStart(4,'0'));
   alterCell(doc,clone,'H',r.sku);
   alterCell(doc,clone,'I',r.name);
   alterCell(doc,clone,'J',[r.color,r.size].filter(Boolean).join('-'));
   alterCell(doc,clone,'K',r.qty);
  }else{alterCell(doc,clone,'A',r.sku);alterCell(doc,clone,'B',r.qty)}
  data.appendChild(clone);
 });
 const dim=doc.getElementsByTagNameNS(ns,'dimension')[0];
 if(dim)dim.setAttribute('ref',(type==='信義'?'A1:P':'A1:D')+(records.length+1));
 const out=enc.encode(new XMLSerializer().serializeToString(doc));
 target.data=out;target.method=0;target.usize=out.length;target.csize=out.length;target.crc=crc(out);
 return zipRepack(entries);
}
async function identifyOriginalWorkbook(file){
 const bytes=new Uint8Array(await file.arrayBuffer()),entries=zipRead(bytes);
 const target=entries.find(x=>x.name==='xl/worksheets/sheet1.xml');
 if(!target)throw Error('找不到原始工作表');
 const doc=new DOMParser().parseFromString(dec.decode(await uncompressed(target)),'application/xml');
 if(doc.querySelector('parsererror'))throw Error('Excel XML 無法解析');
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 const first=doc.getElementsByTagNameNS(ns,'row')[0];if(!first)return null;
 const sharedEntry=entries.find(x=>x.name==='xl/sharedStrings.xml');
 let shared=[];
 if(sharedEntry){
  const sd=new DOMParser().parseFromString(dec.decode(await uncompressed(sharedEntry)),'application/xml');
  shared=Array.from(sd.getElementsByTagNameNS(ns,'si')).map(x=>x.textContent||'');
 }
 const cells=Array.from(first.getElementsByTagNameNS(ns,'c'));
 const values=cells.map(c=>{
  const v=c.getElementsByTagNameNS(ns,'v')[0]?.textContent||'';
  return c.getAttribute('t')==='inlineStr'?(c.getElementsByTagNameNS(ns,'is')[0]?.textContent||''):(c.getAttribute('t')==='s'?shared[Number(v)]:v);
 });
 const sy=['單別','單號','轉出廠別','單據日期','單別','單號','序號','品號','品名','規格','數量','轉出庫','轉入庫','確認碼'];
 const my=['品號','數量','是否為贈品(若是請寫Y)','單價(微日、賣鬧適用)'];
 if(sy.every((x,i)=>values[i]===x))return '實業';
 if(my.every((x,i)=>values[i]===x))return '美學';
 return null;
}
function downloadWorkbook(blob,name){
 const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();
 setTimeout(()=>URL.revokeObjectURL(a.href),15000);
}
async function downloadTemplateBasedFiles(){
 try{
  await templateReadyPromise;
  const missing=['實業','美學'].filter(key=>!originalTemplateFiles.has(key));
  if(missing.length){displayTemplateStatus('尚缺：'+missing.join('、')+' 範本。請使用上方選擇檔案；只需要這兩種，不需要每家門市各上傳一份。');return}
  if(!rows.length)throw Error('沒有配貨資料');
  const periodSet=new Set(rows.map(r=>r.period).filter(Boolean));
  if(periodSet.size>1)throw Error('不同檔期需分開打單');
  const period=[...periodSet][0]||'測試';
  const recordsByFile=new Map();
  for(const r of rows){
   const size=String(r.size||r.sku.split('-').pop()).toUpperCase();
   for(let i=0;i<stores.length;i++){
    const q=r.qty[i]||0;if(q===0)continue;
    const store=stores[i];
    let zones;
    if(store==='信義')zones={'1F倉庫':q};
    else if(store==='光復'){
      if(mode==='restock')throw Error('光復補貨 2F／3F 倉庫分配未提供，請先確認原始分倉數量，不能自動猜測');
      zones={'2F倉庫':q};
    }else if(store==='大隆')zones={'2F倉庫':q};
    else if(mode==='restock'){
      throw Error(store+' 補貨未提供倉庫／陳列的分配數量，無法安全輸出分區打單檔');
    }else{
      if(!['S','M','L','F','XS','XL','XXL'].includes(size))throw Error('商品 '+r.sku+' 缺少尺寸，無法判斷陳列件數');
      const pair={西門:['4F倉庫','2F陳列'],ATT:['B1F倉庫','1F陳列'],文心:['2F倉庫','1F陳列'],南西:['1F倉庫','1F陳列'],民族:['1F倉庫','1F陳列'],板橋:['1F倉庫','1F陳列']}[store];
      if(!pair)throw Error('找不到 '+store+' 的新品分倉設定');
      const display=Math.min(q,store==='南西'?(size==='L'?1:2):1);
      zones={[pair[0]]:q-display,[pair[1]]:display};
    }
    for(const [zone,num] of Object.entries(zones)){
     if(num<=0)continue;
     const key=store+'|'+zone;
     if(!recordsByFile.has(key))recordsByFile.set(key,[]);
     recordsByFile.get(key).push({sku:r.sku,qty:num,name:r.name||'',color:r.color||'',size});
    }
   }
  }
  const files=[],expected=rows.reduce((a,r)=>a+r.qty.reduce((x,y)=>x+y,0),0);
  for(const [key,rec] of recordsByFile.entries()){
   const [store,zone]=key.split('|'),system=store==='信義'?'實業':'美學';
   const blob=await outputFromOriginal(originalTemplateFiles.get(system),rec,store==='信義'?'信義':'美學');
   const prefix=store==='ATT'?'PA ATT':'PA'+store;
   const name=prefix+'店_'+zone+'_'+period+(mode==='new'?'新品':'補貨')+'.xlsx';
   files.push({blob,name,count:rec.reduce((a,x)=>a+x.qty,0)});
  }
  const actual=files.reduce((a,x)=>a+x.count,0);
  if(actual!==expected)throw Error('輸出總件數 '+actual+' 與核准總件數 '+expected+' 不一致，已中止');
  if(!files.length)throw Error('沒有非零配貨數，無可輸出檔案');
  // 一次下載：將全部門市的 Excel 原檔封裝為單一 ZIP。
  // XLSX 本身也是 ZIP；這裡保留每一份原始範本產出的檔案位元組，絕不重建內容。
  const zipEntries=[];
  const names=new Set();
  for(const f of files){
   if(names.has(f.name))throw Error('匯出檔名重複：'+f.name);
   names.add(f.name);
   const data=new Uint8Array(await f.blob.arrayBuffer());
   if(data.length<4||u32(data,0)!==0x04034b50)throw Error('產生的 Excel 無效：'+f.name);
   zipEntries.push({name:f.name,nameBytes:enc.encode(f.name),data,method:0,csize:data.length,usize:data.length,crc:crc(data)});
  }
  const report=['配貨小幫手GOGOGO｜打單 Excel 測試檔',
    '作業：'+(mode==='new'?'新品':'補貨'),'檔期：'+period,
    '門市：'+[...new Set([...recordsByFile.keys()].map(k=>k.split('|')[0]))].join('、'),
    'Excel 檔案數：'+files.length,'合計件數：'+actual,
    '注意：僅含本次有非零配貨數量的門市分區；不代表正式系統匯入 PASS。',
    '',
    '檔名｜件數',...files.map(f=>f.name+'｜'+f.count)].join('\\r\\n');
  const reportData=enc.encode(report);
  zipEntries.push({name:'打單檔案核對.txt',nameBytes:enc.encode('打單檔案核對.txt'),data:reportData,method:0,csize:reportData.length,usize:reportData.length,crc:crc(reportData)});
  const packed=zipRepack(zipEntries);
  // 回讀 ZIP 目錄，確保檔案數量、名稱與位元組長度一致。
  const verified=zipRead(new Uint8Array(await packed.arrayBuffer()));
  if(verified.length!==zipEntries.length||
     verified.some((e,i)=>e.name!==zipEntries[i].name||e.usize!==zipEntries[i].usize||e.crc!==zipEntries[i].crc))
   throw Error('ZIP 封裝驗證失敗，已停止下載');
  downloadWorkbook(packed,'配貨小幫手GOGOGO_'+(mode==='new'?'新品':'補貨')+'_'+period+'_各店打單檔.zip');
  document.getElementById('exportStatus').textContent='ZIP 已產生：'+files.length+' 份 Excel，涵蓋 '+new Set([...recordsByFile.keys()].map(k=>k.split('|')[0])).size+' 家有配貨門市，共 '+actual+' 件。只下載一次；公司系統匯入仍待驗證。';
 }catch(err){alert('原始範本匯出停止：'+err.message)}
}
