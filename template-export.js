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
 const labels=[['信義:warehouse','信義實業・1F倉庫'],['ATT:warehouse','ATT美學・B1F倉庫'],['ATT:display','ATT美學・1F陳列']];
 e.innerHTML='<b>本機範本狀態：</b>'+labels.map(([key,name])=>{
  const f=originalTemplateFiles.get(key);return '<span class="pill">'+name+'：'+(f?'✓ 已保存':'尚未設定')+'</span>'
 }).join('')+'<div class="sub">'+(note||'首次選取一次後會保存在目前裝置的瀏覽器，下次可直接使用；換裝置、無痕模式或清除網站資料可能需要重選。')+'</div>';
}
async function restoreOriginalTemplates(){
 try{
  const labels=['信義:warehouse','ATT:warehouse','ATT:display'];
  for(const key of labels){const value=await localTemplateOperation('get',key);if(value instanceof Blob)originalTemplateFiles.set(key,value)}
  displayTemplateStatus();
 }catch(err){displayTemplateStatus('本機保存無法使用：本次可選取範本，但下次可能需重新選取。')}
}
templateReadyPromise=restoreOriginalTemplates();
async function inspectOriginalTemplates(event){
 await templateReadyPromise;
 const input=event.target;const selected=Array.from(input.files||[]);
 if(!selected.length)return;
 let added=0,failed=[];
 for(const file of selected){
  if(!/\.xlsx$/i.test(file.name)){failed.push(file.name+'（非 xlsx）');continue}
  const key=/PA\s*信義店/.test(file.name)?'信義':(/PA\s*ATT店/i.test(file.name)?'ATT':null);
  if(!key){failed.push(file.name+'（尚未支援此門市範本）');continue}
  const kind=/陳列/.test(file.name)?'display':'warehouse';
  const id=key+':'+kind;
  originalTemplateFiles.set(id,file);added++;
  try{await localTemplateOperation('put',id,file)}catch(e){failed.push(file.name+'（本機保存失敗，僅本次可使用）')}
 }
 displayTemplateStatus(added?'本次新增／取代 '+added+' 份範本。'+(failed.length?'注意：'+failed.join('；'):''):(failed.join('；')||'沒有辨識到可用範本'));
 input.value='';
}
async function clearLocalTemplates(){
 if(!confirm('確定清除儲存在本裝置瀏覽器的原始 Excel 範本？'))return;
 await templateReadyPromise;
 try{await localTemplateOperation('clear');originalTemplateFiles.clear();displayTemplateStatus('已清除本機保存的範本。')}catch(e){alert('清除本機範本失敗：'+e.message)}
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
 if(original.length<2)throw Error('原始工作表缺少商品範例列');
 const first=original[1],header=original[0];
 const sharedEntry=entries.find(x=>x.name==='xl/sharedStrings.xml');
 let shared=[];
 if(sharedEntry){
  const ssdoc=new DOMParser().parseFromString(dec.decode(await uncompressed(sharedEntry)),'application/xml');
  shared=Array.from(ssdoc.getElementsByTagNameNS(ns,'si')).map(si=>si.textContent||'');
 }
 const headerVals=Array.from(header.getElementsByTagNameNS(ns,'c')).map(c=>{
  const v=c.getElementsByTagNameNS(ns,'v')[0]?.textContent||'';
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
function downloadWorkbook(blob,name){
 const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();
 setTimeout(()=>URL.revokeObjectURL(a.href),15000);
}
async function downloadTemplateBasedFiles(){
 try{
  await templateReadyPromise;
  if(!originalTemplateFiles.size){document.getElementById('originalTemplates')?.focus();displayTemplateStatus('尚未設定原始 Excel 範本。請使用上方「選擇範本」一次設定，系統不會自動從 ChatGPT 附件取得檔案。');return;}
  if(!rows.length)throw Error('尚無配貨資料');
  const per=new Set(rows.map(r=>r.period).filter(Boolean));
  if(per.size>1)throw Error('多個不同檔期需分開產生檔案');
  const period=[...per][0]||'測試';
  const details=[];
  for(const store of ['信義','ATT']){
   const storeIdx=stores.indexOf(store);
   const split=(store==='信義')?[['1F倉庫',null]]:[['B1F倉庫','warehouse'],['1F陳列','display']];
   for(const [zone,kind] of split){
    const t=originalTemplateFiles.get(store+':'+(kind||'warehouse'));
    const records=[];
    for(const r of rows){
     const qty=r.qty[storeIdx]||0;if(!qty)continue;
     let n=qty;
     if(store==='ATT'){
      const size=String(r.size||r.sku.split('-').pop()).toUpperCase();
      if(!['S','M','L','F','XS','XL','XXL'].includes(size))throw Error('SKU '+r.sku+' 缺少可識別尺寸，陳列數量不能猜測');
      const display=Math.min(qty,1);n=kind==='display'?display:qty-display;
     }
     if(n)records.push({sku:r.sku,qty:n,name:r.name,color:r.color,size:r.size||r.sku.split('-').pop()});
    }
    if(!records.length)continue;
    if(!t){displayTemplateStatus('尚缺「'+store+' '+zone+'」原始匯入範本，請從上方選擇檔案。');return;}
    const xlsx=await outputFromOriginal(t,records,store);
    const filename=(store==='ATT'?'PA ATT':'PA信義')+'店_'+zone+'_'+period+(mode==='new'?'新品':'補貨')+'.xlsx';
    details.push({blob:xlsx,name:filename,count:records.reduce((a,r)=>a+r.qty,0)});
   }
  }
  if(!details.length)throw Error('目前無信義／ATT 可輸出的配貨數量');
  for(const file of details)downloadWorkbook(file.blob,file.name);
  const box=document.getElementById('exportStatus');
  if(box)box.textContent='已根據原始範本產出 '+details.length+' 份 Excel（'+details.reduce((s,x)=>s+x.count,0)+' 件）。本輪只支援已提供原檔的信義與 ATT；其他七家未提供對應範本，不能擅自生成正式檔。公司系統匯入仍待驗證。';
 }catch(e){alert('原始範本匯出停止：'+e.message)}
}
