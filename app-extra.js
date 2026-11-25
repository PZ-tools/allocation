/* PAZZO allocation test: store export and compact SPU/SKU controls.
   All processing is client-side. No network requests or third-party scripts. */
(function(){
 const baseRender=render,baseApprove=approve,baseSample=sample,baseImport=importCsv;
 const xml=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
 const utf=s=>new TextEncoder().encode(s);
 function safeName(s){return String(s||'').replace(/[\\/:*?"<>|]/g,'-').slice(0,80)}
 function xlsxCells(name,lines){
  const shared=false;
  const types=lines.map(line=>line.map(v=>v===null||v===undefined?'':v));
  let rowsXml=types.map((line,i)=>'<row r="'+(i+1)+'">'+line.map((v,j)=>{
   if(v===''||v===null)return '';
   let col='',x=j+1;while(x){x--;col=String.fromCharCode(65+x%26)+col;x=Math.floor(x/26)}
   return '<c r="'+col+(i+1)+'" '+(typeof v==='number'?'t="n"><v>'+v+'</v>':'t="inlineStr"><is><t xml:space="preserve">'+xml(v)+'</t></is>')+'</c>';
  }).join('')+'</row>').join('');
  const sheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'+rowsXml+'</sheetData></worksheet>';
  return zip({
   '[Content_Types].xml':'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
   '_rels/.rels':'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
   'xl/workbook.xml':'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="'+xml(name)+'" sheetId="1" r:id="rId1"/></sheets></workbook>',
   'xl/_rels/workbook.xml.rels':'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
   'xl/worksheets/sheet1.xml':sheet
  });
 }
 const crctab=(()=>{const t=[];for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?(c>>>1)^0xEDB88320:c>>>1;t[n]=c>>>0}return t})();
 function crc32(bytes){let c=0xFFFFFFFF;for(const b of bytes)c=crctab[(c^b)&255]^(c>>>8);return(c^0xFFFFFFFF)>>>0}
 function zip(contents){let pieces=[],entries=[],at=0;const u16=(b,x)=>{b.push(x&255,(x>>>8)&255)},u32=(b,x)=>{u16(b,x);u16(b,x>>>16)};
  for(const [name,data] of Object.entries(contents)){const n=utf(name),p=typeof data==='string'?utf(data):data,c=crc32(p),h=[];u32(h,0x04034b50);u16(h,20);u16(h,0x0800);u16(h,0);u16(h,0);u16(h,0);u32(h,c);u32(h,p.length);u32(h,p.length);u16(h,n.length);u16(h,0);pieces.push(new Uint8Array(h),n,p);entries.push({n,p,c,at});at+=h.length+n.length+p.length}
  const cd=[];for(const e of entries){const h=[];u32(h,0x02014b50);u16(h,20);u16(h,20);u16(h,0x0800);u16(h,0);u16(h,0);u16(h,0);u32(h,e.c);u32(h,e.p.length);u32(h,e.p.length);u16(h,e.n.length);u16(h,0);u16(h,0);u16(h,0);u16(h,0);u32(h,0);u32(h,e.at);cd.push(new Uint8Array(h),e.n)}
  const cdLen=cd.reduce((x,b)=>x+b.length,0),end=[];u32(end,0x06054b50);u16(end,0);u16(end,0);u16(end,entries.length);u16(end,entries.length);u32(end,cdLen);u32(end,at);u16(end,0);
  return new Blob([...pieces,...cd,new Uint8Array(end)],{type:'application/zip'});
 }
 function trigger(blob,name){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),10000)}
 function meta(r){
  const suffix=(r.sku||'').split('-').pop(),sz=r.size||(['S','M','L','F','XL'].includes(suffix)?suffix:'');
  return {name:r.name||'',color:r.color||'',size:sz,period:r.period||''}
 }
 function zoneSplit(store,r,qty){
  if(store==='信義')return {'1F倉庫':qty};
  if(store==='光復')return {'2F倉庫':qty}; // 3F warehouse routing requires store-specific rule confirmation
  if(store==='大隆')return {'2F倉庫':qty};
  const m=meta(r),display=store==='南西'?(m.size==='L'?1:2):1;
  const used=Math.min(qty,display);
  const zones={'西門':['4F倉庫','2F陳列'],'ATT':['B1F倉庫','1F陳列'],'文心':['2F倉庫','1F陳列'],'南西':['1F倉庫','1F陳列'],'民族':['1F倉庫','1F陳列'],'板橋':['1F倉庫','1F陳列']}[store];
  if(!zones)return {};
  return {[zones[0]]:qty-used,[zones[1]]:used}
 }
 function buildExports(){
  if(!rows.length)throw Error('尚未匯入資料');
  const issues=rows.filter(r=>r.qty.some((v,j)=>v<0||v>r.demand[j]||!Number.isInteger(v))||r.qty.reduce((a,b)=>a+b,0)>cap(r));
  if(issues.length)throw Error('有 '+issues.length+' 筆數量異常，禁止輸出');
  const all={},audit=[],periods=new Set();
  rows.forEach(r=>{const m=meta(r);if(m.period)periods.add(m.period);
   stores.forEach((s,i)=>{const q=r.qty[i];if(!q)return;
    const zs=zoneSplit(s,r,q);
    for(const [zone,num] of Object.entries(zs)){if(num<=0)continue;const key=s+'|'+zone;(all[key]??=[]).push({sku:r.sku,qty:num,...m});audit.push({sku:r.sku,store:s,zone,qty:num})}
   })
  });
  if(periods.size>1)throw Error('同批資料含有不同檔期，請分檔操作');
  const period=periods.size===1?[...periods][0]:'測試';
  const entries={},totals=[];
  for(const [key,items] of Object.entries(all)){const [store,zone]=key.split('|');
   let lines,tab;
   if(store==='信義'){
    tab='工作表1';
    lines=[['單別','單號','轉出廠別','單據日期','單別','單號','序號','品號','品名','規格','數量','轉出庫','轉入庫','確認碼']];
    items.forEach((v,i)=>lines.push(['1225','','001','','1225','',String(i+1).padStart(4,'0'),v.sku,v.name,[v.color,v.size].filter(Boolean).join('-'),v.qty,'1001S','ST01','Y']));
   }else{
    tab='平台資料';lines=[['品號','數量','是否為贈品(若是請寫Y)','單價(微日、賣鬧適用)']];
    items.forEach(v=>lines.push([v.sku,v.qty,'','']));
   }
   const total=items.reduce((a,b)=>a+b.qty,0);
   const file='PA'+store+'店_'+zone+'_'+safeName(period)+(mode==='new'?'新品':'補貨')+'.xlsx';
   entries[file]=new Uint8Array(); // populated by async Blob extraction below
   totals.push({file,blob:xlsxCells(tab,lines),total,items:items.length});
  }
  const expected=rows.reduce((sum,r)=>sum+r.qty.reduce((a,b)=>a+b,0),0);
  if(audit.reduce((sum,r)=>sum+r.qty,0)!==expected)throw Error('分倉數量不平衡，禁止輸出');
  return {totals,expected,period};
 }
 window.downloadStoreFiles=async function(){
  try{
   const {totals,expected,period}=buildExports();
   if(!totals.length)throw Error('沒有可以輸出的非零配貨數量');
   const contents={};for(const t of totals)contents[t.file]=new Uint8Array(await t.blob.arrayBuffer());
   const report=['PAZZO 配貨檔案測試版｜未經系統匯入驗證','作業：'+(mode==='new'?'新品':'補貨'),'檔期：'+period,'總件數：'+expected,'檔案數：'+totals.length,'','重要：信義依實業樣本、其他門市依 ATT 美學樣本；未提供各店倉庫與平台代碼完整驗證。','','檔案,SKU數,件數',...totals.map(x=>x.file+','+x.items+','+x.total)].join('\r\n');
   contents['輸出核對報告.txt']=report;
   trigger(zip(contents),'PAZZO_'+(mode==='new'?'新品':'補貨')+'_各門市匯入測試檔.zip');
   document.getElementById('exportStatus').textContent='已產生 '+totals.length+' 份 Excel，合計 '+expected+' 件。僅供測試，不代表公司系統匯入驗收通過。';
  }catch(e){alert('匯出失敗：'+e.message)}
 };
 function renderInline(){
  if(step!=='work'||mode!=='new')return;
  const node=document.getElementById('newControls');
  node.innerHTML='<div class="box"><h3>當次新品歸零勾選（SPU／SKU 同列）</h3>'+
   Array.from(new Set(rows.map(r=>r.spu))).map(spu=>{
    const items=rows.map((r,i)=>({...r,i})).filter(r=>r.spu===spu),
      blocked=items.every(r=>r.blocked);
    return '<div class="spu-line"><label class="spu-id"><input type="checkbox" data-spu="'+esc(spu)+'" '+(blocked?'checked':'')+'> '+esc(spu)+'（'+items.length+'）</label><div class="sku-items">'+
    items.map(r=>'<label><input type="checkbox" data-idx="'+r.i+'" '+(r.blocked?'checked':'')+' '+(blocked?'disabled':'')+'> '+esc(meta(r).size||r.sku)+'</label>').join('')+'</div></div>';
   }).join('')+'</div>';
  node.querySelectorAll('[data-spu]').forEach(input=>input.addEventListener('change',e=>{
   const spu=e.currentTarget.dataset.spu;rows.filter(r=>r.spu===spu).forEach(r=>r.blocked=e.currentTarget.checked);allocate();
  }));
  node.querySelectorAll('[data-idx]').forEach(input=>input.addEventListener('change',e=>{
   rows[+e.currentTarget.dataset.idx].blocked=e.currentTarget.checked;allocate();
  }));
 }
 render=function(){baseRender();renderInline()};
 approve=function(){baseApprove();if(step==='result'&&mode==='restock')document.getElementById('missing').innerHTML=''};
 sample=function(){baseSample();rows.forEach(r=>{r.period='DEMO';r.name='模擬商品';r.size=r.sku.split('-').pop();r.color='示意色'});render()};
 // Override import only to retain optional descriptive fields while preserving original CSV validation
 importCsv=function(event){const file=event.target.files?.[0];if(!file)return;const reader=new FileReader();
  reader.onload=()=>{
   try{
    const data=csvParse(String(reader.result).replace(/^\uFEFF/,'')),head=data.shift().map(s=>s.trim());
    for(const c of ['SKU','SPU','庫存',...stores])if(!head.includes(c))throw Error('缺少欄位 '+c);
    const get=(v,k)=>{const i=head.indexOf(k);return i<0?'':String(v[i]??'').trim()};
    rows=data.map(v=>({sku:get(v,'SKU'),spu:get(v,'SPU'),stock:Math.max(0,Math.floor(+get(v,'庫存')||0)),demand:stores.map(s=>Math.max(0,Math.floor(+get(v,s)||0))),qty:[],blocked:false,why:'',period:get(v,'檔期'),name:get(v,'品名'),color:get(v,'顏色'),size:get(v,'尺寸')})).filter(r=>r.sku&&r.spu);
    if(!rows.length)throw Error('沒有有效商品');
    if(new Set(rows.map(r=>r.sku)).size!==rows.length)throw Error('有重複 SKU，請修正後匯入');
    allocate();setStep('work');
   }catch(err){alert('CSV 匯入失敗：'+err.message)}
  };
  reader.readAsText(file,'UTF-8');
 };
 downloadTemplate=function(){
  download('PAZZO_測試匯入範本.csv',makeCsv([['SKU','SPU','檔期','品名','顏色','尺寸','庫存',...stores],['PZ-DEMO-001-S','PZ-DEMO-001','DEMO','模擬商品','灰','S',35,4,3,4,4,7,2,2,3,0]]))
 };
})();