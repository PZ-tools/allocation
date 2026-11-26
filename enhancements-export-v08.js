/* v0.8.0: replace legacy export with explicit per-SKU zone routing and single ZIP. */
downloadTemplateBasedFiles=async function(){
 try{
  await templateReadyPromise;
  const missing=['實業','美學'].filter(k=>!originalTemplateFiles.has(k));
  if(missing.length){displayTemplateStatus('尚缺 '+missing.join('、')+' 正式範本，請在上方選擇檔案。');return}
  if(!rows.length)throw Error('沒有配貨資料');
  const fullDate=exportDateYMD(),shortDate=fullDate.slice(4),modeName=mode==='new'?'新品':'補貨';
  const fileRows=new Map(),problems=[];
  for(const [ri,r] of rows.entries()){
   for(const store of stores){
    const qty=Number(r.qty[stores.indexOf(store)]||0);
    if(!qty)continue;
    const info=zoneForSKU(r,store);
    const subtotal=Object.values(info.values).reduce((a,b)=>a+Number(b||0),0);
    if(!info.verified||subtotal!==qty){
     problems.push(store+' '+r.sku+'：分區合計 '+subtotal+'，核准配貨 '+qty);
     continue;
    }
    for(const [zone,value] of Object.entries(info.values)){
     const n=Number(value);if(!n)continue;
     if(!Number.isInteger(n)||n<0)throw Error('分區件數無效 '+store+' '+r.sku);
     const key=store+'|'+zone;
     if(!fileRows.has(key))fileRows.set(key,[]);
     fileRows.get(key).push({sku:r.sku,qty:n,name:r.name||'',color:r.color||'',size:r.size||(/[A-Z]+$/.exec(r.sku||'')||[])[0]||''});
    }
   }
  }
  if(problems.length){
   renderBreakdown();
   const details=problems.slice(0,4).join('；');
   throw Error('有 '+problems.length+' 筆補貨／新品分區尚未確認：'+details+'。請先在結果總覽填寫正確倉庫／陳列數量。');
  }
  const expected=rows.reduce((sum,r)=>sum+r.qty.reduce((a,b)=>a+b,0),0);
  if(!expected)throw Error('目前配貨總件數為 0，沒有打單檔案');
  const files=[];
  for(const [key,rec] of fileRows.entries()){
   const [store,zone]=key.split('|'),system=store==='信義'?'實業':'美學';
   const blob=await outputFromOriginal(originalTemplateFiles.get(system),rec,store==='信義'?'信義':'美學');
   const count=rec.reduce((a,x)=>a+x.qty,0);
   const verified=await verifyExportedWorkbook(blob,rec.length,count,store==='信義'?'信義':'美學');
   if(verified.some((x,i)=>x!==rec[i].sku))throw Error(store+' '+zone+' SKU 核對失敗，停止下載');
   const filename=(store==='ATT'?'PA ATT':'PA'+store)+'店_'+zone+'_'+shortDate+modeName+'.xlsx';
   files.push({name:filename,blob,count,store,zone,skus:rec.length});
  }
  const actual=files.reduce((a,f)=>a+f.count,0);
  if(actual!==expected)throw Error('匯出合計 '+actual+' 與配貨核准 '+expected+' 不符，停止下載');
  const contents=[],used=new Set();
  for(const f of files){
   if(used.has(f.name))throw Error('重複檔名 '+f.name);
   used.add(f.name);
   const bin=new Uint8Array(await f.blob.arrayBuffer());
   if(bin.length<4||u32(bin,0)!==0x04034b50)throw Error('非有效 Excel 檔 '+f.name);
   contents.push({name:f.name,nameBytes:enc.encode(f.name),data:bin,method:0,csize:bin.length,usize:bin.length,crc:crc(bin)});
  }
  const byStore=stores.map(store=>{
   const i=stores.indexOf(store),q=rows.reduce((sum,r)=>sum+(r.qty[i]||0),0);
   const n=files.filter(f=>f.store===store).length;
   return store+'｜'+q+'｜'+n;
  });
  const report=['配貨小幫手GOGOGO｜各店打單檔核對','模式：'+modeName,'打單／檔期日期：'+fullDate,
   '資料列數：'+rows.length,'總配貨件數：'+actual,'Excel檔案數：'+files.length,
   allocationTestOnly?'注意：這是模擬資料，倉庫分區僅用於 ZIP 功能測試。':'注意：正式公司系統匯入仍未通過驗證。',
   '','門市｜配貨件數｜產生檔案數',...byStore,'','檔名｜SKU 數｜件數',...files.map(f=>f.name+'｜'+f.skus+'｜'+f.count)].join('\r\n');
  const txt=enc.encode(report);
  contents.push({name:'打單檔案核對.txt',nameBytes:enc.encode('打單檔案核對.txt'),data:txt,method:0,csize:txt.length,usize:txt.length,crc:crc(txt)});
  const packed=zipRepack(contents),check=zipRead(new Uint8Array(await packed.arrayBuffer()));
  if(check.length!==contents.length||check.some((it,i)=>it.name!==contents[i].name||it.crc!==contents[i].crc||it.usize!==contents[i].usize))
   throw Error('ZIP 內容校驗失敗，停止下載');
  downloadWorkbook(packed,'PAZZO_'+modeName+'_'+fullDate+'.zip');
  const state=document.getElementById('exportStatus');
  if(state)state.textContent='ZIP 已下載：'+files.length+' 份 Excel，共 '+actual+' 件，含全部有實際配貨的門市與分區。正式系統相容性尚待驗證。';
 }catch(err){alert('打單檔案匯出停止：'+err.message)}
};
