
/* v0.8.2: two-sheet allocation overview; summary totals on row 2 only */
downloadResult=function(){
 try{
  const date=exportDateYMD(),title='PAZZO_'+(mode==='new'?'新品':'補貨')+'_'+date+'.xlsx';
  const totals=stores.map((store,i)=>rows.reduce((sum,r)=>sum+Number(r.qty[i]||0),0));
  const mainRows=[
   ['SKU','SPU','庫存',...stores,'庫存剩餘'],
   ['配貨總數','','',...totals,''],
   ...rows.map(r=>[r.sku,r.spu,r.stock,...stores.map((s,i)=>Number(r.qty[i]||0)),
    r.stock-r.qty.reduce((a,b)=>a+Number(b||0),0)])
  ];
  const checkRows=[['門市','倉庫','陳列','總數']];
  for(let i=0;i<stores.length;i++){
   const sum=summarizeZone(stores[i]);
   const warehouse=sum.unknown?'待確認':(sum.warehouse||'');
   const display=sum.unknown?'待確認':(sum.display||'');
   if(!sum.unknown&&Number(sum.warehouse||0)+Number(sum.display||0)!==totals[i])
    throw Error(stores[i]+' 分區合計與配貨總數不相符');
   checkRows.push([stores[i],warehouse,display,totals[i]]);
  }
  if(checkRows.length!==stores.length+1)throw Error('門市數量確認列數不符');
  const grandTotal=totals.reduce((a,b)=>a+b,0);
  if(grandTotal!==rows.reduce((a,r)=>a+r.qty.reduce((v,n)=>v+Number(n||0),0),0))
   throw Error('Excel 件數總合核對失敗');
  downloadWorkbook(summaryXlsxBlob(mainRows,checkRows),title);
 }catch(err){alert('配貨總覽 Excel 匯出停止：'+err.message)}
};
