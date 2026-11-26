/* 配貨總覽 Excel v0.8.1：保留原表頭，第二列總數，第三、四列分區件數。 */
function summaryXmlEscape(x){return String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]))}
function summaryExcelCol(i){let out='';for(let x=i+1;x>0;x=Math.floor((x-1)/26))out=String.fromCharCode(65+(x-1)%26)+out;return out}
function summarySheetXml(records,opts={}){
 const headerLimit=opts.headerLimit||1;
 const nameLimit=records.length;
 const cols=Math.max(...records.map(x=>x.length),1);
 const cellRows=records.map((r,ri)=>{
  const cells=r.map((v,ci)=>{
   if(v===null||v===undefined||v==='')return '';
   const reference=summaryExcelCol(ci)+(ri+1),style=ri===0?' s="1"':ri>0&&ri<headerLimit?' s="2"':'';
   if(typeof v==='number'&&Number.isFinite(v))return '<c r="'+reference+'"'+style+'><v>'+v+'</v></c>';
   return '<c r="'+reference+'"'+style+' t="inlineStr"><is><t xml:space="preserve">'+summaryXmlEscape(v)+'</t></is></c>';
  }).join('');
  return '<row r="'+(ri+1)+'"'+(ri===0?' ht="23" customHeight="1"':'')+'>'+cells+'</row>';
 }).join('');
 const widths=Array.from({length:cols},(_,i)=>'<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+(i===0?24:i===1?20:14)+'" customWidth="1"/>').join('');
 return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:'+summaryExcelCol(cols-1)+nameLimit+'"/>'+
  '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'+
  '<cols>'+widths+'</cols><sheetData>'+cellRows+'</sheetData></worksheet>';
}
function summaryXlsxBlob(mainRows,detailRows){
 const entries={
  '[Content_Types].xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
  '_rels/.rels':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  'xl/workbook.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="配貨總覽" sheetId="1" r:id="rId1"/><sheet name="分區明細" sheetId="2" r:id="rId2"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
  'xl/styles.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><color rgb="FF302B29"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF1EDE9"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
  'xl/worksheets/sheet1.xml':summarySheetXml(mainRows,{headerLimit:4}),
  'xl/worksheets/sheet2.xml':summarySheetXml(detailRows)
 };
 const data=Object.entries(entries).map(([name,contents])=>{
  const bytes=enc.encode(contents);
  return {name,nameBytes:enc.encode(name),data:bytes,method:0,csize:bytes.length,usize:bytes.length,crc:crc(bytes)};
 });
 return zipRepack(data);
}
downloadResult=function(){
 try{
  const fullDate=exportDateYMD(),title='PAZZO_'+(mode==='new'?'新品':'補貨')+'_'+fullDate+'.xlsx';
  const totals=stores.map((store,i)=>rows.reduce((sum,r)=>sum+Number(r.qty[i]||0),0));
  const stockTotal=rows.reduce((sum,r)=>sum+Number(r.stock||0),0);
  const used=totals.reduce((a,b)=>a+b,0);
  const zoneTotals=stores.map(store=>summarizeZone(store));
  const header=['SKU','SPU','庫存',...stores,'庫存剩餘'];
  const sumRow=['總數','',stockTotal,...totals,stockTotal-used];
  const warehouseRow=['倉庫','', '',...zoneTotals.map(z=>z.unknown?'待確認':z.warehouse),''];
  const displayRow=['陳列','', '',...zoneTotals.map(z=>z.unknown?'待確認':(z.display>0?z.display:'')),''];
  const mainRows=[header,sumRow,warehouseRow,displayRow,...rows.map(r=>[r.sku,r.spu,r.stock,...stores.map((s,i)=>Number(r.qty[i]||0)),r.stock-r.qty.reduce((a,b)=>a+Number(b||0),0)])];
  const detail=[['SKU','SPU','門市','區域','件數','狀態']];
  for(const r of rows)for(const store of stores){
   const qty=Number(r.qty[stores.indexOf(store)]||0);
   if(!qty)continue;
   const entry=zoneForSKU(r,store);
   if(!entry.verified){
    detail.push([r.sku,r.spu,store,'待確認',qty,'尚未確認分區']);
    continue;
   }
   for(const [zone,n] of Object.entries(entry.values))if(n>0)detail.push([r.sku,r.spu,store,zone,Number(n),'已確認']);
  }
  if(detail.slice(1).reduce((sum,r)=>sum+Number(r[4]||0),0)!==used)throw Error('分區明細總數與配貨總數不一致，停止匯出');
  for(let i=0;i<stores.length;i++){
   const x=zoneTotals[i];if(!x.unknown&&x.warehouse+x.display!==totals[i])throw Error(stores[i]+' 倉庫、陳列合計不符');
  }
  downloadWorkbook(summaryXlsxBlob(mainRows,detail),title);
 }catch(err){alert('配貨總覽 Excel 匯出停止：'+err.message)}
};
