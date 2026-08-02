const PC = {
  LOG: 'ProductionLog',
  MASTER: 'ProductMaster',
  DASHBOARD: 'Dashboard',
  MONTHLY: 'MonthlySummary',
  SETUP: 'Setup',
  BULK_PER_BATCH: 35000,
  HEADERS: [
    'Timestamp','Record ID','Product','Batch','จำนวน Batch',
    'ยอดผลิต (ชิ้น)','น้ำหนักเฉลี่ย (kg/ชิ้น)','ซองเสีย (ใบ)',
    'กล่องเสีย (ใบ)','Bulk ตามสูตร (kg)','Bulk ตามยอดผลิตจริง (kg)',
    'ส่วนต่าง Bulk (kg)','% สูญเสีย','% Standard','ผล',
    'ซองเสีย/10,000 ชิ้น','กล่องเสีย/10,000 ชิ้น',
    'หมายเหตุ','ผู้รายงาน','Source'
  ],
  PRODUCTS: [
    ['Lovely','Hygiene LD Expert Wash Lovely 1400',-0.005],
    ['Morning Fresh','Hg-EPW Morning Fresh 1400',-0.003],
    ['Forever Bloom','Hg-LD EPW Forever Bloom 1400',-0.003],
    ['Sunkiss Blooming','HG EPW Sunkiss Blooming 1400 NEW',-0.003],
    ['Love Touch','HG EPW Love Touch 1400',-0.003],
    ['Spring Magnolia','Hg-EPW Spring Magnolia 1400',-0.005],
    ['Sunrise Kiss','Hg-EPW Sunrise Kiss 1400',-0.005],
    ['Peony Bloom','Hygiene LD EPW Peony Bloom 1400',-0.005],
    ['Milky Touch','Hygiene LD EPW Milky Touch 1400',-0.003],
    ['Happy Sunshine','Hygiene LD EPW Happy Sunshine 1400',-0.010]
  ]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ProdClose 1400')
    .addItem('ติดตั้ง/ซ่อมโครงสร้างระบบ', 'setupSystem')
    .addItem('เปิด Dashboard', 'openDashboard')
    .addToUi();
}

function openDashboard() {
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(ss.getSheetByName(PC.DASHBOARD));
}

function setupSystem() {
  const ss = SpreadsheetApp.getActive();
  setupSystem_(ss);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('ติดตั้ง ProdClose 1400 เรียบร้อย');
}

function setupSystem_(ss) {
  const log = getOrCreateSheet_(ss, PC.LOG);
  const master = getOrCreateSheet_(ss, PC.MASTER);
  const setup = getOrCreateSheet_(ss, PC.SETUP);
  const dashboard = getOrCreateSheet_(ss, PC.DASHBOARD);
  const monthly = getOrCreateSheet_(ss, PC.MONTHLY);

  setupMaster_(master);
  setupSettings_(setup);
  setupLog_(log);
  setupDashboard_(dashboard);
  setupMonthly_(monthly);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function setupMaster_(sheet) {
  sheet.clear();
  sheet.getRange(1,1,1,3).setValues([['Product','ชื่อเต็ม','% Standard สูญเสีย']]);
  sheet.getRange(2,1,PC.PRODUCTS.length,3).setValues(PC.PRODUCTS);
  sheet.setFrozenRows(1);
  sheet.getRange('A1:C1').setBackground('#0F4C81').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange('C2:C11').setNumberFormat('0.000%');
  sheet.setColumnWidth(1,170);
  sheet.setColumnWidth(2,330);
  sheet.setColumnWidth(3,150);
}

function setupSettings_(sheet) {
  sheet.clear();
  sheet.getRange('A1:B4').setValues([
    ['รายการ','ค่า'],
    ['Bulk ต่อ Batch (kg)',PC.BULK_PER_BATCH],
    ['ชื่อระบบ','ProdClose 1400'],
    ['หมายเหตุ','ระบบบันทึก Timestamp อัตโนมัติ ไม่มีช่องวันที่และกะบนหน้าเว็บ']
  ]);
  sheet.getRange('A1:B1').setBackground('#0F4C81').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setColumnWidth(1,200);
  sheet.setColumnWidth(2,430);
}

function setupLog_(sheet) {
  const hasData = sheet.getLastRow() > 1;
  if (!hasData) sheet.clear();
  sheet.getRange(1,1,1,PC.HEADERS.length).setValues([PC.HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,PC.HEADERS.length)
    .setBackground('#0F4C81').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sheet.getRange('E:E').setNumberFormat('0.0');
  sheet.getRange('F:F').setNumberFormat('#,##0');
  sheet.getRange('G:G').setNumberFormat('0.000');
  sheet.getRange('H:I').setNumberFormat('#,##0');
  sheet.getRange('J:L').setNumberFormat('#,##0.00');
  sheet.getRange('M:N').setNumberFormat('0.000%');
  sheet.getRange('P:Q').setNumberFormat('0.00');
  if (!sheet.getFilter()) sheet.getRange(1,1,Math.max(sheet.getMaxRows(),2),PC.HEADERS.length).createFilter();
}

function setupDashboard_(sheet) {
  sheet.clear();
  sheet.getCharts().forEach(c => sheet.removeChart(c));
  sheet.getRange('A1:J1').merge().setValue('ProdClose 1400 Dashboard')
    .setBackground('#0F4C81').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(20).setHorizontalAlignment('center');

  sheet.getRange('A2').setValue('เลือกเดือน');
  sheet.getRange('B2').setValue(new Date(new Date().getFullYear(), new Date().getMonth(), 1)).setNumberFormat('mmmm yyyy');
  sheet.getRange('D2').setValue('เลือก Product');
  sheet.getRange('E2').setValue('ทั้งหมด');
  const productRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ทั้งหมด'].concat(PC.PRODUCTS.map(p => p[0])), true)
    .setAllowInvalid(false).build();
  sheet.getRange('E2').setDataValidation(productRule);
  sheet.getRange('A2:E2').setBackground('#EAF2F8').setFontWeight('bold').setFontColor('#0F4C81');

  const labels = [
    ['A4','จำนวนรายการผลิต'],['C4','จำนวน Batch รวม'],['E4','ยอดผลิตรวม'],['G4','% สูญเสียรวม'],
    ['A7','Bulk ส่วนต่างรวม'],['C7','ผ่านมาตรฐาน'],['E7','เกินมาตรฐาน'],['G7','Pass Rate'],
    ['A10','ซองเสีย/10,000'],['C10','กล่องเสีย/10,000'],['E10','ซองเสียรวม'],['G10','กล่องเสียรวม']
  ];
  labels.forEach(x => sheet.getRange(x[0]).setValue(x[1]).setBackground('#DCEAF4')
    .setFontWeight('bold').setFontColor('#0F4C81').setHorizontalAlignment('center'));

  const d1 = '$B$2';
  const d2 = 'DATE(YEAR($B$2),MONTH($B$2)+1,1)';
  const product = '$E$2';
  const countBase = `COUNTIFS(${PC.LOG}!$A$2:$A,">="&${d1},${PC.LOG}!$A$2:$A,"<"&${d2})`;
  const countProd = `COUNTIFS(${PC.LOG}!$A$2:$A,">="&${d1},${PC.LOG}!$A$2:$A,"<"&${d2},${PC.LOG}!$C$2:$C,${product})`;
  const sumFormula = col =>
    `=IF(${product}="ทั้งหมด",SUMIFS(${PC.LOG}!$${col}$2:$${col},${PC.LOG}!$A$2:$A,">="&${d1},${PC.LOG}!$A$2:$A,"<"&${d2}),SUMIFS(${PC.LOG}!$${col}$2:$${col},${PC.LOG}!$A$2:$A,">="&${d1},${PC.LOG}!$A$2:$A,"<"&${d2},${PC.LOG}!$C$2:$C,${product}))`;

  sheet.getRange('A5').setFormula(`=IF(${product}="ทั้งหมด",${countBase},${countProd})`);
  sheet.getRange('C5').setFormula(sumFormula('E'));
  sheet.getRange('E5').setFormula(sumFormula('F'));
  sheet.getRange('G5').setFormula(`=IFERROR(${sumFormula('L').substring(1)}/${sumFormula('J').substring(1)},"")`);
  sheet.getRange('A8').setFormula(sumFormula('L'));
  sheet.getRange('C8').setFormula(`=IF(${product}="ทั้งหมด",COUNTIFS(${PC.LOG}!$A$2:$A,">="&${d1},${PC.LOG}!$A$2:$A,"<"&${d2},${PC.LOG}!$O$2:$O,"ผ่าน"),COUNTIFS(${PC.LOG}!$A$2:$A,">="&${d1},${PC.LOG}!$A$2:$A,"<"&${d2},${PC.LOG}!$C$2:$C,${product},${PC.LOG}!$O$2:$O,"ผ่าน"))`);
  sheet.getRange('E8').setFormula(`=IF(${product}="ทั้งหมด",COUNTIFS(${PC.LOG}!$A$2:$A,">="&${d1},${PC.LOG}!$A$2:$A,"<"&${d2},${PC.LOG}!$O$2:$O,"เกินมาตรฐาน"),COUNTIFS(${PC.LOG}!$A$2:$A,">="&${d1},${PC.LOG}!$A$2:$A,"<"&${d2},${PC.LOG}!$C$2:$C,${product},${PC.LOG}!$O$2:$O,"เกินมาตรฐาน"))`);
  sheet.getRange('G8').setFormula('=IFERROR(C8/(C8+E8),0)');
  sheet.getRange('A11').setFormula(`=IFERROR(${sumFormula('H').substring(1)}/${sumFormula('F').substring(1)}*10000,0)`);
  sheet.getRange('C11').setFormula(`=IFERROR(${sumFormula('I').substring(1)}/${sumFormula('F').substring(1)}*10000,0)`);
  sheet.getRange('E11').setFormula(sumFormula('H'));
  sheet.getRange('G11').setFormula(sumFormula('I'));

  ['A5','C5','E5','G5','A8','C8','E8','G8','A11','C11','E11','G11'].forEach(a =>
    sheet.getRange(a).setFontWeight('bold').setFontSize(18).setHorizontalAlignment('center')
      .setBorder(true,true,true,true,true,true,'#CBD5E1',SpreadsheetApp.BorderStyle.SOLID)
  );
  sheet.getRange('G5').setNumberFormat('0.000%');
  sheet.getRange('A8').setNumberFormat('#,##0.00');
  sheet.getRange('G8').setNumberFormat('0.0%');
  sheet.getRange('A11:C11').setNumberFormat('0.00');
  sheet.getRange('E11:G11').setNumberFormat('#,##0');

  sheet.getRange('A14:E14').setValues([['Product','Bulk ตามสูตร','ส่วนต่าง Bulk','% สูญเสีย','ผล']])
    .setBackground('#0F4C81').setFontColor('#FFFFFF').setFontWeight('bold');
  PC.PRODUCTS.forEach((p,i) => {
    const r = 15 + i;
    sheet.getRange(r,1).setValue(p[0]);
    sheet.getRange(r,2).setFormula(`=SUMIFS(${PC.LOG}!$J$2:$J,${PC.LOG}!$A$2:$A,">="&$B$2,${PC.LOG}!$A$2:$A,"<"&DATE(YEAR($B$2),MONTH($B$2)+1,1),${PC.LOG}!$C$2:$C,$A${r})`);
    sheet.getRange(r,3).setFormula(`=SUMIFS(${PC.LOG}!$L$2:$L,${PC.LOG}!$A$2:$A,">="&$B$2,${PC.LOG}!$A$2:$A,"<"&DATE(YEAR($B$2),MONTH($B$2)+1,1),${PC.LOG}!$C$2:$C,$A${r})`);
    sheet.getRange(r,4).setFormula(`=IFERROR(C${r}/B${r},"")`).setNumberFormat('0.000%');
    sheet.getRange(r,5).setFormula(`=IF(D${r}="","",IF(D${r}>=${PC.MASTER}!$C$${2+i},"ผ่าน","เกินมาตรฐาน"))`);
  });

  sheet.getRange('G14:I14').setValues([['วันที่','% สูญเสีย','ยอดผลิต']])
    .setBackground('#0F4C81').setFontColor('#FFFFFF').setFontWeight('bold');
  for (let i=0;i<31;i++) {
    const r=15+i;
    sheet.getRange(r,7).setFormula(`=IF(MONTH($B$2+${i})=MONTH($B$2),$B$2+${i},"")`).setNumberFormat('dd/MM');
    sheet.getRange(r,8).setFormula(`=IF(G${r}="","",IFERROR(SUMIFS(${PC.LOG}!$L$2:$L,${PC.LOG}!$A$2:$A,">="&G${r},${PC.LOG}!$A$2:$A,"<"&G${r}+1)/SUMIFS(${PC.LOG}!$J$2:$J,${PC.LOG}!$A$2:$A,">="&G${r},${PC.LOG}!$A$2:$A,"<"&G${r}+1),""))`).setNumberFormat('0.000%');
    sheet.getRange(r,9).setFormula(`=IF(G${r}="","",SUMIFS(${PC.LOG}!$F$2:$F,${PC.LOG}!$A$2:$A,">="&G${r},${PC.LOG}!$A$2:$A,"<"&G${r}+1))`).setNumberFormat('#,##0');
  }

  const productChart = sheet.newChart().asColumnChart()
    .addRange(sheet.getRange('A14:A24')).addRange(sheet.getRange('D14:D24'))
    .setPosition(2,11,0,0).setOption('title','% สูญเสียแยก Product')
    .setOption('legend',{position:'none'}).setOption('height',320).setOption('width',620).build();
  sheet.insertChart(productChart);

  const trendChart = sheet.newChart().asLineChart()
    .addRange(sheet.getRange('G14:H45'))
    .setPosition(20,11,0,0).setOption('title','แนวโน้ม % สูญเสียรายวัน')
    .setOption('legend',{position:'none'}).setOption('height',320).setOption('width',620).build();
  sheet.insertChart(trendChart);

  sheet.setFrozenRows(2);
}

function setupMonthly_(sheet) {
  sheet.clear();
  sheet.getRange('A1:N1').merge().setValue('สรุปรายเดือน ProdClose 1400')
    .setBackground('#0F4C81').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(18).setHorizontalAlignment('center');
  sheet.getRange('A3:N3').setValues([[
    'เดือน','จำนวนรายการ','Batch','ยอดผลิต','Bulk ตามสูตร','Bulk ตามยอดผลิตจริง',
    'ส่วนต่าง Bulk','% สูญเสีย','ผ่าน','เกินมาตรฐาน','Pass Rate','ซองเสีย','กล่องเสีย','หมายเหตุ'
  ]]).setBackground('#0F4C81').setFontColor('#FFFFFF').setFontWeight('bold');
  for (let m=1;m<=12;m++) {
    const r=3+m;
    sheet.getRange(r,1).setFormula(`=DATE(YEAR(TODAY()),${m},1)`).setNumberFormat('mmmm');
    sheet.getRange(r,2).setFormula(`=COUNTIFS(${PC.LOG}!$A$2:$A,">="&A${r},${PC.LOG}!$A$2:$A,"<"&DATE(YEAR(A${r}),MONTH(A${r})+1,1))`);
    [['C','E'],['D','F'],['E','J'],['F','K'],['G','L'],['L','H'],['M','I']].forEach(x =>
      sheet.getRange(`${x[0]}${r}`).setFormula(`=SUMIFS(${PC.LOG}!$${x[1]}$2:$${x[1]},${PC.LOG}!$A$2:$A,">="&A${r},${PC.LOG}!$A$2:$A,"<"&DATE(YEAR(A${r}),MONTH(A${r})+1,1))`)
    );
    sheet.getRange(r,8).setFormula(`=IFERROR(G${r}/E${r},"")`).setNumberFormat('0.000%');
    sheet.getRange(r,9).setFormula(`=COUNTIFS(${PC.LOG}!$A$2:$A,">="&A${r},${PC.LOG}!$A$2:$A,"<"&DATE(YEAR(A${r}),MONTH(A${r})+1,1),${PC.LOG}!$O$2:$O,"ผ่าน")`);
    sheet.getRange(r,10).setFormula(`=COUNTIFS(${PC.LOG}!$A$2:$A,">="&A${r},${PC.LOG}!$A$2:$A,"<"&DATE(YEAR(A${r}),MONTH(A${r})+1,1),${PC.LOG}!$O$2:$O,"เกินมาตรฐาน")`);
    sheet.getRange(r,11).setFormula(`=IFERROR(I${r}/(I${r}+J${r}),0)`).setNumberFormat('0.0%');
  }
  sheet.setFrozenRows(3);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActive();
    setupSystemIfMissing_(ss);

    const payload = parsePayload_(e);
    const validated = validatePayload_(payload);
    if (!validated.ok) return jsonOutput_({ok:false, error:validated.error});

    const log = ss.getSheetByName(PC.LOG);
    const recordId = String(payload.recordId || createRecordId_());

    if (isDuplicate_(log, recordId)) {
      return jsonOutput_({ok:true, duplicate:true, recordId:recordId});
    }

    const product = String(payload.product).trim();
    const std = getStandard_(product);
    const batchCount = toNumber_(payload.batchCount);
    const qty = toNumber_(payload.qty);
    let avgWeight = toNumber_(payload.avgWeight);
    if (avgWeight > 10) avgWeight = avgWeight / 1000;
    const pouchScrap = Math.max(0, toNumber_(payload.pouchScrap));
    const boxScrap = Math.max(0, toNumber_(payload.boxScrap));

    const bulkPlan = batchCount * PC.BULK_PER_BATCH;
    const bulkActual = qty * avgWeight;
    const diff = bulkActual - bulkPlan;
    const lossPct = bulkPlan ? diff / bulkPlan : 0;
    const result = lossPct >= std ? 'ผ่าน' : 'เกินมาตรฐาน';
    const pouchRate = qty ? pouchScrap / qty * 10000 : 0;
    const boxRate = qty ? boxScrap / qty * 10000 : 0;

    const row = [
      new Date(), recordId, product, String(payload.batch || '').trim(),
      batchCount, qty, avgWeight, pouchScrap, boxScrap,
      bulkPlan, bulkActual, diff, lossPct, std, result,
      pouchRate, boxRate, String(payload.remark || '').trim(),
      String(payload.reporter || '').trim(), 'GitHub Pages'
    ];
    log.appendRow(row);
    const lastRow = log.getLastRow();
    log.getRange(lastRow,1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    log.getRange(lastRow,5).setNumberFormat('0.0');
    log.getRange(lastRow,6).setNumberFormat('#,##0');
    log.getRange(lastRow,7).setNumberFormat('0.000');
    log.getRange(lastRow,8,1,2).setNumberFormat('#,##0');
    log.getRange(lastRow,10,1,3).setNumberFormat('#,##0.00');
    log.getRange(lastRow,13,1,2).setNumberFormat('0.000%');
    log.getRange(lastRow,16,1,2).setNumberFormat('0.00');

    return jsonOutput_({
      ok:true, recordId:recordId, result:result, lossPct:lossPct,
      bulkDiff:diff, timestamp:new Date().toISOString()
    });
  } catch (err) {
    return jsonOutput_({ok:false, error:String(err && err.message ? err.message : err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || 'health');
  const callback = String((e && e.parameter && e.parameter.callback) || '');

  let data;
  if (action === 'dashboard') {
    const ss = SpreadsheetApp.getActive();
    setupSystemIfMissing_(ss);
    data = getDashboardPayload_(
      String(e.parameter.month || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM')),
      String(e.parameter.product || 'ทั้งหมด')
    );
  } else {
    data = {ok:true, app:'ProdClose 1400', time:new Date().toISOString()};
  }

  if (callback && /^[A-Za-z_$][0-9A-Za-z_$\\.]*$/.test(callback)) {
    return ContentService.createTextOutput(`${callback}(${JSON.stringify(data)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOutput_(data);
}

function getDashboardPayload_(monthText, selectedProduct) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(PC.LOG);
  const values = sheet.getLastRow() > 1
    ? sheet.getRange(2,1,sheet.getLastRow()-1,PC.HEADERS.length).getValues()
    : [];

  const parts = monthText.split('-').map(Number);
  const start = new Date(parts[0], (parts[1] || 1)-1, 1);
  const end = new Date(start.getFullYear(), start.getMonth()+1, 1);
  const filtered = values.filter(r => {
    const ts = r[0] instanceof Date ? r[0] : new Date(r[0]);
    const productOk = !selectedProduct || selectedProduct === 'ทั้งหมด' || r[2] === selectedProduct;
    return ts >= start && ts < end && productOk;
  });

  const summary = {
    entries:0,batches:0,qty:0,bulkPlan:0,bulkActual:0,bulkDiff:0,lossPct:0,
    pass:0,fail:0,passRate:0,pouchScrap:0,boxScrap:0,pouchRate:0,boxRate:0
  };
  const byProduct = {};
  const daily = {};

  filtered.forEach(r => {
    summary.entries++;
    summary.batches += Number(r[4]) || 0;
    summary.qty += Number(r[5]) || 0;
    summary.bulkPlan += Number(r[9]) || 0;
    summary.bulkActual += Number(r[10]) || 0;
    summary.bulkDiff += Number(r[11]) || 0;
    summary.pouchScrap += Number(r[7]) || 0;
    summary.boxScrap += Number(r[8]) || 0;
    if (r[14] === 'ผ่าน') summary.pass++; else if (r[14] === 'เกินมาตรฐาน') summary.fail++;

    const p = String(r[2] || 'ไม่ระบุ');
    if (!byProduct[p]) byProduct[p] = {product:p,entries:0,qty:0,bulkPlan:0,bulkDiff:0,lossPct:0,pass:0,fail:0};
    const bp = byProduct[p];
    bp.entries++; bp.qty += Number(r[5]) || 0; bp.bulkPlan += Number(r[9]) || 0; bp.bulkDiff += Number(r[11]) || 0;
    if (r[14] === 'ผ่าน') bp.pass++; else if (r[14] === 'เกินมาตรฐาน') bp.fail++;

    const ts = r[0] instanceof Date ? r[0] : new Date(r[0]);
    const key = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!daily[key]) daily[key] = {date:key,qty:0,bulkPlan:0,bulkDiff:0,lossPct:0};
    daily[key].qty += Number(r[5]) || 0;
    daily[key].bulkPlan += Number(r[9]) || 0;
    daily[key].bulkDiff += Number(r[11]) || 0;
  });

  summary.lossPct = summary.bulkPlan ? summary.bulkDiff / summary.bulkPlan : 0;
  summary.passRate = (summary.pass + summary.fail) ? summary.pass / (summary.pass + summary.fail) : 0;
  summary.pouchRate = summary.qty ? summary.pouchScrap / summary.qty * 10000 : 0;
  summary.boxRate = summary.qty ? summary.boxScrap / summary.qty * 10000 : 0;

  Object.keys(byProduct).forEach(k => {
    byProduct[k].lossPct = byProduct[k].bulkPlan ? byProduct[k].bulkDiff / byProduct[k].bulkPlan : 0;
  });
  Object.keys(daily).forEach(k => {
    daily[k].lossPct = daily[k].bulkPlan ? daily[k].bulkDiff / daily[k].bulkPlan : 0;
  });

  const recent = filtered.slice(-20).reverse().map(r => ({
    timestamp: Utilities.formatDate(r[0] instanceof Date ? r[0] : new Date(r[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    recordId:r[1], product:r[2], batch:r[3], qty:r[5], lossPct:r[12],
    result:r[14], pouchScrap:r[7], boxScrap:r[8], reporter:r[18]
  }));

  return {
    ok:true, month:monthText, product:selectedProduct,
    summary:summary,
    byProduct:Object.keys(byProduct).map(k => byProduct[k]).sort((a,b) => a.lossPct - b.lossPct),
    daily:Object.keys(daily).sort().map(k => daily[k]),
    recent:recent,
    products:PC.PRODUCTS.map(p => p[0])
  };
}

function setupSystemIfMissing_(ss) {
  if (!ss.getSheetByName(PC.LOG) || !ss.getSheetByName(PC.MASTER)) setupSystem_(ss);
}

function parsePayload_(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (_) {}
  }
  return (e && e.parameter) ? e.parameter : {};
}

function validatePayload_(p) {
  if (!p.product) return {ok:false,error:'ไม่พบ Product'};
  if (!p.batch) return {ok:false,error:'ไม่พบ Batch'};
  if (!(toNumber_(p.batchCount) > 0)) return {ok:false,error:'จำนวน Batch ไม่ถูกต้อง'};
  if (!(toNumber_(p.qty) > 0)) return {ok:false,error:'ยอดผลิตไม่ถูกต้อง'};
  if (!(toNumber_(p.avgWeight) > 0)) return {ok:false,error:'น้ำหนักเฉลี่ยไม่ถูกต้อง'};
  if (!String(p.reporter || '').trim()) return {ok:false,error:'ไม่พบผู้รายงาน'};
  return {ok:true};
}

function getStandard_(product) {
  const row = PC.PRODUCTS.find(p => p[0] === product || p[1] === product);
  if (!row) throw new Error(`ไม่พบ Standard ของ Product: ${product}`);
  return Number(row[2]);
}

function isDuplicate_(sheet, recordId) {
  if (!recordId || sheet.getLastRow() < 2) return false;
  return !!sheet.getRange(2,2,sheet.getLastRow()-1,1)
    .createTextFinder(recordId).matchEntireCell(true).findNext();
}

function createRecordId_() {
  return 'PC1400-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random()*9000+1000);
}

function toNumber_(v) {
  const n = Number(String(v === undefined || v === null ? '' : v).replace(/,/g,''));
  return isFinite(n) ? n : 0;
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
