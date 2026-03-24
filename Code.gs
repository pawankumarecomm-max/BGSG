// ============================================================
// BGSG LOGISTICS — FTL Manager
// Sheet ID: 1iXlQefMxvozDAt5eAjVkiDtpd9qgRJqlwpJzhI3Ba-I
// ============================================================
// SETUP STEPS:
// 1. Open: https://docs.google.com/spreadsheets/d/1iXlQefMxvozDAt5eAjVkiDtpd9qgRJqlwpJzhI3Ba-I
// 2. Click Extensions > Apps Script
// 3. Delete everything in Code.gs, paste this file
// 4. Click + (Files) > HTML > name it exactly: Index
// 5. Paste Index.html content there
// 6. Save both files (Ctrl+S)
// 7. Select "testConnection" > Run > check Execution Log
// 8. Select "seedData" > Run (loads 61 trips)
// 9. Deploy > New Deployment > Web App > Execute as Me > Anyone > Deploy
// ============================================================

var SHEET_ID       = '1iXlQefMxvozDAt5eAjVkiDtpd9qgRJqlwpJzhI3Ba-I';
var TRIPS_SHEET    = 'Trips';
var EXPENSES_SHEET = 'Expenses';
var SALARIES_SHEET = 'Salaries';

// Serve the HTML app
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('BGSG Logistics')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ── TEST: Run this first to verify connection ──────────────
function testConnection() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('SUCCESS - Connected to: ' + ss.getName());
    Logger.log('URL: ' + ss.getUrl());
  } catch(e) {
    Logger.log('FAILED: ' + e.message);
  }
}

// ── Get sheet by name, create if missing ───────────────────
function getSheet(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    var headers, bg;
    if (name === TRIPS_SHEET) {
      headers = ['id','date','transporter','from','to','vehicle','type','invoice','customer','driverno','amount','cng','mcd','other','infreight','pl','status','pod'];
      bg = '#0f2137';
    } else if (name === EXPENSES_SHEET) {
      headers = ['id','tripId','date','type','amount','note'];
      bg = '#1a5fa8';
    } else {
      headers = ['id','month','driver','amount','note'];
      bg = '#6d28d9';
    }
    sh.appendRow(headers);
    sh.getRange(1,1,1,headers.length)
      .setBackground(bg)
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── Convert sheet rows to JS objects ──────────────────────
function sheetToObjects(sh) {
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var hdrs = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    hdrs.forEach(function(h, i) {
      var v = row[i];
      obj[h] = (v === null || v === undefined || v === '') ? '' : v;
    });
    ['amount','cng','mcd','other','infreight','pl'].forEach(function(k) {
      if (obj[k] !== undefined) obj[k] = parseFloat(obj[k]) || 0;
    });
    if (!obj.status) obj.status = 'Delivered';
    if (!obj.pod)    obj.pod    = 'Pending';
    return obj;
  });
}

// ── Load all data (called when app opens) ─────────────────
function loadAllData() {
  return {
    trips:    sheetToObjects(getSheet(TRIPS_SHEET)),
    expenses: sheetToObjects(getSheet(EXPENSES_SHEET)),
    salaries: sheetToObjects(getSheet(SALARIES_SHEET))
  };
}

// ── Save or update a trip ─────────────────────────────────
function saveTrip(trip) {
  var sh   = getSheet(TRIPS_SHEET);
  var data = sh.getDataRange().getValues();
  var hdrs = data[0];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(trip.id)) {
      sh.getRange(i+1, 1, 1, hdrs.length)
        .setValues([hdrs.map(function(h){ return trip[h] !== undefined ? trip[h] : ''; })]);
      return { ok: true, action: 'updated' };
    }
  }
  if (!trip.id) trip.id = 'r' + new Date().getTime();
  sh.appendRow(hdrs.map(function(h){ return trip[h] !== undefined ? trip[h] : ''; }));
  return { ok: true, action: 'created', id: trip.id };
}

// ── Delete trip and linked expenses ───────────────────────
function deleteTrip(tripId) {
  var sh   = getSheet(TRIPS_SHEET);
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(tripId)) { sh.deleteRow(i + 1); break; }
  }
  var esh   = getSheet(EXPENSES_SHEET);
  var edata = esh.getDataRange().getValues();
  for (var j = edata.length - 1; j >= 1; j--) {
    if (String(edata[j][1]) === String(tripId)) esh.deleteRow(j + 1);
  }
  return { ok: true };
}

// ── Save or update an expense ─────────────────────────────
function saveExpense(exp) {
  var sh   = getSheet(EXPENSES_SHEET);
  var data = sh.getDataRange().getValues();
  var hdrs = data[0];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(exp.id)) {
      sh.getRange(i+1, 1, 1, hdrs.length)
        .setValues([hdrs.map(function(h){ return exp[h] !== undefined ? exp[h] : ''; })]);
      recalcTripPL(exp.tripId);
      return { ok: true, action: 'updated', id: exp.id };
    }
  }
  if (!exp.id) exp.id = 'exp' + new Date().getTime();
  sh.appendRow(hdrs.map(function(h){ return exp[h] !== undefined ? exp[h] : ''; }));
  recalcTripPL(exp.tripId);
  return { ok: true, action: 'created', id: exp.id };
}

// ── Delete an expense ─────────────────────────────────────
function deleteExpense(expId) {
  var sh   = getSheet(EXPENSES_SHEET);
  var data = sh.getDataRange().getValues();
  var tid  = null;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(expId)) {
      tid = String(data[i][1]);
      sh.deleteRow(i + 1);
      break;
    }
  }
  if (tid) recalcTripPL(tid);
  return { ok: true };
}

// ── Recalculate trip P&L after expense change ─────────────
function recalcTripPL(tripId) {
  var tsh   = getSheet(TRIPS_SHEET);
  var tdata = tsh.getDataRange().getValues();
  var thdrs = tdata[0];
  var esh   = getSheet(EXPENSES_SHEET);
  var edata = esh.getDataRange().getValues();
  var xExp  = 0, xInF = 0;
  for (var e = 1; e < edata.length; e++) {
    if (String(edata[e][1]) === String(tripId)) {
      var a = parseFloat(edata[e][4]) || 0;
      if (edata[e][3] === 'InFreight') xInF += a; else xExp += a;
    }
  }
  var col = function(k) { return thdrs.indexOf(k); };
  for (var t = 1; t < tdata.length; t++) {
    if (String(tdata[t][0]) === String(tripId)) {
      var pl = (parseFloat(tdata[t][col('amount')]) || 0)
             - (parseFloat(tdata[t][col('cng')])    || 0)
             - (parseFloat(tdata[t][col('mcd')])    || 0)
             - (parseFloat(tdata[t][col('other')])  || 0)
             - xExp
             + (parseFloat(tdata[t][col('infreight')]) || 0)
             + xInF;
      tsh.getRange(t + 1, col('pl') + 1).setValue(pl);
      return pl;
    }
  }
}

// ── Save / delete salary ──────────────────────────────────
function saveSalary(sal) {
  var sh   = getSheet(SALARIES_SHEET);
  var data = sh.getDataRange().getValues();
  var hdrs = data[0];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(sal.id)) {
      sh.getRange(i+1, 1, 1, hdrs.length)
        .setValues([hdrs.map(function(h){ return sal[h] !== undefined ? sal[h] : ''; })]);
      return { ok: true, action: 'updated', id: sal.id };
    }
  }
  if (!sal.id) sal.id = 'sal' + new Date().getTime();
  sh.appendRow(hdrs.map(function(h){ return sal[h] !== undefined ? sal[h] : ''; }));
  return { ok: true, action: 'created', id: sal.id };
}

function deleteSalary(salId) {
  var sh   = getSheet(SALARIES_SHEET);
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(salId)) { sh.deleteRow(i + 1); break; }
  }
  return { ok: true };
}

// ── Seed 61 real trips — Run ONCE ─────────────────────────
function seedData() {
  var sh = getSheet(TRIPS_SHEET);
  if (sh.getLastRow() > 1) {
    Logger.log('Sheet already has data. Delete rows 2 onwards to re-seed.');
    return;
  }
  var rows = [
    ['r1','2026-03-02','Amit Transport','BGSG BLR','Tumkur','KA51AL2086','17FT','BLR/25-26/0364','SPROUTLIFE FOODS PRIVATE LIMITED','',7500,7500,0,0,0,0,'Delivered','Pending'],
    ['r2','2026-03-02','RCB Transport','BGSG BLR','Hyderabad, Vijaywada','KA42B0152','20FT','HR/25-26/2621, BLR/25-26/0365','CONNEDIT BUSINESS, MANASH E-COMMERCE','',36000,36000,0,0,0,0,'Delivered','Pending'],
    ['r3','2026-03-02','BGSG Logistics','Jhajjar','Faridabad','DL01LY8139','14FT','HR/25-26/2615','SHOPPERS STOP LIMITED','',6500,0,0,1150,0,5350,'Delivered','Pending'],
    ['r4','2026-03-02','BGSG Logistics','Jhajjar','Howrah','DL01MB4318','22FT','HR/25-26/2622','MANASH E-COMMERCE PRIVATE LIMITED','',58000,44334,10500,0,22500,25666,'Delivered','Pending'],
    ['r5','2026-03-02','BGSG Logistics','Jhajjar','Gurgaon','HR63G9386','22FT','HR/25-26/2609,10,11,12,13,14','VISAGE LINES PERSONAL CARE PRIVATE LIMITED','',4500,0,1500,0,0,3000,'Delivered','Pending'],
    ['r6','2026-03-02','BGSG Logistics','Jhajjar','Gautam Buddha Nagar','HR63G9386','22FT','HR/25-26/2623','DAIWA KASEI INDIA PVT LTD','',7500,5999,0,400,0,1101,'Delivered','Pending'],
    ['r7','2026-03-03','Amit Transport','BGSG BLR','Bangalore','KA51AL2085','17FT','HR/25-26/2627','VISAGE LINES PERSONAL CARE PRIVATE LIMITED','',5500,5500,0,0,0,0,'Delivered','Pending'],
    ['r8','2026-03-03','Amit Transport','BGSG BLR','Bangalore','KA51AL2085','17FT','BLR/25-26/0369','MANASH E-COMMERCE PRIVATE LIMITED','',4500,4500,0,0,0,0,'Delivered','Pending'],
    ['r9','2026-03-04','Ravi Tata Ace','Jhajjar','Sonipat','PICKUP','Bolero','HR/25-26/2631','SRT10 ATHLEISURE PRIVATE LIMITED','',3500,3500,0,0,0,0,'Delivered','Pending'],
    ['r10','2026-03-04','RCB Transport','BGSG BLR','Tumkur','KA51AG9924','20FT','BLR/25-26/0371, 73','SPROUTLIFE FOODS PRIVATE LIMITED','',7500,7500,0,0,0,0,'Delivered','Pending'],
    ['r11','2026-03-05','Ravi Tata Ace','Jhajjar','Sonipat','HR63G3137','Bolero','HR/25-26/2632','SRT10 ATHLEISURE PRIVATE LIMITED','',3000,3000,0,0,0,0,'Delivered','Pending'],
    ['r12','2026-03-05','Amit Transport','BGSG BLR','Bangalore Rural','KA51AL2085','17FT','BLR/25-26/0374','Shiprocket Limited','',6000,6000,0,0,0,0,'Delivered','Pending'],
    ['r13','2026-03-06','RCB Transport','BGSG BLR','Tumkur','KA51AG3753','20FT','BLR/25-26/0375,76,77','SPROUTLIFE FOODS PRIVATE LIMITED','',7500,7500,0,0,0,0,'Delivered','Pending'],
    ['r14','2026-03-06','BGSG Logistics','Jhajjar','Gautam Buddha Nagar','HR63G9386','22FT','HR/25-26/2634','DAIWA KASEI INDIA PVT LTD','',7500,5000,0,400,0,2100,'Delivered','Pending'],
    ['r15','2026-03-07','BGSG Logistics','Jhajjar','Gurgaon','DL01LY8139','14FT','HR/25-26/2639','BLUPIN TECHNOLOGIES PRIVATE LIMITED','',4500,2489,0,0,0,2011,'Delivered','Pending'],
    ['r16','2026-03-07','Amit Transport','BGSG BLR','Bangalore','KA51AL2086','17FT','BLR/25-26/0379','MANASH E-COMMERCE PRIVATE LIMITED','',6500,6500,0,0,0,0,'Delivered','Pending'],
    ['r17','2026-03-07','Amit Transport','BGSG BLR','Bangalore','KA51AL2086','17FT','HR/25-26/2640','VISAGE LINES PERSONAL CARE PRIVATE LIMITED','',5500,5500,0,0,0,0,'Delivered','Pending'],
    ['r18','2026-03-08','BGSG Logistics','Jhajjar','Sonipat','DL01LY8139','14FT','HR/25-26/2646','SRT10 ATHLEISURE PRIVATE LIMITED','',5500,2504,0,0,0,2996,'Delivered','Pending'],
    ['r19','2026-03-08','RCB Transport','Amar Jyothi','Ernakulam','KA51AM0588','20FT','BLR/25-26/0380','MANASH E-COMMERCE PRIVATE LIMITED','',25000,25000,0,0,0,0,'Delivered','Pending'],
    ['r20','2026-03-08','BGSG Logistics','Jhajjar','Allahabad','HR63G9386','22FT','HR/25-26/2647','AROGYAWARDHAK AUSHADHALAY','',38000,19998,2000,0,0,16002,'Delivered','Pending'],
    ['r21','2026-03-09','BGSG Logistics','Jhajjar','Gurgaon','DL01LY8139','14FT','HR/25-26/2651,52','BLUPIN TECHNOLOGIES, Shiprocket Limited','',4500,1806,0,0,0,2694,'Delivered','Pending'],
    ['r22','2026-03-09','Amit Transport','BGSG BLR','Bangalore Rural','KA51AL2085','17FT','BLR/25-26/0381','HAPPILO INTERNATIONAL PRIVATE LIMITED','',4500,4500,0,0,0,0,'Delivered','Pending'],
    ['r23','2026-03-10','BGSG Logistics','Jhajjar','Karnal','DL01LY8139','14FT','HR/25-26/2661,62','SUVIDHA STORES PRIVATE LIMITED','',6500,4505,1003,0,0,992,'Delivered','Pending'],
    ['r24','2026-03-10','JAGESWAR','BGSG BLR','Bangalore Rural','KA52C5063','14FT','BLR/25-26/0382','HAPPILO INTERNATIONAL PRIVATE LIMITED','',4500,4500,0,0,0,0,'Delivered','Pending'],
    ['r25','2026-03-10','Amit Transport','BGSG BLR','Tumkur','KA51AL2086','17FT','BLR/25-26/0383','SPROUTLIFE FOODS PRIVATE LIMITED','',7500,7500,0,0,0,0,'Delivered','Pending'],
    ['r26','2026-03-10','Shree Ganesh Road Carrier','Jhajjar','Gautam Buddha Nagar','HR55AN7682','14FT','HR/25-26/2664','DAIWA KASEI INDIA PVT LTD','',7500,7500,0,0,0,0,'Delivered','Pending'],
    ['r27','2026-03-11','RCB Transport','BGSG BLR','Tumkur','KA51AG3753','20FT','BLR/25-26/0385','SPROUTLIFE FOODS PRIVATE LIMITED','',7500,7500,0,0,0,0,'Delivered','Pending'],
    ['r28','2026-03-11','Amit Transport','BGSG BLR','Bangalore','KA51AL2085','17FT','BLR/25-26/0386','MANASH E-COMMERCE PRIVATE LIMITED','',6500,6500,0,0,0,0,'Delivered','Pending'],
    ['r29','2026-03-11','RCB Transport','BGSG BLR','Bhiwandi, Thane','KA52B7411','22FT','HR/25-26/2668,69','GLOBAL SS BEAUTY, VISAGE','',5500,5500,0,0,0,0,'Delivered','Pending'],
    ['r30','2026-03-12','BGSG Logistics','Jhajjar','Sonipat','DL01LAL0970','17FT','HR/25-26/2673','SRT10 ATHLEISURE PRIVATE LIMITED','',5500,2510,0,0,0,2990,'Delivered','Pending'],
    ['r31','2026-03-13','BGSG Logistics','Jhajjar','Gurgaon','DL01LAL0970','17FT','HR/25-26/2687,88','SANGRILA LIFESTYLE PRIVATE LIMITED','',3500,2007,1000,0,0,493,'Delivered','Pending'],
    ['r32','2026-03-12','BGSG Logistics','Jhajjar','Gurgaon','DL01LY8139','14FT','HR/25-26/2689,90','VISAGE LINES PERSONAL CARE PRIVATE LIMITED','',4500,2006,0,1200,0,1294,'Delivered','Pending'],
    ['r33','2026-03-12','BGSG Logistics','Jhajjar','Thane','HR63G9386','22FT','HR/25-26/2674,76,77,78,79,80,81','SHOPPERS STOP, Vera Moda, Best United','',37000,38948,8000,6000,40000,24052,'Delivered','Pending'],
    ['r34','2026-03-13','BGSG Logistics','Jhajjar','Gautam Buddha Nagar','DL01LAL0970','17FT','HR/25-26/2693','DAIWA KASEI INDIA PVT LTD','',7500,3219,2220,400,0,1661,'Delivered','Pending'],
    ['r35','2026-03-13','BGSG Logistics','Noida','Jhajjar','DL01LAL0970','17FT','inward freight','Accurate Multiplayer Papers LLP','',6500,0,0,0,0,6500,'Delivered','Pending'],
    ['r36','2026-03-13','Routeredar-Divya','BGSG BLR','Coimbatore','KA02AH9639','20FT','BLR/25-26/0389,90','MANASH E-COMMERCE PRIVATE LIMITED','',18000,18000,0,0,0,0,'Delivered','Pending'],
    ['r37','2026-03-13','Khan Transport','BGSG BLR','Tumkur','KA01AA2213','20FT','BLR/25-26/0392','SPROUTLIFE FOODS PRIVATE LIMITED','',7500,7500,0,0,0,0,'Delivered','Pending'],
    ['r38','2026-03-14','BGSG Logistics','Jhajjar','Gurgaon','DL01LY8139','14FT','HR/25-26/2694','VISAGE LINES PERSONAL CARE PRIVATE LIMITED','',4500,3006,55,0,0,1439,'Delivered','Pending'],
    ['r39','2026-03-14','RCB Transport','BGSG BLR','Tumkur','KA52B6241','20FT','BLR/25-26/0394','SPROUTLIFE FOODS PRIVATE LIMITED','',7500,7500,0,0,0,0,'Delivered','Pending'],
    ['r40','2026-03-14','Surindera Transport','Jhajjar','Khordha','HR55AW9978','22FT','HR/25-26/2699,700,701','MANASH E-COMMERCE PRIVATE LIMITED','',68000,68000,0,0,0,0,'Delivered','Pending'],
    ['r41','2026-03-14','RCB Transport','BGSG BLR','Tumkur','KA52B6241','20FT','BLR/25-26/0394b','SPROUTLIFE FOODS PRIVATE LIMITED','',7500,7500,0,0,0,0,'Delivered','Pending'],
    ['r42','2026-03-14','Khan Transport','BGSG BLR','Tumkur','KA01AA2213','20FT','BLR/25-26/0396','SPROUTLIFE FOODS PRIVATE LIMITED','',7500,7500,0,0,0,0,'Delivered','Pending'],
    ['r43','2026-03-14','Routeredar-Divya','BGSG BLR','Bengaluru','KA52C3807','14FT','BLR/25-26/0397','MANASH E-COMMERCE PRIVATE LIMITED','',5000,5000,0,0,0,0,'Delivered','Pending'],
    ['r44','2026-03-14','OM LOGISTICS LTD.','Jhajjar','KOLKATA','HR63F9444','14FT','HR/25-26/2696','CONNEDIT BUSINESS SOLUTIONS PRIVATE LIMITED','',0,0,0,0,0,0,'Delivered','Pending'],
    ['r45','2026-03-14','BGSG Logistics','EPW','Farrukhnagar','DL01LY8139','14FT','EPW invoice','Shiprocket Limited','',5500,1000,0,0,0,4500,'Delivered','Pending'],
    ['r46','2026-03-16','BGSG Logistics','Jhajjar','Gurgaon','DL01LY8139','14FT','HR/25-26/2702','CONNEDIT BUSINESS SOLUTIONS PRIVATE LIMITED','',3500,1002,1750,0,0,748,'Delivered','Pending'],
    ['r47','2026-03-16','BGSG Logistics','Jhajjar','Thane','DL01MB4318','22FT','HR/25-26/2706,07,08','FRUVEGGIE TECHNOLOGY, Vera Moda, Best United','',40000,14886,3003,2000,0,20111,'Delivered','Pending'],
    ['r48','2026-03-16','Khan Transport','BGSG BLR','Bangalore','14FT','14FT','HR/25-26/2709,10','SHOPPERS STOP LIMITED','',4000,4000,0,0,0,0,'Delivered','Pending'],
    ['r49','2026-03-17','RCB Transport','BGSG BLR','Vijayawada','KA51AG1259','20FT','BLR/25-26/0401,102,103','MANASH E-COMMERCE PRIVATE LIMITED','',26000,26000,0,0,0,0,'Delivered','Pending'],
    ['r50','2026-03-17','BGSG Logistics','Jhajjar','Dadri Toye','DL01LAL0970','17FT','internal movement','Arihant (for REELS)','',1500,1116,0,0,0,384,'Delivered','Pending'],
    ['r51','2026-03-17','BGSG Logistics','Jhajjar','Gautam Buddha Nagar','DL01LAL0970','17FT','HR/25-26/2720','DAIWA KASEI INDIA PVT LTD','',7500,3208,0,400,0,3892,'Delivered','Pending'],
    ['r52','2026-03-17','BGSG Logistics','Noida','Jhajjar','DL01LAL0970','17FT','inward freight 2','Accurate Multiplayer Papers LLP','',6500,1001,0,0,0,5499,'Delivered','Pending'],
    ['r53','2026-03-18','Shree Ganesh Road Carrier','Jhajjar','Gurgaon','HR55AH9928','14FT','HR/25-26/2728','Shiprocket Limited','',5500,5500,0,0,0,0,'Delivered','Pending'],
    ['r54','2026-03-18','Amit Transport','BGSG BLR','Bangalore','KA51AL2085','17FT','HR/25-26/2729','VISAGE LINES PERSONAL CARE PRIVATE LIMITED','',5500,5500,0,0,0,0,'Delivered','Pending'],
    ['r55','2026-03-18','Chinar Packaging','BGSG BLR','Tumkur','KA01D4950','14FT','BLR/25-26/0404','SPROUTLIFE FOODS PRIVATE LIMITED','',7500,7500,0,0,0,0,'Delivered','Pending'],
    ['r56','2026-03-19','BGSG Logistics','Jhajjar','Gurgaon','DL01LY8139','14FT','HR/25-26/2730,31','SANGRILA LIFESTYLE PRIVATE LIMITED','',3500,1995,0,0,0,1505,'Delivered','Pending'],
    ['r57','2026-03-19','Wheelz EYE','BGSG BLR','Ernakulam','KA01AM6439','20FT','BLR/25-26/0405,06','MANASH E-COMMERCE, MANASH LIFESTYLE PVT LTD','',25000,25000,0,0,0,0,'Delivered','Pending'],
    ['r58','2026-03-20','BGSG Logistics','Jhajjar','Patna','DL01LAL0970','17FT','HR/25-26/2741,42','MANASH E-COMMERCE PRIVATE LIMITED','',45000,7599,2000,1000,0,34401,'Delivered','Pending'],
    ['r59','2026-03-20','BGSG Logistics','Jhajjar','Gautam Buddha Nagar','DL01LY8139','14FT','HR/25-26/2745','DAIWA KASEI INDIA PVT LTD','',7500,3000,0,400,0,4100,'Delivered','Pending'],
    ['r60','2026-03-20','OM LOGISTICS LTD.','Jhajjar','Thane','HR63F9444','Bolero','HR/25-26/2746,47,50,56','Best United India, Freedom Tree Retail, OVS INDIA RETAIL','',0,0,0,0,0,0,'Delivered','Pending'],
    ['r61','2026-03-21','DS INDIA LOGISTICS','Jhajjar','Karnataka','MH46CU1680','32FT','HR/25-26/2752','BGSG SOLUTIONS PRIVATE LIMITED','',99000,99000,0,0,0,0,'Delivered','Pending']
  ];
  rows.forEach(function(r) { sh.appendRow(r); });
  Logger.log('Done! 61 trips loaded. All status = Delivered.');
}
