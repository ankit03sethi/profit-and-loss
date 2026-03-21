// ============================================================
// MASTER PAYMENTS FILE - Apps Script (Sheets API - FAST)
// ============================================================
// Replaces slow QUERY+IMPORTRANGE with Sheets API batchGet
// Pulls 12 columns from 16 source tabs across 4 spreadsheets
// Headers: TOTAL, S.NO, YEAR, MONTH, PLATFORM, COMPANY, TYPE, TYPE 2, ORDER ID, SKU, QTY, AMOUNT
//
// SETUP:
// 1. Paste in Extensions > Apps Script > Save
// 2. Click "Services" (+) on left > Add "Google Sheets API"
// 3. Reload sheet > use "Master Payments" menu
// ============================================================

var HEADERS = ['TOTAL','S.NO','YEAR','MONTH','PLATFORM','COMPANY','TYPE','TYPE 2','ORDER ID','SKU','QTY','AMOUNT'];
var MASTER_TAB = 'MASTER PAYMENTS FILE';

// VLOOKUP source: Master SKU FILE
var SKU_SHEET_ID = '1iSNEwTlqDWCBC0zYGLUUIJMl_F_mmtjlOwh4AlBjbXE';
var SKU_TAB = 'Master SKU FILE';
// Key = column AA, Return = AC, AD, AG, AJ, AM, AQ → write to M, N, O, P, Q, R

// Revenue types (positive P&L contribution)
var REVENUE_TYPES = {'Order':1,'Transfer':1,'Reimbursements':1,'MP Fee Rebate':1,'Referral Payment':1,'Fee Refund':1};
// Cost types (negative P&L contribution) — everything else

var PLATFORMS = {
  orders1: {
    id: '14zSHcErsm1wWDfdU0e9i_qttzMSbu9lTy4a9jj-2c0E',
    tab: 'ORDERS',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','AK2:AK','I2:I','J2:J','AJ2:AJ','M2:M','AG2:AG']
  },
  paymentAdvice: {
    id: '1PWavO2jatx_48djKk491ahEQM-hgLrVzexNqsNe9ajA',
    tab: 'PAYMENT ADVICE',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','G2:G','J2:J','AD2:AD','Q2:Q','T2:T']
  },
  sr: {
    id: '1PWavO2jatx_48djKk491ahEQM-hgLrVzexNqsNe9ajA',
    tab: 'SR',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','H2:H','J2:J','AS2:AS','S2:S','AT2:AT']
  },
  rto: {
    id: '1PWavO2jatx_48djKk491ahEQM-hgLrVzexNqsNe9ajA',
    tab: 'RTO',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','H2:H','J2:J','AS2:AS','S2:S','AT2:AT']
  },
  orders2: {
    id: '19jHN50rRSIqEKEC6yxiOYKFjwSONKZ1ICPzj7aLA5Gs',
    tab: 'ORDERS',
    cols: ['A3:A','B3:B','C3:C','D3:D','E3:E','F3:F','CF3:CF','BQ3:BQ','N3:N','BM3:BM','BN3:BN','J3:J']
  },
  mpFeeRebate: {
    id: '19jHN50rRSIqEKEC6yxiOYKFjwSONKZ1ICPzj7aLA5Gs',
    tab: 'MP Fee Rebate',
    cols: ['A4:A','B4:B','C4:C','D4:D','E4:E','F4:F','G4:G','H4:H','O4:O','N4:N','AC4:AC','L4:L']
  },
  nonOrderSPF: {
    id: '19jHN50rRSIqEKEC6yxiOYKFjwSONKZ1ICPzj7aLA5Gs',
    tab: 'Non_Order_SPF',
    cols: ['A3:A','B3:B','C3:C','D3:D','E3:E','F3:F','G3:G','H3:H','I3:I','AB3:AB','AC3:AC','K3:K']
  },
  storageRecall: {
    id: '19jHN50rRSIqEKEC6yxiOYKFjwSONKZ1ICPzj7aLA5Gs',
    tab: 'Storage_Recall',
    cols: ['A3:A','B3:B','C3:C','D3:D','E3:E','F3:F','G3:G','H3:H','J3:J','AB3:AB','AC3:AC','L3:L']
  },
  valueAddedServices: {
    id: '19jHN50rRSIqEKEC6yxiOYKFjwSONKZ1ICPzj7aLA5Gs',
    tab: 'Value Added Services',
    cols: ['A3:A','B3:B','C3:C','D3:D','E3:E','F3:F','G3:G','H3:H','I3:I','AB3:AB','AC3:AC','K3:K']
  },
  ads1: {
    id: '19jHN50rRSIqEKEC6yxiOYKFjwSONKZ1ICPzj7aLA5Gs',
    tab: 'Ads',
    cols: ['A3:A','B3:B','C3:C','D3:D','E3:E','F3:F','G3:G','h3:h','I3:I','AC3:AC','AD3:AD','Q3:Q']
  },
  transfer1: {
    id: '19jHN50rRSIqEKEC6yxiOYKFjwSONKZ1ICPzj7aLA5Gs',
    tab: 'Transfer',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','G2:G','J2:J','U2:U','V2:V','M2:M']
  },
  orderPayments: {
    id: '1ov2Z63nO6EpvmAQmlqA-wWwCdqebj20pBjWr2qP2P8c',
    tab: 'Order Payments',
    cols: ['A4:A','B4:B','C4:C','D4:D','E4:E','F4:F','AW4:AW','L4:L','G4:G','K4:K','O4:O','R4:R']
  },
  adsCost: {
    id: '1ov2Z63nO6EpvmAQmlqA-wWwCdqebj20pBjWr2qP2P8c',
    tab: 'Ads Cost',
    cols: ['A4:A','B4:B','C4:C','D4:D','E4:E','F4:F','G4:G','G4:G','J4:J','AB4:AB','AC4:AC','O4:O']
  },
  referralPayments: {
    id: '1ov2Z63nO6EpvmAQmlqA-wWwCdqebj20pBjWr2qP2P8c',
    tab: 'Referral Payments',
    cols: ['A4:A','B4:B','C4:C','D4:D','E4:E','F4:F','G4:G','H4:H','I4:I','AB4:AB','AC4:AC','M4:M']
  },
  compensationRecovery: {
    id: '1ov2Z63nO6EpvmAQmlqA-wWwCdqebj20pBjWr2qP2P8c',
    tab: 'Compensation and Recovery',
    cols: ['A4:A','B4:B','C4:C','D4:D','E4:E','F4:F','G4:G','H4:H','AA4:AA','AB4:AB','AC4:AC','L4:L']
  },
  transfer2: {
    id: '1ov2Z63nO6EpvmAQmlqA-wWwCdqebj20pBjWr2qP2P8c',
    tab: 'Transfer',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','G2:G','M2:M','R2:R','S2:S','N2:N']
  },
  myntraPayments: {
    id: '1kFlxfE7qgEFQovKH37qO34bxCJ8gbS6IQxuq1mtUrCM',
    tab: 'Payments',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','H2:H','AI2:AI','BF2:BF','A2:A','Z2:Z']
  },
  myntraDeductions: {
    id: '1kFlxfE7qgEFQovKH37qO34bxCJ8gbS6IQxuq1mtUrCM',
    tab: 'Deductions',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','H2:H','L2:L','A2:A','A2:A','I2:I'],
    extraCol: {src: 'J2:J', destCol: 15}
  }
};

// ===================== MENU =====================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Master Payments')
    .addItem('>>> RUN EVERYTHING <<<', 'runEverything')
    .addItem('>>> CLEAR & RUN FRESH <<<', 'clearAndRun')
    .addSeparator()
    .addItem('Sync All Sources', 'syncAll')
    .addSeparator()
    .addItem('Sync Orders 1', 'sOrders1')
    .addItem('Sync Payment Advice', 'sPaymentAdvice')
    .addItem('Sync SR', 'sSR')
    .addItem('Sync RTO', 'sRTO')
    .addItem('Sync Orders 2', 'sOrders2')
    .addItem('Sync MP Fee Rebate', 'sMpFeeRebate')
    .addItem('Sync Non Order SPF', 'sNonOrderSPF')
    .addItem('Sync Storage Recall', 'sStorageRecall')
    .addItem('Sync Value Added Services', 'sValueAddedServices')
    .addItem('Sync Ads 1', 'sAds1')
    .addItem('Sync Transfer 1', 'sTransfer1')
    .addItem('Sync Order Payments', 'sOrderPayments')
    .addItem('Sync Ads Cost', 'sAdsCost')
    .addItem('Sync Referral Payments', 'sReferralPayments')
    .addItem('Sync Compensation & Recovery', 'sCompensationRecovery')
    .addItem('Sync Transfer 2', 'sTransfer2')
    .addItem('Sync Myntra Payments', 'sMyntraPayments')
    .addItem('Sync Myntra Deductions', 'sMyntraDeductions')
    .addSeparator()
    .addItem('Run VLOOKUP (SKU)', 'vlookupSKU')
    .addItem('Compute Formulas (S,T,U,V)', 'computeFormulas')
    .addSeparator()
    .addItem('Remove Duplicates', 'removeDuplicates')
    .addSeparator()
    .addItem('Start Auto-Refresh (1 min)', 'setupAutoRefresh')
    .addItem('Stop Auto-Refresh', 'stopAutoRefresh')
    .addSeparator()
    .addItem('Setup Sheet', 'setupSheet')
    .addToUi();

  ui.createMenu('P&L Dashboard')
    .addItem('Test API', 'testPLApi')
    .addToUi();
}

// Individual sync functions
function sOrders1() { syncOne('orders1'); }
function sPaymentAdvice() { syncOne('paymentAdvice'); }
function sSR() { syncOne('sr'); }
function sRTO() { syncOne('rto'); }
function sOrders2() { syncOne('orders2'); }
function sMpFeeRebate() { syncOne('mpFeeRebate'); }
function sNonOrderSPF() { syncOne('nonOrderSPF'); }
function sStorageRecall() { syncOne('storageRecall'); }
function sValueAddedServices() { syncOne('valueAddedServices'); }
function sAds1() { syncOne('ads1'); }
function sTransfer1() { syncOne('transfer1'); }
function sOrderPayments() { syncOne('orderPayments'); }
function sAdsCost() { syncOne('adsCost'); }
function sReferralPayments() { syncOne('referralPayments'); }
function sCompensationRecovery() { syncOne('compensationRecovery'); }
function sTransfer2() { syncOne('transfer2'); }
function sMyntraPayments() { syncOne('myntraPayments'); }
function sMyntraDeductions() { syncOne('myntraDeductions'); }

// ===================== VLOOKUP SKU =====================
// Reads Master SKU FILE: key=AA, returns AC,AD,AG,AJ,AM,AQ
// Writes to MASTER PAYMENTS FILE columns M,N,O,P,Q,R based on key=J

function vlookupSKU() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName(MASTER_TAB);
  if (!ms) { SpreadsheetApp.getUi().alert('Master tab not found. Run Setup Sheet first.'); return; }

  var lr = ms.getLastRow();
  if (lr < 2) { SpreadsheetApp.getUi().alert('No data in master tab.'); return; }

  // 1. Read lookup table from Master SKU FILE via Sheets API
  var lookupCols = ['AA2:AA','AC2:AC','AD2:AD','AG2:AG','AJ2:AJ','AM2:AM','AQ2:AQ'];
  var ranges = [];
  for (var i = 0; i < lookupCols.length; i++) {
    ranges.push("'" + SKU_TAB + "'!" + lookupCols[i]);
  }
  var res = Sheets.Spreadsheets.Values.batchGet(SKU_SHEET_ID, {ranges: ranges});
  var vr = res.valueRanges;

  // Find max rows in lookup
  var maxLookup = 0;
  for (var v = 0; v < vr.length; v++) {
    var len = (vr[v].values || []).length;
    if (len > maxLookup) maxLookup = len;
  }

  // 2. Build lookup map: AA value → [AC, AD, AG, AJ, AM, AQ] (case-insensitive)
  var lookupMap = {};
  for (var r = 0; r < maxLookup; r++) {
    var keyArr = (vr[0].values || []);
    var key = r < keyArr.length && keyArr[r].length > 0 ? String(keyArr[r][0]).trim().toUpperCase() : '';
    if (!key) continue;
    var vals = [];
    for (var c = 1; c < vr.length; c++) {
      var cv = (vr[c].values || []);
      vals.push(r < cv.length && (cv[r] || []).length > 0 ? cv[r][0] : '');
    }
    if (!lookupMap[key]) lookupMap[key] = vals; // first match wins
  }

  // 3. Read column J (col 10) from MASTER PAYMENTS FILE
  var keys = ms.getRange(2, 10, lr - 1, 1).getValues();

  // 4. Build output for columns M-R (cols 13-18)
  var output = [];
  var matched = 0;
  for (var i = 0; i < keys.length; i++) {
    var k = String(keys[i][0]).trim().toUpperCase();
    if (k && lookupMap[k]) {
      output.push(lookupMap[k]);
      matched++;
    } else {
      output.push(['','','','','','']);
    }
  }

  // 5. Write to columns M-R
  if (output.length > 0) {
    ms.getRange(2, 13, output.length, 6).setValues(output);
  }

  SpreadsheetApp.getUi().alert('VLOOKUP Done!\nRows checked: ' + keys.length + '\nMatched: ' + matched);
}

// ===================== COMPUTE FORMULAS (S, T, U, W) =====================
// R = AQ from VLOOKUP (untouched)
// S = SUMIF(A:A, A_value, L:L) — sum of AMOUNT for all rows with same TOTAL
// T = R value (VLOOKUP AQ) shown ONLY when: multiple entries for same A, no RETURN in G, first ORDER row only
// U = S value (SUMIF) shown with priority: if RETURN exists → first RETURN row; else first ORDER row; only when multiple entries
// V = S value on 1st qualifying entry per unique A; rows with G=Order/Return/Shipping Service → always blank, totally ignored

function computeFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName(MASTER_TAB);
  if (!ms) return;
  var lr = ms.getLastRow();
  if (lr < 2) return;

  var numRows = lr - 1;

  // Read required columns
  var dataA = ms.getRange(2, 1, numRows, 1).getValues();   // A (TOTAL)
  var dataG = ms.getRange(2, 7, numRows, 1).getValues();   // G (TYPE)
  var dataL = ms.getRange(2, 12, numRows, 1).getValues();  // L (AMOUNT)
  var dataR = ms.getRange(2, 18, numRows, 1).getValues();  // R (VLOOKUP AQ value)

  // === COLUMN S (col 19): SUMIF(A:A, A_value, L:L) ===
  // NOTE: All A values normalized to UPPERCASE for case-insensitive matching
  var sumByA = {};
  for (var i = 0; i < numRows; i++) {
    var a = String(dataA[i][0]).trim().toUpperCase();
    if (!a) continue;
    var raw = dataL[i][0];
    var amt = 0;
    if (typeof raw === 'number') {
      amt = raw;
    } else {
      var cleaned = String(raw).replace(/,/g, '').trim();
      amt = parseFloat(cleaned) || 0;
    }
    sumByA[a] = (sumByA[a] || 0) + amt;
  }
  var outS = [];
  for (var i = 0; i < numRows; i++) {
    var a = String(dataA[i][0]).trim().toUpperCase();
    outS.push([a ? (Math.round((sumByA[a] || 0) * 100) / 100) : '']);
  }

  // === Build lookup maps for T and U (case-insensitive on A) ===
  var countByA = {};        // A value → count of rows
  var hasReturn = {};       // A value → true if any row has TYPE=RETURN
  var firstOrderRow = {};   // A value → first row index where TYPE=ORDER
  var firstReturnRow = {};  // A value → first row index where TYPE=RETURN

  for (var i = 0; i < numRows; i++) {
    var a = String(dataA[i][0]).trim().toUpperCase();
    if (!a) continue;
    countByA[a] = (countByA[a] || 0) + 1;
    var g = String(dataG[i][0]).trim().toUpperCase();
    if (g === 'RETURN') {
      hasReturn[a] = true;
      if (firstReturnRow[a] === undefined) firstReturnRow[a] = i;
    }
    if (g === 'ORDER' && firstOrderRow[a] === undefined) firstOrderRow[a] = i;
  }

  // === COLUMN T (col 20): R value (VLOOKUP AQ) ===
  // RULES:
  // 1. Single ORDER, no RETURN         → T = R value on that row
  // 2. Multiple ORDERs, no RETURN       → T = R value on 1st ORDER only, rest blank
  // 3. Any RETURN exists (with/without ORDERs) → T = ALL BLANK
  var outT = [];
  for (var i = 0; i < numRows; i++) {
    var a = String(dataA[i][0]).trim().toUpperCase();
    if (!a) { outT.push(['']); continue; }
    if (hasReturn[a]) { outT.push(['']); continue; }            // ANY return → T all blank
    if (firstOrderRow[a] === i) { outT.push([dataR[i][0]]); continue; } // 1st ORDER → show R
    outT.push(['']);                                              // rest → blank
  }

  // === COLUMN U (col 21): S value (SUMIF) ===
  // RULES:
  // 1. Single ORDER, no RETURN              → U = S value on that row
  // 2. Multiple ORDERs, no RETURN            → U = S value on 1st ORDER only, rest blank
  // 3. RETURN exists (+ ORDERs)              → ORDER rows ALL BLANK, U = S on 1st RETURN only
  // 4. Single RETURN only (no ORDER)         → U = S value on that row
  // 5. Multiple RETURNs only (no ORDER)      → U = S value on 1st RETURN only, rest blank
  var outU = [];
  for (var i = 0; i < numRows; i++) {
    var a = String(dataA[i][0]).trim().toUpperCase();
    if (!a) { outU.push(['']); continue; }
    if (hasReturn[a]) {
      outU.push([firstReturnRow[a] === i ? outS[i][0] : '']);
    } else if (firstOrderRow[a] !== undefined) {
      outU.push([firstOrderRow[a] === i ? outS[i][0] : '']);
    } else {
      outU.push(['']);
    }
  }

  // === COLUMN V (col 22): S value on 1st qualifying entry per unique A ===
  // RULES:
  // 1. Rows where G = ORDER, RETURN, or SHIPPING SERVICE → V = blank (totally ignored)
  // 2. Among remaining rows, for each unique A value → V = S value on 1st entry only, rest blank
  var EXCLUDED_TYPES_V = {'ORDER':1, 'RETURN':1, 'SHIPPING SERVICE':1};
  var firstQualifyingRow = {};  // A value → first row index where G is NOT in excluded types
  var allTypesFound = {};       // DEBUG: collect all unique G values
  var qualifyingCount = 0;      // DEBUG: count qualifying rows

  for (var i = 0; i < numRows; i++) {
    var a = String(dataA[i][0]).trim().toUpperCase();
    if (!a) continue;
    var g = String(dataG[i][0]).trim().toUpperCase();
    allTypesFound[g] = (allTypesFound[g] || 0) + 1;  // DEBUG
    if (EXCLUDED_TYPES_V[g]) continue;  // totally ignore these
    qualifyingCount++;
    if (firstQualifyingRow[a] === undefined) firstQualifyingRow[a] = i;
  }

  var vFilledCount = 0;
  var outV = [];
  for (var i = 0; i < numRows; i++) {
    var a = String(dataA[i][0]).trim().toUpperCase();
    if (!a) { outV.push(['']); continue; }
    var g = String(dataG[i][0]).trim().toUpperCase();
    if (EXCLUDED_TYPES_V[g]) { outV.push(['']); continue; }  // excluded type → blank
    if (firstQualifyingRow[a] === i) {
      outV.push([outS[i][0]]);  // 1st qualifying entry → show S value
      vFilledCount++;
    } else {
      outV.push(['']);  // not 1st → blank
    }
  }

  // Write S(19), T(20), U(21), V(22)
  ms.getRange(2, 19, numRows, 1).setValues(outS);
  ms.getRange(2, 20, numRows, 1).setValues(outT);
  ms.getRange(2, 21, numRows, 1).setValues(outU);
  ms.getRange(2, 22, numRows, 1).setValues(outV);

  // DEBUG ALERT: show what types exist and how many qualifying rows
  var typeList = [];
  var typeKeys = Object.keys(allTypesFound).sort();
  for (var ti = 0; ti < typeKeys.length; ti++) {
    typeList.push(typeKeys[ti] + ': ' + allTypesFound[typeKeys[ti]]);
  }
  SpreadsheetApp.getUi().alert(
    'COLUMN V DEBUG:\n' +
    'Total rows: ' + numRows + '\n' +
    'Qualifying rows (not Order/Return/Shipping Service): ' + qualifyingCount + '\n' +
    'Unique A values with qualifying row: ' + Object.keys(firstQualifyingRow).length + '\n' +
    'V cells filled: ' + vFilledCount + '\n\n' +
    'ALL TYPE VALUES IN G:\n' + typeList.join('\n')
  );
}

// ===================== FAST READ (Sheets API batchGet) =====================

function fastRead(sheetId, tabName, colRanges) {
  var ranges = [];
  for (var i = 0; i < colRanges.length; i++) {
    ranges.push("'" + tabName + "'!" + colRanges[i]);
  }
  var res = Sheets.Spreadsheets.Values.batchGet(sheetId, {ranges: ranges});
  var vr = res.valueRanges;

  var maxRows = 0;
  for (var v = 0; v < vr.length; v++) {
    var len = (vr[v].values || []).length;
    if (len > maxRows) maxRows = len;
  }

  var rows = [];
  for (var r = 0; r < maxRows; r++) {
    var row = [];
    for (var c = 0; c < vr.length; c++) {
      var vals = vr[c].values || [];
      row.push(r < vals.length && vals[r].length > 0 ? vals[r][0] : '');
    }
    rows.push(row);
  }
  return rows;
}

// ===================== SYNC ONE SOURCE =====================

function syncOne(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName(MASTER_TAB);
  if (!ms) { SpreadsheetApp.getUi().alert('Run Setup Sheet first.'); return; }

  var p = PLATFORMS[name];

  var raw;
  try {
    raw = fastRead(p.id, p.tab, p.cols);
  } catch(e) {
    SpreadsheetApp.getUi().alert(name + ' error: ' + e.message);
    return;
  }

  // Keep ALL rows where column A (TOTAL) is not blank — NO dedup
  var newRows = [];
  for (var i = 0; i < raw.length; i++) {
    var key = String(raw[i][0]).trim();
    if (!key) continue; // Skip blank TOTAL only
    newRows.push(raw[i]);
  }

  // Fix negative Transfer L values for Amazon
  _fixTransferL(newRows);

  if (newRows.length > 0) {
    var lr = Math.max(ms.getLastRow(), 1);
    ms.getRange(lr + 1, 1, newRows.length, 12).setValues(newRows);

    // Handle extraCol (e.g., myntraDeductions writes src column to destCol)
    if (p.extraCol) {
      try {
        var extraRaw = fastRead(p.id, p.tab, [p.extraCol.src]);
        var extraOut = [];
        var rawIdx = 0;
        for (var ei = 0; ei < raw.length; ei++) {
          if (String(raw[ei][0]).trim() === '') continue;
          extraOut.push([ei < extraRaw.length ? extraRaw[ei][0] : '']);
        }
        if (extraOut.length > 0) {
          ms.getRange(lr + 1, p.extraCol.destCol, extraOut.length, 1).setValues(extraOut);
        }
      } catch(ex) {
        Logger.log('extraCol error for ' + name + ': ' + ex.message);
      }
    }
  }

  SpreadsheetApp.getUi().alert(name + ': ' + raw.length + ' pulled, ' + newRows.length + ' added');
}

// ===================== SYNC ALL =====================

function syncAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName(MASTER_TAB);
  if (!ms) { SpreadsheetApp.getUi().alert('Run Setup Sheet first.'); return; }

  var names = Object.keys(PLATFORMS);
  var totalP = 0, totalA = 0;

  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var p = PLATFORMS[name];

    try {
      var raw = fastRead(p.id, p.tab, p.cols);
      totalP += raw.length;

      // Keep ALL rows where column A is not blank — NO dedup
      var newRows = [];
      for (var r = 0; r < raw.length; r++) {
        var key = String(raw[r][0]).trim();
        if (!key) continue;
        newRows.push(raw[r]);
      }

      // Fix negative Transfer L values for Amazon
      _fixTransferL(newRows);

      if (newRows.length > 0) {
        var lr = Math.max(ms.getLastRow(), 1);
        ms.getRange(lr + 1, 1, newRows.length, 12).setValues(newRows);
        totalA += newRows.length;
      }

      Logger.log(name + ': ' + raw.length + ' pulled, ' + newRows.length + ' added');
      SpreadsheetApp.flush();
    } catch(e) {
      Logger.log('ERROR ' + name + ': ' + e.message);
    }
  }

  SpreadsheetApp.getUi().alert('All done!\nPulled: ' + totalP + '\nAdded: ' + totalA);
}

// ===================== GET EXISTING TOTALS =====================

function getExisting(ms) {
  var map = {};
  var lr = ms.getLastRow();
  if (lr < 2) return map;
  var vals = ms.getRange(2, 1, lr - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var k = String(vals[i][0]).trim();
    if (k) map[k] = true;
  }
  return map;
}

// ===================== REMOVE DUPLICATES =====================

function removeDuplicates() {
  // Dedup DISABLED — same order can have multiple entries (ORDER, RETURN) with same TOTAL value
  // The original QUERY formula keeps ALL rows (SELECT * WHERE Col1 IS NOT NULL)
  SpreadsheetApp.getUi().alert('Dedup is disabled.\nSame order can have multiple entries (ORDER + RETURN) with different amounts.\nUse "Clear & Run Fresh" to re-sync all data.');
}

// ===================== RUN EVERYTHING =====================

function runEverything() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName(MASTER_TAB);
  if (!ms) { ms = ss.insertSheet(MASTER_TAB); }

  var log = [];
  var totalP = 0, totalA = 0;

  var names = Object.keys(PLATFORMS);
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var p = PLATFORMS[name];
    try {
      var raw = fastRead(p.id, p.tab, p.cols);
      totalP += raw.length;
      // Keep ALL rows where column A is not blank — NO dedup
      var newRows = [];
      for (var r = 0; r < raw.length; r++) {
        var key = String(raw[r][0]).trim();
        if (!key) continue;
        newRows.push(raw[r]);
      }
      // Fix negative Transfer L values for Amazon
      _fixTransferL(newRows);
      if (newRows.length > 0) {
        var lr = Math.max(ms.getLastRow(), 1);
        ms.getRange(lr + 1, 1, newRows.length, 12).setValues(newRows);
        totalA += newRows.length;
      }
      log.push(name + ': +' + newRows.length);
      SpreadsheetApp.flush();
    } catch(e) {
      log.push(name + ' ERR: ' + e.message);
    }
  }

  // NO dedup — same order can have ORDER + RETURN entries with same TOTAL value

  // VLOOKUP: fill columns M-R from Master SKU FILE
  try {
    var lrV = ms.getLastRow();
    if (lrV >= 2) {
      var lookupCols = ['AA2:AA','AC2:AC','AD2:AD','AG2:AG','AJ2:AJ','AM2:AM','AQ2:AQ'];
      var lRanges = [];
      for (var li = 0; li < lookupCols.length; li++) {
        lRanges.push("'" + SKU_TAB + "'!" + lookupCols[li]);
      }
      var lRes = Sheets.Spreadsheets.Values.batchGet(SKU_SHEET_ID, {ranges: lRanges});
      var lVr = lRes.valueRanges;
      var maxL = 0;
      for (var lv = 0; lv < lVr.length; lv++) { var ll = (lVr[lv].values || []).length; if (ll > maxL) maxL = ll; }

      var lMap = {};
      for (var lr2 = 0; lr2 < maxL; lr2++) {
        var kArr = (lVr[0].values || []);
        var lk = lr2 < kArr.length && kArr[lr2].length > 0 ? String(kArr[lr2][0]).trim().toUpperCase() : '';
        if (!lk) continue;
        var lv2 = [];
        for (var lc = 1; lc < lVr.length; lc++) {
          var lcv = (lVr[lc].values || []);
          lv2.push(lr2 < lcv.length && (lcv[lr2] || []).length > 0 ? lcv[lr2][0] : '');
        }
        if (!lMap[lk]) lMap[lk] = lv2;
      }

      var lKeys = ms.getRange(2, 10, lrV - 1, 1).getValues();
      var lOut = [];
      var lMatched = 0;
      for (var li2 = 0; li2 < lKeys.length; li2++) {
        var lk2 = String(lKeys[li2][0]).trim().toUpperCase();
        if (lk2 && lMap[lk2]) { lOut.push(lMap[lk2]); lMatched++; }
        else { lOut.push(['','','','','','']); }
      }
      if (lOut.length > 0) {
        ms.getRange(2, 13, lOut.length, 6).setValues(lOut);
      }
      log.push('VLOOKUP: ' + lMatched + ' matched');
    }
  } catch(e) { log.push('VLOOKUP ERR: ' + e.message); }

  // Compute formulas: R, S, T, U
  try {
    computeFormulas();
    log.push('Formulas: S,T,U computed');
  } catch(e) { log.push('Formulas ERR: ' + e.message); }

  ui.alert('ALL DONE!\nPulled: ' + totalP + '\nAdded: ' + totalA + '\n' + log.join('\n'));
}

// ===================== CLEAR & RUN FRESH =====================

function clearAndRun() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('This will DELETE all data and re-sync everything.\n\nContinue?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName(MASTER_TAB);
  if (ms) {
    var lr = ms.getLastRow();
    if (lr > 1) {
      ms.getRange(2, 1, lr - 1, ms.getMaxColumns()).clearContent();
      SpreadsheetApp.flush();
    }
  }
  runEverything();
}

// ===================== SETUP =====================

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName(MASTER_TAB) || ss.insertSheet(MASTER_TAB);
  ms.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  var hr = ms.getRange(1, 1, 1, HEADERS.length);
  hr.setFontWeight('bold');
  hr.setBackground('#4472C4');
  hr.setFontColor('white');
  SpreadsheetApp.getUi().alert(MASTER_TAB + ' tab ready! Headers set in row 1.');
}

// ===================== AUTO-REFRESH =====================

function setupAutoRefresh() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoRefresh') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('autoRefresh').timeBased().everyMinutes(1).create();
  SpreadsheetApp.getUi().alert('Auto-refresh ON! Runs every 1 minute.');
}

function stopAutoRefresh() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoRefresh') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  SpreadsheetApp.getUi().alert('Auto-refresh OFF. Removed ' + removed + ' trigger(s).');
}

function autoRefresh() {
  // Auto-refresh does a CLEAR + full re-pull to avoid duplicating rows on each run
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Logger.log('autoRefresh skipped — previous run still going.'); return; }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(MASTER_TAB);
    if (!ms) { lock.releaseLock(); return; }

    // Clear existing data (keep headers)
    var lr = ms.getLastRow();
    if (lr > 1) {
      ms.getRange(2, 1, lr - 1, ms.getMaxColumns()).clearContent();
      SpreadsheetApp.flush();
    }

    // Re-pull ALL data from all sources — NO dedup
    var names = Object.keys(PLATFORMS);
    for (var i = 0; i < names.length; i++) {
      var name = names[i];
      var p = PLATFORMS[name];
      try {
        var raw = fastRead(p.id, p.tab, p.cols);
        var newRows = [];
        for (var r = 0; r < raw.length; r++) {
          var key = String(raw[r][0]).trim();
          if (!key) continue;
          newRows.push(raw[r]);
        }
        // Fix negative Transfer L values for Amazon
        _fixTransferL(newRows);
        if (newRows.length > 0) {
          var lrNow = Math.max(ms.getLastRow(), 1);
          ms.getRange(lrNow + 1, 1, newRows.length, 12).setValues(newRows);
        }
        SpreadsheetApp.flush();
      } catch(e) {
        Logger.log('autoRefresh ' + name + ' ERR: ' + e.message);
      }
    }

    // Re-run VLOOKUP and formulas after fresh pull
    try { vlookupSKUsilent(); } catch(e) { Logger.log('autoRefresh VLOOKUP ERR: ' + e.message); }
    try { computeFormulas(); } catch(e) { Logger.log('autoRefresh Formulas ERR: ' + e.message); }

    lock.releaseLock();
  } catch(e) {
    Logger.log('autoRefresh ERR: ' + e.message);
    lock.releaseLock();
  }
}

// Silent version of vlookupSKU (no UI alerts) for use in autoRefresh
function vlookupSKUsilent() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName(MASTER_TAB);
  if (!ms) return;
  var lr = ms.getLastRow();
  if (lr < 2) return;

  var lookupCols = ['AA2:AA','AC2:AC','AD2:AD','AG2:AG','AJ2:AJ','AM2:AM','AQ2:AQ'];
  var ranges = [];
  for (var i = 0; i < lookupCols.length; i++) {
    ranges.push("'" + SKU_TAB + "'!" + lookupCols[i]);
  }
  var res = Sheets.Spreadsheets.Values.batchGet(SKU_SHEET_ID, {ranges: ranges});
  var vr = res.valueRanges;
  var maxLookup = 0;
  for (var v = 0; v < vr.length; v++) { var len = (vr[v].values || []).length; if (len > maxLookup) maxLookup = len; }

  var lookupMap = {};
  for (var r = 0; r < maxLookup; r++) {
    var keyArr = (vr[0].values || []);
    var key = r < keyArr.length && keyArr[r].length > 0 ? String(keyArr[r][0]).trim().toUpperCase() : '';
    if (!key) continue;
    var vals = [];
    for (var c = 1; c < vr.length; c++) {
      var cv = (vr[c].values || []);
      vals.push(r < cv.length && (cv[r] || []).length > 0 ? cv[r][0] : '');
    }
    if (!lookupMap[key]) lookupMap[key] = vals;
  }

  var keys = ms.getRange(2, 10, lr - 1, 1).getValues();
  var output = [];
  for (var i = 0; i < keys.length; i++) {
    var k = String(keys[i][0]).trim().toUpperCase();
    if (k && lookupMap[k]) { output.push(lookupMap[k]); }
    else { output.push(['','','','','','']); }
  }
  if (output.length > 0) {
    ms.getRange(2, 13, output.length, 6).setValues(output);
  }
}

// ===================== WEB APP API (doGet) =====================

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.test === '1') {
      return ContentService.createTextOutput(JSON.stringify({ok:true,ts:new Date().toISOString()})).setMimeType(ContentService.MimeType.JSON);
    }

    var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var tab = MASTER_TAB;
    var tz = Session.getScriptTimeZone();
    var now = new Date();

    // Month filter: ?month=January,February&year=2025 (supports comma-separated months)
    var filterMonthRaw = (e && e.parameter && e.parameter.month) ? String(e.parameter.month).trim() : '';
    var filterMonths = {};
    if (filterMonthRaw) {
      var parts = filterMonthRaw.split(',');
      for (var fm = 0; fm < parts.length; fm++) {
        var fmv = parts[fm].trim().toLowerCase();
        if (fmv) filterMonths[fmv] = true;
      }
    }
    var hasMonthFilter = Object.keys(filterMonths).length > 0;
    var filterYear = (e && e.parameter && e.parameter.year) ? String(e.parameter.year).trim() : '';

    // SPLIT into 2 smaller batchGet calls to avoid 413 on Sheets API (117k rows × 17 cols = too large)
    // Batch 1: A,C,D,E,F,G,H,K,L (core columns - 9 cols)
    var batch1Ranges = ['A2:A','C2:C','D2:D','E2:E','F2:F','G2:G','H2:H','K2:K','L2:L'];
    var r1 = [];
    for (var i = 0; i < batch1Ranges.length; i++) r1.push("'" + tab + "'!" + batch1Ranges[i]);
    var res1 = Sheets.Spreadsheets.Values.batchGet(ssId, {ranges: r1, valueRenderOption: 'UNFORMATTED_VALUE'});
    var vr1 = res1.valueRanges;

    // Batch 2: T,U,O (P&L cards + category - 3 cols)
    var batch2Ranges = ['T2:T','U2:U','O2:O'];
    var r2 = [];
    for (var i = 0; i < batch2Ranges.length; i++) r2.push("'" + tab + "'!" + batch2Ranges[i]);
    var res2 = Sheets.Spreadsheets.Values.batchGet(ssId, {ranges: r2, valueRenderOption: 'UNFORMATTED_VALUE'});
    var vr2 = res2.valueRanges;

    // Merge into unified cols array
    var allVr = [];
    for (var v = 0; v < vr1.length; v++) allVr.push(vr1[v]);
    for (var v = 0; v < vr2.length; v++) allVr.push(vr2[v]);

    var maxRows = 0;
    for (var v = 0; v < allVr.length; v++) { var len = (allVr[v].values || []).length; if (len > maxRows) maxRows = len; }

    if (maxRows === 0) {
      return _jsonResp({summary:{},byPlatform:{},byCompany:{},byType:{},byType2:{},byMonth:[],daily:[],plBreakdown:{revenue:0,costs:0,netPL:0},timestamp:new Date().toISOString()});
    }

    var cols = [];
    for (var c = 0; c < allVr.length; c++) cols.push(allVr[c].values || []);
    // cols: 0=A(total), 1=C(year), 2=D(month), 3=E(platform), 4=F(company), 5=G(type), 6=H(type2), 7=K(qty), 8=L(amount), 9=T(col20), 10=U(col21), 11=O(category)

    var summary = {totalEntries:0, totalQty:0, totalAmount:0};
    var byPlatform = {}, byCompany = {}, byType = {}, byType2 = {}, byMonthMap = {}, dailyMap = {};
    var typeDetail = {};
    var plRevenue = 0, plCosts = 0, plCOP = 0, plCardRevenue = 0, plReturns = 0;
    var plRevenueByPlatform = {}, plRevenueByCompany = {};
    var plCOPByPlatform = {}, plCOPByCompany = {};
    var plReturnsByPlatform = {}, plReturnsByCompany = {};
    var byCategory = {};

    // NEW column indices after split: 0=A(total), 1=C(year), 2=D(month), 3=E(platform), 4=F(company), 5=G(type), 6=H(type2), 7=K(qty), 8=L(amount), 9=T, 10=U, 11=O(category)
    var CI = {total:0, year:1, month:2, platform:3, company:4, type:5, type2:6, qty:7, amount:8, colT:9, colU:10, category:11};

    // Collect ALL unique months/years (before filtering) for dropdown population
    var allMonthNames = {}, allYears = {}, monthsByYear = {};
    for (var r0 = 0; r0 < maxRows; r0++) {
      var _t = r0 < cols[CI.total].length && cols[CI.total][r0].length > 0 ? cols[CI.total][r0][0] : '';
      if (!_t || String(_t).trim() === '') continue;
      var _y = String(r0 < cols[CI.year].length && cols[CI.year][r0].length > 0 ? cols[CI.year][r0][0] : '').trim();
      var _m = String(r0 < cols[CI.month].length && cols[CI.month][r0].length > 0 ? cols[CI.month][r0][0] : '').trim();
      if (_m) allMonthNames[_m] = true;
      if (_y) allYears[_y] = true;
      if (_y && _m) {
        if (!monthsByYear[_y]) monthsByYear[_y] = {};
        monthsByYear[_y][_m] = true;
      }
    }
    // Convert monthsByYear sets to sorted arrays (calendar order)
    var monthOrder = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    var monthsByYearSorted = {};
    var yKeys = Object.keys(monthsByYear);
    for (var yi = 0; yi < yKeys.length; yi++) {
      var yKey = yKeys[yi];
      var yMonths = [];
      for (var moi = 0; moi < monthOrder.length; moi++) {
        if (monthsByYear[yKey][monthOrder[moi]]) yMonths.push(monthOrder[moi]);
      }
      var extraM = Object.keys(monthsByYear[yKey]);
      for (var emi = 0; emi < extraM.length; emi++) {
        if (yMonths.indexOf(extraM[emi]) === -1) yMonths.push(extraM[emi]);
      }
      monthsByYearSorted[yKey] = yMonths;
    }

    for (var r = 0; r < maxRows; r++) {
      var total = r < cols[CI.total].length && cols[CI.total][r].length > 0 ? cols[CI.total][r][0] : '';
      if (!total || String(total).trim() === '') continue;

      var year = String(r < cols[CI.year].length && cols[CI.year][r].length > 0 ? cols[CI.year][r][0] : '').trim();
      var month = String(r < cols[CI.month].length && cols[CI.month][r].length > 0 ? cols[CI.month][r][0] : '').trim();
      var platform = String(r < cols[CI.platform].length && cols[CI.platform][r].length > 0 ? cols[CI.platform][r][0] : '').trim() || 'Unknown';
      var company = String(r < cols[CI.company].length && cols[CI.company][r].length > 0 ? cols[CI.company][r][0] : '').trim() || 'Unknown';
      var type = String(r < cols[CI.type].length && cols[CI.type][r].length > 0 ? cols[CI.type][r][0] : '').trim() || 'Unknown';
      var type2 = String(r < cols[CI.type2].length && cols[CI.type2][r].length > 0 ? cols[CI.type2][r][0] : '').trim() || 'Unknown';
      var qty = _toNum(r < cols[CI.qty].length && cols[CI.qty][r].length > 0 ? cols[CI.qty][r][0] : 0);
      var amount = _toNum(r < cols[CI.amount].length && cols[CI.amount][r].length > 0 ? cols[CI.amount][r][0] : 0);
      var colT = _toNum(r < cols[CI.colT].length && cols[CI.colT][r].length > 0 ? cols[CI.colT][r][0] : 0);
      var colU = _toNum(r < cols[CI.colU].length && cols[CI.colU][r].length > 0 ? cols[CI.colU][r][0] : 0);
      var category = String(r < cols[CI.category].length && cols[CI.category][r].length > 0 ? cols[CI.category][r][0] : '').trim() || 'Unknown';

      // Convert Transfer values to positive
      var typeUpper = type.toUpperCase();
      if (typeUpper === 'TRANSFER') {
        amount = Math.abs(amount);
        colT = Math.abs(colT);
        colU = Math.abs(colU);
      }

      // Month key for grouping
      var monthKey = '';
      if (year && month) {
        var mn = _monthToNum(month);
        monthKey = year + '-' + mn;
      }

      // Month/Year filter (column D = month name, column C = year)
      if (hasMonthFilter && !filterMonths[String(month).trim().toLowerCase()]) continue;
      if (filterYear && year !== filterYear) continue;

      // Summary
      summary.totalEntries++;
      summary.totalQty += qty;
      summary.totalAmount += amount;

      // P&L (old logic based on REVENUE_TYPES — kept for byType section)
      if (REVENUE_TYPES[type]) {
        plRevenue += amount;
      } else {
        plCosts += amount;
      }

      // NEW P&L cards: Revenue = SUM(U) where G="Order", COP = SUM(T) where G="Order", Returns = SUM(U) where G="Return"
      if (typeUpper === 'ORDER') {
        plCardRevenue += colU;
        plCOP += colT;
        // Per-platform Revenue/COP drill-down
        if (!plRevenueByPlatform[platform]) plRevenueByPlatform[platform] = {value:0, entries:0, companies:{}};
        plRevenueByPlatform[platform].value += colU;
        plRevenueByPlatform[platform].entries++;
        if (!plRevenueByPlatform[platform].companies[company]) plRevenueByPlatform[platform].companies[company] = {value:0, entries:0};
        plRevenueByPlatform[platform].companies[company].value += colU;
        plRevenueByPlatform[platform].companies[company].entries++;
        if (!plCOPByPlatform[platform]) plCOPByPlatform[platform] = {value:0, entries:0, companies:{}};
        plCOPByPlatform[platform].value += colT;
        plCOPByPlatform[platform].entries++;
        if (!plCOPByPlatform[platform].companies[company]) plCOPByPlatform[platform].companies[company] = {value:0, entries:0};
        plCOPByPlatform[platform].companies[company].value += colT;
        plCOPByPlatform[platform].companies[company].entries++;
        // Per-company Revenue/COP drill-down
        if (!plRevenueByCompany[company]) plRevenueByCompany[company] = {value:0, entries:0};
        plRevenueByCompany[company].value += colU;
        plRevenueByCompany[company].entries++;
        if (!plCOPByCompany[company]) plCOPByCompany[company] = {value:0, entries:0};
        plCOPByCompany[company].value += colT;
        plCOPByCompany[company].entries++;
      }
      if (typeUpper === 'RETURN') {
        plReturns += colU;
        // Per-platform Returns drill-down
        if (!plReturnsByPlatform[platform]) plReturnsByPlatform[platform] = {value:0, entries:0, companies:{}};
        plReturnsByPlatform[platform].value += colU;
        plReturnsByPlatform[platform].entries++;
        if (!plReturnsByPlatform[platform].companies[company]) plReturnsByPlatform[platform].companies[company] = {value:0, entries:0};
        plReturnsByPlatform[platform].companies[company].value += colU;
        plReturnsByPlatform[platform].companies[company].entries++;
        // Per-company Returns drill-down
        if (!plReturnsByCompany[company]) plReturnsByCompany[company] = {value:0, entries:0};
        plReturnsByCompany[company].value += colU;
        plReturnsByCompany[company].entries++;
      }

      // byCategory only (bySubCategory, byProduct, bySKU removed to reduce API read size)
      if (typeUpper === 'ORDER') {
        if (!byCategory[category]) byCategory[category] = {revenue:0, cop:0, returns:0, revCount:0, retCount:0};
        byCategory[category].revenue += colU;
        byCategory[category].cop += colT;
        byCategory[category].revCount++;
      }
      if (typeUpper === 'RETURN') {
        if (!byCategory[category]) byCategory[category] = {revenue:0, cop:0, returns:0, revCount:0, retCount:0};
        byCategory[category].returns += colU;
        byCategory[category].retCount++;
      }

      // byPlatform
      if (!byPlatform[platform]) byPlatform[platform] = {entries:0, qty:0, amount:0};
      byPlatform[platform].entries++;
      byPlatform[platform].qty += qty;
      byPlatform[platform].amount += amount;

      // byCompany
      if (!byCompany[company]) byCompany[company] = {entries:0, qty:0, amount:0};
      byCompany[company].entries++;
      byCompany[company].qty += qty;
      byCompany[company].amount += amount;

      // byType
      if (!byType[type]) byType[type] = {entries:0, qty:0, amount:0};
      byType[type].entries++;
      byType[type].qty += qty;
      byType[type].amount += amount;

      // typeDetail: per-type drill-down by platform and company (using column L amount)
      if (!typeDetail[type]) typeDetail[type] = {total:0, entries:0, byPlatform:{}, byCompany:{}};
      typeDetail[type].total += amount;
      typeDetail[type].entries++;
      if (!typeDetail[type].byPlatform[platform]) typeDetail[type].byPlatform[platform] = {value:0, entries:0, byCompany:{}};
      typeDetail[type].byPlatform[platform].value += amount;
      typeDetail[type].byPlatform[platform].entries++;
      if (!typeDetail[type].byPlatform[platform].byCompany[company]) typeDetail[type].byPlatform[platform].byCompany[company] = {value:0, entries:0};
      typeDetail[type].byPlatform[platform].byCompany[company].value += amount;
      typeDetail[type].byPlatform[platform].byCompany[company].entries++;
      if (!typeDetail[type].byCompany[company]) typeDetail[type].byCompany[company] = {value:0, entries:0};
      typeDetail[type].byCompany[company].value += amount;
      typeDetail[type].byCompany[company].entries++;

      // byType2
      if (!byType2[type2]) byType2[type2] = {entries:0, qty:0, amount:0};
      byType2[type2].entries++;
      byType2[type2].qty += qty;
      byType2[type2].amount += amount;

      // byMonth
      if (monthKey) {
        if (!byMonthMap[monthKey]) byMonthMap[monthKey] = {month:monthKey, entries:0, qty:0, amount:0, revenue:0, costs:0};
        byMonthMap[monthKey].entries++;
        byMonthMap[monthKey].qty += qty;
        byMonthMap[monthKey].amount += amount;
        if (REVENUE_TYPES[type]) byMonthMap[monthKey].revenue += amount;
        else byMonthMap[monthKey].costs += amount;
      }
    }

    // Round
    summary.totalAmount = Math.round(summary.totalAmount * 100) / 100;
    summary.totalQty = Math.round(summary.totalQty * 100) / 100;
    for (var k in byPlatform) { byPlatform[k].amount = Math.round(byPlatform[k].amount * 100) / 100; byPlatform[k].qty = Math.round(byPlatform[k].qty * 100) / 100; }
    for (var k in byCompany) { byCompany[k].amount = Math.round(byCompany[k].amount * 100) / 100; byCompany[k].qty = Math.round(byCompany[k].qty * 100) / 100; }
    for (var k in byType) { byType[k].amount = Math.round(byType[k].amount * 100) / 100; byType[k].qty = Math.round(byType[k].qty * 100) / 100; }
    // Round byCategory
    for (var k in byCategory) { byCategory[k].revenue = Math.round(byCategory[k].revenue * 100) / 100; byCategory[k].cop = Math.round(byCategory[k].cop * 100) / 100; byCategory[k].returns = Math.round(byCategory[k].returns * 100) / 100; }
    // Round typeDetail
    for (var td in typeDetail) {
      typeDetail[td].total = Math.round(typeDetail[td].total * 100) / 100;
      for (var tdp in typeDetail[td].byPlatform) {
        typeDetail[td].byPlatform[tdp].value = Math.round(typeDetail[td].byPlatform[tdp].value * 100) / 100;
        for (var tdpc in typeDetail[td].byPlatform[tdp].byCompany) { typeDetail[td].byPlatform[tdp].byCompany[tdpc].value = Math.round(typeDetail[td].byPlatform[tdp].byCompany[tdpc].value * 100) / 100; }
      }
      for (var tdc in typeDetail[td].byCompany) { typeDetail[td].byCompany[tdc].value = Math.round(typeDetail[td].byCompany[tdc].value * 100) / 100; }
    }
    for (var k in byType2) { byType2[k].amount = Math.round(byType2[k].amount * 100) / 100; byType2[k].qty = Math.round(byType2[k].qty * 100) / 100; }

    var byMonth = [];
    var mKeys = Object.keys(byMonthMap).sort();
    for (var i = 0; i < mKeys.length; i++) {
      var m = byMonthMap[mKeys[i]];
      m.amount = Math.round(m.amount * 100) / 100;
      m.revenue = Math.round(m.revenue * 100) / 100;
      m.costs = Math.round(m.costs * 100) / 100;
      m.qty = Math.round(m.qty * 100) / 100;
      byMonth.push(m);
    }

    var plBreakdown = {
      revenue: Math.round(plRevenue * 100) / 100,
      costs: Math.round(plCosts * 100) / 100,
      netPL: Math.round((plRevenue + plCosts) * 100) / 100
    };

    // NEW P&L cards data: Revenue = SUM(U) where G="Order", COP = SUM(T) where G="Order", Returns = SUM(U) where G="Return"
    // Round per-platform/company values
    for (var rp in plRevenueByPlatform) { plRevenueByPlatform[rp].value = Math.round(plRevenueByPlatform[rp].value * 100) / 100; }
    for (var cp in plCOPByPlatform) { plCOPByPlatform[cp].value = Math.round(plCOPByPlatform[cp].value * 100) / 100; }
    for (var rc in plRevenueByCompany) { plRevenueByCompany[rc].value = Math.round(plRevenueByCompany[rc].value * 100) / 100; }
    for (var cc in plCOPByCompany) { plCOPByCompany[cc].value = Math.round(plCOPByCompany[cc].value * 100) / 100; }
    for (var rtp in plReturnsByPlatform) { plReturnsByPlatform[rtp].value = Math.round(plReturnsByPlatform[rtp].value * 100) / 100; }
    for (var rtc in plReturnsByCompany) { plReturnsByCompany[rtc].value = Math.round(plReturnsByCompany[rtc].value * 100) / 100; }

    var plCards = {
      revenue: Math.round(plCardRevenue * 100) / 100,
      cop: Math.round(plCOP * 100) / 100,
      returns: Math.round(plReturns * 100) / 100,
      revenueByPlatform: plRevenueByPlatform,
      copByPlatform: plCOPByPlatform,
      returnsByPlatform: plReturnsByPlatform,
      revenueByCompany: plRevenueByCompany,
      copByCompany: plCOPByCompany,
      returnsByCompany: plReturnsByCompany
    };

    // Sort months in calendar order for dropdown
    var sortedMonths = [];
    for (var mi = 0; mi < monthOrder.length; mi++) {
      if (allMonthNames[monthOrder[mi]]) sortedMonths.push(monthOrder[mi]);
    }
    var amK = Object.keys(allMonthNames);
    for (var ai = 0; ai < amK.length; ai++) {
      if (sortedMonths.indexOf(amK[ai]) === -1) sortedMonths.push(amK[ai]);
    }

    // === 413 FIX: keep split batch reads, restore full typeDetail for drill-downs ===
    // Only remove bySKU/byProduct/bySubCategory (huge with 117k rows)
    // typeDetail is small so keep full structure for frontend drill-downs

    // Slim typeDetail: remove deepest nesting (byCompany inside byPlatform)
    for (var td in typeDetail) {
      for (var tdp in typeDetail[td].byPlatform) {
        delete typeDetail[td].byPlatform[tdp].byCompany;
      }
    }

    var result = {
      summary: summary,
      byPlatform: byPlatform,
      byCompany: byCompany,
      byType: byType,
      byType2: byType2,
      byMonth: byMonth,
      plBreakdown: plBreakdown,
      plCards: plCards,
      typeDetail: typeDetail,
      byCategory: byCategory,
      bySubCategory: {},
      byProduct: {},
      bySKU: {},
      platforms: Object.keys(byPlatform).sort(),
      companies: Object.keys(byCompany).sort(),
      types: Object.keys(byType).sort(),
      availableMonths: sortedMonths,
      availableYears: Object.keys(allYears).sort(),
      monthsByYear: monthsByYearSorted,
      timestamp: new Date().toISOString()
    };

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return _jsonResp({ error: err.message });
  }
}

function _jsonResp(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function _toNum(v) { if (v === '' || v === null || v === undefined) return 0; var n = Number(v); return isNaN(n) ? 0 : n; }

// Fix Transfer column L: if platform contains "Amazon" and type = "Transfer", make L positive
function _fixTransferL(rows) {
  for (var i = 0; i < rows.length; i++) {
    var plat = String(rows[i][4] || '').trim().toUpperCase();
    var typ = String(rows[i][6] || '').trim().toUpperCase();
    if (plat.indexOf('AMAZON') > -1 && typ === 'TRANSFER') {
      var lVal = Number(rows[i][11]);
      if (!isNaN(lVal) && lVal < 0) rows[i][11] = Math.abs(lVal);
    }
  }
}
function _monthToNum(m) {
  var months = {'january':'01','february':'02','march':'03','april':'04','may':'05','june':'06','july':'07','august':'08','september':'09','october':'10','november':'11','december':'12',
    'jan':'01','feb':'02','mar':'03','apr':'04','jun':'06','jul':'07','aug':'08','sep':'09','oct':'10','nov':'11','dec':'12'};
  var s = String(m).toLowerCase().trim();
  if (months[s]) return months[s];
  // If numeric
  var n = parseInt(s, 10);
  if (n >= 1 && n <= 12) return String(n).length === 1 ? '0' + n : String(n);
  return '01';
}

// ===================== TEST =====================

function testPLApi() {
  var result = doGet({parameter:{}});
  var json = JSON.parse(result.getContent());
  Logger.log('Total Entries: ' + json.summary.totalEntries);
  Logger.log('Total Amount: ' + json.summary.totalAmount);
  Logger.log('P&L Revenue: ' + json.plBreakdown.revenue);
  Logger.log('P&L Costs: ' + json.plBreakdown.costs);
  Logger.log('Net P&L: ' + json.plBreakdown.netPL);
  Logger.log('Platforms: ' + (json.platforms || []).join(', '));
  Logger.log('Types: ' + (json.types || []).join(', '));
  SpreadsheetApp.getActiveSpreadsheet().toast('Entries: ' + json.summary.totalEntries + ' | Amount: ' + json.summary.totalAmount + ' | Net P/L: ' + json.plBreakdown.netPL, 'API Test OK', 10);
}
