// ═══════════════════════════════════════════════════════════════════
//  QualiCheck LIS — Google Apps Script  (v8 — ADD CATEGORY + ADD TEST)
//
//  CHANGES from v7:
//  ┌─────────────────────────────────────────────────────────────────┐
//  │  1. handleAddCategory() — adds a new category header row to    │
//  │     TESTS & PRICES sheet, styled bold + purple background.     │
//  │  2. handleAddTest() — inserts a new test row under its        │
//  │     category block with Regular + Senior pricing.             │
//  │  3. doPost() updated — routes 'addcategory' and 'addtest'.    │
//  └─────────────────────────────────────────────────────────────────┘
// ═══════════════════════════════════════════════════════════════════

// ── Sheet IDs ─────────────────────────────────────────────────────
const SHEET_ID     = '11w99hIQyQKSn8m1LPkq6SYgQ4_HxAQsIpf9-_2sF3LI';
const RAW_SHEET_ID = '10vx1OqNuKOKZjT4161JS_3TOb2HOaEHMZDCofoomtVM';

// ── Tab names ─────────────────────────────────────────────────────
const TAB = {
  LOGINS:       'Logins',
  DOCTORS:      'DOCTORS',
  TESTS:        'TESTS & PRICES',
  PRICELIST:    'PRICELIST',
  DISCOUNTS:    'DISCOUNTS',
  DAILY:        'DAILY REG SHEET',
  RECEIPT_LOGS: 'RECEIPT LOGS',
};

// ── Column indices — DAILY REG SHEET (1-based, A=1 … AE=31) ───────
const COL = {
  DATE:         1,   // A
  NO:           2,   // B
  NAME:         3,   // C
  BIRTHDAY:     4,   // D
  AGE_SEX:      5,   // E
  ADDRESS:      6,   // F
  OP_TYPE:      7,   // G
  CONTACT:      8,   // H
  REFERRED:     9,   // I
  COMPANY:      10,  // J
  PROCEDURES:   11,  // K  LAB PROCEDURES
  LAB_SUB:      12,  // L  LAB SUBTOTAL
  CHECKUP:      13,  // M  CHECKUP
  CHECKUP_SUB:  14,  // N  CHECKUP SUBTOTAL
  CARDIOLOGY:   15,  // O  CARDIOLOGY
  CAD_SUB:      16,  // P  CAD SUBTOTAL
  KITS:         17,  // Q  KITS
  KITS_SUB:     18,  // R  KITS SUBTOTAL
  XRAY_TESTS:   19,  // S  XRAY TESTS
  XRAY_SUB:     20,  // T  XRAY SUBTOTAL
  DRUG_TEST:    21,  // U  DRUG TEST
  DRUG_SUB:     22,  // V  DRUG SUBTOTAL
  UTZ_TESTS:    23,  // W  UTZ TESTS
  UTZ_SUB:      24,  // X  UTZ SUBTOTAL
  DISCOUNT:     25,  // Y  DISCOUNT
  DIS_TOTAL:    26,  // Z  DIS. TOTAL
  AMOUNT:       27,  // AA AMOUNT
  STATUS:       28,  // AB STATUS
  P_METHODS:    29,  // AC P. METHODS
  PAYMENT:      30,  // AD PAYMENT
  BALANCE:      31,  // AE BALANCE
  REMARKS:      32,  // AF REMARKS
  ENCODED_BY:   33,  // AG ENCODED BY
};

// ── Column indices — RECEIPT LOGS (1-based, A=1 … AB=28) ──────────
const RLOG = {
  DATE:         1,   // A
  NO:           2,   // B
  NAME:         3,   // C
  AGE_SEX:      4,   // D
  REFERRED:     5,   // E
  ITEMS_JSON:   6,   // F
  LAB_PROC:     7,   // G
  CHECKUP:      8,   // H
  CARDIOLOGY:   9,   // I
  KITS:         10,  // J
  XRAY:         11,  // K
  DRUG:         12,  // L
  UTZ:          13,  // M
  LAB_SUB:      14,  // N
  CHECKUP_SUB:  15,  // O
  CAD_SUB:      16,  // P
  KITS_SUB:     17,  // Q
  XRAY_SUB:     18,  // R
  DRUG_SUB:     19,  // S
  UTZ_SUB:      20,  // T
  DISC_LABEL:   21,  // U
  DISC_AMT:     22,  // V
  GRAND:        23,  // W
  STATUS:       24,  // X
  PAY_TYPE:     25,  // Y
  PAY_AMT:      26,  // Z
  BALANCE:      27,  // AA
  CASH_AMT:     28,  // AB (Split payment: cash portion)
  GCASH_AMT:    29,  // AC (Split payment: GCash portion)
  GCASH_REF:    30,  // AD (Split payment: GCash reference number)
  TIMESTAMP:    31,  // AE
  ENCODED_BY:   32,  // AF ENCODED BY
};

const RLOG_HEADERS = [
  'DATE', 'NO.', 'NAME', 'AGE/SEX', 'REFERRED BY',
  'ITEMS JSON',
  'LAB PROCEDURES', 'CHECKUP', 'CARDIOLOGY', 'KITS',
  'XRAY TESTS', 'DRUG TESTS', 'UTZ TESTS',
  'LAB SUBTOTAL', 'CHECKUP SUBTOTAL', 'CAD SUBTOTAL', 'KITS SUBTOTAL',
  'XRAY SUBTOTAL', 'DRUG SUBTOTAL', 'UTZ SUBTOTAL',
  'DISC LABEL', 'DISC AMT', 'GRAND TOTAL',
  'STATUS', 'PAY TYPE', 'PAY AMT', 'BALANCE',
  'CASH AMT', 'GCASH AMT', 'GCASH REF',
  'TIMESTAMP',
  'ENCODED BY',
];

const RLOG_PESO_COLS = [
  RLOG.LAB_SUB, RLOG.CHECKUP_SUB, RLOG.CAD_SUB, RLOG.KITS_SUB,
  RLOG.XRAY_SUB, RLOG.DRUG_SUB, RLOG.UTZ_SUB,
  RLOG.DISC_AMT, RLOG.GRAND, RLOG.PAY_AMT, RLOG.BALANCE,
  RLOG.CASH_AMT, RLOG.GCASH_AMT,
];

const STATUS_STYLES = {
  'PENDING':     { bg: '#EDE8FF', fg: '#3D0099' },
  'FULLY PAID':  { bg: '#1B5E20', fg: '#FFFFFF' },
  'PARTLY PAID': { bg: '#FFF9C4', fg: '#000000' },
  'UNPAID':      { bg: '#B71C1C', fg: '#FFFFFF' },
  'REFUNDED':    { bg: '#212121', fg: '#FFFFFF' },
};

const PMETHOD_STYLES = {
  'CASH':           { bg: '#B8860B', fg: '#FFFFFF' },
  'GCASH':          { bg: '#0D47A1', fg: '#FFFFFF' },
  'CASH + GCASH':   { bg: '#8E24AA', fg: '#FFFFFF' },
  'YAKAP':          { bg: '#1B5E20', fg: '#FFFFFF' },
  'BSA':            { bg: '#4A148C', fg: '#FFFFFF' },
  'BANK':           { bg: '#B71C1C', fg: '#FFFFFF' },
  'CARD':           { bg: '#B71C1C', fg: '#FFFFFF' },
  'CHARGED':        { bg: '#37474F', fg: '#FFFFFF' },
};

const DAILY_HEADERS = [
  'DATE', 'NO.', 'NAME', 'BIRTHDAY', 'AGE/SEX', 'ADDRESS',
  'OP TYPE', 'CONTACT', 'REFERRED BY', 'COMPANY',
  'LAB PROCEDURES',  'LAB SUBTOTAL',
  'CHECKUP',         'CHECKUP SUBTOTAL',
  'CARDIOLOGY',      'CAD SUBTOTAL',
  'KITS',            'KITS SUBTOTAL',
  'XRAY TESTS',      'XRAY SUBTOTAL',
  'DRUG TEST',       'DRUG SUBTOTAL',
  'UTZ TESTS',       'UTZ SUBTOTAL',
  'DISCOUNT', 'DIS. TOTAL', 'AMOUNT',
  'STATUS', 'P. METHODS', 'PAYMENT', 'BALANCE', 'REMARKS', 'ENCODED BY',
];

const DAILY_COL_WIDTHS = [
  90, 60, 180, 90, 80, 150, 100, 110, 150, 130,  // A-J
  280, 110,                                        // K-L  LAB
  200, 110, 200, 110, 200, 110,                    // M-R  CHECKUP/CARDIOLOGY/KITS
  200, 110, 150, 110, 200, 110,                    // S-X  XRAY/DRUG/UTZ
  200, 100, 100, 120, 120, 100, 100, 100, 130           // Y-AE DISCOUNT→ENCODED BY
];

const PESO_COLS = [
  COL.LAB_SUB, COL.CHECKUP_SUB, COL.CAD_SUB, COL.KITS_SUB,
  COL.XRAY_SUB, COL.DRUG_SUB, COL.UTZ_SUB,
  COL.DIS_TOTAL, COL.AMOUNT, COL.PAYMENT, COL.BALANCE,
];

// ═══════════════════════════════════════════════════════════════════
//  ROUTING
// ═══════════════════════════════════════════════════════════════════
function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();
  try {
    if (action === 'login')        return json(handleLogin(e.parameter.user, e.parameter.pass));
    if (action === 'getprices')    return json(handleGetPrices());
    if (action === 'getdiscounts') return json(handleGetDiscounts());
    if (action === 'getdoctors')   return json(handleGetDoctors());
    if (action === 'adddoctor')    return json(handleAddDoctor(e.parameter.name, e.parameter.details));
    if (action === 'gettoday')     return json(handleGetToday());
    if (action === 'getpatients')  return json(handleGetPatients(e.parameter));
    return json({ success: false, message: 'Unknown action: ' + action });
  } catch (err) {
    return json({ success: false, message: err.message });
  }
}

function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = (data.action || '').toLowerCase();
    if (action === 'register')      return json(handleRegister(data));
    if (action === 'deletepatient') return json(handleDeletePatient(data));
    if (action === 'refundpatient') return json(handleRefundPatient(data));
    if (action === 'editpatient')   return json(handleEditPatient(data));
    if (action === 'addcategory')   return json(handleAddCategory(data));   // ← NEW v8
    if (action === 'addtest')       return json(handleAddTest(data));       // ← NEW v8
    if (action === 'edittest')      return json(handleEditTest(data));      // ← NEW v9
    return json({ success: false, message: 'Unknown POST action.' });
  } catch (err) {
    return json({ success: false, message: 'Error: ' + err.message });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════
//  LOGIN
// ═══════════════════════════════════════════════════════════════════
function handleLogin(username, password) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB.LOGINS);

  if (!sheet) {
    const DEMO = [{ u: 'admin', p: 'admin2026' }, { u: 'labtech', p: 'lab@lis2026' }];
    const ok   = DEMO.find(x => x.u === username && x.p === password);
    return { success: !!ok, message: ok ? 'OK' : 'Invalid credentials.' };
  }

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === username &&
        String(rows[i][1]).trim() === password) {
      const displayName = String(rows[i][2] || '').trim() || _toTitleCase(username);
      return { success: true, message: 'OK', displayName };
    }
  }
  return { success: false, message: 'Invalid username or password.' };
}

function _toTitleCase(str) {
  return (str || '').replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
}

// ═══════════════════════════════════════════════════════════════════
//  GET PRICES
// ═══════════════════════════════════════════════════════════════════
function handleGetPrices() {
  const ss         = SpreadsheetApp.openById(SHEET_ID);
  const testsSheet = ss.getSheetByName(TAB.TESTS);
  const plSheet    = ss.getSheetByName(TAB.PRICELIST);
 
  if (!testsSheet && !plSheet)
    return { success: false, message: 'No price sheet found. Run initialSetup() first.' };
 
  const prices     = {};
  const categories = {};
 
  function parsePrice(val) {
    if (typeof val === 'number' && !isNaN(val) && val > 0) return val;
    if (typeof val === 'string') {
      const n = parseFloat(val.replace(/[\u20b1,\s]/g, '').trim());
      if (!isNaN(n) && n > 0) return n;
    }
    return null;
  }
 
  if (testsSheet) {
    const rows = testsSheet.getDataRange().getValues();
    let currentCat = '';
    for (let i = 1; i < rows.length; i++) {
      const catCell  = String(rows[i][0] || '').trim();
      const testName = String(rows[i][1] || '').trim();
 
      // ── Price columns ─────────────────────────────────────────────
      const r     = parsePrice(rows[i][2]);   // Col C — Regular
      const s     = parsePrice(rows[i][3]);   // Col D — Senior
      // Col E (index 4) is reserved / notes — skip
      const bsa   = parsePrice(rows[i][5]);   // Col F — BSA Price      ← NEW
      const om    = parsePrice(rows[i][6]);   // Col G — QSD OM Price   ← NEW
      const yakap = parsePrice(rows[i][7]);   // Col H — YAKAP Price    ← NEW
      const amante = parsePrice(rows[i][8]);   // Col I — AMANTE Price    ← NEW
 
      // Category header row — col A has content, col B is blank
      if (catCell) {
        currentCat = catCell;
        if (!categories[currentCat]) categories[currentCat] = [];
        continue;
      }
 
      // Test row — must have a name and at least one price
      if (!testName || (r === null && s === null)) continue;
 
      // Store all price tiers; null means "not set → fall back to regular"
      prices[testName.toUpperCase()] = { r, s, bsa, om, yakap, amante };  // ← bsa/om/yakap/amante added
 
      if (currentCat) {
        if (!categories[currentCat]) categories[currentCat] = [];
        categories[currentCat].push(testName);
      }
    }
  } else {
    // Fallback: read from legacy PRICELIST tab (no BSA/OM/YAKAP/AMANTE columns there)
    const rows = plSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const name = String(rows[i][0] || '').trim().toUpperCase();
      if (!name) continue;
      prices[name] = {
        r:     parsePrice(rows[i][1]),
        s:     parsePrice(rows[i][2]),
        bsa:   null,
        om:    null,
        yakap: null,
        amante: null,
      };
    }
  }
 
  return { success: true, prices, categories };
}

// ═══════════════════════════════════════════════════════════════════
//  GET DISCOUNTS
// ═══════════════════════════════════════════════════════════════════
function handleGetDiscounts() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB.DISCOUNTS);
  if (!sheet) return { success: false, message: '"DISCOUNTS" tab not found.' };

  const rows      = sheet.getDataRange().getValues();
  const discounts = [];
  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][0] || '').trim();
    if (!name) continue;
    let rate = rows[i][1];
    if (typeof rate === 'string') rate = parseFloat(rate.replace('%', '').trim());
    if (typeof rate === 'number' && !isNaN(rate)) {
      if (rate > 1) rate = rate / 100;
      discounts.push({ name, rate });
    }
  }
  return { success: true, discounts };
}

// ═══════════════════════════════════════════════════════════════════
//  GET DOCTORS
// ═══════════════════════════════════════════════════════════════════
function handleGetDoctors() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB.DOCTORS);
  if (!sheet) return { success: true, doctors: [] };

  const rows    = sheet.getDataRange().getValues();
  const doctors = [];
  for (let i = 1; i < rows.length; i++) {
    const name    = String(rows[i][0] || '').trim();
    const details = String(rows[i][1] || '').trim();
    if (name) doctors.push({ name, details });
  }
  return { success: true, doctors };
}

// ═══════════════════════════════════════════════════════════════════
//  ADD DOCTOR
// ═══════════════════════════════════════════════════════════════════
function handleAddDoctor(name, details) {
  if (!name) return { success: false, message: 'Name is required.' };

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(TAB.DOCTORS);
  if (!sheet) sheet = _createDoctorsTab(ss);

  const rows = sheet.getDataRange().getValues();
  const dup  = rows.slice(1).some(r =>
    String(r[0]).trim().toLowerCase() === name.trim().toLowerCase()
  );
  if (dup) return { success: true, message: 'Already exists.' };

  sheet.appendRow([name.trim().toUpperCase(), (details || '').trim(), new Date()]);
  return { success: true, message: 'Added.' };
}

// ═══════════════════════════════════════════════════════════════════
//  ADD CATEGORY — NEW v8
//  Appends a new category header row to TESTS & PRICES, then
//  re-syncs the PRICELIST tab.
// ═══════════════════════════════════════════════════════════════════
function handleAddCategory(data) {
  const categoryName = String(data.categoryName || '').trim().toUpperCase();
  if (!categoryName) return { success: false, message: 'Category name is required.' };

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB.TESTS);
  if (!sheet) return { success: false, message: '"TESTS & PRICES" tab not found.' };

  // Check for duplicate category
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const cat = String(rows[i][0] || '').trim().toUpperCase();
    if (cat === categoryName) {
      return { success: false, message: 'Category "' + categoryName + '" already exists.' };
    }
  }

  // Append category header row — col A = name, B/C blank, D = notes
  sheet.appendRow([categoryName, '', '', data.notes || '']);

  // Style it: bold, purple-tinted background
  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 1, 1, 4)
       .setFontWeight('bold')
       .setBackground('#EDE8FF')
       .setFontColor('#4A0E8F')
       .setFontSize(10);

  // Re-sync PRICELIST
  try { syncPriceList(); } catch(e) {
    Logger.log('syncPriceList after addCategory failed (non-fatal): ' + e.message);
  }

  Logger.log('handleAddCategory: Added "' + categoryName + '" at row ' + newRow);
  return { success: true, message: 'Category "' + categoryName + '" added.' };
}

// ═══════════════════════════════════════════════════════════════════
//  ADD TEST — NEW v8
//  Inserts a new test row directly after the last test in its
//  category block, so the sheet stays grouped by category.
//  Then re-syncs the PRICELIST tab.
// ═══════════════════════════════════════════════════════════════════
function handleAddTest(data) {
  const categoryName = String(data.categoryName || '').trim().toUpperCase();
  const testName     = String(data.testName     || '').trim().toUpperCase();
  const priceRegular = parseFloat(data.priceRegular) || null;
  const priceSenior  = parseFloat(data.priceSenior)  || priceRegular;

  if (!categoryName) return { success: false, message: 'Category name is required.' };
  if (!testName)     return { success: false, message: 'Test name is required.' };
  if (!priceRegular) return { success: false, message: 'Regular price is required.' };

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB.TESTS);
  if (!sheet) return { success: false, message: '"TESTS & PRICES" tab not found.' };

  const rows = sheet.getDataRange().getValues();

  // Walk the sheet to find the category block and its last row
  let inTargetCategory  = false;
  let lastRowOfCategory = -1; // 1-based sheet row number

  for (let i = 1; i < rows.length; i++) {
    const catCell  = String(rows[i][0] || '').trim().toUpperCase();
    const testCell = String(rows[i][1] || '').trim().toUpperCase();

    if (catCell) {
      if (catCell === categoryName) {
        inTargetCategory  = true;
        lastRowOfCategory = i + 1; // the category header itself is the starting point
      } else if (inTargetCategory) {
        // Hit the next category — stop
        break;
      }
    }

    if (inTargetCategory) {
      // Duplicate check
      if (testCell && testCell === testName) {
        return {
          success: false,
          message: 'Test "' + testName + '" already exists in "' + categoryName + '".',
        };
      }
      // Keep advancing the last-row pointer
      if (testCell || catCell === categoryName) {
        lastRowOfCategory = i + 1;
      }
    }
  }

  if (lastRowOfCategory === -1) {
    return { success: false, message: 'Category "' + categoryName + '" not found. Add it first.' };
  }

  // Insert a new row right after the last row of this category
  sheet.insertRowAfter(lastRowOfCategory);
  const newRow = lastRowOfCategory + 1;

  // Write: col A = blank (not a category), col B = test name, C = regular, D = senior
  sheet.getRange(newRow, 1, 1, 4).setValues([
    ['', testName, priceRegular, priceSenior || priceRegular]
  ]);

  // Format price columns as number
  sheet.getRange(newRow, 3, 1, 2).setNumberFormat('#,##0.00');

  // Optional notes in col E
  if (data.notes) {
    sheet.getRange(newRow, 5).setValue(String(data.notes).trim());
  }

  // Re-sync PRICELIST
  try { syncPriceList(); } catch(e) {
    Logger.log('syncPriceList after addTest failed (non-fatal): ' + e.message);
  }

  Logger.log('handleAddTest: Added "' + testName + '" under "' + categoryName + '" at row ' + newRow);
  return {
    success: true,
    message: 'Test "' + testName + '" added under "' + categoryName + '".',
    row: newRow,
  };
}

// ═══════════════════════════════════════════════════════════════════
//  EDIT TEST — NEW v9
//  Updates the Regular + Senior price of an existing test row
//  inside its category block in TESTS & PRICES, then re-syncs
//  the PRICELIST tab.
// ═══════════════════════════════════════════════════════════════════
function handleEditTest(data) {
  const categoryName = String(data.categoryName || '').trim().toUpperCase();
  const testName     = String(data.testName     || '').trim().toUpperCase();
  const priceRegular = parseFloat(data.priceRegular) || null;
  const priceSenior  = parseFloat(data.priceSenior)  || priceRegular;

  if (!categoryName) return { success: false, message: 'Category name is required.' };
  if (!testName)     return { success: false, message: 'Test name is required.' };
  if (!priceRegular) return { success: false, message: 'Regular price is required.' };

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB.TESTS);
  if (!sheet) return { success: false, message: '"TESTS & PRICES" tab not found.' };

  const rows = sheet.getDataRange().getValues();
  let inTargetCategory = false;
  let foundRow = -1;

  for (let i = 1; i < rows.length; i++) {
    const catCell  = String(rows[i][0] || '').trim().toUpperCase();
    const testCell = String(rows[i][1] || '').trim().toUpperCase();

    if (catCell) {
      if (catCell === categoryName) {
        inTargetCategory = true;
      } else if (inTargetCategory) {
        break; // passed the category block without finding test
      }
    }

    if (inTargetCategory && testCell === testName) {
      foundRow = i + 1; // 1-based sheet row
      break;
    }
  }

  if (foundRow === -1) {
    return {
      success: false,
      message: 'Test "' + testName + '" not found in category "' + categoryName + '".',
    };
  }

  // Update columns C (regular) and D (senior)
  sheet.getRange(foundRow, 3, 1, 2).setValues([[priceRegular, priceSenior || priceRegular]]);
  sheet.getRange(foundRow, 3, 1, 2).setNumberFormat('#,##0.00');

  // Re-sync PRICELIST
  try { syncPriceList(); } catch(e) {
    Logger.log('syncPriceList after editTest failed (non-fatal): ' + e.message);
  }

  Logger.log('handleEditTest: Updated "' + testName + '" in "' + categoryName + '" at row ' + foundRow);
  return {
    success: true,
    message: 'Test "' + testName + '" updated successfully.',
    row: foundRow,
  };
}

// ═══════════════════════════════════════════════════════════════════
//  DELETE PATIENT
// ═══════════════════════════════════════════════════════════════════
function handleDeletePatient(data) {
  const numPadded  = String(data.numPadded  || '').trim();
  const nameCheck  = String(data.name       || '').trim().toLowerCase();
  if (!numPadded) return { success: false, message: 'numPadded is required.' };

  const ss = SpreadsheetApp.openById(RAW_SHEET_ID);

  function _deleteRowByNum(sheet, numCol, nameCol) {
    if (!sheet) return false;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;

    const range  = sheet.getRange(2, 1, lastRow - 1, Math.max(numCol, nameCol));
    const values = range.getValues();

    for (let i = values.length - 1; i >= 0; i--) {
      const cellNum  = String(values[i][numCol  - 1] || '').trim();
      const cellName = String(values[i][nameCol  - 1] || '').trim().toLowerCase();

      if (cellNum !== numPadded) continue;
      if (nameCheck && cellName && cellName !== nameCheck) continue;

      sheet.deleteRow(i + 2);
      return true;
    }
    return false;
  }

  const daily        = ss.getSheetByName(TAB.DAILY);
  const dailyDeleted = _deleteRowByNum(daily, COL.NO, COL.NAME);

  const logSheet   = ss.getSheetByName(TAB.RECEIPT_LOGS);
  const logDeleted = _deleteRowByNum(logSheet, RLOG.NO, RLOG.NAME);

  if (!dailyDeleted && !logDeleted) {
    return {
      success: false,
      message: 'Record #' + numPadded + ' not found in either sheet.',
    };
  }

  Logger.log('handleDeletePatient: Deleted #' + numPadded +
    ' | daily=' + dailyDeleted + ' | receiptLog=' + logDeleted);

  return {
    success: true,
    message: 'Record #' + numPadded + ' deleted.',
    dailyDeleted,
    logDeleted,
  };
}

// ═══════════════════════════════════════════════════════════════════
//  REFUND PATIENT
// ═══════════════════════════════════════════════════════════════════
function handleRefundPatient(data) {
  const numPadded = String(data.numPadded || '').trim();
  if (!numPadded) return { success: false, message: 'numPadded is required.' };

  const remarks = String(data.remarks || '').trim();
  const ss = SpreadsheetApp.openById(RAW_SHEET_ID);

  // ── 1. Update DAILY REG SHEET ────────────────────────────────────
  let dailyUpdated = false;
  const daily = ss.getSheetByName(TAB.DAILY);
  if (daily) {
    const lastRow = daily.getLastRow();
    if (lastRow >= 2) {
      const numVals = daily.getRange(2, COL.NO, lastRow - 1, 1).getValues();
      for (let i = numVals.length - 1; i >= 0; i--) {
        if (String(numVals[i][0] || '').trim() !== numPadded) continue;

        const sheetRow = i + 2;
        // Ensure the status validation allows REFUNDED before writing it.
        _setupDailyDropdowns(daily, sheetRow);

        // Set amounts to 0 and status to REFUNDED
        daily.getRange(sheetRow, COL.AMOUNT).setValue(0);        // AA
        daily.getRange(sheetRow, COL.STATUS).setValue('REFUNDED'); // AB
        daily.getRange(sheetRow, COL.PAYMENT).setValue(0);       // AD
        daily.getRange(sheetRow, COL.BALANCE).setValue(0);       // AE
        daily.getRange(sheetRow, COL.REMARKS).setValue(remarks);  // AF

        // Apply GRAY formatting to STATUS column
        daily.getRange(sheetRow, COL.STATUS).setFontColor('#757575')
                                            .setFontWeight('bold');
        
        dailyUpdated = true;
        Logger.log('handleRefundPatient: Refunded DAILY row ' + sheetRow + ' for #' + numPadded);
        break;
      }
    }
  }

  // ── 2. Update RECEIPT LOGS ───────────────────────────────────────
  let logUpdated = false;
  const logSheet = ss.getSheetByName(TAB.RECEIPT_LOGS);
  if (logSheet) {
    const lastRow = logSheet.getLastRow();
    if (lastRow >= 2) {
      const numVals = logSheet.getRange(2, RLOG.NO, lastRow - 1, 1).getValues();
      for (let i = numVals.length - 1; i >= 0; i--) {
        if (String(numVals[i][0] || '').trim() !== numPadded) continue;

        const sheetRow = i + 2;
        // Zero out amounts and set status to REFUNDED
        logSheet.getRange(sheetRow, RLOG.GRAND).setValue(0);      // W (Grand Total)
        logSheet.getRange(sheetRow, RLOG.STATUS).setValue('REFUNDED'); // X (Status)
        logSheet.getRange(sheetRow, RLOG.PAY_AMT).setValue(0);     // Z (Payment Amount)
        logSheet.getRange(sheetRow, RLOG.BALANCE).setValue(0);     // AA (Balance)

        // Apply GRAY formatting to STATUS column
        logSheet.getRange(sheetRow, RLOG.STATUS).setFontColor('#757575')
                                               .setFontWeight('bold');
        
        logUpdated = true;
        Logger.log('handleRefundPatient: Refunded RECEIPT LOGS row ' + sheetRow + ' for #' + numPadded);
        break;
      }
    }
  }

  if (!dailyUpdated && !logUpdated) {
    return {
      success: false,
      message: 'Record #' + numPadded + ' not found in either sheet.',
    };
  }

  return {
    success:      true,
    message:      'Record #' + numPadded + ' refunded successfully.',
    dailyUpdated,
    logUpdated,
  };
}

// ═══════════════════════════════════════════════════════════════════
//  EDIT PATIENT
// ═══════════════════════════════════════════════════════════════════
function handleEditPatient(data) {
  const numPadded = String(data.numPadded || '').trim();
  if (!numPadded) return { success: false, message: 'numPadded is required.' };

  const labSub      = parseFloat(data.labSubtotal)      || 0;
  const checkupSub  = parseFloat(data.checkupSubtotal)  || 0;
  const cadSub      = parseFloat(data.cadSubtotal)      || 0;
  const kitsSub     = parseFloat(data.kitsSubtotal)     || 0;
  const xraySub     = parseFloat(data.xraySubtotal)     || 0;
  const drugSub     = parseFloat(data.drugSubtotal)     || 0;
  const utzSub      = parseFloat(data.utzSubtotal)      || 0;
  const disAmt      = parseFloat(data.discountAmt)      || 0;
  const finalAmount = parseFloat(data.amount)
    || Math.max(0, labSub + checkupSub + cadSub + kitsSub +
                   xraySub + drugSub + utzSub - disAmt);
  const payStatus = (data.paymentStatus || 'PENDING').toString().trim().toUpperCase();
  let payType   = (data.paymentType   || 'CASH').toString().trim().toUpperCase().replace('G-CASH','GCASH');
  let payAmt    = parseFloat(data.paymentAmount) || 0;
  const balAmt    = parseFloat(data.balance)       || 0;

  // ── Split Payment Support ───────────────────────────────────────
  const cashAmount  = parseFloat(data.cashAmount)  || 0;
  const gcashAmount = parseFloat(data.gcashAmount) || 0;
  const gcashRef    = String(data.gcashReference || '').trim();

  // If both cash and GCash amounts are provided, it's a split payment
  if (cashAmount > 0 && gcashAmount > 0) {
    payType = 'CASH + GCASH';
    payAmt = cashAmount + gcashAmount;
  }

  const ss = SpreadsheetApp.openById(RAW_SHEET_ID);

  // ── 1. Update DAILY REG SHEET ────────────────────────────────────
  let dailyUpdated = false;
  const daily = ss.getSheetByName(TAB.DAILY);
  if (daily) {
    const lastRow = daily.getLastRow();
    if (lastRow >= 2) {
      const numVals = daily.getRange(2, COL.NO, lastRow - 1, 1).getValues();
      for (let i = numVals.length - 1; i >= 0; i--) {
        if (String(numVals[i][0] || '').trim() !== numPadded) continue;

        const sheetRow = i + 2;
        const rowData = [
          data.date           || daily.getRange(sheetRow, COL.DATE).getValue(),
          numPadded,
          data.name           || '',
          data.birthday       || '',
          data.ageSex         || '',
          data.address        || '',
          data.opType         || '',
          data.contact        || '',
          data.referred       || '',
          data.company        || '',
          data.procedures     || '',
          labSub    > 0 ? labSub    : '',
          data.checkupTests   || '',
          checkupSub > 0 ? checkupSub : '',
          data.cardiologyTests|| '',
          cadSub    > 0 ? cadSub    : '',
          data.kitsTests      || '',
          kitsSub   > 0 ? kitsSub   : '',
          data.xrayTests      || '',
          xraySub   > 0 ? xraySub   : '',
          data.drugTests      || '',
          drugSub   > 0 ? drugSub   : '',
          data.utzTests       || '',
          utzSub    > 0 ? utzSub    : '',
          data.discountLabel  || '',
          disAmt    > 0 ? disAmt    : '',
          finalAmount,
          payStatus,
          payType,
          payAmt,
          balAmt,
          data.remarks   || '',   
          data.encodedBy
        ];

        daily.getRange(sheetRow, 1, 1, rowData.length).setValues([rowData]);
        _formatDailyRow(daily, sheetRow, payStatus, payType, disAmt,
                        xraySub, drugSub, utzSub);
        _setupDailyDropdowns(daily, sheetRow);
        dailyUpdated = true;
        Logger.log('handleEditPatient: Updated DAILY row ' + sheetRow + ' for #' + numPadded);
        break;
      }
    }
  }

  // ── 2. Update RECEIPT LOGS ───────────────────────────────────────
  let logUpdated = false;
  const logSheet = ss.getSheetByName(TAB.RECEIPT_LOGS);
  if (logSheet) {
    const lastRow = logSheet.getLastRow();
    if (lastRow >= 2) {
      const numVals = logSheet.getRange(2, RLOG.NO, lastRow - 1, 1).getValues();
      for (let i = numVals.length - 1; i >= 0; i--) {
        if (String(numVals[i][0] || '').trim() !== numPadded) continue;

        const sheetRow = i + 2;
        const origTs = logSheet.getRange(sheetRow, RLOG.TIMESTAMP).getValue();
        const tsStr  = origTs instanceof Date
          ? origTs.getHours().toString().padStart(2,'0') + ':' +
            origTs.getMinutes().toString().padStart(2,'0') + ':' +
            origTs.getSeconds().toString().padStart(2,'0')
          : String(origTs || '');

        const rowData = [
          data.date           || logSheet.getRange(sheetRow, RLOG.DATE).getValue(),
          numPadded,
          data.name           || '',
          data.ageSex         || '',
          data.referred       || '',
          data.itemsJson      || '[]',
          data.procedures     || '',
          data.checkupTests   || '',
          data.cardiologyTests|| '',
          data.kitsTests      || '',
          data.xrayTests      || '',
          data.drugTests      || '',
          data.utzTests       || '',
          labSub    > 0 ? labSub    : 0,
          checkupSub > 0 ? checkupSub : 0,
          cadSub    > 0 ? cadSub    : 0,
          kitsSub   > 0 ? kitsSub   : 0,
          xraySub   > 0 ? xraySub   : 0,
          drugSub   > 0 ? drugSub   : 0,
          utzSub    > 0 ? utzSub    : 0,
          data.discountLabel  || '',
          disAmt,
          finalAmount,
          payStatus,
          payType,
          payAmt,
          balAmt,
          cashAmount || 0,
          gcashAmount || 0,
          gcashRef || '',
          tsStr,
          data.encodedBy || '',
        ];

        logSheet.getRange(sheetRow, 1, 1, 32).setValues([rowData]);
        _formatReceiptLogRow(logSheet, sheetRow, payStatus, payType, disAmt,
                             labSub, checkupSub, cadSub, kitsSub,
                             xraySub, drugSub, utzSub);
        logUpdated = true;
        Logger.log('handleEditPatient: Updated RECEIPT LOGS row ' + sheetRow + ' for #' + numPadded);
        break;
      }
    }
  }

  if (!dailyUpdated && !logUpdated) {
    return {
      success: false,
      message: 'Record #' + numPadded + ' not found in either sheet.',
    };
  }

  return {
    success:      true,
    message:      'Record #' + numPadded + ' updated successfully.',
    dailyUpdated,
    logUpdated,
  };
}

// ═══════════════════════════════════════════════════════════════════
//  REGISTER
// ═══════════════════════════════════════════════════════════════════
function handleRegister(data) {
  const ss    = SpreadsheetApp.openById(RAW_SHEET_ID);
  let   sheet = ss.getSheetByName(TAB.DAILY);
  if (!sheet) {
    sheet = ss.insertSheet(TAB.DAILY);
    _writeDailyHeader(sheet);
  }

  const lastRow = sheet.getLastRow();
  let nextNum   = 1;
  if (lastRow >= 2) {
    const nums = sheet.getRange(2, COL.NO, lastRow - 1, 1).getValues();
    for (let i = nums.length - 1; i >= 0; i--) {
      const raw = String(nums[i][0] || '').replace(/^0+/, '') || '0';
      const n   = parseInt(raw);
      if (!isNaN(n) && n > 0) { nextNum = n + 1; break; }
    }
  }
  const paddedNum = String(nextNum).padStart(4, '0');

  const labSubtotal      = parseFloat(data.labSubtotal)      || 0;
  const checkupSubtotal  = parseFloat(data.checkupSubtotal)  || 0;
  const cadSubtotal      = parseFloat(data.cadSubtotal)      || 0;
  const kitsSubtotal     = parseFloat(data.kitsSubtotal)     || 0;
  const xraySubtotal     = parseFloat(data.xraySubtotal)     || 0;
  const drugSubtotal     = parseFloat(data.drugSubtotal)     || 0;
  const utzSubtotal      = parseFloat(data.utzSubtotal)      || 0;
  const disAmt           = parseFloat(data.discountAmt)      || 0;
  const finalAmount      = parseFloat(data.amount)
    || Math.max(0, labSubtotal + checkupSubtotal + cadSubtotal + kitsSubtotal
                 + xraySubtotal + drugSubtotal + utzSubtotal - disAmt);

  const paymentStatus = data.paymentStatus || 'PENDING';
  let paymentType   = (data.paymentType   || 'CASH').toString().trim().toUpperCase().replace('G-CASH','GCASH');
  let paymentAmt    = parseFloat(data.paymentAmount) || 0;
  const balanceAmt    = parseFloat(data.balance)       || 0;

  // ── Split Payment Support ───────────────────────────────────────
  const cashAmount  = parseFloat(data.cashAmount)  || 0;
  const gcashAmount = parseFloat(data.gcashAmount) || 0;
  const gcashRef    = String(data.gcashReference || '').trim();

  // If both cash and GCash amounts are provided, it's a split payment
  if (cashAmount > 0 && gcashAmount > 0) {
    paymentType = 'CASH + GCASH';
    paymentAmt = cashAmount + gcashAmount;
  }

  const row = [
    data.date          || '',
    paddedNum,
    data.name          || '',
    data.birthday      || '',
    data.ageSex        || '',
    data.address       || '',
    data.opType        || '',
    data.contact       || '',
    data.referred      || '',
    data.company       || '',
    data.procedures    || '',
    labSubtotal  > 0 ? labSubtotal  : '',
    data.checkupTests  || '',
    checkupSubtotal > 0 ? checkupSubtotal : '',
    data.cardiologyTests || '',
    cadSubtotal  > 0 ? cadSubtotal  : '',
    data.kitsTests     || '',
    kitsSubtotal > 0 ? kitsSubtotal : '',
    data.xrayTests     || '',
    xraySubtotal > 0 ? xraySubtotal : '',
    data.drugTests     || '',
    drugSubtotal > 0 ? drugSubtotal : '',
    data.utzTests      || '',
    utzSubtotal  > 0 ? utzSubtotal  : '',
    data.discountLabel || '',
    disAmt       > 0 ? disAmt       : '',
    finalAmount,
    paymentStatus,
    paymentType,
    paymentAmt,
    balanceAmt,
    data.remarks       || '',
    data.encodedBy     || '',
  ];

  sheet.appendRow(row);
  const newRow = sheet.getLastRow();
  _formatDailyRow(sheet, newRow, paymentStatus, paymentType, disAmt,
                  xraySubtotal, drugSubtotal, utzSubtotal);
  _setupDailyDropdowns(sheet, newRow);

  try {
    _writeReceiptLog(ss, data, nextNum, paddedNum,
                     labSubtotal, checkupSubtotal, cadSubtotal, kitsSubtotal,
                     xraySubtotal, drugSubtotal, utzSubtotal,
                     disAmt, finalAmount, paymentStatus, paymentType,
                     paymentAmt, balanceAmt, cashAmount, gcashAmount, gcashRef);
  } catch(logErr) {
    Logger.log('_writeReceiptLog failed (non-fatal): ' + logErr.message);
  }

  return { success: true, number: nextNum };
}

// ═══════════════════════════════════════════════════════════════════
//  _writeReceiptLog
// ═══════════════════════════════════════════════════════════════════
function _writeReceiptLog(ss, data, num, paddedNum,
                          labSub, checkupSub, cadSub, kitsSub,
                          xraySub, drugSub, utzSub,
                          disAmt, finalAmount, payStatus, payType,
                          payAmt, balAmt, cashAmt, gcashAmt, gcashRef) {

  let logSheet = ss.getSheetByName(TAB.RECEIPT_LOGS);
  if (!logSheet) {
    logSheet = ss.insertSheet(TAB.RECEIPT_LOGS);
    _writeReceiptLogHeader(logSheet);
  }

  const ts = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'HH:mm:ss'
  );

  logSheet.appendRow([
    data.date               || '',
    paddedNum,
    data.name               || '',
    data.ageSex             || '',
    data.referred           || '',
    data.itemsJson          || '[]',
    data.procedures         || '',
    data.checkupTests       || '',
    data.cardiologyTests    || '',
    data.kitsTests          || '',
    data.xrayTests          || '',
    data.drugTests          || '',
    data.utzTests           || '',
    labSub     > 0 ? labSub    : 0,
    checkupSub > 0 ? checkupSub: 0,
    cadSub     > 0 ? cadSub    : 0,
    kitsSub    > 0 ? kitsSub   : 0,
    xraySub    > 0 ? xraySub   : 0,
    drugSub    > 0 ? drugSub   : 0,
    utzSub     > 0 ? utzSub    : 0,
    data.discountLabel      || '',
    disAmt,
    finalAmount,
    payStatus,
    payType,
    payAmt,
    balAmt,
    cashAmt || 0,
    gcashAmt || 0,
    gcashRef || '',
    ts,
    data.encodedBy || '',
  ]);

  const newRow = logSheet.getLastRow();
  _formatReceiptLogRow(logSheet, newRow, payStatus, payType, disAmt,
                       labSub, checkupSub, cadSub, kitsSub,
                       xraySub, drugSub, utzSub);
}

// ═══════════════════════════════════════════════════════════════════
//  _writeReceiptLogHeader
// ═══════════════════════════════════════════════════════════════════
function _writeReceiptLogHeader(sheet) {
  const n = RLOG_HEADERS.length;

  const hdr = sheet.getRange(1, 1, 1, n);
  hdr.setValues([RLOG_HEADERS])
     .setFontFamily('Arial')
     .setFontWeight('bold')
     .setFontSize(10)
     .setFontColor('#FFFFFF')
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle')
     .setBackground('#283593');

  sheet.getRange(1, RLOG.DATE,        1, 5 ).setBackground('#4A0E8F');
  sheet.getRange(1, RLOG.ITEMS_JSON,  1, 1 ).setBackground('#37474F');
  sheet.getRange(1, RLOG.LAB_PROC,    1, 1 ).setBackground('#6F2DBD');
  sheet.getRange(1, RLOG.CHECKUP,     1, 1 ).setBackground('#7B1FA2');
  sheet.getRange(1, RLOG.CARDIOLOGY,  1, 1 ).setBackground('#AD1457');
  sheet.getRange(1, RLOG.KITS,        1, 1 ).setBackground('#00838F');
  sheet.getRange(1, RLOG.XRAY,        1, 1 ).setBackground('#1565C0');
  sheet.getRange(1, RLOG.DRUG,        1, 1 ).setBackground('#00695C');
  sheet.getRange(1, RLOG.UTZ,         1, 1 ).setBackground('#283593');
  sheet.getRange(1, RLOG.LAB_SUB,     1, 7 ).setBackground('#4A0E8F');
  sheet.getRange(1, RLOG.DISC_LABEL,  1, 2 ).setBackground('#37474F');
  sheet.getRange(1, RLOG.GRAND,       1, 1 ).setBackground('#1B5E20');
  sheet.getRange(1, RLOG.STATUS,      1, 1 ).setBackground('#6A1B9A');
  sheet.getRange(1, RLOG.PAY_TYPE,    1, 3 ).setBackground('#283593');
  sheet.getRange(1, RLOG.TIMESTAMP,   1, 1 ).setBackground('#37474F');

  sheet.setRowHeight(1, 38);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  sheet.setColumnWidth(RLOG.DATE,       90);
  sheet.setColumnWidth(RLOG.NO,         55);
  sheet.setColumnWidth(RLOG.NAME,       180);
  sheet.setColumnWidth(RLOG.AGE_SEX,    75);
  sheet.setColumnWidth(RLOG.REFERRED,   140);
  sheet.setColumnWidth(RLOG.ITEMS_JSON, 400);
  sheet.setColumnWidth(RLOG.LAB_PROC,   260);
  sheet.setColumnWidth(RLOG.CHECKUP,    160);
  sheet.setColumnWidth(RLOG.CARDIOLOGY, 160);
  sheet.setColumnWidth(RLOG.KITS,       160);
  sheet.setColumnWidth(RLOG.XRAY,       160);
  sheet.setColumnWidth(RLOG.DRUG,       130);
  sheet.setColumnWidth(RLOG.UTZ,        160);
  for (let c = RLOG.LAB_SUB; c <= RLOG.UTZ_SUB; c++) {
    sheet.setColumnWidth(c, 105);
  }
  sheet.setColumnWidth(RLOG.DISC_LABEL, 180);
  sheet.setColumnWidth(RLOG.DISC_AMT,   95);
  sheet.setColumnWidth(RLOG.GRAND,      105);
  sheet.setColumnWidth(RLOG.STATUS,     105);
  sheet.setColumnWidth(RLOG.PAY_TYPE,   90);
  sheet.setColumnWidth(RLOG.PAY_AMT,    95);
  sheet.setColumnWidth(RLOG.BALANCE,    95);
  sheet.setColumnWidth(RLOG.TIMESTAMP,  80);
  sheet.setColumnWidth(RLOG.ENCODED_BY, 130);
}

// ═══════════════════════════════════════════════════════════════════
//  _formatReceiptLogRow
// ═══════════════════════════════════════════════════════════════════
function _formatReceiptLogRow(sheet, row, payStatus, payType, disAmt,
                               labSub, checkupSub, cadSub, kitsSub,
                               xraySub, drugSub, utzSub) {

  const n        = RLOG_HEADERS.length;
  const rowRange = sheet.getRange(row, 1, 1, n);

  rowRange
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true,
               '#DDDDDD', SpreadsheetApp.BorderStyle.SOLID);

  rowRange.setBackground(row % 2 === 0 ? '#F8F6FF' : '#FFFFFF');

  [RLOG.ITEMS_JSON, RLOG.LAB_PROC, RLOG.CHECKUP,
   RLOG.CARDIOLOGY, RLOG.KITS, RLOG.XRAY, RLOG.DRUG, RLOG.UTZ]
    .forEach(c => sheet.getRange(row, c).setWrap(true).setVerticalAlignment('top'));

  sheet.getRange(row, RLOG.NO)
       .setHorizontalAlignment('center')
       .setFontFamily('DM Mono, Courier New, monospace')
       .setFontWeight('bold');

  RLOG_PESO_COLS.forEach(c =>
    sheet.getRange(row, c).setNumberFormat('₱#,##0.00')
  );

  [
    { c: RLOG.LAB_SUB,    v: labSub     },
    { c: RLOG.CHECKUP_SUB,v: checkupSub },
    { c: RLOG.CAD_SUB,    v: cadSub     },
    { c: RLOG.KITS_SUB,   v: kitsSub    },
    { c: RLOG.XRAY_SUB,   v: xraySub    },
    { c: RLOG.DRUG_SUB,   v: drugSub    },
    { c: RLOG.UTZ_SUB,    v: utzSub     },
  ].forEach(({ c, v }) => {
    const cell = sheet.getRange(row, c);
    if (v > 0) {
      cell.setFontWeight('bold').setFontColor('#1a1a1a');
    } else {
      cell.setFontColor('#BBBBBB').setFontStyle('italic');
    }
  });

  if (checkupSub > 0)
    sheet.getRange(row, RLOG.CHECKUP_SUB)
         .setBackground('#F3E5F5').setFontColor('#7B1FA2').setFontWeight('bold');
  if (cadSub > 0)
    sheet.getRange(row, RLOG.CAD_SUB)
         .setBackground('#FCE4EC').setFontColor('#AD1457').setFontWeight('bold');
  if (kitsSub > 0)
    sheet.getRange(row, RLOG.KITS_SUB)
         .setBackground('#E0F7FA').setFontColor('#00838F').setFontWeight('bold');
  if (xraySub > 0)
    sheet.getRange(row, RLOG.XRAY_SUB)
         .setBackground('#E3F2FD').setFontColor('#1565C0').setFontWeight('bold');
  if (drugSub > 0)
    sheet.getRange(row, RLOG.DRUG_SUB)
         .setBackground('#E0F2F1').setFontColor('#00695C').setFontWeight('bold');
  if (utzSub > 0)
    sheet.getRange(row, RLOG.UTZ_SUB)
         .setBackground('#E8EAF6').setFontColor('#283593').setFontWeight('bold');

  if (disAmt > 0) {
    sheet.getRange(row, RLOG.DISC_AMT)
         .setBackground('#FFF3E0').setFontColor('#E65100').setFontWeight('bold');
  } else {
    sheet.getRange(row, RLOG.DISC_AMT)
         .setFontColor('#BBBBBB').setFontStyle('italic');
    sheet.getRange(row, RLOG.DISC_LABEL)
         .setFontColor('#BBBBBB').setFontStyle('italic');
  }

  sheet.getRange(row, RLOG.GRAND)
       .setBackground('#E8F5E9').setFontColor('#1B5E20').setFontWeight('bold');

  const ss = STATUS_STYLES[payStatus] || STATUS_STYLES['PENDING'];
  sheet.getRange(row, RLOG.STATUS)
       .setBackground(ss.bg).setFontColor(ss.fg)
       .setFontWeight('bold').setHorizontalAlignment('center');

  const ps = PMETHOD_STYLES[payType] || PMETHOD_STYLES['CASH'];
  sheet.getRange(row, RLOG.PAY_TYPE)
       .setBackground(ps.bg).setFontColor(ps.fg)
       .setFontWeight('bold').setHorizontalAlignment('center');

  const balCell = sheet.getRange(row, RLOG.BALANCE);
  balCell.setFontWeight('bold');
  const balVal = balCell.getValue();
  if (typeof balVal === 'number' && balVal > 0) {
    balCell.setFontColor('#B71C1C');
  }

  sheet.getRange(row, RLOG.TIMESTAMP)
       .setFontFamily('DM Mono, Courier New, monospace')
       .setFontColor('#888888')
       .setHorizontalAlignment('center');
}

// ═══════════════════════════════════════════════════════════════════
//  rebuildReceiptLogFormatting
// ═══════════════════════════════════════════════════════════════════
function rebuildReceiptLogFormatting() {
  const ss       = SpreadsheetApp.openById(RAW_SHEET_ID);
  const logSheet = ss.getSheetByName(TAB.RECEIPT_LOGS);

  if (!logSheet) {
    Logger.log('rebuildReceiptLogFormatting: RECEIPT LOGS tab not found — nothing to do.');
    return;
  }

  const lastRow = logSheet.getLastRow();
  _writeReceiptLogHeader(logSheet);

  if (lastRow < 2) {
    Logger.log('rebuildReceiptLogFormatting: Header written. No data rows yet.');
    return;
  }

  const data = logSheet.getRange(2, 1, lastRow - 1, RLOG_HEADERS.length).getValues();

  data.forEach((r, i) => {
    const row      = i + 2;
    const payStatus= String(r[RLOG.STATUS   - 1] || 'PENDING').trim().toUpperCase();
    const payType  = String(r[RLOG.PAY_TYPE - 1] || 'CASH'   ).trim().toUpperCase();
    const disAmt   = parseFloat(r[RLOG.DISC_AMT  - 1]) || 0;
    const labSub   = parseFloat(r[RLOG.LAB_SUB   - 1]) || 0;
    const ckupSub  = parseFloat(r[RLOG.CHECKUP_SUB-1]) || 0;
    const cadSub   = parseFloat(r[RLOG.CAD_SUB   - 1]) || 0;
    const kitsSub  = parseFloat(r[RLOG.KITS_SUB  - 1]) || 0;
    const xraySub  = parseFloat(r[RLOG.XRAY_SUB  - 1]) || 0;
    const drugSub  = parseFloat(r[RLOG.DRUG_SUB  - 1]) || 0;
    const utzSub   = parseFloat(r[RLOG.UTZ_SUB   - 1]) || 0;

    _formatReceiptLogRow(logSheet, row, payStatus, payType, disAmt,
                         labSub, ckupSub, cadSub, kitsSub,
                         xraySub, drugSub, utzSub);
  });

  Logger.log('rebuildReceiptLogFormatting: Formatted ' + (lastRow - 1) + ' rows.');

  try {
    SpreadsheetApp.getUi().alert(
      '[OK] RECEIPT LOGS reformatted!\n\n' +
      (lastRow - 1) + ' data row(s) updated.'
    );
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════
//  handleGetToday
// ═══════════════════════════════════════════════════════════════════
function handleGetToday() {
  const ss       = SpreadsheetApp.openById(RAW_SHEET_ID);
  const logSheet = ss.getSheetByName(TAB.RECEIPT_LOGS);

  if (!logSheet) return { success: true, records: [] };

  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return { success: true, records: [] };

  const now      = new Date();
  const todayStr = (now.getMonth() + 1).toString().padStart(2, '0') + '/' +
                   now.getDate().toString().padStart(2, '0') + '/' +
                   now.getFullYear().toString();

  const rows = logSheet.getRange(2, 1, lastRow - 1, 29).getValues();
  const records = [];

  rows.forEach(r => {
    if (String(r[0] || '').trim() !== todayStr) return;

    let items = [];
    try { items = JSON.parse(String(r[5] || '[]')); } catch(e) { items = []; }

    records.push({
      date: (function(v) {
        if (v instanceof Date)
          return (v.getMonth()+1).toString().padStart(2,'0') + '/' +
                 v.getDate().toString().padStart(2,'0') + '/' +
                 v.getFullYear();
        return String(v||'').replace(/\s*(GMT|UTC|\().*$/, '').trim();
      })(r[0]),
      num:             parseInt(String(r[1]).replace(/^0+/, '') || '0'),
      name:            String(r[2]  || ''),
      ageSex:          String(r[3]  || ''),
      referred:        String(r[4]  || ''),
      items,
      procedures:      String(r[6]  || ''),
      checkupTests:    String(r[7]  || ''),
      cardiologyTests: String(r[8]  || ''),
      kitsTests:       String(r[9]  || ''),
      xrayTests:       String(r[10] || ''),
      drugTests:       String(r[11] || ''),
      utzTests:        String(r[12] || ''),
      labSubtotal:     parseFloat(r[13]) || 0,
      checkupSubtotal: parseFloat(r[14]) || 0,
      cadSubtotal:     parseFloat(r[15]) || 0,
      kitsSubtotal:    parseFloat(r[16]) || 0,
      xraySubtotal:    parseFloat(r[17]) || 0,
      drugSubtotal:    parseFloat(r[18]) || 0,
      utzSubtotal:     parseFloat(r[19]) || 0,
      discountLabel:   String(r[20] || ''),
      discAmt:         parseFloat(r[21]) || 0,
      grand:           parseFloat(r[22]) || 0,
      paymentStatus:   String(r[23] || 'PENDING'),
      paymentType:     String(r[24] || 'CASH'),
      paymentAmt:      parseFloat(r[25]) || 0,
      balance:         parseFloat(r[26]) || 0,
      ts:              String(r[27] || ''),
      encodedBy:       String(r[28] || ''),
    });
  });

  return { success: true, records };
}

// ═══════════════════════════════════════════════════════════════════
//  handleGetPatients
// ═══════════════════════════════════════════════════════════════════
function handleGetPatients(p) {
  const ss    = SpreadsheetApp.openById(RAW_SHEET_ID);
  const sheet = ss.getSheetByName(TAB.DAILY);
  if (!sheet) return { success: true, records: [], total: 0, page: 1, pages: 1 };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, records: [], total: 0, page: 1, pages: 1 };

  const PAGE_SIZE = 100;
  const page      = Math.max(1, parseInt(p.page) || 1);
  const search    = (p.search    || '').trim().toLowerCase();
  const status    = (p.status    || '').trim().toUpperCase();
  const referred  = (p.referred  || '').trim().toLowerCase();
  const dateFrom  = (p.dateFrom  || '').trim();
  const dateTo    = (p.dateTo    || '').trim();

  const rows = sheet.getRange(2, 1, lastRow - 1, 33).getValues();

  // Build lookup map from RECEIPT LOGS
  const logMap = {};
  try {
    const logSheet = ss.getSheetByName(TAB.RECEIPT_LOGS);
    if (logSheet && logSheet.getLastRow() >= 2) {
      const logRows = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 29).getValues();
      logRows.forEach(lr => {
        const key = String(lr[RLOG.NO - 1] || '').trim();
        if (!key) return;

        let ts = lr[RLOG.TIMESTAMP - 1];
        if (ts instanceof Date) {
          ts = ts.getHours().toString().padStart(2,'0') + ':' +
               ts.getMinutes().toString().padStart(2,'0') + ':' +
               ts.getSeconds().toString().padStart(2,'0');
        } else {
          ts = String(ts || '').trim();
          const m = ts.match(/(\d{1,2}:\d{2}:\d{2})/);
          ts = m ? m[1] : ts;
        }

        logMap[key] = {
          ts,
          itemsJson:       String(lr[RLOG.ITEMS_JSON    - 1] || '[]'),
          procedures:      String(lr[RLOG.LAB_PROC      - 1] || ''),
          checkupTests:    String(lr[RLOG.CHECKUP       - 1] || ''),
          cardiologyTests: String(lr[RLOG.CARDIOLOGY    - 1] || ''),
          kitsTests:       String(lr[RLOG.KITS          - 1] || ''),
          xrayTests:       String(lr[RLOG.XRAY          - 1] || ''),
          drugTests:       String(lr[RLOG.DRUG          - 1] || ''),
          utzTests:        String(lr[RLOG.UTZ           - 1] || ''),
          labSubtotal:     parseFloat(lr[RLOG.LAB_SUB      - 1]) || 0,
          checkupSubtotal: parseFloat(lr[RLOG.CHECKUP_SUB  - 1]) || 0,
          cadSubtotal:     parseFloat(lr[RLOG.CAD_SUB      - 1]) || 0,
          kitsSubtotal:    parseFloat(lr[RLOG.KITS_SUB     - 1]) || 0,
          xraySubtotal:    parseFloat(lr[RLOG.XRAY_SUB     - 1]) || 0,
          drugSubtotal:    parseFloat(lr[RLOG.DRUG_SUB     - 1]) || 0,
          utzSubtotal:     parseFloat(lr[RLOG.UTZ_SUB      - 1]) || 0,
          discountLabel:   String(lr[RLOG.DISC_LABEL    - 1] || ''),
          discAmt:         parseFloat(lr[RLOG.DISC_AMT     - 1]) || 0,
          encodedBy:       String(lr[RLOG.ENCODED_BY    - 1] || ''),
        };
      });
    }
  } catch(e) { /* non-fatal */ }

  const all = [];
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    const num  = String(r[COL.NO   - 1] || '').trim();
    const name = String(r[COL.NAME - 1] || '').trim();
    if (!num && !name) continue;

    let dateStr = '';
    const rawDate = r[COL.DATE - 1];
    if (rawDate instanceof Date) {
      dateStr = (rawDate.getMonth()+1).toString().padStart(2,'0') + '/' +
                rawDate.getDate().toString().padStart(2,'0') + '/' +
                rawDate.getFullYear().toString();
    } else if (rawDate) {
      const s = String(rawDate).trim();
      const match = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (match) {
        dateStr = match[0];
      } else {
        const d = new Date(s);
        if (!isNaN(d)) {
          dateStr = (d.getMonth()+1).toString().padStart(2,'0') + '/' +
                    d.getDate().toString().padStart(2,'0') + '/' +
                    d.getFullYear().toString();
        } else {
          dateStr = s.replace(/GMT.*/,'').replace(/\(.*\)/,'').trim();
        }
      }
    }

    const payStatus = String(r[COL.STATUS    - 1] || 'PENDING').trim().toUpperCase();
    const payType   = String(r[COL.P_METHODS - 1] || 'CASH'   ).trim();
    const payAmt    = parseFloat(r[COL.PAYMENT    - 1]) || 0;
    const balance   = parseFloat(r[COL.BALANCE    - 1]) || 0;
    const grand     = parseFloat(r[COL.AMOUNT     - 1]) || 0;
    const ref       = String(r[COL.REFERRED  - 1] || '').trim();
    const ageSex    = String(r[COL.AGE_SEX   - 1] || '').trim();
    const birthday  = String(r[COL.BIRTHDAY  - 1] || '').trim();
    const address   = String(r[COL.ADDRESS   - 1] || '').trim();
    const opType    = String(r[COL.OP_TYPE   - 1] || '').trim();
    const contact   = String(r[COL.CONTACT   - 1] || '').trim();
    const company   = String(r[COL.COMPANY   - 1] || '').trim();

    if (search && !name.toLowerCase().includes(search) &&
        !num.includes(search)) continue;
    if (status && payStatus !== status) continue;
    if (referred && !ref.toLowerCase().includes(referred)) continue;
    if (dateFrom || dateTo) {
      const parts = dateStr.split('/');
      if (parts.length === 3) {
        const d = new Date(parts[2], parts[0]-1, parts[1]);
        if (dateFrom) {
          const fp = dateFrom.split('-');
          if (d < new Date(fp[0], fp[1]-1, fp[2])) continue;
        }
        if (dateTo) {
          const tp = dateTo.split('-');
          if (d > new Date(tp[0], tp[1]-1, tp[2])) continue;
        }
      }
    }

    const log = logMap[num] || {};
    let items = [];
    try { items = JSON.parse(log.itemsJson || '[]'); } catch(e) { items = []; }

    all.push({
      date:            dateStr,
      num:             parseInt(String(num).replace(/^0+/,'')) || 0,
      numPadded:       num,
      name,
      ageSex,
      referred:        ref,
      ts:              log.ts || '',
      birthday,
      address,
      opType,
      contact,
      company,
      procedures:      log.procedures      || String(r[COL.PROCEDURES - 1] || '').trim(),
      checkupTests:    log.checkupTests    || String(r[COL.CHECKUP    - 1] || '').trim(),
      cardiologyTests: log.cardiologyTests || String(r[COL.CARDIOLOGY - 1] || '').trim(),
      kitsTests:       log.kitsTests       || String(r[COL.KITS       - 1] || '').trim(),
      xrayTests:       log.xrayTests       || String(r[COL.XRAY_TESTS - 1] || '').trim(),
      drugTests:       log.drugTests       || String(r[COL.DRUG_TEST  - 1] || '').trim(),
      utzTests:        log.utzTests        || String(r[COL.UTZ_TESTS  - 1] || '').trim(),
      labSubtotal:     log.labSubtotal      || parseFloat(r[COL.LAB_SUB     - 1]) || 0,
      checkupSubtotal: log.checkupSubtotal  || parseFloat(r[COL.CHECKUP_SUB - 1]) || 0,
      cadSubtotal:     log.cadSubtotal      || parseFloat(r[COL.CAD_SUB     - 1]) || 0,
      kitsSubtotal:    log.kitsSubtotal     || parseFloat(r[COL.KITS_SUB    - 1]) || 0,
      xraySubtotal:    log.xraySubtotal     || parseFloat(r[COL.XRAY_SUB    - 1]) || 0,
      drugSubtotal:    log.drugSubtotal     || parseFloat(r[COL.DRUG_SUB    - 1]) || 0,
      utzSubtotal:     log.utzSubtotal      || parseFloat(r[COL.UTZ_SUB     - 1]) || 0,
      discountLabel:   log.discountLabel   || String(r[COL.DISCOUNT  - 1] || '').trim(),
      discAmt:         log.discAmt         || parseFloat(r[COL.DIS_TOTAL - 1]) || 0,
      grand,
      paymentStatus:   payStatus,
      paymentType:     payType,
      paymentAmt:      payAmt,
      balance,
      items,
      itemsJson:       log.itemsJson || '[]',
      fromSheet: true,
      encodedBy: log.encodedBy || String(r[COL.ENCODED_BY - 1] || '').trim(),
    });
  }

  const total = all.length;
  const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  const start = (page - 1) * PAGE_SIZE;
  const slice = all.slice(start, start + PAGE_SIZE);

  return { success: true, records: slice, total, page, pages };
}

// ═══════════════════════════════════════════════════════════════════
//  _writeDailyHeader
// ═══════════════════════════════════════════════════════════════════
function _writeDailyHeader(sheet) {
  const n   = DAILY_HEADERS.length;
  const hdr = sheet.getRange(1, 1, 1, n);

  hdr.setValues([DAILY_HEADERS])
     .setFontFamily('Arial')
     .setFontWeight('bold')
     .setFontSize(10)
     .setFontColor('#FFFFFF')
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');

  hdr.setBackground('#4A0E8F');
  sheet.getRange(1, COL.PROCEDURES, 1, 2).setBackground('#6F2DBD');
  sheet.getRange(1, COL.CHECKUP,    1, 2).setBackground('#7B1FA2');
  sheet.getRange(1, COL.CARDIOLOGY, 1, 2).setBackground('#AD1457');
  sheet.getRange(1, COL.KITS,       1, 2).setBackground('#00838F');
  sheet.getRange(1, COL.XRAY_TESTS, 1, 2).setBackground('#1565C0');
  sheet.getRange(1, COL.DRUG_TEST,  1, 2).setBackground('#00695C');
  sheet.getRange(1, COL.UTZ_TESTS,  1, 2).setBackground('#283593');
  sheet.getRange(1, COL.DISCOUNT,   1, 7).setBackground('#37474F');

  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  DAILY_COL_WIDTHS.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

// ═══════════════════════════════════════════════════════════════════
//  _formatDailyRow
// ═══════════════════════════════════════════════════════════════════
function _formatDailyRow(sheet, row, status, payMethod, disAmt,
                          xraySub, drugSub, utzSub) {
  const n        = DAILY_HEADERS.length;
  const rowRange = sheet.getRange(row, 1, 1, n);

  rowRange.setFontFamily('Arial')
          .setFontSize(10)
          .setVerticalAlignment('middle')
          .setBorder(true, true, true, true, true, true,
                     '#DDDDDD', SpreadsheetApp.BorderStyle.SOLID);

  rowRange.setBackground(row % 2 === 0 ? '#F8F6FF' : '#FFFFFF');

  [COL.PROCEDURES, COL.CHECKUP, COL.CARDIOLOGY, COL.KITS,
   COL.XRAY_TESTS, COL.DRUG_TEST, COL.UTZ_TESTS].forEach(c =>
    sheet.getRange(row, c).setWrap(true).setVerticalAlignment('top')
  );

  PESO_COLS.forEach(c => sheet.getRange(row, c).setNumberFormat('₱#,##0.00'));

  sheet.getRange(row, COL.NO).setHorizontalAlignment('center')
       .setFontFamily('DM Mono, monospace').setFontWeight('bold');

  if (xraySub > 0)
    sheet.getRange(row, COL.XRAY_SUB)
         .setBackground('#E3F2FD').setFontColor('#1565C0').setFontWeight('bold');
  if (drugSub > 0)
    sheet.getRange(row, COL.DRUG_SUB)
         .setBackground('#E0F2F1').setFontColor('#00695C').setFontWeight('bold');
  if (utzSub > 0)
    sheet.getRange(row, COL.UTZ_SUB)
         .setBackground('#E8EAF6').setFontColor('#283593').setFontWeight('bold');

  if (!disAmt || disAmt === 0)
    sheet.getRange(row, COL.DIS_TOTAL).setFontColor('#BBBBBB').setFontStyle('italic');

  sheet.getRange(row, COL.AMOUNT) .setFontWeight('bold');
  sheet.getRange(row, COL.BALANCE).setFontWeight('bold');

  const ss = STATUS_STYLES[status] || STATUS_STYLES['PENDING'];
  sheet.getRange(row, COL.STATUS)
       .setBackground(ss.bg).setFontColor(ss.fg)
       .setFontWeight('bold').setHorizontalAlignment('center');

  const ps = PMETHOD_STYLES[payMethod] || PMETHOD_STYLES['CASH'];
  sheet.getRange(row, COL.P_METHODS)
       .setBackground(ps.bg).setFontColor(ps.fg)
       .setFontWeight('bold').setHorizontalAlignment('center');
}

// ═══════════════════════════════════════════════════════════════════
//  _setupDailyDropdowns
// ═══════════════════════════════════════════════════════════════════
function _setupDailyDropdowns(sheet, row) {
  sheet.getRange(row, COL.STATUS).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['PENDING','FULLY PAID','PARTLY PAID','UNPAID','REFUNDED'], true)
      .setAllowInvalid(false).build()
  );
  sheet.getRange(row, COL.P_METHODS).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['CASH','GCASH','YAKAP','BSA','BANK'], true)
      .setAllowInvalid(false).build()
  );
}

// ═══════════════════════════════════════════════════════════════════
//  onEdit TRIGGER
// ═══════════════════════════════════════════════════════════════════
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TAB.DAILY) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  const get  = c => sheet.getRange(row, c).getValue();
  const set  = (c, v) => sheet.getRange(row, c).setValue(v);
  const setF = (c, f) => sheet.getRange(row, c).setFormula(f);
  const clr  = c => sheet.getRange(row, c).clearContent();

  const amtRef     = _col(COL.AMOUNT)  + row;
  const pmtRef     = _col(COL.PAYMENT) + row;
  const balFormula = '=' + amtRef + '-' + pmtRef;

  function applyStatusRules(status) {
    status = (status || '').toString().trim().toUpperCase();
    const amount = parseFloat(get(COL.AMOUNT)) || 0;

    if (status === 'FULLY PAID') {
      set(COL.PAYMENT, amount);
      set(COL.BALANCE, 0);
    } else if (status === 'PARTLY PAID') {
      clr(COL.PAYMENT);
      setF(COL.BALANCE, balFormula);
      const pm = (get(COL.P_METHODS) || '').toString().trim().toUpperCase();
      if (pm === 'YAKAP' || pm === 'BSA') {
        set(COL.P_METHODS, 'CASH');
        _applyPMethodStyle(sheet, row, 'CASH');
      }
    } else if (status === 'UNPAID') {
      set(COL.PAYMENT, 0);
      set(COL.BALANCE, amount);
    } else if (status === 'PENDING') {
      set(COL.PAYMENT, 0);
      setF(COL.BALANCE, balFormula);
    } else if (status === 'REFUNDED') {
      [COL.LAB_SUB, COL.CHECKUP_SUB, COL.CAD_SUB, COL.KITS_SUB,
       COL.XRAY_SUB, COL.DRUG_SUB, COL.UTZ_SUB,
       COL.DIS_TOTAL, COL.AMOUNT, COL.PAYMENT, COL.BALANCE]
        .forEach(c => set(c, 0));
      clr(COL.P_METHODS);
      sheet.getRange(row, COL.P_METHODS)
           .setBackground('#EEEEEE').setFontColor('#888888')
           .setFontWeight('normal').setHorizontalAlignment('center');
    } else {
      set(COL.PAYMENT, 0);
      setF(COL.BALANCE, balFormula);
    }
    _applyStatusStyle(sheet, row, status);
  }

  if (col === COL.P_METHODS) {
    const pm = (e.value || '').toString().trim().toUpperCase();
    _applyPMethodStyle(sheet, row, pm);
    if (pm === 'YAKAP' || pm === 'BSA') {
      set(COL.STATUS, 'FULLY PAID');
      applyStatusRules('FULLY PAID');
    } else {
      const currentStatus = (get(COL.STATUS) || '').toString().trim();
      if (currentStatus) applyStatusRules(currentStatus);
    }
    return;
  }

  if (col === COL.STATUS) {
    applyStatusRules((e.value || '').toString().trim());
    return;
  }

  if (col === COL.PAYMENT) {
    const status  = (get(COL.STATUS) || '').toString().trim().toUpperCase();
    const amount  = parseFloat(get(COL.AMOUNT)) || 0;
    const payment = parseFloat(e.value) || 0;
    if (status === 'FULLY PAID') {
      set(COL.BALANCE, Math.max(0, amount - payment));
    } else {
      setF(COL.BALANCE, balFormula);
    }
    return;
  }
}

// ═══════════════════════════════════════════════════════════════════
//  _applyStatusStyle / _applyPMethodStyle
// ═══════════════════════════════════════════════════════════════════
function _applyStatusStyle(sheet, row, status) {
  const s = STATUS_STYLES[status] || STATUS_STYLES['PENDING'];
  sheet.getRange(row, COL.STATUS)
       .setBackground(s.bg).setFontColor(s.fg)
       .setFontWeight('bold').setHorizontalAlignment('center');
}

function _applyPMethodStyle(sheet, row, method) {
  const s = PMETHOD_STYLES[method] || PMETHOD_STYLES['CASH'];
  sheet.getRange(row, COL.P_METHODS)
       .setBackground(s.bg).setFontColor(s.fg)
       .setFontWeight('bold').setHorizontalAlignment('center');
}

// ═══════════════════════════════════════════════════════════════════
//  MONTHLY RESET
// ═══════════════════════════════════════════════════════════════════
function monthlyReset() {
  const ss    = SpreadsheetApp.openById(RAW_SHEET_ID);
  const sheet = ss.getSheetByName(TAB.DAILY);

  if (!sheet) {
    Logger.log('monthlyReset: DAILY REG SHEET not found.');
    return;
  }

  const lastRow = sheet.getLastRow();

  const now         = new Date();
  const archiveDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const archiveName = archiveDate.toLocaleDateString('en-US', {
    month: 'long', year: 'numeric',
  }).toUpperCase();

  if (lastRow >= 2) {
    const oldArchive = ss.getSheetByName(archiveName);
    if (oldArchive) ss.deleteSheet(oldArchive);

    const archiveSheet = ss.insertSheet(archiveName);
    const numCols      = DAILY_HEADERS.length;
    const allData      = sheet.getRange(1, 1, lastRow, numCols).getValues();

    archiveSheet.getRange(1, 1, lastRow, numCols).setValues(allData);
    _writeDailyHeader(archiveSheet);

    for (let r = 2; r <= lastRow; r++) {
      const rd = allData[r - 1];
      if (!rd[1] && !rd[2]) continue;
      const status  = String(rd[COL.STATUS    - 1] || 'PENDING').trim().toUpperCase();
      const pm      = String(rd[COL.P_METHODS - 1] || 'CASH').trim().toUpperCase();
      const dis     = parseFloat(rd[COL.DIS_TOTAL - 1]) || 0;
      const xraySub = parseFloat(rd[COL.XRAY_SUB  - 1]) || 0;
      const drugSub = parseFloat(rd[COL.DRUG_SUB  - 1]) || 0;
      const utzSub  = parseFloat(rd[COL.UTZ_SUB   - 1]) || 0;
      _formatDailyRow(archiveSheet, r, status, pm, dis, xraySub, drugSub, utzSub);
    }

    DAILY_COL_WIDTHS.forEach((w, i) => archiveSheet.setColumnWidth(i + 1, w));
    archiveSheet.setFrozenRows(1);
    archiveSheet.setFrozenColumns(3);

    ss.setActiveSheet(archiveSheet);
    ss.moveActiveSheet(ss.getNumSheets());
  }

  if (lastRow >= 2) sheet.deleteRows(2, lastRow - 1);
  _writeDailyHeader(sheet);

  _archiveReceiptLogs(ss, archiveName);
}

// ═══════════════════════════════════════════════════════════════════
//  _archiveReceiptLogs
// ═══════════════════════════════════════════════════════════════════
function _archiveReceiptLogs(ss, archiveName) {
  const logSheet = ss.getSheetByName(TAB.RECEIPT_LOGS);
  if (!logSheet) return;

  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return;

  const archiveTabName = 'RECEIPTS ' + archiveName;
  const oldArchive     = ss.getSheetByName(archiveTabName);
  if (oldArchive) ss.deleteSheet(oldArchive);

  const archiveSheet = ss.insertSheet(archiveTabName);
  const numCols      = RLOG_HEADERS.length;
  const allData      = logSheet.getRange(1, 1, lastRow, numCols).getValues();

  archiveSheet.getRange(1, 1, lastRow, numCols).setValues(allData);

  _writeReceiptLogHeader(archiveSheet);
  const dataRows = allData.slice(1);
  dataRows.forEach((r, i) => {
    const row      = i + 2;
    const payStatus= String(r[RLOG.STATUS    - 1] || 'PENDING').trim().toUpperCase();
    const payType  = String(r[RLOG.PAY_TYPE  - 1] || 'CASH'   ).trim().toUpperCase();
    const disAmt   = parseFloat(r[RLOG.DISC_AMT   - 1]) || 0;
    const labSub   = parseFloat(r[RLOG.LAB_SUB    - 1]) || 0;
    const ckupSub  = parseFloat(r[RLOG.CHECKUP_SUB- 1]) || 0;
    const cadSub   = parseFloat(r[RLOG.CAD_SUB    - 1]) || 0;
    const kitsSub  = parseFloat(r[RLOG.KITS_SUB   - 1]) || 0;
    const xraySub  = parseFloat(r[RLOG.XRAY_SUB   - 1]) || 0;
    const drugSub  = parseFloat(r[RLOG.DRUG_SUB   - 1]) || 0;
    const utzSub   = parseFloat(r[RLOG.UTZ_SUB    - 1]) || 0;
    _formatReceiptLogRow(archiveSheet, row, payStatus, payType, disAmt,
                         labSub, ckupSub, cadSub, kitsSub,
                         xraySub, drugSub, utzSub);
  });

  ss.setActiveSheet(archiveSheet);
  ss.moveActiveSheet(ss.getNumSheets());

  if (lastRow >= 2) logSheet.deleteRows(2, lastRow - 1);
}

// ═══════════════════════════════════════════════════════════════════
//  setupDailyRegSheet — RUN ONCE
// ═══════════════════════════════════════════════════════════════════
function setupDailyRegSheet() {
  const ss       = SpreadsheetApp.openById(RAW_SHEET_ID);
  const existing = ss.getSheetByName(TAB.DAILY);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(TAB.DAILY);
  _writeDailyHeader(sheet);

  SpreadsheetApp.getUi().alert(
    '[OK] DAILY REG SHEET created!\n\n' +
    'Now run createMonthlyResetTrigger() ONCE\n' +
    'to enable the 1AM auto-archive on the 1st of every month.'
  );
}

// ═══════════════════════════════════════════════════════════════════
//  createMonthlyResetTrigger — RUN ONCE ONLY
// ═══════════════════════════════════════════════════════════════════
function createMonthlyResetTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'monthlyReset') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('monthlyReset')
    .timeBased()
    .onMonthDay(1)
    .atHour(1)
    .create();

  try {
    SpreadsheetApp.getUi().alert(
      '[OK] Trigger installed!\n\n' +
      'monthlyReset() will run automatically at 1:00 AM\n' +
      'on the 1st of every month.\n\n' +
      'Verify: Apps Script → Triggers (clock icon on left).'
    );
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════
//  manualMonthlyReset
// ═══════════════════════════════════════════════════════════════════
function manualMonthlyReset() {
  const ui     = SpreadsheetApp.getUi();
  const result = ui.alert(
    '[!] Manual Monthly Reset',
    'This will archive all current DAILY REG SHEET data and reset it.\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );
  if (result === ui.Button.YES) {
    monthlyReset();
    ui.alert('[OK] Done! Sheet archived and reset.');
  }
}

// ═══════════════════════════════════════════════════════════════════
//  syncPriceList
// ═══════════════════════════════════════════════════════════════════
function syncPriceList() {
  const ss       = SpreadsheetApp.openById(SHEET_ID);
  const srcSheet = ss.getSheetByName(TAB.TESTS);
  if (!srcSheet) {
    SpreadsheetApp.getUi().alert('[X] "' + TAB.TESTS + '" tab not found.');
    return;
  }

  const srcData = srcSheet.getDataRange().getValues();
  const output  = [['TEST NAME', 'REGULAR', 'SENIOR']];

  for (let i = 1; i < srcData.length; i++) {
    const cat  = String(srcData[i][0] || '').trim();
    const test = String(srcData[i][1] || '').trim();
    const reg  = srcData[i][2];
    const sen  = srcData[i][3];
    if (cat || !test) continue;
    const r = typeof reg === 'number' && reg > 0 ? reg : null;
    const s = typeof sen === 'number' && sen > 0 ? sen : null;
    if (r === null && s === null) continue;
    output.push([test.toUpperCase(), r || '', s || '']);
  }

  let plSheet = ss.getSheetByName(TAB.PRICELIST);
  if (plSheet) ss.deleteSheet(plSheet);
  plSheet = ss.insertSheet(TAB.PRICELIST);

  plSheet.getRange(1, 1, output.length, 3).setValues(output);
  plSheet.getRange(1, 1, 1, 3)
         .setBackground('#6F2DBD').setFontColor('#FFFFFF')
         .setFontWeight('bold').setFontSize(11);
  plSheet.getRange(1, 1, 1, 3).setNote(
    '[!] Do not edit manually. Edit TESTS & PRICES, then run syncPriceList().'
  );
  if (output.length > 1)
    plSheet.getRange(2, 2, output.length - 1, 2).setNumberFormat('#,##0.00');
  plSheet.setColumnWidth(1, 340);
  plSheet.setColumnWidth(2, 120);
  plSheet.setColumnWidth(3, 120);
  plSheet.setFrozenRows(1);

  try {
    SpreadsheetApp.getUi().alert('[OK] PRICELIST synced — ' + (output.length - 1) + ' tests.');
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════
//  _createDoctorsTab
// ═══════════════════════════════════════════════════════════════════
function _createDoctorsTab(ss) {
  const sheet = ss.insertSheet(TAB.DOCTORS);
  sheet.getRange(1, 1, 1, 3)
       .setValues([['NAME', 'DETAILS', 'DATE']])
       .setBackground('#6F2DBD').setFontColor('#FFFFFF')
       .setFontWeight('bold').setFontSize(11);
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 150);
  sheet.setFrozenRows(1);
  return sheet;
}

// ═══════════════════════════════════════════════════════════════════
//  _col — column number → letter(s)
// ═══════════════════════════════════════════════════════════════════
function _col(n) {
  let l = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    l = String.fromCharCode(65 + r) + l;
    n = Math.floor((n - 1) / 26);
  }
  return l;
}

// ═══════════════════════════════════════════════════════════════════
//  initialSetup — RUN ONCE
// ═══════════════════════════════════════════════════════════════════
function initialSetup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let logins = ss.getSheetByName(TAB.LOGINS);
  if (!logins) {
    logins = ss.insertSheet(TAB.LOGINS);
    logins.getRange(1, 1, 1, 3)
          .setValues([['USERNAME', 'PASSWORD', 'ROLE']])
          .setBackground('#6F2DBD').setFontColor('#FFFFFF').setFontWeight('bold');
    logins.getRange(2, 1, 2, 3)
          .setValues([['admin', 'admin2026', 'Admin'], ['labtech', 'lab@lis2026', 'Staff']]);
    logins.setColumnWidth(1, 160); logins.setColumnWidth(2, 160); logins.setColumnWidth(3, 120);
    logins.setFrozenRows(1);
  }

  if (!ss.getSheetByName(TAB.DOCTORS)) {
    const d = _createDoctorsTab(ss);
    d.getRange(2, 1, 4, 3).setValues([
      ['DR. JUAN DELA CRUZ', 'Internal Medicine — City Medical Center', ''],
      ['DR. MARIA SANTOS',   'OB-GYN — San Pedro General Hospital',    ''],
      ['WALK-IN',            'No referral',                             ''],
      ['YAKAP PROGRAM',      'Community health program — 100% free',    ''],
    ]);
  }

  syncPriceList();

  if (!ss.getSheetByName(TAB.DISCOUNTS)) {
    const disc = ss.insertSheet(TAB.DISCOUNTS);
    disc.getRange(1, 1, 1, 3)
        .setValues([['DISCOUNT NAME', 'RATE (%)', 'NOTES']])
        .setBackground('#6F2DBD').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
    const discData = [
      ['Senior/PWD', 20, ''], ['YAKAP Program', 100, ''],
      ['BSA (Brgy. San Antonio)', 100, ''], ['Mika Discount (Regular)', 15, ''],
      ['Mika Discount (Senior)', 25, ''], ['PWD', 20, ''],
      ['Promo', 20, ''], ['Special Discount', 15, ''],
      ['Charged to Company', 100, ''], ['Charged to OM', 100, ''],
    ];
    disc.getRange(2, 1, discData.length, 3).setValues(discData);
    disc.setColumnWidth(1, 220); disc.setColumnWidth(2, 100); disc.setColumnWidth(3, 240);
    disc.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert(
    '[OK] Setup complete!\n\nNext:\n' +
    '1. Run setupDailyRegSheet()\n' +
    '2. Run createMonthlyResetTrigger()\n' +
    '3. Deploy as Web App'
  );
}

// ═══════════════════════════════════════════════════════════════════
//  normalizeGcash — RUN ONCE after deploying v7/v8
// ═══════════════════════════════════════════════════════════════════
function normalizeGcash() {
  const ss = SpreadsheetApp.openById(RAW_SHEET_ID);
  let dailyFixed = 0;
  let logFixed   = 0;

  const daily = ss.getSheetByName(TAB.DAILY);
  if (daily) {
    const lastRow = daily.getLastRow();
    if (lastRow >= 2) {
      const pmVals = daily.getRange(2, COL.P_METHODS, lastRow - 1, 1).getValues();
      pmVals.forEach((r, i) => {
        const raw = String(r[0] || '').trim();
        if (raw.toUpperCase() === 'G-CASH') {
          const row = i + 2;
          daily.getRange(row, COL.P_METHODS).setValue('GCASH');
          _applyPMethodStyle(daily, row, 'GCASH');
          _setupDailyDropdowns(daily, row);
          dailyFixed++;
        }
      });
    }
  }

  const logSheet = ss.getSheetByName(TAB.RECEIPT_LOGS);
  if (logSheet) {
    const lastRow = logSheet.getLastRow();
    if (lastRow >= 2) {
      const ptVals = logSheet.getRange(2, RLOG.PAY_TYPE, lastRow - 1, 1).getValues();
      ptVals.forEach((r, i) => {
        const raw = String(r[0] || '').trim();
        if (raw.toUpperCase() === 'G-CASH') {
          const row = i + 2;
          logSheet.getRange(row, RLOG.PAY_TYPE).setValue('GCASH');
          const ps = PMETHOD_STYLES['GCASH'] || PMETHOD_STYLES['CASH'];
          logSheet.getRange(row, RLOG.PAY_TYPE)
            .setBackground(ps.bg).setFontColor(ps.fg)
            .setFontWeight('bold').setHorizontalAlignment('center');
          logFixed++;
        }
      });
    }
  }

  Logger.log('normalizeGcash: dailyFixed=' + dailyFixed + ', logFixed=' + logFixed);

  try {
    SpreadsheetApp.getUi().alert(
      '[OK] GCASH normalisation complete!\n\n' +
      'DAILY REG SHEET rows fixed: ' + dailyFixed + '\n' +
      'RECEIPT LOGS rows fixed:    ' + logFixed + '\n\n' +
      '"G-CASH" no longer exists in any row.\n' +
      'All payment method dropdowns now only accept GCASH.'
    );
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════
//  rebuildDailyFormatting
// ═══════════════════════════════════════════════════════════════════
function rebuildDailyFormatting() {
  const ss    = SpreadsheetApp.openById(RAW_SHEET_ID);
  const sheet = ss.getSheetByName(TAB.DAILY);

  if (!sheet) {
    Logger.log('rebuildDailyFormatting: DAILY REG SHEET not found.');
    return;
  }

  _writeDailyHeader(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('rebuildDailyFormatting: No data rows.');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 31).getValues();
  data.forEach((r, i) => {
    const row      = i + 2;
    const status   = String(r[COL.STATUS    - 1] || 'PENDING').trim().toUpperCase();
    const payMethod= String(r[COL.P_METHODS - 1] || 'CASH'   ).trim().toUpperCase()
                       .replace('G-CASH','GCASH');
    const disAmt   = parseFloat(r[COL.DIS_TOTAL  - 1]) || 0;
    const xraySub  = parseFloat(r[COL.XRAY_SUB   - 1]) || 0;
    const drugSub  = parseFloat(r[COL.DRUG_SUB   - 1]) || 0;
    const utzSub   = parseFloat(r[COL.UTZ_SUB    - 1]) || 0;
    _formatDailyRow(sheet, row, status, payMethod, disAmt, xraySub, drugSub, utzSub);
    _setupDailyDropdowns(sheet, row);
  });

  Logger.log('rebuildDailyFormatting: Formatted ' + (lastRow - 1) + ' rows.');

  try {
    SpreadsheetApp.getUi().alert(
      '[OK] DAILY REG SHEET reformatted!\n\n' +
      (lastRow - 1) + ' data row(s) updated.\n' +
      'All "G-CASH" values also normalized to "GCASH".'
    );
  } catch(e) {}
}