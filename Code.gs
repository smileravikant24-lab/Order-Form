var FIRM_SHEET_URL  = 'https://docs.google.com/spreadsheets/d/1WLgaZ21Jt4uckn1zyi01F5yLcwtWTmUTpmwQGW7hQGI/edit';
var FIRM_SHEET_NAME = 'WMS_Firms';
var FIRM_COLUMN     = 1;
var GST_COLUMN      = 2;

var PARTY_MASTER_URL       = 'https://docs.google.com/spreadsheets/d/1R1eMFIwOdN5CyfrZa0K8NSVaJZnVmjPd95bOp2FYz90/edit';
var PARTY_MASTER_SHEET     = 'PartyMaster';

// Product master lives as a second tab in the same master spreadsheet as
// the firm master. New products confirmed by the user (see
// addNewProductToMaster) get appended here.
var PRODUCT_SHEET_NAME = 'WMS_Products';
var DEFAULT_PACKING_RATIO_SERVER = 5;

// Seeds the product master with the products already known to the form,
// so deploying this doesn't suddenly treat all of them as "new". Ream Per
// Box values match the existing packingMap in index.html.
var DEFAULT_PRODUCT_SEED = [
  ['DOUBLE A A4 70 GSM', 5], ['DOUBLE A A4 75 GSM', 5], ['DOUBLE A A4 80 GSM', 5], ['DOUBLE A A4 100 GSM', 4],
  ['N R COPIER A4 70 GSM', 10], ['N R COPIER A4 75 GSM', 10], ['N R COPIER A4 80 GSM', 10], ['N R COPIER A3 70 GSM', 5], ['N R COPIER A3 75 GSM', 5],
  ['KHANNA A4 KBOLT (S) 65 GSM', 10], ['KHANNA A4 EPRINT (S) 70 GSM', 10], ['KHANNA A4 ECOPY (S) 75 GSM', 10], ['KHANNA A3 KBOLT (S) 65 GSM', 5],
  ['SATIA A4 65 GSM', 10], ['SATIA A4 70 GSM', 10], ['SATIA A4 75 GSM', 10], ['SATIA A3 70 GSM', 5], ['SATIA FS 70 GSM', 10],
  ['ROZAANA A4 70 GSM', 10], ['ROZAANA A4 75 GSM', 10],
  ['I K COPIER A4 70 GSM', 10], ['P P LITE A4 70 GSM', 10],
  ['PAPER ONE A4 75 GSM', 5], ['PAPER ONE A4 100 GSM', 4], ['PAPER ONE A3 100 GSM', 4],
  ['MAPLE A4 70 GSM', 10], ['SPECTRA A4 75 GSM', 10], ['CENTURY STAR A4 75 GSM', 10], ['J K RED A4 75 GSM', 10],
  ['REFLECTION A4 65 GSM', 10], ['REFLECTION A4 70 GSM', 10], ['REFLECTION A4 75 GSM', 10], ['REFLECTION A4 100 GSM', 5],
  ['REFLECTION FS 70 GSM', 10], ['REFLECTION A3 70 GSM', 5],
  ['FINEPRINT A4 70 GSM', 10], ['FINEPRINT A4 75 GSM', 10],
  ['Tarang Yellow A4 75 GSM', 5], ['Tarang Pink A4 75 GSM', 5], ['Tarang Green A4 75 GSM', 5], ['Tarang Blue A4 75 GSM', 5], ['Tarang Ivory A4 75 GSM', 5]
];

var SALES_PERSONS = [
  'MR. PAWAN NEGI',
  'MR. PRANAV SATIJA',
  'MR. MUKESH SHUKLA',
  'MR. RISHABH JAIN',
  'MR. SONU CHAUHAN'
];

var WHATSAPP_CONFIG = {
  INSTANCE: 'instance171119',
  TOKEN:    '9tqcp752y7at6z2j',
  PHONE:    '919899708020'
};

// PIN required to modify an existing order (search-by-PI-No. feature).
// Change this to whatever PIN you want the team to use.
var MODIFY_ORDER_PIN = '1234';


function doGet(e) {
  if (e.parameter.action === 'getInitialData') {
    var firms = getFirmData();
    var salesPersons = getSalesPersons();
    var products = getProductData();
    return ContentService.createTextOutput(JSON.stringify({ firms: firms, salesPersons: salesPersons, products: products }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (e.parameter.action === 'addNewFirm') {
    var firmResult = addNewFirmToMaster(e.parameter.name, e.parameter.gst);
    return ContentService.createTextOutput(JSON.stringify(firmResult))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (e.parameter.action === 'addNewProduct') {
    var productResult = addNewProductToMaster(e.parameter.name, e.parameter.perBox);
    return ContentService.createTextOutput(JSON.stringify(productResult))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (e.parameter.action === 'getFormNumber') {
    var fno = getNextFormNumber(e.parameter.formType);
    return ContentService.createTextOutput(JSON.stringify({ formNumber: fno }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (e.parameter.action === 'getOrderByFormNumber') {
    var result = getOrderByFormNumber(e.parameter.formNumber);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (e.parameter.action === 'verifyModifyPin') {
    var isValid = (e.parameter.pin || '').toString().trim() === MODIFY_ORDER_PIN;
    return ContentService.createTextOutput(JSON.stringify({ valid: isValid }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Satija Paper — Order / Enquiry')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Scans the sheet once to find the highest serial already used for this
// form-type + financial-year combo. Only called the very first time a
// combo is needed after deployment (or a new FY starts) — after that,
// getNextFormNumber() reads/writes a cached counter instead, which is why
// the PI No. used to take a while to show up (it re-scanned the whole
// Orders sheet on every single page load / new entry) and is now instant.
function scanMaxSerialFromSheet(formType, fy) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var isOrder = (formType === 'Order');
  var baseType = isOrder ? 'PI-SP' : 'Enq-SP';
  var sheetName = isOrder ? 'Orders' : 'Enquiries';
  var sheet = ss.getSheetByName(sheetName);
  var maxSerial = 0;

  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      // NOTE: Form Numbers live in Column B. If that ever changes,
      // update idColumn accordingly.
      var idColumn = 2;
      var allFormNos = sheet.getRange(2, idColumn, lastRow - 1, 1).getValues();

      for (var i = 0; i < allFormNos.length; i++) {
        var cellVal = (allFormNos[i][0] || '').toString().trim();
        if (cellVal.includes(baseType) && cellVal.includes(fy)) {
          var match = cellVal.match(/\d+$/);
          if (match) {
            var serialNum = parseInt(match[0], 10);
            if (!isNaN(serialNum) && serialNum > maxSerial) {
              maxSerial = serialNum;
            }
          }
        }
      }
    }
  }
  return maxSerial;
}

function getSerialCacheKey(formType, fy) {
  var isOrder = (formType === 'Order');
  // "V2" so any stale/inflated counter left over from an earlier buggy
  // version is ignored and re-derived from the actual sheet instead of
  // being trusted as-is.
  return 'SERIAL_V2_' + (isOrder ? 'Order' : 'Enquiry') + '_' + fy;
}

function getCachedMaxSerial(formType, fy) {
  var props = PropertiesService.getScriptProperties();
  var cacheKey = getSerialCacheKey(formType, fy);
  var cached = parseInt(props.getProperty(cacheKey), 10);
  if (isNaN(cached)) {
    cached = scanMaxSerialFromSheet(formType, fy);
    props.setProperty(cacheKey, String(cached));
  }
  return cached;
}

// Called only after an order/enquiry is actually written to the sheet, so
// the cache tracks reality. This is the ONLY place the counter moves
// forward — merely previewing a number (page load, tab switch, refresh)
// must never consume one.
function bumpSerialCache(formType, formNumber) {
  try {
    var match = (formNumber || '').toString().match(/(\d+)$/);
    if (!match) return;
    var serial = parseInt(match[1], 10);
    if (isNaN(serial)) return;

    var fy = getCurrentFY();
    var props = PropertiesService.getScriptProperties();
    var cacheKey = getSerialCacheKey(formType, fy);
    var current = parseInt(props.getProperty(cacheKey), 10);
    if (isNaN(current) || serial > current) {
      props.setProperty(cacheKey, String(serial));
    }
  } catch (e) {
    Logger.log('bumpSerialCache error: ' + e.message);
  }
}

// PEEK ONLY — this is what the form shows and what refreshes on every page
// load / tab switch. It must be side-effect-free: it never reserves or
// consumes a number, it just reports "next one would be X" based on the
// cached max (falling back to a one-time sheet scan if nothing is cached
// yet). The number only actually gets used once bumpSerialCache() runs,
// which happens after a real save — see processOrder / processEnquiry.
function getNextFormNumber(formType) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    var fy = getCurrentFY();
    var isOrder = (formType === 'Order');
    var prefix = isOrder ? ('PI-SP/' + fy + '/') : ('Enq-SP/' + fy + '/');

    var maxSerial = getCachedMaxSerial(formType, fy);

    // Generate the next number and pad it to 5 digits (e.g., 00078)
    return prefix + String(maxSerial + 1).padStart(5, '0');

  } catch (e) {
    console.error("Error generating form number: " + e.message);
    throw new Error("System busy, please try again.");
  } finally {
    lock.releaseLock();
  }
}

function getCurrentFY() {
  var now = new Date();
  var startYear = now.getFullYear();
  if (now.getMonth() < 3) startYear--;
  return String(startYear).slice(-2) + '-' + String(startYear + 1).slice(-2);
}

var FIRM_CACHE_KEY = 'FIRM_DATA_CACHE_V1';
var FIRM_CACHE_SECONDS = 300; // 5 min

function getFirmData() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(FIRM_CACHE_KEY);
    if (cached) return JSON.parse(cached);

    var externalSS = SpreadsheetApp.openByUrl(FIRM_SHEET_URL);
    var sheet = externalSS.getSheetByName(FIRM_SHEET_NAME) || externalSS.getSheets()[0];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var colsToRead = Math.max(FIRM_COLUMN, GST_COLUMN);
    var values = sheet.getRange(2, 1, lastRow - 1, colsToRead).getValues();
    var result = [];
    for (var i = 0; i < values.length; i++) {
      var firmName = (values[i][FIRM_COLUMN - 1] || '').toString().trim().toUpperCase();
      var gstNo    = (values[i][GST_COLUMN - 1]  || '').toString().trim().toUpperCase();
      if (firmName) result.push({ name: firmName, gst: gstNo });
    }
    try { cache.put(FIRM_CACHE_KEY, JSON.stringify(result), FIRM_CACHE_SECONDS); } catch (cacheErr) { Logger.log('Firm cache put error: ' + cacheErr.message); }
    return result;
  } catch (e) {
    Logger.log('getFirmData Error: ' + e.message);
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var localSheet = ss.getSheetByName('WMS_Firms');
      if (!localSheet) return [];
      var lr = localSheet.getLastRow();
      if (lr < 2) return [];
      var vals = localSheet.getRange(2, 1, lr - 1, 2).getValues();
      var res = [];
      for (var j = 0; j < vals.length; j++) {
        var fn = (vals[j][0] || '').toString().trim().toUpperCase();
        var gn = (vals[j][1] || '').toString().trim().toUpperCase();
        if (fn) res.push({ name: fn, gst: gn });
      }
      return res;
    } catch(e2) { return []; }
  }
}

function getFirmNames() {
  return getFirmData().map(function(d){ return d.name; });
}

function getProductSheet_() {
  var externalSS = SpreadsheetApp.openByUrl(FIRM_SHEET_URL);
  var sheet = externalSS.getSheetByName(PRODUCT_SHEET_NAME);
  if (!sheet) {
    sheet = externalSS.insertSheet(PRODUCT_SHEET_NAME);
    sheet.appendRow(['Product Name', 'Ream Per Box']);
    sheet.setFrozenRows(1);
    if (DEFAULT_PRODUCT_SEED.length) {
      sheet.getRange(2, 1, DEFAULT_PRODUCT_SEED.length, 2).setValues(DEFAULT_PRODUCT_SEED);
    }
  }
  return sheet;
}

var PRODUCT_CACHE_KEY = 'PRODUCT_DATA_CACHE_V1';
var PRODUCT_CACHE_SECONDS = 300; // 5 min

function getProductData() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(PRODUCT_CACHE_KEY);
    if (cached) return JSON.parse(cached);

    var sheet = getProductSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    var result = [];
    for (var i = 0; i < values.length; i++) {
      var name = (values[i][0] || '').toString().trim();
      var perBox = parseFloat(values[i][1]) || DEFAULT_PACKING_RATIO_SERVER;
      if (name) result.push({ name: name, perBox: perBox });
    }
    try { cache.put(PRODUCT_CACHE_KEY, JSON.stringify(result), PRODUCT_CACHE_SECONDS); } catch (cacheErr) { Logger.log('Product cache put error: ' + cacheErr.message); }
    return result;
  } catch (e) {
    Logger.log('getProductData Error: ' + e.message);
    return [];
  }
}

// Saves a newly added party only to the PartyMaster sheet.
// WMS_Firms is read-only from this app; it is managed separately.
function addNewFirmToMaster(name, gst) {
  name = (name || '').toString().trim().toUpperCase();
  gst = (gst || '').toString().trim().toUpperCase();
  if (!name) return { success: false, message: 'Firm name required.' };

  try {
    var pmSS = SpreadsheetApp.openByUrl(PARTY_MASTER_URL);
    var pmSheet = pmSS.getSheetByName(PARTY_MASTER_SHEET);
    if (!pmSheet) return { success: false, message: 'PartyMaster sheet not found.' };
    pmSheet.appendRow([name, gst, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')]);
    return { success: true };
  } catch (e) {
    Logger.log('addNewFirmToMaster Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Adds a product the user has confirmed is new to the WMS_Products master
// sheet. Silently no-ops if the name is already present (case-insensitive).
function addNewProductToMaster(name, perBox) {
  name = (name || '').toString().trim();
  perBox = parseFloat(perBox) || DEFAULT_PACKING_RATIO_SERVER;
  if (!name) return { success: false, message: 'Product name required.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = getProductSheet_();
    var lastRow = sheet.getLastRow();
    var nameUpper = name.toUpperCase();

    if (lastRow >= 2) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if ((existing[i][0] || '').toString().trim().toUpperCase() === nameUpper) {
          return { success: true, alreadyExists: true };
        }
      }
    }

    sheet.appendRow([name, perBox]);
    try { CacheService.getScriptCache().remove(PRODUCT_CACHE_KEY); } catch (cacheErr) {}
    return { success: true };
  } catch (e) {
    Logger.log('addNewProductToMaster Error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getSalesPersons() {
  return SALES_PERSONS;
}

// ✅ NEW FUNCTION: To force ALL CAPS everywhere safely
function toAllCaps(str) {
  if (!str) return '';
  return str.toString().trim().toUpperCase();
}

function cleanFirmSlug(firmName) {
  return (firmName || 'PARTY')
    .replace(/[^a-zA-Z0-9 ]/g, '')
    .trim()
    .replace(/\s+/g, '_')
    .substring(0, 30)
    .toUpperCase();
}

function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    var hdrRange = sheet.getRange(1, 1, 1, headers.length);
    hdrRange.setFontWeight('bold');
    hdrRange.setBackground('#1a5c2a');
    hdrRange.setFontColor('#ffffff');
    hdrRange.setBorder(true, true, true, true, true, true);
  } else {
    // Append any missing columns at the end (e.g. newly added 'Filled By (Name)')
    var lastCol = sheet.getLastColumn();
    for (var c = lastCol; c < headers.length; c++) {
      var cell = sheet.getRange(1, c + 1);
      cell.setValue(headers[c]);
      cell.setFontWeight('bold');
      cell.setBackground('#1a5c2a');
      cell.setFontColor('#ffffff');
      cell.setBorder(true, true, true, true, true, true);
    }
  }
  return sheet;
}

var _ss = null;
var _tz = null;
function getSS(){ if(!_ss) _ss = SpreadsheetApp.getActiveSpreadsheet(); return _ss; }
function getTZ(){ if(!_tz) _tz = getSS().getSpreadsheetTimeZone(); return _tz; }

function applyRowFormat(sheet, row, numCols) {
  sheet.getRange(row, 1, 1, numCols).setBorder(true, true, true, true, true, true);
}

function applyBatchRowFormat(sheet, startRow, numRows, numCols) {
  if(numRows < 1) return;
  sheet.getRange(startRow, 1, numRows, numCols).setBorder(true, true, true, true, true, true);
}

function processForm(formData) {
  try {
    _ss = null; _tz = null;
    var timestamp = new Date();

    if (formData.customerName) {
      var n = formData.customerName.trim();
      if (!n.toLowerCase().startsWith('mr.')) n = 'MR. ' + n;
      formData.customerName = n.toUpperCase();
    }
    if (formData.formType === 'Enquiry' && formData.referenceName) {
      var r = formData.referenceName.trim();
      if (!r.toLowerCase().startsWith('mr.')) r = 'MR. ' + r;
      formData.referenceName = r.toUpperCase();
    }

    var typeCheck = (formData.formType || '').toString().trim().toLowerCase();

    if (typeCheck === 'order') {
      if (formData.isUpdate) {
        return updateOrder(formData, timestamp);
      }
      return processOrder(formData, timestamp);
    }

    return processEnquiry(formData, timestamp);

  } catch (e) {
    Logger.log('processForm Error: ' + e.message);
    return { success: false, message: 'Server error: ' + e.message };
  }
}


var ORDER_DETAIL_HEADERS = [
  'Timestamp', 'Form No', 'Firm Name', 'GST No', 'Area', 'Customer Name',
  'WhatsApp', 'Sales Person', 'Order Date', 'Delivery Date',
  'Client Type', 'Payment Terms', 'Payment Days',
  'Transport Mode', 'Transport Name', 'Transporter Number', 'Vehicle Number', 'Driver Number',
  'Product Name', 'Ream Qty', 'Box Qty', 'Original Rate', 'Discount', 'Net Rate', 'Freight (Item)',
  'Item Value', 'Item Value + Freight', 'GST (18%)', 'Item Grand Total',
  'Remarks', 'Filled By (Name)', 'PDF URL'
];

var ORDER_SUMMARY_HEADERS = [
  'Timestamp', 'Form No', 'Firm Name', 'GST No', 'Area', 'Customer Name',
  'WhatsApp', 'Sales Person', 'Order Date', 'Delivery Date',
  'Client Type', 'Payment Terms', 'Payment Days',
  'Transport Mode', 'Transport Name', 'Transporter Number', 'Vehicle Number', 'Driver Number',
  'Items (Name | Ream | Box | Rate)', 'Total Reams', 'Total Boxes',
  'Gross Value', 'Total Discount', 'Net Value',
  'Freight', 'GST (18%)', 'Grand Total',
  'Remarks', 'Filled By (Name)', 'PDF URL'
];

// Builds the detail + summary rows for an order without touching the sheet.
// Shared by processOrder (new order) and updateOrder (modify existing order)
// so the two flows can never drift out of sync with each other.
function buildOrderRowsAndSummary(formData, timestamp) {
  var items = formData.items || [];

  var fmtDate = function(d) {
    if (!d) return '';
    try { return Utilities.formatDate(new Date(d), getTZ(), 'dd/MM/yyyy'); }
    catch(e) { return d; }
  };

  var totalFreight = parseFloat(formData.freightAmount) || 0;
  var round2 = function(v){ return Math.round(v * 100) / 100; };

  var totalReams = 0;
  items.forEach(function(it){ totalReams += (parseFloat(it.ream) || 0); });

  var fmFirmName   = toAllCaps(formData.firmName);
  var fmGstNo      = toAllCaps(formData.gstNo);
  var fmArea       = toAllCaps(formData.areaName);
  var fmCustomer   = toAllCaps(formData.customerName);
  var fmEmail      = toAllCaps(formData.filledByEmail || '');
  var fmOrderDate  = fmtDate(formData.orderDate);
  var fmDelivDate  = fmtDate(formData.deliveryDate);
  var fmClientType = toAllCaps(formData.clientType || '');
  var fmSalesPerson = toAllCaps(formData.salesPerson || '');
  var fmPaymentTerms = toAllCaps(formData.paymentTerms || '');
  var fmTransportMode = toAllCaps(formData.transportMode || '');
  var fmTransName  = toAllCaps(formData.transportName || '');
  var fmVehicle    = toAllCaps(formData.vehicleNumber);
  var fmRemarks    = toAllCaps(formData.remarks);

  var detailRows = [];
  var totalReamsSummary = 0, totalBoxesSummary = 0;
  var grossVal = 0, totalDiscSummary = 0, netVal = 0;
  var itemsStrParts = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var isFirst = (i === 0);
    var itemReams    = parseFloat(item.ream) || 0;
    var itemRate     = parseFloat(item.rate) || 0;
    var itemOrigRate = parseFloat(item.origRate) || itemRate;
    var itemDiscount = parseFloat(item.discount) || 0;
    var itemValue    = round2(itemReams * itemRate);
    var itemFreight  = totalReams > 0
      ? round2(totalFreight * (itemReams / totalReams))
      : (i === 0 ? totalFreight : 0);
    var itemWithFreight = round2(itemValue + itemFreight);
    var itemGST         = round2(itemWithFreight * 0.18);
    var itemGrand       = round2(itemWithFreight + itemGST);

    var fmItemName   = toAllCaps(item.name);

    totalReamsSummary += itemReams;
    totalBoxesSummary += (parseFloat(item.box) || 0);
    grossVal          += round2(itemReams * itemOrigRate);
    totalDiscSummary  += itemDiscount;
    netVal            += itemValue;

    itemsStrParts.push(
      fmItemName + ' | ' + itemReams + ' REAM | ' +
      (parseFloat(item.box)||0) + ' BOX | RS.' + round2(itemOrigRate)
    );

    detailRows.push([
      timestamp,
      formData.formNumber || '',
      fmFirmName, fmGstNo, fmArea, fmCustomer,
      formData.whatsapp || '', fmSalesPerson,
      fmOrderDate, fmDelivDate, fmClientType,
      fmPaymentTerms, formData.paymentDays || '',
      fmTransportMode, fmTransName,
      formData.transporterNumber || '', fmVehicle, formData.driverNumber || '',
      fmItemName, itemReams, item.box || 0,
      round2(itemOrigRate), round2(itemDiscount), round2(itemRate),
      round2(itemFreight), itemValue, itemWithFreight, itemGST, itemGrand,
      isFirst ? fmRemarks : '',
      isFirst ? fmEmail : '',
      ''
    ]);
  }

  var freightSummary = totalFreight;
  var subTotal   = round2(netVal + freightSummary);
  var gstSummary = round2(subTotal * 0.18);
  var grandTotal = round2(subTotal + gstSummary);

  var summaryRow = [
    timestamp,
    formData.formNumber || '',
    fmFirmName, fmGstNo, fmArea, fmCustomer,
    formData.whatsapp || '', fmSalesPerson,
    fmOrderDate, fmDelivDate, fmClientType,
    fmPaymentTerms, formData.paymentDays || '',
    fmTransportMode, fmTransName,
    formData.transporterNumber || '', fmVehicle, formData.driverNumber || '',
    itemsStrParts.join('\n'),
    totalReamsSummary, round2(totalBoxesSummary),
    round2(grossVal), round2(totalDiscSummary), round2(netVal),
    round2(freightSummary), gstSummary, grandTotal,
    fmRemarks,
    fmEmail,
    ''
  ];

  return {
    detailHeaders: ORDER_DETAIL_HEADERS,
    summaryHeaders: ORDER_SUMMARY_HEADERS,
    detailRows: detailRows,
    summaryRow: summaryRow,
    grandTotal: grandTotal,
    netVal: round2(netVal),
    gstSummary: gstSummary,
    freightSummary: round2(freightSummary),
    totalDiscSummary: round2(totalDiscSummary),
    grossVal: round2(grossVal)
  };
}


function processOrder(formData, timestamp) {
  var items = formData.items || [];
  if (!items.length) return { success: false, message: 'KOI ITEM NAHI MILA.' };

  var built = buildOrderRowsAndSummary(formData, timestamp);

  var detailSheet  = getOrCreateSheet('Orders', built.detailHeaders);
  var summarySheet = getOrCreateSheet('Orders_Summary', built.summaryHeaders);

  var detailStartRow = detailSheet.getLastRow() + 1;
  detailSheet.getRange(detailStartRow, 1, built.detailRows.length, built.detailHeaders.length)
             .setValues(built.detailRows);
  applyBatchRowFormat(detailSheet, detailStartRow, built.detailRows.length, built.detailHeaders.length);

  var summaryStartRow = summarySheet.getLastRow() + 1;
  summarySheet.getRange(summaryStartRow, 1, 1, built.summaryHeaders.length)
              .setValues([built.summaryRow]);
  applyBatchRowFormat(summarySheet, summaryStartRow, 1, built.summaryHeaders.length);

  bumpSerialCache('Order', formData.formNumber);

  var pdfUrl = '';
  try {
    if (formData.pdfBase64) {
      var firmSlug = cleanFirmSlug(formData.firmName);
      var fnoClean = (formData.formNumber || 'ORDER').replace(/\//g, '-');
      var pdfFileName = fnoClean + '_' + firmSlug + '.pdf';

      pdfUrl = savePDFToDriveAndGetUrl(
        formData.pdfBase64,
        pdfFileName,
        'Order',
        formData.formNumber
      );

      if (pdfUrl) {
        detailSheet.getRange(detailStartRow, built.detailHeaders.length).setValue(pdfUrl);
        summarySheet.getRange(summaryStartRow, built.summaryHeaders.length).setValue(pdfUrl);
        Logger.log('PDF SAVED: ' + pdfUrl);
      }
    }
  } catch(pdfErr) {
    Logger.log('PDF ERROR: ' + pdfErr.message);
  }

  try {
    sendWhatsAppOrder(formData, built.grandTotal, built.netVal, built.gstSummary, built.freightSummary, built.totalDiscSummary, built.grossVal, pdfUrl, false);
  } catch(waErr) {
    Logger.log('WHATSAPP ERROR: ' + waErr.message);
  }

  return {
    success: true,
    message: 'ORDER SAVED! (' + items.length + ' ITEMS)',
    pdfUrl: pdfUrl
  };
}


// Modifies an existing order in-place: same PI No., rows for that order are
// replaced (deleted + re-inserted at the same position) instead of a new
// entry being appended. Used by the "Modify Order" search-by-PI feature.
function updateOrder(formData, timestamp) {
  var items = formData.items || [];
  if (!items.length) return { success: false, message: 'KOI ITEM NAHI MILA.' };

  var formNumber = (formData.formNumber || '').toString().trim();
  if (!formNumber) return { success: false, message: 'FORM NUMBER MISSING.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var detailSheet  = getOrCreateSheet('Orders', ORDER_DETAIL_HEADERS);
  var summarySheet = getOrCreateSheet('Orders_Summary', ORDER_SUMMARY_HEADERS);

  var origTimestamp = timestamp;

  // Locate the contiguous block of detail rows belonging to this PI No.
  // (all items of one order are always written together, so they sit
  // next to each other) and remove it.
  var detailStartRow = -1, detailCount = 0;
  var lastRow = detailSheet.getLastRow();
  if (lastRow >= 2) {
    var fnCol = detailSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = 0; i < fnCol.length; i++) {
      var cellVal = (fnCol[i][0] || '').toString().trim();
      if (cellVal === formNumber) {
        if (detailStartRow === -1) detailStartRow = i + 2;
        detailCount++;
      } else if (detailStartRow !== -1) {
        break;
      }
    }
  }
  if (detailStartRow !== -1) {
    var origTsCell = detailSheet.getRange(detailStartRow, 1).getValue();
    if (origTsCell) origTimestamp = origTsCell;
    detailSheet.deleteRows(detailStartRow, detailCount);
  }

  // Locate & remove the matching summary row.
  var summaryRowIdx = -1;
  var sLastRow = summarySheet.getLastRow();
  if (sLastRow >= 2) {
    var sFnCol = summarySheet.getRange(2, 2, sLastRow - 1, 1).getValues();
    for (var j = 0; j < sFnCol.length; j++) {
      if ((sFnCol[j][0] || '').toString().trim() === formNumber) { summaryRowIdx = j + 2; break; }
    }
  }
  if (summaryRowIdx !== -1) summarySheet.deleteRow(summaryRowIdx);

  var built = buildOrderRowsAndSummary(formData, origTimestamp);

  var detailInsertAt = (detailStartRow !== -1) ? detailStartRow : (detailSheet.getLastRow() + 1);
  detailSheet.insertRowsBefore(detailInsertAt, built.detailRows.length);
  detailSheet.getRange(detailInsertAt, 1, built.detailRows.length, built.detailHeaders.length)
             .setValues(built.detailRows);
  applyBatchRowFormat(detailSheet, detailInsertAt, built.detailRows.length, built.detailHeaders.length);

  var summaryInsertAt = (summaryRowIdx !== -1) ? summaryRowIdx : (summarySheet.getLastRow() + 1);
  summarySheet.insertRowsBefore(summaryInsertAt, 1);
  summarySheet.getRange(summaryInsertAt, 1, 1, built.summaryHeaders.length)
              .setValues([built.summaryRow]);
  applyBatchRowFormat(summarySheet, summaryInsertAt, 1, built.summaryHeaders.length);

  var pdfUrl = '';
  try {
    if (formData.pdfBase64) {
      var firmSlug = cleanFirmSlug(formData.firmName);
      var fnoClean = formNumber.replace(/\//g, '-');
      var pdfFileName = fnoClean + '_' + firmSlug + '.pdf';

      pdfUrl = savePDFToDriveAndGetUrl(formData.pdfBase64, pdfFileName, 'Order', formNumber);

      if (pdfUrl) {
        detailSheet.getRange(detailInsertAt, built.detailHeaders.length).setValue(pdfUrl);
        summarySheet.getRange(summaryInsertAt, built.summaryHeaders.length).setValue(pdfUrl);
      }
    }
  } catch(pdfErr) {
    Logger.log('PDF ERROR (update): ' + pdfErr.message);
  }

  try {
    sendWhatsAppOrder(formData, built.grandTotal, built.netVal, built.gstSummary, built.freightSummary, built.totalDiscSummary, built.grossVal, pdfUrl, true);
  } catch(waErr) {
    Logger.log('WHATSAPP ERROR (update): ' + waErr.message);
  }

  return {
    success: true,
    message: 'ORDER UPDATED! (' + items.length + ' ITEMS)',
    pdfUrl: pdfUrl
  };
}


// Looks up an existing order by its PI No. (Form No) so it can be loaded
// back into the form for editing. Returns all item rows + the shared
// header fields from the first row.
var ORDER_LOOKUP_COLS = {
  timestamp: 0, formNo: 1, firmName: 2, gstNo: 3, area: 4, customerName: 5,
  whatsapp: 6, salesPerson: 7, orderDate: 8, deliveryDate: 9,
  clientType: 10, paymentTerms: 11, paymentDays: 12,
  transportMode: 13, transportName: 14, transporterNumber: 15, vehicleNumber: 16, driverNumber: 17,
  productName: 18, ream: 19, box: 20, origRate: 21, discount: 22, netRate: 23, freight: 24,
  remarks: 29, email: 30
};

// Accepts a full PI No. (e.g. "PI-SP/26-27/00425") OR just a fragment/number
// (e.g. "425" or "00425") so the user doesn't have to type the whole thing.
// Exact match wins outright; otherwise every order whose PI No. contains the
// typed text is offered back as a pick-list for the client to disambiguate.
function getOrderByFormNumber(query) {
  query = (query || '').toString().trim();
  if (!query) return { found: false, message: 'PI No. ya order number daalein.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Orders');
  if (!sheet) return { found: false, message: 'Orders sheet not found.' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { found: false, message: 'No orders found.' };

  var numCols = ORDER_DETAIL_HEADERS.length;
  var values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var col = ORDER_LOOKUP_COLS;

  // Distinct PI numbers in the sheet, most-recently-added first.
  var seen = {};
  var order = [];
  for (var i = 0; i < values.length; i++) {
    var fn = (values[i][col.formNo] || '').toString().trim();
    if (!fn || seen[fn]) continue;
    seen[fn] = { formNumber: fn, firmName: values[i][col.firmName] || '', customerName: values[i][col.customerName] || '' };
    order.push(fn);
  }
  order.reverse();

  var qUpper = query.toUpperCase();
  var formNumber = null;
  for (var k = 0; k < order.length; k++) {
    if (order[k].toUpperCase() === qUpper) { formNumber = order[k]; break; }
  }

  if (!formNumber) {
    var matches = order.filter(function(fn){ return fn.toUpperCase().indexOf(qUpper) !== -1; });
    if (!matches.length) {
      return { found: false, message: '"' + query + '" ke liye koi order nahi mila.' };
    }
    if (matches.length > 1) {
      return {
        found: false,
        multiple: true,
        message: matches.length + ' orders mile, ek select karein.',
        matches: matches.slice(0, 20).map(function(fn){ return seen[fn]; })
      };
    }
    formNumber = matches[0];
  }

  var matchedRows = [];
  for (var m = 0; m < values.length; m++) {
    var cellFn = (values[m][col.formNo] || '').toString().trim();
    if (cellFn === formNumber) matchedRows.push(values[m]);
  }

  if (!matchedRows.length) {
    return { found: false, message: 'PI No. "' + formNumber + '" ke liye koi order nahi mila.' };
  }

  var first = matchedRows[0];

  var toISODate = function(v) {
    if (!v) return '';
    try {
      if (v instanceof Date) {
        if (isNaN(v.getTime())) return '';
        return Utilities.formatDate(v, getTZ(), 'yyyy-MM-dd');
      }
      var parts = v.toString().trim().split('/');
      if (parts.length === 3) {
        var d = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
        if (!isNaN(d.getTime())) return Utilities.formatDate(d, getTZ(), 'yyyy-MM-dd');
      }
      return '';
    } catch (e) { return ''; }
  };

  var items = matchedRows.map(function(r) {
    return {
      name: r[col.productName] || '',
      ream: parseFloat(r[col.ream]) || 0,
      box: parseFloat(r[col.box]) || 0,
      origRate: parseFloat(r[col.origRate]) || 0,
      rate: parseFloat(r[col.netRate]) || 0,
      discount: parseFloat(r[col.discount]) || 0
    };
  });

  var totalFreight = 0;
  matchedRows.forEach(function(r) { totalFreight += parseFloat(r[col.freight]) || 0; });

  var tsVal = first[col.timestamp];
  var order = {
    formNumber: formNumber,
    timestamp: (tsVal instanceof Date) ? tsVal.toISOString() : (tsVal || ''),
    firmName: first[col.firmName] || '',
    gstNo: first[col.gstNo] || '',
    areaName: first[col.area] || '',
    customerName: first[col.customerName] || '',
    whatsapp: first[col.whatsapp] || '',
    salesPerson: first[col.salesPerson] || '',
    orderDateISO: toISODate(first[col.orderDate]),
    deliveryDateISO: toISODate(first[col.deliveryDate]),
    clientType: first[col.clientType] || '',
    paymentTerms: first[col.paymentTerms] || '',
    paymentDays: first[col.paymentDays] || '',
    transportMode: first[col.transportMode] || '',
    transportName: first[col.transportName] || '',
    transporterNumber: first[col.transporterNumber] || '',
    vehicleNumber: first[col.vehicleNumber] || '',
    driverNumber: first[col.driverNumber] || '',
    freightAmount: Math.round(totalFreight * 100) / 100,
    remarks: first[col.remarks] || '',
    items: items
  };

  return { found: true, order: order };
}


function processEnquiry(formData, timestamp) {
  var headers = [
    'Timestamp', 'Form No', 'Firm Name', 'GST No', 'Area', 'Customer Name',
    'WhatsApp', 'Client Type', 'Client Source', 'Reference Name', 'Sales Person',
    'Product Name', 'Ream Qty', 'Box Qty', 'Rate',
    'Remarks', 'Visiting Card Link', 'Shop Photo Link', 'Filled By (Name)', 'PDF URL'
  ];

  var sheet = getOrCreateSheet('Enquiries', headers);
  var items = formData.items || [];
  if (!items.length) return { success: false, message: 'KOI ITEM NAHI MILA.' };

  var vcLink = '', shopLink = '';
  try {
    if (formData.visitingCardData)
      vcLink = saveBase64ToDrive(
        formData.visitingCardData,
        (formData.formNumber||'ENQ').replace(/\//g,'-')+'_VC.jpg',
        'SP_Enquiry_Uploads'
      );
    if (formData.shopPhotoData)
      shopLink = saveBase64ToDrive(
        formData.shopPhotoData,
        (formData.formNumber||'ENQ').replace(/\//g,'-')+'_Shop.jpg',
        'SP_Enquiry_Uploads'
      );
  } catch(imgErr) {
    Logger.log('IMAGE ERROR: ' + imgErr.message);
  }

  // Convert Enquiry Details to ALL CAPS
  var fmFirmName      = toAllCaps(formData.firmName);
  var fmGstNo         = toAllCaps(formData.gstNo);
  var fmArea          = toAllCaps(formData.areaName);
  var fmCustomer      = toAllCaps(formData.customerName);
  var fmClientType    = toAllCaps(formData.clientType || '');
  var fmClientSource  = toAllCaps(formData.clientSource || '');
  var fmReferenceName = toAllCaps(formData.referenceName || '');
  var fmSalesPerson   = toAllCaps(formData.salesPerson || '');
  var fmRemarks       = toAllCaps(formData.remarks || '');
  var fmEmail         = toAllCaps(formData.filledByEmail || '');

  var allRows = [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var isFirst = (i === 0);
    allRows.push([
      timestamp,
      formData.formNumber || '',
      fmFirmName,
      fmGstNo,
      fmArea,
      fmCustomer,
      formData.whatsapp || '',
      fmClientType,
      fmClientSource,
      fmReferenceName,
      fmSalesPerson,
      toAllCaps(item.name || ''),
      parseFloat(item.ream) || 0,
      item.box || 0,
      parseFloat(item.rate) || 0,
      isFirst ? fmRemarks : '',
      isFirst ? vcLink   : '',
      isFirst ? shopLink : '',
      isFirst ? fmEmail : '',
      ''
    ]);
  }

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, allRows.length, headers.length).setValues(allRows);
  applyBatchRowFormat(sheet, startRow, allRows.length, headers.length);

  bumpSerialCache('Enquiry', formData.formNumber);

  var pdfUrl = '';
  try {
    if (formData.pdfBase64) {
      var firmSlug = cleanFirmSlug(formData.firmName);
      var fnoClean = (formData.formNumber || 'ENQUIRY').replace(/\//g, '-');
      var pdfFileName = fnoClean + '_' + firmSlug + '.pdf';

      pdfUrl = savePDFToDriveAndGetUrl(
        formData.pdfBase64,
        pdfFileName,
        'Enquiry',
        formData.formNumber
      );

      if (pdfUrl) {
        sheet.getRange(startRow, headers.length).setValue(pdfUrl);
        Logger.log('ENQUIRY PDF SAVED: ' + pdfUrl);
      }
    }
  } catch(pdfErr) {
    Logger.log('PDF ERROR: ' + pdfErr.message);
  }

  try {
    sendWhatsAppEnquiry(formData);
  } catch(waErr) {
    Logger.log('WHATSAPP ENQUIRY ERROR: ' + waErr.message);
  }

  return {
    success: true,
    message: 'ENQUIRY SAVED! (' + items.length + ' ITEMS)',
    pdfUrl: pdfUrl
  };
}


function sendWhatsAppOrder(formData, grandTotal, netVal, gst, freight, totalDisc, grossVal, pdfUrl, isUpdate) {
  var lines = [];
  lines.push(isUpdate ? '🟠 *ORDER MODIFIED — SATIJA PAPER*' : '🟢 *NEW ORDER — SATIJA PAPER*');
  lines.push('━━━━━━━━━━━━━━━━━━━━');
  lines.push('📋 *PO ID:* ' + (formData.formNumber || '—'));
  lines.push('📅 *ORDER DATE:* ' + (formData.orderDate || '—'));
  lines.push('🚚 *DELIVERY:* ' + (formData.deliveryDate || '—'));
  lines.push('');
  lines.push('🏢 *PARTY:* ' + toAllCaps(formData.firmName));
  lines.push('📍 *AREA:* ' + toAllCaps(formData.areaName));
  lines.push('👤 *CUSTOMER:* ' + toAllCaps(formData.customerName));
  lines.push('📱 *MOBILE:* ' + (formData.whatsapp || '—'));
  lines.push('👨‍💼 *SALES:* ' + toAllCaps(formData.salesPerson || '—'));
  lines.push('💳 *PAYMENT:* ' + (
    toAllCaps(formData.paymentTerms) === 'CREDIT DAYS'
      ? (formData.paymentDays + ' DAYS CREDIT')
      : toAllCaps(formData.paymentTerms || '—')
  ));
  lines.push('');
  lines.push('📦 *ITEMS:*');

  (formData.items || []).forEach(function(item, i) {
    var origR    = parseFloat(item.origRate || item.rate) || 0;
    var reamQty  = parseFloat(item.ream) || 0;
    var boxQty   = Math.round(parseFloat(item.box) || 0);
    var disc     = parseFloat(item.discount) || 0;
    var netRate  = parseFloat(item.rate) || 0;
    var netAmt   = Math.round(reamQty * netRate);
    lines.push(
      (i + 1) + '. *' + toAllCaps(item.name || '') + '*\n' +
      '   ' + reamQty + ' REAM (' + boxQty + ' BOX)' +
      ' | RATE: ₹' + Math.round(origR) +
      (disc > 0 ? ' [-₹' + disc + ']' : '') +
      ' = ₹' + netAmt
    );
  });

  lines.push('');
  lines.push('━━━━━━━━━━━━━━━━━━━━');
  if (totalDisc > 0) {
    lines.push('💰 GROSS:    ₹' + Math.round(grossVal).toLocaleString());
    lines.push('🎁 SUPPORT: -₹' + Math.round(totalDisc).toLocaleString());
    lines.push('   NET:      ₹' + Math.round(netVal).toLocaleString());
  } else {
    lines.push('💰 NET VALUE: ₹' + Math.round(netVal).toLocaleString());
  }
  if (freight > 0) lines.push('🚛 Freight:   ₹' + Math.round(freight).toLocaleString());
  lines.push('🧾 GST (18%): ₹' + Math.round(gst).toLocaleString());
  lines.push('✅ *GRAND TOTAL: ₹' + Math.round(grandTotal).toLocaleString() + '*');

  if (formData.transportMode) {
    lines.push('');
    lines.push('🚚 *TRANSPORT:* ' + toAllCaps(formData.transportMode) +
      (formData.transportName ? ' — ' + toAllCaps(formData.transportName) : '') +
      (formData.vehicleNumber ? ' | ' + formData.vehicleNumber.toUpperCase() : ''));
  }

  if (formData.remarks) {
    lines.push('');
    lines.push('📝 *REMARKS:* ' + toAllCaps(formData.remarks));
  }

  if (pdfUrl) {
    lines.push('');
    lines.push('📄 *PDF:* ' + pdfUrl);
  }

  lines.push('━━━━━━━━━━━━━━━━━━━━');
  lines.push('_SATIJA PAPER | AUTO-GENERATED_');

  callMeBot(lines.join('\n'));
}


function sendWhatsAppEnquiry(formData) {
  var lines = [];
  lines.push('🔵 *NEW ENQUIRY — SATIJA PAPER*');
  lines.push('━━━━━━━━━━━━━━━━━━━━');
  lines.push('📋 *ENQ ID:* ' + (formData.formNumber || '—'));
  lines.push('🏢 *PARTY:* ' + toAllCaps(formData.firmName));
  lines.push('📍 *AREA:* ' + toAllCaps(formData.areaName));
  lines.push('👤 *CUSTOMER:* ' + toAllCaps(formData.customerName));
  lines.push('📱 *MOBILE:* ' + (formData.whatsapp || '—'));
  lines.push('👨‍💼 *SALES:* ' + toAllCaps(formData.salesPerson || '—'));
  lines.push('📌 *SOURCE:* ' + toAllCaps(formData.clientSource || '—'));
  if (formData.referenceName) lines.push('🔗 *REFERENCE:* ' + toAllCaps(formData.referenceName));
  lines.push('');
  lines.push('📦 *PRODUCTS ENQUIRED:*');
  (formData.items || []).forEach(function(item, i) {
    lines.push(
      (i + 1) + '. ' + toAllCaps(item.name || '') +
      ' — ' + (parseFloat(item.ream) || 0) + ' REAM' +
      ' (' + Math.round(parseFloat(item.box) || 0) + ' BOX)'
    );
  });
  if (formData.remarks) {
    lines.push('');
    lines.push('📝 *REMARKS:* ' + toAllCaps(formData.remarks));
  }
  lines.push('━━━━━━━━━━━━━━━━━━━━');
  lines.push('_SATIJA PAPER | AUTO-GENERATED_');

  callMeBot(lines.join('\n'));
}


// WhatsApp sends are queued and flushed by a short-lived trigger instead of
// being sent inline, so doPost can write to the sheet and return to the
// client without waiting on the UltraMsg network call (this was the main
// cause of "slow saving").
function callMeBot(message) {
  queueWhatsAppMessage(message);
}

function queueWhatsAppMessage(message) {
  try {
    var props = PropertiesService.getScriptProperties();
    var key = 'WA_MSG_' + new Date().getTime() + '_' + Math.floor(Math.random() * 100000);
    props.setProperty(key, message);
    ScriptApp.newTrigger('flushWhatsAppQueue').timeBased().after(1000).create();
  } catch (e) {
    Logger.log('WhatsApp queue error, sending inline instead: ' + e.message);
    sendWhatsAppMsg(WHATSAPP_CONFIG.PHONE, message);
  }
}

function flushWhatsAppQueue() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  for (var key in all) {
    if (key.indexOf('WA_MSG_') === 0) {
      try {
        sendWhatsAppMsg(WHATSAPP_CONFIG.PHONE, all[key]);
      } catch (e) {
        Logger.log('WhatsApp flush send error: ' + e.message);
      }
      props.deleteProperty(key);
    }
  }
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'flushWhatsAppQueue') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function sendWhatsAppMsg(toNumber, message) {
  if (!WHATSAPP_CONFIG.TOKEN || !WHATSAPP_CONFIG.INSTANCE) {
    Logger.log('UltraMsg config set nahi hai — WhatsApp skip.');
    return;
  }
  var phone = toNumber.toString().replace(/[^0-9]/g, '');
  if (!phone.startsWith('91')) phone = '91' + phone;

  var url = 'https://api.ultramsg.com/' + WHATSAPP_CONFIG.INSTANCE + '/messages/chat';
  var payload = {
    token: WHATSAPP_CONFIG.TOKEN,
    to:    '+' + phone,
    body:  message,
    priority: 1
  };
  var options = {
    method: 'POST',
    contentType: 'application/x-www-form-urlencoded',
    payload: payload,
    muteHttpExceptions: true
  };
  try {
    var resp = UrlFetchApp.fetch(url, options);
    Logger.log('UltraMsg: ' + resp.getResponseCode() + ' — ' + resp.getContentText().substring(0, 200));
  } catch(e) {
    Logger.log('UltraMsg fetch error: ' + e.message);
  }
}


function savePDFToDriveAndGetUrl(pdfBase64, fileName, formType, formNumber) {
  try {
    var folderName = (formType === 'Order') ? 'Order Form' : 'Enquiry PDF';
    var folder = getOrCreateFolder(folderName);
    var data = pdfBase64.indexOf(',') > -1 ? pdfBase64.split(',')[1] : pdfBase64;
    var decoded = Utilities.base64Decode(data);
    var blob = Utilities.newBlob(decoded, 'application/pdf', fileName);
    var file = folder.createFile(blob);
    if (formNumber) file.setDescription('SYSTEM GENERATED ' + formType.toUpperCase() + ' NO: ' + formNumber);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = file.getUrl();
    Logger.log('PDF SAVED TO DRIVE: ' + fileName + ' -> ' + url);
    return url;
  } catch (e) {
    Logger.log('PDF SAVE ERROR: ' + e.message);
    return '';
  }
}

function savePDFToDrive(pdfBase64, fileName, formType, formNumber) {
  return savePDFToDriveAndGetUrl(pdfBase64, fileName, formType, formNumber) !== '';
}

function saveBase64ToDrive(base64Data, fileName, folderName) {
  try {
    var data = base64Data.indexOf(',') > -1 ? base64Data.split(',')[1] : base64Data;
    var folder = getOrCreateFolder(folderName);
    var decoded = Utilities.base64Decode(data);
    var mime = fileName.toLowerCase().endsWith('.pdf') ? 'application/pdf' : 'image/jpeg';
    var blob = Utilities.newBlob(decoded, mime, fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log('FILE SAVE ERROR: ' + e.message);
    return '';
  }
}

function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

function testFolderCreation() {
  Logger.log("Order Folder URL: " + getOrCreateFolder('Order Form').getUrl());
  Logger.log("Enquiry Folder URL: " + getOrCreateFolder('Enquiry PDF').getUrl());
  Logger.log("Uploads Folder URL: " + getOrCreateFolder('Enquiry Uploads').getUrl());
}


function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.formType === 'Order' || data.formType === 'Enquiry' && data.items) {
      var result = processForm(data);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var formNumber = getNextFormNumber('Enquiry');
    var formData = {
      formType: 'Enquiry',
      formNumber: formNumber,
      firmName: '',
      gstNo: '',
      areaName: '',
      customerName: toAllCaps(data.name || ''),
      whatsapp: data.whatsapp || '',
      clientType: '',
      clientSource: 'WEBSITE ENQUIRY',
      referenceName: '',
      salesPerson: '',
      remarks: '',
      email: toAllCaps(data.email || ''),
      items: [{
        name: toAllCaps(data.requirements || 'GENERAL ENQUIRY'),
        ream: parseFloat((data.quantity || '').toString().replace(/[^\d.]/g, '')) || 0,
        box: 0,
        rate: 0
      }],
      visitingCardData: '',
      shopPhotoData: '',
      pdfBase64: ''
    };
    var webhookResult = processEnquiry(formData, new Date());
    return ContentService.createTextOutput(JSON.stringify(webhookResult))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


function migrateOrdersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var oldSheet = ss.getSheetByName('Orders');
  if (oldSheet) {
    var timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmm');
    oldSheet.setName('Orders_Backup_' + timestamp);
    Logger.log('Old sheet renamed to: Orders_Backup_' + timestamp);
  }
  var oldSummary = ss.getSheetByName('Orders_Summary');
  if (oldSummary) {
    var ts2 = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmm');
    oldSummary.setName('Orders_Summary_Backup_' + ts2);
    Logger.log('Old summary sheet renamed.');
  }
  Logger.log('Done! Next order will create fresh Orders + Orders_Summary sheets.');
}

function givePermission() {
  UrlFetchApp.fetch('https://www.google.com');
  Logger.log('Permission granted!');
}

function testWhatsApp() {
  callMeBot('✅ TEST MESSAGE FROM SATIJA PAPER WMS!\nULTRAMSG SETUP SUCCESSFUL 🎉\nAB ORDERS PE AUTO WHATSAPP AAYEGA!');
}

function syncLastRowToFirebase(e) {
  if (e && e.changeType === 'REMOVE_ROW') {
    Logger.log('Row deletion detected. Stopping sync to avoid duplicate.');
    return;
  }

  Utilities.sleep(1500);

  var projectId = "sp-dashboard-1e9c8";
  var firestoreUrl = "https://firestore.googleapis.com/v1/projects/" + projectId + "/databases/(default)/documents/enquiries";

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Enquiries");
  var lastRow = sheet.getLastRow();

  var timestamp = sheet.getRange(lastRow, 1).getDisplayValue();
  var name       = sheet.getRange(lastRow, 6).getValue();
  var whatsapp  = sheet.getRange(lastRow, 7).getValue();
  var req       = sheet.getRange(lastRow, 12).getValue();
  var qty       = sheet.getRange(lastRow, 13).getValue();
  var email     = sheet.getRange(lastRow, 19).getValue();

  Logger.log("Processing Row: " + lastRow + " | Name: " + name + " | WA: " + whatsapp);

  if (!name || !whatsapp) {
    Logger.log("Required fields are empty. Aborting Firebase sync.");
    return;
  }

  var payload = {
    "fields": {
      "name": { "stringValue": String(name) },
      "wa": { "stringValue": String(whatsapp) },
      "email": { "stringValue": String(email || "") },
      "qty": { "stringValue": String(qty || "0") },
      "req": { "stringValue": String(req || "") },
      "status": { "stringValue": "NEW" },
      "date": { "stringValue": String(timestamp) }
    }
  };

  var options = {
    "method" : "post",
    "contentType": "application/json",
    "payload" : JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(firestoreUrl, options);
  Logger.log("Firebase Response: " + response.getContentText());
}
