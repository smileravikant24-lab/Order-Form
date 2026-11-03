var FIRM_SHEET_URL  = 'https://docs.google.com/spreadsheets/d/1WLgaZ21Jt4uckn1zyi01F5yLcwtWTmUTpmwQGW7hQGI/edit';
var FIRM_SHEET_NAME = 'WMS_Firms';
var FIRM_COLUMN     = 1;
var GST_COLUMN      = 2;

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

    if (e.parameter.action === 'fixShiftedRows') {
    var fixResult = fixShiftedRows773_774();
    return ContentService.createTextOutput(JSON.stringify(fixResult))
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
    .setTitle('Satija Paper ΓÇö Order / Enquiry')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Scans the sheet once to find the highest serial already used for this
// form-type + financial-year combo. Only called the very first time a
// combo is needed after deployment (or a new FY starts) ΓÇö after that,
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
// forward ΓÇö merely previewing a number (page load, tab switch, refresh)
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

// PEEK ONLY ΓÇö this is what the form shows and what refreshes on every page
// load / tab switch. It must be side-effect-free: it never reserves or
// consumes a number, it just reports "next one would be X" based on the
// cached max (falling back to a one-time sheet scan if nothing is cached
// yet). The number only actually gets used once bumpSerialCache() runs,
// which happens after a real save ΓÇö see processOrder / processEnquiry.
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

// Adds a party the user has confirmed is new to the WMS_Firms master
// sheet (the same sheet getFirmData() reads from). Silently no-ops if the
// name is already present (case-insensitive) so a retry can't duplicate it.
function addNewFirmToMaster(name, gst) {
  name = (name || '').toString().trim().toUpperCase();
  gst = (gst || '').toString().trim().toUpperCase();
  if (!name) return { success: false, message: 'Firm name required.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var externalSS = SpreadsheetApp.openByUrl(FIRM_SHEET_URL);
    var sheet = externalSS.getSheetByName(FIRM_SHEET_NAME) || externalSS.getSheets()[0];
    var lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      var existing = sheet.getRange(2, FIRM_COLUMN, lastRow - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if ((existing[i][0] || '').toString().trim().toUpperCase() === name) {
          return { success: true, alreadyExists: true };
        }
      }
    }

    var numCols = Math.max(FIRM_COLUMN, GST_COLUMN);
    var newRow = new Array(numCols).fill('');
    newRow[FIRM_COLUMN - 1] = name;
    newRow[GST_COLUMN - 1] = gst;
    sheet.appendRow(newRow);
    try { CacheService.getScriptCache().remove(FIRM_CACHE_KEY); } catch (cacheErr) {}
    return { success: true };
  } catch (e) {
    Logger.log('addNewFirmToMaster Error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
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

// Γ£à NEW FUNCTION: To force ALL CAPS everywhere safely
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
  'Remarks', 'Form Filler Name', 'PDF URL'
];

var ORDER_SUMMARY_HEADERS = [
  'Timestamp', 'Form No', 'Firm Name', 'GST No', 'Area', 'Customer Name',
  'WhatsApp', 'Sales Person', 'Order Date', 'Delivery Date',
  'Client Type', 'Payment Terms', 'Payment Days',
  'Transport Mode', 'Transport Name', 'Transporter Number', 'Vehicle Number', 'Driver Number',
  'Items (Name | Ream | Box | Rate)', 'Total Reams', 'Total Boxes',
  'Gross Value', 'Total Discount', 'Net Value',
  'Freight', 'GST (18%)', 'Grand Total',
  'Remarks', 'Form Filler Name', 'PDF URL'
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
  var fmEmail      = (formData.email || '').toString().trim().toLowerCase();
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
      isFirst ? toAllCaps(formData.filledByEmail || formData.formFilledBy || '') : '',
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
    toAllCaps(formData.filledByEmail || formData.formFilledBy || ''),
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
    formNumber: formData.formNumber,
    message: 'ORDER SAVED! (' + items.length + ' ITEMS)',
    pdfUrl: pdfUrl
  };
}


// Clean text helper for robust comparison
function cleanOrderText(str) {
  return String(str || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

// Modifies an existing order IN-PLACE: SAME ROW, SAME PO NUMBER.
// No row deletion/insertion if row count is unchanged.
// NEVER creates a new entry at the bottom.
function updateOrder(formData, timestamp) {
  var items = formData.items || [];
  if (!items.length) return { success: false, message: 'KOI ITEM NAHI MILA.' };

  var formNumber = (formData.formNumber || '').toString().trim();
  if (!formNumber) return { success: false, message: 'FORM NUMBER (PO NO) MISSING.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var detailSheet  = getOrCreateSheet('Orders', ORDER_DETAIL_HEADERS);
  var summarySheet = getOrCreateSheet('Orders_Summary', ORDER_SUMMARY_HEADERS);

  var cleanTargetFn = cleanOrderText(formNumber);
  var lastRow = detailSheet.getLastRow();
  if (lastRow < 2) {
    return { success: false, message: 'Orders sheet me koi data nahi mila.' };
  }

  var detailStartRow = -1;
  var detailCount = 0;

  // 1. If frontend provided rowNumber, verify it first for fast and 100% accurate matching
  if (formData.rowNumber && typeof formData.rowNumber === 'number' && formData.rowNumber >= 2 && formData.rowNumber <= lastRow) {
    var checkVal = cleanOrderText(detailSheet.getRange(formData.rowNumber, 2).getValue());
    if (checkVal === cleanTargetFn || checkVal.includes(cleanTargetFn) || cleanTargetFn.includes(checkVal)) {
      detailStartRow = formData.rowNumber;
    }
  }

  // 2. Scan Column B (Form No) if not verified by rowNumber
  var fnCol = detailSheet.getRange(2, 2, lastRow - 1, 1).getValues();
  if (detailStartRow === -1) {
    for (var i = 0; i < fnCol.length; i++) {
      var cellVal = cleanOrderText(fnCol[i][0]);
      if (cellVal === cleanTargetFn) {
        detailStartRow = i + 2;
        break;
      }
    }
    // Fallback: match by trailing digits/number if exact match failed
    if (detailStartRow === -1) {
      var numMatch = formNumber.match(/\d+$/);
      if (numMatch) {
        var numStr = numMatch[0];
        for (var k = 0; k < fnCol.length; k++) {
          var cVal = cleanOrderText(fnCol[k][0]);
          if (cVal.endsWith(numStr) || cVal.includes('/' + numStr)) {
            detailStartRow = k + 2;
            break;
          }
        }
      }
    }
  }

  // CRITICAL SAFETY: If order row is NOT found, NEVER append as a new entry!
  if (detailStartRow === -1) {
    return { 
      success: false, 
      message: 'Order "' + formNumber + '" sheet me nahi mila! Kripya PO/PI number check karein.' 
    };
  }

  // Count how many consecutive rows currently belong to this order in the sheet
  detailCount = 0;
  for (var r = detailStartRow; r <= lastRow; r++) {
    var rVal = cleanOrderText(detailSheet.getRange(r, 2).getValue());
    if (rVal === cleanTargetFn || rVal.includes(cleanTargetFn) || cleanTargetFn.includes(rVal)) {
      detailCount++;
    } else {
      break;
    }
  }
  if (detailCount === 0) detailCount = 1;

  // Preserve original creation timestamp
  var origTsCell = detailSheet.getRange(detailStartRow, 1).getValue();
  var origTimestamp = origTsCell ? origTsCell : timestamp;

  // Build the updated detail and summary rows
  var built = buildOrderRowsAndSummary(formData, origTimestamp);
  var newRowCount = built.detailRows.length;
  var numCols = built.detailHeaders.length;

  // IN-PLACE UPDATE IN 'Orders' SHEET (SAME ROW):
  if (newRowCount === detailCount) {
    // Exact same number of items: Overwrite in place directly!
    detailSheet.getRange(detailStartRow, 1, newRowCount, numCols).setValues(built.detailRows);
    applyBatchRowFormat(detailSheet, detailStartRow, newRowCount, numCols);
  } else if (newRowCount < detailCount) {
    // Items removed: Overwrite the first newRowCount rows in-place, delete leftover extra rows
    detailSheet.getRange(detailStartRow, 1, newRowCount, numCols).setValues(built.detailRows);
    applyBatchRowFormat(detailSheet, detailStartRow, newRowCount, numCols);
    var extraToDelete = detailCount - newRowCount;
    detailSheet.deleteRows(detailStartRow + newRowCount, extraToDelete);
  } else {
    // Items added: Overwrite existing detailCount rows in-place, insert only the additional rows below
    var existingSlice = built.detailRows.slice(0, detailCount);
    detailSheet.getRange(detailStartRow, 1, detailCount, numCols).setValues(existingSlice);
    applyBatchRowFormat(detailSheet, detailStartRow, detailCount, numCols);

    var extraToAdd = newRowCount - detailCount;
    detailSheet.insertRowsAfter(detailStartRow + detailCount - 1, extraToAdd);
    var addedSlice = built.detailRows.slice(detailCount);
    detailSheet.getRange(detailStartRow + detailCount, 1, extraToAdd, numCols).setValues(addedSlice);
    applyBatchRowFormat(detailSheet, detailStartRow + detailCount, extraToAdd, numCols);
  }

  // Clear any shifted ghost columns beyond numCols for these updated rows
  var maxCols = Math.max(detailSheet.getLastColumn(), numCols + 2);
  if (maxCols > numCols) {
    detailSheet.getRange(detailStartRow, numCols + 1, newRowCount, maxCols - numCols).clearContent();
  }

  // IN-PLACE UPDATE IN 'Orders_Summary' SHEET (SAME ROW):
  var summaryRowIdx = -1;
  var sLastRow = summarySheet.getLastRow();
  if (sLastRow >= 2) {
    var sFnCol = summarySheet.getRange(2, 2, sLastRow - 1, 1).getValues();
    for (var j = 0; j < sFnCol.length; j++) {
      var sVal = cleanOrderText(sFnCol[j][0]);
      if (sVal === cleanTargetFn || sVal.includes(cleanTargetFn) || cleanTargetFn.includes(sVal)) {
        summaryRowIdx = j + 2;
        break;
      }
    }
  }

  if (summaryRowIdx !== -1) {
    // In-place overwrite of summary row! No delete, no insert!
    summarySheet.getRange(summaryRowIdx, 1, 1, built.summaryHeaders.length).setValues([built.summaryRow]);
    applyBatchRowFormat(summarySheet, summaryRowIdx, 1, built.summaryHeaders.length);
    var sMaxCols = Math.max(summarySheet.getLastColumn(), built.summaryHeaders.length + 2);
    if (sMaxCols > built.summaryHeaders.length) {
      summarySheet.getRange(summaryRowIdx, built.summaryHeaders.length + 1, 1, sMaxCols - built.summaryHeaders.length).clearContent();
    }
  } else {
    // If not found in summary, insert at bottom of summary
    var sInsertAt = summarySheet.getLastRow() + 1;
    summarySheet.getRange(sInsertAt, 1, 1, built.summaryHeaders.length).setValues([built.summaryRow]);
    applyBatchRowFormat(summarySheet, sInsertAt, 1, built.summaryHeaders.length);
    summaryRowIdx = sInsertAt;
  }

  // PDF Handling:
  var pdfUrl = '';
  try {
    if (formData.pdfBase64) {
      var firmSlug = cleanFirmSlug(formData.firmName);
      var fnoClean = formNumber.replace(/\//g, '-');
      var pdfFileName = fnoClean + '_' + firmSlug + '.pdf';

      pdfUrl = savePDFToDriveAndGetUrl(formData.pdfBase64, pdfFileName, 'Order', formNumber);

      if (pdfUrl) {
        detailSheet.getRange(detailStartRow, numCols).setValue(pdfUrl);
        if (summaryRowIdx !== -1) {
          summarySheet.getRange(summaryRowIdx, built.summaryHeaders.length).setValue(pdfUrl);
        }
      }
    }
  } catch(pdfErr) {
    Logger.log('PDF ERROR (update): ' + pdfErr.message);
  }

  // WhatsApp Notification (sent as Update)
  try {
    sendWhatsAppOrder(formData, built.grandTotal, built.netVal, built.gstSummary, built.freightSummary, built.totalDiscSummary, built.grossVal, pdfUrl, true);
  } catch(waErr) {
    Logger.log('WHATSAPP ERROR (update): ' + waErr.message);
  }

  return {
    success: true,
    isUpdate: true,
    formNumber: formNumber,
    rowNumber: detailStartRow,
    message: 'ORDER UPDATED IN PLACE! (SAME ROW: ' + detailStartRow + ', SAME PO: ' + formNumber + ')',
    pdfUrl: pdfUrl
  };
}


// Looks up an existing order by its PI No. (Form No) so it can be loaded
// back into the form for editing. Uses DYNAMIC column matching based on
// row 1 of the 'Orders' sheet so column layout changes never break data population.
function getOrderByFormNumber(query) {
  query = (query || '').toString().trim();
  if (!query) return { found: false, message: 'PI No. ya order number daalein.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Orders');
  if (!sheet) return { found: false, message: 'Orders sheet not found.' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { found: false, message: 'No orders found.' };

  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  function findColIndex(candidates, fallback) {
    for (var c = 0; c < candidates.length; c++) {
      var target = candidates[c].toLowerCase().replace(/[^a-z0-9]/g, '');
      for (var h = 0; h < headerRow.length; h++) {
        var hStr = (headerRow[h] || '').toString().toLowerCase().replace(/[^a-z0-9]/g, '');
        if (hStr === target) return h;
      }
    }
    return (typeof fallback === 'number') ? fallback : -1;
  }

  var col = {
    timestamp: findColIndex(['Timestamp', 'Time'], 0),
    formNo: findColIndex(['Form No', 'Form Number', 'PI No', 'PI Number'], 1),
    firmName: findColIndex(['Firm Name', 'Firm'], 2),
    gstNo: findColIndex(['GST No', 'GST', 'GSTIN'], 3),
    area: findColIndex(['Area', 'Area Name'], 4),
    customerName: findColIndex(['Customer Name', 'Customer'], 5),
    whatsapp: findColIndex(['WhatsApp', 'Mobile', 'Phone'], 6),
    salesPerson: findColIndex(['Sales Person', 'Salesperson', 'SP'], 7),
    orderDate: findColIndex(['Order Date'], 8),
    deliveryDate: findColIndex(['Delivery Date'], 9),
    clientType: findColIndex(['Client Type'], 10),
    paymentTerms: findColIndex(['Payment Terms'], 11),
    paymentDays: findColIndex(['Payment Days'], 12),
    transportMode: findColIndex(['Transport Mode'], 13),
    transportName: findColIndex(['Transport Name'], 14),
    transporterNumber: findColIndex(['Transporter Number'], 15),
    vehicleNumber: findColIndex(['Vehicle Number'], 16),
    driverNumber: findColIndex(['Driver Number'], 17),
    productName: findColIndex(['Product Name', 'Item Name', 'Product'], 18),
    ream: findColIndex(['Ream Qty', 'Ream', 'Reams'], 19),
    box: findColIndex(['Box Qty', 'Box', 'Boxes'], 20),
    origRate: findColIndex(['Original Rate', 'Orig Rate'], 21),
    discount: findColIndex(['Discount', 'Flat Discount'], 22),
    netRate: findColIndex(['Net Rate', 'Rate'], 23),
    freight: findColIndex(['Freight (Item)', 'Freight'], 24),
    itemValue: findColIndex(['Item Value'], 25),
    itemValueFreight: findColIndex(['Item Value + Freight'], 26),
    gst: findColIndex(['GST (18%)', 'GST'], 27),
    itemGrandTotal: findColIndex(['Item Grand Total'], 28),
    remarks: findColIndex(['Remarks', 'Remark'], 29),
    filledBy: findColIndex(['Form Filler Name', 'Filled By (Name)', 'Form Filled By', 'Filled By', 'User'], 30),
    pdfUrl: findColIndex(['PDF URL', 'PDF'], 31),
    email: findColIndex(['eMail Address', 'Email Address', 'Email', 'eMail'], -1)
  };

  function getCell(row, idx) {
    if (idx === -1 || idx === undefined || !row || idx >= row.length) return '';
    return row[idx];
  }

  // Detects if a specific row has been shifted right by 1 column starting from Column H (index 7)
  function isRowShifted(r) {
    if (!r) return false;
    var spVal = (getCell(r, col.salesPerson) || '').toString().trim();
    var odVal = (getCell(r, col.orderDate) || '').toString().trim();
    if (!spVal && odVal) {
      if (/^(MR\.|MS\.|DA)/i.test(odVal) || SALES_PERSONS.indexOf(odVal.toUpperCase()) !== -1) {
        return true;
      }
      var ddVal = (getCell(r, col.deliveryDate) || '').toString().trim();
      if (ddVal && !/^\d{4}-\d{2}-\d{2}/.test(odVal) && !/^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(odVal)) {
        return true;
      }
    }
    return false;
  }

  function getRowVal(r, colIdx, shifted) {
    if (colIdx === -1 || colIdx === undefined || !r) return '';
    var actualIdx = (shifted && colIdx >= col.salesPerson) ? (colIdx + 1) : colIdx;
    return getCell(r, actualIdx);
  }

  // Distinct PI numbers in the sheet, most-recently-added first.
  var seen = {};
  var orderList = [];
  for (var i = 0; i < values.length; i++) {
    var fn = (getCell(values[i], col.formNo) || '').toString().trim();
    if (!fn || seen[fn]) continue;
    seen[fn] = {
      formNumber: fn,
      firmName: getCell(values[i], col.firmName) || '',
      customerName: getCell(values[i], col.customerName) || ''
    };
    orderList.push(fn);
  }
  orderList.reverse();

  var qUpper = query.toUpperCase();
  var formNumber = null;
  for (var k = 0; k < orderList.length; k++) {
    if (orderList[k].toUpperCase() === qUpper) { formNumber = orderList[k]; break; }
  }

  if (!formNumber) {
    var matches = orderList.filter(function(fn){ return fn.toUpperCase().indexOf(qUpper) !== -1; });
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
  var firstMatchedRowIndex = -1;
  for (var m = 0; m < values.length; m++) {
    var cellFn = (getCell(values[m], col.formNo) || '').toString().trim();
    if (cellFn === formNumber) {
      if (firstMatchedRowIndex === -1) {
        firstMatchedRowIndex = m + 2; // Row index in sheet (data starts at row 2)
      }
      matchedRows.push(values[m]);
    }
  }

  if (!matchedRows.length) {
    return { found: false, message: 'PI No. "' + formNumber + '" ke liye koi order nahi mila.' };
  }

  var first = matchedRows[0];
  var firstShifted = isRowShifted(first);

  var toISODate = function(v) {
    if (!v) return '';
    try {
      if (v instanceof Date) {
        if (isNaN(v.getTime())) return '';
        return Utilities.formatDate(v, getTZ(), 'yyyy-MM-dd');
      }
      var str = v.toString().trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.slice(0, 10);
      var parts = str.split('/');
      if (parts.length === 3) {
        var d = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
        if (!isNaN(d.getTime())) return Utilities.formatDate(d, getTZ(), 'yyyy-MM-dd');
      }
      var d2 = new Date(str);
      if (!isNaN(d2.getTime())) return Utilities.formatDate(d2, getTZ(), 'yyyy-MM-dd');
      return '';
    } catch (e) { return ''; }
  };

  var items = matchedRows.map(function(r) {
    var rShifted = isRowShifted(r);
    var origR = parseFloat(getRowVal(r, col.origRate, rShifted)) || 0;
    var netR = parseFloat(getRowVal(r, col.netRate, rShifted)) || 0;
    return {
      name: getRowVal(r, col.productName, rShifted) || '',
      ream: parseFloat(getRowVal(r, col.ream, rShifted)) || 0,
      box: parseFloat(getRowVal(r, col.box, rShifted)) || 0,
      origRate: origR || netR || 0,
      rate: netR || origR || 0,
      discount: parseFloat(getRowVal(r, col.discount, rShifted)) || 0
    };
  });

  var totalFreight = 0;
  if (col.freight !== -1) {
    matchedRows.forEach(function(r) {
      var rShifted = isRowShifted(r);
      totalFreight += parseFloat(getRowVal(r, col.freight, rShifted)) || 0;
    });
  }

  var tsVal = getRowVal(first, col.timestamp, firstShifted);
  var filledByVal = getRowVal(first, col.filledBy, firstShifted);

  var order = {
    rowNumber: firstMatchedRowIndex,
    itemCount: matchedRows.length,
    formNumber: formNumber,
    timestamp: (tsVal instanceof Date) ? tsVal.toISOString() : (tsVal || ''),
    firmName: getRowVal(first, col.firmName, firstShifted) || '',
    gstNo: getRowVal(first, col.gstNo, firstShifted) || '',
    areaName: getRowVal(first, col.area, firstShifted) || '',
    customerName: getRowVal(first, col.customerName, firstShifted) || '',
    whatsapp: getRowVal(first, col.whatsapp, firstShifted) || '',
    email: col.email !== -1 ? (getRowVal(first, col.email, firstShifted) || '') : '',
    salesPerson: getRowVal(first, col.salesPerson, firstShifted) || '',
    orderDateISO: toISODate(getRowVal(first, col.orderDate, firstShifted)),
    deliveryDateISO: toISODate(getRowVal(first, col.deliveryDate, firstShifted)),
    clientType: getRowVal(first, col.clientType, firstShifted) || '',
    paymentTerms: getRowVal(first, col.paymentTerms, firstShifted) || '',
    paymentDays: getRowVal(first, col.paymentDays, firstShifted) || '',
    transportMode: getRowVal(first, col.transportMode, firstShifted) || '',
    transportName: getRowVal(first, col.transportName, firstShifted) || '',
    transporterNumber: getRowVal(first, col.transporterNumber, firstShifted) || '',
    vehicleNumber: getRowVal(first, col.vehicleNumber, firstShifted) || '',
    driverNumber: getRowVal(first, col.driverNumber, firstShifted) || '',
    freightAmount: Math.round(totalFreight * 100) / 100,
    remarks: getRowVal(first, col.remarks, firstShifted) || '',
    filledBy: filledByVal || '',
    formFilledBy: filledByVal || '',
    filledByEmail: filledByVal || '',
    items: items
  };

  return { found: true, order: order };
}


function processEnquiry(formData, timestamp) {
  var headers = [
    'Timestamp', 'Form No', 'Firm Name', 'GST No', 'Area', 'Customer Name',
    'WhatsApp', 'Client Type', 'Client Source', 'Reference Name', 'Sales Person',
    'Product Name', 'Ream Qty', 'Box Qty', 'Rate',
    'Remarks', 'Visiting Card Link', 'Shop Photo Link', 'eMail Address', 'PDF URL'
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
  var fmEmail         = toAllCaps(formData.email || '');

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
  lines.push(isUpdate ? '≡ƒƒá *ORDER MODIFIED ΓÇö SATIJA PAPER*' : '≡ƒƒó *NEW ORDER ΓÇö SATIJA PAPER*');
  lines.push('ΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöü');
  lines.push('≡ƒôï *PO ID:* ' + (formData.formNumber || 'ΓÇö'));
  lines.push('≡ƒôà *ORDER DATE:* ' + (formData.orderDate || 'ΓÇö'));
  lines.push('≡ƒÜÜ *DELIVERY:* ' + (formData.deliveryDate || 'ΓÇö'));
  lines.push('');
  lines.push('≡ƒÅó *PARTY:* ' + toAllCaps(formData.firmName));
  lines.push('≡ƒôì *AREA:* ' + toAllCaps(formData.areaName));
  lines.push('≡ƒæñ *CUSTOMER:* ' + toAllCaps(formData.customerName));
  lines.push('≡ƒô▒ *MOBILE:* ' + (formData.whatsapp || 'ΓÇö'));
  lines.push('≡ƒæ¿ΓÇì≡ƒÆ╝ *SALES:* ' + toAllCaps(formData.salesPerson || 'ΓÇö'));
  lines.push('≡ƒÆ│ *PAYMENT:* ' + (
    toAllCaps(formData.paymentTerms) === 'CREDIT DAYS'
      ? (formData.paymentDays + ' DAYS CREDIT')
      : toAllCaps(formData.paymentTerms || 'ΓÇö')
  ));
  lines.push('');
  lines.push('≡ƒôª *ITEMS:*');

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
      ' | RATE: Γé╣' + Math.round(origR) +
      (disc > 0 ? ' [-Γé╣' + disc + ']' : '') +
      ' = Γé╣' + netAmt
    );
  });

  lines.push('');
  lines.push('ΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöü');
  if (totalDisc > 0) {
    lines.push('≡ƒÆ░ GROSS:    Γé╣' + Math.round(grossVal).toLocaleString());
    lines.push('≡ƒÄü SUPPORT: -Γé╣' + Math.round(totalDisc).toLocaleString());
    lines.push('   NET:      Γé╣' + Math.round(netVal).toLocaleString());
  } else {
    lines.push('≡ƒÆ░ NET VALUE: Γé╣' + Math.round(netVal).toLocaleString());
  }
  if (freight > 0) lines.push('≡ƒÜ¢ Freight:   Γé╣' + Math.round(freight).toLocaleString());
  lines.push('≡ƒº╛ GST (18%): Γé╣' + Math.round(gst).toLocaleString());
  lines.push('Γ£à *GRAND TOTAL: Γé╣' + Math.round(grandTotal).toLocaleString() + '*');

  if (formData.transportMode) {
    lines.push('');
    lines.push('≡ƒÜÜ *TRANSPORT:* ' + toAllCaps(formData.transportMode) +
      (formData.transportName ? ' ΓÇö ' + toAllCaps(formData.transportName) : '') +
      (formData.vehicleNumber ? ' | ' + formData.vehicleNumber.toUpperCase() : ''));
  }

  if (formData.remarks) {
    lines.push('');
    lines.push('≡ƒô¥ *REMARKS:* ' + toAllCaps(formData.remarks));
  }

  if (pdfUrl) {
    lines.push('');
    lines.push('≡ƒôä *PDF:* ' + pdfUrl);
  }

  lines.push('ΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöü');
  lines.push('_SATIJA PAPER | AUTO-GENERATED_');

  callMeBot(lines.join('\n'));
}


function sendWhatsAppEnquiry(formData) {
  var lines = [];
  lines.push('≡ƒö╡ *NEW ENQUIRY ΓÇö SATIJA PAPER*');
  lines.push('ΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöü');
  lines.push('≡ƒôï *ENQ ID:* ' + (formData.formNumber || 'ΓÇö'));
  lines.push('≡ƒÅó *PARTY:* ' + toAllCaps(formData.firmName));
  lines.push('≡ƒôì *AREA:* ' + toAllCaps(formData.areaName));
  lines.push('≡ƒæñ *CUSTOMER:* ' + toAllCaps(formData.customerName));
  lines.push('≡ƒô▒ *MOBILE:* ' + (formData.whatsapp || 'ΓÇö'));
  lines.push('≡ƒæ¿ΓÇì≡ƒÆ╝ *SALES:* ' + toAllCaps(formData.salesPerson || 'ΓÇö'));
  lines.push('≡ƒôî *SOURCE:* ' + toAllCaps(formData.clientSource || 'ΓÇö'));
  if (formData.referenceName) lines.push('≡ƒöù *REFERENCE:* ' + toAllCaps(formData.referenceName));
  lines.push('');
  lines.push('≡ƒôª *PRODUCTS ENQUIRED:*');
  (formData.items || []).forEach(function(item, i) {
    lines.push(
      (i + 1) + '. ' + toAllCaps(item.name || '') +
      ' ΓÇö ' + (parseFloat(item.ream) || 0) + ' REAM' +
      ' (' + Math.round(parseFloat(item.box) || 0) + ' BOX)'
    );
  });
  if (formData.remarks) {
    lines.push('');
    lines.push('≡ƒô¥ *REMARKS:* ' + toAllCaps(formData.remarks));
  }
  lines.push('ΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöüΓöü');
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
    Logger.log('UltraMsg config set nahi hai ΓÇö WhatsApp skip.');
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
    Logger.log('UltraMsg: ' + resp.getResponseCode() + ' ΓÇö ' + resp.getContentText().substring(0, 200));
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

    var isUpdate = (data.isUpdate === true || data.isUpdate === 'true' || data.action === 'updateOrder');

    if (data.formType === 'Order') {
      if (isUpdate) {
        var updateResult = updateOrder(data, new Date());
        return ContentService.createTextOutput(JSON.stringify(updateResult))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var orderResult = processOrder(data, new Date());
      return ContentService.createTextOutput(JSON.stringify(orderResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.formType === 'Enquiry' && data.items && data.items.length > 0) {
      var result = processForm(data);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Website enquiry webhook fallback
    if (data.name || data.whatsapp || data.requirements) {
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
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Ignored: No valid form or enquiry data' }))
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
  callMeBot('Γ£à TEST MESSAGE FROM SATIJA PAPER WMS!\nULTRAMSG SETUP SUCCESSFUL ≡ƒÄë\nAB ORDERS PE AUTO WHATSAPP AAYEGA!');
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


// Optional utility function to delete test/dummy entries if needed
function removeTestOrders(testFnos) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fnList = testFnos || ['PI-SP/26-27/00770', 'PI-SP/26-27/00771'];
  ['Orders', 'Orders_Summary'].forEach(function(sName) {
    var sheet = ss.getSheetByName(sName);
    if (!sheet) return;
    var lr = sheet.getLastRow();
    if (lr < 2) return;
    var vals = sheet.getRange(2, 2, lr - 1, 1).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var v = (vals[i][0] || '').toString().trim();
      if (fnList.indexOf(v) !== -1) {
        sheet.deleteRow(i + 2);
        Logger.log('Deleted ' + v + ' from ' + sName + ' row ' + (i + 2));
      }
    }
  });
}


// Utility to directly repair shifted rows (such as PO 773, 774) in the sheet.
// Can be run directly from Google Apps Script Editor with 1-click OR called via URL action=fixShiftedRows.
function fixShiftedRows773_774() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fixedDetails = [];
  var fixedSummaries = [];

  // 1. Fix in 'Orders' sheet
  var detailSheet = ss.getSheetByName('Orders');
  if (detailSheet) {
    var dLastRow = detailSheet.getLastRow();
    var dLastCol = Math.max(detailSheet.getLastColumn(), 33);
    if (dLastRow >= 2) {
      var dRange = detailSheet.getRange(2, 1, dLastRow - 1, dLastCol);
      var dValues = dRange.getValues();

      for (var i = 0; i < dValues.length; i++) {
        var row = dValues[i];
        var rowNum = i + 2;
        var formNo = (row[1] || '').toString().trim(); // Col B (Form No)
        var spVal = (row[7] || '').toString().trim();  // Col H (Sales Person)
        var odVal = (row[8] || '').toString().trim();  // Col I (Order Date)

        // Check if shifted: Col H is empty AND (matches 773/774 OR odVal looks like a person name / not date)
        var isTarget = formNo.includes('00773') || formNo.includes('00774') || /773|774/.test(formNo);
        var looksShifted = (!spVal && odVal && (/^(MR\.|MS\.|DA)/i.test(odVal) || SALES_PERSONS.indexOf(odVal.toUpperCase()) !== -1));

        if (isTarget || looksShifted) {
          // Columns 9..dLastCol shift left into Columns 8..dLastCol-1
          var updatedRow = row.slice();
          for (var c = 7; c < updatedRow.length - 1; c++) {
            updatedRow[c] = updatedRow[c + 1];
          }
          updatedRow[updatedRow.length - 1] = ''; // blank out the trailing column

          // Update columns 8 to dLastCol in this row
          var shiftLength = dLastCol - 7;
          var newSlice = [updatedRow.slice(7, dLastCol)];
          detailSheet.getRange(rowNum, 8, 1, shiftLength).setValues(newSlice);

          fixedDetails.push({ row: rowNum, formNo: formNo });
        }
      }
    }
  }

  // 2. Fix in 'Orders_Summary' sheet
  var summarySheet = ss.getSheetByName('Orders_Summary');
  if (summarySheet) {
    var sLastRow = summarySheet.getLastRow();
    var sLastCol = Math.max(summarySheet.getLastColumn(), 31);
    if (sLastRow >= 2) {
      var sRange = summarySheet.getRange(2, 1, sLastRow - 1, sLastCol);
      var sValues = sRange.getValues();

      for (var j = 0; j < sValues.length; j++) {
        var sRow = sValues[j];
        var sRowNum = j + 2;
        var sFormNo = (sRow[1] || '').toString().trim();
        var sSpVal = (sRow[7] || '').toString().trim();
        var sOdVal = (sRow[8] || '').toString().trim();

        var sIsTarget = sFormNo.includes('00773') || sFormNo.includes('00774') || /773|774/.test(sFormNo);
        var sLooksShifted = (!sSpVal && sOdVal && (/^(MR\.|MS\.|DA)/i.test(sOdVal) || SALES_PERSONS.indexOf(sOdVal.toUpperCase()) !== -1));

        if (sIsTarget || sLooksShifted) {
          var updatedSRow = sRow.slice();
          for (var sc = 7; sc < updatedSRow.length - 1; sc++) {
            updatedSRow[sc] = updatedSRow[sc + 1];
          }
          updatedSRow[updatedSRow.length - 1] = '';

          var sShiftLength = sLastCol - 7;
          var sNewSlice = [updatedSRow.slice(7, sLastCol)];
          summarySheet.getRange(sRowNum, 8, 1, sShiftLength).setValues(sNewSlice);

          fixedSummaries.push({ row: sRowNum, formNo: sFormNo });
        }
      }
    }
  }

  Logger.log('FIX RESULT: ' + JSON.stringify({ fixedDetails: fixedDetails, fixedSummaries: fixedSummaries }));
  return {
    success: true,
    message: 'Shifted rows (773, 774) fixed successfully!',
    fixedDetails: fixedDetails,
    fixedSummaries: fixedSummaries
  };
}
