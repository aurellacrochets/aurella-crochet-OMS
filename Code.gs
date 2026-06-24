// ═══════════════════════════════════════════════════════════════════
// AURELLA OMS — Google Apps Script Backend
// Deploy as: Web App → Execute as: Me → Who has access: Anyone
// ═══════════════════════════════════════════════════════════════════

const SHEET_ID = '1Y_V9OawA9lys0Pw5XR0i5QcZGsU1T9BPkwsbGgGLehY';
const ORDERS_SHEET   = 'Orders';
const ITEMS_SHEET    = 'OrderItems';
const LOG_SHEET      = 'ActivityLog';

// ── CORS helper ─────────────────────────────────────────────────────
function corsOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── CORS preflight ───────────────────────────────────────────────────
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── GET router ───────────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'getOrders')  return corsOutput({ data: getSheetData(ORDERS_SHEET) });
    if (action === 'getItems')   return corsOutput({ data: getSheetData(ITEMS_SHEET) });
    if (action === 'getLog')     return corsOutput({ data: getLog(e.parameter.orderId) });
    if (action === 'ping')       return corsOutput({ ok: true });
    return corsOutput({ error: 'Unknown action' });
  } catch(err) {
    return corsOutput({ error: err.message });
  }
}

// ── POST router ──────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'createOrder') return corsOutput(createOrder(body));
    if (action === 'updateOrder') return corsOutput(updateOrder(body));
    if (action === 'bulkCreate')  return corsOutput(bulkCreate(body));
    if (action === 'bulkEdit')    return corsOutput(bulkEdit(body));

    return corsOutput({ error: 'Unknown action' });
  } catch(err) {
    return corsOutput({ error: err.message });
  }
}

// ── Sheet helpers ────────────────────────────────────────────────────
function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(sheet, name);
  }
  return sheet;
}

function initSheet(sheet, name) {
  if (name === ORDERS_SHEET) {
    sheet.appendRow(['OrderID','CustomerName','Phone','Address','Landmark','City','Pincode',
                     'DeliveryType','Courier','AWB','PaymentStatus','OrderStatus',
                     'CreatedAt','UpdatedAt']);
  }
  if (name === ITEMS_SHEET) {
    sheet.appendRow(['OrderID','ProductName','Quantity','Price']);
  }
  if (name === LOG_SHEET) {
    sheet.appendRow(['OrderID','Description','Timestamp']);
  }
}

function getSheetData(name) {
  const sheet = getSheet(name);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] !== undefined ? String(row[i]) : '');
    return obj;
  });
}

function findOrderRow(sheet, orderId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === orderId) return i + 1; // 1-indexed sheet row
  }
  return -1;
}

// ── createOrder ──────────────────────────────────────────────────────
function createOrder(body) {
  const { order, products } = body;
  const sheet = getSheet(ORDERS_SHEET);
  const now = new Date().toISOString();

  sheet.appendRow([
    order.OrderID, order.CustomerName, order.Phone, order.Address,
    order.Landmark||'', order.City||'', order.Pincode, order.DeliveryType, order.Courier||'',
    order.AWB||'', order.PaymentStatus, order.OrderStatus, now, now
  ]);

  const itemSheet = getSheet(ITEMS_SHEET);
  (products||[]).forEach(p => {
    itemSheet.appendRow([order.OrderID, p.ProductName, p.Quantity, p.Price||0]);
  });

  addLog(order.OrderID, `Order created — ${order.OrderStatus} / ${order.PaymentStatus}`);
  return { ok: true, orderId: order.OrderID };
}

// ── updateOrder ──────────────────────────────────────────────────────
function updateOrder(body) {
  const { orderId, fields, products } = body;
  const sheet = getSheet(ORDERS_SHEET);
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const rowNum  = findOrderRow(sheet, orderId);

  if (rowNum < 0) return { error: 'Order not found: ' + orderId };

  const fieldMap = {
    CustomerName:1, Phone:2, Address:3, Landmark:4, City:5, Pincode:6,
    DeliveryType:7, Courier:8, AWB:9, PaymentStatus:10, OrderStatus:11
  };
  // Headers are 0-indexed; columns start at 1. Adjust per actual sheet columns.
  const headerIdx = {};
  headers.forEach((h,i) => headerIdx[h] = i+1);

  const changes = [];
  Object.entries(fields||{}).forEach(([key, val]) => {
    const col = headerIdx[key];
    if (col) {
      sheet.getRange(rowNum, col).setValue(val);
      changes.push(`${key}→${val}`);
    }
  });
  // UpdatedAt
  const updCol = headerIdx['UpdatedAt'];
  if (updCol) sheet.getRange(rowNum, updCol).setValue(new Date().toISOString());

  // Update products if provided
  if (products) {
    const itemSheet = getSheet(ITEMS_SHEET);
    const itemData  = itemSheet.getDataRange().getValues();
    // Delete existing rows for this order (from bottom to avoid index shift)
    for (let i = itemData.length - 1; i >= 1; i--) {
      if (String(itemData[i][0]) === orderId) itemSheet.deleteRow(i+1);
    }
    products.forEach(p => itemSheet.appendRow([orderId, p.ProductName, p.Quantity, p.Price||0]));
  }

  if (changes.length) addLog(orderId, `Updated: ${changes.join(', ')}`);
  return { ok: true };
}

// ── bulkCreate ───────────────────────────────────────────────────────
function bulkCreate(body) {
  const { orders } = body; // array of { order, products }
  let created = 0, skipped = 0;
  const existing = getSheetData(ORDERS_SHEET).map(o => o.OrderID);

  orders.forEach(({ order, products }) => {
    if (existing.includes(order.OrderID)) { skipped++; return; }
    createOrder({ order, products });
    created++;
  });
  return { ok: true, created, skipped };
}

// ── bulkEdit ─────────────────────────────────────────────────────────
function bulkEdit(body) {
  const { updates } = body; // array of { orderId, fields }
  let updated = 0, notFound = 0;

  updates.forEach(({ orderId, fields }) => {
    const res = updateOrder({ orderId, fields });
    if (res.error) notFound++;
    else updated++;
  });
  return { ok: true, updated, notFound };
}

// ── Activity Log ─────────────────────────────────────────────────────
function addLog(orderId, description) {
  try {
    const sheet = getSheet(LOG_SHEET);
    sheet.appendRow([orderId, description, new Date().toISOString()]);
  } catch(e) {}
}

function getLog(orderId) {
  if (!orderId) return [];
  const data = getSheetData(LOG_SHEET);
  return data.filter(r => r.OrderID === orderId).reverse();
}
