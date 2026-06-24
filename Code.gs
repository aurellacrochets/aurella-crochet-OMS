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

// ═══════════════════════════════════════════════════════════════════
// AUTH SYSTEM
// Credentials are stored in Apps Script Properties (never in code).
//
// HOW TO SET UP USERS:
//   1. In Apps Script editor → Project Settings → Script Properties
//   2. Add a property named exactly:  auth_users
//   3. Value is a JSON array of user objects, e.g.:
//      [{"username":"alice","password":"mypass123"},{"username":"bob","password":"secure456"}]
//
// HOW TOKENS WORK:
//   - On login, Apps Script generates a random token and stores it
//     in Script Properties with key  token_<token>  and value  <expiry ISO string>
//   - Tokens expire after TOKEN_TTL_HOURS hours
//   - Every API call must send the token; it's validated server-side
// ═══════════════════════════════════════════════════════════════════

const TOKEN_TTL_HOURS = 8; // Token lifetime — adjust as needed

function getUsers() {
  const raw = PropertiesService.getScriptProperties().getProperty('auth_users');
  if (!raw) return [];
  try { return JSON.parse(raw); } catch(e) { return []; }
}

function validateToken(token) {
  if (!token) return false;
  const props = PropertiesService.getScriptProperties();
  const expiry = props.getProperty('token_' + token);
  if (!expiry) return false;
  if (new Date() > new Date(expiry)) {
    props.deleteProperty('token_' + token); // clean up expired token
    return false;
  }
  return true;
}

function login(username, password) {
  const users = getUsers();
  const user = users.find(u => u.username === username && u.password === password);
  if (!user) return { error: 'Invalid username or password' };

  // Generate a secure random token
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const expiry = new Date();
  expiry.setHours(expiry.getHours() + TOKEN_TTL_HOURS);

  PropertiesService.getScriptProperties().setProperty('token_' + token, expiry.toISOString());
  return { ok: true, token, expiresAt: expiry.toISOString(), username: user.username };
}

function logout(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('token_' + token);
  return { ok: true };
}

// ── GET router ───────────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action;

    // Login is the only unauthenticated GET (ping also allowed for API config test)
    if (action === 'ping') return corsOutput({ ok: true });

    // Auth check for all other GET requests
    if (!validateToken(e.parameter.token)) {
      return corsOutput({ error: 'Unauthorized', code: 401 });
    }

    if (action === 'getOrders')  return corsOutput({ data: getSheetData(ORDERS_SHEET) });
    if (action === 'getItems')   return corsOutput({ data: getSheetData(ITEMS_SHEET) });
    if (action === 'getLog')     return corsOutput({ data: getLog(e.parameter.orderId) });
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

    // Login is the only unauthenticated POST
    if (action === 'login')  return corsOutput(login(body.username, body.password));
    if (action === 'logout') return corsOutput(logout(body.token));

    // Auth check for all other POST requests
    if (!validateToken(body.token)) {
      return corsOutput({ error: 'Unauthorized', code: 401 });
    }

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
    order.Landmark, order.City||'', order.Pincode, order.DeliveryType, order.Courier||'',
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
  const updCol = headerIdx['UpdatedAt'];
  if (updCol) sheet.getRange(rowNum, updCol).setValue(new Date().toISOString());

  if (products) {
    const itemSheet = getSheet(ITEMS_SHEET);
    const itemData  = itemSheet.getDataRange().getValues();
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
  const { orders } = body;
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
  const { updates } = body;
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
