/**
 * ======================================================
 *  DocuTracker - Consolidated Backend (Code.gs)
 * ======================================================
 *  Single-file backend for Google Apps Script.
 *  All pages are served from index.html as a SPA.
 * ======================================================
 */

const SHEET_ID            = '1msKYtUvpTsl4lo4a1RaUWqPJMDrw0RqKXzHdDOrdt70';
const DRIVE_FOLDER_ID     = '1ewUQh2MbRSg7ld-w1yEpfqdiRGTlPGw-';
const USERS_SHEET         = 'Users';
const DOCS_SHEET          = 'Documents';
const HISTORY_SHEET       = 'Document_History';
const CONFIG_SHEET        = 'Config';
const ACTIVITY_LOG_SHEET  = 'ActivityLog';
const OVERDUE_THRESHOLD   = 5;

// Per-request session token — set by doPost so internal helper functions
// (logActivity, logDocumentHistory, etc.) can call getCurrentUser() without
// an explicit token param. In native google.script.run calls the token is
// passed explicitly as a function argument.
let _doPostToken = '';

// =====================================================================
//  MAIN ENTRY POINT  –  All pages served from one index.html (SPA)
// =====================================================================
function doGet(e) {
  const scriptUrl = ScriptApp.getService().getUrl();
  // Capture URL parameters (e.g. ?token=xxx&page=reset) — GAS does not expose
  // query-string params to the browser via window.location, so we inject them
  // into the HTML as a JavaScript variable that the SPA reads on load.
  const params    = (e && e.parameter) ? e.parameter : {};
  const token     = params.token     || '';
  const page      = params.page      || '';

  // Build a small inline script that seeds the SPA's URL-parameter state
  const injected  = `<script id="__gas_init__">
    window.__GAS_SCRIPT_URL__ = ${JSON.stringify(scriptUrl)};
    window.__GAS_PARAMS__     = ${JSON.stringify({ token: token, page: page })};
  <\/script>`;

  let html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('DocuTracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  // Inject the init script right after <head> so it runs before any app code
  const content = html.getContent().replace('<head>', '<head>' + injected);
  html.setContent(content);
  return html;
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// =====================================================================
//  REST API ENDPOINT — handles JSON POST calls from standalone HTML
//  (i.e., when index.html is opened directly, not via doGet)
// =====================================================================
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const body   = JSON.parse(e.postData.contents);
    const method = body.method;
    const args   = body.args || [];

    // Set the module-level token so internal helpers (getCurrentUser, logActivity, etc.)
    // work without needing an explicit token argument during this execution.
    _doPostToken = body.sessionToken || body.sessionEmail || '';

    const allowedMethods = [
      'checkLogin', 'logout', 'getCurrentUser', 'updateUserProfile',
      'sendResetEmail', 'validateResetToken', 'setNewPassword',
      'getDropdownOptions', 'updateDropdownOptions',
      'getAllDocuments', 'getDocumentById', 'getDocumentStats',
      'addDocument', 'updateDocument', 'deleteDocument',
      'getDocumentHistory', 'logDocumentHistory',
      'getAllActivityLogs', 'logActivity',
      'uploadPDFToGoogleDrive', 'getScriptUrl',
      'getUsers', 'addUser', 'updateUser', 'toggleUserStatus',
      'getInitialData'
    ];

    if (!allowedMethods.includes(method)) {
      output.setContent(JSON.stringify({ error: 'Method not allowed: ' + method }));
      return output;
    }

    const fn = this[method] || global[method];
    if (typeof fn !== 'function') {
      output.setContent(JSON.stringify({ error: 'Unknown method: ' + method }));
      return output;
    }

    const result = fn.apply(null, args);
    output.setContent(JSON.stringify({ result: result }));
  } catch (err) {
    output.setContent(JSON.stringify({ error: err.toString() }));
  }

  return output;
}

// =====================================================================
//  CONFIG / DROPDOWN OPTIONS
// =====================================================================
function initializeConfigSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let configSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG_SHEET);
    configSheet.getRange(1, 1, 1, 5).setValues([['Document Types', 'Suppliers', 'Offices', 'Status Options', 'EndUser']]);
    configSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    const defaultDocTypes = ['Purchase Request','Purchase Order','Notice to Proceed',
      'Notice of Award','COA AOM','Memos','DTR','Leave','PAR','PRS','Clearance'];
    const defaultOffices  = ['GSO-Supp','GSO-Admin','GSO-Rec','BAC','CADMIN',
      'CMO','PNP','CBO','CACCTO','CAGRIO','COA','CCRO','CPDO','CECON','CVET'];
    const defaultStatuses = ['Received','Incoming','In Review','In Process',
      'Approved','Forwarded','Hold','Completed'];
    const defaultEndUsers = ['GSO','BAC','CADMIN','CMO','CBO','CACCTO','COA','CPDO','CECON'];
    defaultDocTypes.forEach((v, i) => configSheet.getRange(i + 2, 1).setValue(v));
    defaultOffices.forEach((v, i)  => configSheet.getRange(i + 2, 3).setValue(v));
    defaultStatuses.forEach((v, i) => configSheet.getRange(i + 2, 4).setValue(v));
    defaultEndUsers.forEach((v, i) => configSheet.getRange(i + 2, 5).setValue(v));
    Logger.log('Config sheet created with defaults');
  } else {
    // Ensure EndUser header exists in col 5 for existing sheets
    const lastCol = configSheet.getLastColumn();
    if (lastCol < 5) {
      configSheet.getRange(1, 5).setValue('EndUser');
      configSheet.getRange(1, 5).setFontWeight('bold');
    } else {
      const headerVal = configSheet.getRange(1, 5).getValue();
      if (!headerVal || headerVal.toString().trim() === '') {
        configSheet.getRange(1, 5).setValue('EndUser');
        configSheet.getRange(1, 5).setFontWeight('bold');
      }
    }
  }
  return configSheet;
}

function getDropdownOptions() {
  try {
    const cached = _cacheGet(CACHE_OPTS);
    if (cached) return cached;
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let configSheet = ss.getSheetByName(CONFIG_SHEET);
    if (!configSheet) configSheet = initializeConfigSheet();
    // Ensure EndUser column exists without a redundant extra call
    const lastCol = configSheet.getLastColumn();
    if (lastCol < 5 || !configSheet.getRange(1, 5).getValue()) {
      configSheet.getRange(1, 5).setValue('EndUser').setFontWeight('bold');
    }
    const data = configSheet.getDataRange().getValues();
    const options = { docTypes: [], suppliers: [], offices: [], statuses: [], endUsers: [] };
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) options.docTypes.push(String(data[i][0]).trim());
      if (data[i][1]) options.suppliers.push(String(data[i][1]).trim());
      if (data[i][2]) options.offices.push(String(data[i][2]).trim());
      if (data[i][3]) options.statuses.push(String(data[i][3]).trim());
      if (data[i][4]) options.endUsers.push(String(data[i][4]).trim());
    }
    _cachePut(CACHE_OPTS, options);
    return options;
  } catch (error) {
    Logger.log('getDropdownOptions error: ' + error);
    return { docTypes: ['Purchase Request'], suppliers: [], offices: ['GSO-Admin'], statuses: ['Received'], endUsers: [] };
  }
}

function updateDropdownOptions(type, values) {
  try {
    const currentUser = getCurrentUser();
    if (!currentUser || currentUser.role !== 'Admin') {
      return { status: 'error', message: 'Only administrators can update dropdown options' };
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let configSheet = ss.getSheetByName(CONFIG_SHEET) || initializeConfigSheet();
    const columnMap = { docTypes: 1, suppliers: 2, offices: 3, statuses: 4, endUsers: 5 };
    const column = columnMap[type];
    if (!column) return { status: 'error', message: 'Invalid type' };
    const lastRow = configSheet.getLastRow();
    if (lastRow > 1) configSheet.getRange(2, column, lastRow - 1, 1).clearContent();
    values.forEach((v, i) => { if (v.trim()) configSheet.getRange(i + 2, column).setValue(v.trim()); });
    return { status: 'success', message: 'Dropdown options updated successfully' };
  } catch (error) {
    Logger.log('updateDropdownOptions error: ' + error);
    return { status: 'error', message: error.toString() };
  }
}

// =====================================================================
//  SESSION HELPERS  (token-based, per-browser, multi-user safe)
//  Each browser stores its own UUID in localStorage.
//  ScriptProperties key: 'sess_<UUID>' → JSON { email, expires }
// =====================================================================
function _createSession(email) {
  const token   = Utilities.getUuid();
  const expires = Date.now() + 24 * 60 * 60 * 1000; // 24-hour TTL
  PropertiesService.getScriptProperties().setProperty(
    'sess_' + token, JSON.stringify({ email: email.toLowerCase().trim(), expires: expires })
  );
  return token;
}

function _validateSession(token) {
  if (!token) return null;
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('sess_' + token);
    if (!raw) return null;
    const sess = JSON.parse(raw);
    if (sess.expires < Date.now()) {
      PropertiesService.getScriptProperties().deleteProperty('sess_' + token);
      return null;
    }
    return sess.email;
  } catch (e) { return null; }
}

function _destroySession(token) {
  if (!token) return;
  try { PropertiesService.getScriptProperties().deleteProperty('sess_' + token); } catch (e) {}
}

function _getDefaultPermissions(role) {
  const r = (role || 'Staff').toLowerCase();
  if (r === 'admin')   return { addDoc: true,  editDoc: true,  deleteDoc: true,  viewDoc: true, printExport: true,  manageSettings: true,  manageUsers: true,  viewAnalytics: true,  trackHistory: true  };
  if (r === 'manager') return { addDoc: true,  editDoc: true,  deleteDoc: false, viewDoc: true, printExport: true,  manageSettings: false, manageUsers: false, viewAnalytics: true,  trackHistory: true  };
  return                       { addDoc: true,  editDoc: true,  deleteDoc: false, viewDoc: true, printExport: false, manageSettings: false, manageUsers: false, viewAnalytics: false, trackHistory: true  };
}

function _ensureUsersColumns(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    const cols = [ { idx: 6, name: 'Status' }, { idx: 7, name: 'Permissions' },
                   { idx: 8, name: 'CreatedAt' }, { idx: 9, name: 'LastLogin' } ];
    cols.forEach(c => {
      if (!headers[c.idx] || headers[c.idx].toString().trim() === '') {
        sheet.getRange(1, c.idx + 1).setValue(c.name).setFontWeight('bold');
        if (c.name === 'Status') {
          const lr = sheet.getLastRow();
          if (lr > 1) {
            const rng = sheet.getRange(2, c.idx + 1, lr - 1, 1);
            rng.setValues(rng.getValues().map(r => [r[0] || 'Active']));
          }
        }
      }
    });
  } catch (e) { Logger.log('_ensureUsersColumns: ' + e); }
}

// =====================================================================
//  USER SESSION MANAGEMENT
// =====================================================================
function getCurrentUser(tokenParam) {
  try {
    const token = tokenParam || _doPostToken;
    const email = _validateSession(token);
    if (!email) return null;
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(USERS_SHEET);
    if (!sheet) return null;
    _ensureUsersColumns(sheet);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() !== email) continue;
      const role   = String(data[i][2] || 'Staff');
      const status = String(data[i][6] || 'Active');
      if (status.toLowerCase() === 'inactive') return null;
      const perms  = _getDefaultPermissions(role);
      try { const c = data[i][7] ? JSON.parse(data[i][7]) : null; if (c) Object.assign(perms, c); } catch (e) {}
      return { email: data[i][0], name: data[i][3] || 'User', role: role, status: status, permissions: perms };
    }
    return null;
  } catch (error) { Logger.log('getCurrentUser error: ' + error); return null; }
}

function logout(tokenParam) {
  try {
    const token = tokenParam || _doPostToken;
    const email = _validateSession(token);
    if (email) logActivity('Logout', '', `User logged out: ${email}`);
    _destroySession(token);
  } catch (error) { Logger.log('logout error: ' + error); }
  try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; }
}

function updateUserProfile(tokenParam, updates) {
  try {
    const token = tokenParam || _doPostToken;
    const email = _validateSession(token);
    if (!email) return { status: 'error', message: 'No user session found' };
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(USERS_SHEET);
    if (!sheet) return { status: 'error', message: 'Users sheet not found' };
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() !== email.toLowerCase()) continue;
      if (updates.name)     sheet.getRange(i + 1, 4).setValue(updates.name);
      if (updates.password) sheet.getRange(i + 1, 2).setValue(updates.password);
      return { status: 'success', message: 'Profile updated successfully' };
    }
    return { status: 'error', message: 'User not found' };
  } catch (error) { Logger.log('updateUserProfile error: ' + error); return { status: 'error', message: error.toString() }; }
}

// =====================================================================
//  AUTHENTICATION
// =====================================================================
function checkLogin(email, password) {
  if (!email || !password) return { status: 'invalid', message: 'Please enter email and password' };
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) return { status: 'invalid', message: 'System error' };
  _ensureUsersColumns(sheet);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const userEmail = String(data[i][0]).trim().toLowerCase();
    const userPass  = String(data[i][1]).trim();
    if (email.toLowerCase() !== userEmail || password !== userPass) continue;
    const status = String(data[i][6] || 'Active');
    if (status.toLowerCase() === 'inactive')
      return { status: 'invalid', message: 'Account is deactivated. Contact your administrator.' };
    const token = _createSession(userEmail);
    logActivity('Login', '', `User logged in: ${userEmail}`);
    return { status: 'success', token: token };
  }
  return { status: 'invalid', message: 'Invalid email or password' };
}

// =====================================================================
//  PASSWORD RESET
// =====================================================================
function sendResetEmail(email) {
  if (!email) return 'notfound';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) return 'notfound';
  const data   = sheet.getDataRange().getValues();
  const token  = Utilities.getUuid();
  const expiry = new Date(Date.now() + 1000 * 60 * 30); // 30-minute expiry

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase()) {
      // Persist token + expiry into Users sheet (columns E & F)
      sheet.getRange(i + 1, 5).setValue(token);
      sheet.getRange(i + 1, 6).setValue(expiry);

      // Build the reset link using the CURRENT deployed Web App URL.
      // We pass both `page=reset` (so doGet knows to show the reset view)
      // and `token` (for validation). GAS injects these via __GAS_PARAMS__.
      const baseUrl    = ScriptApp.getService().getUrl();
      const resetLink  = baseUrl + '?page=reset&token=' + encodeURIComponent(token);

      try {
        MailApp.sendEmail({
          to: email,
          subject: 'DocuTracker — Password Reset Request',
          htmlBody: [
            '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:24px;border:1px solid #e2e8f0;border-radius:8px">',
            '<h2 style="margin:0 0 8px;color:#0f172a">&#128196; DocuTracker</h2>',
            '<h3 style="margin:0 0 16px;color:#374151">Password Reset Request</h3>',
            '<p style="color:#4b5563;font-size:14px">A password reset was requested for your account.</p>',
            '<a href="' + resetLink + '" style="display:inline-block;margin:16px 0;background:#0f172a;color:#fff;',
            'padding:12px 24px;border-radius:6px;text-decoration:none;font-weight:600;font-size:14px">',
            'Reset My Password</a>',
            '<p style="color:#6b7280;font-size:13px">This link expires in <strong>30 minutes</strong>.</p>',
            '<p style="color:#6b7280;font-size:13px">If you did not request this, you can safely ignore this email.</p>',
            '<hr style="border:none;border-top:1px solid #e2e8f0;margin:16px 0">',
            '<p style="color:#9ca3af;font-size:12px">DocuTracker System</p>',
            '</div>'
          ].join(''),
          body: 'Hello,\n\nA password reset was requested for your DocuTracker account.\n\nReset link:\n' + resetLink + '\n\nThis link is valid for 30 minutes. If you did not request this, ignore this email.\n\nDocuTracker System'
        });
      } catch (error) {
        Logger.log('sendResetEmail — MailApp error: ' + error);
        return 'error';
      }
      return 'sent';
    }
  }
  return 'notfound';
}

function validateResetToken(token) {
  if (!token) return false;
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === token && new Date(data[i][5]) > new Date()) return true;
  }
  return false;
}

function setNewPassword(token, newPassword) {
  if (!token || !newPassword) return 'failed';
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === token) {
      sheet.getRange(i + 1, 2).setValue(newPassword);
      sheet.getRange(i + 1, 5, 1, 2).clearContent();
      return 'success';
    }
  }
  return 'failed';
}

// =====================================================================
//  DOCUMENT HISTORY
// =====================================================================
function initializeHistorySheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let historySheet = ss.getSheetByName(HISTORY_SHEET);
  if (!historySheet) {
    historySheet = ss.insertSheet(HISTORY_SHEET);
    historySheet.getRange(1, 1, 1, 6).setValues([['Document ID','Action','Status','User','Timestamp','Remarks']]);
    historySheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return historySheet;
}

function logDocumentHistory(docId, action, status, user, remarks) {
  try {
    const historySheet = initializeHistorySheet();
    if (!user || user === 'System') {
      const cu = getCurrentUser();
      if (cu) user = cu.name;
    }
    historySheet.appendRow([docId, action, status || '', user || 'System', new Date(), remarks || '']);
    return true;
  } catch (error) {
    Logger.log('logDocumentHistory error: ' + error);
    return false;
  }
}

function getDocumentHistory(docId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const historySheet = ss.getSheetByName(HISTORY_SHEET);
    if (!historySheet) return [];
    const data    = historySheet.getDataRange().getValues();
    const history = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === docId) {
        history.push({
          documentId: data[i][0],
          action:     data[i][1],
          status:     data[i][2],
          user:       data[i][3],
          timestamp:  Utilities.formatDate(new Date(data[i][4]), Session.getScriptTimeZone(), 'MMM dd, yyyy hh:mm a'),
          remarks:    data[i][5]
        });
      }
    }
    history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    return history;
  } catch (error) {
    Logger.log('getDocumentHistory error: ' + error);
    return [];
  }
}

function getDocumentById(docId) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(DOCS_SHEET);
    if (!sheet) return null;
    const data    = sheet.getDataRange().getValues();
    if (data.length <= 1) return null;
    const headers = data[0].map(h => h.toString().trim());
    const idCol   = headers.indexOf('ID/Barcode');
    if (idCol === -1) return null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === String(docId).trim()) {
        const doc = {};
        for (let j = 0; j < headers.length; j++) {
          const value = data[i][j];
          const h = headers[j];
          if (value instanceof Date) {
            if (h === 'Doc Time Stamp' || h === 'PO Time Stamp') {
              doc[h] = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
            } else {
              doc[h] = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            }
          } else {
            doc[h] = value;
          }
        }
        return doc;
      }
    }
    return null;
  } catch (error) {
    Logger.log('getDocumentById error: ' + error);
    return null;
  }
}

// =====================================================================
//  ACTIVITY LOG
// =====================================================================
function initializeActivityLogSheet() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let logSheet = ss.getSheetByName(ACTIVITY_LOG_SHEET);
    if (!logSheet) {
      logSheet = ss.insertSheet(ACTIVITY_LOG_SHEET);
      logSheet.getRange(1, 1, 1, 6).setValues([['Timestamp','Action','User','Document ID','Details','IP Address']]);
      logSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      logSheet.getRange(1, 1, 1, 6).setBackground('#f3f4f6');
      logSheet.setFrozenRows(1);
    }
    return logSheet;
  } catch (error) {
    Logger.log('initializeActivityLogSheet error: ' + error);
    return null;
  }
}

function logActivity(action, documentId, details) {
  try {
    const logSheet = initializeActivityLogSheet();
    if (!logSheet) return false;
    const cu       = getCurrentUser();
    const userName = cu ? cu.name : 'System';
    logSheet.appendRow([
      new Date(),
      String(action || '').trim(),
      userName,
      String(documentId || '').trim(),
      String(details || '').trim(),
      'N/A'
    ]);
    return true;
  } catch (error) {
    Logger.log('logActivity error: ' + error);
    return false;
  }
}

function getAllActivityLogs() {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(ACTIVITY_LOG_SHEET);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const logs = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      logs.push({
        Timestamp:  Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'MMM dd, yyyy hh:mm a'),
        Action:     data[i][1] || '',
        User:       data[i][2] || '',
        DocumentID: data[i][3] || '',
        Details:    data[i][4] || '',
        IPAddress:  data[i][5] || 'N/A'
      });
    }
    logs.sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));
    return logs;
  } catch (error) {
    Logger.log('getAllActivityLogs error: ' + error);
    return [];
  }
}

// =====================================================================
//  DOCUMENT HELPERS
// =====================================================================
function calculateOverdueStatus(dateReceived, status) {
  const s = String(status || '').toLowerCase().trim();
  if (s.includes('forwarded') || s.includes('completed') || s.includes('complete')) return 'On time';
  const received = new Date(dateReceived);
  const dueDate  = new Date(received);
  dueDate.setDate(dueDate.getDate() + OVERDUE_THRESHOLD);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  dueDate.setHours(0, 0, 0, 0);
  if (today > dueDate) {
    return `${Math.floor((today - dueDate) / 86400000)} days`;
  }
  return 'On time';
}

function formatAmount(amount) {
  if (!amount || amount === '') return '';
  const num = parseFloat(amount.toString().replace(/[^\d.-]/g, ''));
  if (isNaN(num)) return amount;
  return '₱' + num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function checkDuplicateDocNo(docNo, excludeId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DOCS_SHEET);
    if (!sheet) return false;
    const data    = sheet.getDataRange().getValues();
    if (data.length <= 1) return false;
    const headers = data[0].map(h => h.toString().trim());
    const idCol   = headers.indexOf('ID/Barcode');
    const noCol   = headers.indexOf('Doc No');
    if (noCol === -1) return false;
    for (let i = 1; i < data.length; i++) {
      const rowId = idCol !== -1 ? data[i][idCol] : data[i][0];
      if (data[i][noCol] === docNo && rowId !== excludeId) return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}

// Fast variant of logActivity — accepts pre-fetched userName to avoid a
// redundant getCurrentUser() call inside the same request.
function _logActivityFast(action, documentId, details, userName) {
  try {
    const logSheet = initializeActivityLogSheet();
    if (!logSheet) return;
    logSheet.appendRow([
      new Date(),
      String(action || '').trim(),
      userName || 'System',
      String(documentId || '').trim(),
      String(details || '').trim(),
      'N/A'
    ]);
  } catch(e) { Logger.log('_logActivityFast error: ' + e); }
}

// Fast variant of logDocumentHistory — accepts pre-fetched userName.
function _logDocHistoryFast(docId, action, status, userName, remarks) {
  try {
    const histSheet = initializeHistorySheet();
    if (!histSheet) return;
    histSheet.appendRow([docId, action, status || '', userName || 'System', new Date(), remarks || '']);
  } catch(e) { Logger.log('_logDocHistoryFast error: ' + e); }
}

// =====================================================================
//  SCHEMA MIGRATION HELPER
//  Ensures the Documents sheet has all 20 columns in the correct order.
//  Safe to run on both new and existing sheets.
// =====================================================================
// Column count constant — update here only when schema changes
const DOCS_COL_COUNT = 21;
const EXPECTED_HEADERS = [
  'Doc Time Stamp','ID/Barcode','Doc Type','Doc No','PR Date','Description','Amount',
  'EndUser','PO Time Stamp','PO No','PO Date','PO Amount',
  'Supplier','Requisitioner','Endorsed To','Status',
  'Date Received','Due Date','Overdue','Notes','PDF Link'
];

function ensureDocumentSheetHeaders(sheet) {
  try {
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) {
      sheet.getRange(1, 1, 1, DOCS_COL_COUNT).setValues([EXPECTED_HEADERS]);
      sheet.getRange(1, 1, 1, DOCS_COL_COUNT).setFontWeight('bold');
      return;
    }
    const current = data[0].map(h => h.toString().trim());
    const needsUpdate = current.length < DOCS_COL_COUNT ||
      EXPECTED_HEADERS.some((h, i) => current[i] !== h);
    if (needsUpdate) {
      sheet.getRange(1, 1, 1, DOCS_COL_COUNT).setValues([EXPECTED_HEADERS]);
      sheet.getRange(1, 1, 1, DOCS_COL_COUNT).setFontWeight('bold');
      Logger.log('Document sheet headers normalized to expected schema.');
    }
  } catch (e) {
    Logger.log('ensureDocumentSheetHeaders error: ' + e);
  }
}

// =====================================================================
//  CACHE HELPERS  — 5-minute CacheService invalidated on writes
// =====================================================================
const CACHE_TTL     = 300; // seconds
const CACHE_DOCS    = 'dt_docs_v2';
const CACHE_OPTS    = 'dt_opts_v1';
const CACHE_STATS   = 'dt_stats_v1';

function _cacheGet(key) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch(e) { return null; }
}
function _cachePut(key, value) {
  try {
    const raw = JSON.stringify(value);
    if (raw.length < 100000) CacheService.getScriptCache().put(key, raw, CACHE_TTL);
  } catch(e) {}
}
function _cacheInvalidateDocs() {
  try { CacheService.getScriptCache().removeAll([CACHE_DOCS, CACHE_STATS]); } catch(e) {}
}

// =====================================================================
//  BATCH LOADER — single round-trip for dashboard init
//  The sessionToken arg is sent by the frontend; doPost already put it
//  in _doPostToken, so getCurrentUser() picks it up automatically.
//  We still accept (and optionally override) the arg for flexibility.
// =====================================================================
function getInitialData(sessionToken) {
  try {
    // Override the module-level token if an arg was passed directly
    if (sessionToken) _doPostToken = sessionToken;
    const user  = getCurrentUser();
    // Only continue loading data if the session is valid
    if (!user) return { user: null, opts: {}, docs: [], stats: {} };
    const opts  = getDropdownOptions();
    const docs  = getAllDocuments();
    const stats = getDocumentStats();
    return { user: user, opts: opts, docs: docs, stats: stats };
  } catch(e) {
    Logger.log('getInitialData error: ' + e);
    return { user: null, opts: {}, docs: [], stats: {} };
  }
}

// =====================================================================
//  DOCUMENT CRUD
// =====================================================================
function getAllDocuments() {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(DOCS_SHEET);
    if (!sheet) return [];
    ensureDocumentSheetHeaders(sheet);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const headers    = data[0].map(h => h.toString().trim());
    const idCol      = headers.indexOf('ID/Barcode');
    const noCol      = headers.indexOf('Doc No');
    const dateRcvdCol= headers.indexOf('Date Received');
    const statusCol  = headers.indexOf('Status');
    const documents  = [];
    const seenDocNos = new Set();
    for (let i = 1; i < data.length; i++) {
      const rowId = idCol !== -1 ? data[i][idCol] : data[i][0];
      if (!rowId || rowId.toString().trim() === '') continue;
      const docNo = noCol !== -1 ? data[i][noCol] : data[i][2];
      if (seenDocNos.has(docNo)) continue;
      seenDocNos.add(docNo);
      const doc = {};
      for (let j = 0; j < headers.length; j++) {
        const value = data[i][j];
        // Timestamp fields: format with time; date-only fields: format date only
        if (value instanceof Date) {
          const h = headers[j];
          if (h === 'Doc Time Stamp' || h === 'PO Time Stamp') {
            doc[h] = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
          } else {
            doc[h] = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }
        } else {
          doc[headers[j]] = value;
        }
      }
      const drVal = dateRcvdCol !== -1 ? doc['Date Received'] : doc['Date Received'];
      const stVal = statusCol   !== -1 ? doc['Status']        : doc['Status'];
      if (drVal) {
        doc['Overdue'] = calculateOverdueStatus(drVal, stVal);
      }
      documents.push(doc);
    }
    return documents;
  } catch (error) {
    Logger.log('getAllDocuments error: ' + error);
    return [];
  }
}

function addDocument(docData) {
  try {
    const ss  = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(DOCS_SHEET);
    const tz  = Session.getScriptTimeZone();

    // ── ONE read covers: header-check + duplicate-check ────────────────
    let existingData = null;
    if (!sheet) {
      sheet = ss.insertSheet(DOCS_SHEET);
      sheet.getRange(1, 1, 1, DOCS_COL_COUNT).setValues([EXPECTED_HEADERS]).setFontWeight('bold');
    } else {
      existingData = sheet.getDataRange().getValues();
      // Fix headers if needed (uses the already-loaded data, no extra read)
      if (existingData.length > 0) {
        const curr = existingData[0].map(h => h.toString().trim());
        if (curr.length < DOCS_COL_COUNT || EXPECTED_HEADERS.some((h, i) => curr[i] !== h)) {
          sheet.getRange(1, 1, 1, DOCS_COL_COUNT).setValues([EXPECTED_HEADERS]).setFontWeight('bold');
        }
      }
      // Duplicate check from the same data (NO extra Sheets read)
      if (existingData.length > 1) {
        const hdr   = existingData[0].map(h => h.toString().trim());
        const noIdx = hdr.indexOf('Doc No');
        if (noIdx !== -1) {
          const target = (docData.docNo || '').trim();
          for (let i = 1; i < existingData.length; i++) {
            if (String(existingData[i][noIdx] || '').trim() === target) {
              return { status: 'error', message: 'duplicate', docNo: docData.docNo };
            }
          }
        }
      }
    }

    // ── Build new row ────────────────────────────────────────────────────
    const now    = new Date();
    const docId  = `DOC-${Utilities.formatDate(now, tz, 'yyyyMMddHHmmss')}-${String(Math.floor(Math.random()*10000)).padStart(4,'0')}`;
    const dateRcvd = new Date(docData.dateReceived || now);
    const dueDate  = new Date(dateRcvd);
    dueDate.setDate(dueDate.getDate() + OVERDUE_THRESHOLD);
    const tsStr  = Utilities.formatDate(now,      tz, 'yyyy-MM-dd HH:mm:ss');
    const drStr  = Utilities.formatDate(dateRcvd, tz, 'yyyy-MM-dd');
    const ddStr  = Utilities.formatDate(dueDate,  tz, 'yyyy-MM-dd');
    const overdue = calculateOverdueStatus(dateRcvd, docData.status);

    const newRow = [
      tsStr,
      docId,
      docData.docType       || '',
      docData.docNo         || '',
      docData.prDate        || '',
      docData.description   || '',
      formatAmount(docData.amount),
      docData.endUser       || '',
      tsStr,
      docData.poNo          || '',
      docData.poDate        || '',
      formatAmount(docData.poAmount),
      docData.supplier      || '',
      docData.requisitioner || '',
      docData.office        || '',
      docData.status        || 'Received',
      drStr, ddStr, overdue,
      docData.notes         || '',
      docData.pdfLink       || ''
    ];

    // ── ONE write (setValues on specific row is reliable for large sheets) ─
    const nextRow = existingData ? existingData.length + 1 : 2;
    sheet.getRange(nextRow, 1, 1, DOCS_COL_COUNT).setValues([newRow]);
    _cacheInvalidateDocs();

    // ── Build doc object for optimistic frontend update ──────────────────
    const docObj = {};
    EXPECTED_HEADERS.forEach((h, i) => { docObj[h] = newRow[i]; });
    docObj['Overdue'] = overdue;

    // ── Logging (reuse already-fetched user; no extra getCurrentUser call) ─
    const cu      = getCurrentUser();
    const cuName  = cu ? cu.name : 'System';
    const remarks = [
      `${docData.docType} — ${docData.docNo}`,
      docData.office   ? `Endorsed To: ${docData.office}`   : '',
      docData.endUser  ? `End User: ${docData.endUser}`     : '',
      docData.prDate   ? `PR Date: ${docData.prDate}`       : ''
    ].filter(Boolean).join(' | ');
    _logDocHistoryFast(docId, 'Document Created', docData.status || 'Received', cuName, remarks);
    _logActivityFast('Create Document', docId, `Document created: ${docData.docType} - ${docData.docNo}`, cuName);

    return { status: 'success', docId: docId, doc: docObj };
  } catch (error) {
    Logger.log('addDocument error: ' + error);
    return { status: 'error', message: error.toString() };
  }
}

function updateDocument(docId, docData) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(DOCS_SHEET);
    if (!sheet) return { status: 'error', message: 'Documents sheet not found' };
    const tz    = Session.getScriptTimeZone();

    // ── ONE read covers: header-check + duplicate-check + row-find ──────
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());

    // Fix headers if needed (inline — no extra read)
    if (headers.length < DOCS_COL_COUNT || EXPECTED_HEADERS.some((h, i) => headers[i] !== h)) {
      sheet.getRange(1, 1, 1, DOCS_COL_COUNT).setValues([EXPECTED_HEADERS]).setFontWeight('bold');
    }

    const colIdx   = {};
    headers.forEach((h, i) => { colIdx[h] = i; });
    const idColIdx = colIdx['ID/Barcode'] !== undefined ? colIdx['ID/Barcode'] : 0;
    const noColIdx = colIdx['Doc No']     !== undefined ? colIdx['Doc No']     : 2;
    const targetNo = (docData.docNo || '').trim();

    // ── Single-pass: duplicate check AND row-find simultaneously ────────
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][idColIdx] || '').trim();
      const rowNo = String(data[i][noColIdx] || '').trim();
      // Duplicate: same Doc No but different ID
      if (rowNo === targetNo && rowId !== docId) {
        return { status: 'error', message: 'duplicate', docNo: docData.docNo };
      }
      if (rowId === docId) rowIndex = i + 1; // 1-indexed sheet row
    }
    if (rowIndex === -1) return { status: 'error', message: 'Document not found' };

    // ── Compute values ──────────────────────────────────────────────────
    const now      = new Date();
    const dateRcvd = docData.dateReceived ? new Date(docData.dateReceived) : now;
    const dueDate  = new Date(dateRcvd);
    dueDate.setDate(dueDate.getDate() + OVERDUE_THRESHOLD);
    const origRow   = data[rowIndex - 1];
    const docTsCol  = colIdx['Doc Time Stamp'];
    const origDocTs = (docTsCol !== undefined && origRow[docTsCol])
      ? origRow[docTsCol]
      : Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
    const pdfLinkCol  = colIdx['PDF Link'];
    const existingPdf = pdfLinkCol !== undefined ? origRow[pdfLinkCol] : '';
    const poTs = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
    const drStr = Utilities.formatDate(dateRcvd, tz, 'yyyy-MM-dd');
    const ddStr = Utilities.formatDate(dueDate,  tz, 'yyyy-MM-dd');
    const overdue = calculateOverdueStatus(dateRcvd, docData.status);

    const updRow = [
      origDocTs,
      docId,
      docData.docType       || '',
      docData.docNo         || '',
      docData.prDate        || '',
      docData.description   || '',
      formatAmount(docData.amount),
      docData.endUser       || '',
      poTs,
      docData.poNo          || '',
      docData.poDate        || '',
      formatAmount(docData.poAmount),
      docData.supplier      || '',
      docData.requisitioner || '',
      docData.office        || '',
      docData.status        || 'Received',
      drStr, ddStr, overdue,
      docData.notes         || '',
      docData.pdfLink       || existingPdf || ''
    ];

    // ── ONE write ───────────────────────────────────────────────────────
    sheet.getRange(rowIndex, 1, 1, DOCS_COL_COUNT).setValues([updRow]);
    _cacheInvalidateDocs();

    // ── Build doc object for optimistic frontend update ─────────────────
    const docObj = {};
    EXPECTED_HEADERS.forEach((h, i) => { docObj[h] = updRow[i]; });
    docObj['Overdue'] = overdue;

    // ── Logging (reuse already-fetched user) ────────────────────────────
    const cu     = getCurrentUser();
    const cuName = cu ? cu.name : 'System';
    const statusKey  = (docData.status || '').toLowerCase().trim();
    const actionMap  = {
      'incoming':'Status: Incoming','received':'Document Received',
      'in review':'Status: In Review','in process':'Status: In Process',
      'approved':'Document Approved','forwarded':'Document Forwarded',
      'hold':'Document On Hold','completed':'Document Completed','complete':'Document Completed'
    };
    const actionName = actionMap[statusKey] || 'Document Updated';
    const remarks = [
      `Status: ${docData.status}`,
      docData.office  ? `Endorsed To: ${docData.office}` : '',
      docData.endUser ? `End User: ${docData.endUser}`   : '',
      docData.poNo    ? `PO No: ${docData.poNo}`         : ''
    ].filter(Boolean).join(' | ');
    _logDocHistoryFast(docId, actionName, docData.status || 'Received', cuName, remarks);
    _logActivityFast('Update Document', docId, `Document updated: Status → ${docData.status}`, cuName);

    return { status: 'success', docId: docId, doc: docObj };
  } catch (error) {
    Logger.log('updateDocument error: ' + error);
    return { status: 'error', message: error.toString() };
  }
}

function deleteDocument(docId) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(DOCS_SHEET);
    if (!sheet) return { status: 'error', message: 'Documents sheet not found' };
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const idCol  = headers.indexOf('ID/Barcode');
    const noCol  = headers.indexOf('Doc No');
    const tyCol  = headers.indexOf('Doc Type');
    const stCol  = headers.indexOf('Status');
    const cu   = getCurrentUser();
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol === -1 ? 0 : idCol] === docId) {
        logDocumentHistory(docId, 'Document Deleted', stCol !== -1 ? data[i][stCol] : '', cu ? cu.name : 'System', 'Document permanently deleted');
        logActivity('Delete Document', docId, `Document deleted: ${tyCol !== -1 ? data[i][tyCol] : ''} - ${noCol !== -1 ? data[i][noCol] : ''}`);
        sheet.deleteRow(i + 1);
        _cacheInvalidateDocs();
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'Document not found' };
  } catch (error) {
    Logger.log('deleteDocument error: ' + error);
    return { status: 'error', message: error.toString() };
  }
}

function getDocumentStats() {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(DOCS_SHEET);
    const zero  = { total:0, incoming:0, received:0, outgoing:0, hold:0, complete:0, overdue:0 };
    if (!sheet) return zero;
    const data       = sheet.getDataRange().getValues();
    if (data.length <= 1) return zero;
    const headers    = data[0].map(h => h.toString().trim());
    const idCol      = headers.indexOf('ID/Barcode');
    const noCol      = headers.indexOf('Doc No');
    const statusCol  = headers.indexOf('Status');
    const dateRcvdCol= headers.indexOf('Date Received');
    const stats      = Object.assign({}, zero);
    const seenDocNos = new Set();
    for (let i = 1; i < data.length; i++) {
      const rowId = idCol !== -1 ? data[i][idCol] : data[i][0];
      if (!rowId || rowId.toString().trim() === '') continue;
      const docNo = noCol !== -1 ? data[i][noCol] : data[i][2];
      if (seenDocNos.has(docNo)) continue;
      seenDocNos.add(docNo);
      stats.total++;
      const status   = String(statusCol !== -1 ? data[i][statusCol] : data[i][8]).toLowerCase().trim();
      const dateRcvd = dateRcvdCol !== -1 ? data[i][dateRcvdCol] : data[i][9];
      if (status.includes('forwarded'))                               stats.outgoing++;
      else if (status.includes('completed') || status.includes('complete')) stats.complete++;
      else                                                            stats.incoming++;
      if (status.includes('received')) stats.received++;
      else if (status.includes('hold')) stats.hold++;
      if (dateRcvd) {
        const ov = calculateOverdueStatus(dateRcvd, status);
        if (ov && ov !== 'On time') stats.overdue++;
      }
    }
    return stats;
  } catch (error) {
    Logger.log('getDocumentStats error: ' + error);
    return { total:0, incoming:0, received:0, outgoing:0, hold:0, complete:0, overdue:0 };
  }
}

// =====================================================================
//  PDF UPLOAD
// =====================================================================
function uploadPDFToGoogleDrive(base64Data, fileName, docId) {
  try {
    const blob    = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', fileName);
    const folders = DriveApp.getFoldersByName('DocuTracker_PDFs');
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder('DocuTracker_PDFs');
    const file    = folder.createFile(blob);
    file.setName(`${docId}_${fileName}`);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    logActivity('Upload PDF', docId, `PDF uploaded: ${fileName}`);
    return { status: 'success', fileId: file.getId(), fileUrl: file.getUrl(), fileName: file.getName() };
  } catch (error) {
    Logger.log('uploadPDFToGoogleDrive error: ' + error);
    return { status: 'error', message: error.toString() };
  }
}

// =====================================================================
//  USER MANAGEMENT  (Admin only)
// =====================================================================
function getUsers(tokenParam) {
  try {
    const cu = getCurrentUser(tokenParam || _doPostToken);
    if (!cu || !cu.permissions.manageUsers) return { status: 'error', message: 'Access denied' };
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET);
    if (!sheet) return { status: 'error', message: 'Users sheet not found' };
    _ensureUsersColumns(sheet);
    const data  = sheet.getDataRange().getValues();
    const users = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      users.push({
        email:       data[i][0],
        role:        data[i][2] || 'Staff',
        name:        data[i][3] || '',
        status:      data[i][6] || 'Active',
        createdAt:   data[i][8] ? String(data[i][8]).split('T')[0] : ''
      });
    }
    return { status: 'success', users: users };
  } catch (e) { Logger.log('getUsers error: ' + e); return { status: 'error', message: e.toString() }; }
}

function addUser(tokenParam, userData) {
  try {
    const cu = getCurrentUser(tokenParam || _doPostToken);
    if (!cu || !cu.permissions.manageUsers) return { status: 'error', message: 'Access denied' };
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET);
    if (!sheet) return { status: 'error', message: 'Users sheet not found' };
    _ensureUsersColumns(sheet);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === userData.email.toLowerCase())
        return { status: 'error', message: 'Email already exists' };
    }
    sheet.appendRow([
      userData.email.toLowerCase().trim(),
      userData.password || 'changeme123',
      userData.role     || 'Staff',
      userData.name     || '',
      '', '',                              // ResetToken, ResetExpiry
      userData.status   || 'Active',
      '',                                  // custom permissions (blank = use role defaults)
      new Date(), ''
    ]);
    logActivity('Add User', '', `User added: ${userData.email} (${userData.role})`);
    return { status: 'success' };
  } catch (e) { Logger.log('addUser error: ' + e); return { status: 'error', message: e.toString() }; }
}

function updateUser(tokenParam, targetEmail, userData) {
  try {
    const cu = getCurrentUser(tokenParam || _doPostToken);
    if (!cu || !cu.permissions.manageUsers) return { status: 'error', message: 'Access denied' };
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET);
    if (!sheet) return { status: 'error', message: 'Users sheet not found' };
    _ensureUsersColumns(sheet);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() !== targetEmail.toLowerCase()) continue;
      if (userData.name     !== undefined) sheet.getRange(i + 1, 4).setValue(userData.name);
      if (userData.role     !== undefined) sheet.getRange(i + 1, 3).setValue(userData.role);
      if (userData.status   !== undefined) sheet.getRange(i + 1, 7).setValue(userData.status);
      if (userData.password)               sheet.getRange(i + 1, 2).setValue(userData.password);
      logActivity('Update User', '', `User updated: ${targetEmail}`);
      return { status: 'success' };
    }
    return { status: 'error', message: 'User not found' };
  } catch (e) { Logger.log('updateUser error: ' + e); return { status: 'error', message: e.toString() }; }
}

function toggleUserStatus(tokenParam, targetEmail) {
  try {
    const cu = getCurrentUser(tokenParam || _doPostToken);
    if (!cu || !cu.permissions.manageUsers) return { status: 'error', message: 'Access denied' };
    if (targetEmail.toLowerCase() === cu.email.toLowerCase())
      return { status: 'error', message: 'Cannot deactivate your own account' };
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET);
    if (!sheet) return { status: 'error', message: 'Users sheet not found' };
    _ensureUsersColumns(sheet);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() !== targetEmail.toLowerCase()) continue;
      const newStatus = String(data[i][6] || 'Active').toLowerCase() === 'active' ? 'Inactive' : 'Active';
      sheet.getRange(i + 1, 7).setValue(newStatus);
      logActivity('Toggle User Status', '', `User ${targetEmail} → ${newStatus}`);
      return { status: 'success', newStatus: newStatus };
    }
    return { status: 'error', message: 'User not found' };
  } catch (e) { Logger.log('toggleUserStatus error: ' + e); return { status: 'error', message: e.toString() }; }
}