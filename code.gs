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

// Request-scoped spreadsheet handle — avoids repeated openById() calls within
// a single GAS execution (each openById() adds ~200-400ms overhead).
let _ssCache = null;
function _getSS() {
  if (!_ssCache) _ssCache = SpreadsheetApp.openById(SHEET_ID);
  return _ssCache;
}
// Reset per-request cache (call at start of doPost)
function _resetRequestCache() { _ssCache = null; }

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
    _resetRequestCache(); // reset per-request SS handle

    const allowedMethods = [
      'checkLogin', 'checkSocialLogin', 'logout', 'getCurrentUser', 'updateUserProfile',
      'sendResetEmail', 'validateResetToken', 'setNewPassword',
      'getDropdownOptions', 'updateDropdownOptions', 'updateAllDropdownOptions',
      'getAllDocuments', 'getDocumentById', 'getDocumentStats',
      'addDocument', 'updateDocument', 'deleteDocument',
      'getDocumentHistory', 'logDocumentHistory',
      'getAllActivityLogs', 'logActivity',
      'uploadPDFToGoogleDrive', 'updatePdfLink', 'getScriptUrl',
      'getUsers', 'addUser', 'updateUser', 'toggleUserStatus', 'deleteUser',
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
  const ss = _getSS();
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
    const hit = CacheService.getScriptCache().get('dropdown_opts');
    if (hit) { try { return JSON.parse(hit); } catch(e) {} }
    const ss = _getSS();
    let configSheet = ss.getSheetByName(CONFIG_SHEET) || initializeConfigSheet();
    const data = configSheet.getDataRange().getValues();
    const options = { docTypes: [], suppliers: [], offices: [], statuses: [], endUsers: [] };
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) options.docTypes.push(data[i][0]);
      if (data[i][1]) options.suppliers.push(data[i][1]);
      if (data[i][2]) options.offices.push(data[i][2]);
      if (data[i][3]) options.statuses.push(data[i][3]);
      if (data[i][4]) options.endUsers.push(data[i][4]);
    }
    try { CacheService.getScriptCache().put('dropdown_opts', JSON.stringify(options), 300); } catch(e) {}
    return options;
  } catch (error) {
    Logger.log('getDropdownOptions error: ' + error);
    return { docTypes: ['Purchase Request'], suppliers: [], offices: ['GSO-Admin'], statuses: ['Received'], endUsers: [] };
  }
}

function updateDropdownOptions(type, values, tokenParam) {
  try {
    const currentUser = getCurrentUser(tokenParam || _doPostToken);
    const role = currentUser ? (currentUser.role || '').toLowerCase() : '';
    if (role !== 'admin' && role !== 'manager') {
      return { status: 'error', message: 'Access restricted to Admin or Manager' };
    }
    const ss = _getSS();
    let configSheet = ss.getSheetByName(CONFIG_SHEET) || initializeConfigSheet();
    const columnMap = { docTypes: 1, suppliers: 2, offices: 3, statuses: 4, endUsers: 5 };
    const column = columnMap[type];
    if (!column) return { status: 'error', message: 'Invalid type' };
    
    const uniqueValues = [...new Set(values.map(v => String(v).trim()).filter(v => v))];
    const lastRow = configSheet.getLastRow();
    if (lastRow > 1) {
      configSheet.getRange(2, column, lastRow - 1, 1).clearContent();
    }
    if (uniqueValues.length > 0) {
      const dataArr = uniqueValues.map(v => [v]);
      configSheet.getRange(2, column, dataArr.length, 1).setValues(dataArr);
    }
    
    _invalidateDropdownCache();
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
    // CacheService first (fast path — avoids ScriptProperties quota hits)
    const ck  = 'sv_' + token.replace(/-/g, '');
    const hit = CacheService.getScriptCache().get(ck);
    if (hit) {
      // Sliding renewal: refresh cache TTL on every successful hit so active
      // users never get kicked out just because the 6-min cache window expired.
      CacheService.getScriptCache().put(ck, hit, 600);
      return hit;
    }
    // Slow path — read from ScriptProperties
    const raw = PropertiesService.getScriptProperties().getProperty('sess_' + token);
    if (!raw) return null;
    const sess = JSON.parse(raw);
    const now  = Date.now();
    if (sess.expires < now) {
      PropertiesService.getScriptProperties().deleteProperty('sess_' + token);
      return null;
    }
    // Sliding window: extend the session by another 24h on each successful validation
    // (keeps the session alive as long as the user is active)
    sess.expires = now + 24 * 60 * 60 * 1000;
    PropertiesService.getScriptProperties().setProperty('sess_' + token, JSON.stringify(sess));
    // Cache for 10 minutes — long enough to survive normal page interactions
    CacheService.getScriptCache().put(ck, sess.email, 600);
    return sess.email;
  } catch (e) { Logger.log('_validateSession error: ' + e); return null; }
}

function _destroySession(token) {
  if (!token) return;
  try {
    PropertiesService.getScriptProperties().deleteProperty('sess_' + token);
    CacheService.getScriptCache().remove('sv_' + token.replace(/-/g, ''));
  } catch (e) {}
}
function _invalidateDocsCache() {
  try {
    const cache = CacheService.getScriptCache();
    const nStr  = cache.get('all_docs_n');
    const n     = nStr ? parseInt(nStr) : 1;
    const keys  = ['all_docs_n'];
    if (n === 1) keys.push('all_docs'); else for (let i = 0; i < n; i++) keys.push('all_docs_' + i);
    cache.removeAll(keys);
  } catch(e) {}
}
function _invalidateDropdownCache() { try { CacheService.getScriptCache().remove('dropdown_opts'); } catch(e) {} }

function _getDefaultPermissions(role) {
  const r = (role || 'Staff').toLowerCase();
  if (r === 'admin')   return { addDoc: true,  editDoc: true,  deleteDoc: true,  viewDoc: true, printExport: true,  manageSettings: true,  manageUsers: true,  viewAnalytics: true,  trackHistory: true  };
  if (r === 'manager') return { addDoc: true,  editDoc: true,  deleteDoc: false, viewDoc: true, printExport: true,  manageSettings: false, manageUsers: false, viewAnalytics: true,  trackHistory: true  };
  if (r === 'viewer')  return { addDoc: false, editDoc: false, deleteDoc: false, viewDoc: true, printExport: true,  manageSettings: false, manageUsers: false, viewAnalytics: false, trackHistory: true  };
  return                       { addDoc: true,  editDoc: true,  deleteDoc: false, viewDoc: true, printExport: false, manageSettings: false, manageUsers: false, viewAnalytics: false, trackHistory: true  };
}

function _ensureUsersColumns(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    const cols = [ { idx: 6, name: 'Status' }, { idx: 7, name: 'Permissions' },
                   { idx: 8, name: 'CreatedAt' }, { idx: 9, name: 'LastLogin' },
                   { idx: 10, name: 'Team' } ];
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
    // Short-lived user-object cache (60 s) — avoids repeated sheet reads within one
    // browser session (initDashboard, logActivity, logDocumentHistory all call this).
    const ck = 'cu_' + token.replace(/-/g, '').substring(0, 24);
    try {
      const hit = CacheService.getScriptCache().get(ck);
      if (hit) return JSON.parse(hit);
    } catch (e) {}
    const ss    = _getSS();
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
      const user = { email: data[i][0], name: data[i][3] || 'User', role: role, status: status, permissions: perms, team: String(data[i][10] || '').trim() };
      try { CacheService.getScriptCache().put(ck, JSON.stringify(user), 300); } catch (e) {}
      return user;
    }
    return null;
  } catch (error) { Logger.log('getCurrentUser error: ' + error); return null; }
}

function logout(tokenParam) {
  try {
    const token = tokenParam || _doPostToken;
    const email = _validateSession(token);
    if (email) {
      // Look up the user's display name so the Activity Log shows the real name,
      // not 'System'. We use the cached user object first (fast), then fall back
      // to a sheet read if not cached.
      let displayName = email; // sensible fallback
      try {
        const ck  = 'cu_' + email.replace(/[^a-zA-Z0-9]/g, '_');
        const hit = CacheService.getScriptCache().get(ck);
        if (hit) {
          const u = JSON.parse(hit);
          displayName = u.name || email;
        } else {
          // Direct sheet lookup — only runs on cache miss
          const ss    = _getSS();
          const sheet = ss.getSheetByName(USERS_SHEET);
          if (sheet) {
            const data = sheet.getDataRange().getValues();
            for (let i = 1; i < data.length; i++) {
              if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase()) {
                displayName = String(data[i][3] || '').trim() || email;
                break;
              }
            }
          }
        }
      } catch(e) { Logger.log('logout name lookup error: ' + e); }
      logActivity('Logout', '', `User logged out: ${email}`, displayName);
    }
    _destroySession(token);
  } catch (error) { Logger.log('logout error: ' + error); }
  try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; }
}

function updateUserProfile(tokenParam, updates) {
  try {
    const token = tokenParam || _doPostToken;
    const email = _validateSession(token);
    if (!email) return { status: 'error', message: 'No user session found' };
    const ss    = _getSS();
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
  const ss    = _getSS();
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
    // Pass the user's display name explicitly — at login time there is no active
    // session token yet, so getCurrentUser() inside logActivity would return null
    // and fall back to 'System'. We pass the name directly to avoid that.
    const userName = String(data[i][3] || '').trim() || userEmail;
    logActivity('Login', '', `User logged in: ${userEmail}`, userName);
    return { status: 'success', token: token };
  }
  return { status: 'invalid', message: 'Invalid email or password' };
}

function checkSocialLogin(email) {
  if (!email) return { status: 'invalid', message: 'No email provided' };
  const ss    = _getSS();
  const sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) return { status: 'invalid', message: 'System error' };
  _ensureUsersColumns(sheet);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const userEmail = String(data[i][0]).trim().toLowerCase();
    if (email.toLowerCase() !== userEmail) continue;
    
    const status = String(data[i][6] || 'Active');
    if (status.toLowerCase() === 'inactive')
      return { status: 'invalid', message: 'Account is deactivated. Contact your administrator.' };
      
    const token = _createSession(userEmail);
    const userName = String(data[i][3] || '').trim() || userEmail;
    logActivity('Login', '', `User logged in via Social: ${userEmail}`, userName);
    return { status: 'success', token: token };
  }
  return { status: 'invalid', message: 'Social login email not found. Contact Admin for approval.' };
}

// =====================================================================
//  PASSWORD RESET
// =====================================================================
function sendResetEmail(email) {
  if (!email) return 'notfound';
  const ss = _getSS();
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
  const sheet = _getSS().getSheetByName(USERS_SHEET);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === token && new Date(data[i][5]) > new Date()) return true;
  }
  return false;
}

function setNewPassword(token, newPassword) {
  if (!token || !newPassword) return 'failed';
  const sheet = _getSS().getSheetByName(USERS_SHEET);
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
  const ss = _getSS();
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
    const ss = _getSS();
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
          timestamp:  (function(d){
            const dt = new Date(d);
            if(isNaN(dt)) return String(d);
            const p=n=>String(n).padStart(2,'0');
            const h=dt.getHours()%12||12, ampm=dt.getHours()<12?'AM':'PM';
            return `${dt.toLocaleString('en-US',{month:'short'})} ${String(dt.getDate()).padStart(2,'0')}, ${dt.getFullYear()} ${String(h).padStart(2,'0')}:${p(dt.getMinutes())} ${ampm}`;
          })(data[i][4]),
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
    const ss    = _getSS();
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
            if (h === 'Doc Time Stamp' || h === 'PO Time Stamp' || h === 'Endorsement Time Stamp') {
              doc[h] = _fmtTs(value);
            } else {
              doc[h] = _fmtDate(value);
            }
          } else if (value === '') {
            doc[h] = '';
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
    const ss = _getSS();
    let logSheet = ss.getSheetByName(ACTIVITY_LOG_SHEET);
    if (!logSheet) {
      logSheet = ss.insertSheet(ACTIVITY_LOG_SHEET);
      logSheet.getRange(1, 1, 1, 7).setValues([['Timestamp','Action','User','Document ID','Doc No','Details','Description']]);
      logSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
      logSheet.getRange(1, 1, 1, 7).setBackground('#f3f4f6');
      logSheet.setFrozenRows(1);
    } else {
      // Migrate existing sheet: check if Doc No column already exists (col 5 header)
      const headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
      if (!headers.includes('Doc No')) {
        // Insert Doc No column at position 5 (after Document ID col 4)
        logSheet.insertColumnAfter(4);
        logSheet.getRange(1, 5).setValue('Doc No');
        logSheet.getRange(1, 5).setFontWeight('bold').setBackground('#f3f4f6');
      }
    }
    return logSheet;
  } catch (error) {
    Logger.log('initializeActivityLogSheet error: ' + error);
    return null;
  }
}

// logActivity: action, documentId, details, [explicitUserName], [docNo]
// explicitUserName: pass when getCurrentUser() would return null (e.g. during login)
function logActivity(action, documentId, details, explicitUserName, docNo) {
  try {
    const logSheet = initializeActivityLogSheet();
    if (!logSheet) return false;
    let userName = explicitUserName || '';
    if (!userName) {
      const cu = getCurrentUser();
      userName = cu ? cu.name : 'System';
    }
    logSheet.appendRow([
      new Date(),
      String(action    || '').trim(),
      userName,
      String(documentId || '').trim(),
      String(docNo      || '').trim(),
      String(details    || '').trim(),
      ''                                  // Description column (reserved)
    ]);
    CacheService.getScriptCache().remove('all_activity_logs');
    return true;
  } catch (error) {
    Logger.log('logActivity error: ' + error);
    return false;
  }
}

function getAllActivityLogs() {
  try {
    // 30-second cache — activity logs are append-only; this is safe
    const cache = CacheService.getScriptCache();
    const hit   = cache.get('all_activity_logs');
    if (hit) { try { return JSON.parse(hit); } catch(e) {} }

    const ss    = _getSS();
    const sheet = ss.getSheetByName(ACTIVITY_LOG_SHEET);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    // Format all dates in one pass using JS (no Utilities.formatDate per row)
    const tz     = Session.getScriptTimeZone();
    const locale = 'en-US';
    const logs   = [];
    for (let i = data.length - 1; i >= 1; i--) {          // reversed = newest first, no sort needed
      if (!data[i][0]) continue;
      const ts  = new Date(data[i][0]);
      const fmt = isNaN(ts) ? String(data[i][0]) :
        ts.toLocaleDateString(locale, { month:'short', day:'numeric', year:'numeric' }) + ' ' +
        ts.toLocaleTimeString(locale, { hour:'2-digit', minute:'2-digit' });
      logs.push({
        Timestamp:  fmt,
        Action:     data[i][1] || '',
        User:       data[i][2] || '',
        DocumentID: data[i][3] || '',
        DocNo:      data[i][4] || '',   // Doc No — col 5 (index 4)
        Details:    data[i][5] || '',   // Details — col 6 (index 5)
        IPAddress:  data[i][6] || ''    // Description/reserved — col 7 (index 6)
      });
    }
    try {
      const json = JSON.stringify(logs);
      if (json.length < 98304) cache.put('all_activity_logs', json, 120);
    } catch(e) {}
    return logs;
  } catch (error) { Logger.log('getAllActivityLogs error: ' + error); return []; }
}

// =====================================================================
//  BATCH LOG HELPER — writes Document History + Activity Log in a single
//  sequence using the already-open request-scoped SS handle.
//  This replaces two separate logDocumentHistory() + logActivity() calls
//  (which were each opening sheets independently) with one fast pass.
// =====================================================================
function _logBoth(docId, histAction, histStatus, histUser, histRemarks, actAction, actDetails) {
  try {
    const ss = _getSS();
    // --- Document History ---
    let histSheet = ss.getSheetByName(HISTORY_SHEET);
    if (!histSheet) {
      histSheet = ss.insertSheet(HISTORY_SHEET);
      histSheet.getRange(1,1,1,6).setValues([['Document ID','Action','Status','User','Timestamp','Remarks']]);
      histSheet.getRange(1,1,1,6).setFontWeight('bold');
    }
    const cu = getCurrentUser();
    const userName = cu ? cu.name : 'System';
    const ts = new Date();
    histSheet.appendRow([docId, histAction, histStatus || '', histUser || userName || 'System', ts, histRemarks || '']);

    // --- Resolve Doc No from docId using cached docs (zero extra sheet read) ---
    let docNo = '';
    try {
      const cachedDocs = getAllDocuments();
      const match = cachedDocs.find(d => d['ID/Barcode'] === docId);
      if (match) docNo = match['Doc No'] || '';
    } catch(e) {}

    // --- Activity Log (7 columns: Timestamp, Action, User, Document ID, Doc No, Details, Description) ---
    const logSheet = initializeActivityLogSheet();
    if (logSheet) {
      logSheet.appendRow([ts, String(actAction||'').trim(), userName, String(docId||'').trim(), docNo, String(actDetails||'').trim(), '']);
    }
    // Invalidate activity cache
    try { CacheService.getScriptCache().remove('all_activity_logs'); } catch(e) {}
  } catch(e) { Logger.log('_logBoth error: ' + e); }
}

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
    const sheet = _getSS().getSheetByName(DOCS_SHEET);
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

// =====================================================================
//  SCHEMA MIGRATION HELPER
//  Ensures the Documents sheet has all 20 columns in the correct order.
//  Safe to run on both new and existing sheets.
// =====================================================================
function ensureDocumentSheetHeaders(sheet) {
  try {
    const EXPECTED_HEADERS = [
      'Doc Time Stamp','ID/Barcode','Doc Type','Doc No','PR Date','Description','Amount',
      'EndUser','PO Time Stamp','Date Received From BAC','PO No','PO Date','PO Amount',
      'Supplier','Requisitioner','Endorsed To','Endorsement Time Stamp','Status',
      'Date Endorsed To Acctng','Date Endorsed From CMO','Delivery Status',
      'Date Endorse To COA','Date Endorse To CTO',
      'Date Received','Due Date','Overdue','Notes','PDF Link'
    ];
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) {
      sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).setValues([EXPECTED_HEADERS]);
      sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).setFontWeight('bold');
      return;
    }
    const current = data[0].map(h => h.toString().trim());
    // Validate EVERY header — fixes spelling variants like 'PO TimeStamp' vs 'PO Time Stamp'
    const needsUpdate = current.length < EXPECTED_HEADERS.length ||
      EXPECTED_HEADERS.some((h, i) => current[i] !== h);
    if (needsUpdate) {
      sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).setValues([EXPECTED_HEADERS]);
      sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).setFontWeight('bold');
      Logger.log('Document sheet headers normalized to expected schema.');
    }
  } catch (e) {
    Logger.log('ensureDocumentSheetHeaders error: ' + e);
  }
}

// =====================================================================
//  DOCUMENT CRUD
// =====================================================================
// Pure-JS date formatters — zero GAS API overhead, replaces Utilities.formatDate()
function _fmtDate(d) {
  // yyyy-MM-dd
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}`;
}
function _fmtTs(d) {
  // yyyy-MM-dd HH:mm:ss
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

// Split a JSON string into ≤99 KB chunks for CacheService storage
function _cacheSet(key, json) {
  try {
    const cache = CacheService.getScriptCache();
    const CHUNK = 99000;
    if (json.length <= CHUNK) {
      cache.put(key, json, 600);
      cache.put(key + '_n', '1', 600);
    } else {
      const n = Math.ceil(json.length / CHUNK);
      for (let i = 0; i < n; i++) cache.put(key + '_' + i, json.slice(i*CHUNK, (i+1)*CHUNK), 600);
      cache.put(key + '_n', String(n), 600);
    }
  } catch(e) {}
}
function _cacheGet(key) {
  try {
    const cache = CacheService.getScriptCache();
    const nStr  = cache.get(key + '_n');
    if (!nStr) return null;
    const n = parseInt(nStr);
    if (n === 1) return cache.get(key);
    const parts = [];
    for (let i = 0; i < n; i++) { const p = cache.get(key + '_' + i); if (!p) return null; parts.push(p); }
    return parts.join('');
  } catch(e) { return null; }
}

function getAllDocuments() {
  try {
    // 5-minute chunked cache — works for datasets beyond the 100 KB CacheService limit
    const hit = _cacheGet('all_docs');
    if (hit) { try { return JSON.parse(hit); } catch(e) {} }

    const ss    = _getSS();
    const sheet = ss.getSheetByName(DOCS_SHEET);
    if (!sheet) return [];
    // NOTE: ensureDocumentSheetHeaders() removed from here — it costs a full
    // sheet read on every cold load. It's still called in addDocument() where
    // a schema migration is actually needed.
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers    = data[0].map(h => h.toString().trim());
    const idCol      = headers.indexOf('ID/Barcode');
    const noCol      = headers.indexOf('Doc No');
    const statusCol  = headers.indexOf('Status');
    // Pre-build a set of timestamp column indices so we format them differently
    const tsColSet   = new Set(['Doc Time Stamp','PO Time Stamp','Endorsement Time Stamp'].map(h => headers.indexOf(h)).filter(i => i !== -1));
    const documents  = [];
    const seenIds    = new Set();

    for (let i = 1; i < data.length; i++) {
      const rowId = idCol !== -1 ? data[i][idCol] : data[i][0];
      if (!rowId || rowId.toString().trim() === '') continue;
      const rowIdStr = String(rowId).trim();
      if (seenIds.has(rowIdStr)) continue;
      seenIds.add(rowIdStr);

      const doc = {};
      for (let j = 0; j < headers.length; j++) {
        const v = data[i][j];
        if (v instanceof Date) {
          // Pure JS formatting — no GAS API call, orders of magnitude faster
          doc[headers[j]] = tsColSet.has(j) ? _fmtTs(v) : _fmtDate(v);
        } else {
          doc[headers[j]] = v;
        }
      }
      if (doc['Date Received']) {
        doc['Overdue'] = calculateOverdueStatus(doc['Date Received'], doc['Status']);
      }
      documents.push(doc);
    }

    _cacheSet('all_docs', JSON.stringify(documents));
    return documents;
  } catch (error) { Logger.log('getAllDocuments error: ' + error); return []; }
}

function addDocument(docData) {
  try {
    // Use cached docs for duplicate check — avoids a cold sheet read
    const cachedDocs = getAllDocuments();
    if (cachedDocs.some(d => d['Doc No'] && d['Doc No'].trim().toLowerCase() === (docData.docNo || '').trim().toLowerCase())) {
      return { status: 'error', message: 'duplicate', field: 'docNo', docNo: docData.docNo };
    }
    if (docData.poNo && docData.poNo.trim() !== '' &&
        cachedDocs.some(d => d['PO No'] && d['PO No'].trim().toLowerCase() === docData.poNo.trim().toLowerCase())) {
      return { status: 'error', message: 'duplicate', field: 'poNo', poNo: docData.poNo };
    }
    const ss = _getSS();
    let sheet = ss.getSheetByName(DOCS_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(DOCS_SHEET);
      sheet.getRange(1, 1, 1, 28).setValues([[
        'Doc Time Stamp','ID/Barcode','Doc Type','Doc No','PR Date','Description','Amount',
        'EndUser','PO Time Stamp','Date Received From BAC','PO No','PO Date','PO Amount',
        'Supplier','Requisitioner','Endorsed To','Endorsement Time Stamp','Status',
        'Date Endorsed To Acctng','Date Endorsed From CMO','Delivery Status',
        'Date Endorse To COA','Date Endorse To CTO',
        'Date Received','Due Date','Overdue','Notes','PDF Link'
      ]]);
      sheet.getRange(1, 1, 1, 28).setFontWeight('bold');
    } else {
      // Ensure new columns exist in existing sheets
      ensureDocumentSheetHeaders(sheet);
    }
    const now      = new Date();
    const docId    = `DOC-${_fmtTs(now).replace(/[-: ]/g,'').substring(0,14)}-${String(Math.floor(Math.random() * 10000)).padStart(4, '0')}`;
    const dateRcvd = new Date(docData.dateReceived || now);
    const dueDate  = new Date(dateRcvd);
    dueDate.setDate(dueDate.getDate() + OVERDUE_THRESHOLD);
    // Doc Time Stamp: auto-generated once on create, never changes
    const docTimeStamp = _fmtTs(now);
    // PO Time Stamp: recorded only when PO No. is first provided on creation
    const poTimeStamp = (docData.poNo && docData.poNo.trim() !== '') ? _fmtTs(now) : '';
    // Endorsement Time Stamp: recorded when Endorsed To (office) is set on creation
    const endorsementTimeStamp = (docData.office && docData.office.trim() !== '') ? _fmtTs(now) : '';
    sheet.appendRow([
      docTimeStamp,
      docId,
      docData.docType              || '',
      docData.docNo                || '',
      docData.prDate               || '',
      docData.description          || '',
      formatAmount(docData.amount),
      docData.endUser              || '',
      poTimeStamp,
      docData.dateReceivedFromBAC  || '',
      docData.poNo                 || '',
      docData.poDate               || '',
      formatAmount(docData.poAmount),
      docData.supplier             || '',
      docData.requisitioner        || '',
      docData.office               || '',
      endorsementTimeStamp,
      docData.status               || 'Received',
      docData.dateEndorsedToAcctng || '',
      docData.dateEndorsedFromCMO  || '',
      docData.deliveryStatus       || '',
      docData.dateEndorseToCOA     || '',
      docData.dateEndorseToCTO     || '',
      _fmtDate(dateRcvd),
      _fmtDate(dueDate),
      calculateOverdueStatus(dateRcvd, docData.status),
      docData.notes                || '',
      docData.pdfLink              || ''
    ]);
    _invalidateDocsCache();
    const createRemarks = [
      `${docData.docType} — ${docData.docNo}`,
      docData.office    ? `Endorsed To: ${docData.office}`       : '',
      docData.endUser   ? `End User: ${docData.endUser}`         : '',
      docData.prDate    ? `PR Date: ${docData.prDate}`           : '',
      docData.supplier  ? `Supplier: ${docData.supplier}`        : ''
    ].filter(Boolean).join(' | ');
    const cu = getCurrentUser();
    // Build a descriptive create log message
    const createDetail = 'Created New Document: ' + (docData.description || '(no description)') +
      (docData.docNo    ? ' | Doc No: '       + docData.docNo    : '') +
      (docData.docType  ? ' | Type: '         + docData.docType  : '') +
      (docData.endUser  ? ' | End User: '     + docData.endUser  : '') +
      (docData.supplier ? ' | Supplier: '     + docData.supplier : '') +
      (docData.office   ? ' | Endorsed To: '  + docData.office   : '') +
      (docData.status   ? ' | Status: '       + docData.status   : '');
    _logBoth(docId, 'Document Created', docData.status || 'Received', cu ? cu.name : 'System', createRemarks,
             'Create Document', createDetail);
    return { status: 'success', docId: docId };
  } catch (error) {
    Logger.log('addDocument error: ' + error);
    return { status: 'error', message: error.toString() };
  }
}

function updateDocument(docId, docData) {
  try {
    // Use cached docs for duplicate check — avoids a full sheet read
    const cachedDocs = getAllDocuments();
    const dupDocNo = cachedDocs.some(d =>
      d['Doc No'] && d['Doc No'].trim().toLowerCase() === (docData.docNo || '').trim().toLowerCase() && d['ID/Barcode'] !== docId
    );
    if (dupDocNo) return { status: 'error', message: 'duplicate', field: 'docNo', docNo: docData.docNo };
    if (docData.poNo && docData.poNo.trim() !== '') {
      const dupPoNo = cachedDocs.some(d =>
        d['PO No'] && d['PO No'].trim().toLowerCase() === docData.poNo.trim().toLowerCase() && d['ID/Barcode'] !== docId
      );
      if (dupPoNo) return { status: 'error', message: 'duplicate', field: 'poNo', poNo: docData.poNo };
    }

    const ss    = _getSS();
    const sheet = ss.getSheetByName(DOCS_SHEET);
    if (!sheet) return { status: 'error', message: 'Documents sheet not found' };
    // Migrate sheet headers first — ensures 'Endorsement Time Stamp' column exists
    // before we read or write data. Safe to call every time (no-op if already migrated).
    ensureDocumentSheetHeaders(sheet);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    // Build column index map from the (now-migrated) headers
    const colIdx = {};
    headers.forEach((h, i) => { colIdx[h] = i; });
    let rowIndex = -1;
    const idCol = colIdx['ID/Barcode'] !== undefined ? colIdx['ID/Barcode'] : 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === docId) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) return { status: 'error', message: 'Document not found' };
    const dateRcvd = docData.dateReceived ? new Date(docData.dateReceived) : new Date();
    const dueDate  = new Date(dateRcvd);
    dueDate.setDate(dueDate.getDate() + OVERDUE_THRESHOLD);
    const cu = getCurrentUser();
    const now = new Date();
    // Preserve original Doc Time Stamp — never changes after creation
    const origRow      = data[rowIndex - 1];
    const docTsCol     = colIdx['Doc Time Stamp'];
    const origDocTs    = (docTsCol !== undefined && origRow[docTsCol]) ? (origRow[docTsCol] instanceof Date ? _fmtTs(origRow[docTsCol]) : origRow[docTsCol]) : _fmtTs(now);
    // PO Time Stamp: preserve existing value; only set if it was blank and poNo is now provided
    const poTsCol      = colIdx['PO Time Stamp'];
    const origPoTs     = (poTsCol !== undefined && origRow[poTsCol])
                          ? (origRow[poTsCol] instanceof Date ? _fmtTs(origRow[poTsCol]) : origRow[poTsCol])
                          : '';
    const poNoCol      = colIdx['PO No'];
    const origPoNo     = (poNoCol !== undefined) ? String(origRow[poNoCol] || '').trim() : '';
    const newPoNo      = String(docData.poNo || '').trim();
    // Set PO TS only when: (a) existing PO TS is blank AND a new PO No. is now being entered
    const poTimeStamp  = origPoTs
                          ? origPoTs                                       // preserve existing
                          : (newPoNo !== '' ? _fmtTs(now) : '');          // first-time entry only
    // Endorsement Time Stamp: update only when Endorsed To value actually changes
    const endTsCol     = colIdx['Endorsement Time Stamp'];
    const origEndTs    = (endTsCol !== undefined && origRow[endTsCol])
                          ? (origRow[endTsCol] instanceof Date ? _fmtTs(origRow[endTsCol]) : origRow[endTsCol])
                          : '';
    const officeCol    = colIdx['Endorsed To'];
    const origOffice   = (officeCol !== undefined) ? String(origRow[officeCol] || '').trim() : '';
    const newOffice    = String(docData.office || '').trim();
    // Only stamp if office value changed, or if it's being set for the first time
    const endorsementTimeStamp = (newOffice !== origOffice && newOffice !== '')
                                  ? _fmtTs(now)
                                  : origEndTs;
    const pdfLinkCol   = colIdx['PDF Link'];
    const existingPdf  = (pdfLinkCol !== undefined) ? origRow[pdfLinkCol] : '';
    // ── Column-name-aware row write ────────────────────────────────────────────
    // Builds the updated row from the original row data, overwriting only the
    // fields that belong to this update. Works correctly with BOTH the old
    // 27-column schema and the new 28-column schema — no positional assumptions.
    const updatedRow = Array.from(origRow);
    // Pad to header length in case this row predates the new column (old data)
    while (updatedRow.length < headers.length) updatedRow.push('');
    const _setF = (name, val) => { const ci = colIdx[name]; if (ci !== undefined) updatedRow[ci] = val; };
    _setF('Doc Time Stamp',          origDocTs);
    _setF('Doc Type',                docData.docType              || '');
    _setF('Doc No',                  docData.docNo                || '');
    _setF('PR Date',                 docData.prDate               || '');
    _setF('Description',             docData.description          || '');
    _setF('Amount',                  formatAmount(docData.amount));
    _setF('EndUser',                 docData.endUser              || '');
    _setF('PO Time Stamp',           poTimeStamp);
    _setF('Date Received From BAC',  docData.dateReceivedFromBAC  || '');
    _setF('PO No',                   docData.poNo                 || '');
    _setF('PO Date',                 docData.poDate               || '');
    _setF('PO Amount',               formatAmount(docData.poAmount));
    _setF('Supplier',                docData.supplier             || '');
    _setF('Requisitioner',           docData.requisitioner        || '');
    _setF('Endorsed To',             docData.office               || '');
    _setF('Endorsement Time Stamp',  endorsementTimeStamp);
    _setF('Status',                  docData.status               || 'Received');
    _setF('Date Endorsed To Acctng', docData.dateEndorsedToAcctng || '');
    _setF('Date Endorsed From CMO',  docData.dateEndorsedFromCMO  || '');
    _setF('Delivery Status',         docData.deliveryStatus       || '');
    _setF('Date Endorse To COA',     docData.dateEndorseToCOA     || '');
    _setF('Date Endorse To CTO',     docData.dateEndorseToCTO     || '');
    _setF('Date Received',           _fmtDate(dateRcvd));
    _setF('Due Date',                _fmtDate(dueDate));
    _setF('Overdue',                 calculateOverdueStatus(dateRcvd, docData.status));
    _setF('Notes',                   docData.notes                || '');
    _setF('PDF Link',                docData.pdfLink              || existingPdf || '');
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([updatedRow]);    _invalidateDocsCache();
    const statusActionMap = {
      'incoming':   'Status: Incoming',
      'received':   'Document Received',
      'in review':  'Status: In Review',
      'in process': 'Status: In Process',
      'approved':   'Document Approved',
      'forwarded':  'Document Forwarded',
      'hold':       'Document On Hold',
      'completed':  'Document Completed',
      'complete':   'Document Completed'
    };
    const statusKey  = (docData.status || '').toLowerCase().trim();
    const actionName = statusActionMap[statusKey] || 'Document Updated';
    const updateRemarks = [
      `Status: ${docData.status}`,
      docData.office   ? `Endorsed To: ${docData.office}`       : '',
      docData.endUser  ? `End User: ${docData.endUser}`         : '',
      docData.dateReceived ? `Date Received: ${docData.dateReceived}` : '',
      docData.poNo     ? `PO No: ${docData.poNo}`               : ''
    ].filter(Boolean).join(' | ');

    // ── Field-level change detection ────────────────────────────────────────
    // Compare each tracked field's old value (origRow) with the incoming new
    // value (docData) and surface only the changed fields in the Activity Log.
    const _fldDefs = [
      { label: 'Doc No',              col: 'Doc No',                  nw: docData.docNo                },
      { label: 'PR Date',             col: 'PR Date',                 nw: docData.prDate               },
      { label: 'Description',         col: 'Description',             nw: docData.description          },
      { label: 'Amount',              col: 'Amount',                  nw: formatAmount(docData.amount) },
      { label: 'End User',            col: 'EndUser',                 nw: docData.endUser              },
      { label: 'Date Rcvd From BAC',  col: 'Date Received From BAC',  nw: docData.dateReceivedFromBAC  },
      { label: 'PO No',               col: 'PO No',                   nw: docData.poNo                 },
      { label: 'PO Date',             col: 'PO Date',                 nw: docData.poDate               },
      { label: 'PO Amount',           col: 'PO Amount',               nw: formatAmount(docData.poAmount) },
      { label: 'Supplier',            col: 'Supplier',                nw: docData.supplier             },
      { label: 'Requisitioner',       col: 'Requisitioner',           nw: docData.requisitioner        },
      { label: 'Endorsed To',         col: 'Endorsed To',             nw: docData.office               },
      { label: 'Status',              col: 'Status',                  nw: docData.status               },
      { label: 'Date End. To Acctng', col: 'Date Endorsed To Acctng', nw: docData.dateEndorsedToAcctng },
      { label: 'Date End. From CMO',  col: 'Date Endorsed From CMO',  nw: docData.dateEndorsedFromCMO  },
      { label: 'Delivery Status',     col: 'Delivery Status',         nw: docData.deliveryStatus       },
      { label: 'Date End. To COA',    col: 'Date Endorse To COA',     nw: docData.dateEndorseToCOA     },
      { label: 'Date End. To CTO',    col: 'Date Endorse To CTO',     nw: docData.dateEndorseToCTO     },
      { label: 'Notes',               col: 'Notes',                   nw: docData.notes                }
    ];
    const _fChanges = [];
    _fldDefs.forEach(function(f) {
      const ci = colIdx[f.col];
      if (ci === undefined) return;
      const raw    = origRow[ci];
      const oldStr = String(raw instanceof Date ? _fmtDate(raw) : (raw || '')).trim();
      const newStr = String(f.nw || '').trim();
      if (oldStr !== newStr) {
        _fChanges.push(f.label + ': "' + (oldStr || '—') + '" → "' + (newStr || '—') + '"');
      }
    });
    const changeDetail = _fChanges.length > 0
      ? 'Updated [' + (docData.docNo || docId) + ']: ' + _fChanges.join(' | ')
      : 'Updated [' + (docData.docNo || docId) + ']: No field changes detected';
    _logBoth(docId, actionName, docData.status || 'Received', cu ? cu.name : 'System', updateRemarks,
             'Update Document', changeDetail);
    return { status: 'success', docId: docId };
  } catch (error) {
    Logger.log('updateDocument error: ' + error);
    return { status: 'error', message: error.toString() };
  }
}

function deleteDocument(docId) {
  try {
    const ss    = _getSS();
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
        const delStatus = stCol !== -1 ? data[i][stCol] : '';
        const delType   = tyCol !== -1 ? data[i][tyCol] : '';
        const delNo     = noCol !== -1 ? data[i][noCol] : '';
        sheet.deleteRow(i + 1);
        _invalidateDocsCache();
        _logBoth(docId, 'Document Deleted', delStatus, cu ? cu.name : 'System', 'Document permanently deleted',
                 'Delete Document', `Document deleted: ${delType} - ${delNo}`);
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
    // Reuse the cached getAllDocuments result — no extra sheet read needed
    const docs = getAllDocuments();
    const stats = { total:0, incoming:0, received:0, outgoing:0, hold:0, complete:0, overdue:0 };
    docs.forEach(doc => {
      stats.total++;
      const st = String(doc.Status || '').toLowerCase().trim();
      if      (st.includes('forwarded'))                                  stats.outgoing++;
      else if (st.includes('completed') || st.includes('complete'))      stats.complete++;
      else                                                                stats.incoming++;
      if (st.includes('received')) stats.received++;
      if (st.includes('hold'))     stats.hold++;
      if (doc.Overdue && doc.Overdue !== 'On time')                      stats.overdue++;
    });
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
    const blob   = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', fileName);
    // Always save to the configured DRIVE_FOLDER_ID — never create/search by name
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const file   = folder.createFile(blob);
    const safeName = (docId && docId !== 'temp') ? `${docId}_${fileName}` : fileName;
    file.setName(safeName);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    logActivity('Upload PDF', docId || '', `PDF uploaded: ${fileName}`);
    return { status: 'success', fileId: file.getId(), fileUrl: file.getUrl(), fileName: file.getName() };
  } catch (error) {
    Logger.log('uploadPDFToGoogleDrive error: ' + error);
    return { status: 'error', message: error.toString() };
  }
}

// =====================================================================
//  UPDATE PDF LINK — lightweight single-cell write, no full row rewrite.
//  Called from the frontend after a background PDF upload completes.
// =====================================================================
function updatePdfLink(docId, pdfUrl) {
  try {
    const ss    = _getSS();
    const sheet = ss.getSheetByName(DOCS_SHEET);
    if (!sheet) return { status: 'error', message: 'Documents sheet not found' };
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const idCol   = headers.indexOf('ID/Barcode');
    const pdfCol  = headers.indexOf('PDF Link');
    if (pdfCol === -1) return { status: 'error', message: 'PDF Link column not found' };
    for (let i = 1; i < data.length; i++) {
      const rowId = idCol !== -1 ? data[i][idCol] : data[i][1];
      if (String(rowId).trim() === String(docId).trim()) {
        sheet.getRange(i + 1, pdfCol + 1).setValue(pdfUrl || '');
        _invalidateDocsCache();
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'Document not found' };
  } catch (error) {
    Logger.log('updatePdfLink error: ' + error);
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
    const sheet = _getSS().getSheetByName(USERS_SHEET);
    if (!sheet) return { status: 'error', message: 'Users sheet not found' };
    _ensureUsersColumns(sheet);
    const data  = sheet.getDataRange().getValues();
    const users = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const rawPerms = data[i][7] ? (() => { try { return JSON.parse(data[i][7]); } catch(e) { return null; } })() : null;
      const basePerms = _getDefaultPermissions(data[i][2] || 'Staff');
      if (rawPerms) Object.assign(basePerms, rawPerms);
      users.push({
        email:       data[i][0],
        role:        data[i][2] || 'Staff',
        name:        data[i][3] || '',
        status:      data[i][6] || 'Active',
        permissions: basePerms,
        team:        String(data[i][10] || '').trim(),
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
    const sheet = _getSS().getSheetByName(USERS_SHEET);
    if (!sheet) return { status: 'error', message: 'Users sheet not found' };
    _ensureUsersColumns(sheet);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === userData.email.toLowerCase())
        return { status: 'error', message: 'Email already exists' };
    }
    // Serialize custom permissions if provided (non-default overrides)
    let permJson = '';
    if (userData.permissions && typeof userData.permissions === 'object') {
      try { permJson = JSON.stringify(userData.permissions); } catch(e) {}
    }
    sheet.appendRow([
      userData.email.toLowerCase().trim(),
      userData.password || 'changeme123',
      userData.role     || 'Staff',
      userData.name     || '',
      '', '',                              // ResetToken, ResetExpiry
      userData.status   || 'Active',
      permJson,                            // custom permissions JSON
      new Date(), '',                      // CreatedAt, LastLogin
      userData.team     || ''              // Team (col K)
    ]);
    logActivity('Add User', '', `User added: ${userData.email} (${userData.role})`);
    return { status: 'success' };
  } catch (e) { Logger.log('addUser error: ' + e); return { status: 'error', message: e.toString() }; }
}

function updateUser(tokenParam, targetEmail, userData) {
  try {
    const cu = getCurrentUser(tokenParam || _doPostToken);
    if (!cu || !cu.permissions.manageUsers) return { status: 'error', message: 'Access denied' };
    const sheet = _getSS().getSheetByName(USERS_SHEET);
    if (!sheet) return { status: 'error', message: 'Users sheet not found' };
    _ensureUsersColumns(sheet);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() !== targetEmail.toLowerCase()) continue;
      if (userData.name     !== undefined) sheet.getRange(i + 1, 4).setValue(userData.name);
      if (userData.role     !== undefined) sheet.getRange(i + 1, 3).setValue(userData.role);
      if (userData.status   !== undefined) sheet.getRange(i + 1, 7).setValue(userData.status);
      if (userData.password)               sheet.getRange(i + 1, 2).setValue(userData.password);
      if (userData.team     !== undefined) sheet.getRange(i + 1, 11).setValue(userData.team || '');
      // Persist custom permissions JSON if provided
      if (userData.permissions && typeof userData.permissions === 'object') {
        try { sheet.getRange(i + 1, 8).setValue(JSON.stringify(userData.permissions)); } catch(e) {}
      }
      // Invalidate per-user cache so next request picks up new permissions/team
      try {
        const ck = 'cu_' + targetEmail.replace(/[^a-zA-Z0-9]/g, '_');
        CacheService.getScriptCache().remove(ck);
      } catch(e) {}
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
    const sheet = _getSS().getSheetByName(USERS_SHEET);
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

function deleteUser(tokenParam, targetEmail) {
  try {
    const cu = getCurrentUser(tokenParam || _doPostToken);
    // Only admins can delete users
    if (!cu || cu.role !== 'Admin') return { status: 'error', message: 'Access denied — Admin only' };
    // Cannot delete yourself
    if (targetEmail.toLowerCase() === cu.email.toLowerCase())
      return { status: 'error', message: 'You cannot delete your own account' };
    const sheet = _getSS().getSheetByName(USERS_SHEET);
    if (!sheet) return { status: 'error', message: 'Users sheet not found' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() !== targetEmail.toLowerCase()) continue;
      const deletedName = String(data[i][3] || targetEmail).trim();
      sheet.deleteRow(i + 1);
      // Invalidate user cache for this email
      try { CacheService.getScriptCache().remove('cu_' + targetEmail.replace(/[^a-zA-Z0-9]/g, '_')); } catch(e) {}
      logActivity('Delete User', '', `User deleted: ${deletedName} (${targetEmail})`);
      return { status: 'success', deletedEmail: targetEmail };
    }
    return { status: 'error', message: 'User not found' };
  } catch (e) { Logger.log('deleteUser error: ' + e); return { status: 'error', message: e.toString() }; }
}

// =====================================================================
//  BATCH INIT — returns user + documents + dropdown options + stats
//  in a single round-trip so initDashboard() only needs ONE server call.
// =====================================================================
function getInitialData(token) {
  try {
    _doPostToken = token || _doPostToken;
    const user = getCurrentUser(token);
    if (!user) return { user: null, docs: [], opts: { docTypes:[], suppliers:[], offices:[], statuses:[], endUsers:[] }, stats: {}, logs: [], users: [] };

    // ── Fast path: serve entirely from cache ─────────────────────────────────
    const cache    = CacheService.getScriptCache();
    const docsCached = _cacheGet('all_docs');
    const optsCached = cache.get('dropdown_opts');

    if (docsCached && optsCached) {
      // All data cached — zero sheet reads, returns in ~100ms
      try {
        const docs = JSON.parse(docsCached);
        const opts = JSON.parse(optsCached);
        let logs = [];
        try { const lc = cache.get('all_activity_logs'); if (lc) logs = JSON.parse(lc); } catch(e) {}
        let users = [];
        if (user.permissions && user.permissions.manageUsers) {
          try { const r = getUsers(token); users = (r && r.users) ? r.users : []; } catch(e) {}
        }
        const stats = _computeStatsGAS(docs);
        return { user, docs, opts, stats, logs, users };
      } catch(e) {
        // Cache parse error — fall through to cold path
      }
    }

    // ── Cold path: read all needed sheets in ONE batch ────────────────────────
    // Load all sheet data in one pass rather than sheet-by-sheet
    const ss     = _getSS();
    const sheets = ss.getSheets();
    const sheetMap = {};
    sheets.forEach(s => { sheetMap[s.getName()] = s; });

    // Read docs sheet
    let docs = [];
    if (!docsCached) {
      const docSheet = sheetMap[DOCS_SHEET];
      if (docSheet) docs = _parseDocsSheet(docSheet);
    } else {
      try { docs = JSON.parse(docsCached); } catch(e) { docs = []; }
    }

    // Read config/dropdown sheet
    let opts = { docTypes: [], suppliers: [], offices: [], statuses: [], endUsers: [] };
    if (!optsCached) {
      const cfgSheet = sheetMap[CONFIG_SHEET] || initializeConfigSheet();
      if (cfgSheet) {
        const cfgData = cfgSheet.getDataRange().getValues();
        for (let i = 1; i < cfgData.length; i++) {
          if (cfgData[i][0]) opts.docTypes.push(cfgData[i][0]);
          if (cfgData[i][1]) opts.suppliers.push(cfgData[i][1]);
          if (cfgData[i][2]) opts.offices.push(cfgData[i][2]);
          if (cfgData[i][3]) opts.statuses.push(cfgData[i][3]);
          if (cfgData[i][4]) opts.endUsers.push(cfgData[i][4]);
        }
        try { cache.put('dropdown_opts', JSON.stringify(opts), 600); } catch(e) {}
      }
    } else {
      try { opts = JSON.parse(optsCached); } catch(e) {}
    }

    // Logs: only from cache (not read on cold path — too slow for startup)
    let logs = [];
    try { const lc = cache.get('all_activity_logs'); if (lc) logs = JSON.parse(lc); } catch(e) {}

    // Users: only for admin/manager
    let users = [];
    if (user.permissions && user.permissions.manageUsers) {
      try { const r = getUsers(token); users = (r && r.users) ? r.users : []; } catch(e) {}
    }

    const stats = _computeStatsGAS(docs);
    return { user, docs, opts, stats, logs, users };
  } catch (e) {
    Logger.log('getInitialData error: ' + e);
    return { user: null, docs: [], opts: { docTypes:[], suppliers:[], offices:[], statuses:[], endUsers:[] }, stats: {}, logs: [], users: [] };
  }
}

// Pure-GAS stats computation (mirrors client-side _computeStats)
function _computeStatsGAS(docs) {
  const s = { total:0, incoming:0, received:0, outgoing:0, hold:0, complete:0, overdue:0 };
  (docs || []).forEach(doc => {
    s.total++;
    const st = String(doc.Status || '').toLowerCase().trim();
    if      (st.includes('forwarded'))                             s.outgoing++;
    else if (st.includes('completed') || st.includes('complete')) s.complete++;
    else                                                           s.incoming++;
    if (st.includes('received')) s.received++;
    if (st.includes('hold'))     s.hold++;
    if (doc.Overdue && doc.Overdue !== 'On time')                 s.overdue++;
  });
  return s;
}

// Parse a document sheet into the standard docs array (extracted from getAllDocuments)
function _parseDocsSheet(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers   = data[0].map(h => h.toString().trim());
  const idCol     = headers.indexOf('ID/Barcode');
  const tsColSet  = new Set(['Doc Time Stamp','PO Time Stamp'].map(h => headers.indexOf(h)).filter(i => i !== -1));
  const seenIds   = new Set();
  const documents = [];
  for (let i = 1; i < data.length; i++) {
    const rowId = idCol !== -1 ? data[i][idCol] : data[i][0];
    if (!rowId || rowId.toString().trim() === '') continue;
    const rowIdStr = String(rowId).trim();
    if (seenIds.has(rowIdStr)) continue;
    seenIds.add(rowIdStr);
    const doc = {};
    for (let j = 0; j < headers.length; j++) {
      const v = data[i][j];
      doc[headers[j]] = (v instanceof Date) ? (tsColSet.has(j) ? _fmtTs(v) : _fmtDate(v)) : v;
    }
    if (doc['Date Received']) doc['Overdue'] = calculateOverdueStatus(doc['Date Received'], doc['Status']);
    documents.push(doc);
  }
  _cacheSet('all_docs', JSON.stringify(documents));
  return documents;
}

// =====================================================================
//  BATCH DROPDOWN UPDATE — saves all 5 categories in ONE server call
//  instead of 5 separate google.script.run calls.
// =====================================================================
function updateAllDropdownOptions(optionsObj, tokenParam) {
  try {
    const currentUser = getCurrentUser(tokenParam || _doPostToken);
    const role = currentUser ? (currentUser.role || '').toLowerCase() : '';
    if (role !== 'admin' && role !== 'manager') {
      return { status: 'error', message: 'Access restricted to Admin or Manager' };
    }
    const ss = _getSS();
    let configSheet = ss.getSheetByName(CONFIG_SHEET) || initializeConfigSheet();

    const keys = ['docTypes','suppliers','offices','statuses','endUsers'];
    // Build per-column value arrays
    const cols = keys.map(k =>
      [...new Set((optionsObj[k] || []).map(v => String(v).trim()).filter(v => v))]
    );
    const maxRows = Math.max(...cols.map(c => c.length), 1);

    // Clear existing data rows (all 5 columns in one API call)
    const lastRow = configSheet.getLastRow();
    if (lastRow > 1) configSheet.getRange(2, 1, lastRow - 1, 5).clearContent();

    // Write all 5 columns in ONE setValues call (build 2D array)
    const grid = Array.from({ length: maxRows }, (_, r) =>
      cols.map(col => col[r] !== undefined ? col[r] : '')
    );
    configSheet.getRange(2, 1, maxRows, 5).setValues(grid);

    _invalidateDropdownCache();
    return { status: 'success', message: 'All dropdown options updated successfully' };
  } catch (error) {
    Logger.log('updateAllDropdownOptions error: ' + error);
    return { status: 'error', message: error.toString() };
  }
}