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

// =====================================================================
//  MAIN ENTRY POINT  –  All pages served from one index.html (SPA)
// =====================================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('DocuTracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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

    const allowedMethods = [
      'checkLogin', 'logout', 'getCurrentUser', 'updateUserProfile',
      'sendResetEmail', 'validateResetToken', 'setNewPassword',
      'getDropdownOptions', 'updateDropdownOptions',
      'getAllDocuments', 'getDocumentById', 'getDocumentStats',
      'addDocument', 'updateDocument', 'deleteDocument',
      'getDocumentHistory', 'logDocumentHistory',
      'getAllActivityLogs', 'logActivity',
      'uploadPDFToGoogleDrive', 'getScriptUrl'
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

    // For methods that need a session, we pass the email token from the request
    if (body.sessionEmail) {
      PropertiesService.getUserProperties().setProperty('currentUserEmail', body.sessionEmail);
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
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let configSheet = ss.getSheetByName(CONFIG_SHEET) || initializeConfigSheet();
    initializeConfigSheet(); // ensure EndUser column exists
    const data = configSheet.getDataRange().getValues();
    const options = { docTypes: [], suppliers: [], offices: [], statuses: [], endUsers: [] };
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) options.docTypes.push(data[i][0]);
      if (data[i][1]) options.suppliers.push(data[i][1]);
      if (data[i][2]) options.offices.push(data[i][2]);
      if (data[i][3]) options.statuses.push(data[i][3]);
      if (data[i][4]) options.endUsers.push(data[i][4]);
    }
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
//  USER SESSION MANAGEMENT
// =====================================================================
function getCurrentUser() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const userEmail = userProperties.getProperty('currentUserEmail');
    if (!userEmail) return null;
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(USERS_SHEET);
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === userEmail.toLowerCase()) {
        return { email: data[i][0], name: data[i][3] || 'User', role: data[i][2] || 'User' };
      }
    }
    return null;
  } catch (error) {
    Logger.log('getCurrentUser error: ' + error);
    return null;
  }
}

function logout() {
  try {
    const currentUser = getCurrentUser();
    if (currentUser) logActivity('Logout', '', `User logged out: ${currentUser.email}`);
    PropertiesService.getUserProperties().deleteProperty('currentUserEmail');
    return ScriptApp.getService().getUrl();
  } catch (error) {
    Logger.log('logout error: ' + error);
    return ScriptApp.getService().getUrl();
  }
}

function updateUserProfile(updates) {
  try {
    const userEmail = PropertiesService.getUserProperties().getProperty('currentUserEmail');
    if (!userEmail) return { status: 'error', message: 'No user session found' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(USERS_SHEET);
    if (!sheet) return { status: 'error', message: 'Users sheet not found' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === userEmail.toLowerCase()) {
        if (updates.name)     sheet.getRange(i + 1, 4).setValue(updates.name);
        if (updates.password) sheet.getRange(i + 1, 2).setValue(updates.password);
        return { status: 'success', message: 'Profile updated successfully' };
      }
    }
    return { status: 'error', message: 'User not found' };
  } catch (error) {
    Logger.log('updateUserProfile error: ' + error);
    return { status: 'error', message: error.toString() };
  }
}

// =====================================================================
//  AUTHENTICATION
// =====================================================================
function checkLogin(email, password) {
  if (!email || !password) return { status: 'invalid', message: 'Please enter email and password' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) return { status: 'invalid', message: 'System error' };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const userEmail = String(data[i][0]).trim().toLowerCase();
    const userPass  = String(data[i][1]).trim();
    if (email.toLowerCase() === userEmail && password === userPass) {
      PropertiesService.getUserProperties().setProperty('currentUserEmail', userEmail);
      logActivity('Login', '', `User logged in: ${userEmail}`);
      return { status: 'success' };
    }
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
  const data = sheet.getDataRange().getValues();
  const token  = Utilities.getUuid();
  const expiry = new Date(Date.now() + 1000 * 60 * 30);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, 5).setValue(token);
      sheet.getRange(i + 1, 6).setValue(expiry);
      const resetLink = ScriptApp.getService().getUrl() + '?token=' + encodeURIComponent(token);
      try {
        MailApp.sendEmail({
          to: email,
          subject: 'DocuTracker Password Reset',
          body: `Hello,\n\nA password reset was requested for your DocuTracker account.\n\nClick the link below to reset your password:\n${resetLink}\n\nThis link is valid for 30 minutes.\n\nIf you did not request this, you can safely ignore this email.\n\nThank you,\nDocuTracker System`
        });
      } catch (error) {
        Logger.log('Email send error: ' + error);
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

// =====================================================================
//  SCHEMA MIGRATION HELPER
//  Ensures the Documents sheet has all 20 columns in the correct order.
//  Safe to run on both new and existing sheets.
// =====================================================================
function ensureDocumentSheetHeaders(sheet) {
  try {
    const EXPECTED_HEADERS = [
      'Doc Time Stamp','ID/Barcode','Doc Type','Doc No','PR Date','Description','Amount',
      'EndUser','PO Time Stamp','PO No','PO Date',
      'Supplier','Requisitioner','Endorsed To','Status',
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
    if (checkDuplicateDocNo(docData.docNo)) {
      return { status: 'error', message: 'duplicate', docNo: docData.docNo };
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(DOCS_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(DOCS_SHEET);
      sheet.getRange(1, 1, 1, 20).setValues([[
        'Doc Time Stamp','ID/Barcode','Doc Type','Doc No','PR Date','Description','Amount',
        'EndUser','PO Time Stamp','PO No','PO Date',
        'Supplier','Requisitioner','Endorsed To','Status',
        'Date Received','Due Date','Overdue','Notes','PDF Link'
      ]]);
      sheet.getRange(1, 1, 1, 20).setFontWeight('bold');
    } else {
      // Ensure new columns exist in existing sheets
      ensureDocumentSheetHeaders(sheet);
    }
    const now      = new Date();
    const docId    = `DOC-${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss')}-${String(Math.floor(Math.random() * 10000)).padStart(4, '0')}`;
    const dateRcvd = new Date(docData.dateReceived || now);
    const dueDate  = new Date(dateRcvd);
    dueDate.setDate(dueDate.getDate() + OVERDUE_THRESHOLD);
    const cu = getCurrentUser();
    // Doc Time Stamp: auto-generated on create
    const docTimeStamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    // PO Time Stamp: auto-generated on create/update
    const poTimeStamp  = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([
      docTimeStamp,
      docId,
      docData.docType        || '',
      docData.docNo          || '',
      docData.prDate         || '',
      docData.description    || '',
      formatAmount(docData.amount),
      docData.endUser        || '',
      poTimeStamp,
      docData.poNo           || '',
      docData.poDate         || '',
      docData.supplier       || '',
      docData.requisitioner  || '',
      docData.office         || '',
      docData.status         || 'Received',
      Utilities.formatDate(dateRcvd, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      Utilities.formatDate(dueDate,  Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      calculateOverdueStatus(dateRcvd, docData.status),
      docData.notes          || '',
      docData.pdfLink        || ''
    ]);
    const createRemarks = [
      `${docData.docType} — ${docData.docNo}`,
      docData.office    ? `Endorsed To: ${docData.office}`       : '',
      docData.endUser   ? `End User: ${docData.endUser}`         : '',
      docData.prDate    ? `PR Date: ${docData.prDate}`           : '',
      docData.supplier  ? `Supplier: ${docData.supplier}`        : ''
    ].filter(Boolean).join(' | ');
    logDocumentHistory(docId, 'Document Created', docData.status || 'Received', cu ? cu.name : 'System', createRemarks);
    logActivity('Create Document', docId, `Document created: ${docData.docType} - ${docData.docNo}`);
    return { status: 'success', docId: docId };
  } catch (error) {
    Logger.log('addDocument error: ' + error);
    return { status: 'error', message: error.toString() };
  }
}

function updateDocument(docId, docData) {
  try {
    if (checkDuplicateDocNo(docData.docNo, docId)) {
      return { status: 'error', message: 'duplicate', docNo: docData.docNo };
    }
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(DOCS_SHEET);
    if (!sheet) return { status: 'error', message: 'Documents sheet not found' };
    ensureDocumentSheetHeaders(sheet);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    // Build column index map from actual headers (handles both old and new schemas)
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
    // Preserve original Doc Time Stamp; generate new PO Time Stamp on update
    const origRow      = data[rowIndex - 1];
    const docTsCol     = colIdx['Doc Time Stamp'];
    const origDocTs    = (docTsCol !== undefined && origRow[docTsCol]) ? origRow[docTsCol] : Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const poTimeStamp  = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const pdfLinkCol   = colIdx['PDF Link'];
    const existingPdf  = (pdfLinkCol !== undefined) ? origRow[pdfLinkCol] : '';
    sheet.getRange(rowIndex, 1, 1, 20).setValues([[
      origDocTs,
      docId,
      docData.docType        || '',
      docData.docNo          || '',
      docData.prDate         || '',
      docData.description    || '',
      formatAmount(docData.amount),
      docData.endUser        || '',
      poTimeStamp,
      docData.poNo           || '',
      docData.poDate         || '',
      docData.supplier       || '',
      docData.requisitioner  || '',
      docData.office         || '',
      docData.status         || 'Received',
      Utilities.formatDate(dateRcvd, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      Utilities.formatDate(dueDate,  Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      calculateOverdueStatus(dateRcvd, docData.status),
      docData.notes          || '',
      docData.pdfLink        || existingPdf || ''
    ]]);
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
    logDocumentHistory(docId, actionName, docData.status || 'Received', cu ? cu.name : 'System', updateRemarks);
    logActivity('Update Document', docId, `Document updated: Status changed to ${docData.status}`);
    return { status: 'success', docId: docId };
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