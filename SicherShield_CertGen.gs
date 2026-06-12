// ═══════════════════════════════════════════════════════════════════════════
// SicherShield CertGen — Google Apps Script (standalone)
// Handles: saveCertificate, backupSettings, restoreSettings, getCertHistory
//
// DEPLOY: Extensions → Apps Script → Deploy → New deployment
//         Type: Web App, Execute as: Me, Access: Anyone
// ═══════════════════════════════════════════════════════════════════════════

// ── CONFIGURATION ───────────────────────────────────────────────────────────
const CONFIG = {
  SHEET_CERTIFICATES: 'Certificate',
  SHEET_CERT_SETTINGS:'CertSettings',
  FOLDER_CERTS:       'SicherShield_Certificates',
  FOLDER_BACKUPS:     'SicherShield_Backups',
};

// ── ENTRY POINT ──────────────────────────────────────────────────────────────

function doGet(e) {
  // Guard: e or e.parameter can be undefined when triggered without parameters
  // (e.g. direct browser open, test run from Apps Script editor)
  if (!e || !e.parameter) {
    return jsonp({ ok: true, ping: true, info: 'SicherShield CertGen GAS — no action' }, 'cb');
  }

  let p = e.parameter;
  // Frontend sends params wrapped as ?payload=<json> for GET requests
  if (p.payload) {
    try {
      const parsed = JSON.parse(p.payload);
      p = Object.assign({}, p, parsed);
    } catch(_) {}
  }
  const action = p.action || '';
  const cb     = p.callback || 'cb';

  try {
    const result = dispatch(action, p, null);
    return jsonp(result, cb);
  } catch(err) {
    return jsonp({ ok: false, error: err.message }, cb);
  }
}

function doPost(e) {
  // Guard: e or e.postData can be undefined
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'No POST body received' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(_) {}

  const action = body.action || '';
  let parsed = null;
  if (body.data) {
    try { parsed = JSON.parse(body.data); } catch(_) { parsed = body.data; }
  }

  try {
    const result = dispatch(action, body, parsed);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function dispatch(action, params, data) {
  switch(action) {
    case 'ping':             return { ok: true, ts: new Date().toISOString() };
    case 'getToken':         return { ok: true, token: Session.getActiveUser().getEmail() };

    case 'saveCertificate':  return certSave(params);
    case 'backupSettings':   return certBackupSettings(params);
    case 'restoreSettings':  return certRestoreSettings();
    case 'getCertHistory':   return certGetHistory();

    default:
      return { ok: false, error: 'Unknown action: ' + action };
  }
}

// ── HELPERS ──────────────────────────────────────────────────────────────────

function jsonp(obj, cb) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getOrCreateFolder(name, parent) {
  const base = parent || DriveApp.getRootFolder();
  const it = base.getFoldersByName(name);
  return it.hasNext() ? it.next() : base.createFolder(name);
}

function nowStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

// ═══════════════════════════════════════════════════════════════════════════
// CERTIFICATE GENERATOR FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

function certSave(p) {
  // 1. Log to Certificates sheet
  const sheet = getOrCreateSheet(CONFIG.SHEET_CERTIFICATES);

  // Write headers if empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Nr. Certificat','Prenume','Nume','Funcție','Companie','Departament',
      'Curs ID','Titlu Curs','Normă','Durată','Limbă',
      'Data Emiterii','Valabil Până','Nr. Serie',
      'Trainer','Emis de','Timestamp','Drive File ID','Drive URL'
    ]);
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, 19);
    headerRange.setBackground('#1a3a5c');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
  }

  // 2. Save PDF to Drive
  let fileId = '', fileUrl = '';
  if (p.pdfBase64) {
    try {
      const folder  = getOrCreateFolder(CONFIG.FOLDER_CERTS);
      const bytes   = Utilities.base64Decode(p.pdfBase64.replace(/^data:[^;]+;base64,/,''));
      const filename = `${p.certNo||'CERT'}_${(p.firstName||'')+' '+(p.lastName||'')}_${p.language||'XX'}.pdf`.replace(/\s+/g,'_');
      const blob    = Utilities.newBlob(bytes, 'application/pdf', filename);
      const file    = folder.createFile(blob);
      fileId  = file.getId();
      fileUrl = file.getUrl();
    } catch(err) {
      console.log('PDF save error:', err.message);
    }
  }

  // 3. Append row to sheet
  sheet.appendRow([
    p.certNo        || '',
    p.firstName     || '',
    p.lastName      || '',
    p.position      || '',
    p.company       || '',
    p.department    || '',
    p.courseId      || '',
    p.courseTitle   || '',
    p.norm          || '',
    p.duration      || '',
    p.language      || '',
    p.issueDate     || '',
    p.validUntil    || '',
    p.certNumber    || p.certNo || '',
    p.trainerId     || '',
    p.company_issuer|| '',
    nowStr(),
    fileId,
    fileUrl
  ]);

  // 4. Auto-resize columns
  try { sheet.autoResizeColumns(1, 19); } catch(_) {}

  return {
    ok:      true,
    certNo:  p.certNo,
    fileId:  fileId,
    fileUrl: fileUrl,
    ts:      nowStr()
  };
}

function certBackupSettings(p) {
  if (!p.settingsJson) return { ok: false, error: 'No settings data' };
  try {
    // Save to Sheets
    const sheet = getOrCreateSheet(CONFIG.SHEET_CERT_SETTINGS);
    sheet.clearContents();
    sheet.appendRow(['Key', 'Value', 'Updated']);
    const settings = JSON.parse(p.settingsJson);
    Object.entries(settings).forEach(([k, v]) => {
      sheet.appendRow([k, typeof v === 'object' ? JSON.stringify(v) : v, nowStr()]);
    });

    // Save to Drive as JSON backup
    const folder = getOrCreateFolder(CONFIG.FOLDER_BACKUPS);
    const blob = Utilities.newBlob(
      p.settingsJson, 'application/json',
      'CertGen_Settings_' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.json'
    );
    const file = folder.createFile(blob);
    return { ok: true, fileId: file.getId(), fileUrl: file.getUrl(), ts: nowStr() };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

function certRestoreSettings() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_CERT_SETTINGS);
    if (sheet.getLastRow() < 2) return { ok: true, settings: null };

    const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, 3).getValues();
    const settings = {};
    rows.forEach(([k, v]) => {
      if (!k) return;
      try { settings[k] = JSON.parse(v); } catch(_) { settings[k] = v; }
    });
    return { ok: true, settings: settings };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

function certGetHistory() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_CERTIFICATES);
    if (sheet.getLastRow() < 2) return { ok: true, history: [] };

    const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, 19).getValues();
    const history = rows.map(r => ({
      certNo:      r[0],
      firstName:   r[1],
      lastName:    r[2],
      position:    r[3],
      company:     r[4],
      courseId:    r[6],
      courseTitle: r[7],
      language:    r[10],
      issueDate:   r[11],
      validUntil:  r[12],
      fileUrl:     r[18]
    }));
    return { ok: true, history: history };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}
