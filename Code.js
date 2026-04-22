// ============================================================
//  BUDGET TRACKER — Google Sheets Editor Add-on
//  Menu-only (no sidebar card / homepage trigger)
//
//  ARCHITECTURE
//  ─────────────────────────────────────────────────────────
//  ScriptProperties  (shared across all users, set by Finance)
//    PERMISSIONS_SHEET_ID   Spreadsheet ID of the permissions sheet
//    COA_SHEET_ID           Spreadsheet ID of the chart of accounts
//
//  Permissions sheet layout  (Finance-maintained)
//    Col A  SuiteKey           e.g.  WaxhawIT-U_SIL
//    Col B  Viewer emails      comma-separated; everyone who can see this SuiteKey
//    Col C  Owner email        single email; the one person who can edit the budget
//
//  UserProperties  (per user, per device)
//    GL_SHEET_ID              Spreadsheet ID of the user's GL export
//    GL_TAB_NAME              Tab name within that spreadsheet
//    BUDGET_FILE_ID_{key}_FY{year}   Drive file ID of each budget sheet they own
//
//  Budget sheets
//    Created in the owner's Drive (private)
//    Columns: A=Code  B=Name  C=Group  D=Section  E=Desc  F=Budget  G=Notes
//
//  Dashboard flow
//    Level 1  — one card per authorized SuiteKey (budget vs actuals)
//    Level 2  — click SuiteKey → category breakdown (filtered by SuiteKey)
//    Level 3  — click category → individual GL transactions
// ============================================================


// ── CONSTANTS ────────────────────────────────────────────────
const FISCAL_YEAR_START_MONTH = 10;   // 10 = October

// ── EXECUTION-SCOPED CACHE ───────────────────────────────────
// These are reset for every Apps Script execution (every server call).
// They prevent redundant SpreadsheetApp.openById() and sheet reads
// within a single request.
let _glSheet       = undefined;   // cached GL sheet object
let _suiteKeyCache = undefined;   // cached getUserSuiteKeys() result
let _permData      = undefined;   // cached permissions sheet data
let _budgetCache   = {};          // { "suiteKey_FY": { rows, readonly, fileId } }

// ScriptProperties keys
const SPROP_PERMISSIONS_ID    = 'PERMISSIONS_SHEET_ID';
const SPROP_COA_ID            = 'COA_SHEET_ID';
const SPROP_PERMISSIONS_CACHE = 'PERMISSIONS_CACHE';

// UserProperties keys
const UPROP_GL_SHEET_ID    = 'GL_SHEET_ID';
const UPROP_GL_TAB_NAME    = 'GL_TAB_NAME';
const BUDGET_PROP_PREFIX   = 'BUDGET_FILE_ID_'; // + suitekey + _FY + year

// Budget sheet column indexes (0-based) — never reorder
const BC_CODE    = 0;  // A
const BC_NAME    = 1;  // B
const BC_GROUP   = 2;  // C
const BC_SECTION = 3;  // D
const BC_DESC    = 4;  // E
const BC_BUDGET  = 5;  // F  ← editable
const BC_NOTES   = 6;  // G

// GL export column indexes (0-based) — NetSuite GL Multicurrency
const TXCOL_HDR_ACCOUNT = 0;
const TXCOL_CODE        = 1;
const TXCOL_TYPE        = 2;
const TXCOL_POST_DATE   = 3;
const TXCOL_TXN_DATE    = 4;
const TXCOL_DOC_NUM     = 5;
const TXCOL_DESC        = 6;
const TXCOL_NAME        = 7;
const TXCOL_COST_CENTER = 8;
const TXCOL_SUITEKEY    = 10;  // Col K — ICJE Suite Key
const TXCOL_DEBIT       = 11;
const TXCOL_CREDIT      = 12;
const TXCOL_LOCAL_AMT   = 14;


// ============================================================
//  WEB APP ENTRY POINT
// ============================================================

/**
 * Serves the Budget Tracker as a web app.
 * Deploy as: Execute as = User accessing the web app
 *            Who has access = Anyone in your organization
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Budget Tracker')
    .setFaviconUrl('https://raw.githubusercontent.com/kode-monki/budget-tracker/main/budget-tracker-icon.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Single batched call that returns everything the dashboard needs
 * on first load — eliminates 2-3 sequential round trips.
 */
function getInitialData() {
  const t0 = Date.now();
  const log = [];

  try {
    log.push('start');

    // Step 1: user properties (fast — no network)
    const up = PropertiesService.getUserProperties();
    const sp = PropertiesService.getScriptProperties();
    log.push('props:' + (Date.now()-t0) + 'ms');

    // Step 2: current user email
    let userEmail = '';
    try { userEmail = Session.getEffectiveUser().getEmail(); } catch(e) {
      try { userEmail = Session.getActiveUser().getEmail(); } catch(e2) { userEmail = 'unknown'; }
    }
    log.push('email:' + (Date.now()-t0) + 'ms');

    // Step 3: permissions data (from ScriptProperties cache or direct sheet read)
    let suiteKeys = [], skError = null;
    console.log('getInitialData: step3 permissions permId=' + (sp.getProperty(SPROP_PERMISSIONS_ID)||'(none)'));
    const permResult = getPermissionsData();
    log.push('perm:' + (Date.now()-t0) + 'ms fromCache:' + permResult.fromCache);
    if (permResult.data) {
      const pdata  = permResult.data;
      const uEmail = userEmail.toLowerCase();
      for (let i = 1; i < pdata.length; i++) {
        const row   = pdata[i];
        const sk    = String(row[0]||'').trim();
        const owner = String(row[2]||'').trim().toLowerCase();
        if (!sk) continue;
        const viewers = String(row[1]||'').split(',').map(e=>e.trim().toLowerCase()).filter(Boolean);
        if (owner === uEmail || viewers.includes(uEmail)) {
          suiteKeys.push({ suiteKey: sk, isOwner: owner===uEmail, ownerEmail: owner, viewers });
        }
      }
      _suiteKeyCache = { error: null, suiteKeys };
      log.push('perm_done:' + (Date.now()-t0) + 'ms suitekeys:' + suiteKeys.length);
    } else {
      skError = permResult.error || 'cannot_open_permissions';
      _suiteKeyCache = { error: skError, suiteKeys: [] };
      log.push('perm_err:' + skError);
    }

    // Step 4: GL sheet — open and cache data
    const glId  = up.getProperty(UPROP_GL_SHEET_ID)  || '';
    const glTab = up.getProperty(UPROP_GL_TAB_NAME)   || '';
    const fySet = new Set([getCurrentFiscalYear()]);
    const mSet  = new Set();
    let   glError = null;

    if (glId) {
      console.log('getInitialData: step4 opening GL sheet ' + glId + ' tab=' + (glTab||'(first)'));
      try {
        const gss = SpreadsheetApp.openById(glId);
        log.push('gl_open:' + (Date.now()-t0) + 'ms');
        const glSheet = glTab ? (gss.getSheetByName(glTab) || gss.getSheets()[0]) : gss.getSheets()[0];
        _glSheet = glSheet;
        const glData = glSheet.getDataRange().getValues();
        _glData = glData;
        log.push('gl_read:' + (Date.now()-t0) + 'ms rows:' + glData.length);
        for (let i = 0; i < glData.length; i++) {
          const colA = String(glData[i][TXCOL_HDR_ACCOUNT]||'').trim();
          const colB = String(glData[i][TXCOL_CODE]||'').trim();
          if (colA||!colB||isSkipRow(colB)) continue;
          const d = parseGLDate(glData[i][TXCOL_POST_DATE]);
          if (!d) continue;
          fySet.add(getFiscalYear(d));
          mSet.add(d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'));
        }
        log.push('gl_parse:' + (Date.now()-t0) + 'ms');
      } catch(e) {
        _glSheet = null; _glData = null;
        glError = { id: glId, message: e.message };
        log.push('gl_err:' + e.message);
      }
    } else {
      log.push('no_gl_sheet');
    }

    // Step 5: build periods
    const MONTHS = ['January','February','March','April','May','June',
                    'July','August','September','October','November','December'];
    const props2 = up.getProperties();
    Object.keys(props2).forEach(k => {
      const m = k.match(/BUDGET_FILE_ID_.+_FY(\d{4})$/);
      if (m) fySet.add(parseInt(m[1]));
    });
    const periods = {
      fiscalYears : [...fySet].sort().reverse(),
      months      : [...mSet].sort().reverse().map(key => {
        const [yr,mo] = key.split('-');
        return { value: key, label: MONTHS[parseInt(mo)-1] + ' ' + yr };
      }),
      currentFY   : getCurrentFiscalYear(),
    };

    // Step 6: settings
    const settings = {
      permissionsSheetId : sp.getProperty(SPROP_PERMISSIONS_ID) || '',
      coaSheetId         : sp.getProperty(SPROP_COA_ID) || '',
      glSheetId          : glId,
      glTabName          : glTab,
      userEmail,
    };

    log.push('done:' + (Date.now()-t0) + 'ms');
    console.log('getInitialData: ' + log.join(' | '));

    return { periods, settings, suiteKeys, skError, glError, currentFY: periods.currentFY, _log: log };

  } catch(e) {
    console.error('getInitialData FATAL: ' + e.message + ' | log: ' + log.join(' | '));
    return {
      periods   : { fiscalYears:[getCurrentFiscalYear()], months:[], currentFY:getCurrentFiscalYear() },
      settings  : {},
      suiteKeys : [],
      skError   : 'fatal_error',
      currentFY : getCurrentFiscalYear(),
      _error    : e.message,
      _log      : log,
    };
  }
}

/**
 * Instant ping — returns immediately with no API calls.
 * Used to verify google.script.run is working before anything else.
 */
function ping() {
  return { ok: true, ts: new Date().toISOString() };
}

/**
 * Minimal init — only reads Properties, no Sheets API calls at all.
 * Call this first to confirm the script itself runs fast.
 */
function getInitialDataFast() {
  const t0 = Date.now();
  try {
    const sp = PropertiesService.getScriptProperties();
    const up = PropertiesService.getUserProperties();
    const permId = sp.getProperty(SPROP_PERMISSIONS_ID) || '';
    const coaId  = sp.getProperty(SPROP_COA_ID) || '';
    const glId   = up.getProperty(UPROP_GL_SHEET_ID) || '';
    const glTab  = up.getProperty(UPROP_GL_TAB_NAME) || '';
    return {
      ok: true,
      ms: Date.now() - t0,
      hasPermSheet: !!permId,
      hasGLSheet:   !!glId,
      hasCOA:       !!coaId,
      permId, glId, glTab,
    };
  } catch(e) {
    return { ok: false, error: e.message, ms: Date.now() - t0 };
  }
}

/**
 * Helper called from Index.html to inline CSS/JS files if needed.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ============================================================
//  PERMISSIONS
// ============================================================

/**
 * Returns the current user's email.
 */
function getCurrentUserEmail() {
  try {
    return (Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail()).toLowerCase();
  } catch(e) {
    return Session.getActiveUser().getEmail().toLowerCase();
  }
}

/**
 * Reads the permissions sheet and returns all SuiteKeys the current
 * user is authorised to see, plus whether they are the owner of each.
 *
 * Returns: [{ suiteKey, isOwner, ownerEmail, viewers }]
 */
function getUserSuiteKeys() {
  if (_suiteKeyCache !== undefined) return _suiteKeyCache;

  const permResult = getPermissionsData();
  if (!permResult.data) {
    _suiteKeyCache = { error: permResult.error || 'cannot_open_permissions', suiteKeys: [] };
    return _suiteKeyCache;
  }

  const data      = permResult.data;
  const userEmail = getCurrentUserEmail();
  const results   = [];

  for (let i = 1; i < data.length; i++) {
    const row        = data[i];
    const suiteKey   = String(row[0] || '').trim();
    const viewersRaw = String(row[1] || '');
    const ownerEmail = String(row[2] || '').trim().toLowerCase();
    if (!suiteKey) continue;
    const viewers  = viewersRaw.split(',').map(e => e.trim().toLowerCase()).filter(Boolean);
    const isOwner  = ownerEmail === userEmail;
    const isViewer = viewers.includes(userEmail);
    if (isOwner || isViewer) {
      results.push({ suiteKey, isOwner, ownerEmail, viewers });
    }
  }

  _suiteKeyCache = { error: null, suiteKeys: results };
  return _suiteKeyCache;
}


// ============================================================
//  SETTINGS  (admin + per-user)
// ============================================================

/**
 * Reads the permissions sheet and stores it in ScriptProperties so end users
 * never need direct access to the sheet. Finance runs this after any change.
 * The sheet can then be restricted to Finance only.
 */
function syncPermissionsCache() {
  const sp     = PropertiesService.getScriptProperties();
  const permId = sp.getProperty(SPROP_PERMISSIONS_ID);
  if (!permId) return { ok: false, error: 'No permissions sheet ID saved in Settings.' };
  try {
    const data = SpreadsheetApp.openById(permId).getSheets()[0].getDataRange().getValues();
    const json = JSON.stringify(data);
    sp.setProperty(SPROP_PERMISSIONS_CACHE, json);
    return { ok: true, rows: Math.max(0, data.length - 1) };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

/**
 * Returns the raw permissions sheet data from ScriptProperties cache.
 * Falls back to reading the sheet directly if the cache is empty (first-run).
 */
function getPermissionsData() {
  const sp     = PropertiesService.getScriptProperties();
  const cached = sp.getProperty(SPROP_PERMISSIONS_CACHE);
  if (cached) {
    try { return { data: JSON.parse(cached), fromCache: true }; } catch(e) {}
  }
  // Cache empty — try direct read (requires sheet to be shared)
  const permId = sp.getProperty(SPROP_PERMISSIONS_ID);
  if (!permId) return { data: null, error: 'no_permissions_sheet' };
  try {
    const data = SpreadsheetApp.openById(permId).getSheets()[0].getDataRange().getValues();
    return { data, fromCache: false };
  } catch(e) {
    return { data: null, error: 'cannot_open_permissions' };
  }
}

/** Returns all settings needed to render the Settings tab. */
function getAllSettings() {
  const sp  = PropertiesService.getScriptProperties();
  const up  = PropertiesService.getUserProperties();
  return {
    // Admin (ScriptProperties — anyone can read, Finance sets these)
    permissionsSheetId : sp.getProperty(SPROP_PERMISSIONS_ID) || '',
    coaSheetId         : sp.getProperty(SPROP_COA_ID)         || '',
    // Per-user (UserProperties)
    glSheetId          : up.getProperty(UPROP_GL_SHEET_ID)    || '',
    glTabName          : up.getProperty(UPROP_GL_TAB_NAME)     || '',
    userEmail          : getCurrentUserEmail(),
  };
}

/** Saves admin-level settings to ScriptProperties. Finance only. */
function saveAdminSettings(permissionsSheetId, coaSheetId) {
  permissionsSheetId = String(permissionsSheetId || '').trim();
  coaSheetId         = String(coaSheetId         || '').trim();

  if (permissionsSheetId) {
    try { SpreadsheetApp.openById(permissionsSheetId); }
    catch(e) { return { ok: false, error: 'Cannot open permissions spreadsheet: ' + e.message }; }
  }
  if (coaSheetId) {
    try { SpreadsheetApp.openById(coaSheetId); }
    catch(e) { return { ok: false, error: 'Cannot open COA spreadsheet: ' + e.message }; }
  }

  const sp = PropertiesService.getScriptProperties();
  if (permissionsSheetId) sp.setProperty(SPROP_PERMISSIONS_ID, permissionsSheetId);
  if (coaSheetId)         sp.setProperty(SPROP_COA_ID,         coaSheetId);
  return { ok: true };
}

/** Saves per-user GL sheet settings. */
function saveUserSettings(glSheetId, glTabName) {
  glSheetId = String(glSheetId || '').trim();
  glTabName = String(glTabName || '').trim();

  if (glSheetId) {
    try {
      const ss    = SpreadsheetApp.openById(glSheetId);
      const sheet = glTabName ? ss.getSheetByName(glTabName) : ss.getSheets()[0];
      if (!sheet) return { ok: false, error: 'Tab "' + glTabName + '" not found.' };
    } catch(e) {
      return { ok: false, error: 'Cannot open GL spreadsheet: ' + e.message };
    }
  }

  const up = PropertiesService.getUserProperties();
  up.setProperty(UPROP_GL_SHEET_ID, glSheetId);
  up.setProperty(UPROP_GL_TAB_NAME, glTabName);
  return { ok: true };
}

/** Returns tab names for any spreadsheet ID (used for GL picker). */
function getSheetTabs(sheetId) {
  sheetId = String(sheetId || '').trim();
  if (!sheetId) return { ok: false, tabs: [] };
  try {
    return { ok: true, tabs: SpreadsheetApp.openById(sheetId).getSheets().map(s => s.getName()) };
  } catch(e) {
    return { ok: false, tabs: [], error: e.message };
  }
}

/** Extracts spreadsheet ID from a full URL or raw ID string. */
function parseSheetId(input) {
  const m = String(input || '').match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : String(input || '').trim();
}


// ============================================================
//  BUDGET SHEET MANAGEMENT  (per SuiteKey, per FY)
// ============================================================

/**
 * Returns budget sheet status for all SuiteKeys the user can access,
 * for the given fiscal year.
 */
function getBudgetStatus(fiscalYear) {
  const fy       = parseInt(fiscalYear);
  const { suiteKeys, error } = getUserSuiteKeys();
  const up       = PropertiesService.getUserProperties();
  const results  = [];

  suiteKeys.forEach(sk => {
    const propKey = BUDGET_PROP_PREFIX + sk.suiteKey + '_FY' + fy;
    const fileId  = up.getProperty(propKey);
    let   exists  = false, name = null, url = null;

    if (fileId) {
      try {
        const f = DriveApp.getFileById(fileId);
        exists  = true;
        name    = f.getName();
        url     = f.getUrl();
      } catch(e) {
        // File deleted — clear stale property
        up.deleteProperty(propKey);
      }
    }

    results.push({
      suiteKey : sk.suiteKey,
      isOwner  : sk.isOwner,
      ownerEmail: sk.ownerEmail,
      exists, fileId: exists ? fileId : null, name, url,
    });
  });

  return { budgets: results, error, currentFY: getCurrentFiscalYear() };
}

/**
 * Creates a new budget sheet for the given SuiteKey + FY.
 * Only the owner of the SuiteKey can create/edit its budget.
 */
function createBudgetForSuiteKey(suiteKey, fiscalYear) {
  const fy = parseInt(fiscalYear);

  // Verify ownership
  const { suiteKeys } = getUserSuiteKeys();
  const sk = suiteKeys.find(s => s.suiteKey === suiteKey);
  if (!sk)        throw new Error('You do not have access to SuiteKey: ' + suiteKey);
  if (!sk.isOwner) throw new Error('Only the owner (' + sk.ownerEmail + ') can create a budget for ' + suiteKey);

  const propKey = BUDGET_PROP_PREFIX + suiteKey + '_FY' + fy;
  const up      = PropertiesService.getUserProperties();
  const existing = up.getProperty(propKey);

  if (existing) {
    try {
      const f = DriveApp.getFileById(existing);
      return { fileId: existing, name: f.getName(), url: f.getUrl(), alreadyExisted: true };
    } catch(e) { up.deleteProperty(propKey); }
  }

  const name = 'Budget — ' + suiteKey + ' — FY' + fy;
  let ss;
  try {
    ss = SpreadsheetApp.create(name);
  } catch(e) {
    if (e.message && e.message.toLowerCase().includes('permission')) {
      throw new Error(
        'Cannot create the budget sheet — your authorization needs to be refreshed. ' +
        'Open the app URL in a new browser tab; if prompted, click Allow to re-authorize. ' +
        'If the problem continues, ask IT to check that the app is allowed to create Google Sheets in your Workspace Admin Console. ' +
        '(Original error: ' + e.message + ')'
      );
    }
    throw e;
  }
  const sheet= ss.getActiveSheet();
  sheet.setName('Budget FY' + fy);

  // Reference data for new columns
  const priorData = readBudgetSheet(suiteKey, fy - 1);
  const priorMap  = {};
  (priorData.rows || []).forEach(r => { priorMap[r.code] = Math.abs(r.budget); });
  const ytdMap = calcTotalsForSuiteKey(suiteKey, fy, null, 'fiscal_year', null, null);

  // Header row
  const headers = ['Code', 'Account name', 'Group', 'Section', 'Description', 'FY' + fy + ' Budget', 'Notes', 'FY' + (fy - 1) + ' Budget', 'YTD Actuals'];
  sheet.getRange(1, 1, 1, headers.length)
       .setValues([headers])
       .setBackground('#1a73e8')
       .setFontColor('#ffffff')
       .setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Data rows from COA
  const accounts = getCanonicalAccounts();
  const rows     = [];
  let lastSection = '';

  accounts.forEach(a => {
    if (a.section !== lastSection) {
      rows.push(['', a.section.toUpperCase(), '', '', '', '', '', '', '']);
      lastSection = a.section;
    }
    const prior = priorMap[a.code] !== undefined ? priorMap[a.code] : '';
    const ytd   = Math.abs(ytdMap[a.code] || 0) || '';
    rows.push([a.code, a.name, a.group, a.section, a.desc, 0, '', prior, ytd]);
  });

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, BC_BUDGET + 1, rows.length, 1).setNumberFormat('$#,##0.00');
    // Format reference columns as currency, light gray to indicate read-only
    const refCols = [headers.length - 1, headers.length]; // FY prior + YTD
    refCols.forEach(col => {
      sheet.getRange(2, col, rows.length, 1)
           .setNumberFormat('$#,##0.00')
           .setBackground('#f8f9fa')
           .setFontColor('#5f6368');
    });

    // Style section divider rows
    rows.forEach((r, i) => {
      if (!r[0] && r[1]) {
        sheet.getRange(i + 2, 1, 1, headers.length)
             .setBackground('#f1f3f4').setFontWeight('bold').setFontColor('#5f6368');
      }
    });
  }

  // Column widths
  [70, 240, 130, 80, 320, 130, 200, 130, 120].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  up.setProperty(propKey, ss.getId());
  return { fileId: ss.getId(), name, url: ss.getUrl(), alreadyExisted: false };
}

/**
 * Reads budget rows from a user's own budget sheet for a SuiteKey+FY.
 * Returns rows with a `readonly` flag if the caller is not the owner.
 */
function readBudgetSheet(suiteKey, fiscalYear) {
  const fy       = parseInt(fiscalYear);
  const cacheKey = suiteKey + '_FY' + fy;
  if (_budgetCache[cacheKey]) return _budgetCache[cacheKey];

  const up     = PropertiesService.getUserProperties();
  let   fileId = up.getProperty(BUDGET_PROP_PREFIX + suiteKey + '_FY' + fy);

  if (!fileId) {
    fileId = PropertiesService.getScriptProperties()
               .getProperty('SHARED_' + BUDGET_PROP_PREFIX + suiteKey + '_FY' + fy);
  }

  if (!fileId) {
    _budgetCache[cacheKey] = { rows: [], readonly: true, missing: true };
    return _budgetCache[cacheKey];
  }

  let ss;
  try { ss = SpreadsheetApp.openById(fileId); }
  catch(e) {
    const detail = suiteKey + ' FY' + fy + ' (file ID: ' + fileId + '): ' + e.message;
    _budgetCache[cacheKey] = { rows: [], readonly: true, missing: true, error: detail };
    return _budgetCache[cacheKey];
  }

  const { suiteKeys } = getUserSuiteKeys();  // cached — no extra round trip
  const sk       = suiteKeys.find(s => s.suiteKey === suiteKey);
  const readonly = !sk || !sk.isOwner;

  const data = ss.getSheets()[0].getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const code = String(row[BC_CODE] || '').trim();
    if (!code) continue;
    rows.push({
      code,
      name    : String(row[BC_NAME]    || '').trim(),
      group   : String(row[BC_GROUP]   || '').trim(),
      section : String(row[BC_SECTION] || '').trim(),
      desc    : String(row[BC_DESC]    || '').trim(),
      budget  : parseMoney(row[BC_BUDGET]),
      notes   : String(row[BC_NOTES]   || '').trim(),
      fullName: code + ': ' + String(row[BC_NAME] || '').trim(),
    });
  }

  _budgetCache[cacheKey] = { rows, readonly, fileId };
  return _budgetCache[cacheKey];
}

/** Returns budget rows plus prior-year budget map and YTD actuals for the Budget tab. */
function getBudgetPageData(suiteKey, fiscalYear) {
  const fy      = parseInt(fiscalYear);
  const current = readBudgetSheet(suiteKey, fy);
  const prior   = readBudgetSheet(suiteKey, fy - 1);
  const ytd     = calcTotalsForSuiteKey(suiteKey, fiscalYear, null, 'fiscal_year', null, null);

  const priorBudget = {};
  (prior.rows || []).forEach(r => { priorBudget[r.code] = Math.abs(r.budget); });

  return {
    rows        : current.rows,
    readonly    : current.readonly,
    missing     : current.missing,
    fileId      : current.fileId,
    priorBudget,
    ytdActuals  : ytd,
  };
}

/** Saves edited budget amounts. Only works if the caller is the owner. */
function saveBudgetAmounts(suiteKey, fiscalYear, updates) {
  const fy  = parseInt(fiscalYear);
  const up  = PropertiesService.getUserProperties();
  const key = BUDGET_PROP_PREFIX + suiteKey + '_FY' + fy;
  const fileId = up.getProperty(key);
  if (!fileId) throw new Error('No budget sheet found for ' + suiteKey + ' FY' + fy);

  // Verify ownership
  const { suiteKeys } = getUserSuiteKeys();
  const sk = suiteKeys.find(s => s.suiteKey === suiteKey);
  if (!sk || !sk.isOwner) throw new Error('Only the owner can edit this budget.');

  const sheet = SpreadsheetApp.openById(fileId).getSheets()[0];
  const data  = sheet.getDataRange().getValues();
  const map   = {};
  updates.forEach(u => map[String(u.code)] = u.amount);

  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][BC_CODE] || '').trim();
    if (code && map[code] !== undefined) {
      sheet.getRange(i + 1, BC_BUDGET + 1).setValue(map[code]);
    }
  }

  // Share file ID in ScriptProperties so viewers can read it
  PropertiesService.getScriptProperties()
    .setProperty('SHARED_' + key, fileId);

  return { ok: true, saved: updates.length };
}

/**
 * Renames a budget sheet and removes it from the add-on interface
 * (clears the property) without deleting the Drive file.
 */
function removeBudgetFromAddon(suiteKey, fiscalYear) {
  const fy  = parseInt(fiscalYear);
  const up  = PropertiesService.getUserProperties();
  const key = BUDGET_PROP_PREFIX + suiteKey + '_FY' + fy;
  const fileId = up.getProperty(key);

  if (fileId) {
    try {
      const file    = DriveApp.getFileById(fileId);
      const oldName = file.getName();
      file.setName('[Archived] ' + oldName);
    } catch(e) { /* file already gone */ }
    up.deleteProperty(key);
    // Also remove shared reference
    PropertiesService.getScriptProperties()
      .deleteProperty('SHARED_' + key);
  }
  return { ok: true };
}


// ============================================================
//  DASHBOARD DATA API
// ============================================================

/**
 * Top-level dashboard: returns budget vs actuals for each
 * SuiteKey the user is authorized to see.
 */
function getDashboardSummary(fiscalYear, period, periodType, dateFrom, dateTo) {
  const fy            = parseInt(fiscalYear);
  const { suiteKeys } = getUserSuiteKeys();
  const actuals       = calcAllTotals(fiscalYear, period, periodType, dateFrom, dateTo);
  const results       = [];

  suiteKeys.forEach(sk => {
    const { rows, readonly, missing } = readBudgetSheet(sk.suiteKey, fy);
    const skActuals = actuals[sk.suiteKey] || {};

    let totalBudget  = 0;
    let totalActuals = 0;

    rows.forEach(r => {
      totalBudget  += Math.abs(r.budget);
      totalActuals += Math.abs(skActuals[r.code] || 0);
    });

    const pct = totalBudget ? Math.round(totalActuals / totalBudget * 100) : 0;
    results.push({
      suiteKey    : sk.suiteKey,
      isOwner     : sk.isOwner,
      ownerEmail  : sk.ownerEmail,
      totalBudget,
      totalActuals,
      remaining   : totalBudget - totalActuals,
      pct,
      hasBudget   : !missing && rows.length > 0,
      readonly,
    });
  });

  return results;
}

/**
 * Level-2 drill: returns category breakdown for one SuiteKey.
 */
function getSuiteKeyDetail(suiteKey, fiscalYear, period, periodType, dateFrom, dateTo) {
  const fy            = parseInt(fiscalYear);
  const { rows, readonly } = readBudgetSheet(suiteKey, fy);
  const actuals       = calcTotalsForSuiteKey(suiteKey, fiscalYear, period, periodType, dateFrom, dateTo);

  return {
    suiteKey,
    readonly,
    categories: rows.map(r => {
      const spent     = Math.abs(actuals[r.code] || 0);
      const budgetAbs = Math.abs(r.budget);
      const remaining = budgetAbs - spent;
      const pct       = budgetAbs ? Math.round(spent / budgetAbs * 100) : 0;
      return { ...r, actuals: spent, remaining, pct };
    }),
  };
}

/**
 * Level-3 drill: returns individual GL transactions for one
 * account code within one SuiteKey.
 */
function getTransactions(suiteKey, accountCode, fiscalYear, period, periodType, dateFrom, dateTo) {
  const data = getGLData();
  if (!data) return [];
  const result = [];
  let   currentSectionCode = null;

  for (let i = 0; i < data.length; i++) {
    const row  = data[i];
    const colA = String(row[TXCOL_HDR_ACCOUNT] || '').trim();
    const colB = String(row[TXCOL_CODE]         || '').trim();

    if (colA && !colA.startsWith('Total') && !colB) {
      currentSectionCode = normalizeCode(extractCode(colA));
      continue;
    }
    if (colA.startsWith('Total')) continue;

    if (!colA && colB && !isSkipRow(colB)) {
      const txSuiteKey  = String(row[TXCOL_SUITEKEY] || '').trim();
      if (txSuiteKey !== suiteKey) continue;

      const sectionCode = normalizeCode(currentSectionCode || colB);
      if (sectionCode !== accountCode) continue;

      const d = parseGLDate(row[TXCOL_POST_DATE]);
      if (!d || !inPeriod(d, fiscalYear, period, periodType, dateFrom, dateTo)) continue;

      const txnDate = parseGLDate(row[TXCOL_TXN_DATE]);
      result.push({
        postedDate : formatDate(d),
        txnDate    : formatDate(txnDate || d),
        docNum     : String(row[TXCOL_DOC_NUM]     || ''),
        desc       : String(row[TXCOL_DESC]        || ''),
        name       : String(row[TXCOL_NAME]        || ''),
        costCenter : String(row[TXCOL_COST_CENTER] || ''),
        type       : String(row[TXCOL_TYPE]        || ''),
        amount     : calcAmount(row, sectionCode),
      });
    }
  }
  return result;
}

/** Returns available fiscal years from GL data + any saved budgets. */
function getAvailablePeriods() {
  const fySet = new Set();
  const mSet  = new Set();

  const glData = getGLData();
  if (glData) {
    for (let i = 0; i < glData.length; i++) {
      const colA = String(glData[i][TXCOL_HDR_ACCOUNT] || '').trim();
      const colB = String(glData[i][TXCOL_CODE]         || '').trim();
      if (colA || !colB || isSkipRow(colB)) continue;
      const d = parseGLDate(glData[i][TXCOL_POST_DATE]);
      if (!d) continue;
      fySet.add(getFiscalYear(d));
      mSet.add(d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'));
    }
  }

  const curFY = getCurrentFiscalYear();
  fySet.add(curFY);

  // Also include FYs that have saved budgets
  const up    = PropertiesService.getUserProperties();
  const props = up.getProperties();
  Object.keys(props).forEach(k => {
    const m = k.match(/BUDGET_FILE_ID_.+_FY(\d{4})$/);
    if (m) fySet.add(parseInt(m[1]));
  });

  const MONTHS = ['January','February','March','April','May','June',
                  'July','August','September','October','November','December'];

  return {
    fiscalYears : [...fySet].sort().reverse(),
    months      : [...mSet].sort().reverse().map(key => {
      const [yr, mo] = key.split('-');
      return { value: key, label: MONTHS[parseInt(mo) - 1] + ' ' + yr };
    }),
    currentFY   : curFY,
  };
}


// ============================================================
//  GL PARSER
// ============================================================

/**
 * Returns totals keyed as { suiteKey: { accountCode: amount } }
 * across ALL SuiteKeys the user can see.
 */
// Module-level GL data cache — read the sheet once per execution
let _glData = undefined;
function getGLData() {
  if (_glData !== undefined) return _glData;
  const sheet = getGLSheet();
  _glData = sheet ? sheet.getDataRange().getValues() : null;
  return _glData;
}

function calcAllTotals(fiscalYear, period, periodType, dateFrom, dateTo) {
  const data = getGLData();
  if (!data) return {};
  const totals = {};
  let   currentSectionCode = null;

  for (let i = 0; i < data.length; i++) {
    const row  = data[i];
    const colA = String(row[TXCOL_HDR_ACCOUNT] || '').trim();
    const colB = String(row[TXCOL_CODE]         || '').trim();

    if (colA && !colA.startsWith('Total') && !colB) {
      currentSectionCode = normalizeCode(extractCode(colA));
      continue;
    }
    if (colA.startsWith('Total')) continue;

    if (!colA && colB && !isSkipRow(colB)) {
      const d = parseGLDate(row[TXCOL_POST_DATE]);
      if (!d || !inPeriod(d, fiscalYear, period, periodType, dateFrom, dateTo)) continue;

      const suiteKey = String(row[TXCOL_SUITEKEY] || '').trim();
      if (!suiteKey) continue;

      const code   = currentSectionCode || normalizeCode(colB);
      const amount = calcAmount(row, code);

      if (!totals[suiteKey])        totals[suiteKey]       = {};
      if (!totals[suiteKey][code])  totals[suiteKey][code] = 0;
      totals[suiteKey][code] += amount;
    }
  }
  return totals;
}

/** Returns totals for a single SuiteKey: { accountCode: amount } */
function calcTotalsForSuiteKey(suiteKey, fiscalYear, period, periodType, dateFrom, dateTo) {
  const all = calcAllTotals(fiscalYear, period, periodType, dateFrom, dateTo);
  return all[suiteKey] || {};
}

let _accountTypeCache = null;

function getAccountTypeMap() {
  if (_accountTypeCache) return _accountTypeCache;
  const accounts = getCanonicalAccounts();
  const map = {};
  accounts.forEach(a => {
    map[a.code] = a.type || (a.section === 'Income' ? 'income' : 'expense');
  });
  _accountTypeCache = map;
  return map;
}

function calcAmount(row, code) {
  const typeMap = getAccountTypeMap();
  const type    = typeMap[code];
  const debit   = parseMoney(row[TXCOL_DEBIT]);
  const credit  = parseMoney(row[TXCOL_CREDIT]);
  const local   = parseMoney(row[TXCOL_LOCAL_AMT]);
  if (type === 'income')  return credit - debit;
  if (type === 'expense') return debit  - credit;
  // Fallback for codes not in CoA
  const n = parseInt(code, 10);
  if (n >= 4000 && n <= 6999) return credit - debit;
  if (n >= 7000 && n <= 9999) return debit  - credit;
  return local;
}

function getGLSheet() {
  if (_glSheet !== undefined) return _glSheet;
  const up  = PropertiesService.getUserProperties();
  const id  = up.getProperty(UPROP_GL_SHEET_ID);
  const tab = up.getProperty(UPROP_GL_TAB_NAME);
  if (!id) { _glSheet = null; return null; }
  try {
    const ss = SpreadsheetApp.openById(id);
    _glSheet = tab ? (ss.getSheetByName(tab) || ss.getSheets()[0]) : ss.getSheets()[0];
  } catch(e) { _glSheet = null; }
  return _glSheet;
}


// ============================================================
//  CHART OF ACCOUNTS
// ============================================================

function getCanonicalAccounts() {
  const coaId = PropertiesService.getScriptProperties().getProperty(SPROP_COA_ID);
  if (coaId) {
    try {
      const ss    = SpreadsheetApp.openById(coaId);
      const sheet = ss.getSheets()[0];
      return parseCOASheet(sheet);
    } catch(e) {
      console.warn('COA sheet unavailable: ' + e.message);
    }
  }
  return getBuiltinAccounts();
}

function isIncomeCategory(cat) {
  const c = String(cat || '').toLowerCase().trim();
  return c.includes('income') || c.includes('sales') || c.includes('revenue') ||
         c.includes('contribut') || c.includes('grant');
}

function parseCOASheet(sheet) {
  const data    = sheet.getDataRange().getValues();
  const results = [];
  const seen    = new Set();

  for (let i = 0; i < data.length; i++) {
    const row     = data[i];
    const rawCat  = String(row[0] || '').trim();  // Col A — NetSuite Category
    const rawCode = String(row[1] || '').trim();  // Col B — Account Code
    const rawName = String(row[2] || '').trim();  // Col C — Account Description
    const rawDesc = String(row[3] || '').trim();  // Col D — Description (tooltip)

    if (!rawCode || rawCode.toLowerCase() === 'account code') continue;
    const codeNum = parseInt(rawCode, 10);
    if (isNaN(codeNum) || codeNum < 4000)  continue;
    if (/summary/i.test(rawName))          continue;

    const codeStr = String(codeNum);
    if (seen.has(codeStr)) continue;
    seen.add(codeStr);

    const isIncome = isIncomeCategory(rawCat);
    results.push({
      code     : codeStr,
      name     : rawName || ('Account ' + codeStr),
      group    : deriveGroupFromCategory(rawCat, codeNum),
      section  : isIncome ? 'Income' : 'Expense',
      type     : isIncome ? 'income' : 'expense',
      desc     : rawDesc,
      fullName : codeStr + ': ' + (rawName || ''),
    });
  }

  results.sort((a, b) => parseInt(a.code) - parseInt(b.code));
  return results;
}

function deriveGroupFromCategory(category, codeNum) {
  const cat = String(category || '').toLowerCase().trim();
  const map = {
    'income':'Income','sales & service income':'Sales & Service',
    'contributions':'Contributions','other income':'Other Income',
    'cost of goods sold':'Cost of Goods Sold','expense':'Expense',
    'other expense':'Other Expense','payroll expense':'Personnel',
    'personnel':'Personnel','labor':'Personnel',
    'facilities':'Facilities','facilities, equipment, & maintenance':'Facilities',
    'technology':'Technology','information technology':'Technology',
    'operations':'Operations','program expense':'Program','program':'Program',
    'depreciation':'Depreciation','amortization':'Depreciation',
    'internal':'Internal','intercompany':'Internal',
  };
  if (map[cat]) return map[cat];
  if (cat.includes('payroll')||cat.includes('personnel')||cat.includes('labor')) return 'Personnel';
  if (cat.includes('facilit'))   return 'Facilities';
  if (cat.includes('technolog')||cat.includes('computer')) return 'Technology';
  if (cat.includes('program'))   return 'Program';
  if (cat.includes('depreciat')||cat.includes('amortiz'))  return 'Depreciation';
  if (cat.includes('internal')||cat.includes('interco'))   return 'Internal';
  if (cat.includes('income')||cat.includes('revenue')||cat.includes('sales')) return 'Sales & Service';
  if (cat.includes('contribut')||cat.includes('grant'))    return 'Contributions';
  return deriveGroup(String(codeNum));
}

function getBuiltinAccounts() {
  const accounts = [
    { code:'4000', name:'Contributions',                     group:'Contributions',   section:'Income',  desc:'Donor contributions and grants' },
    { code:'5010', name:'Product Sales',                     group:'Sales & Service', section:'Income',  desc:'Revenue from product sales' },
    { code:'5110', name:'Service Income - Non-Professional', group:'Sales & Service', section:'Income',  desc:'Non-professional service fees' },
    { code:'5200', name:'Subscriptions',                     group:'Sales & Service', section:'Income',  desc:'Subscription and SaaS revenue' },
    { code:'5400', name:'Tuition / Fees',                    group:'Sales & Service', section:'Income',  desc:'Tuition, registration, and program fees' },
    { code:'8016', name:'Internal Non-Inventory Sales',      group:'Internal Income', section:'Income',  desc:'Internal sales of non-inventory items to other OUs' },
    { code:'8026', name:'Internal Non-Inventory Sales to SIL OU', group:'Internal Income', section:'Income', desc:'Internal sales to other SIL OUs' },
    { code:'7130', name:'Employee Salaries',                 group:'Personnel',       section:'Expense', desc:'Full-time and part-time employee wages and salaries' },
    { code:'7160', name:'Payroll Taxes - Employer Portion',  group:'Personnel',       section:'Expense', desc:'Employer share of payroll taxes' },
    { code:'7170', name:'Employee Benefits',                 group:'Personnel',       section:'Expense', desc:'Health, dental, vision, retirement, and other benefits' },
    { code:'7405', name:'Travel',                            group:'Operations',      section:'Expense', desc:'Flights, hotels, ground transport, and per diem' },
    { code:'7525', name:'Other Services',                    group:'Operations',      section:'Expense', desc:'Contracted services and professional fees' },
    { code:'7605', name:'Computer & Information Technology', group:'Technology',      section:'Expense', desc:'Hardware, software licenses, cloud services, IT support' },
    { code:'7645', name:'Dues & Subscriptions',              group:'Technology',      section:'Expense', desc:'Professional memberships and SaaS subscriptions' },
    { code:'7790', name:'Miscellaneous Expense',             group:'Other Expense',   section:'Expense', desc:'Minor or infrequent expenses not classified elsewhere' },
    { code:'8285', name:'Internal Insurance',                group:'Internal',        section:'Expense', desc:'Insurance premiums allocated internally between OUs' },
    { code:'8525', name:'Internal Other Services & Fees',    group:'Internal',        section:'Expense', desc:'Fees for services provided by other SIL operating units' },
  ];
  accounts.forEach(a => { a.type = a.section === 'Income' ? 'income' : 'expense'; });
  return accounts;
}


// ============================================================
//  WATCHLIST
// ============================================================

const UPROP_WATCH_RULES = 'WATCH_RULES';

function getWatchRules() {
  const raw = PropertiesService.getUserProperties().getProperty(UPROP_WATCH_RULES);
  try { return raw ? JSON.parse(raw) : []; } catch(e) { return []; }
}

function saveWatchRules(rules) {
  PropertiesService.getUserProperties().setProperty(UPROP_WATCH_RULES, JSON.stringify(rules || []));
  return { ok: true };
}

/** Returns individual GL transactions matching a single watch rule. */
function getWatchTransactions(rule, fiscalYear, period, periodType, dateFrom, dateTo) {
  if (!rule || !rule.term) return [];
  const data = getGLData();
  if (!data) return [];

  const result = [];
  const term = rule.term.toLowerCase();
  let currentSectionCode = null;

  for (let i = 0; i < data.length; i++) {
    const row  = data[i];
    const colA = String(row[TXCOL_HDR_ACCOUNT] || '').trim();
    const colB = String(row[TXCOL_CODE]         || '').trim();

    if (colA && !colA.startsWith('Total') && !colB) {
      currentSectionCode = normalizeCode(extractCode(colA));
      continue;
    }
    if (colA.startsWith('Total')) continue;

    if (!colA && colB && !isSkipRow(colB)) {
      const d = parseGLDate(row[TXCOL_POST_DATE]);
      if (!d || !inPeriod(d, fiscalYear, period, periodType, dateFrom, dateTo)) continue;

      const txSuiteKey = String(row[TXCOL_SUITEKEY] || '').trim();
      if (rule.suiteKey && rule.suiteKey !== txSuiteKey) continue;

      const txName = String(row[TXCOL_NAME] || '').toLowerCase();
      const txDesc = String(row[TXCOL_DESC] || '').toLowerCase();
      const matchName = rule.field !== 'desc' && txName.includes(term);
      const matchDesc = rule.field !== 'name' && txDesc.includes(term);
      if (!matchName && !matchDesc) continue;

      const sectionCode = normalizeCode(currentSectionCode || colB);
      const txnDate = parseGLDate(row[TXCOL_TXN_DATE]);
      result.push({
        postedDate : formatDate(d),
        txnDate    : formatDate(txnDate || d),
        docNum     : String(row[TXCOL_DOC_NUM]     || ''),
        desc       : String(row[TXCOL_DESC]        || ''),
        name       : String(row[TXCOL_NAME]        || ''),
        account    : sectionCode,
        suiteKey   : txSuiteKey,
        amount     : calcAmount(row, sectionCode),
      });
    }
  }
  return result;
}

/**
 * Scans the GL for transactions matching each watch rule and returns totals.
 * Each rule: { id, term, field ('name'|'desc'|'both'), suiteKey ('' = any), label }
 */
function getWatchListData(rules, fiscalYear, period, periodType, dateFrom, dateTo) {
  if (!rules || !rules.length) return [];
  const data = getGLData();
  if (!data) return rules.map(r => ({ id: r.id, total: 0, count: 0 }));

  const totals = {};
  const counts = {};
  rules.forEach(r => { totals[r.id] = 0; counts[r.id] = 0; });

  let currentSectionCode = null;

  for (let i = 0; i < data.length; i++) {
    const row  = data[i];
    const colA = String(row[TXCOL_HDR_ACCOUNT] || '').trim();
    const colB = String(row[TXCOL_CODE]         || '').trim();

    if (colA && !colA.startsWith('Total') && !colB) {
      currentSectionCode = normalizeCode(extractCode(colA));
      continue;
    }
    if (colA.startsWith('Total')) continue;

    if (!colA && colB && !isSkipRow(colB)) {
      const d = parseGLDate(row[TXCOL_POST_DATE]);
      if (!d || !inPeriod(d, fiscalYear, period, periodType, dateFrom, dateTo)) continue;

      const txSuiteKey = String(row[TXCOL_SUITEKEY] || '').trim();
      const txName     = String(row[TXCOL_NAME]     || '').toLowerCase();
      const txDesc     = String(row[TXCOL_DESC]     || '').toLowerCase();
      const sectionCode = normalizeCode(currentSectionCode || colB);
      const amount     = calcAmount(row, sectionCode);

      for (let j = 0; j < rules.length; j++) {
        const rule = rules[j];
        if (rule.suiteKey && rule.suiteKey !== txSuiteKey) continue;
        const term = (rule.term || '').toLowerCase();
        if (!term) continue;
        const matchName = rule.field !== 'desc' && txName.includes(term);
        const matchDesc = rule.field !== 'name' && txDesc.includes(term);
        if (matchName || matchDesc) {
          totals[rule.id] += amount;
          counts[rule.id]++;
        }
      }
    }
  }

  return rules.map(r => ({ id: r.id, total: totals[r.id] || 0, count: counts[r.id] || 0 }));
}


// ============================================================
//  HELPERS
// ============================================================

function getCurrentFiscalYear() {
  const now = new Date();
  return (now.getMonth() + 1) >= FISCAL_YEAR_START_MONTH
    ? now.getFullYear() + 1 : now.getFullYear();
}

function extractCode(raw) {
  const m = String(raw).match(/^(\d{4,5})/);
  return m ? m[1] : null;
}

function normalizeCode(code) {
  if (!code) return code;
  const m = String(code).match(/^(\d{4,5})/);
  return m ? m[1] : code;
}

function parseMoney(val) {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') return val;
  const s = String(val).trim();
  if (!s || s === '-' || s === '--' || s === '$-') return 0;
  const n = parseFloat(s.replace(/[$,\s]/g, '').replace(/\(([^)]+)\)/, '-$1'));
  return isNaN(n) ? 0 : n;
}

function getFiscalYear(date) {
  const m = date.getMonth() + 1;
  return m >= FISCAL_YEAR_START_MONTH ? date.getFullYear() + 1 : date.getFullYear();
}

function getFiscalQuarter(date) {
  return Math.floor(((date.getMonth() + 1 - FISCAL_YEAR_START_MONTH + 12) % 12) / 3) + 1;
}

function inPeriod(date, fiscalYear, period, periodType, dateFrom, dateTo) {
  // Custom date range — ignore fiscal year entirely
  if (periodType === 'custom') {
    if (!dateFrom || !dateTo) return true;
    const from = new Date(dateFrom + 'T00:00:00');
    const to   = new Date(dateTo   + 'T23:59:59');
    return date >= from && date <= to;
  }
  if (getFiscalYear(date) !== parseInt(fiscalYear, 10)) return false;
  if (periodType === 'fiscal_year') return true;
  if (periodType === 'quarter') return getFiscalQuarter(date) === parseInt(period, 10);
  if (periodType === 'month') {
    return (date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0')) === String(period);
  }
  return true;
}

function parseGLDate(raw) {
  if (!raw) return null;
  if (raw instanceof Date) return isNaN(raw) ? null : raw;
  const s = String(raw).trim();
  const m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})$/);
  if (m) {
    let year = parseInt(m[3], 10);
    if (year < 100) year += 2000;
    const d = new Date(year, parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    return isNaN(d) ? null : d;
  }
  const d = new Date(raw);
  return isNaN(d) ? null : d;
}

function formatDate(d) {
  if (!d) return '';
  const M = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return M[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}

function isSkipRow(s) {
  return !s || /^total/i.test(s) || /^<[-<]/.test(s);
}

function deriveGroup(code) {
  const n = parseInt(code, 10);
  if (isNaN(n))               return 'Other';
  if (n >= 4000 && n <= 4999) return 'Contributions';
  if (n >= 5000 && n <= 5999) return 'Sales & Service';
  if (n >= 6000 && n <= 6999) return 'Other Income';
  if (n >= 7000 && n <= 7199) return 'Personnel';
  if (n >= 7200 && n <= 7399) return 'Facilities';
  if (n >= 7400 && n <= 7599) return 'Operations';
  if (n >= 7600 && n <= 7699) return 'Technology';
  if (n >= 7700 && n <= 7999) return 'Other Expense';
  if (n >= 8000 && n <= 8999) return 'Internal';
  return 'Uncategorized';
}