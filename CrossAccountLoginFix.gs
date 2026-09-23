// ---------------------------------------------------------------------------
// CROSS-ACCOUNT LOGIN FIX - diagnostics and recovery tools
// ---------------------------------------------------------------------------
// Symptom this file addresses:
//   "I imported the same project into another Gmail account and bound it to the
//    same sheet. Login works on the original account but not on the new one."
//
// Cause:
//   The password salt was stored in Script Properties, which belong to the Apps
//   Script PROJECT - not to the spreadsheet. An imported/copied project starts
//   with empty script properties, silently generates a new random salt, and
//   every hash already stored in the Users sheet stops matching.
//   Security.gs now keeps the salt in the spreadsheet (hidden "_SystemConfig"
//   sheet) so it travels with the data; these helpers let you verify and repair
//   an installation that was copied before that change.
//
// HOW TO USE (run these from the Apps Script editor: pick the function, Run)
//   1. On the ORIGINAL project (the one where login still works):
//        run  publishPasswordSaltToSheet()
//      -> writes the salt into the spreadsheet.
//   2. On the NEW/imported project:
//        run  syncPasswordSaltFromSheet()
//      -> adopts the salt from the spreadsheet. Login now works.
//   3. If that is not possible, copy the salt across manually:
//        original project: run exportPasswordSalt()   (Admin must be signed in)
//        new project:      run setPasswordSalt('<paste the value here>')
//
//   At any time run diagnoseLoginEnvironment() to see exactly what that copy of
//   the project sees (spreadsheet binding, write access, salt source, hash
//   formats, and whether a given username/password verifies).
// ---------------------------------------------------------------------------

/**
 * True when the caller is an Admin with a valid session, or the account the
 * script runs as can edit the spreadsheet (i.e. the owner running these tools by
 * hand from the Apps Script editor).
 * @returns {boolean}
 */
function isRecoveryAuthorised_() {
  try {
    if (typeof checkSession === 'function' && checkSession()) {
      const user = typeof getLoggedInUser === 'function' ? getLoggedInUser() : null;
      if (user && String(user.role || '').trim().toLowerCase() === 'admin') return true;
    }
  } catch (e) {
    /* no session - fall through */
  }
  return isEffectiveUserSpreadsheetEditor_();
}

/**
 * Whether the account the script executes as can edit the bound spreadsheet.
 * @returns {boolean}
 */
function isEffectiveUserSpreadsheetEditor_() {
  try {
    const ss = typeof getSpreadsheet_ === 'function' ? getSpreadsheet_() : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return false;
    const email = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
    if (!email) return false;
    try {
      if (String(ss.getOwner().getEmail() || '').toLowerCase() === email) return true;
    } catch (e) { /* owner not visible - ignore */ }
    const editors = ss.getEditors().map(function(u) { return String(u.getEmail() || '').toLowerCase(); });
    return editors.indexOf(email) !== -1;
  } catch (e) {
    return false;
  }
}

/**
 * Users sheet layout, read defensively so this file still works even if the
 * constants in User.gs are not in scope.
 * @returns {{firstRow: number, firstCol: number, colCount: number}}
 */
function usersSheetLayout_() {
  return {
    firstRow: (typeof USER_SHEET_FIRST_DATA_ROW !== 'undefined') ? USER_SHEET_FIRST_DATA_ROW : 3,
    firstCol: (typeof USER_SHEET_FIRST_COL !== 'undefined') ? USER_SHEET_FIRST_COL : 2,
    colCount: (typeof USER_SHEET_COL_COUNT !== 'undefined') ? USER_SHEET_COL_COUNT : 9
  };
}

/** Short, non-secret identifier for a salt so two copies can be compared. */
function saltFingerprint_(salt) {
  const value = String(salt || '');
  if (!value) return 'none';
  return value.length + ':' + value.slice(0, 6) + '…';
}

/**
 * Full report of what THIS copy of the project sees.
 * Optionally pass a username/password to test that credential.
 * Run it from the Apps Script editor and read the Logger output (View > Logs),
 * or run it and inspect the returned object in the Apps Script dashboard.
 *
 * @param {string} [username] optional
 * @param {string} [password] optional
 * @returns {object} diagnostic report
 */
function diagnoseLoginEnvironment(username, password) {
  if (!isRecoveryAuthorised_()) {
    return { authorised: false, message: 'Run this from the Apps Script editor with an account that can edit the spreadsheet.' };
  }

  const report = {
    authorised: true,
    ranAs: '',
    spreadsheet: {},
    salt: {},
    users: {},
    session: {},
    credentialCheck: null,
    hints: []
  };

  try { report.ranAs = Session.getEffectiveUser().getEmail(); } catch (e) { report.ranAs = 'unknown'; }

  // --- spreadsheet binding -------------------------------------------------
  let ss = null;
  let bound = false;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    bound = !!ss;
  } catch (e) { /* ignore */ }
  if (!ss && typeof getSpreadsheet_ === 'function') ss = getSpreadsheet_();

  report.spreadsheet = {
    boundToSpreadsheet: bound,
    connected: !!ss,
    name: ss ? ss.getName() : null,
    id: ss ? ss.getId() : null,
    url: ss ? ss.getUrl() : null,
    canEdit: isEffectiveUserSpreadsheetEditor_()
  };

  if (!bound) {
    report.hints.push('This project is NOT container-bound to a spreadsheet. Open the spreadsheet and use Extensions > Apps Script, or run setSpreadsheetId("<spreadsheet id>") once.');
  }
  if (ss && !report.spreadsheet.canEdit) {
    report.hints.push('The account running this script (' + report.ranAs + ') is not an editor of the spreadsheet. Share the sheet as Editor, or deploy the web app with "Execute as: me".');
  }

  // --- salt ----------------------------------------------------------------
  const sheetSalt = String(readSaltFromSheet_() || '').trim();
  let scriptSalt = '';
  let previousSalt = '';
  try {
    const props = PropertiesService.getScriptProperties();
    scriptSalt = String(props.getProperty(SALT_PROPERTY_KEY) || '').trim();
    previousSalt = String(props.getProperty(SALT_PREVIOUS_KEY) || '').trim();
  } catch (e) { /* ignore */ }
  const activeSalt = String(getSalt() || '').trim();

  const provenance = readSaltProvenance_();
  report.salt = {
    activeSource: sheetSalt ? 'spreadsheet (_SystemConfig)' : (scriptSalt ? 'script properties (legacy - not portable)' : 'newly generated'),
    activeFingerprint: saltFingerprint_(activeSalt),
    inSpreadsheet: !!sheetSalt,
    inScriptProperties: !!scriptSalt,
    spreadsheetFingerprint: saltFingerprint_(sheetSalt),
    scriptFingerprint: saltFingerprint_(scriptSalt),
    previousFingerprint: saltFingerprint_(previousSalt),
    saltsAgree: !sheetSalt || !scriptSalt || sheetSalt === scriptSalt,
    setBy: provenance.setBy,
    setAt: provenance.setAt,
    knownCandidates: listSaltCandidates_().map(function(c) { return c.name + ' = ' + c.fingerprint; })
  };

  if (provenance.setBy && report.ranAs && provenance.setBy.toLowerCase() !== String(report.ranAs).toLowerCase()) {
    report.hints.push('The salt in the spreadsheet was published by ' + provenance.setBy + ', but this copy runs as ' + report.ranAs + '. If that account is the imported copy, it overwrote the salt with its own - take the salt from the original project (exportPasswordSalt) and run setPasswordSalt("<value>") here.');
  }

  if (sheetSalt && scriptSalt && sheetSalt !== scriptSalt) {
    report.hints.push('This project had generated its own salt before it could read the spreadsheet one. Run syncPasswordSaltFromSheet() (the current salt has already been auto-aligned to the spreadsheet).');
  }
  if (!sheetSalt) {
    report.hints.push('No salt is stored in the spreadsheet yet. Run publishPasswordSaltToSheet() on the ORIGINAL project, then syncPasswordSaltFromSheet() here - otherwise every copy keeps its own salt and passwords keep breaking.');
  }

  // --- users ---------------------------------------------------------------
  const users = { total: 0, hashed: 0, sha256Prefixed: 0, plainText: 0, empty: 0, active: 0, locked: 0, sample: [] };
  if (ss) {
    const sheet = ss.getSheetByName('Users');
    if (!sheet) {
      report.hints.push('There is no "Users" sheet in this spreadsheet - this copy is pointing at the wrong file.');
    } else if (sheet.getLastRow() >= usersSheetLayout_().firstRow) {
      const layout = usersSheetLayout_();
      const data = sheet.getRange(
        layout.firstRow,
        layout.firstCol,
        sheet.getLastRow() - (layout.firstRow - 1),
        layout.colCount
      ).getValues();

      data.forEach(function(row) {
        const stored = String(row[6] || '').trim();
        if (!String(row[5] || '').trim() && !stored) return; // blank row
        users.total++;
        const format = getPasswordHashFormat_(stored);
        if (format === 'hashed') users.hashed++;
        else if (format === 'sha256-prefixed') users.sha256Prefixed++;
        else if (format === 'plain-text') users.plainText++;
        else users.empty++;

        if (String(row[4] || '').trim().toLowerCase() === 'active') users.active++;
        if (isUserAccountLocked_(parseInt(row[7], 10) || 0, row[8])) users.locked++;

        if (users.sample.length < 5) {
          users.sample.push({
            username: String(row[5] || '').trim(),
            status: String(row[4] || '').trim(),
            storedFormat: format,
            failedAttempts: parseInt(row[7], 10) || 0
          });
        }
      });
    }
  }
  report.users = users;

  if (users.plainText > 0 || users.sha256Prefixed > 0) {
    report.hints.push('Some users still use an old password format. They are upgraded automatically the first time they log in successfully.');
  }
  if (users.locked > 0) {
    report.hints.push(users.locked + ' account(s) are currently locked after failed attempts. Clear columns I and J on the Users sheet for those rows, or wait 30 minutes.');
  }

  // --- sessions ------------------------------------------------------------
  try {
    const props = PropertiesService.getUserProperties();
    const sessionId = props.getProperty('CURRENT_SESSION_ID');
    report.session = {
      hasSessionId: !!sessionId,
      sessionValid: sessionId ? !!validateSession(sessionId) : false,
      note: 'Web apps deployed with "Execute as: me" share one session store between all visitors, because the script always runs as the deployer.'
    };
  } catch (e) {
    report.session = { error: e.message };
  }

  // --- optional credential test -------------------------------------------
  if (username) {
    report.credentialCheck = testUserCredential(username, password || '');
  }

  report.ok = report.hints.length === 0;
  Logger.log('LOGIN DIAGNOSTICS\n' + JSON.stringify(report, null, 2));
  return report;
}

/**
 * Check one credential and report which storage format matched.
 * Returns no secret material.
 * @param {string} username
 * @param {string} password
 * @returns {object}
 */
function testUserCredential(username, password) {
  if (!isRecoveryAuthorised_()) {
    return { authorised: false, message: 'Not authorised.' };
  }

  const ss = typeof getSpreadsheet_ === 'function' ? getSpreadsheet_() : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { found: false, reason: 'No spreadsheet connected to this project.' };

  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { found: false, reason: 'No "Users" sheet.' };

  const normalized = String(username || '').trim();
  const layout = usersSheetLayout_();
  if (sheet.getLastRow() >= layout.firstRow) {
    const data = sheet.getRange(
      layout.firstRow,
      layout.firstCol,
      sheet.getLastRow() - (layout.firstRow - 1),
      layout.colCount
    ).getValues();

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][5]).trim() !== normalized) continue;

      const stored = String(data[i][6] || '').trim();
      const status = String(data[i][4] || '').trim();
      const locked = isUserAccountLocked_(parseInt(data[i][7], 10) || 0, data[i][8]);
      const check = password ? verifyPasswordWithDetails_(password, stored) : { matched: false, mode: 'not-tested', salt: '' };

      return {
        authorised: true,
        found: true,
        row: i + layout.firstRow,
        username: normalized,
        accountStatus: status,
        locked: locked,
        storedFormat: getPasswordHashFormat_(stored),
        passwordMatches: check.matched,
        matchMode: check.mode,
        usedSaltFingerprint: saltFingerprint_(check.salt),
        currentSaltFingerprint: saltFingerprint_(getSalt())
      };
    }
  }

  return { authorised: true, found: false, reason: 'No user row with that username on this copy of the project.' };
}

/**
 * Every salt value this copy of the project knows about.
 * @returns {Array<{name: string, value: string, fingerprint: string}>}
 */
function listSaltCandidates_() {
  const out = [];
  const push = function(name, value) {
    const clean = String(value || '').trim();
    if (!clean) return;
    for (let i = 0; i < out.length; i++) {
      if (out[i].value === clean) return; // de-duplicate
    }
    out.push({ name: name, value: clean, fingerprint: saltFingerprint_(clean) });
  };

  push('spreadsheet (_SystemConfig)', readSaltFromSheet_());
  try {
    const props = PropertiesService.getScriptProperties();
    push('script properties (current)', props.getProperty(SALT_PROPERTY_KEY));
    push('script properties (previous)', props.getProperty(SALT_PREVIOUS_KEY));
  } catch (e) { /* ignore */ }
  push('active', getSalt());
  return out;
}

/**
 * Admin only: reveal every salt this project knows, so one of them can be
 * pasted into another copy with setPasswordSalt(). Run it on the ORIGINAL
 * project (where login still works).
 * @returns {object}
 */
function exportPasswordSalt() {
  requireLogin();
  const user = getLoggedInUser();
  if (!user || String(user.role || '').trim().toLowerCase() !== 'admin') {
    return { success: false, message: 'Admin only.' };
  }

  const candidates = listSaltCandidates_();
  return {
    success: true,
    salt: getSalt(),
    fingerprint: saltFingerprint_(getSalt()),
    candidates: candidates,
    note: 'On the other copy of the project run setPasswordSalt("<value>") with each value until testUserCredential() reports passwordMatches: true.'
  };
}

/**
 * Find which known salt a given username/password was actually hashed with.
 * Does not change anything - use it to confirm a credential before applying.
 * @param {string} username
 * @param {string} password
 * @returns {object}
 */
function findWorkingSaltFor(username, password) {
  if (!isRecoveryAuthorised_()) return { authorised: false, message: 'Not authorised.' };
  if (!username || !password) return { found: false, message: 'Pass a username and a password.' };

  const ss = typeof getSpreadsheet_ === 'function' ? getSpreadsheet_() : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss ? ss.getSheetByName('Users') : null;
  if (!sheet) return { found: false, message: 'No "Users" sheet.' };

  const layout = usersSheetLayout_();
  const normalized = String(username).trim();
  let stored = '';
  let foundRow = -1;
  if (sheet.getLastRow() >= layout.firstRow) {
    const data = sheet.getRange(layout.firstRow, layout.firstCol, sheet.getLastRow() - (layout.firstRow - 1), layout.colCount).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][5]).trim() === normalized) {
        stored = String(data[i][6] || '').trim();
        foundRow = i + layout.firstRow;
        break;
      }
    }
  }
  if (foundRow === -1) return { found: false, message: 'No user row with username "' + normalized + '".' };
  if (!stored) return { found: false, message: 'That user has no stored password.' };

  const candidates = listSaltCandidates_();
  for (let i = 0; i < candidates.length; i++) {
    if (hashWithSalt_(password, candidates[i].value) === stored) {
      return {
        found: true,
        username: normalized,
        salt: candidates[i].value,
        fingerprint: candidates[i].fingerprint,
        source: candidates[i].name,
        alreadyActive: candidates[i].value === getSalt(),
        message: candidates[i].value === getSalt()
          ? 'This credential already verifies with the active salt - login should work. If it does not, the problem is sessions/permissions, not the salt.'
          : 'Found it. Run setPasswordSalt("' + candidates[i].value + '") to apply it.'
      };
    }
  }

  if (verifyPasswordWithDetails_(password, stored).matched) {
    return { found: true, username: normalized, salt: '', source: 'legacy format (no salt)', message: 'This password is stored in a legacy format and verifies as-is; it is upgraded automatically on the next successful login.' };
  }

  return {
    found: false,
    username: normalized,
    message: 'None of the salts known to THIS copy of the project match. Get the salt from the original project with exportPasswordSalt() and apply it here with setPasswordSalt("<value>").',
    tried: candidates.map(function(c) { return c.fingerprint; })
  };
}

/**
 * One-shot repair: work out which salt this credential was hashed with and
 * apply it to this copy of the project.
 * @param {string} username
 * @param {string} password
 * @returns {object}
 */
function applySaltForUser(username, password) {
  const hit = findWorkingSaltFor(username, password);
  if (!hit.found) return hit;

  const check = testUserCredential(username, password);
  if (check.passwordMatches && hit.alreadyActive !== false) {
    return { success: true, changed: false, message: 'No change needed - the active salt already verifies this credential.', credential: check };
  }

  const applied = setPasswordSalt(hit.salt);
  return {
    success: applied.success,
    changed: true,
    message: applied.message,
    fingerprint: hit.fingerprint,
    credential: testUserCredential(username, password)
  };
}

/**
 * Test one specific salt value against a user's stored password, and apply it
 * when it matches. Use with the values returned by exportPasswordSalt().
 * @param {string} salt
 * @param {string} username
 * @param {string} password
 * @param {boolean} [apply] set true to store it when it matches
 * @returns {object}
 */
function tryPasswordSalt(salt, username, password, apply) {
  if (!isRecoveryAuthorised_()) return { authorised: false, message: 'Not authorised.' };

  const value = String(salt || '').trim();
  if (!value) return { matched: false, message: 'No salt supplied.' };

  const check = testUserCredential(username, password);
  if (!check.found) return check;

  const ss = typeof getSpreadsheet_ === 'function' ? getSpreadsheet_() : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const stored = String(sheet.getRange(check.row, 8).getValue() || '').trim();
  const matched = hashWithSalt_(password, value) === stored;

  const result = {
    matched: matched,
    fingerprint: saltFingerprint_(value),
    storedFormat: getPasswordHashFormat_(stored),
    applied: false
  };

  if (matched && apply) {
    const applied = setPasswordSalt(value);
    result.applied = applied.success;
    result.message = applied.message;
  } else if (!matched) {
    result.message = 'That salt does not produce the stored hash for this user.';
  }

  return result;
}

/**
 * Store the salt in every available store (spreadsheet + script properties).
 * Use this to repair a copy of the project that generated its own salt.
 * @param {string} salt
 * @returns {object}
 */
function setPasswordSalt(salt) {
  const value = String(salt || '').trim();
  if (!value) return { success: false, message: 'No salt supplied.' };
  if (!isRecoveryAuthorised_()) {
    return { success: false, message: 'Run this from the Apps Script editor with an account that can edit the spreadsheet (or signed in as Admin).' };
  }

  const current = String(getSalt() || '').trim();
  invalidateSaltCache_();
  if (current && current !== value) {
    try {
      PropertiesService.getScriptProperties().setProperty(SALT_PREVIOUS_KEY, current);
    } catch (e) { /* ignore */ }
  }

  const sheetWritten = writeSaltToSheet_(value);
  let propsWritten = false;
  try {
    PropertiesService.getScriptProperties().setProperty(SALT_PROPERTY_KEY, value);
    propsWritten = true;
  } catch (e) { /* ignore */ }

  try {
    if (typeof invalidateCacheOnModify === 'function') invalidateCacheOnModify('Users');
  } catch (e) { /* ignore */ }

  SALT_CACHE = value;

  return {
    success: sheetWritten || propsWritten,
    storedInSpreadsheet: sheetWritten,
    storedInScriptProperties: propsWritten,
    fingerprint: saltFingerprint_(value),
    message: sheetWritten
      ? 'Salt saved. Ask users to sign in again - passwords from the original project will now verify.'
      : 'Salt saved to script properties only; the spreadsheet could not be written, so other copies of the project will not inherit it.'
  };
}

/**
 * Write this project's salt into the spreadsheet so imported copies inherit it.
 * Run once on the ORIGINAL project.
 * @returns {object}
 */
function publishPasswordSaltToSheet() {
  if (!isRecoveryAuthorised_()) {
    return { success: false, message: 'Run this from the Apps Script editor with an account that can edit the spreadsheet.' };
  }
  const salt = getSalt();
  const written = writeSaltToSheet_(salt);
  return {
    success: written,
    fingerprint: saltFingerprint_(salt),
    message: written
      ? 'Salt published to the spreadsheet. Any other copy of the project bound to this file can now run syncPasswordSaltFromSheet().'
      : 'Could not write to the spreadsheet. Check that this account can edit the file.'
  };
}

/**
 * Adopt the salt stored in the spreadsheet (repairs an imported copy).
 * @returns {object}
 */
function syncPasswordSaltFromSheet() {
  if (!isRecoveryAuthorised_()) {
    return { success: false, message: 'Run this from the Apps Script editor with an account that can edit the spreadsheet.' };
  }

  const sheetSalt = String(readSaltFromSheet_() || '').trim();
  if (!sheetSalt) {
    return {
      success: false,
      message: 'The spreadsheet does not contain a salt yet. Run publishPasswordSaltToSheet() on the ORIGINAL project first (or use setPasswordSalt("<salt>")).'
    };
  }

  const result = setPasswordSalt(sheetSalt);
  result.message = result.success
    ? 'Salt synced from the spreadsheet (' + result.fingerprint + '). Existing passwords from the original project will now verify.'
    : result.message;
  return result;
}

/**
 * Show the diagnostics report in a dialog (Spreadsheet menu: "Login Fix").
 * Useful when you cannot log in at all: it tells you why.
 */
function showLoginDiagnostics() {
  const report = diagnoseLoginEnvironment();
  const json = JSON.stringify(report, null, 2)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  const html =
    '<div style="font-family:monospace;font-size:12px;white-space:pre-wrap;">' + json + '</div>' +
    '<p style="font-family:sans-serif;font-size:12px;">Run <b>publishPasswordSaltToSheet</b> on the ORIGINAL project, ' +
    'then <b>syncPasswordSaltFromSheet</b> here, if the salt sources differ.</p>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(760).setHeight(520),
    'Login diagnostics'
  );
}

/**
 * Point a standalone copy of the project at the right spreadsheet.
 * Only needed when the project is NOT opened from the spreadsheet
 * (Extensions > Apps Script). The id is the long string in the sheet URL:
 * https://docs.google.com/spreadsheets/d/<THIS PART>/edit
 * @param {string} id
 * @returns {object}
 */
function setSpreadsheetId(id) {
  const value = String(id || '').trim();
  if (!value) return { success: false, message: 'No spreadsheet id supplied.' };

  try {
    const ss = SpreadsheetApp.openById(value);
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', value);
    return { success: true, spreadsheetName: ss.getName(), url: ss.getUrl() };
  } catch (e) {
    return { success: false, message: 'Could not open that spreadsheet: ' + e.message };
  }
}
