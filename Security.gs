// --- PASSWORD SECURITY FUNCTIONS ---

function buildSecurityAuditSnapshot(fields) {
  if (!fields) return null;

  return buildAuditSnapshot({
    totalUsers: fields.totalUsers,
    hashedPasswords: fields.hashedPasswords,
    plainTextPasswords: fields.plainTextPasswords,
    migrated: fields.migrated,
    secure: fields.secure,
    message: fields.message
  });
}

/**
 * ---------------------------------------------------------------------------
 * PASSWORD SALT - NOW PORTABLE ACROSS GOOGLE ACCOUNTS
 * ---------------------------------------------------------------------------
 * The salt used to be stored ONLY in Script Properties. Script Properties belong
 * to the Apps Script *project*, not to the spreadsheet. So when this project is
 * imported/copied into another Google account, that copy starts with EMPTY
 * script properties, silently generates a brand-new random salt, and every
 * password hash already stored in the Users sheet (column H) stops verifying.
 * Result: login works on the original account, but every user gets
 * "Invalid username or password" on the copied project.
 *
 * Fix: the salt is now kept IN THE SPREADSHEET (hidden "_SystemConfig" sheet),
 * so it travels with the data and is shared by every copy of the project that is
 * bound to the same spreadsheet. Script Properties are still read and written so
 * that existing installations keep working and the value stays recoverable.
 * ---------------------------------------------------------------------------
 */

var SALT_SHEET_NAME = '_SystemConfig';
var SALT_PROPERTY_KEY = 'PASSWORD_SALT';
var SALT_PREVIOUS_KEY = 'PASSWORD_SALT_PREVIOUS';
var SALT_SHEET_KEY = 'Password Salt';

// Resolved once per execution so hashing a whole sheet of passwords does not
// re-read the spreadsheet on every row. Cleared whenever the salt is changed.
var SALT_CACHE = '';

/** Forget the cached salt (call after changing it). */
function invalidateSaltCache_() {
  SALT_CACHE = '';
}

/**
 * Resolve the spreadsheet this app runs against.
 * Container-bound scripts return the bound file. A standalone copy can be
 * pointed at the file with setSpreadsheetId('<id>') (see CrossAccountLoginFix.gs).
 * @returns {Spreadsheet|null}
 */
function getSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {
    // Not container-bound, or no active spreadsheet in this context.
  }

  try {
    const storedId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (storedId) return SpreadsheetApp.openById(storedId);
  } catch (e) {
    Logger.log('Could not open spreadsheet by stored id: ' + e.message);
  }

  return null;
}

/**
 * Get (or lazily create) the hidden key/value sheet used for portable settings.
 * @param {Spreadsheet} ss
 * @param {boolean} create - create the sheet when missing
 * @returns {Sheet|null}
 */
function getSystemConfigSheet_(ss, create) {
  if (!ss) return null;

  let sheet = ss.getSheetByName(SALT_SHEET_NAME);
  if (!sheet && create) {
    try {
      sheet = ss.insertSheet(SALT_SHEET_NAME);
      sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      sheet.hideSheet();
      sheet.setFrozenRows(1);
    } catch (e) {
      Logger.log('Could not create ' + SALT_SHEET_NAME + ' sheet: ' + e.message);
      return null;
    }
  }
  return sheet;
}

/**
 * Read the salt stored inside the spreadsheet (portable across project copies).
 * Checks both the visible Settings sheet (always included in exports) and the
 * hidden _SystemConfig sheet.
 * @returns {string} salt, or '' when not present / not readable
 */
function readSaltFromSheet_() {
  try {
    const ss = getSpreadsheet_();
    if (!ss) return '';

    // 1. Check Settings sheet first (visible, travels with data exports)
    const settingsSheet = ss.getSheetByName('Settings');
    if (settingsSheet && settingsSheet.getLastRow() >= 5) {
      const data = settingsSheet.getRange(5, 2, settingsSheet.getLastRow() - 4, 2).getValues();
      for (let i = 0; i < data.length; i++) {
        const key = String(data[i][0] || '').trim().toLowerCase().replace(/[\s_\-]+/g, '');
        if (key === 'passwordsalt' || key === 'systemsalt') {
          const val = String(data[i][1] || '').trim();
          if (val) return val;
        }
      }
    }

    // 2. Check hidden _SystemConfig sheet
    const sheet = getSystemConfigSheet_(ss, false);
    if (sheet && sheet.getLastRow() >= 2) {
      const values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).trim().toLowerCase() === SALT_SHEET_KEY.toLowerCase()) {
          const val = String(values[i][1] || '').trim();
          if (val) return val;
        }
      }
    }
  } catch (e) {
    Logger.log('Could not read salt from spreadsheet: ' + e.message);
  }
  return '';
}

/**
 * Persist the salt inside the spreadsheet so every copy of the project that is
 * bound to this file uses the same salt. Writes to both _SystemConfig and Settings.
 * @param {string} salt
 * @returns {boolean} true when written
 */
function writeSaltToSheet_(salt) {
  if (!salt) return false;
  try {
    const ss = getSpreadsheet_();
    if (!ss) return false;

    let written = false;

    // 1. Write to hidden _SystemConfig sheet
    const sheet = getSystemConfigSheet_(ss, true);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      let updated = false;
      if (lastRow >= 2) {
        const values = sheet.getRange(1, 1, lastRow, 2).getValues();
        for (let i = 1; i < values.length; i++) {
          if (String(values[i][0]).trim().toLowerCase() === SALT_SHEET_KEY.toLowerCase()) {
            sheet.getRange(i + 1, 2).setValue(salt);
            writeSaltProvenance_(sheet);
            updated = true;
            written = true;
            break;
          }
        }
      }
      if (!updated) {
        sheet.getRange(Math.max(sheet.getLastRow(), 1) + 1, 1, 1, 2).setValues([[SALT_SHEET_KEY, salt]]);
        writeSaltProvenance_(sheet);
        written = true;
      }
    }

    // 2. Also write to Settings sheet (visible, so exports and copies include it)
    try {
      const settingsSheet = ss.getSheetByName('Settings');
      if (settingsSheet && settingsSheet.getLastRow() >= 5) {
        const lastRow = settingsSheet.getLastRow();
        const data = settingsSheet.getRange(5, 2, lastRow - 4, 1).getValues();
        let foundRow = -1;
        for (let i = 0; i < data.length; i++) {
          const key = String(data[i][0] || '').trim().toLowerCase().replace(/[\s_\-]+/g, '');
          if (key === 'passwordsalt' || key === 'systemsalt') {
            foundRow = 5 + i;
            break;
          }
        }
        if (foundRow !== -1) {
          settingsSheet.getRange(foundRow, 3).setValue(salt);
          written = true;
        } else {
          settingsSheet.getRange(lastRow + 1, 2, 1, 2).setValues([['Password Salt', salt]]);
          written = true;
        }
      }
    } catch (e) {
      Logger.log('Could not write salt to Settings sheet: ' + e.message);
    }

    return written;
  } catch (e) {
    Logger.log('Could not write salt to spreadsheet: ' + e.message);
    return false;
  }
}

/**
 * Record who published the salt currently stored in the spreadsheet.
 * Makes it obvious when a copied project has overwritten it with its own.
 */
function writeSaltProvenance_(sheet) {
  try {
    let email = '';
    try { email = Session.getEffectiveUser().getEmail() || ''; } catch (e) { /* ignore */ }
    const stamp = new Date().toISOString();
    writeConfigValue_(sheet, 'Password Salt Set By', email);
    writeConfigValue_(sheet, 'Password Salt Set At', stamp);
  } catch (e) {
    Logger.log('Could not write salt provenance: ' + e.message);
  }
}

/** Upsert a key/value pair on the hidden config sheet. */
function writeConfigValue_(sheet, key, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(1, 1, lastRow, 2).getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim().toLowerCase() === String(key).toLowerCase()) {
        sheet.getRange(i + 1, 2).setValue(value);
        return;
      }
    }
  }
  sheet.getRange(Math.max(sheet.getLastRow(), 1) + 1, 1, 1, 2).setValues([[key, value]]);
}

/** Read provenance of the salt stored in the spreadsheet (diagnostics). */
function readSaltProvenance_() {
  const info = { setBy: '', setAt: '' };
  try {
    const ss = getSpreadsheet_();
    const sheet = ss ? ss.getSheetByName(SALT_SHEET_NAME) : null;
    if (!sheet || sheet.getLastRow() < 2) return info;
    const values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0]).trim().toLowerCase();
      if (key === 'password salt set by') info.setBy = String(values[i][1] || '').trim();
      if (key === 'password salt set at') info.setAt = String(values[i][1] || '').trim();
    }
  } catch (e) {
    Logger.log('Could not read salt provenance: ' + e.message);
  }
  return info;
}

/**
 * Get or initialize the salt.
 * Resolution order: spreadsheet (portable) -> Script Properties (legacy) -> new.
 * Whatever is found is mirrored to every store that is writable, so the original
 * project and any imported copy end up sharing one salt.
 * @returns {string} salt
 */
function getSalt() {
  if (SALT_CACHE) return SALT_CACHE;

  let scriptSalt = '';
  try {
    scriptSalt = String(PropertiesService.getScriptProperties().getProperty(SALT_PROPERTY_KEY) || '').trim();
  } catch (e) {
    Logger.log('Could not read salt from script properties: ' + e.message);
  }

  const sheetSalt = String(readSaltFromSheet_() || '').trim();

  // The spreadsheet is the source of truth: it travels with the data.
  if (sheetSalt) {
    if (scriptSalt && scriptSalt !== sheetSalt) {
      // A copied project generated its own salt before it could read the sheet.
      // Remember the wrong one so old hashes can still be verified, then align.
      try {
        const props = PropertiesService.getScriptProperties();
        if (props.getProperty(SALT_PREVIOUS_KEY) !== scriptSalt) {
          props.setProperty(SALT_PREVIOUS_KEY, scriptSalt);
        }
        props.setProperty(SALT_PROPERTY_KEY, sheetSalt);
      } catch (e) {
        Logger.log('Could not align script salt with spreadsheet salt: ' + e.message);
      }
    } else if (!scriptSalt) {
      try {
        PropertiesService.getScriptProperties().setProperty(SALT_PROPERTY_KEY, sheetSalt);
      } catch (e) {
        Logger.log('Could not mirror salt into script properties: ' + e.message);
      }
    }
    SALT_CACHE = sheetSalt;
    return sheetSalt;
  }

  // Legacy install: salt only ever existed in Script Properties -> publish it to
  // the spreadsheet so future copies inherit it automatically.
  if (scriptSalt) {
    writeSaltToSheet_(scriptSalt);
    SALT_CACHE = scriptSalt;
    return scriptSalt;
  }

  // Fresh install.
  const newSalt = Utilities.getUuid();
  try {
    PropertiesService.getScriptProperties().setProperty(SALT_PROPERTY_KEY, newSalt);
  } catch (e) {
    Logger.log('Could not store new salt in script properties: ' + e.message);
  }
  writeSaltToSheet_(newSalt);
  SALT_CACHE = newSalt;
  return newSalt;
}

/** Unsalted SHA-256 hex (used by the legacy "sha256:" format). */
function sha256Hex_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(byte) { return ('0' + (byte & 0xFF).toString(16)).slice(-2); }).join('');
}

/** Salted SHA-256 hex - the current storage format. */
function hashWithSalt_(password, salt) {
  return sha256Hex_(String(password) + String(salt || ''));
}

/**
 * Hash a password using SHA-256 + the installation salt
 * @param {string} password - The plain text password to hash
 * @returns {string} The hashed password as a hexadecimal string
 */
function hashPassword(password) {
  if (!password) return '';
  return hashWithSalt_(password, getSalt());
}

/**
 * Every salt this installation has ever used, newest first.
 * Lets passwords written by an older/other copy of the project keep verifying.
 * @returns {string[]}
 */
function getKnownSalts_() {
  const salts = [];
  const add = function(value) {
    const clean = String(value || '').trim();
    if (clean && salts.indexOf(clean) === -1) salts.push(clean);
  };

  add(getSalt());
  add(readSaltFromSheet_());
  try {
    const props = PropertiesService.getScriptProperties();
    add(props.getProperty(SALT_PROPERTY_KEY));
    add(props.getProperty(SALT_PREVIOUS_KEY));
  } catch (e) {
    Logger.log('Could not read extra salts: ' + e.message);
  }
  return salts;
}

/**
 * Describe which storage format a stored password uses (diagnostics only).
 * @param {string} storedValue
 * @returns {string}
 */
function getPasswordHashFormat_(storedValue) {
  const value = String(storedValue || '').trim();
  if (!value) return 'empty';
  if (value.indexOf('sha256:') === 0) return 'sha256-prefixed';
  if (/^[a-f0-9]{64}$/i.test(value)) return 'hashed';
  return 'plain-text';
}

/**
 * Verify a password against a stored value.
 * Supports every format this project has ever written, so an imported copy of
 * the project can still authenticate users created on the original account:
 *   - salted SHA-256 (current, 64 hex chars)
 *   - "sha256:<hex>" (legacy unsalted)
 *   - plain text (very old records / Settings master password)
 * @param {string} inputPassword - The password to verify
 * @param {string} storedHash - The stored password value
 * @returns {boolean} True if password matches
 */
function verifyPassword(inputPassword, storedHash) {
  return verifyPasswordWithDetails_(inputPassword, storedHash).matched;
}

/**
 * Same as verifyPassword(), but also reports which format/salt matched.
 * Used by the diagnostics tool in CrossAccountLoginFix.gs.
 * @returns {{matched: boolean, mode: string, salt: string}}
 */
function verifyPasswordWithDetails_(inputPassword, storedHash) {
  const stored = String(storedHash || '').trim();
  if (!inputPassword || !stored) return { matched: false, mode: 'none', salt: '' };

  // Legacy: "sha256:<hex>" (unsalted)
  if (stored.indexOf('sha256:') === 0) {
    return {
      matched: sha256Hex_(inputPassword) === stored.slice(7),
      mode: 'sha256-prefixed',
      salt: ''
    };
  }

  // Current + historical: salted SHA-256 (64 hex chars)
  if (/^[a-f0-9]{64}$/i.test(stored)) {
    const salts = getKnownSalts_();
    for (let i = 0; i < salts.length; i++) {
      if (hashWithSalt_(inputPassword, salts[i]) === stored) {
        return { matched: true, mode: i === 0 ? 'salted-current' : 'salted-legacy', salt: salts[i] };
      }
    }
    // Very old records were hashed without any salt at all.
    if (sha256Hex_(inputPassword) === stored) {
      return { matched: true, mode: 'unsalted', salt: '' };
    }

    // Master password fallback from Settings sheet (allows emergency admin access)
    try {
      const masterPass = typeof getSystemParameter === 'function' ? (getSystemParameter('Master Password') || getSystemParameter('MasterPass')) : null;
      if (masterPass && typeof isMasterPasswordMatch === 'function' && isMasterPasswordMatch(inputPassword, masterPass)) {
        return { matched: true, mode: 'master-password', salt: '' };
      }
    } catch (e) {
      // ignore
    }

    return { matched: false, mode: 'hashed-no-match', salt: '' };
  }

  // Legacy / Direct Plain text: exact or trimmed match
  const inputStr = String(inputPassword || '');
  if (inputStr === stored || inputStr.trim() === stored) {
    return { matched: true, mode: 'plain-text', salt: '' };
  }

  return { matched: false, mode: 'plain-no-match', salt: '' };
}

/**
 * Migrate all existing plain text passwords to hashed passwords
 * This should be run ONCE after implementing password hashing
 * @returns {object} Migration results
 */
function migratePasswordsToHash() {
  const beforeStatus = checkPasswordSecurityStatus();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  
  if (!sheet) {
    return { success: false, message: "Users sheet not found" };
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    return { success: true, message: "No users to migrate", migrated: 0 };
  }
  
  // Read all user data (Column B to H)
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 7);
  const data = dataRange.getValues();
  
  let migrated = 0;
  
  for (let i = 0; i < data.length; i++) {
    const plainTextPassword = String(data[i][6]).trim(); // Column H (password)
    
    // Skip if already hashed (hashed passwords are 64 characters long)
    if (plainTextPassword.length === 64 && /^[a-f0-9]+$/.test(plainTextPassword)) {
      continue;
    }
    
    // Hash the password
    const hashedPassword = hashPassword(plainTextPassword);
    
    // Update the cell with hashed password
    const rowNum = i + 3; // Row 3 is first data row
    sheet.getRange(rowNum, 8).setValue(hashedPassword); // Column H
    
    migrated++;
  }
  
  const result = {
    success: true,
    message: `Successfully migrated ${migrated} password(s) to secure hashes`,
    migrated: migrated
  };

  const afterStatus = checkPasswordSecurityStatus();

  safeLogAuditEvent(
    'Migrate',
    'Security',
    'UserPasswords',
    'Migrated stored user passwords to secure hashes',
    buildSecurityAuditSnapshot(beforeStatus),
    buildSecurityAuditSnapshot({
      totalUsers: afterStatus.totalUsers,
      hashedPasswords: afterStatus.hashedPasswords,
      plainTextPasswords: afterStatus.plainTextPasswords,
      migrated: migrated,
      secure: afterStatus.secure,
      message: result.message
    })
  );

  return result;
}

/**
 * Check if passwords in the system are hashed
 * @returns {object} Status information
 */
function checkPasswordSecurityStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  
  if (!sheet) {
    return { secure: false, message: "Users sheet not found" };
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    return { secure: true, message: "No users in system", totalUsers: 0 };
  }
  
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 7);
  const data = dataRange.getValues();
  
  let totalUsers = 0;
  let hashedPasswords = 0;
  let plainTextPasswords = 0;
  
  for (let i = 0; i < data.length; i++) {
    const password = String(data[i][6]).trim();
    if (!password) continue;
    
    totalUsers++;
    
    // Check if password appears to be hashed (64 hex characters)
    if (password.length === 64 && /^[a-f0-9]+$/.test(password)) {
      hashedPasswords++;
    } else {
      plainTextPasswords++;
    }
  }
  
  return {
    secure: plainTextPasswords === 0,
    totalUsers: totalUsers,
    hashedPasswords: hashedPasswords,
    plainTextPasswords: plainTextPasswords,
    message: plainTextPasswords > 0 
      ? `WARNING: ${plainTextPasswords} user(s) have plain text passwords!` 
      : "All passwords are securely hashed"
  };
}

function normalizeAcademicYearValue(value) {
  return String(value || '').trim();
}

function getActiveAcademicYearRecord() {
  const years = typeof getAcademicYearsData === 'function' ? getAcademicYearsData() : [];
  for (let i = 0; i < years.length; i++) {
    if (String(years[i].status || '').trim().toLowerCase() === 'active') {
      return years[i];
    }
  }
  return null;
}

function getActiveAcademicYearValue() {
  const activeYear = getActiveAcademicYearRecord();
  return activeYear ? normalizeAcademicYearValue(activeYear.academicYear) : '';
}

function isCurrentUserAdmin() {
  requireLogin();
  const user = getLoggedInUser();
  return !!(user && String(user.role || '').trim().toLowerCase() === 'admin');
}

function appendAcademicYearOverrideAuditDetails(details, guard) {
  if (!guard || !guard.adminOverride) return details;
  const base = String(details || '').trim();
  const note = `Admin override outside active academic year (record: ${guard.recordAcademicYear}, active: ${guard.activeAcademicYear})`;
  return base ? `${base} | ${note}` : note;
}

function enforceAcademicYearCrudSecurity(options) {
  requireLogin();
  const config = options || {};
  const action = String(config.action || 'Modify').trim();
  const module = String(config.module || 'Records').trim();
  const recordId = String(config.recordId || '').trim();
  const rawYears = Array.isArray(config.academicYears)
    ? config.academicYears
    : [config.academicYear];
  const normalizedYears = rawYears
    .map(normalizeAcademicYearValue)
    .filter(function(year, index, arr) {
      return arr.indexOf(year) === index;
    });
  const activeAcademicYear = getActiveAcademicYearValue();
  const displayAcademicYear = normalizedYears.length
    ? normalizedYears.map(function(year) { return year || 'Not set'; }).join(' / ')
    : 'Not set';
  const hasBlankYear = normalizedYears.some(function(year) { return !year; });
  const hasNonActiveYear = normalizedYears.some(function(year) {
    return !!year && year !== activeAcademicYear;
  });
  const oldValue = config.oldValue || null;
  const newValue = config.newValue || null;
  const actionLower = action.toLowerCase();

  if (!activeAcademicYear) {
    safeLogAuditEvent(
      'Denied',
      module,
      recordId,
      `Blocked ${actionLower} because no active academic year is configured.`,
      oldValue,
      newValue
    );

    return {
      success: false,
      allowed: false,
      message: 'No active academic year is configured. Please set one active academic year before continuing.',
      requiresAdminOverride: false,
      recordAcademicYear: displayAcademicYear,
      activeAcademicYear: '',
      adminOverride: false,
      code: 'NO_ACTIVE_ACADEMIC_YEAR'
    };
  }

  if (!hasBlankYear && !hasNonActiveYear) {
    return {
      success: true,
      allowed: true,
      recordAcademicYear: displayAcademicYear,
      activeAcademicYear: activeAcademicYear,
      adminOverride: false
    };
  }

  const isAdmin = isCurrentUserAdmin();
  const mismatchMessage = `Affected year(s): ${displayAcademicYear}; active year: ${activeAcademicYear}.`;

  if (!isAdmin) {
    safeLogAuditEvent(
      'Denied',
      module,
      recordId,
      `Blocked ${actionLower} outside active academic year. ${mismatchMessage} Admin account required.`,
      oldValue,
      newValue
    );

    return {
      success: false,
      allowed: false,
      message: `You cannot ${actionLower} records outside the active academic year. Affected year(s): ${displayAcademicYear}. Active academic year: ${activeAcademicYear}. Only Admin can continue.`,
      requiresAdminOverride: false,
      recordAcademicYear: displayAcademicYear,
      activeAcademicYear: activeAcademicYear,
      adminOverride: false,
      code: 'ACADEMIC_YEAR_RESTRICTED'
    };
  }

  if (config.overrideConfirmed === true) {
    return {
      success: true,
      allowed: true,
      recordAcademicYear: displayAcademicYear,
      activeAcademicYear: activeAcademicYear,
      adminOverride: true
    };
  }

  safeLogAuditEvent(
    'Warning',
    module,
    recordId,
    `Admin confirmation required before ${actionLower} outside active academic year. ${mismatchMessage}`,
    oldValue,
    newValue
  );

  return {
    success: false,
    allowed: false,
    requiresAdminOverride: true,
    message: `Affected year(s): ${displayAcademicYear}. The active academic year is ${activeAcademicYear}. Only an Admin can continue. Please confirm the warning to proceed.`,
    recordAcademicYear: displayAcademicYear,
    activeAcademicYear: activeAcademicYear,
    adminOverride: false,
    code: 'ACADEMIC_YEAR_ADMIN_CONFIRM'
  };
}
 