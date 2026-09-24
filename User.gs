// --- USERS FUNCTIONS ---

const MAX_LOGIN_ATTEMPTS = 4;
const LOGIN_LOCK_MINUTES = 30;
const USER_SHEET_FIRST_DATA_ROW = 3;
const USER_SHEET_FIRST_COL = 2;
const USER_SHEET_COL_COUNT = 9; // B through J

function getUserSheetColumnCount_(sheet) {
  const lastCol = sheet.getLastColumn();
  return Math.max(USER_SHEET_COL_COUNT, lastCol >= USER_SHEET_FIRST_COL ? lastCol - USER_SHEET_FIRST_COL + 1 : USER_SHEET_COL_COUNT);
}

function parseLockedUntilValue_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatLockedUntilForDisplay_(value) {
  const lockedUntil = parseLockedUntilValue_(value);
  if (!lockedUntil) return '';
  try {
    return Utilities.formatDate(lockedUntil, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  } catch (e) {
    return lockedUntil.toISOString();
  }
}

function isUserAccountLocked_(loginTrials, lockedUntilValue) {
  const lockedUntil = parseLockedUntilValue_(lockedUntilValue);
  if (!lockedUntil) return false;
  return lockedUntil.getTime() > Date.now();
}

function getRemainingLockMinutes_(lockedUntilValue) {
  const lockedUntil = parseLockedUntilValue_(lockedUntilValue);
  if (!lockedUntil) return 0;
  const remainingMs = lockedUntil.getTime() - Date.now();
  return remainingMs > 0 ? Math.ceil(remainingMs / 60000) : 0;
}

function resetUserLoginLock_(sheet, rowNum) {
  sheet.getRange(rowNum, 9, 1, 2).setValues([[0, '']]);
}

/**
 * Same as resetUserLoginLock_() but never throws.
 * Clearing the failed-attempt counter is bookkeeping; it must not be able to
 * block a successful sign-in (e.g. when the effective account can only view the
 * spreadsheet, which is a common state on an imported copy of the project).
 */
function safeResetUserLoginLock_(sheet, rowNum) {
  try {
    resetUserLoginLock_(sheet, rowNum);
    return true;
  } catch (e) {
    Logger.log('Could not reset login lock for row ' + rowNum + ': ' + e.message);
    return false;
  }
}

/**
 * Same as recordFailedLoginAttempt_() but degrades gracefully: if the counter
 * cannot be written we still report "invalid credentials" instead of crashing.
 */
function safeRecordFailedLoginAttempt_(sheet, rowNum, currentAttempts) {
  try {
    return recordFailedLoginAttempt_(sheet, rowNum, currentAttempts);
  } catch (e) {
    Logger.log('Could not record failed login attempt for row ' + rowNum + ': ' + e.message);
    return { locked: false, attempts: (parseInt(currentAttempts, 10) || 0) + 1, message: 'Invalid username or password' };
  }
}

/**
 * Re-write a legacy-format password (plain text or "sha256:" prefixed) using the
 * current salt, so every copy of the project stores the same hash.
 * Failures are ignored on purpose.
 */
function safeUpgradeStoredPassword_(sheet, rowNum, plainPassword) {
  try {
    const newHash = hashPassword(plainPassword);
    if (!newHash) return false;
    sheet.getRange(rowNum, 8).setValue(newHash);
    invalidateCacheOnModify('Users');
    return true;
  } catch (e) {
    Logger.log('Could not upgrade stored password hash on row ' + rowNum + ': ' + e.message);
    return false;
  }
}

function recordFailedLoginAttempt_(sheet, rowNum, currentAttempts) {
  const newAttempts = (parseInt(currentAttempts, 10) || 0) + 1;

  if (newAttempts >= MAX_LOGIN_ATTEMPTS) {
    const lockUntil = new Date(Date.now() + LOGIN_LOCK_MINUTES * 60 * 1000);
    sheet.getRange(rowNum, 9, 1, 2).setValues([[newAttempts, lockUntil]]);
    return {
      locked: true,
      attempts: newAttempts,
      message: 'Too many failed login attempts. Account locked for ' + LOGIN_LOCK_MINUTES + ' minutes.'
    };
  }

  sheet.getRange(rowNum, 9).setValue(newAttempts);
  return {
    locked: false,
    attempts: newAttempts,
    message: 'Invalid username or password'
  };
}

// 1. Fetch Users Data (EXCLUDES Password for security)
function getUsersData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < USER_SHEET_FIRST_DATA_ROW) return [];

  const colCount = getUserSheetColumnCount_(sheet);
  const dataRange = sheet.getRange(USER_SHEET_FIRST_DATA_ROW, USER_SHEET_FIRST_COL, lastRow - (USER_SHEET_FIRST_DATA_ROW - 1), colCount);
  const data = dataRange.getValues();

  return data.map((row) => {
    const loginTrials = parseInt(row[7], 10) || 0;
    const lockedUntil = row[8];
    const locked = isUserAccountLocked_(loginTrials, lockedUntil);

    return {
      userId: String(row[0]).trim(),
      googleEmail: String(row[1]).trim(),
      fullName: String(row[2]).trim(),
      role: String(row[3]).trim(),
      accountStatus: String(row[4]).trim(),
      username: String(row[5]).trim(),
      password: "••••••••",
      loginTrials: loginTrials,
      lockedUntil: formatLockedUntilForDisplay_(lockedUntil),
      isLocked: locked,
      lockMinutesRemaining: locked ? getRemainingLockMinutes_(lockedUntil) : 0
    };
  });
}

// 2. Helper to generate User ID (e.g. USR-1001)
function generateNextUserId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "USR-1001";

  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  let maxIdNum = 1000;
  values.forEach((row) => {
    const idStr = String(row[0]).trim();
    if (idStr.startsWith("USR-")) {
      const num = parseInt(idStr.substring(4), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return "USR-" + (maxIdNum + 1);
}

// 3. Helper: Find row by User ID - OPTIMIZED WITH INDEX
function findUserRowById(sheet, userId) {
  // Use indexed lookup for O(1) performance
  const rowNumber = findRowByIdIndexed('Users', userId);
  
  if (rowNumber !== -1) {
    return rowNumber;
  }
  
  // Fallback to linear search
  Logger.log('WARNING: User index lookup failed, falling back to linear search');
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return -1;
  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(userId).trim()) {
      return i + 3;
    }
  }
  return -1;
}

function buildUserAuditSnapshot(userData) {
  if (!userData) return null;

  return buildAuditSnapshot({
    userId: userData.userId,
    googleEmail: userData.googleEmail,
    fullName: userData.fullName,
    role: userData.role,
    accountStatus: userData.accountStatus,
    username: userData.username,
    passwordChanged: userData.passwordChanged === true ? true : undefined
  });
}

// 4. Add New User (Create) - Now includes Username and Password
function addUser(userData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) throw new Error("Users worksheet not found.");

  const nextId = generateNextUserId(sheet);
  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  // Hash the password before storing
  const plainPassword = userData.password || "Password123";
  const hashedPassword = hashPassword(plainPassword);

  // Write to Column B to Column J
  sheet.getRange(targetRow, USER_SHEET_FIRST_COL, 1, USER_SHEET_COL_COUNT).setValues([
    [
      nextId,
      userData.googleEmail,
      userData.fullName,
      userData.role,
      userData.accountStatus || "Active",
      userData.username,
      hashedPassword,
      0,
      ""
    ],
  ]);

  // Invalidate cache and index
  invalidateCacheOnModify('Users');
  invalidateIndex('Users');

  const createdUserSnapshot = buildUserAuditSnapshot({
    userId: nextId,
    googleEmail: userData.googleEmail,
    fullName: userData.fullName,
    role: userData.role,
    accountStatus: userData.accountStatus || 'Active',
    username: userData.username,
    passwordChanged: true
  });

  safeLogAuditEvent(
    'Create',
    'Users',
    nextId,
    'Created user ' + String(userData.fullName || nextId).trim(),
    null,
    createdUserSnapshot
  );

  return { success: true, userId: nextId };
}

// 5. Update User (Update) - Now includes Username and Password
function updateUser(userId, userData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) throw new Error("Users worksheet not found.");

  const existingUser = getUsersData().find((user) => String(user.userId).trim() === String(userId).trim()) || null;
  const row = findUserRowById(sheet, userId);
  if (row === -1) throw new Error("User record not found.");

  // Keep the existing password unless the admin actually types a new one.
  const submittedPassword = typeof userData.password === 'string' ? userData.password.trim() : '';
  const currentStoredPassword = String(sheet.getRange(row, 8).getValue() || '').trim();
  let passwordToStore = currentStoredPassword;

  if (submittedPassword) {
    // Only hash if it's not already a hash (64 hex characters)
    if (!(submittedPassword.length === 64 && /^[a-f0-9]+$/.test(submittedPassword))) {
      passwordToStore = hashPassword(submittedPassword);
    } else {
      passwordToStore = submittedPassword;
    }
  }

  // Update from Col C(3) to Col H(8) -> 6 columns
  sheet
    .getRange(row, 3, 1, 6)
    .setValues([
      [
        userData.googleEmail,
        userData.fullName,
        userData.role,
        userData.accountStatus,
        userData.username,
        passwordToStore,
      ],
    ]);

  if (String(userData.accountStatus || '').trim().toLowerCase() === 'active') {
    resetUserLoginLock_(sheet, row);
  }

  // Invalidate cache and index
  invalidateCacheOnModify('Users');
  invalidateIndex('Users');

  const updatedUserSnapshot = buildUserAuditSnapshot(Object.assign({}, existingUser || {}, userData, {
    userId: userId,
    passwordChanged: !!userData.password
  }));

  safeLogAuditEvent(
    'Update',
    'Users',
    userId,
    'Updated user ' + String((updatedUserSnapshot && updatedUserSnapshot.fullName) || userId).trim(),
    buildUserAuditSnapshot(existingUser),
    updatedUserSnapshot
  );

  return { success: true };
}

// 6b. Unlock a temporarily locked user account (Admin only)
function unlockUserAccount(userId) {
  if (typeof isCurrentUserAdmin === 'function' && !isCurrentUserAdmin()) {
    throw new Error('Only administrators can unlock user accounts.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) throw new Error('Users worksheet not found.');

  const existingUser = getUsersData().find((user) => String(user.userId).trim() === String(userId).trim()) || null;
  const row = findUserRowById(sheet, userId);
  if (row === -1) throw new Error('User record not found.');

  if (!existingUser || (!existingUser.isLocked && !(existingUser.loginTrials > 0))) {
    return { success: true, message: 'Account is not locked.' };
  }

  resetUserLoginLock_(sheet, row);

  invalidateCacheOnModify('Users');
  invalidateIndex('Users');

  safeLogAuditEvent(
    'Update',
    'Users',
    userId,
    'Unlocked user account ' + String((existingUser.fullName || userId)).trim(),
    buildUserAuditSnapshot(existingUser),
    buildUserAuditSnapshot(Object.assign({}, existingUser, {
      loginTrials: 0,
      lockedUntil: '',
      isLocked: false,
      lockMinutesRemaining: 0
    }))
  );

  return { success: true, message: 'Account unlocked successfully.' };
}

// 6. Delete User (Delete)
function deleteUser(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) throw new Error("Users worksheet not found.");

  const existingUser = getUsersData().find((user) => String(user.userId).trim() === String(userId).trim()) || null;
  const row = findUserRowById(sheet, userId);
  if (row === -1) throw new Error("User record not found.");

  // SECURITY: Prevent deleting the very last Admin
  const allUsers = getUsersData();
  const admins = allUsers.filter((u) => u.role === "Admin");

  if (admins.length === 1 && admins[0].userId === userId) {
    throw new Error(
      "Cannot delete the last Admin user. You would lock yourself out of the system!",
    );
  }

  sheet.deleteRow(row);
  
  // Invalidate cache and index
  invalidateCacheOnModify('Users');
  invalidateIndex('Users');

  safeLogAuditEvent(
    'Delete',
    'Users',
    userId,
    'Deleted user ' + String(((existingUser && existingUser.fullName) || userId)).trim(),
    buildUserAuditSnapshot(existingUser),
    null
  );
  
  return { success: true };
}

// --- NEW LOGIN LOGIC BASED ON THE SHEET ---

// 7. Verify Credentials directly from the Users Sheet - SECURE VERSION
function loginUser(username, password) {
  const ss = typeof getSpreadsheet_ === 'function' ? getSpreadsheet_() : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return {
      success: false,
      message: "This Apps Script project is not connected to a spreadsheet. Open the spreadsheet, then Extensions > Apps Script, or run setSpreadsheetId('<spreadsheet id>') once."
    };
  }
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "Users worksheet not found" };

  const lastRow = sheet.getLastRow();
  if (lastRow < USER_SHEET_FIRST_DATA_ROW) return { success: false, message: "No users found" };

  const colCount = getUserSheetColumnCount_(sheet);
  const dataRange = sheet.getRange(USER_SHEET_FIRST_DATA_ROW, USER_SHEET_FIRST_COL, lastRow - (USER_SHEET_FIRST_DATA_ROW - 1), colCount);
  const data = dataRange.getValues();
  const normalizedUsername = String(username || '').trim();

  for (let i = 0; i < data.length; i++) {
    const rowNum = i + USER_SHEET_FIRST_DATA_ROW;
    const userId = String(data[i][0]).trim();
    const fullName = String(data[i][2]).trim();
    const role = String(data[i][3]).trim() || "User";
    const accountStatus = String(data[i][4]).trim();
    const user = String(data[i][5]).trim();
    const storedPasswordHash = String(data[i][6]).trim();
    const loginTrials = parseInt(data[i][7], 10) || 0;
    const lockedUntilValue = data[i][8];

    // Case-insensitive username matching
    if (user.toLowerCase() !== normalizedUsername.toLowerCase()) continue;

    // Check account status: only block if explicitly inactive / disabled
    if (accountStatus && ['inactive', 'suspended', 'disabled', 'blocked'].indexOf(accountStatus.toLowerCase()) !== -1) {
      return { success: false, message: "Account is inactive. Contact administrator." };
    }

    // Verify password first
    const check = verifyPasswordWithDetails_(password, storedPasswordHash);
    if (check.matched) {
      // Correct password! Immediately clear failed attempts and unlock account
      safeResetUserLoginLock_(sheet, rowNum);

      // Quietly normalise plain-text or legacy formats to the current salted hash
      if (check.mode !== 'salted-current' && check.mode !== 'salted-legacy') {
        safeUpgradeStoredPassword_(sheet, rowNum, password);
      }

      const sessionId = createSession(userId, normalizedUsername, role, fullName);
      storeSessionIdInProperties(sessionId);

      return {
        success: true,
        role: role,
        sessionId: sessionId,
        username: normalizedUsername,
        fullName: fullName
      };
    }

    // Password did NOT match: check account lockout
    const lockedUntil = parseLockedUntilValue_(lockedUntilValue);
    if (lockedUntil && lockedUntil.getTime() <= Date.now()) {
      safeResetUserLoginLock_(sheet, rowNum);
    } else if (isUserAccountLocked_(loginTrials, lockedUntilValue)) {
      const minutesLeft = getRemainingLockMinutes_(lockedUntilValue) || LOGIN_LOCK_MINUTES;
      return {
        success: false,
        message: 'Account temporarily locked after ' + MAX_LOGIN_ATTEMPTS + ' failed attempts. Try again in ' + minutesLeft + ' minute(s).'
      };
    }

    const failureResult = safeRecordFailedLoginAttempt_(sheet, rowNum, loginTrials);
    return { success: false, message: failureResult.message };
  }

  return { success: false, message: "Invalid username or password" };
}

// 8. Check if user is logged in - SECURE VERSION
function checkSession(paramSessionId) {
  let sessionId = paramSessionId || null;
  
  if (!sessionId) {
    try {
      const props = PropertiesService.getUserProperties();
      sessionId = props.getProperty('CURRENT_SESSION_ID');
    } catch (e) {
      // UserProperties might be unavailable in some environments
    }
  }
  
  if (!sessionId) {
    // Try to migrate old session format
    const migrationResult = migrateOldSessionToSecure();
    if (migrationResult.migrated) {
      Logger.log('Auto-migrated old session for: ' + migrationResult.username);
      return true;
    }
    return false;
  }
  
  // Validate session token
  const session = validateSession(sessionId);
  if (!session) {
    // Session expired or invalid - clean up
    try {
      storeSessionIdInProperties(null);
    } catch (e) {}
    return false;
  }
  
  // If a valid sessionId was provided, keep UserProperties in sync
  try {
    const props = PropertiesService.getUserProperties();
    if (props && sessionId) {
      props.setProperty('CURRENT_SESSION_ID', sessionId);
    }
  } catch (e) {}
  
  // Session is valid and activity timestamp was updated by validateSession
  return true;
}

// 9. Log out - SECURE VERSION
function logoutUser(paramSessionId) {
  let sessionId = paramSessionId || null;
  if (!sessionId) {
    try {
      const props = PropertiesService.getUserProperties();
      sessionId = props.getProperty('CURRENT_SESSION_ID');
    } catch (e) {}
  }
  
  if (sessionId) {
    // Destroy the session in cache
    destroySession(sessionId);
  }
  
  // Clear session ID from properties
  storeSessionIdInProperties(null);
  
  // Also clean up old session format (backward compatibility)
  try {
    const oldProps = PropertiesService.getScriptProperties();
    oldProps.deleteProperty("IS_LOGGED_IN");
    oldProps.deleteProperty("LOGGED_IN_ROLE");
    oldProps.deleteProperty("LOGGED_IN_USER");
    oldProps.deleteProperty("LOGGED_IN_AT");
  } catch (e) {}
  
  return { success: true };
}

// 10. Get logged-in user info - SECURE & DYNAMIC VERSION
function getLoggedInUser(paramSessionId) {
  let sessionId = paramSessionId || null;
  
  // 1. If not provided directly, try UserProperties
  if (!sessionId) {
    try {
      const props = PropertiesService.getUserProperties();
      sessionId = props.getProperty('CURRENT_SESSION_ID');
    } catch (e) {}
  }
  
  // 2. If still no session, check for legacy migration
  if (!sessionId) {
    try {
      const oldProps = PropertiesService.getScriptProperties();
      if (oldProps && oldProps.getProperty("IS_LOGGED_IN") === "true") {
        const migrationResult = migrateOldSessionToSecure();
        if (migrationResult.migrated) {
          sessionId = migrationResult.sessionId;
        }
      }
    } catch (e) {}
  }
  
  if (!sessionId) return null;
  
  // 3. Validate session from CacheService
  const session = validateSession(sessionId);
  if (!session) {
    storeSessionIdInProperties(null);
    return null;
  }
  
  // 4. LIVE LOOKUP FROM USERS SHEET:
  // Instead of only returning stale session.fullName from the cache, dynamically
  // query the Users sheet so edits to Full Name, Role, or Status in the sheet or UI
  // are immediately reflected without requiring the user to re-login!
  let resolvedFullName = session.fullName || session.username || 'User';
  let resolvedRole = session.role || 'User';
  let resolvedUsername = session.username || '';
  let resolvedUserId = session.userId || '';

  try {
    const users = typeof getUsersData === 'function' ? getUsersData() : [];
    const matchedUser = users.find(function(u) {
      if (resolvedUserId && String(u.userId).trim().toLowerCase() === String(resolvedUserId).trim().toLowerCase()) {
        return true;
      }
      if (resolvedUsername && String(u.username).trim().toLowerCase() === String(resolvedUsername).trim().toLowerCase()) {
        return true;
      }
      return false;
    });

    if (matchedUser) {
      if (matchedUser.fullName) resolvedFullName = matchedUser.fullName;
      if (matchedUser.role) resolvedRole = matchedUser.role;
      if (matchedUser.username) resolvedUsername = matchedUser.username;
      if (matchedUser.userId) resolvedUserId = matchedUser.userId;

      // Keep cached session data synchronized with the sheet
      if (session.fullName !== resolvedFullName || session.role !== resolvedRole) {
        session.fullName = resolvedFullName;
        session.role = resolvedRole;
        try {
          const cache = CacheService.getUserCache();
          cache.put('session_' + sessionId, JSON.stringify(session), 28800);
          CacheService.getScriptCache().put('session_' + sessionId, JSON.stringify(session), 28800);
        } catch (cErr) {}
      }
    }
  } catch (lookupErr) {
    Logger.log('Warning in live user lookup for getLoggedInUser: ' + lookupErr.message);
  }

  return {
    userId: resolvedUserId,
    username: resolvedUsername,
    fullName: resolvedFullName,
    role: resolvedRole,
    initials: getInitials(resolvedFullName),
    sessionId: sessionId
  };
}

// Helper to get initials from full name
function getInitials(name) {
  if (!name) return "U";
  const parts = name.trim().split(/\s+/);
  if (parts.length === 1) return parts[0].charAt(0).toUpperCase();
  return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
}

// --- ADMIN FORGOT PASSWORD RECOVERY ---

/**
 * Helper to mask an email address for privacy (e.g. s***8@gmail.com)
 */
function maskEmail_(email) {
  if (!email || email.indexOf('@') === -1) return '***';
  const parts = email.split('@');
  const user = parts[0];
  const domain = parts[1];
  if (user.length <= 2) return user.charAt(0) + '***@' + domain;
  return user.charAt(0) + '***' + user.charAt(user.length - 1) + '@' + domain;
}

/**
 * 1. Request Password Reset for an Admin user.
 * Generates a 6-digit verification code and emails it to the Admin's registered email.
 * @param {string} identifier - Admin username or email
 * @returns {object} { success: boolean, message: string, resetToken?: string, maskedEmail?: string }
 */
function requestAdminPasswordReset(identifier) {
  try {
    const rawId = String(identifier || '').trim();
    if (!rawId) {
      return { success: false, message: 'Please enter your Admin username or email address.' };
    }

    const ss = typeof getSpreadsheet_ === 'function' ? getSpreadsheet_() : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { success: false, message: 'Spreadsheet connection unavailable.' };
    }
    const sheet = ss.getSheetByName('Users');
    if (!sheet) {
      return { success: false, message: 'Users sheet not found.' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < USER_SHEET_FIRST_DATA_ROW) {
      return { success: false, message: 'No users found in the system.' };
    }

    const colCount = getUserSheetColumnCount_(sheet);
    const dataRange = sheet.getRange(USER_SHEET_FIRST_DATA_ROW, USER_SHEET_FIRST_COL, lastRow - (USER_SHEET_FIRST_DATA_ROW - 1), colCount);
    const data = dataRange.getValues();

    const normalizedId = rawId.toLowerCase();
    let targetUser = null;

    for (let i = 0; i < data.length; i++) {
      const userId = String(data[i][0] || '').trim();
      const googleEmail = String(data[i][1] || '').trim();
      const fullName = String(data[i][2] || '').trim();
      const role = String(data[i][3] || '').trim();
      const username = String(data[i][5] || '').trim();

      if (username.toLowerCase() === normalizedId || googleEmail.toLowerCase() === normalizedId) {
        targetUser = {
          userId: userId,
          googleEmail: googleEmail,
          fullName: fullName,
          role: role,
          username: username
        };
        break;
      }
    }

    if (!targetUser) {
      return { success: false, message: 'No user account found with that username or email.' };
    }

    // SECURITY CHECK: Must be an Admin role
    if (String(targetUser.role || '').trim().toLowerCase() !== 'admin') {
      return {
        success: false,
        message: 'Password self-reset is only available for Administrator accounts. Please contact an Admin to reset your password.'
      };
    }

    // Determine destination email: user's googleEmail or system owner email
    let recipientEmail = targetUser.googleEmail;
    if (!recipientEmail || recipientEmail.indexOf('@') === -1) {
      if (typeof getSystemOwnerEmail === 'function') {
        recipientEmail = getSystemOwnerEmail();
      }
    }

    if (!recipientEmail || recipientEmail.indexOf('@') === -1) {
      return {
        success: false,
        message: 'No valid email address is linked to this Admin account in the Users sheet. Please update Column C in the Users sheet or use the Master Password.'
      };
    }

    // Generate 6-digit verification code
    const code = String(Math.floor(100000 + Math.random() * 900000));
    const resetToken = Utilities.getUuid();
    const expiresAt = Date.now() + 15 * 60 * 1000; // 15 minutes

    const sessionPayload = JSON.stringify({
      userId: targetUser.userId,
      username: targetUser.username,
      email: recipientEmail,
      code: code,
      expiresAt: expiresAt,
      attempts: 0
    });

    // Store in ScriptCache (15 minutes = 900 seconds)
    try {
      CacheService.getScriptCache().put('pw_reset_' + resetToken, sessionPayload, 900);
    } catch (e) {
      Logger.log('CacheService put error: ' + e.message);
    }

    // Also store in ScriptProperties as backup
    try {
      PropertiesService.getScriptProperties().setProperty('pw_reset_' + resetToken, sessionPayload);
    } catch (e) {}

    // Send the email with the code
    const schoolName = typeof getSchoolNameFromSettings === 'function' ? getSchoolNameFromSettings() : 'School Management System';
    const subject = '🔐 Password Reset Code for Admin (' + targetUser.username + ')';
    const plainBody = 
      'Hello ' + (targetUser.fullName || targetUser.username) + ',\n\n' +
      'A request was received to reset the password for your Admin account (' + targetUser.username + ') on ' + schoolName + '.\n\n' +
      'Your 6-digit verification code is:\n\n' +
      '   👉 ' + code + ' 👈\n\n' +
      'This code is valid for 15 minutes.\n\n' +
      'Enter this code on the login page to choose a new password and unlock your account.\n\n' +
      'If you did not request this reset, your account password remains unchanged. Please check with your team.';

    try {
      GmailApp.sendEmail(recipientEmail, subject, plainBody, {
        name: schoolName,
        noReply: true
      });
    } catch (gErr) {
      try {
        MailApp.sendEmail({
          to: recipientEmail,
          subject: subject,
          body: plainBody,
          name: schoolName
        });
      } catch (mErr) {
        Logger.log('Failed to send reset email: ' + mErr.message);
        return {
          success: false,
          message: 'Failed to send email to ' + maskEmail_(recipientEmail) + ': ' + mErr.message + '. You can also use the Master Password to reset.'
        };
      }
    }

    return {
      success: true,
      resetToken: resetToken,
      maskedEmail: maskEmail_(recipientEmail),
      message: 'A 6-digit verification code has been sent to ' + maskEmail_(recipientEmail) + '.'
    };
  } catch (err) {
    Logger.log('Error in requestAdminPasswordReset: ' + err.toString());
    return { success: false, message: 'System error: ' + err.message };
  }
}

/**
 * 2. Verify Code & Set New Password for Admin.
 * @param {string} resetToken
 * @param {string} code
 * @param {string} newPassword
 * @returns {object} { success: boolean, message: string }
 */
function verifyAndResetAdminPassword(resetToken, code, newPassword) {
  try {
    const token = String(resetToken || '').trim();
    const providedCode = String(code || '').trim();
    const pass = String(newPassword || '').trim();

    if (!token) return { success: false, message: 'Invalid or missing reset token.' };
    if (!providedCode) return { success: false, message: 'Please enter the 6-digit verification code.' };
    if (!pass || pass.length < 6) return { success: false, message: 'Password must be at least 6 characters long.' };

    // Retrieve session payload from CacheService or ScriptProperties
    let payloadStr = '';
    try {
      payloadStr = CacheService.getScriptCache().get('pw_reset_' + token);
    } catch (e) {}

    if (!payloadStr) {
      try {
        payloadStr = PropertiesService.getScriptProperties().getProperty('pw_reset_' + token);
      } catch (e) {}
    }

    if (!payloadStr) {
      return { success: false, message: 'Password reset session has expired or is invalid. Please request a new code.' };
    }

    const session = JSON.parse(payloadStr);
    if (!session || Date.now() > session.expiresAt) {
      return { success: false, message: 'This verification code has expired. Please request a new one.' };
    }

    // Check code attempt count
    session.attempts = (session.attempts || 0) + 1;
    if (session.attempts > 5) {
      try {
        CacheService.getScriptCache().remove('pw_reset_' + token);
        PropertiesService.getScriptProperties().deleteProperty('pw_reset_' + token);
      } catch (e) {}
      return { success: false, message: 'Too many incorrect attempts. Please request a new code.' };
    }

    if (session.code !== providedCode) {
      try {
        CacheService.getScriptCache().put('pw_reset_' + token, JSON.stringify(session), 900);
      } catch (e) {}
      return { success: false, message: 'Invalid verification code. Please check your email and try again.' };
    }

    // CODE MATCHES! Perform password update in Users sheet
    const ss = typeof getSpreadsheet_ === 'function' ? getSpreadsheet_() : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet connection unavailable.' };
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, message: 'Users sheet not found.' };

    const rowNum = findUserRowById(sheet, session.userId);
    if (rowNum === -1) {
      return { success: false, message: 'User record could not be found in Users sheet.' };
    }

    // Hash password with active salt (and ensure salt is stored in sheet)
    const activeSalt = getSalt();
    writeSaltToSheet_(activeSalt);
    const hashedPassword = hashPassword(pass);

    // Update password in Column H (8)
    sheet.getRange(rowNum, 8).setValue(hashedPassword);
    // Ensure status is Active in Column F (6)
    sheet.getRange(rowNum, 6).setValue('Active');
    // Clear failed attempts and lockout in Column I & J (9 & 10)
    sheet.getRange(rowNum, 9, 1, 2).setValues([[0, '']]);

    // Clean up reset token
    try {
      CacheService.getScriptCache().remove('pw_reset_' + token);
      PropertiesService.getScriptProperties().deleteProperty('pw_reset_' + token);
    } catch (e) {}

    // Invalidate caches
    invalidateCacheOnModify('Users');
    invalidateIndex('Users');

    safeLogAuditEvent(
      'Update',
      'Users',
      session.userId,
      'Admin password reset via self-service email verification for ' + session.username,
      null,
      null
    );

    return {
      success: true,
      message: 'Admin password reset successfully! You can now sign in with your new password.'
    };
  } catch (err) {
    Logger.log('Error in verifyAndResetAdminPassword: ' + err.toString());
    return { success: false, message: 'System error: ' + err.message };
  }
}

/**
 * 3. Immediate Admin Password Reset using Master Password (no email required).
 * @param {string} username
 * @param {string} masterPassword
 * @param {string} newPassword
 * @returns {object} { success: boolean, message: string }
 */
function resetAdminPasswordWithMasterPassword(username, masterPassword, newPassword) {
  try {
    const userNorm = String(username || '').trim().toLowerCase();
    const masterPass = String(masterPassword || '').trim();
    const pass = String(newPassword || '').trim();

    if (!userNorm) return { success: false, message: 'Please enter your Admin username.' };
    if (!masterPass) return { success: false, message: 'Please enter the Master Password.' };
    if (!pass || pass.length < 6) return { success: false, message: 'New password must be at least 6 characters long.' };

    // Verify Master Password from Settings sheet
    const storedMaster = typeof getSystemParameter === 'function' ? (getSystemParameter('Master Password') || getSystemParameter('MasterPass')) : null;
    if (!storedMaster) {
      return { success: false, message: 'No Master Password is configured in the Settings sheet. Please use email verification or reset via Google Sheets menu.' };
    }

    if (typeof isMasterPasswordMatch === 'function') {
      if (!isMasterPasswordMatch(masterPass, storedMaster)) {
        return { success: false, message: 'Incorrect Master Password.' };
      }
    } else {
      if (masterPass !== String(storedMaster)) {
        return { success: false, message: 'Incorrect Master Password.' };
      }
    }

    const ss = typeof getSpreadsheet_ === 'function' ? getSpreadsheet_() : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet connection unavailable.' };
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, message: 'Users sheet not found.' };

    const lastRow = sheet.getLastRow();
    if (lastRow < USER_SHEET_FIRST_DATA_ROW) return { success: false, message: 'No users found.' };

    const colCount = getUserSheetColumnCount_(sheet);
    const data = sheet.getRange(USER_SHEET_FIRST_DATA_ROW, USER_SHEET_FIRST_COL, lastRow - (USER_SHEET_FIRST_DATA_ROW - 1), colCount).getValues();

    let targetRow = -1;
    let targetUserId = '';
    let targetRole = '';

    for (let i = 0; i < data.length; i++) {
      const u = String(data[i][5] || '').trim().toLowerCase();
      if (u === userNorm) {
        targetRow = USER_SHEET_FIRST_DATA_ROW + i;
        targetUserId = String(data[i][0] || '').trim();
        targetRole = String(data[i][3] || '').trim();
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: 'Admin username "' + username + '" not found.' };
    }

    if (targetRole.toLowerCase() !== 'admin') {
      return { success: false, message: 'Master Password reset is only available for Admin accounts.' };
    }

    const activeSalt = getSalt();
    writeSaltToSheet_(activeSalt);
    const hashedPassword = hashPassword(pass);

    sheet.getRange(targetRow, 8).setValue(hashedPassword);
    sheet.getRange(targetRow, 6).setValue('Active');
    sheet.getRange(targetRow, 9, 1, 2).setValues([[0, '']]);

    invalidateCacheOnModify('Users');
    invalidateIndex('Users');

    safeLogAuditEvent(
      'Update',
      'Users',
      targetUserId,
      'Admin password reset using Master Password for ' + username,
      null,
      null
    );

    return {
      success: true,
      message: 'Admin password reset successfully! You can now sign in with your new password.'
    };
  } catch (err) {
    Logger.log('Error in resetAdminPasswordWithMasterPassword: ' + err.toString());
    return { success: false, message: 'System error: ' + err.message };
  }
}
