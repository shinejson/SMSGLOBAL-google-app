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

  // Hash password if it's being updated and not already hashed
  let passwordToStore = userData.password;
  if (passwordToStore) {
    // Only hash if it's not already a hash (64 hex characters)
    if (!(passwordToStore.length === 64 && /^[a-f0-9]+$/.test(passwordToStore))) {
      passwordToStore = hashPassword(passwordToStore);
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

    if (user !== normalizedUsername) continue;

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

    if (accountStatus.toLowerCase() !== "active") {
      return { success: false, message: "Account is inactive. Contact administrator." };
    }

    const check = verifyPasswordWithDetails_(password, storedPasswordHash);
    if (check.matched) {
      // A correct password must never be blocked by a bookkeeping write
      // (e.g. the account running the script only has view access to the file).
      safeResetUserLoginLock_(sheet, rowNum);

      // Quietly normalise legacy formats to the current salted hash so that the
      // original project and any imported copy keep agreeing on the stored hash.
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

    const failureResult = safeRecordFailedLoginAttempt_(sheet, rowNum, loginTrials);
    return { success: false, message: failureResult.message };
  }

  return { success: false, message: "Invalid username or password" };
}

// 8. Check if user is logged in - SECURE VERSION
function checkSession() {
  // Get session ID from user properties
  const props = PropertiesService.getUserProperties();
  const sessionId = props.getProperty('CURRENT_SESSION_ID');
  
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
    storeSessionIdInProperties(null);
    return false;
  }
  
  // Session is valid and activity timestamp was updated by validateSession
  return true;
}

// 9. Log out - SECURE VERSION
function logoutUser() {
  // Get session ID
  const props = PropertiesService.getUserProperties();
  const sessionId = props.getProperty('CURRENT_SESSION_ID');
  
  if (sessionId) {
    // Destroy the session in cache
    destroySession(sessionId);
  }
  
  // Clear session ID from properties
  storeSessionIdInProperties(null);
  
  // Also clean up old session format (backward compatibility)
  const oldProps = PropertiesService.getScriptProperties();
  oldProps.deleteProperty("IS_LOGGED_IN");
  oldProps.deleteProperty("LOGGED_IN_ROLE");
  oldProps.deleteProperty("LOGGED_IN_USER");
  oldProps.deleteProperty("LOGGED_IN_AT");
  
  return { success: true };
}

// 10. Get logged-in user info - SECURE VERSION
function getLoggedInUser() {
  // Get session ID from user properties
  const props = PropertiesService.getUserProperties();
  const sessionId = props.getProperty('CURRENT_SESSION_ID');
  
  if (!sessionId) {
    // Try old format for backward compatibility
    const oldProps = PropertiesService.getScriptProperties();
    if (oldProps.getProperty("IS_LOGGED_IN") === "true") {
      const migrationResult = migrateOldSessionToSecure();
      if (migrationResult.migrated) {
        return getLoggedInUser(); // Recursive call after migration
      }
    }
    return null;
  }
  
  // Validate and get session data
  const session = validateSession(sessionId);
  if (!session) {
    storeSessionIdInProperties(null);
    return null;
  }
  
  return {
    userId: session.userId,
    username: session.username,
    fullName: session.fullName,
    role: session.role,
    initials: getInitials(session.fullName),
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
