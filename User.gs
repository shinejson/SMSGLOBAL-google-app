// --- USERS FUNCTIONS ---

// 1. Fetch Users Data (EXCLUDES Password for security)
function getUsersData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  // Row 3 to Last, Col B(2) to Col H(8) => 7 columns
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 7);
  const data = dataRange.getValues();

  return data.map((row) => ({
    userId: String(row[0]).trim(),
    googleEmail: String(row[1]).trim(),
    fullName: String(row[2]).trim(),
    role: String(row[3]).trim(),
    accountStatus: String(row[4]).trim(),
    username: String(row[5]).trim(),
    // Password is intentionally excluded for security
    password: "••••••••", // Masked password for display purposes
  }));
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

  // Write to Column B to Column H (7 columns)
  sheet.getRange(targetRow, 2, 1, 7).setValues([
    [
      nextId,
      userData.googleEmail,
      userData.fullName,
      userData.role,
      userData.accountStatus || "Active",
      userData.username,
      hashedPassword, // Store hashed password
    ],
  ]);

  // Invalidate cache and index
  invalidateCacheOnModify('Users');
  invalidateIndex('Users');

  return { success: true, userId: nextId };
}

// 5. Update User (Update) - Now includes Username and Password
function updateUser(userId, userData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) throw new Error("Users worksheet not found.");

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

  // Invalidate cache and index
  invalidateCacheOnModify('Users');
  invalidateIndex('Users');

  return { success: true };
}

// 6. Delete User (Delete)
function deleteUser(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) throw new Error("Users worksheet not found.");

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
  
  return { success: true };
}

// --- NEW LOGIN LOGIC BASED ON THE SHEET ---

// 7. Verify Credentials directly from the Users Sheet - SECURE VERSION
function loginUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "Users worksheet not found" };

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return { success: false, message: "No users found" };

  // Read User data columns B through H so we can verify username/password and capture role.
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 7);
  const data = dataRange.getValues();

  for (let i = 0; i < data.length; i++) {
    const userId = String(data[i][0]).trim();
    const fullName = String(data[i][2]).trim();
    const role = String(data[i][3]).trim() || "User";
    const accountStatus = String(data[i][4]).trim();
    const user = String(data[i][5]).trim();
    const storedPasswordHash = String(data[i][6]).trim();

    // Check if account is active
    if (accountStatus.toLowerCase() !== "active") {
      if (user === username) {
        return { success: false, message: "Account is inactive. Contact administrator." };
      }
      continue;
    }

    // Verify username and password using secure hash comparison
    if (user === username && verifyPassword(password, storedPasswordHash)) {
      // Create secure session with unique token
      const sessionId = createSession(userId, username, role, fullName);
      
      // Store session ID in user properties (client-accessible)
      storeSessionIdInProperties(sessionId);
      
      return { 
        success: true, 
        role: role,
        sessionId: sessionId,
        username: username,
        fullName: fullName
      };
    }
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
