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
 * Get or initialize the salt from Script Properties
 * The salt is stored securely and used for all password hashing
 */
function getSalt() {
  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty('PASSWORD_SALT');
  
  if (!salt) {
    // Generate a unique salt for this installation
    salt = Utilities.getUuid();
    props.setProperty('PASSWORD_SALT', salt);
  }
  
  return salt;
}

/**
 * Hash a password using SHA-256
 * @param {string} password - The plain text password to hash
 * @returns {string} The hashed password as a hexadecimal string
 */
function hashPassword(password) {
  if (!password) return '';
  
  const salt = getSalt();
  const saltedPassword = password + salt;
  
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    saltedPassword,
    Utilities.Charset.UTF_8
  );
  
  // Convert byte array to hexadecimal string
  return rawHash.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * Verify a password against a stored hash
 * @param {string} inputPassword - The password to verify
 * @param {string} storedHash - The stored password hash
 * @returns {boolean} True if password matches
 */
function verifyPassword(inputPassword, storedHash) {
  if (!inputPassword || !storedHash) return false;
  return hashPassword(inputPassword) === storedHash;
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
 