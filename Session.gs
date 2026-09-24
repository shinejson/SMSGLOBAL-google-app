// --- SECURE SESSION MANAGEMENT ---

/**
 * Create a new session with a unique token
 * @param {string} userId - User ID
 * @param {string} username - Username
 * @param {string} role - User role
 * @param {string} fullName - User's full name
 * @returns {string} Session token (UUID)
 */
function createSession(userId, username, role, fullName) {
  // Generate unique session ID (token)
  const sessionId = Utilities.getUuid();
  
  // Prepare session data
  const sessionData = {
    userId: userId,
    username: username,
    role: role,
    fullName: fullName || username,
    createdAt: new Date().getTime(),
    lastActivity: new Date().getTime()
  };
  
  // Store in user-specific cache (isolated per user)
  const cache = CacheService.getUserCache();
  const SESSION_TIMEOUT = 28800; // 8 hours in seconds
  
  cache.put('session_' + sessionId, JSON.stringify(sessionData), SESSION_TIMEOUT);
  
  // Also store in ScriptCache as backup (for recovery scenarios)
  try {
    CacheService.getScriptCache().put('session_' + sessionId, JSON.stringify(sessionData), SESSION_TIMEOUT);
  } catch (e) {
    Logger.log('Warning: Could not store backup session: ' + e.message);
  }
  
  Logger.log('Session created: ' + sessionId + ' for user: ' + username);
  return sessionId;
}

/**
 * Validate and retrieve session data
 * @param {string} sessionId - Session token to validate
 * @returns {object|null} Session data or null if invalid/expired
 */
function validateSession(sessionId) {
  if (!sessionId) return null;
  
  const cache = CacheService.getUserCache();
  const sessionKey = 'session_' + sessionId;
  
  // Try to get session from user cache
  let sessionDataStr = cache.get(sessionKey);
  
  // Fallback to script cache if user cache is empty
  if (!sessionDataStr) {
    try {
      sessionDataStr = CacheService.getScriptCache().get(sessionKey);
      // Restore to user cache if found in script cache
      if (sessionDataStr) {
        cache.put(sessionKey, sessionDataStr, 28800);
      }
    } catch (e) {
      Logger.log('Cache retrieval error: ' + e.message);
    }
  }
  
  if (!sessionDataStr) {
    Logger.log('Session not found: ' + sessionId);
    return null;
  }
  
  try {
    const session = JSON.parse(sessionDataStr);
    
    // Check session age (8 hours = 28800000 ms)
    const age = new Date().getTime() - session.createdAt;
    const SESSION_TIMEOUT_MS = 28800000; // 8 hours
    
    if (age > SESSION_TIMEOUT_MS) {
      Logger.log('Session expired: ' + sessionId);
      destroySession(sessionId);
      return null;
    }
    
    // Check inactivity timeout (30 minutes = 1800000 ms)
    const inactivity = new Date().getTime() - session.lastActivity;
    const INACTIVITY_TIMEOUT_MS = 1800000; // 30 minutes
    
    if (inactivity > INACTIVITY_TIMEOUT_MS) {
      Logger.log('Session timed out due to inactivity: ' + sessionId);
      destroySession(sessionId);
      return null;
    }
    
    // Update last activity timestamp
    session.lastActivity = new Date().getTime();
    cache.put(sessionKey, JSON.stringify(session), 28800);
    
    return session;
  } catch (e) {
    Logger.log('Session parse error: ' + e.message);
    return null;
  }
}

/**
 * Destroy a session (logout)
 * @param {string} sessionId - Session token to destroy
 */
function destroySession(sessionId) {
  if (!sessionId) return;
  
  const sessionKey = 'session_' + sessionId;
  
  // Remove from user cache
  try {
    CacheService.getUserCache().remove(sessionKey);
  } catch (e) {
    Logger.log('Error removing from user cache: ' + e.message);
  }
  
  // Remove from script cache
  try {
    CacheService.getScriptCache().remove(sessionKey);
  } catch (e) {
    Logger.log('Error removing from script cache: ' + e.message);
  }
  
  Logger.log('Session destroyed: ' + sessionId);
}

/**
 * Get session ID from request (from cookie or parameter)
 * @param {object} e - Event object from doGet/doPost
 * @returns {string|null} Session ID or null
 */
function getSessionIdFromRequest(e) {
  // Try to get from cookie first
  if (e && e.parameter && e.parameter.sessionId) {
    return e.parameter.sessionId;
  }
  
  // Try to get from properties as fallback (for backward compatibility)
  const props = PropertiesService.getUserProperties();
  return props.getProperty('CURRENT_SESSION_ID');
}

/**
 * Store session ID in user properties (client-side storage alternative)
 * @param {string} sessionId - Session token
 */
function storeSessionIdInProperties(sessionId) {
  const props = PropertiesService.getUserProperties();
  if (sessionId) {
    props.setProperty('CURRENT_SESSION_ID', sessionId);
  } else {
    props.deleteProperty('CURRENT_SESSION_ID');
  }
}

/**
 * Clean up old/expired sessions (maintenance function)
 * Note: CacheService automatically removes expired entries,
 * but this can be used to force cleanup
 */
function cleanupExpiredSessions() {
  // CacheService handles expiration automatically
  // This function is kept for manual cleanup if needed
  Logger.log('Session cleanup completed (automatic via CacheService)');
  return { success: true, message: 'CacheService handles automatic cleanup' };
}

/**
 * Get current session info (for debugging/admin purposes)
 * @returns {object} Session information
 */
function getCurrentSessionInfo() {
  const props = PropertiesService.getUserProperties();
  const sessionId = props.getProperty('CURRENT_SESSION_ID');
  
  if (!sessionId) {
    return { loggedIn: false, message: 'No active session' };
  }
  
  const session = validateSession(sessionId);
  
  if (!session) {
    return { loggedIn: false, message: 'Session expired or invalid' };
  }
  
  const now = new Date().getTime();
  const sessionAge = Math.floor((now - session.createdAt) / 1000 / 60); // minutes
  const inactiveTime = Math.floor((now - session.lastActivity) / 1000 / 60); // minutes
  
  return {
    loggedIn: true,
    sessionId: sessionId,
    username: session.username,
    role: session.role,
    sessionAgeMinutes: sessionAge,
    inactiveMinutes: inactiveTime,
    expiresIn: Math.floor((28800000 - (now - session.createdAt)) / 1000 / 60) + ' minutes'
  };
}

/**
 * Extend session timeout (refresh activity)
 * @param {string} sessionId - Session token
 * @returns {boolean} Success status
 */
function extendSession(sessionId) {
  const session = validateSession(sessionId);
  if (!session) return false;
  
  // validateSession already updates lastActivity
  return true;
}

/**
 * List all active sessions for current user (debugging)
 * Note: CacheService doesn't provide enumeration, so we track in properties
 * @returns {array} List of session info
 */
function listUserSessions() {
  const props = PropertiesService.getUserProperties();
  const sessionId = props.getProperty('CURRENT_SESSION_ID');
  
  if (!sessionId) {
    return [];
  }
  
  const session = validateSession(sessionId);
  if (!session) {
    return [];
  }
  
  return [getCurrentSessionInfo()];
}

/**
 * Migrate from old ScriptProperties session to new secure session
 * This allows seamless transition without logging everyone out
 */
function migrateOldSessionToSecure() {
  const oldProps = PropertiesService.getScriptProperties();
  const isLoggedIn = oldProps.getProperty('IS_LOGGED_IN');
  
  if (isLoggedIn !== 'true') {
    return { migrated: false, message: 'No old session found' };
  }
  
  const username = oldProps.getProperty('LOGGED_IN_USER');
  const role = oldProps.getProperty('LOGGED_IN_ROLE');
  
  if (!username || !role) {
    return { migrated: false, message: 'Incomplete session data' };
  }
  
  // Look up full name from Users sheet if available
  let fullName = username;
  try {
    const users = typeof getUsersData === 'function' ? getUsersData() : [];
    const matched = users.find(u => String(u.username).trim().toLowerCase() === String(username).trim().toLowerCase());
    if (matched && matched.fullName) fullName = matched.fullName;
  } catch (e) {}

  // Create new secure session
  const sessionId = createSession('MIGRATED', username, role, fullName);
  storeSessionIdInProperties(sessionId);
  
  // Clear old session data
  oldProps.deleteProperty('IS_LOGGED_IN');
  oldProps.deleteProperty('LOGGED_IN_USER');
  oldProps.deleteProperty('LOGGED_IN_ROLE');
  oldProps.deleteProperty('LOGGED_IN_AT');
  
  return {
    migrated: true,
    sessionId: sessionId,
    username: username,
    message: 'Successfully migrated to secure session'
  };
}
