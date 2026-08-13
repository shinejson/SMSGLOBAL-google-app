/**
 * AUDIT LOG SYSTEM - BACKEND
 * Separate backend script for audit trail functionality
 * Tracks all user actions in the application
 */

/**
 * Create or get the Audit Logs sheet
 */
function getAuditLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Audit Logs');
  
  if (!sheet) {
    sheet = ss.insertSheet('Audit Logs');
    
    // Set up headers
    const headers = [
      'Log ID',
      'Timestamp',
      'User ID',
      'User Name',
      'Action',
      'Module',
      'Record ID',
      'Details',
      'IP Address',
      'User Agent',
      'Old Value',
      'New Value'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set column widths
    sheet.setColumnWidth(1, 150); // Log ID
    sheet.setColumnWidth(2, 180); // Timestamp
    sheet.setColumnWidth(3, 120); // User ID
    sheet.setColumnWidth(4, 150); // User Name
    sheet.setColumnWidth(5, 100); // Action
    sheet.setColumnWidth(6, 120); // Module
    sheet.setColumnWidth(7, 150); // Record ID
    sheet.setColumnWidth(8, 250); // Details
    sheet.setColumnWidth(9, 130); // IP Address
    sheet.setColumnWidth(10, 200); // User Agent
    sheet.setColumnWidth(11, 200); // Old Value
    sheet.setColumnWidth(12, 200); // New Value
  }
  
  return sheet;
}

function auditToString(value) {
  if (value == null) return '';
  if (value instanceof Date) {
    try {
      return value.toISOString();
    } catch (e) {
      return String(value);
    }
  }
  if (typeof value === 'object') {
    try {
      return JSON.stringify(value);
    } catch (e) {
      return String(value);
    }
  }
  return String(value);
}

function safeLogAuditEvent(action, module, recordId, details, oldValue, newValue) {
  try {
    return logAuditEvent(action, module, recordId, details, oldValue, newValue);
  } catch (error) {
    Logger.log('Audit logging failed for ' + module + ' [' + action + ']: ' + error.message);
    return { success: false, message: error.message };
  }
}

function buildAuditSnapshot(fields) {
  const snapshot = {};
  Object.keys(fields || {}).forEach(function(key) {
    const value = fields[key];
    if (value !== undefined) {
      snapshot[key] = value;
    }
  });
  return snapshot;
}

/**
 * Log an audit event
 * @param {string} action - The action performed (Create, Update, Delete, View, Login, Logout, etc.)
 * @param {string} module - The module/entity affected (Students, Attendance, Users, etc.)
 * @param {string} recordId - The ID of the record affected
 * @param {string} details - Additional details about the action
 * @param {string} oldValue - Previous value (for updates)
 * @param {string} newValue - New value (for updates)
 */
function logAuditEvent(action, module, recordId, details, oldValue, newValue) {
  try {
    const sheet = getAuditLogSheet();
    const user = getLoggedInUser();
    const timestamp = new Date();
    
    // Generate unique log ID
    const logId = 'LOG-' + timestamp.getTime() + '-' + Math.floor(Math.random() * 1000);
    
    // Get user information
    const userId = user ? user.userId : 'Unknown';
    const userName = user ? user.fullName : 'Unknown User';
    
    // Get IP address (approximation - not always accurate in Apps Script)
    const ipAddress = Session.getTemporaryActiveUserKey() || 'N/A';
    
    // Get user agent
    const userAgent = 'Google Apps Script'; // Limited in Apps Script environment
    
    // Prepare row data
    const rowData = [
      logId,
      timestamp,
      userId,
      userName,
      action,
      module,
      recordId || '',
      details || '',
      ipAddress,
      userAgent,
      auditToString(oldValue),
      auditToString(newValue)
    ];
    
    // Append to sheet
    sheet.appendRow(rowData);
    
    return { success: true, logId: logId };
  } catch (error) {
    Logger.log('Error logging audit event: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * Get all audit logs with optional filters
 */
function getAuditLogs(filters) {
  try {
    const sheet = getAuditLogSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return [];
    }
    
    // Get all data
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    
    // Convert to objects
    const logs = data.map(function(row) {
      let ts = '';
      if (row[1] instanceof Date) {
        ts = row[1].toISOString();
      } else if (row[1]) {
        ts = String(row[1]);
      }

      return {
        logId: String(row[0] || ''),
        timestamp: ts,
        userId: String(row[2] || ''),
        userName: String(row[3] || ''),
        action: String(row[4] || ''),
        module: String(row[5] || ''),
        recordId: String(row[6] || ''),
        details: String(row[7] || ''),
        ipAddress: String(row[8] || ''),
        userAgent: String(row[9] || ''),
        oldValue: String(row[10] || ''),
        newValue: String(row[11] || '')
      };
    }).filter(log => log.logId); // Filter out empty rows
    
    // Sort by timestamp descending (newest first)
    logs.sort(function(a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });
    
    return logs;
  } catch (error) {
    Logger.log('Error getting audit logs: ' + error.message);
    throw new Error('Failed to retrieve audit logs: ' + error.message);
  }
}

/**
 * Clear old audit logs (older than specified days)
 */
function clearOldAuditLogs(days) {
  try {
    const sheet = getAuditLogSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: true, message: 'No logs to delete', deletedCount: 0 };
    }
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    
    // Get timestamp column (column 2)
    const timestamps = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    
    // Find rows to delete (from bottom to top to avoid index shifting)
    const rowsToDelete = [];
    for (let i = timestamps.length - 1; i >= 0; i--) {
      const timestamp = new Date(timestamps[i][0]);
      if (timestamp < cutoffDate) {
        rowsToDelete.push(i + 2); // +2 because array is 0-indexed and sheet has header
      }
    }
    
    // Delete rows
    rowsToDelete.forEach(function(rowIndex) {
      sheet.deleteRow(rowIndex);
    });
    
    const deletedCount = rowsToDelete.length;
    
    return {
      success: true,
      message: `Deleted ${deletedCount} audit log(s) older than ${days} days`,
      deletedCount: deletedCount
    };
  } catch (error) {
    Logger.log('Error clearing old audit logs: ' + error.message);
    return { success: false, message: 'Failed to clear logs: ' + error.message };
  }
}

/**
 * Get audit logs for a specific record
 */
function getAuditLogsForRecord(module, recordId) {
  try {
    const allLogs = getAuditLogs();
    return allLogs.filter(function(log) {
      return log.module === module && log.recordId === recordId;
    });
  } catch (error) {
    Logger.log('Error getting audit logs for record: ' + error.message);
    return [];
  }
}

/**
 * Get audit logs for a specific user
 */
function getAuditLogsForUser(userId) {
  try {
    const allLogs = getAuditLogs();
    return allLogs.filter(function(log) {
      return log.userId === userId;
    });
  } catch (error) {
    Logger.log('Error getting audit logs for user: ' + error.message);
    return [];
  }
}

/**
 * Export audit logs to CSV format
 */
function exportAuditLogsToCSV(filters) {
  try {
    const logs = getAuditLogs(filters);
    
    if (logs.length === 0) {
      return { success: false, message: 'No logs to export' };
    }
    
    // Create CSV content
    const headers = ['Log ID', 'Timestamp', 'User ID', 'User Name', 'Action', 'Module', 'Record ID', 'Details', 'IP Address', 'User Agent'];
    let csv = headers.join(',') + '\n';
    
    logs.forEach(function(log) {
      const row = [
        escapeCSV(log.logId),
        escapeCSV(log.timestamp),
        escapeCSV(log.userId),
        escapeCSV(log.userName),
        escapeCSV(log.action),
        escapeCSV(log.module),
        escapeCSV(log.recordId),
        escapeCSV(log.details),
        escapeCSV(log.ipAddress),
        escapeCSV(log.userAgent)
      ];
      csv += row.join(',') + '\n';
    });
    
    return { success: true, csv: csv };
  } catch (error) {
    Logger.log('Error exporting audit logs: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * Helper function to escape CSV values
 */
function escapeCSV(value) {
  if (value == null || value === '') return '""';
  const str = String(value);
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return '"' + str + '"';
}

/**
 * Get audit statistics
 */
function getAuditStatistics() {
  try {
    const logs = getAuditLogs();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const stats = {
      totalLogs: logs.length,
      todayLogs: logs.filter(function(log) {
        const logDate = new Date(log.timestamp);
        logDate.setHours(0, 0, 0, 0);
        return logDate.getTime() === today.getTime();
      }).length,
      uniqueUsers: new Set(logs.map(log => log.userId)).size,
      actionBreakdown: {},
      moduleBreakdown: {}
    };
    
    // Count actions
    logs.forEach(function(log) {
      stats.actionBreakdown[log.action] = (stats.actionBreakdown[log.action] || 0) + 1;
      stats.moduleBreakdown[log.module] = (stats.moduleBreakdown[log.module] || 0) + 1;
    });
    
    return stats;
  } catch (error) {
    Logger.log('Error getting audit statistics: ' + error.message);
    return null;
  }
}

// ==================== INTEGRATION HELPERS ====================
// Use these functions in your existing code to log events

/**
 * Log student actions
 */
function logStudentAction(action, studentId, details, oldValue, newValue) {
  return logAuditEvent(action, 'Students', studentId, details, oldValue, newValue);
}

/**
 * Log attendance actions
 */
function logAttendanceAction(action, attendanceId, details, oldValue, newValue) {
  return logAuditEvent(action, 'Attendance', attendanceId, details, oldValue, newValue);
}

/**
 * Log user actions
 */
function logUserAction(action, userId, details, oldValue, newValue) {
  return logAuditEvent(action, 'Users', userId, details, oldValue, newValue);
}

/**
 * Log invoice actions
 */
function logInvoiceAction(action, invoiceId, details, oldValue, newValue) {
  return logAuditEvent(action, 'Invoices', invoiceId, details, oldValue, newValue);
}

/**
 * Log payment actions
 */
function logPaymentAction(action, paymentId, details, oldValue, newValue) {
  return logAuditEvent(action, 'Payments', paymentId, details, oldValue, newValue);
}

/**
 * Log class actions
 */
function logClassAction(action, classId, details, oldValue, newValue) {
  return logAuditEvent(action, 'Classes', classId, details, oldValue, newValue);
}

/**
 * Log course actions
 */
function logCourseAction(action, courseId, details, oldValue, newValue) {
  return logAuditEvent(action, 'Courses', courseId, details, oldValue, newValue);
}

/**
 * Log performance actions
 */
function logPerformanceAction(action, performanceId, details, oldValue, newValue) {
  return logAuditEvent(action, 'Performance', performanceId, details, oldValue, newValue);
}

/**
 * Log enrollment actions
 */
function logEnrollmentAction(action, enrollmentId, details, oldValue, newValue) {
  return logAuditEvent(action, 'Enrollments', enrollmentId, details, oldValue, newValue);
}

/**
 * Log teacher actions
 */
function logTeacherAction(action, teacherId, details, oldValue, newValue) {
  return logAuditEvent(action, 'Teachers', teacherId, details, oldValue, newValue);
}

/**
 * Log academic year actions
 */
function logAcademicYearAction(action, academicYear, details, oldValue, newValue) {
  return logAuditEvent(action, 'Academic Years', academicYear, details, oldValue, newValue);
}

/**
 * Log authentication events
 */
function logAuthEvent(action, userId, details) {
  return logAuditEvent(action, 'Authentication', userId, details, null, null);
}

/**
 * Log settings changes
 */
function logSettingsChange(action, settingName, details, oldValue, newValue) {
  return logAuditEvent(action, 'Settings', settingName, details, oldValue, newValue);
}

/**
 * Log export events
 */
function logExportEvent(module, details) {
  return logAuditEvent('Export', module, null, details, null, null);
}

/**
 * Log import events
 */
function logImportEvent(module, details) {
  return logAuditEvent('Import', module, null, details, null, null);
}
