// --- BACKUP AND RESTORE SYSTEM ---

/**
 * Initialize and authorize Drive access
 * Run this function ONCE to grant permissions
 * This must be run from the Apps Script Editor before using backup features
 */
function authorizeDriveAccess() {
  try {
    // Test Drive access
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const file = DriveApp.getFileById(ss.getId());
    const folder = getOrCreateBackupFolder();
    
    Logger.log('✅ Drive access authorized successfully!');
    Logger.log('Spreadsheet: ' + file.getName());
    Logger.log('Backup folder: ' + folder.getName());
    
    return {
      success: true,
      message: 'Drive access authorized. You can now use backup features.',
      spreadsheet: file.getName(),
      backupFolder: folder.getName()
    };
  } catch (error) {
    Logger.log('❌ Authorization failed: ' + error.message);
    return {
      success: false,
      message: 'Authorization failed. Please run this function from Apps Script Editor.',
      error: error.message
    };
  }
}

/**
 * Create a complete backup of all sheets
 * @returns {object} Backup result with file ID and download URL
 */
function createFullBackup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const backupName = 'SMS_Backup_' + timestamp;
    
    // Create a copy of the entire spreadsheet
    const backupFile = DriveApp.getFileById(ss.getId()).makeCopy(backupName);
    
    // Move to backups folder (create if doesn't exist)
    const backupFolder = getOrCreateBackupFolder();
    backupFile.moveTo(backupFolder);
    
    // Set sharing to anyone with link (view only)
    backupFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const downloadUrl = 'https://docs.google.com/spreadsheets/d/' + backupFile.getId() + '/export?format=xlsx';
    
    Logger.log('Backup created: ' + backupName);
    
    return {
      success: true,
      fileId: backupFile.getId(),
      fileName: backupName,
      downloadUrl: downloadUrl,
      viewUrl: backupFile.getUrl(),
      folderId: backupFolder.getId(),
      folderUrl: backupFolder.getUrl(),
      timestamp: timestamp,
      size: formatFileSize(backupFile.getSize()),
      message: 'Backup created successfully'
    };
    
  } catch (error) {
    Logger.log('Backup error: ' + error.message);
    return {
      success: false,
      message: 'Backup failed: ' + error.message
    };
  }
}

/**
 * Create backup of specific sheets only
 * @param {array} sheetNames - Array of sheet names to backup
 * @returns {object} Backup result
 */
function createSelectiveBackup(sheetNames) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const backupName = 'SMS_Selective_Backup_' + timestamp;
    
    // Create new spreadsheet for selective backup
    const newSS = SpreadsheetApp.create(backupName);
    const newSSId = newSS.getId();
    
    // Copy selected sheets
    sheetNames.forEach(function(sheetName) {
      const sourceSheet = ss.getSheetByName(sheetName);
      if (sourceSheet) {
        sourceSheet.copyTo(newSS);
      }
    });
    
    // Remove the default "Sheet1" if it exists
    const defaultSheet = newSS.getSheetByName('Sheet1');
    if (defaultSheet && newSS.getSheets().length > 1) {
      newSS.deleteSheet(defaultSheet);
    }
    
    // Move to backups folder
    const backupFolder = getOrCreateBackupFolder();
    const backupFile = DriveApp.getFileById(newSSId);
    backupFile.moveTo(backupFolder);
    
    backupFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const downloadUrl = 'https://docs.google.com/spreadsheets/d/' + newSSId + '/export?format=xlsx';
    
    return {
      success: true,
      fileId: newSSId,
      fileName: backupName,
      downloadUrl: downloadUrl,
      viewUrl: backupFile.getUrl(),
      folderId: backupFolder.getId(),
      folderUrl: backupFolder.getUrl(),
      sheetsBackedUp: sheetNames,
      message: 'Selective backup created successfully'
    };
    
  } catch (error) {
    Logger.log('Selective backup error: ' + error.message);
    return {
      success: false,
      message: 'Selective backup failed: ' + error.message
    };
  }
}

/**
 * Get or create the backups folder in Drive
 * @returns {Folder} Backup folder
 */
function getOrCreateBackupFolder() {
  const folderName = 'SMS_Backups';
  
  // Check if folder exists
  const folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  }
  
  // Create folder if it doesn't exist
  return DriveApp.createFolder(folderName);
}

/**
 * List all available backups
 * @param {number} limit - Maximum number of backups to return
 * @returns {array} Array of backup info objects
 */
function listBackups(limit) {
  try {
    limit = limit || 20;
    
    const backupFolder = getOrCreateBackupFolder();
    const files = backupFolder.getFiles();
    
    const backups = [];
    
    while (files.hasNext() && backups.length < limit) {
      const file = files.next();
      const fileName = file.getName();
      
      // Only include files that start with SMS_Backup
      if (fileName.indexOf('SMS_Backup') === 0 || fileName.indexOf('SMS_Selective_Backup') === 0) {
        backups.push({
          fileId: file.getId(),
          fileName: fileName,
          created: file.getDateCreated(),
          modified: file.getLastUpdated(),
          size: formatFileSize(file.getSize()),
          downloadUrl: 'https://docs.google.com/spreadsheets/d/' + file.getId() + '/export?format=xlsx',
          viewUrl: file.getUrl()
        });
      }
    }
    
    // Sort by creation date (newest first)
    backups.sort(function(a, b) {
      return b.created.getTime() - a.created.getTime();
    });
    
    return backups;
    
  } catch (error) {
    Logger.log('List backups error: ' + error.message);
    return [];
  }
}

/**
 * Delete a backup file
 * @param {string} fileId - File ID to delete
 * @returns {object} Delete result
 */
function deleteBackup(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    
    // Safety check - only delete backup files
    if (fileName.indexOf('SMS_Backup') !== 0 && fileName.indexOf('SMS_Selective_Backup') !== 0) {
      throw new Error('Not a backup file');
    }
    
    file.setTrashed(true);
    
    Logger.log('Backup deleted: ' + fileName);
    
    return {
      success: true,
      message: 'Backup deleted: ' + fileName
    };
    
  } catch (error) {
    Logger.log('Delete backup error: ' + error.message);
    return {
      success: false,
      message: 'Delete failed: ' + error.message
    };
  }
}

/**
 * Restore from a backup file
 * WARNING: This will overwrite current data!
 * @param {string} backupFileId - Backup file ID to restore from
 * @param {array} sheetNames - Optional: specific sheets to restore
 * @returns {object} Restore result
 */
function restoreFromBackup(backupFileId, sheetNames) {
  try {
    const currentSS = SpreadsheetApp.getActiveSpreadsheet();
    const backupSS = SpreadsheetApp.openById(backupFileId);
    
    // If no specific sheets specified, restore all
    if (!sheetNames || sheetNames.length === 0) {
      sheetNames = backupSS.getSheets().map(function(sheet) {
        return sheet.getName();
      });
    }
    
    const restored = [];
    
    sheetNames.forEach(function(sheetName) {
      const backupSheet = backupSS.getSheetByName(sheetName);
      
      if (!backupSheet) {
        Logger.log('Sheet not found in backup: ' + sheetName);
        return;
      }
      
      // Check if sheet exists in current spreadsheet
      let currentSheet = currentSS.getSheetByName(sheetName);
      
      if (currentSheet) {
        // Clear existing data
        currentSheet.clear();
      } else {
        // Create new sheet
        currentSheet = currentSS.insertSheet(sheetName);
      }
      
      // Copy data from backup
      const backupData = backupSheet.getDataRange();
      const numRows = backupData.getNumRows();
      const numCols = backupData.getNumColumns();
      
      if (numRows > 0 && numCols > 0) {
        const values = backupData.getValues();
        currentSheet.getRange(1, 1, numRows, numCols).setValues(values);
        
        // Copy formatting
        backupData.copyFormatToRange(currentSheet, 1, numCols, 1, numRows);
        
        restored.push(sheetName);
      }
    });
    
    // Invalidate all caches and indexes after restore
    clearAllCache();
    clearAllIndexes();
    
    Logger.log('Restore completed: ' + restored.length + ' sheets');
    
    return {
      success: true,
      restoredSheets: restored,
      message: 'Restored ' + restored.length + ' sheet(s) successfully'
    };
    
  } catch (error) {
    Logger.log('Restore error: ' + error.message);
    return {
      success: false,
      message: 'Restore failed: ' + error.message
    };
  }
}

/**
 * Schedule automatic backups (run daily)
 * Call this once to set up the trigger
 */
function setupAutomaticBackups() {
  try {
    // Delete existing triggers for this function
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'performAutomaticBackup') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create new daily trigger at 2 AM
    ScriptApp.newTrigger('performAutomaticBackup')
      .timeBased()
      .atHour(2)
      .everyDays(1)
      .create();
    
    Logger.log('Automatic backup trigger created');
    
    return {
      success: true,
      message: 'Automatic daily backups enabled (2 AM)'
    };
    
  } catch (error) {
    Logger.log('Setup automatic backups error: ' + error.message);
    return {
      success: false,
      message: 'Failed to setup automatic backups: ' + error.message
    };
  }
}

/**
 * Disable automatic backups
 */
function disableAutomaticBackups() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deleted = 0;
    
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'performAutomaticBackup') {
        ScriptApp.deleteTrigger(trigger);
        deleted++;
      }
    });
    
    Logger.log('Automatic backup triggers deleted: ' + deleted);
    
    return {
      success: true,
      message: 'Automatic backups disabled'
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Failed to disable automatic backups: ' + error.message
    };
  }
}

/**
 * Perform automatic backup (called by trigger)
 */
function performAutomaticBackup() {
  try {
    const result = createFullBackup();
    
    if (result.success) {
      // Clean up old backups (keep last 30 days)
      cleanupOldBackups(30);
    }
    
    return result;
    
  } catch (error) {
    Logger.log('Automatic backup error: ' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Clean up old backups
 * @param {number} daysToKeep - Number of days to keep backups
 * @returns {object} Cleanup result
 */
function cleanupOldBackups(daysToKeep) {
  try {
    daysToKeep = daysToKeep || 30;
    
    const backupFolder = getOrCreateBackupFolder();
    const files = backupFolder.getFiles();
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    
    let deleted = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      if ((fileName.indexOf('SMS_Backup') === 0 || fileName.indexOf('SMS_Selective_Backup') === 0) &&
          file.getDateCreated() < cutoffDate) {
        file.setTrashed(true);
        deleted++;
        Logger.log('Deleted old backup: ' + fileName);
      }
    }
    
    return {
      success: true,
      deleted: deleted,
      message: 'Deleted ' + deleted + ' old backup(s)'
    };
    
  } catch (error) {
    Logger.log('Cleanup error: ' + error.message);
    return {
      success: false,
      message: 'Cleanup failed: ' + error.message
    };
  }
}

/**
 * Export specific sheet to CSV
 * @param {string} sheetName - Name of sheet to export
 * @returns {object} Export result with download URL
 */
function exportSheetToCSV(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('Sheet not found: ' + sheetName);
    }
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const fileName = sheetName + '_Export_' + timestamp + '.csv';
    
    // Get sheet data
    const data = sheet.getDataRange().getValues();
    
    // Convert to CSV
    const csv = data.map(function(row) {
      return row.map(function(cell) {
        // Escape quotes and wrap in quotes if contains comma
        const cellStr = String(cell);
        if (cellStr.indexOf(',') !== -1 || cellStr.indexOf('"') !== -1) {
          return '"' + cellStr.replace(/"/g, '""') + '"';
        }
        return cellStr;
      }).join(',');
    }).join('\n');
    
    // Create file in backups folder
    const backupFolder = getOrCreateBackupFolder();
    const csvFile = backupFolder.createFile(fileName, csv, MimeType.CSV);
    
    csvFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      success: true,
      fileId: csvFile.getId(),
      fileName: fileName,
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + csvFile.getId(),
      viewUrl: csvFile.getUrl(),
      folderId: backupFolder.getId(),
      folderUrl: backupFolder.getUrl(),
      message: 'CSV export successful'
    };
    
  } catch (error) {
    Logger.log('CSV export error: ' + error.message);
    return {
      success: false,
      message: 'CSV export failed: ' + error.message
    };
  }
}

/**
 * Get backup statistics
 * @returns {object} Backup statistics
 */
function getBackupStats() {
  try {
    const backupFolder = getOrCreateBackupFolder();
    const files = backupFolder.getFiles();
    
    let totalBackups = 0;
    let totalSize = 0;
    let oldestBackup = null;
    let newestBackup = null;
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      if (fileName.indexOf('SMS_Backup') === 0 || fileName.indexOf('SMS_Selective_Backup') === 0) {
        totalBackups++;
        totalSize += file.getSize();
        
        const created = file.getDateCreated();
        
        if (!oldestBackup || created < oldestBackup) {
          oldestBackup = created;
        }
        
        if (!newestBackup || created > newestBackup) {
          newestBackup = created;
        }
      }
    }
    
    // Check if automatic backups are enabled
    const triggers = ScriptApp.getProjectTriggers();
    let autoBackupEnabled = false;
    
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'performAutomaticBackup') {
        autoBackupEnabled = true;
      }
    });
    
    return {
      totalBackups: totalBackups,
      totalSize: formatFileSize(totalSize),
      oldestBackup: oldestBackup ? Utilities.formatDate(oldestBackup, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : 'N/A',
      newestBackup: newestBackup ? Utilities.formatDate(newestBackup, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : 'N/A',
      autoBackupEnabled: autoBackupEnabled,
      backupFolder: backupFolder.getName(),
      backupFolderId: backupFolder.getId()
    };
    
  } catch (error) {
    Logger.log('Backup stats error: ' + error.message);
    return {
      error: error.message
    };
  }
}

/**
 * Format file size in human-readable format
 * @param {number} bytes - Size in bytes
 * @returns {string} Formatted size
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  
  return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

/**
 * Get list of all sheet names for backup selection
 * @returns {array} Array of sheet names
 */
function getAvailableSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().map(function(sheet) {
    return {
      name: sheet.getName(),
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn()
    };
  });
}
