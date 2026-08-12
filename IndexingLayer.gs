// --- INDEXING LAYER FOR FAST LOOKUPS ---

/**
 * Build index for fast student lookups by ID
 * Maps studentId → sheet row number
 * @returns {object} Index map { studentId: rowNumber }
 */
function buildStudentIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  
  if (!sheet) return {};
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};
  
  // Read only the Student ID column (Column B)
  const ids = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  
  const index = {};
  ids.forEach((row, i) => {
    const studentId = String(row[0]).trim();
    if (studentId) {
      index[studentId] = i + 3; // +3 for sheet row offset (Row 1 blank, Row 2 headers, data starts Row 3)
    }
  });
  
  // Cache for 1 hour
  try {
    const cache = CacheService.getScriptCache();
    cache.put('index_students', JSON.stringify(index), 3600);
  } catch (e) {
    Logger.log('Index caching error: ' + e.message);
  }
  
  Logger.log('Student index built: ' + Object.keys(index).length + ' entries');
  return index;
}

/**
 * Build index for fast user lookups by ID
 * @returns {object} Index map { userId: rowNumber }
 */
function buildUserIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  
  if (!sheet) return {};
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};
  
  const ids = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  
  const index = {};
  ids.forEach((row, i) => {
    const userId = String(row[0]).trim();
    if (userId) {
      index[userId] = i + 3;
    }
  });
  
  try {
    const cache = CacheService.getScriptCache();
    cache.put('index_users', JSON.stringify(index), 3600);
  } catch (e) {
    Logger.log('Index caching error: ' + e.message);
  }
  
  Logger.log('User index built: ' + Object.keys(index).length + ' entries');
  return index;
}

/**
 * Build index for invoices
 * @returns {object} Index map { invoiceId: rowNumber }
 */
function buildInvoiceIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Invoices");
  
  if (!sheet) return {};
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};
  
  const ids = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  
  const index = {};
  ids.forEach((row, i) => {
    const invoiceId = String(row[0]).trim();
    if (invoiceId) {
      index[invoiceId] = i + 3;
    }
  });
  
  try {
    const cache = CacheService.getScriptCache();
    cache.put('index_invoices', JSON.stringify(index), 3600);
  } catch (e) {
    Logger.log('Index caching error: ' + e.message);
  }
  
  return index;
}

/**
 * Build index for classes
 * @returns {object} Index map { classId: rowNumber }
 */
function buildClassIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Classes");
  
  if (!sheet) return {};
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};
  
  const ids = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  
  const index = {};
  ids.forEach((row, i) => {
    const classId = String(row[0]).trim();
    if (classId) {
      index[classId] = i + 3;
    }
  });
  
  try {
    const cache = CacheService.getScriptCache();
    cache.put('index_classes', JSON.stringify(index), 3600);
  } catch (e) {
    Logger.log('Index caching error: ' + e.message);
  }
  
  return index;
}

/**
 * Get or build index for a specific data type
 * @param {string} dataType - Type of data (Students, Users, Invoices, Classes)
 * @returns {object} Index map
 */
function getOrBuildIndex(dataType) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'index_' + dataType.toLowerCase();
  
  // Try to get from cache
  let indexStr = cache.get(cacheKey);
  
  if (indexStr) {
    try {
      return JSON.parse(indexStr);
    } catch (e) {
      Logger.log('Index parse error: ' + e.message);
      // Fall through to rebuild
    }
  }
  
  // Cache miss - build index
  switch(dataType.toLowerCase()) {
    case 'students':
      return buildStudentIndex();
    case 'users':
      return buildUserIndex();
    case 'invoices':
      return buildInvoiceIndex();
    case 'classes':
      return buildClassIndex();
    default:
      return {};
  }
}

/**
 * Fast row lookup using index - REPLACES LINEAR SEARCH
 * @param {string} sheetName - Name of sheet (Students, Users, etc.)
 * @param {string} id - ID to find
 * @returns {number} Row number or -1 if not found
 */
function findRowByIdIndexed(sheetName, id) {
  const index = getOrBuildIndex(sheetName);
  const rowNumber = index[id];
  
  if (rowNumber) {
    Logger.log('Index HIT: ' + sheetName + '[' + id + '] → Row ' + rowNumber);
    return rowNumber;
  }
  
  Logger.log('Index MISS: ' + sheetName + '[' + id + '] - rebuilding index');
  
  // Index miss - rebuild and try again
  const freshIndex = getOrBuildIndex(sheetName);
  return freshIndex[id] || -1;
}

/**
 * Invalidate index when data is modified
 * @param {string} dataType - Type of data (Students, Users, etc.)
 */
function invalidateIndex(dataType) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'index_' + dataType.toLowerCase();
  cache.remove(cacheKey);
  Logger.log('Index invalidated: ' + dataType);
}

/**
 * Invalidate multiple indexes at once
 * @param {array} dataTypes - Array of data types to invalidate
 */
function invalidateMultipleIndexes(dataTypes) {
  dataTypes.forEach(function(dataType) {
    invalidateIndex(dataType);
  });
}

/**
 * Rebuild all indexes (maintenance function)
 */
function rebuildAllIndexes() {
  Logger.log('Rebuilding all indexes...');
  
  const startTime = new Date().getTime();
  
  const indexes = {
    Students: buildStudentIndex(),
    Users: buildUserIndex(),
    Invoices: buildInvoiceIndex(),
    Classes: buildClassIndex()
  };
  
  const elapsed = new Date().getTime() - startTime;
  
  const summary = {
    elapsed: elapsed + 'ms',
    students: Object.keys(indexes.Students).length,
    users: Object.keys(indexes.Users).length,
    invoices: Object.keys(indexes.Invoices).length,
    classes: Object.keys(indexes.Classes).length
  };
  
  Logger.log('All indexes rebuilt: ' + JSON.stringify(summary));
  return summary;
}

/**
 * Clear all indexes
 */
function clearAllIndexes() {
  const cache = CacheService.getScriptCache();
  const indexKeys = [
    'index_students',
    'index_users',
    'index_invoices',
    'index_classes'
  ];
  
  cache.removeAll(indexKeys);
  Logger.log('All indexes cleared');
  
  return { success: true, cleared: indexKeys.length };
}

/**
 * Get index statistics (for monitoring)
 * @returns {object} Index stats
 */
function getIndexStats() {
  const cache = CacheService.getScriptCache();
  
  const stats = {};
  
  ['students', 'users', 'invoices', 'classes'].forEach(function(type) {
    const cacheKey = 'index_' + type;
    const indexStr = cache.get(cacheKey);
    
    if (indexStr) {
      try {
        const index = JSON.parse(indexStr);
        stats[type] = {
          cached: true,
          entries: Object.keys(index).length
        };
      } catch (e) {
        stats[type] = {
          cached: false,
          error: 'Parse error'
        };
      }
    } else {
      stats[type] = {
        cached: false,
        entries: 0
      };
    }
  });
  
  return stats;
}

/**
 * Benchmark index performance vs linear search
 * @returns {object} Performance comparison
 */
function benchmarkIndexPerformance() {
  Logger.log('Starting index benchmark...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  
  if (!sheet || sheet.getLastRow() < 3) {
    return { error: 'Not enough data to benchmark' };
  }
  
  // Get a random student ID
  const students = getStudentsDataUncached();
  if (students.length === 0) {
    return { error: 'No students found' };
  }
  
  const testStudentId = students[Math.floor(Math.random() * students.length)].studentId;
  
  // Clear indexes first
  clearAllIndexes();
  
  // Benchmark 1: Linear search (old method)
  const linearStart = new Date().getTime();
  const linearRow = findRowByIdLinear(sheet, testStudentId);
  const linearTime = new Date().getTime() - linearStart;
  
  // Benchmark 2: Indexed search (new method - first time, cache miss)
  const indexedStart1 = new Date().getTime();
  const indexedRow1 = findRowByIdIndexed('Students', testStudentId);
  const indexedTime1 = new Date().getTime() - indexedStart1;
  
  // Benchmark 3: Indexed search (second time, cache hit)
  const indexedStart2 = new Date().getTime();
  const indexedRow2 = findRowByIdIndexed('Students', testStudentId);
  const indexedTime2 = new Date().getTime() - indexedStart2;
  
  const improvement = Math.round((linearTime / indexedTime2) * 100) / 100;
  
  return {
    testId: testStudentId,
    linearSearch: linearTime + 'ms (row: ' + linearRow + ')',
    indexedSearchMiss: indexedTime1 + 'ms (row: ' + indexedRow1 + ')',
    indexedSearchHit: indexedTime2 + 'ms (row: ' + indexedRow2 + ')',
    speedup: improvement + 'x faster',
    recommendation: improvement > 5 ? 'Indexing highly effective' : 'Consider other optimizations'
  };
}

/**
 * Linear search (old method) - kept for benchmarking
 * @param {Sheet} sheet - Google Sheets sheet object
 * @param {string} id - ID to find
 * @returns {number} Row number or -1
 */
function findRowByIdLinear(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return -1;
  
  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(id).trim()) {
      return i + 3;
    }
  }
  
  return -1;
}

/**
 * Build composite index for faster multi-column searches
 * Example: Search students by class AND status
 * @param {string} sheetName - Sheet name
 * @param {array} columns - Column numbers to index (e.g., [2, 9] for ID and Class)
 * @param {string} indexName - Custom index name
 * @returns {object} Composite index
 */
function buildCompositeIndex(sheetName, columns, indexName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return {};
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};
  
  // Calculate range to read all indexed columns
  const minCol = Math.min.apply(null, columns);
  const maxCol = Math.max.apply(null, columns);
  const numCols = maxCol - minCol + 1;
  
  const data = sheet.getRange(3, minCol, lastRow - 2, numCols).getValues();
  
  const index = {};
  
  data.forEach((row, i) => {
    // Build composite key from all indexed columns
    const keyParts = columns.map(function(col) {
      const adjustedIndex = col - minCol;
      return String(row[adjustedIndex]).trim();
    });
    
    const compositeKey = keyParts.join('|');
    
    if (!index[compositeKey]) {
      index[compositeKey] = [];
    }
    
    index[compositeKey].push(i + 3); // Sheet row number
  });
  
  // Cache composite index
  try {
    const cache = CacheService.getScriptCache();
    cache.put('composite_' + indexName, JSON.stringify(index), 3600);
  } catch (e) {
    Logger.log('Composite index caching error: ' + e.message);
  }
  
  Logger.log('Composite index built: ' + indexName + ' (' + Object.keys(index).length + ' keys)');
  return index;
}

/**
 * Search using composite index
 * Example: Find all students in "Class 5" with status "Active"
 * @param {string} indexName - Composite index name
 * @param {array} keyParts - Values to search for
 * @returns {array} Array of row numbers
 */
function searchCompositeIndex(indexName, keyParts) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'composite_' + indexName;
  
  let indexStr = cache.get(cacheKey);
  
  if (!indexStr) {
    Logger.log('Composite index not found: ' + indexName);
    return [];
  }
  
  try {
    const index = JSON.parse(indexStr);
    const compositeKey = keyParts.join('|');
    return index[compositeKey] || [];
  } catch (e) {
    Logger.log('Composite index search error: ' + e.message);
    return [];
  }
}

/**
 * Warm all indexes (preload during off-peak hours)
 */
function warmAllIndexes() {
  Logger.log('Warming all indexes...');
  return rebuildAllIndexes();
}

/**
 * Smart index invalidation - only invalidate affected indexes
 * @param {string} operation - Operation type (add, update, delete)
 * @param {string} dataType - Data type (Students, Users, etc.)
 */
function smartInvalidateIndex(operation, dataType) {
  // Always invalidate primary index
  invalidateIndex(dataType);
  
  // Invalidate related indexes based on operation
  if (dataType === 'Students') {
    // If students change, dashboard stats might change
    invalidateCache('dashboard_stats');
  }
  
  if (dataType === 'Invoices' || dataType === 'Payments') {
    // Financial data affects dashboard
    invalidateCache('dashboard_stats');
    invalidateIndex('Invoices');
    invalidateIndex('Payments');
  }
}
