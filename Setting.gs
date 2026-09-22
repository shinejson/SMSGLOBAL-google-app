// --- SETTINGS FUNCTIONS ---

// 1. Fetch Main System Parameters (Left side)
function getSystemParameters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return [];

  // Read from Row 5 downward, Column B and C
  const dataRange = sheet.getRange(5, 2, lastRow - 4, 2);
  const data = dataRange.getValues();

  return data
    .map((row) => ({
      param: String(row[0]).trim(),
      value: /password/i.test(String(row[0]).trim()) ? '••••••' : String(row[1]).trim()
    }))
    .filter(item => item.param !== '');
}

// Returns a flat { paramName: value } map for quick key-based lookups.
// Passwords are masked (shown as '••••••') in the returned map.
function getParameters() {
  const params = getSystemParameters();
  const map = {};
  params.forEach(function(item) { map[item.param] = item.value; });
  try {
    if (typeof isTestModeActive === 'function') {
      const tm = isTestModeActive();
      if (tm && tm.active) {
        map['License Status'] = 'Test Mode';
        map['Student Limit'] = 999999;
        map['isTestMode'] = true;
        map['testModeDeadline'] = tm.deadline;
      }
    }
  } catch (e) {
    Logger.log('Error checking test mode in getParameters: ' + e.message);
  }
  return map;
}


// 2. Update a specific System Parameter
function normalizeParamName(name) {
  return String(name || '').trim().toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-9]/g, '');
}

function hashString(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8);
  return bytes.map(function(byte) {
    return (byte + 256).toString(16).slice(-2);
  }).join('');
}

function isStrongPassword(value) {
  return /^(?=.*[A-Za-z])(?=.*\d)(?=.*[^A-Za-z\d]).{12,}$/.test(String(value || ''));
}

function isSensitiveSettingName(paramName) {
  return /password|secret|token|key/i.test(String(paramName || '').trim());
}

function maskSettingAuditValue(paramName, value) {
  if (value == null || value === '') return '';
  return isSensitiveSettingName(paramName) ? '••••••' : value;
}

function buildSettingAuditSnapshot(fields) {
  if (!fields) return null;

  return buildAuditSnapshot({
    param: fields.param,
    value: maskSettingAuditValue(fields.param, fields.value),
    listType: fields.listType,
    columnLetter: fields.columnLetter,
    oldValue: fields.oldValue,
    newValue: fields.newValue
  });
}

function getSettingParameterSnapshot(paramName) {
  const params = getSystemParameters();
  const normalizedTarget = normalizeParamName(paramName);
  const match = params.find(function(item) {
    return normalizeParamName(item.param) === normalizedTarget;
  });

  return match ? buildSettingAuditSnapshot({ param: match.param, value: match.value }) : null;
}

function getSettingsListType(columnLetter) {
  const normalized = String(columnLetter || '').trim().toUpperCase();
  if (normalized === 'E') return 'Categories';
  if (normalized === 'F') return 'Payment Methods';
  if (normalized === 'G') return 'Statuses';
  return 'Settings List';
}

function updateSystemParameter(paramName, newValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) throw new Error("Settings worksheet not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 5) throw new Error("No system parameters found.");

  const beforeSnapshot = getSettingParameterSnapshot(paramName);
  const normalizedTarget = normalizeParamName(paramName);
  const paramsRange = sheet.getRange(5, 2, lastRow - 4, 1);
  const params = paramsRange.getValues();

  let foundRow = -1;
  for (let i = 0; i < params.length; i++) {
    if (normalizeParamName(params[i][0]) === normalizedTarget) {
      foundRow = 5 + i;
      break;
    }
  }

  let storedValue = String(newValue || '').trim();
  if (normalizedTarget === 'masterpassword' || normalizedTarget === 'masterpass') {
    if (!isStrongPassword(storedValue)) {
      throw new Error('Master password must be at least 12 characters and include letters, numbers, and special characters.');
    }
    if (!storedValue.startsWith('sha256:')) {
      storedValue = 'sha256:' + hashString(storedValue);
    }
  }

  if (foundRow === -1) {
    sheet.getRange(lastRow + 1, 2, 1, 2).setValues([[paramName, storedValue]]);

    const createdSnapshot = buildSettingAuditSnapshot({
      param: paramName,
      value: storedValue
    });

    safeLogAuditEvent(
      'Create',
      'Settings',
      String(paramName || '').trim(),
      'Created system parameter ' + String(paramName || '').trim(),
      null,
      createdSnapshot
    );

    return { success: true };
  }

  sheet.getRange(foundRow, 3).setValue(storedValue);

  const afterSnapshot = buildSettingAuditSnapshot({
    param: paramName,
    value: storedValue
  });

  safeLogAuditEvent(
    'Update',
    'Settings',
    String(paramName || '').trim(),
    'Updated system parameter ' + String(paramName || '').trim(),
    beforeSnapshot,
    afterSnapshot
  );

  return { success: true };
}

// 3. Fetch Lists (Fee Categories, Payment Methods, Statuses)
function getSystemLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) return { categories: [], methods: [], statuses: [] };

  // Column E (5), F (6), G (7)
  // Get all data from Row 6 downwards
  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return { categories: [], methods: [], statuses: [] };

  // Get the grid of data
  const dataRange = sheet.getRange(6, 5, lastRow - 5, 3); 
  const data = dataRange.getValues();

  const categories = [];
  const methods = [];
  const statuses = [];

  data.forEach(row => {
    if (String(row[0]).trim()) categories.push(String(row[0]).trim());
    if (String(row[1]).trim()) methods.push(String(row[1]).trim());
    if (String(row[2]).trim()) statuses.push(String(row[2]).trim());
  });

  return { categories, methods, statuses };
}

// 4. Update a List Item (Used for Categories and Methods)
function updateSystemListItem(columnLetter, oldValue, newValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) throw new Error("Settings worksheet not found.");

  // Convert Column Letter (E, F, G) to index (5, 6, 7)
  const colIndex = columnLetter.charCodeAt(0) - 64; 
  if (colIndex < 5 || colIndex > 7) throw new Error("Invalid column.");

  const listType = getSettingsListType(columnLetter);
  const oldSnapshot = buildSettingAuditSnapshot({
    listType: listType,
    columnLetter: String(columnLetter || '').trim().toUpperCase(),
    oldValue: oldValue,
    newValue: oldValue
  });

  const lastRow = sheet.getLastRow();
  // Search from Row 6 down to find the old value
  for (let r = 6; r <= lastRow; r++) {
    const cell = sheet.getRange(r, colIndex);
    if (String(cell.getValue()).trim() === oldValue) {
      cell.setValue(newValue);

      const newSnapshot = buildSettingAuditSnapshot({
        listType: listType,
        columnLetter: String(columnLetter || '').trim().toUpperCase(),
        oldValue: oldValue,
        newValue: newValue
      });

      safeLogAuditEvent(
        'Update',
        'Settings',
        listType,
        'Updated settings list item in ' + listType,
        oldSnapshot,
        newSnapshot
      );

      return { success: true };
    }
  }
  throw new Error("List item not found.");
}


// 5. Get Score Percentages for Performance Calculations
function getScorePercentages() {
  const params = getSystemParameters();
  let classScorePercentage = 50; // Default
  let examScorePercentage = 50;  // Default
  
  params.forEach(function(item) {
    const paramName = String(item.param || '').trim();
    const paramValue = String(item.value || '').trim();
    
    if (paramName === 'Class Score' && paramValue && !isNaN(paramValue)) {
      classScorePercentage = parseFloat(paramValue);
    }
    if (paramName === 'Exams Score' && paramValue && !isNaN(paramValue)) {
      examScorePercentage = parseFloat(paramValue);
    }
  });
  
  return {
    classScorePercentage: classScorePercentage,
    examScorePercentage: examScorePercentage,
    classScoreDecimal: classScorePercentage / 100,
    examScoreDecimal: examScorePercentage / 100
  };
}
