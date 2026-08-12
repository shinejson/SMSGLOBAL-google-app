// --- PERMISSIONS FUNCTIONS ---

// 1. Fetch Permissions Data
function getPermissionsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Permissions");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; 

  // Row 3 to Last, Col B(2) to Col D(4) => 3 columns
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 3); 
  const data = dataRange.getValues();

  return data.map((row) => ({
    role: String(row[0]).trim(),
    accessLevel: String(row[1]).trim(),
    actions: String(row[2]).trim(),
  }));
}

// 2. Add New Permission
function addPermission(role, accessLevel, actions) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Permissions");
  if (!sheet) throw new Error("Permissions sheet not found");

  // Check if role already exists
  const existingData = getPermissionsData();
  if (existingData.some(p => p.role.toLowerCase() === role.toLowerCase())) {
    throw new Error("Permission role already exists");
  }

  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;

  // Add to row starting from column B
  sheet.getRange(newRow, 2, 1, 3).setValues([
    [role, accessLevel, actions]
  ]);

  return { success: true };
}

// 3. Update Permission
function updatePermission(role, accessLevel, actions) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Permissions");
  if (!sheet) throw new Error("Permissions sheet not found");

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) throw new Error("No permissions found to update");

  // Get all data and find the row
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 3);
  const data = dataRange.getValues();

  let foundRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(role).trim().toLowerCase()) {
      foundRow = i + 3; // +3 because data starts at row 3
      break;
    }
  }

  if (foundRow === -1) {
    throw new Error("Permission role not found");
  }

  // Update the row
  sheet.getRange(foundRow, 2, 1, 3).setValues([
    [role, accessLevel, actions]
  ]);

  return { success: true };
}

// 4. Delete Permission
function deletePermission(role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Permissions");
  if (!sheet) throw new Error("Permissions sheet not found");

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) throw new Error("No permissions found to delete");

  // Get all data and find the row
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 3);
  const data = dataRange.getValues();

  let foundRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(role).trim().toLowerCase()) {
      foundRow = i + 3; // +3 because data starts at row 3
      break;
    }
  }

  if (foundRow === -1) {
    throw new Error("Permission role not found");
  }

  // Delete the row
  sheet.deleteRow(foundRow);

  return { success: true };
}