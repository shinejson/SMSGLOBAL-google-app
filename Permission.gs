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

function buildPermissionAuditSnapshot(permissionData) {
  if (!permissionData) return null;

  return buildAuditSnapshot({
    role: permissionData.role,
    accessLevel: permissionData.accessLevel,
    actions: permissionData.actions
  });
}

function findPermissionRowByRole(sheet, role) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return -1;

  const data = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(role).trim().toLowerCase()) {
      return i + 3;
    }
  }

  return -1;
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

  const createdPermission = buildPermissionAuditSnapshot({
    role: role,
    accessLevel: accessLevel,
    actions: actions
  });

  safeLogAuditEvent(
    'Create',
    'Permissions',
    String(role || '').trim(),
    'Created permission role ' + String(role || '').trim(),
    null,
    createdPermission
  );

  return { success: true };
}

// 3. Update Permission
function updatePermission(role, accessLevel, actions) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Permissions");
  if (!sheet) throw new Error("Permissions sheet not found");

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) throw new Error("No permissions found to update");

  const existingPermission = getPermissionsData().find((permission) => String(permission.role).trim().toLowerCase() === String(role).trim().toLowerCase()) || null;
  const foundRow = findPermissionRowByRole(sheet, role);

  if (foundRow === -1) {
    throw new Error("Permission role not found");
  }

  // Update the row
  sheet.getRange(foundRow, 2, 1, 3).setValues([
    [role, accessLevel, actions]
  ]);

  const updatedPermission = buildPermissionAuditSnapshot({
    role: role,
    accessLevel: accessLevel,
    actions: actions
  });

  safeLogAuditEvent(
    'Update',
    'Permissions',
    String(role || '').trim(),
    'Updated permission role ' + String(role || '').trim(),
    buildPermissionAuditSnapshot(existingPermission),
    updatedPermission
  );

  return { success: true };
}

// 4. Delete Permission
function deletePermission(role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Permissions");
  if (!sheet) throw new Error("Permissions sheet not found");

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) throw new Error("No permissions found to delete");

  const existingPermission = getPermissionsData().find((permission) => String(permission.role).trim().toLowerCase() === String(role).trim().toLowerCase()) || null;
  const foundRow = findPermissionRowByRole(sheet, role);

  if (foundRow === -1) {
    throw new Error("Permission role not found");
  }

  // Delete the row
  sheet.deleteRow(foundRow);

  safeLogAuditEvent(
    'Delete',
    'Permissions',
    String(role || '').trim(),
    'Deleted permission role ' + String(role || '').trim(),
    buildPermissionAuditSnapshot(existingPermission),
    null
  );

  return { success: true };
}
