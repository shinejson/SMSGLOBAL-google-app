// --- PARENTS FUNCTIONS ---

// 1. Fetch Parents Data
function getParentsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Parents");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; 

  // Row 3 to Last, Col B(2) to Col I(9) => 8 columns (added Parent Name column)
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 8); 
  const data = dataRange.getValues();

  return data
    .filter(row => String(row[0]).trim() !== "") // Filter out empty rows
    .map((row) => ({
      mappingId: String(row[0]).trim(),
      parentUserId: String(row[1]).trim(),
      parentName: String(row[2] || '').trim(), // Parent Name
      studentId: String(row[3]).trim(),
      studentName: String(row[4]).trim(),
      relationship: String(row[5]).trim(),
      isPrimary: row[6] === true || row[6] === "TRUE" || row[6] === 1, // Handles Checkbox/Boolean
      billing: String(row[7]).trim(),
    }));
}

// 2. Helper to generate Mapping ID (e.g. PAR-1001)
function generateNextParentId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "PAR-1001";

  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues(); 
  let maxIdNum = 1000;
  values.forEach((row) => {
    const idStr = String(row[0]).trim();
    if (idStr.startsWith("PAR-")) {
      const num = parseInt(idStr.substring(4), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return "PAR-" + (maxIdNum + 1);
}

// 3. Helper: Get just User IDs and Full Names for dropdown
function getUserPairs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  // Get ID (Col B), Full Name (Col D)
  const data = sheet.getRange(3, 2, lastRow - 2, 3).getValues();
  return data.map(row => ({
    userId: String(row[0]).trim(),
    fullName: String(row[2]).trim()
  }));
}

// 4. Helper: Find a student's Name by their ID
function getStudentNameById(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  if (!sheet) return "";

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "";

  // Column B: ID, Column C: First, Column D: Last
  const data = sheet.getRange(3, 2, lastRow - 2, 3).getValues();
  for (let row of data) {
    if (String(row[0]).trim() === studentId) {
      return `${String(row[1]).trim()} ${String(row[2]).trim()}`;
    }
  }
  return "";
}

// 5. Add New Mapping (Create)
function addParentMapping(mapData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Parents");
  if (!sheet) throw new Error("Parents worksheet not found.");

  const nextId = generateNextParentId(sheet);
  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  // Write to Column B to Column I (8 columns including Parent Name)
  sheet.getRange(targetRow, 2, 1, 8).setValues([
    [
      nextId,
      mapData.parentUserId,
      mapData.parentName || '', // Parent Name
      mapData.studentId,
      mapData.studentName,
      mapData.relationship,
      mapData.isPrimary, // Google Sheets will automatically treat TRUE/FALSE as checkboxes
      mapData.billing
    ]
  ]);

  return { success: true, mappingId: nextId };
}

// 6. Update Mapping (Update)
function updateParentMapping(mappingId, mapData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Parents");
  if (!sheet) throw new Error("Parents worksheet not found.");

  // Because mappingId is unique, we must find the row
  const lastRow = sheet.getLastRow();
  let rowIndex = -1;
  for (let r = 3; r <= lastRow; r++) {
    const cell = sheet.getRange(r, 2).getValue();
    if (String(cell).trim() === mappingId) {
      rowIndex = r;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Mapping record not found.");

  // Update from Col C(3) to Col I(9) -> 7 columns (including Parent Name)
  sheet.getRange(rowIndex, 3, 1, 7).setValues([
    [
      mapData.parentUserId,
      mapData.parentName || '', // Parent Name
      mapData.studentId,
      mapData.studentName,
      mapData.relationship,
      mapData.isPrimary,
      mapData.billing
    ]
  ]);

  return { success: true };
}

// 7. Delete Mapping (Delete)
function deleteParentMapping(mappingId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Parents");
  if (!sheet) throw new Error("Parents worksheet not found.");

  const lastRow = sheet.getLastRow();
  let rowIndex = -1;
  for (let r = 3; r <= lastRow; r++) {
    const cell = sheet.getRange(r, 2).getValue();
    if (String(cell).trim() === mappingId) {
      rowIndex = r;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Mapping record not found.");

  sheet.deleteRow(rowIndex);
  return { success: true };
}