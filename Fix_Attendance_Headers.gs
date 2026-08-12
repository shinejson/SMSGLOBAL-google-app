/**
 * ATTENDANCE HEADER FIX UTILITY
 * 
 * This script helps you verify and fix the column headers in your Attendance sheet.
 * Run this function from the Apps Script editor to check your current headers
 * and optionally fix them to match the expected format.
 */

function checkAttendanceHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  
  if (!sheet) {
    Logger.log("ERROR: 'Attendance' sheet not found!");
    return;
  }
  
  const lastCol = Math.max(sheet.getLastColumn(), 10);
  const numCols = lastCol - 1;
  
  // Get current headers from Row 2, starting from Column B
  const currentHeaders = sheet.getRange(2, 2, 1, numCols).getValues()[0];
  
  Logger.log("===== CURRENT HEADERS IN ROW 2 (Starting from Column B) =====");
  currentHeaders.forEach((header, index) => {
    const columnLetter = String.fromCharCode(66 + index); // B=66, C=67, etc.
    Logger.log(`Column ${columnLetter} (Index ${index}): "${header}"`);
  });
  
  Logger.log("\n===== EXPECTED HEADERS =====");
  const expectedHeaders = [
    "Attendance ID",    // Column B (index 0)
    "Date",             // Column C (index 1)
    "Academic Year",    // Column D (index 2)
    "Class",            // Column E (index 3)
    "Student ID",       // Column F (index 4)
    "Student Name",     // Column G (index 5)
    "Course ID",        // Column H (index 6)
    "Status",           // Column I (index 7)
    "Term"              // Column J (index 8)
  ];
  
  expectedHeaders.forEach((header, index) => {
    const columnLetter = String.fromCharCode(66 + index);
    const current = currentHeaders[index] || "(empty)";
    const match = String(current).trim().toLowerCase() === header.toLowerCase() ? "✓" : "✗";
    Logger.log(`${match} Column ${columnLetter}: Expected "${header}", Found "${current}"`);
  });
  
  return {
    current: currentHeaders,
    expected: expectedHeaders
  };
}

function fixAttendanceHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  
  if (!sheet) {
    Logger.log("ERROR: 'Attendance' sheet not found!");
    return { success: false, message: "'Attendance' sheet not found!" };
  }
  
  const correctHeaders = [
    "Attendance ID",    // Column B
    "Date",             // Column C
    "Academic Year",    // Column D
    "Class",            // Column E
    "Student ID",       // Column F
    "Student Name",     // Column G
    "Course ID",        // Column H
    "Status",           // Column I
    "Term"              // Column J
  ];
  
  try {
    // Write correct headers to Row 2, starting from Column B
    sheet.getRange(2, 2, 1, correctHeaders.length).setValues([correctHeaders]);
    
    // Format the header row
    const headerRange = sheet.getRange(2, 2, 1, correctHeaders.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4A90E2");
    headerRange.setFontColor("#FFFFFF");
    
    Logger.log("✓ Headers fixed successfully!");
    Logger.log("New headers:");
    correctHeaders.forEach((header, index) => {
      const columnLetter = String.fromCharCode(66 + index);
      Logger.log(`  Column ${columnLetter}: ${header}`);
    });
    
    return { 
      success: true, 
      message: "Headers fixed successfully! Please refresh your Attendance page." 
    };
  } catch (e) {
    Logger.log("ERROR: " + e.message);
    return { 
      success: false, 
      message: "Error fixing headers: " + e.message 
    };
  }
}

/**
 * Alternative: Check if your data is in the correct columns
 * This will show you a sample of your data row by row
 */
function inspectAttendanceData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  
  if (!sheet) {
    Logger.log("ERROR: 'Attendance' sheet not found!");
    return;
  }
  
  const lastRow = Math.min(sheet.getLastRow(), 7); // Check first 5 data rows
  if (lastRow < 3) {
    Logger.log("No data rows found (sheet has less than 3 rows)");
    return;
  }
  
  const headers = sheet.getRange(2, 2, 1, 9).getValues()[0];
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 9);
  const data = dataRange.getValues();
  
  Logger.log("===== SAMPLE DATA ROWS =====\n");
  
  data.forEach((row, rowIndex) => {
    Logger.log(`--- Row ${rowIndex + 3} ---`);
    row.forEach((cell, colIndex) => {
      const header = headers[colIndex] || `Column ${colIndex}`;
      Logger.log(`  ${header}: "${cell}"`);
    });
    Logger.log("");
  });
}
