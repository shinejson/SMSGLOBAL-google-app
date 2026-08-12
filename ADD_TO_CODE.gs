/**
 * ADD THIS CODE TO YOUR Code.js FILE
 * 
 * These are server-side functions that can be called from the Attendance.html page
 * to help diagnose and fix header issues directly from the web interface.
 */

// Add this function to Code.js - it will be callable from the HTML page
function checkAndFixAttendanceHeaders() {
  requireLogin();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  
  if (!sheet) {
    return { 
      success: false, 
      message: "Attendance sheet not found!" 
    };
  }
  
  const lastCol = Math.max(sheet.getLastColumn(), 10);
  const numCols = lastCol - 1;
  
  // Get current headers from Row 2, starting from Column B
  const currentHeaders = sheet.getRange(2, 2, 1, numCols).getValues()[0];
  
  const expectedHeaders = [
    "Attendance ID",
    "Date",
    "Academic Year",
    "Class",
    "Student ID",
    "Student Name",
    "Course ID",
    "Status",
    "Term"
  ];
  
  // Check if headers match
  let mismatchCount = 0;
  const issues = [];
  
  for (let i = 0; i < expectedHeaders.length; i++) {
    const current = String(currentHeaders[i] || "").trim().toLowerCase();
    const expected = expectedHeaders[i].toLowerCase();
    
    if (current !== expected && current.indexOf(expected.replace(/\s+/g, "")) === -1) {
      mismatchCount++;
      const columnLetter = String.fromCharCode(66 + i);
      issues.push({
        column: columnLetter,
        expected: expectedHeaders[i],
        found: currentHeaders[i] || "(empty)"
      });
    }
  }
  
  if (mismatchCount === 0) {
    return {
      success: true,
      message: "Headers are correctly configured!",
      needsFix: false
    };
  }
  
  // Auto-fix headers
  try {
    sheet.getRange(2, 2, 1, expectedHeaders.length).setValues([expectedHeaders]);
    
    // Format the header row
    const headerRange = sheet.getRange(2, 2, 1, expectedHeaders.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4A90E2");
    headerRange.setFontColor("#FFFFFF");
    
    return {
      success: true,
      message: `Fixed ${mismatchCount} header(s). Please refresh the page.`,
      needsFix: true,
      fixed: true,
      issues: issues
    };
  } catch (e) {
    return {
      success: false,
      message: "Error fixing headers: " + e.message,
      needsFix: true,
      fixed: false,
      issues: issues
    };
  }
}

// Add this function to get diagnostic info about attendance data
function getAttendanceDiagnostics() {
  requireLogin();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  
  if (!sheet) {
    return { error: "Attendance sheet not found" };
  }
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  // Get headers
  const headers = sheet.getRange(2, 2, 1, Math.min(lastCol - 1, 9)).getValues()[0];
  
  // Get first 3 data rows as samples
  const sampleData = lastRow >= 3 
    ? sheet.getRange(3, 2, Math.min(lastRow - 2, 3), Math.min(lastCol - 1, 9)).getValues()
    : [];
  
  const hMap = getAttendanceHeaderMap(sheet);
  
  return {
    sheetInfo: {
      totalRows: lastRow,
      totalColumns: lastCol,
      dataRows: lastRow - 2
    },
    headers: headers,
    headerMap: hMap,
    sampleData: sampleData.map((row, idx) => ({
      rowNumber: idx + 3,
      data: row
    }))
  };
}
