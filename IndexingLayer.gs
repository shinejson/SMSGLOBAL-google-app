// --- DATA ACCESS LAYER WITH CACHING AND SANITIZATION ---

/**
 * This file provides cached and sanitized versions of data access functions
 * All data returned from these functions is:
 * 1. Cached for performance (10 minute TTL)
 * 2. Sanitized to prevent XSS attacks
 */

// ==================== STUDENTS ====================

/**
 * Get students data - UNCACHED version (direct sheet access)
 * This is the original function renamed for caching layer
 */
function getStudentsDataUncached() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Students");
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  var headerRowIndex = -1;
  for (var i = 0; i < values.length; i++) {
    if (values[i].some(function (c) { return c !== "" && c !== null; })) {
      headerRowIndex = i;
      break;
    }
  }
  if (headerRowIndex === -1) return [];

  var headers = values[headerRowIndex].map(function (h) {
    return String(h).trim();
  });
  var idx = {};
  headers.forEach(function (h, i) {
    idx[h] = i;
  });

  var rows = values.slice(headerRowIndex + 1).filter(function (r) {
    return r.some(function (c) {
      return c !== "" && c !== null;
    });
  });

  var out = rows.map(function (r) {
    function val(name) {
      if (idx[name] === undefined) return "";
      var v = r[idx[name]];
      if (v instanceof Date) return v.toISOString();
      return v === "" || v === null ? "" : String(v);
    }
    return {
      studentId: val("Student ID") || val("studentId") || "",
      firstName: val("First Name") || "",
      lastName: val("Last Name") || "",
      email: val("Email") || "",
      dob: val("Date of Birth") || val("dob") || "",
      gender: val("Gender") || "",
      enrollmentDate: val("Enrollment Date") || val("enrollmentDate") || "",
      status: val("Status") || "",
      class: val("Class") || val("Class Name") || "",
    };
  });

  return out;
}

/**
 * Get students data - CACHED AND SANITIZED (recommended for use)
 * This replaces the original getStudentsData() function
 */
function getStudentsData() {
  const cachedData = getCachedData('data_Students', getStudentsDataUncached, 600);
  
  // Sanitize each student record before returning
  return cachedData.map(function(student) {
    return sanitizeStudentData(student);
  });
}

// ==================== USERS ====================

/**
 * Get users data - UNCACHED version
 */
function getUsersDataUncached() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const colCount = Math.max(9, sheet.getLastColumn() >= 2 ? sheet.getLastColumn() - 1 : 9);
  const dataRange = sheet.getRange(3, 2, lastRow - 2, colCount);
  const data = dataRange.getValues();

  return data.map((row) => ({
    userId: String(row[0]).trim(),
    googleEmail: String(row[1]).trim(),
    fullName: String(row[2]).trim(),
    role: String(row[3]).trim(),
    accountStatus: String(row[4]).trim(),
    username: String(row[5]).trim(),
    password: "••••••••",
    loginTrials: parseInt(row[7], 10) || 0,
    lockedUntil: row[8] ? String(row[8]) : '',
    isLocked: false,
    lockMinutesRemaining: 0
  }));
}

/**
 * Override getUsersData to use caching and sanitization
 * Note: This is already defined in User.gs, so we create a cached wrapper
 */
function getUsersDataCached() {
  const cachedData = getCachedData('data_Users', getUsersDataUncached, 600);
  
  // Sanitize each user record
  return cachedData.map(function(user) {
    return sanitizeUserData(user);
  });
}

// ==================== COURSES ====================

/**
 * Wrapper to add caching for courses
 */
function getCoursesDataUncached() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Courses");
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  const data = sheet.getRange(3, 2, lastRow - 2, 6).getValues();
  
  return data.map(row => ({
    courseId: String(row[0]).trim(),
    courseName: String(row[1]).trim(),
    courseCode: String(row[2]).trim(),
    department: String(row[3]).trim(),
    credits: String(row[4]).trim(),
    status: String(row[5]).trim()
  }));
}

function getCoursesData() {
  const cachedData = getCachedData('data_Courses', getCoursesDataUncached, 600);
  
  // Sanitize course data
  return cachedData.map(function(course) {
    return {
      courseId: sanitizeHtml(course.courseId),
      courseName: sanitizeHtml(course.courseName),
      courseCode: sanitizeHtml(course.courseCode),
      department: sanitizeHtml(course.department),
      credits: sanitizeHtml(course.credits),
      status: sanitizeHtml(course.status)
    };
  });
}

// ==================== CLASSES ====================

function getClassesDataUncached() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Classes");
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  // Get class names from Column A, starting from Row 2
  const classRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const classData = classRange.getValues();
  
  // Get all students data and count them per class
  let studentCounts = {};
  
  try {
    // Use the existing getStudentsDataUncached which properly handles headers
    const studentsData = getStudentsDataUncached();
    
    // Count students per class
    studentsData.forEach(function(student) {
      const className = String(student.class || '').trim();
      if (className) {
        studentCounts[className] = (studentCounts[className] || 0) + 1;
      }
    });
  } catch (e) {
    Logger.log('Error counting students per class: ' + e.message);
  }
  
  return classData.map((row, idx) => {
    const className = String(row[0]).trim();
    return {
      id: idx + 2, // Sheet row number for reference
      className: className,
      studentCount: studentCounts[className] || 0
    };
  }).filter((r) => r.className);
}

function getClassesData() {
  const cachedData = getCachedData('data_Classes', getClassesDataUncached, 600);
  
  return cachedData.map(function(cls) {
    return {
      id: cls.id,
      className: sanitizeHtml(cls.className),
      studentCount: cls.studentCount
    };
  });
}

// ==================== ATTENDANCE ====================

function getAttendanceDataUncached() {
  // Delegate to Code.js authoritative reader (correct column mapping via header map)
  if (typeof getAttendanceDataFromSheet === "function") {
    return getAttendanceDataFromSheet();
  }
  return [];
}

function getAttendanceData() {
  const cachedData = getCachedData('data_Attendance', getAttendanceDataUncached, 300); // 5 min cache
  
  return cachedData.map(function(att) {
    let dateVal = att.date;
    if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
      try {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } catch (e) {
        dateVal = dateVal.toISOString().split("T")[0];
      }
    } else if (dateVal && String(dateVal).indexOf("T") !== -1) {
      dateVal = String(dateVal).split("T")[0];
    }
    return {
      attendanceId: sanitizeHtml(att.attendanceId),
      date: dateVal,
      academicYear: sanitizeHtml(att.academicYear),
      className: sanitizeHtml(att.className),
      studentId: sanitizeHtml(att.studentId),
      studentName: sanitizeHtml(att.studentName),
      courseId: sanitizeHtml(att.courseId),
      status: sanitizeHtml(att.status),
      term: sanitizeHtml(att.term)
    };
  });
}

// ==================== PERFORMANCE ====================

function getPerformanceDataUncached() {
  return getPerformanceDataFromSheet();
}

function getPerformanceData() {
  const cachedData = getCachedData('data_Performance', getPerformanceDataUncached, 600);
  if (!Array.isArray(cachedData)) return [];

  return cachedData.map(function(perf) {
    return {
      performanceId: sanitizeHtml(perf.performanceId),
      rowNumber: perf.rowNumber,
      studentId: sanitizeHtml(perf.studentId),
      studentName: sanitizeHtml(perf.studentName),
      studentClass: sanitizeHtml(perf.studentClass),
      academicYear: sanitizeHtml(perf.academicYear),
      term: sanitizeHtml(perf.term),
      course: sanitizeHtml(perf.course),
      classScore: perf.classScore,
      examScore100: perf.examScore100,
      examScore60: perf.examScore60,
      examScore50: perf.examScore60,
      total: perf.total,
      rank: sanitizeHtml(perf.rank)
    };
  });
}

// ==================== INVOICES ====================

function getInvoicesDataUncached() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Invoices");
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  const data = sheet.getRange(3, 2, lastRow - 2, 8).getValues();
  
  return data.map(row => ({
    invoiceId: String(row[0]).trim(),
    studentId: String(row[1]).trim(),
    studentName: String(row[2]).trim(),
    category: String(row[3]).trim(),
    amountDue: Number(row[4]) || 0,
    status: String(row[5]).trim(),
    academicYear: String(row[6]).trim(),
    term: String(row[7]).trim()
  }));
}

function getInvoicesData() {
  const cachedData = getCachedData('data_Invoices', getInvoicesDataUncached, 300); // 5 min
  
  return cachedData.map(function(inv) {
    return {
      invoiceId: sanitizeHtml(inv.invoiceId),
      studentId: sanitizeHtml(inv.studentId),
      studentName: sanitizeHtml(inv.studentName),
      category: sanitizeHtml(inv.category),
      amountDue: inv.amountDue,
      status: sanitizeHtml(inv.status),
      academicYear: sanitizeHtml(inv.academicYear),
      term: sanitizeHtml(inv.term)
    };
  });
}

// ==================== PAYMENTS ====================

function getPaymentsDataUncached() {
  return getPaymentsDataFromSheet();
}

function getPaymentsData() {
  const cachedData = getCachedData('data_Payments', getPaymentsDataUncached, 300); // 5 min
  if (!Array.isArray(cachedData)) return [];

  return cachedData.map(function(pay) {
    return {
      transactionId: sanitizeHtml(pay.transactionId),
      invoiceId: sanitizeHtml(pay.invoiceId),
      studentId: sanitizeHtml(pay.studentId),
      studentName: sanitizeHtml(pay.studentName),
      amountPaid: pay.amountPaid,
      paymentDate: pay.paymentDate,
      paymentMethod: sanitizeHtml(pay.paymentMethod),
      referenceNo: sanitizeHtml(pay.referenceNo),
      academicYear: sanitizeHtml(pay.academicYear),
      term: sanitizeHtml(pay.term),
      studentClass: sanitizeHtml(pay.studentClass),
      balance: pay.balance,
      balanceStatus: sanitizeHtml(pay.balanceStatus),
      paymentStatus: sanitizeHtml(pay.paymentStatus || pay.balanceStatus)
    };
  });
}

// ==================== CACHE INVALIDATION HELPERS ====================

/**
 * Invalidate cache when student data changes
 */
function invalidateStudentsCache() {
  invalidateCacheOnModify('Students');
}

/**
 * Invalidate cache when user data changes
 */
function invalidateUsersCache() {
  invalidateCacheOnModify('Users');
}

/**
 * Invalidate cache when performance data changes
 */
function invalidatePerformanceCache() {
  invalidateCacheOnModify('Performance');
}

/**
 * Invalidate cache when invoice/payment data changes
 */
function invalidateFinancialCache() {
  invalidateMultipleCache(['data_Invoices', 'data_Payments', 'dashboard_stats']);
}

/**
 * Invalidate cache when attendance data changes
 */
function invalidateAttendanceCache() {
  invalidateCacheOnModify('Attendance');
}

// ==================== ROW INDEX HELPERS ====================

/**
 * Look up a sheet row number by ID using a cached index.
 * Returns -1 when the index is missing so callers can fall back to linear search.
 */
function findRowByIdIndexed(sheetName, id) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('index_' + sheetName);
    if (!cached) return -1;

    const index = JSON.parse(cached);
    const row = index[String(id).trim()];
    return typeof row === 'number' ? row : -1;
  } catch (e) {
    return -1;
  }
}

/**
 * Clear the cached row index for a sheet after add/update/delete operations.
 */
function invalidateIndex(sheetName) {
  try {
    CacheService.getScriptCache().remove('index_' + sheetName);
  } catch (e) {
    Logger.log('Could not invalidate index for ' + sheetName + ': ' + e.message);
  }
}
