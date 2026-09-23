function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || "Index";
  var embedded = e && e.parameter && e.parameter.embedded === "true";

  // Special page for password migration utility (Admin only)
  if (page === "Migration_Utility") {
    if (!checkSession()) {
      return HtmlService.createTemplateFromFile("Login").evaluate();
    }
    var user = getLoggedInUser();
    if (!user || user.role !== "Admin") {
      return ContentService.createTextOutput("Access Denied: Admin only");
    }
    return HtmlService.createTemplateFromFile("Migration_Utility")
      .evaluate()
      .setTitle("Password Security Migration")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Whitelist of pages allowed to be loaded inside the iframe
  var EMBEDDED_ALLOWED = [
    "Student",
    "Courses",
    "Attendance",
    "Enrollments",
    "Classes",
    "Teachers",
    "Academic Years",
    "Performance",
    "Invoices",
    "Payments",
    "Settings",
    "Users",
    "Parents",
    "Permissions",
    "Reports",
    "Billings",
    "Billing Categories",
    "AuditTrials",
  ];

  // Special download endpoint: stream a Drive PDF to the browser
  if (page === "download") {
    var fileId = (e && e.parameter && e.parameter.fileId) || null;
    if (!fileId) return ContentService.createTextOutput("Missing fileId");
    // Ensure user is authenticated for downloads
    if (!checkSession()) {
      return HtmlService.createTemplateFromFile("Login").evaluate();
    }
    try {
      var file = DriveApp.getFileById(fileId);
      // Ensure the file is at least viewable by link
      try {
        file.setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW,
        );
      } catch (e) {
        /* ignore */
      }
      var downloadUrl =
        "https://docs.google.com/uc?export=download&id=" +
        encodeURIComponent(fileId);
      var html =
        '<!doctype html><html><head><meta charset="utf-8"><title>Download</title></head><body>' +
        '<p>If your download does not start automatically, <a id="lnk" href="' +
        downloadUrl +
        '">click here</a>.</p>' +
        '<script>window.location.replace("' +
        downloadUrl +
        '");<\/script></body></html>';
      return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL,
      );
    } catch (err) {
      return ContentService.createTextOutput(
        "File not found or access denied.",
      );
    }
  }

  // Allow whitelisted pages loaded in an iframe to bypass session re-check
  var skipSessionCheck = embedded && EMBEDDED_ALLOWED.indexOf(page) !== -1;

  // Protect whitelisted pages / protect all pages except Login/skipSessionCheck whitelisted
  if (page !== "Login" && !skipSessionCheck && !checkSession()) {
    page = "Login";
  }

  if (page === "Login") {
    var template = HtmlService.createTemplateFromFile("Login");
    template.scriptUrl = ScriptApp.getService().getUrl();
    return template
      .evaluate()
      .setTitle("Login - SMS")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var template = HtmlService.createTemplateFromFile(page);
  if (page === "Index") {
    template.scriptUrl = ScriptApp.getService().getUrl();
  }

  return template
    .evaluate()
    .setTitle(
      page === "Student"
        ? "Student Directory"
        : page === "Reports"
          ? "System Reports"
          : "Student Management System",
    )
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Get audit logs statistics for dashboard display
 * Returns key metrics about audit log activity
 */
function getAuditStats() {
  try {
    const stats = getAuditStatistics();
    
    if (!stats) {
      return {
        error: true,
        message: 'Failed to retrieve audit statistics'
      };
    }
    
    // Calculate deletions from action breakdown
    const deletions = stats.actionBreakdown['Delete'] || 0;
    
    return {
      totalActions: stats.totalLogs || 0,
      todayActions: stats.todayLogs || 0,
      activeUsers: stats.uniqueUsers || 0,
      deletions: deletions
    };
  } catch (error) {
    Logger.log('Error in getAuditStats: ' + error.message);
    return {
      error: true,
      message: error.message
    };
  }
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// Session helpers live in User.gs (checkSession, loginUser, logoutUser).

// --- ROLE-BASED ACCESS CONTROL HELPERS ---

function getAccessLevelForRole(role) {
  if (!role) return "General Access";
  // Access getPermissionsData() defined globally in Permission.gs
  try {
    const permissions = getPermissionsData();
    const found = permissions.find(
      (p) => p.role.trim().toLowerCase() === role.trim().toLowerCase(),
    );
    return found ? found.accessLevel : role;
  } catch (e) {
    console.error("Error fetching permissions from sheet:", e);
    return role; // Fallback to raw role name if sheets are not readable
  }
}

function getAllowedPagesForRole(role) {
  const allPages = [
    "Student",
    "Courses",
    "Attendance",
    "Enrollments",
    "Classes",
    "Teachers",
    "Academic Years",
    "Performance",
    "Invoices",
    "Payments",
    "Settings",
    "Users",
    "Parents",
    "Permissions",
    "Reports",
    "Billings",
    "Billing Categories",
  ];

  const rawAccessLevel = getAccessLevelForRole(role);
  const access = rawAccessLevel.toLowerCase().trim();

  // Map Access Level keywords to permitted pages
  if (access.includes("full access") || access === "admin") {
    return allPages;
  }
  if (access.includes("financial") || access === "bursar") {
    return [
      "Student",
      "Parents",
      "Invoices",
      "Payments",
      "Billings",
      "Billing Categories",
      "Reports",
      "Classes",
      "Courses",
    ];
  }
  if (access.includes("read only") || access === "viewer") {
    return [
      "Student",
      "Courses",
      "Classes",
      "Teachers",
      "Attendance",
      "Academic Years",
      "Parents",
      "Reports",
    ];
  }
  if (access.includes("restricted") || access === "parent") {
    return ["Invoices", "Payments"];
  }
  if (access.includes("class management") || access.includes("teacher")) {
    return [
      "Student",
      "Courses",
      "Classes",
      "Attendance",
      "Performance",
      "Reports",
    ];
  }
  if (access.includes("general") || access === "staff") {
    return [
      "Student",
      "Courses",
      "Classes",
      "Teachers",
      "Attendance",
      "Academic Years",
    ];
  }

  // Safe fallback list (public/common views)
  return ["Student", "Courses", "Classes", "Attendance"];
}

function getPageNameFromSheetName(sheetName) {
  const nameLower = sheetName.trim().toLowerCase();
  if (nameLower.includes("billing categories")) return "Billing Categories";
  if (nameLower.includes("billings")) return "Billings";
  if (nameLower.includes("student")) return "Student";
  if (nameLower.includes("course") || nameLower.includes("subject")) return "Courses";
  if (nameLower.includes("attend")) return "Attendance";
  if (nameLower.includes("enroll")) return "Enrollments";
  if (nameLower.includes("classes")) return "Classes";
  if (nameLower.includes("teacher")) return "Teachers";
  if (nameLower.includes("academic")) return "Academic Years";
  if (nameLower.includes("performance")) return "Performance";
  if (nameLower.includes("invoice")) return "Invoices";
  if (nameLower.includes("payment")) return "Payments";
  if (nameLower.includes("settings")) return "Settings";
  if (nameLower.includes("user")) return "Users";
  if (nameLower.includes("parent")) return "Parents";
  if (nameLower.includes("permission")) return "Permissions";
  if (nameLower.includes("report")) return "Reports";
  return sheetName;
}

/**
 * Returns the evaluated HTML string for a given page name.
 * Called via google.script.run so session is already validated by the
 * parent Index page — no second session check needed here.
 */
function getPageHtml(pageName) {
  requireLogin();
  const user = getLoggedInUser();
  const role = user ? user.role : "";
  const allowedPages = getAllowedPagesForRole(role);

  if (allowedPages.indexOf(pageName) === -1) {
    throw new Error(
      "Access Denied: You do not have permission to access the '" +
        pageName +
        "' module.",
    );
  }

  try {
    return HtmlService.createTemplateFromFile(pageName).evaluate().getContent();
  } catch (e) {
    throw new Error('Could not load page "' + pageName + '": ' + e.message);
  }
}

/**
 * Search across common entities (pages, students, invoices) and return lightweight matches.
 * Called from client-side search to provide jump-to behaviour.
 */
function searchIndex(query) {
  query = String(query || '').trim().toLowerCase();
  if (query === '') return [];

  const results = [];
  try {
    // Pages (sheet-backed)
    const sheets = getSheetNames();
    sheets.forEach(function(name) {
      if (String(name || '').toLowerCase().indexOf(query) !== -1) {
        results.push({ type: 'page', page: getPageNameFromSheetName(name), title: name });
      }
    });

    // Students
    const students = typeof getStudentsData === 'function' ? getStudentsData() : [];
    students.forEach(function(s) {
      const full = ((s.firstName || '') + ' ' + (s.lastName || '')).trim().toLowerCase();
      const id = String(s.studentId || '').toLowerCase();
      if (full.indexOf(query) !== -1 || id.indexOf(query) !== -1) {
        results.push({ type: 'student', studentId: s.studentId, name: (s.firstName || '') + ' ' + (s.lastName || '') });
      }
    });

    // Invoices (search by id, student name or description)
    if (typeof getInvoicesData === 'function') {
      const invoices = getInvoicesData();
      invoices.forEach(function(inv) {
        const invId = String(inv.invoiceId || '').toLowerCase();
        const studentName = String(inv.studentName || '').toLowerCase();
        const desc = String(inv.items || inv.description || '').toLowerCase();
        if (invId.indexOf(query) !== -1 || studentName.indexOf(query) !== -1 || desc.indexOf(query) !== -1) {
          results.push({ type: 'invoice', invoiceId: inv.invoiceId, label: inv.invoiceId || inv.studentName || inv.items });
        }
      });
    }

    // Payments (search by transactionId, invoiceId or student name)
    if (typeof getPaymentsData === 'function') {
      const payments = getPaymentsData();
      payments.forEach(function(p) {
        const tx = String(p.transactionId || '').toLowerCase();
        const invRef = String(p.invoiceId || '').toLowerCase();
        const studentName = String(p.studentName || '').toLowerCase();
        if (tx.indexOf(query) !== -1 || invRef.indexOf(query) !== -1 || studentName.indexOf(query) !== -1) {
          results.push({ type: 'payment', transactionId: p.transactionId, label: p.transactionId || p.invoiceId || p.studentName });
        }
      });
    }

    // Billings (search billing items)
    if (typeof getBillingsData === 'function') {
      const billings = getBillingsData();
      billings.forEach(function(b) {
        const id = String(b.billingId || '').toLowerCase();
        const item = String(b.item || '').toLowerCase();
        const desc = String(b.description || '').toLowerCase();
        if (id.indexOf(query) !== -1 || item.indexOf(query) !== -1 || desc.indexOf(query) !== -1) {
          results.push({ type: 'billing', billingId: b.billingId, label: b.item || b.billingId });
        }
      });
    }

    // Users
    if (typeof getUsersData === 'function') {
      const users = getUsersData();
      users.forEach(function(u) {
        const id = String(u.userId || '').toLowerCase();
        const name = String(u.fullName || '').toLowerCase();
        if (id.indexOf(query) !== -1 || name.indexOf(query) !== -1) {
          results.push({ type: 'user', userId: u.userId, label: u.fullName || u.userId });
        }
      });
    }

    // Parents
    if (typeof getParentsData === 'function') {
      const parents = getParentsData();
      parents.forEach(function(p) {
        const id = String(p.parentId || p.id || '').toLowerCase();
        const name = String(p.parentName || p.fullName || '').toLowerCase();
        if (id.indexOf(query) !== -1 || name.indexOf(query) !== -1) {
          results.push({ type: 'parent', parentId: p.parentId || p.id, label: p.parentName || p.fullName });
        }
      });
    }
  } catch (e) {
    // swallow errors and return whatever matches we gathered
    console.error('searchIndex error', e);
  }

  return results.slice(0, 50);
}

// Function to return the names of the sheets in the sidebar
function getSheetNames() {
  requireLogin();
  const user = getLoggedInUser();
  const role = user ? user.role : "";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Filter out sheets that don't have corresponding HTML files
  const sheetsToSkip = [];

  const names = sheets
    .map((sheet) => sheet.getName())
    .filter((name) => !sheetsToSkip.includes(name));

  // Reports is a virtual dashboard page not backed by a physical sheet, so we manually append it.
  if (names.indexOf("Reports") === -1) {
    names.push("Reports");
  }

  // Filter sheets list according to user's permissions
  const allowedPages = getAllowedPagesForRole(role);
  return names.filter((name) => {
    const pageName = getPageNameFromSheetName(name);
    return allowedPages.indexOf(pageName) !== -1;
  });
}

// Helper to parse date string/object into a Date object at local midnight
function parseDashboardDate(input) {
  if (!input) return null;
  if (input instanceof Date) return isNaN(input.getTime()) ? null : input;
  const str = String(input).trim();
  if (!str) return null;
  // If ISO date string YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    const parts = str.substring(0, 10).split('-');
    return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  }
  // If YYYY/MM/DD
  if (/^\d{4}\/\d{2}\/\d{2}/.test(str)) {
    const parts = str.substring(0, 10).split('/');
    return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  }
  // If DD/MM/YYYY or MM/DD/YYYY
  const parts = str.split(/[\/\-\.]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    } else if (parts[2].length === 4) {
      let p1 = Number(parts[0]), p2 = Number(parts[1]), y = Number(parts[2]);
      if (p1 > 12) return new Date(y, p2 - 1, p1);
      return new Date(y, p2 - 1, p1);
    }
  }
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

// Helper to determine if an attendance status counts as present
function isAttendancePresentStatus(statusVal, statusesList) {
  const s = String(statusVal || '').trim().toLowerCase();
  if (!s || s === 'undefined' || s === 'null') return false;
  if (s === 'present' || s === 'p' || s.startsWith('present') || s === 'late' || s === 'on time') return true;
  if (statusesList && statusesList.length > 0) {
    const firstStatus = String(statusesList[0].value || '').trim().toLowerCase();
    if (firstStatus && s === firstStatus) return true;
  }
  return false;
}

// ============================================================
// DASHBOARD STATS (Index page)
// ------------------------------------------------------------
// Filter semantics, matching the filter bar in Index.html:
//   * Academic Year + Term -> Attendance, Invoices and Payments.
//     When the year dropdown is left on its default, the ACTIVE academic year
//     is used, so the billing snapshot always reflects the current year.
//     The sentinel value DASHBOARD_ALL_YEARS ("ALL") opts out of year filtering.
//   * Date -> Attendance only. Invoices and payments are a year + term snapshot
//     and are deliberately NOT narrowed down to a single day.
//   * Students and Courses are never filtered: neither the Students sheet nor
//     the Courses sheet has an Academic Year column, so "Total Students" and
//     "Students by Class" always report the whole student population.
// ============================================================

// Sentinel value used by the dashboard year dropdown for "All Academic Years".
var DASHBOARD_ALL_YEARS = "ALL";

// Normalise an academic year for comparison so "2025/2026", "2025-2026" and
// the sanitized "2025&#x2F;2026" all collapse to the same key.
function normalizeAcademicYearKey(value) {
  const raw = typeof decodeSanitizedHtml === "function" ? decodeSanitizedHtml(value) : value;
  return String(raw || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

// Resolve the dashboard year filter.
// "" (default) -> the active academic year, "ALL" -> no year filtering,
// anything else -> that exact year.
function resolveDashboardAcademicYear(selectedYear) {
  const requested = String(
    selectedYear === null || selectedYear === undefined ? "" : selectedYear
  ).trim();

  if (requested.toUpperCase() === DASHBOARD_ALL_YEARS) {
    return { year: "", wantsAllYears: true };
  }

  if (requested !== "") {
    return { year: requested, wantsAllYears: false };
  }

  // Nothing picked -> fall back to the active academic year so billing never
  // silently mixes several academic years together.
  let active = "";
  try {
    if (typeof getActiveAcademicYearValue === "function") {
      active = String(getActiveAcademicYearValue() || "").trim();
    }
  } catch (e) {
    active = "";
  }

  return { year: active, wantsAllYears: false };
}

// Invoice rows carry free text payment statuses ("Paid", "Unpaid", "Partial",
// "Overdue", ...). Bucket them for the Quick Insights breakdown.
function classifyInvoicePaymentStatus(statusValue) {
  const s = String(statusValue || "").trim().toLowerCase();
  if (s === "" || s.indexOf("unpaid") !== -1 || s.indexOf("pending") !== -1 || s.indexOf("due") !== -1) {
    return "unpaid";
  }
  if (s.indexOf("partial") !== -1 || s.indexOf("part ") !== -1) {
    return "partial";
  }
  if (s.indexOf("paid") !== -1 || s.indexOf("settled") !== -1 || s.indexOf("cleared") !== -1) {
    return "paid";
  }
  return "unpaid";
}

function getDashboardStats(selectedYear, selectedTerm, selectedDate) {
  requireLogin();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // getStudentsData()/getInvoicesData()/getPaymentsData() return sanitized
  // values, while the year/term filters arrive raw. Decode before comparing.
  const decode =
    typeof decodeSanitizedHtml === "function"
      ? decodeSanitizedHtml
      : function (v) {
          return v;
        };

  // --- Resolve the filters ---------------------------------------------
  const yearFilter = resolveDashboardAcademicYear(selectedYear);
  const normYear = yearFilter.year ? normalizeAcademicYearKey(yearFilter.year) : "";
  const normTerm = selectedTerm ? String(decode(selectedTerm)).trim().toLowerCase() : "";

  // Date filter: attendance only. Billing is a year/term snapshot.
  const filterDate = parseDashboardDate(selectedDate);

  function matchYear(itemYear) {
    if (!normYear) return true; // "All Academic Years"
    const y = normalizeAcademicYearKey(itemYear);
    // Records without an academic year are not attributed to the selected year
    // (same rule as the Invoices/Payments page year filters).
    return y !== "" && y === normYear;
  }

  function matchTerm(itemTerm) {
    if (!normTerm) return true;
    const t = String(decode(itemTerm) || "").trim().toLowerCase();
    if (!t) return true; // legacy rows without a term still count
    return t === normTerm || t.indexOf(normTerm) !== -1 || normTerm.indexOf(t) !== -1;
  }

  function matchYT(itemYear, itemTerm) {
    return matchYear(itemYear) && matchTerm(itemTerm);
  }

  function matchDate(itemDate) {
    if (!filterDate) return true;
    if (!itemDate) return false;

    try {
      const compareDate = parseDashboardDate(itemDate);
      if (!compareDate) return false;

      // Match if the dates are on the same day
      return (
        compareDate.getFullYear() === filterDate.getFullYear() &&
        compareDate.getMonth() === filterDate.getMonth() &&
        compareDate.getDate() === filterDate.getDate()
      );
    } catch (e) {
      return false;
    }
  }

  const scopeYearLabel = yearFilter.year ? decode(yearFilter.year) : "All academic years";
  const scopeTermLabel = normTerm ? String(decode(selectedTerm)).trim() : "All terms";
  const billingScopeLabel = scopeYearLabel + " · " + scopeTermLabel;

  // --- Students: whole population, independent of every dashboard filter --
  // The Students sheet has no Academic Year column, so students are never
  // filtered by year, term or date.
  let totalStudents = 0;
  const studentClassCounts = {};
  let allStudents = [];
  try {
    allStudents = typeof getStudentsData === "function" ? getStudentsData() : [];
  } catch (e) {
    allStudents = [];
  }

  totalStudents = allStudents.length;
  allStudents.forEach(function (student) {
    const className =
      String(student.class || student.Class || student["Class Name"] || "").trim() || "Unassigned";
    studentClassCounts[className] = (studentClassCounts[className] || 0) + 1;
  });

  // --- Active courses: independent of academic year and date filters ------
  let activeCourses = 0;
  const coursesSheet = ss.getSheetByName("Courses");
  if (coursesSheet) {
    const lastRow = coursesSheet.getLastRow();
    if (lastRow >= 3) {
      const data = coursesSheet.getRange(3, 2, lastRow - 2, 7).getValues();
      activeCourses = data.filter((row) => {
        const id = String(row[0]).trim();
        const status = String(row[5] || row[4] || "").trim().toLowerCase();
        if (id === "") return false;
        if (status !== "active") return false;
        return true; // Count all active courses regardless of year/term
      }).length;
    }
  }

  // --- Attendance: year + term, plus the single day when one is selected --
  let avgAttendance = "0%";
  const attendanceTrend = [];
  const attendanceSheet = ss.getSheetByName("Attendance");
  if (attendanceSheet) {
    const lastRow = attendanceSheet.getLastRow();
    if (lastRow >= 3) {
      const rawAttendanceData = typeof getAttendanceData === "function" ? getAttendanceData() : [];
      // getAttendanceData() returns HTML-escaped values, while the year/term
      // filters arrive raw. Decode so names like "2025/2026" match correctly.
      const attendanceData = rawAttendanceData.map(function (row) {
        return Object.assign({}, row, {
          academicYear: decode(row.academicYear),
          term: decode(row.term),
          status: decode(row.status),
        });
      });
      const statuses = typeof getAttendanceStatuses === "function" ? getAttendanceStatuses() : [];

      // Valid records matching Year, Term, and Date filter (if selected)
      const validRecords = attendanceData.filter(
        (row) => matchYT(row.academicYear, row.term) && matchDate(row.date)
      );

      if (validRecords.length > 0) {
        const presentCount = validRecords.filter((row) =>
          isAttendancePresentStatus(row.status, statuses)
        ).length;
        const percentage = Math.round((presentCount / validRecords.length) * 100);
        avgAttendance = percentage + "%";
      }

      // Year & Term records for the monthly trend chart (independent of the
      // single date filter)
      const yearTermRecords = attendanceData.filter((row) => matchYT(row.academicYear, row.term));

      const trendMap = {};
      yearTermRecords.forEach((row) => {
        const dateValue = parseDashboardDate(row.date);
        if (!dateValue || isNaN(dateValue.getTime())) return;

        const monthLabel = Utilities.formatDate(
          dateValue,
          Session.getScriptTimeZone(),
          "MMM yyyy"
        );
        const sortKey = dateValue.getFullYear() * 100 + (dateValue.getMonth() + 1);

        if (!trendMap[monthLabel]) {
          trendMap[monthLabel] = { present: 0, total: 0, sortKey: sortKey };
        }
        trendMap[monthLabel].total += 1;
        if (isAttendancePresentStatus(row.status, statuses)) {
          trendMap[monthLabel].present += 1;
        }
      });

      const sortedTrendKeys = Object.keys(trendMap).sort(
        (a, b) => trendMap[a].sortKey - trendMap[b].sortKey
      );
      sortedTrendKeys.slice(-6).forEach((key) => {
        const item = trendMap[key];
        attendanceTrend.push({
          label: key,
          presentPct: item.total > 0 ? Math.round((item.present / item.total) * 100) : 0,
          presentCount: item.present,
          totalCount: item.total,
        });
      });
    }
  }

  // --- Invoices and payments: academic year + term only -------------------
  let invoiceCount = 0;
  let totalInvoiced = 0;
  let totalPaid = 0;
  let totalUnpaid = 0;
  let paymentCount = 0;
  const invoiceStatusCounts = { paid: 0, partial: 0, unpaid: 0 };

  // Class-based payment statistics (year/term scoped, never date scoped)
  const classPaymentStats = {};

  try {
    const invoices = typeof getInvoicesData === "function" ? getInvoicesData() : [];
    const payments = typeof getPaymentsData === "function" ? getPaymentsData() : [];

    // NOTE: invoices expose issueDate/dueDate (there is no invoiceDate field),
    // so billing is intentionally not narrowed by the dashboard date filter.
    const filteredInvoices = invoices.filter((inv) => matchYT(inv.academicYear, inv.term));
    const filteredPayments = payments.filter((p) => matchYT(p.academicYear, p.term));

    invoiceCount = filteredInvoices.length;
    totalInvoiced = filteredInvoices.reduce(
      (sum, inv) => sum + (Number(inv.amountDue) || 0),
      0
    );

    paymentCount = filteredPayments.length;
    totalPaid = filteredPayments.reduce((sum, p) => sum + (Number(p.amountPaid) || 0), 0);
    totalUnpaid = Math.max(0, totalInvoiced - totalPaid);

    filteredInvoices.forEach((inv) => {
      invoiceStatusCounts[classifyInvoicePaymentStatus(decode(inv.paymentStatus))] += 1;
    });

    // Students are not year scoped, so every student is counted per class and
    // "paid" means the student has at least one payment in the selected
    // academic year/term.
    const paidStudentIds = new Set();
    filteredPayments.forEach((p) => {
      const id = String(p.studentId || "").trim();
      if (id) paidStudentIds.add(id);
    });

    allStudents.forEach((student) => {
      const className =
        String(student.class || student.Class || student["Class Name"] || "").trim() ||
        "Unassigned";
      const studentId = String(student.studentId || "").trim();

      if (!classPaymentStats[className]) {
        classPaymentStats[className] = { totalStudents: 0, paidStudents: 0 };
      }

      classPaymentStats[className].totalStudents += 1;

      if (studentId && paidStudentIds.has(studentId)) {
        classPaymentStats[className].paidStudents += 1;
      }
    });
  } catch (e) {
    invoiceCount = 0;
    totalInvoiced = 0;
    totalPaid = 0;
    totalUnpaid = 0;
    paymentCount = 0;
  }

  // Collected vs outstanding always add up to 100% of what was billed.
  const collectionRate =
    totalInvoiced > 0
      ? Math.round((Math.min(totalPaid, totalInvoiced) / totalInvoiced) * 100)
      : 0;

  return {
    totalStudents: totalStudents,
    activeCourses: activeCourses,
    avgAttendance: avgAttendance,
    invoiceCount: invoiceCount,
    totalInvoiced: totalInvoiced,
    totalPaid: totalPaid,
    totalUnpaid: totalUnpaid,
    paymentCount: paymentCount,
    paidInvoiceCount: invoiceStatusCounts.paid,
    partialInvoiceCount: invoiceStatusCounts.partial,
    unpaidInvoiceCount: invoiceStatusCounts.unpaid,
    collectionRate: collectionRate,
    billingAcademicYear: yearFilter.year,
    billingScopeLabel: billingScopeLabel,
    attendanceTrend: attendanceTrend,
    studentClassCounts: studentClassCounts,
    classPaymentStats: classPaymentStats,
  };
}

function requireLogin() {
  if (!checkSession()) {
    throw new Error("Authentication required. Please sign in.");
  }
}

// --- LICENSE & DEMO LIMIT HELPERS ---
function getLicenseDetails() {
  const status = String(getSystemParameter('License Status') || 'Demo').trim();
  const limitValue = parseInt(getSystemParameter('Student Limit'), 10);
  const limit = isNaN(limitValue) ? 10 : limitValue;
  const devEmail = String(getSystemParameter('Developer Email') || '').trim();

  // Test Mode overrides Demo restrictions — grant unlimited access
  const testMode = isTestModeActive();
  if (testMode.active) {
    return {
      status: 'Test Mode',
      limit: 999999,
      developerEmail: devEmail,
      testMode: true,
      testModeDeadline: testMode.deadline
    };
  }

  return { status: status, limit: limit, developerEmail: devEmail };
}

// --- TEST MODE ---
// Owner constant — the only account that can activate/deactivate Test Mode
var TEST_MODE_OWNER_EMAIL = 'shineakakpo08@gmail.com';

/**
 * Checks whether Test Mode is currently active.
 * If the deadline has already passed it auto-clears the stored flags.
 * Returns { active: Boolean, deadline: ISO-string | null }
 */
function isTestModeActive() {
  try {
    const props = PropertiesService.getScriptProperties();
    const active   = props.getProperty('test_mode_active');
    const deadline = props.getProperty('test_mode_deadline');

    if (active !== 'true') return { active: false, deadline: null };

    // Check if deadline has expired
    if (deadline) {
      const deadlineMs = parseInt(deadline, 10);
      if (!isNaN(deadlineMs) && Date.now() > deadlineMs) {
        // Expired — clear and return inactive
        props.deleteProperty('test_mode_active');
        props.deleteProperty('test_mode_deadline');
        return { active: false, deadline: null };
      }
      return { active: true, deadline: new Date(deadlineMs).toISOString() };
    }

    return { active: true, deadline: null };
  } catch (e) {
    return { active: false, deadline: null };
  }
}

/**
 * Generates and emails a one-time OTP to the owner email for Test Mode activation.
 */
function generateTestModeOtp() {
  requireLogin();
  const userEmail = getLoggedInUserEmail_();
  if (userEmail.toLowerCase() !== TEST_MODE_OWNER_EMAIL.toLowerCase()) {
    return { success: false, message: 'Only the system owner can activate Test Mode.' };
  }

  const code = String(Math.floor(100000 + Math.random() * 900000));
  const timestamp = Date.now();

  const props = PropertiesService.getScriptProperties();
  props.setProperty('test_mode_otp', code);
  props.setProperty('test_mode_otp_ts', String(timestamp));

  const subject = '🧪 SMS Test Mode Activation Code';
  const body =
    'Your Test Mode activation code is: ' + code + '\n\n' +
    'This code expires in 10 minutes.\n' +
    'Use it in the SMS application to enable Test Mode.\n\n' +
    'If you did not request this, please ignore this email.';

  try {
    GmailApp.sendEmail(TEST_MODE_OWNER_EMAIL, subject, body, {
      name: 'SMS School App',
      noReply: true
    });
    return { success: true, message: 'OTP sent to ' + TEST_MODE_OWNER_EMAIL };
  } catch (e1) {
    try {
      MailApp.sendEmail({ to: TEST_MODE_OWNER_EMAIL, subject: subject, body: body, name: 'SMS School App' });
      return { success: true, message: 'OTP sent to ' + TEST_MODE_OWNER_EMAIL };
    } catch (e2) {
      // Clean up stored OTP on failure
      props.deleteProperty('test_mode_otp');
      props.deleteProperty('test_mode_otp_ts');
      return { success: false, message: 'Failed to send OTP: ' + e2.message };
    }
  }
}

/**
 * Verifies the OTP code to confirm the owner email before setting the deadline.
 * Valid for 15 minutes to allow setting the deadline.
 */
function confirmTestModeOwnerEmail(providedCode) {
  requireLogin();
  const userEmail = getLoggedInUserEmail_();
  if (userEmail.toLowerCase() !== TEST_MODE_OWNER_EMAIL.toLowerCase()) {
    return { success: false, message: 'Only the system owner can confirm this email.' };
  }

  const props = PropertiesService.getScriptProperties();
  const storedCode = props.getProperty('test_mode_otp');
  const storedTs   = props.getProperty('test_mode_otp_ts');

  if (!storedCode || !storedTs) {
    return { success: false, message: 'No confirmation code found. Please request a new one.' };
  }

  // OTP expires in 10 minutes
  if (Date.now() - parseInt(storedTs, 10) > 600000) {
    props.deleteProperty('test_mode_otp');
    props.deleteProperty('test_mode_otp_ts');
    return { success: false, message: 'Confirmation code has expired. Please request a new one.' };
  }

  if (String(providedCode).trim() !== String(storedCode).trim()) {
    return { success: false, message: 'Incorrect confirmation code. Please check and try again.' };
  }

  // Code is valid! Mark email as confirmed (valid for 15 minutes to configure deadline)
  props.deleteProperty('test_mode_otp');
  props.deleteProperty('test_mode_otp_ts');
  props.setProperty('test_mode_owner_confirmed', String(Date.now()));

  return { success: true, message: 'Email confirmed successfully! You can now set the deadline.' };
}

/**
 * Activates Test Mode after the email has been confirmed.
 * Accepts durationMs OR customDeadlineIso.
 */
function activateTestModeWithDeadline(durationMs, customDeadlineIso) {
  requireLogin();
  const userEmail = getLoggedInUserEmail_();
  if (userEmail.toLowerCase() !== TEST_MODE_OWNER_EMAIL.toLowerCase()) {
    return { success: false, message: 'Only the system owner can activate Test Mode.' };
  }

  const props = PropertiesService.getScriptProperties();
  const confirmedTs = props.getProperty('test_mode_owner_confirmed');
  if (!confirmedTs || (Date.now() - parseInt(confirmedTs, 10) > 900000)) {
    return { success: false, message: 'Confirmation expired. Please verify your email again.' };
  }

  let deadlineMs = 0;
  if (customDeadlineIso) {
    const parsed = new Date(customDeadlineIso).getTime();
    if (!isNaN(parsed) && parsed > Date.now()) {
      deadlineMs = parsed;
    }
  }

  if (!deadlineMs && durationMs) {
    deadlineMs = Date.now() + Number(durationMs);
  }

  if (!deadlineMs || deadlineMs <= Date.now()) {
    deadlineMs = Date.now() + 86400000; // Default 24 hours
  }

  props.deleteProperty('test_mode_owner_confirmed');
  props.setProperty('test_mode_active', 'true');
  props.setProperty('test_mode_deadline', String(deadlineMs));

  return {
    success: true,
    message: 'Test Mode activated!',
    deadline: new Date(deadlineMs).toISOString()
  };
}

/**
 * Activates Test Mode after verifying the OTP and setting a deadline (all-in-one fallback).
 * @param {string} providedCode   - The 6-digit OTP the user entered.
 * @param {number} durationMs     - Duration in milliseconds (e.g. 3600000 = 1 hour).
 */
function activateTestMode(providedCode, durationMs) {
  requireLogin();
  const userEmail = getLoggedInUserEmail_();
  if (userEmail.toLowerCase() !== TEST_MODE_OWNER_EMAIL.toLowerCase()) {
    return { success: false, message: 'Only the system owner can activate Test Mode.' };
  }

  const props = PropertiesService.getScriptProperties();
  const storedCode = props.getProperty('test_mode_otp');
  const storedTs   = props.getProperty('test_mode_otp_ts');

  if (!storedCode || !storedTs) {
    return { success: false, message: 'No OTP found. Please request a new one.' };
  }

  // OTP expires in 10 minutes
  if (Date.now() - parseInt(storedTs, 10) > 600000) {
    props.deleteProperty('test_mode_otp');
    props.deleteProperty('test_mode_otp_ts');
    return { success: false, message: 'OTP has expired. Please request a new one.' };
  }

  if (String(providedCode).trim() !== String(storedCode)) {
    return { success: false, message: 'Incorrect OTP. Please try again.' };
  }

  // OTP is valid — activate Test Mode
  props.deleteProperty('test_mode_otp');
  props.deleteProperty('test_mode_otp_ts');

  const deadlineMs = Date.now() + Number(durationMs);
  props.setProperty('test_mode_active', 'true');
  props.setProperty('test_mode_deadline', String(deadlineMs));

  return {
    success: true,
    message: 'Test Mode activated!',
    deadline: new Date(deadlineMs).toISOString()
  };
}

/**
 * Deactivates Test Mode immediately (owner only).
 */
function deactivateTestMode() {
  requireLogin();
  const userEmail = getLoggedInUserEmail_();
  if (userEmail.toLowerCase() !== TEST_MODE_OWNER_EMAIL.toLowerCase()) {
    return { success: false, message: 'Only the system owner can deactivate Test Mode.' };
  }

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('test_mode_active');
  props.deleteProperty('test_mode_deadline');
  props.deleteProperty('test_mode_otp');
  props.deleteProperty('test_mode_otp_ts');

  return { success: true, message: 'Test Mode deactivated.' };
}

/**
 * Returns the current Test Mode status — safe to call from the client.
 * Also returns whether the calling user is the owner (to show/hide the UI toggle).
 */
function getTestModeStatus() {
  requireLogin();
  const userEmail = getLoggedInUserEmail_();
  const isOwner = userEmail.toLowerCase() === TEST_MODE_OWNER_EMAIL.toLowerCase();
  const testMode = isTestModeActive();
  return {
    isOwner: isOwner,
    active: testMode.active,
    deadline: testMode.deadline
  };
}

/**
 * Internal helper: resolves the googleEmail for the currently logged-in user
 * by looking up their userId in the Users sheet.
 * @returns {string} lowercase email, or '' if not found.
 */
function getLoggedInUserEmail_() {
  try {
    const user = getLoggedInUser();
    if (!user || !user.userId) return '';

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return '';

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return '';

    // Users sheet: col B = userId, col C = googleEmail
    const data = sheet.getRange(3, 2, lastRow - 2, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(user.userId).trim()) {
        return String(data[i][1]).trim().toLowerCase();
      }
    }
    return '';
  } catch (e) {
    return '';
  }
}

function activateFullLicense(providedPassword) {
  requireLogin();
  const masterPass = getSystemParameter('Master Password');
  if (!verifyPassword(providedPassword, masterPass || '')) {
    return { success: false, message: 'Invalid developer master password.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) return { success: false, message: 'Settings sheet not found.' };

  const data = settingsSheet.getDataRange().getValues();
  const norm = s => String(s || '').trim().toLowerCase().replace(/[\s_\-]+/g, '');

  let statusRow = -1;
  let statusCol = 3; // Default value column: C (3)

  for (let i = 0; i < data.length; i++) {
    const colAKey = norm(data[i][0]);
    const colBKey = norm(data[i][1]);
    if (colAKey === 'licensestatus') {
      statusRow = i + 1;
      statusCol = 2;
      break;
    } else if (colBKey === 'licensestatus') {
      statusRow = i + 1;
      statusCol = 3;
      break;
    }
  }

  if (statusRow === -1) {
    statusRow = 9;
    statusCol = 3;
    if (!settingsSheet.getRange(9, 2).getValue()) {
      settingsSheet.getRange(9, 2).setValue('License Status');
    }
  }

  settingsSheet.getRange(statusRow, statusCol).setValue('Full Access');
  return { success: true, message: 'Full lifetime access activated successfully!' };
}

function verifyMasterPassword(providedPassword) {
  requireLogin();
  const masterPass = getSystemParameter('Master Password');
  if (!verifyPassword(providedPassword, masterPass || '')) {
    return { success: false, message: 'Invalid master password.' };
  }
  const license = getLicenseDetails();
  return { success: true, license: license };
}

function updateLicenseSettings(status, limit, providedPassword) {
  requireLogin();
  
  // If password is empty, it means verification was done via email code
  if (providedPassword && providedPassword.trim() !== '') {
    const masterPass = getSystemParameter('Master Password');
    if (!verifyPassword(providedPassword, masterPass || '')) {
      return { success: false, message: 'Invalid master password.' };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) return { success: false, message: 'Settings sheet not found.' };

  const data = settingsSheet.getDataRange().getValues();
  const norm = s => String(s || '').trim().toLowerCase().replace(/[\s_\-]+/g, '');

  let statusRow = -1;
  let statusValCol = 3; // Default: Cell C9 (Column 3)

  let limitRow = -1;
  let limitValCol = 3; // Default: Cell C10 (Column 3)

  for (let i = 0; i < data.length; i++) {
    const colAKey = norm(data[i][0]);
    const colBKey = norm(data[i][1]);

    if (colAKey === 'licensestatus') {
      statusRow = i + 1;
      statusValCol = 2;
    } else if (colBKey === 'licensestatus') {
      statusRow = i + 1;
      statusValCol = 3;
    }

    if (colAKey === 'studentlimit') {
      limitRow = i + 1;
      limitValCol = 2;
    } else if (colBKey === 'studentlimit') {
      limitRow = i + 1;
      limitValCol = 3;
    }
  }

  // Explicitly target Row 9 for License Status and Row 10 for Student Limit if not found elsewhere
  if (statusRow === -1) {
    statusRow = 9;
    statusValCol = 3;
  }
  if (limitRow === -1) {
    limitRow = 10;
    limitValCol = 3;
  }

  // Ensure parameter names are present in Column B if missing
  if (!settingsSheet.getRange(statusRow, 2).getValue() && !settingsSheet.getRange(statusRow, 1).getValue()) {
    settingsSheet.getRange(statusRow, 2).setValue('License Status');
  }
  if (!settingsSheet.getRange(limitRow, 2).getValue() && !settingsSheet.getRange(limitRow, 1).getValue()) {
    settingsSheet.getRange(limitRow, 2).setValue('Student Limit');
  }

  // Write values into cells C9 (Row 9, Column C) and C10 (Row 10, Column C)
  settingsSheet.getRange(statusRow, statusValCol).setValue(status);
  settingsSheet.getRange(limitRow, limitValCol).setValue(limit);

  return { success: true, message: 'License settings updated successfully.' };
}

// --- EMAIL VERIFICATION FOR LICENSE EDITS ---
function generateAndEmailVerificationCode() {
  requireLogin();
  const devEmail = (getSystemParameter('Developer Email') || '').trim();
  if (!devEmail) {
    return { success: false, message: 'Developer Email is not configured in settings. Please set it first.' };
  }

  // Generate a 6-digit random code
  const code = String(Math.floor(100000 + Math.random() * 900000));
  const timestamp = Date.now();
  
  // Store verification code in PropertiesService (expires in 10 minutes)
  const props = PropertiesService.getUserProperties();
  props.setProperty('license_verify_code', code);
  props.setProperty('license_verify_timestamp', String(timestamp));

  try {
    const subject = '🔐 SMS License Settings Verification Code';
    const plainBody =
      'Your verification code is: ' + code + '\n\n' +
      'This code expires in 10 minutes.\n' +
      'Use this code in the SMS application to update License Status and Student Limit.\n\n' +
      'If you did not request this, please ignore this email.';

    // Use GmailApp to send only to the developer email (no automatic CC to script owner)
    GmailApp.sendEmail(devEmail, subject, plainBody, {
      name: 'SMS School App',
      noReply: true
    });

    return { success: true, message: 'Verification code sent to ' + devEmail };
  } catch (e1) {
    // Fallback to MailApp if GmailApp fails (different account types)
    try {
      MailApp.sendEmail({
        to: devEmail,
        subject: '🔐 SMS License Settings Verification Code',
        body:
          'Your verification code is: ' + code + '\n\n' +
          'This code expires in 10 minutes.\n' +
          'Use this code in the SMS application to update License Status and Student Limit.\n\n' +
          'If you did not request this, please ignore this email.',
        name: 'SMS School App'
      });
      return { success: true, message: 'Verification code sent to ' + devEmail };
    } catch (e2) {
      return { success: false, message: 'Failed to send email: ' + e2.message };
    }
  }
}

function verifyEmailCode(providedCode) {
  requireLogin();
  const props = PropertiesService.getUserProperties();
  const storedCode = props.getProperty('license_verify_code');
  const timestamp = props.getProperty('license_verify_timestamp');
  
  if (!storedCode || !timestamp) {
    return { success: false, message: 'No verification code found. Please request a new one.' };
  }

  // Check if code expired (10 minutes = 600000 ms)
  const now = Date.now();
  if (now - parseInt(timestamp) > 600000) {
    props.deleteProperty('license_verify_code');
    props.deleteProperty('license_verify_timestamp');
    return { success: false, message: 'Verification code has expired. Please request a new one.' };
  }

  // Verify code
  if (String(providedCode).trim() !== String(storedCode)) {
    return { success: false, message: 'Incorrect verification code.' };
  }

  // Code is valid - return license details
  const license = getLicenseDetails();
  
  // Clear the code after successful verification (one-time use)
  props.deleteProperty('license_verify_code');
  props.deleteProperty('license_verify_timestamp');
  
  return { success: true, license: license };
}

// 1. Serve the Student Page
function showStudentPage() {
  return HtmlService.createTemplateFromFile("Student")
    .evaluate()
    .setTitle("Student Directory")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. Fetch Data from the 'Students' sheet
// 2. Fetch Data from the 'Students' sheet - UPDATED FOR FIX
// ...existing code...
/**
 * Return an array of student objects based on the "Students" sheet.
 * Make sure the sheet name and header names match what's in the sheet.
 */
function getStudentsData() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Students");
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  // Find the first non-empty row to use as the header
  // (handles sheets where row 1 is blank and headers are on row 2)
  var headerRowIndex = -1;
  for (var i = 0; i < values.length; i++) {
    if (
      values[i].some(function (c) {
        return c !== "" && c !== null;
      })
    ) {
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

  // Data rows come after the header row; skip any fully blank rows
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

  Logger.log("getStudentsData -> rows: " + out.length);
  return out;
}

function findDuplicateStudent(studentData, ignoreStudentId) {
  const students = getStudentsData();
  const firstName = String(studentData.firstName || "")
    .trim()
    .toLowerCase();
  const lastName = String(studentData.lastName || "")
    .trim()
    .toLowerCase();

  for (let i = 0; i < students.length; i++) {
    const student = students[i];
    if (ignoreStudentId && student.studentId === ignoreStudentId) continue;

    const existingFirst = String(student.firstName || "")
      .trim()
      .toLowerCase();
    const existingLast = String(student.lastName || "")
      .trim()
      .toLowerCase();

    if (
      firstName &&
      lastName &&
      existingFirst === firstName &&
      existingLast === lastName
    ) {
      return { type: "name", studentId: student.studentId };
    }
  }
  return null;
}

/**
 * Return an array of class names from the "Classes" sheet (first column).
 * NOTE: getClassesData() is defined later in the file (keeping one definition)
 */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 3. Helper to locate student row by ID - OPTIMIZED WITH INDEX
function findRowById(sheet, studentId) {
  // Use indexed lookup for O(1) performance instead of O(n) linear search
  const rowNumber = findRowByIdIndexed('Students', studentId);
  
  if (rowNumber !== -1) {
    return rowNumber;
  }
  
  // Fallback to linear search if index fails (shouldn't happen)
  Logger.log('WARNING: Index lookup failed, falling back to linear search');
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return -1;
  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(studentId).trim()) {
      return i + 3;
    }
  }
  return -1;
}

// 4. Helper to auto-generate the next incrementing Student ID (e.g. STU-1001)
function generateNextStudentId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "STU-1001";

  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues(); // Column B
  let maxIdNum = 1000;
  values.forEach((row) => {
    const idStr = String(row[0]).trim();
    if (idStr.startsWith("STU-")) {
      const num = parseInt(idStr.substring(4), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return "STU-" + (maxIdNum + 1);
}

// 5. Create operation logic
function addStudent(studentData) {
  requireLogin();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  if (!sheet) throw new Error("Students worksheet not found.");

  const license = getLicenseDetails();
  const currentStudents = getStudentsData();
  if (currentStudents.length >= license.limit) {
    return {
      success: false,
      licenseLimitReached: true,
      limit: license.limit,
      currentCount: currentStudents.length,
      developerEmail: license.developerEmail,
      message:
        'Student Limit Reached! You have reached the maximum allowance of ' +
        license.limit +
        ' students.',
    };
  }

  const duplicate = findDuplicateStudent(studentData);
  if (duplicate) {
    return {
      success: false,
      message:
        "A student with the same first name and last name already exists.",
    };
  }

  const nextId = generateNextStudentId(sheet);
  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  // Set values starting from Column 2 (B) leaving Column A empty
  // Include Class as the 9th column (Column J)
  sheet.getRange(targetRow, 2, 1, 9).setValues([
    [
      nextId,
      studentData.firstName,
      studentData.lastName,
      studentData.email,
      studentData.dob ? new Date(studentData.dob) : "",
      studentData.gender,
      studentData.enrollmentDate
        ? new Date(studentData.enrollmentDate)
        : new Date(),
      studentData.status || "Active",
      studentData.class || "",
    ],
  ]);

  // Invalidate cache and index after modification
  invalidateStudentsCache();
  invalidateIndex('Students');

  safeLogAuditEvent(
    'Create',
    'Students',
    nextId,
    'Created student record',
    null,
    buildAuditSnapshot({
      studentId: nextId,
      firstName: studentData.firstName,
      lastName: studentData.lastName,
      email: studentData.email,
      dob: studentData.dob || '',
      gender: studentData.gender || '',
      enrollmentDate: studentData.enrollmentDate || '',
      status: studentData.status || 'Active',
      class: studentData.class || ''
    })
  );

  return { success: true, studentId: nextId };
}

// 6. Update operation logic
/**
 * Update a student without replacing fields that were not supplied.
 *
 * The old students page called this function with positional arguments while
 * the current page sends one data object.  Treating the old call as an object
 * caused every property to be undefined and the subsequent setValues() call
 * blanked the student's name, email, dates, gender, and class.  Supporting the
 * legacy shape here and merging against the existing row makes the server side
 * operation safe even when an older deployed page is still open.
 */
function updateStudent(studentId, studentData, legacyLastName, legacyEmail, legacyDob, legacyGender, legacyEnrollmentDate, legacyStatus, legacyClass) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  if (!sheet) throw new Error("Students worksheet not found.");

  const isObjectPayload = studentData !== null &&
    typeof studentData === "object" &&
    !Array.isArray(studentData);

  if (!isObjectPayload) {
    // Backward compatibility for older Student.html deployments.  If an
    // invalid request contains no legacy fields, reject it rather than writing
    // undefined values to the sheet.
    if (arguments.length < 3) {
      return { success: false, message: "Student update data is invalid." };
    }
    studentData = {
      firstName: studentData,
      lastName: legacyLastName,
      email: legacyEmail,
      dob: legacyDob,
      gender: legacyGender,
      enrollmentDate: legacyEnrollmentDate,
      status: legacyStatus,
      class: legacyClass
    };
  }

  const row = findRowById(sheet, studentId);
  if (row === -1) throw new Error("Student record not found.");

  const currentRow = sheet.getRange(row, 1, 1, 10).getValues()[0];
  const hasField = function(field) {
    return Object.prototype.hasOwnProperty.call(studentData, field) &&
      studentData[field] !== undefined;
  };
  const suppliedOrCurrent = function(field, currentValue) {
    return hasField(field) ? studentData[field] : currentValue;
  };
  const textValue = function(value) {
    return value === null || value === undefined ? "" : String(value);
  };
  const dateValue = function(value, fallback) {
    if (value === null || value === undefined || value === "") return value === "" ? "" : fallback;
    if (value instanceof Date) {
      return isNaN(value.getTime()) ? fallback : value;
    }
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? fallback : parsed;
  };

  // Merge with the row currently in the sheet.  This is deliberately done
  // before writing so a partial/legacy request can never erase unrelated data.
  const mergedStudentData = {
    firstName: textValue(suppliedOrCurrent("firstName", currentRow[2])),
    lastName: textValue(suppliedOrCurrent("lastName", currentRow[3])),
    email: textValue(suppliedOrCurrent("email", currentRow[4])),
    dob: dateValue(suppliedOrCurrent("dob", currentRow[5]), currentRow[5]),
    gender: textValue(suppliedOrCurrent("gender", currentRow[6])),
    enrollmentDate: dateValue(suppliedOrCurrent("enrollmentDate", currentRow[7]), currentRow[7]),
    status: textValue(suppliedOrCurrent("status", currentRow[8])),
    class: textValue(suppliedOrCurrent("class", currentRow[9]))
  };

  const duplicate = findDuplicateStudent(mergedStudentData, studentId);
  if (duplicate) {
    return {
      success: false,
      message:
        "Another student with the same first name and last name already exists.",
    };
  }

  // Check if student is Completed JHS 3 (locked record)
  const currentStatus = textValue(currentRow[8]).trim();
  const currentClass = textValue(currentRow[9]).trim();
  if (currentStatus.toLowerCase() === "completed" && currentClass.toUpperCase().includes("JHS 3")) {
    return {
      success: false,
      message: "Cannot update this student: Completed JHS 3 record is locked.",
    };
  }

  const oldSnapshot = buildAuditSnapshot({
    studentId: textValue(currentRow[1]).trim(),
    firstName: textValue(currentRow[2]).trim(),
    lastName: textValue(currentRow[3]).trim(),
    email: textValue(currentRow[4]).trim(),
    dob: currentRow[5] instanceof Date ? currentRow[5].toISOString() : textValue(currentRow[5]).trim(),
    gender: textValue(currentRow[6]).trim(),
    enrollmentDate: currentRow[7] instanceof Date ? currentRow[7].toISOString() : textValue(currentRow[7]).trim(),
    status: currentStatus,
    class: currentClass
  });

  // Update only the student detail columns (C:J).  Every value is either from
  // the request or the existing row; undefined can never be sent to Sheets.
  sheet.getRange(row, 3, 1, 8).setValues([[
    mergedStudentData.firstName,
    mergedStudentData.lastName,
    mergedStudentData.email,
    mergedStudentData.dob,
    mergedStudentData.gender,
    mergedStudentData.enrollmentDate,
    mergedStudentData.status || "Active",
    mergedStudentData.class
  ]]);

  // Invalidate cache and index after modification
  invalidateStudentsCache();
  invalidateIndex('Students');

  safeLogAuditEvent(
    'Update',
    'Students',
    studentId,
    'Updated student record',
    oldSnapshot,
    buildAuditSnapshot({
      studentId: studentId,
      firstName: mergedStudentData.firstName,
      lastName: mergedStudentData.lastName,
      email: mergedStudentData.email,
      dob: mergedStudentData.dob || '',
      gender: mergedStudentData.gender,
      enrollmentDate: mergedStudentData.enrollmentDate || '',
      status: mergedStudentData.status || 'Active',
      class: mergedStudentData.class
    })
  );

  return { success: true };
}

/**
 * Update the status and/or class for multiple students in one request.
 * Only these two fields are accepted so other student details cannot be
 * accidentally overwritten by a bulk quick action.
 */
function bulkUpdateStudents(studentIds, updates) {
  requireLogin();

  if (!Array.isArray(studentIds) || studentIds.length === 0) {
    return { success: false, message: "Select at least one student to update." };
  }
  if (studentIds.length > 1000) {
    return { success: false, message: "A maximum of 1000 students can be updated at once." };
  }

  updates = updates || {};
  const hasStatus = Object.prototype.hasOwnProperty.call(updates, "status");
  const hasClass = Object.prototype.hasOwnProperty.call(updates, "class");
  if (!hasStatus && !hasClass) {
    return { success: false, message: "Choose a status or class to update." };
  }

  let newStatus = "";
  if (hasStatus) {
    newStatus = String(updates.status || "").trim();
    const validStatuses = ["Active", "Inactive", "Completed"];
    if (validStatuses.indexOf(newStatus) === -1) {
      return { success: false, message: "The selected student status is invalid." };
    }
  }

  let newClass = "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (hasClass) {
    newClass = String(updates.class || "").trim();
    if (!newClass) {
      return { success: false, message: "Choose a class to update the selected students." };
    }

    const classesSheet = ss.getSheetByName("Classes");
    if (!classesSheet) {
      return { success: false, message: "Classes worksheet not found." };
    }

    const classNames = classesSheet.getLastRow() < 2
      ? []
      : classesSheet.getRange(2, 1, classesSheet.getLastRow() - 1, 1).getValues()
          .map(function(row) { return String(row[0] || "").trim(); })
          .filter(function(className) { return className !== ""; });
    const selectedClass = classNames.find(function(className) {
      return className.toLowerCase() === newClass.toLowerCase();
    });

    if (!selectedClass) {
      return { success: false, message: "The selected class no longer exists. Refresh the page and try again." };
    }
    newClass = selectedClass;
  }

  const sheet = ss.getSheetByName("Students");
  if (!sheet) {
    return { success: false, message: "Students worksheet not found." };
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const values = sheet.getDataRange().getValues();
    let headerRowIndex = -1;
    for (let i = 0; i < values.length; i++) {
      if (values[i].some(function(value) { return value !== "" && value !== null; })) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      return { success: false, message: "The Students worksheet is empty." };
    }

    const normalizedHeaders = values[headerRowIndex].map(function(header) {
      return String(header || "").trim().toLowerCase();
    });
    const studentIdColumn = normalizedHeaders.indexOf("student id");
    const statusColumn = normalizedHeaders.indexOf("status");
    let classColumn = normalizedHeaders.indexOf("class");
    if (classColumn === -1) classColumn = normalizedHeaders.indexOf("class name");

    if (studentIdColumn === -1 || statusColumn === -1 || classColumn === -1) {
      return {
        success: false,
        message: "Students worksheet must contain Student ID, Status, and Class columns."
      };
    }

    const rowsByStudentId = {};
    for (let rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex++) {
      const id = String(values[rowIndex][studentIdColumn] || "").trim();
      if (id) rowsByStudentId["$" + id] = rowIndex;
    }

    const uniqueStudentIds = [];
    const seenStudentIds = {};
    studentIds.forEach(function(studentId) {
      const id = String(studentId || "").trim();
      if (id && !seenStudentIds["$" + id]) {
        seenStudentIds["$" + id] = true;
        uniqueStudentIds.push(id);
      }
    });

    const statusRanges = [];
    const classRanges = [];
    const auditEntries = [];
    const errors = [];
    let unchangedCount = 0;

    uniqueStudentIds.forEach(function(studentId) {
      const rowIndex = rowsByStudentId["$" + studentId];
      if (rowIndex === undefined) {
        errors.push(studentId + ": student record not found.");
        return;
      }

      const currentStatus = String(values[rowIndex][statusColumn] || "").trim();
      const currentClass = String(values[rowIndex][classColumn] || "").trim();
      const isLocked = currentStatus.toLowerCase() === "completed" &&
        currentClass.toUpperCase().indexOf("JHS 3") !== -1;
      if (isLocked) {
        errors.push(studentId + ": Completed JHS 3 record is locked.");
        return;
      }

      const statusChanged = hasStatus && currentStatus !== newStatus;
      const classChanged = hasClass && currentClass !== newClass;
      if (!statusChanged && !classChanged) {
        unchangedCount++;
        return;
      }

      const sheetRow = rowIndex + 1;
      if (statusChanged) {
        statusRanges.push(sheet.getRange(sheetRow, statusColumn + 1).getA1Notation());
      }
      if (classChanged) {
        classRanges.push(sheet.getRange(sheetRow, classColumn + 1).getA1Notation());
      }

      auditEntries.push({
        studentId: studentId,
        oldStatus: currentStatus,
        oldClass: currentClass,
        newStatus: hasStatus ? newStatus : currentStatus,
        newClass: hasClass ? newClass : currentClass
      });
    });

    if (statusRanges.length > 0) sheet.getRangeList(statusRanges).setValue(newStatus);
    if (classRanges.length > 0) sheet.getRangeList(classRanges).setValue(newClass);

    if (auditEntries.length > 0) {
      invalidateStudentsCache();
      invalidateIndex("Students");

      const detail = hasStatus && hasClass
        ? "Bulk updated student status and class"
        : (hasStatus ? "Bulk updated student status" : "Bulk updated student class");
      auditEntries.forEach(function(entry) {
        safeLogAuditEvent(
          "Update",
          "Students",
          entry.studentId,
          detail,
          buildAuditSnapshot({
            studentId: entry.studentId,
            status: entry.oldStatus,
            class: entry.oldClass
          }),
          buildAuditSnapshot({
            studentId: entry.studentId,
            status: entry.newStatus,
            class: entry.newClass
          })
        );
      });
    }

    return {
      success: true,
      updatedCount: auditEntries.length,
      unchangedCount: unchangedCount,
      failedCount: errors.length,
      errors: errors
    };
  } finally {
    lock.releaseLock();
  }
}

// 7. Delete operation logic
/**
 * Check if a student has related records in other sheets
 * Returns an object with hasRecords (boolean) and details (array of sheet names)
 */
function checkStudentRelatedRecords(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const relatedSheets = [];
  const normalizedStudentId = String(studentId).trim();
  
  // Check Payments sheet (column C - Student ID)
  const paymentsSheet = ss.getSheetByName("Payments");
  if (paymentsSheet && paymentsSheet.getLastRow() > 2) {
    const paymentsData = paymentsSheet.getRange(3, 3, paymentsSheet.getLastRow() - 2, 1).getValues();
    for (let i = 0; i < paymentsData.length; i++) {
      if (String(paymentsData[i][0]).trim() === normalizedStudentId) {
        relatedSheets.push("Payments");
        break;
      }
    }
  }
  
  // Check Invoices sheet (column C - Student ID)
  const invoicesSheet = ss.getSheetByName("Invoices");
  if (invoicesSheet && invoicesSheet.getLastRow() > 2) {
    const invoicesData = invoicesSheet.getRange(3, 3, invoicesSheet.getLastRow() - 2, 1).getValues();
    for (let i = 0; i < invoicesData.length; i++) {
      if (String(invoicesData[i][0]).trim() === normalizedStudentId) {
        relatedSheets.push("Invoices");
        break;
      }
    }
  }
  
  // Check Attendance sheet (column C - Student ID)
  const attendanceSheet = ss.getSheetByName("Attendance");
  if (attendanceSheet && attendanceSheet.getLastRow() > 2) {
    const attendanceData = attendanceSheet.getRange(3, 3, attendanceSheet.getLastRow() - 2, 1).getValues();
    for (let i = 0; i < attendanceData.length; i++) {
      if (String(attendanceData[i][0]).trim() === normalizedStudentId) {
        relatedSheets.push("Attendance");
        break;
      }
    }
  }
  
  // Check Enrollments sheet (column C - Student ID)
  const enrollmentsSheet = ss.getSheetByName("Enrollments");
  if (enrollmentsSheet && enrollmentsSheet.getLastRow() > 2) {
    const enrollmentsData = enrollmentsSheet.getRange(3, 3, enrollmentsSheet.getLastRow() - 2, 1).getValues();
    for (let i = 0; i < enrollmentsData.length; i++) {
      if (String(enrollmentsData[i][0]).trim() === normalizedStudentId) {
        relatedSheets.push("Enrollments");
        break;
      }
    }
  }
  
  // Check Performance sheet (column C - Student ID)
  const performanceSheet = ss.getSheetByName("Performance");
  if (performanceSheet && performanceSheet.getLastRow() > 2) {
    const performanceData = performanceSheet.getRange(3, 3, performanceSheet.getLastRow() - 2, 1).getValues();
    for (let i = 0; i < performanceData.length; i++) {
      if (String(performanceData[i][0]).trim() === normalizedStudentId) {
        relatedSheets.push("Performance");
        break;
      }
    }
  }
  
  // Check Billings sheet (column C - Student ID)
  const billingsSheet = ss.getSheetByName("Billings");
  if (billingsSheet && billingsSheet.getLastRow() > 2) {
    const billingsData = billingsSheet.getRange(3, 3, billingsSheet.getLastRow() - 2, 1).getValues();
    for (let i = 0; i < billingsData.length; i++) {
      if (String(billingsData[i][0]).trim() === normalizedStudentId) {
        relatedSheets.push("Billings");
        break;
      }
    }
  }
  
  return {
    hasRecords: relatedSheets.length > 0,
    relatedSheets: relatedSheets
  };
}

function deleteStudent(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  if (!sheet) throw new Error("Students worksheet not found.");

  const row = findRowById(sheet, studentId);
  if (row === -1) throw new Error("Student record not found.");

  // Check for related records before deletion
  const relatedCheck = checkStudentRelatedRecords(studentId);
  if (relatedCheck.hasRecords) {
    const sheetList = relatedCheck.relatedSheets.join(", ");
    throw new Error(
      "Cannot delete student " + studentId + ". This student has related records in: " + sheetList + ". " +
      "Please delete or reassign these records first to maintain data integrity."
    );
  }

  const currentRow = sheet.getRange(row, 1, 1, 10).getValues()[0];
  const deletedSnapshot = buildAuditSnapshot({
    studentId: String(currentRow[1] || '').trim(),
    firstName: String(currentRow[2] || '').trim(),
    lastName: String(currentRow[3] || '').trim(),
    email: String(currentRow[4] || '').trim(),
    dob: currentRow[5] instanceof Date ? currentRow[5].toISOString() : String(currentRow[5] || '').trim(),
    gender: String(currentRow[6] || '').trim(),
    enrollmentDate: currentRow[7] instanceof Date ? currentRow[7].toISOString() : String(currentRow[7] || '').trim(),
    status: String(currentRow[8] || '').trim(),
    class: String(currentRow[9] || '').trim()
  });

  sheet.deleteRow(row);
  
  // Invalidate cache and index after modification
  invalidateStudentsCache();
  invalidateIndex('Students');

  safeLogAuditEvent(
    'Delete',
    'Students',
    studentId,
    'Deleted student record',
    deletedSnapshot,
    null
  );
  
  return { success: true };
}

// 8. Import Multiple Students from CSV Data
function importStudents(studentsArray) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Students");
    
    if (!sheet) {
      return { 
        success: false, 
        message: "Students worksheet not found." 
      };
    }

    const license = getLicenseDetails();
    const currentStudents = getStudentsData();
    let currentCount = currentStudents.length;

    if (currentCount >= license.limit) {
      return {
        success: false,
        licenseLimitReached: true,
        limit: license.limit,
        currentCount: currentCount,
        developerEmail: license.developerEmail,
        message: 'Student Limit Reached! Cannot import students because you have reached the maximum allowance of ' + license.limit + ' students.'
      };
    }

    let imported = 0;
    let failed = 0;
    let duplicates = 0;
    let limitHitDuringImport = false;
    const errors = [];

    for (let index = 0; index < studentsArray.length; index++) {
      const studentData = studentsArray[index];
      try {
        if (currentCount + imported >= license.limit) {
          limitHitDuringImport = true;
          errors.push(`Row ${studentData.rowNumber}: Student limit of ${license.limit} reached. Remaining students were skipped.`);
          break;
        }

        // Validate required fields
        if (!studentData.firstName || !studentData.lastName || !studentData.email) {
          errors.push(`Row ${studentData.rowNumber}: Missing required fields (First Name, Last Name, or Email)`);
          failed++;
          continue;
        }

        // Check for duplicates (same first name and last name)
        const duplicate = findDuplicateStudent(studentData, null);
        if (duplicate) {
          errors.push(`Row ${studentData.rowNumber}: Duplicate student - ${studentData.firstName} ${studentData.lastName} already exists`);
          duplicates++;
          continue;
        }

        // Validate email format (basic check)
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(studentData.email)) {
          errors.push(`Row ${studentData.rowNumber}: Invalid email format - ${studentData.email}`);
          failed++;
          continue;
        }

        // Validate date formats (YYYY-MM-DD)
        const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
        if (studentData.dob && !dateRegex.test(studentData.dob)) {
          errors.push(`Row ${studentData.rowNumber}: Invalid date of birth format - ${studentData.dob} (expected YYYY-MM-DD)`);
          failed++;
          continue;
        }
        if (studentData.enrollmentDate && !dateRegex.test(studentData.enrollmentDate)) {
          errors.push(`Row ${studentData.rowNumber}: Invalid enrollment date format - ${studentData.enrollmentDate} (expected YYYY-MM-DD)`);
          failed++;
          continue;
        }

        // Validate gender
        const validGenders = ['Male', 'Female', 'Other', 'male', 'female', 'other'];
        if (studentData.gender && !validGenders.includes(studentData.gender)) {
          errors.push(`Row ${studentData.rowNumber}: Invalid gender - ${studentData.gender} (expected Male, Female, or Other)`);
          failed++;
          continue;
        }

        // Validate status
        const validStatuses = ['Active', 'Inactive', 'active', 'inactive'];
        if (studentData.status && !validStatuses.includes(studentData.status)) {
          errors.push(`Row ${studentData.rowNumber}: Invalid status - ${studentData.status} (expected Active or Inactive)`);
          failed++;
          continue;
        }

        // Generate new student ID
        const studentId = generateNextStudentId(sheet);

        // Convert date strings to Date objects
        const dobDate = studentData.dob ? new Date(studentData.dob) : '';
        const enrollDate = studentData.enrollmentDate ? new Date(studentData.enrollmentDate) : new Date();

        // Prepare row data (matching your Students sheet structure)
        const newRow = [
          '', // Column A (blank)
          studentId,
          studentData.firstName.trim(),
          studentData.lastName.trim(),
          studentData.email.trim().toLowerCase(),
          dobDate,
          studentData.gender || '',
          enrollDate,
          studentData.status || 'Active',
          studentData.class || ''
        ];

        // Append to sheet
        sheet.appendRow(newRow);
        imported++;

      } catch (err) {
        errors.push(`Row ${studentData.rowNumber}: ${err.message}`);
        failed++;
      }
    }

    if (imported > 0) {
      invalidateStudentsCache();
      invalidateIndex('Students');
      safeLogAuditEvent(
        'Import',
        'Students',
        null,
        'Imported students in bulk',
        null,
        buildAuditSnapshot({
          imported: imported,
          failed: failed,
          duplicates: duplicates,
          totalRows: studentsArray.length
        })
      );
    }

    return {
      success: true,
      imported: imported,
      failed: failed,
      duplicates: duplicates,
      errors: errors,
      licenseLimitReached: limitHitDuringImport,
      limit: license.limit,
      developerEmail: license.developerEmail
    };

  } catch (err) {
    Logger.log('Import error: ' + err.message);
    return {
      success: false,
      message: err.message,
      imported: 0,
      failed: 0,
      duplicates: 0,
      errors: []
    };
  }
}

// Robust helper to format attendance dates as YYYY-MM-DD
function formatAttendanceDate(rawD, ss) {
  if (!rawD) return "";
  if (rawD instanceof Date && !isNaN(rawD.getTime())) {
    try {
      const tz = ss ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
      return Utilities.formatDate(rawD, tz, "yyyy-MM-dd");
    } catch (e) {
      const yyyy = rawD.getFullYear();
      const mm = String(rawD.getMonth() + 1).padStart(2, "0");
      const dd = String(rawD.getDate()).padStart(2, "0");
      return `${yyyy}-${mm}-${dd}`;
    }
  }
  const str = String(rawD).trim();
  if (!str) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  const parts = str.split(/[\/\-\.]/);
  if (parts.length === 3) {
    let m, d, y;
    if (parts[0].length === 4) {
      y = parts[0]; m = parts[1].padStart(2, "0"); d = parts[2].padStart(2, "0");
    } else {
      m = parts[0].padStart(2, "0"); d = parts[1].padStart(2, "0"); y = parts[2];
    }
    if (y && m && d && y.length === 4) return `${y}-${m}-${d}`;
  }
  return str;
}

// Build a Date whose calendar day in the spreadsheet's timezone matches the
// given "yyyy-MM-dd" string, so writing then reading it back (formatAttendanceDate)
// yields the same day regardless of the spreadsheet timezone. Parsing with
// new Date("yyyy-MM-dd") alone yields UTC midnight, which shifts a day backwards
// in timezones west of UTC.
function buildAttendanceSheetDate(dateString, tz) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dateString || "").trim());
  if (!m) return dateString ? new Date(dateString) : null;
  const y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10);
  const d = parseInt(m[3], 10);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return new Date(dateString);

  // Anchor at noon UTC (safe midpoint for offsets up to ±12h), then adjust so
  // the local calendar date in tz equals the picked date.
  let dt = new Date(Date.UTC(y, mo - 1, d, 12, 0, 0));
  try {
    const effectiveTz = tz || Session.getScriptTimeZone();
    const local = Utilities.formatDate(dt, effectiveTz, "yyyy-MM-dd");
    if (local !== m[0]) {
      const shiftDays = local > m[0] ? -1 : 1;
      dt = new Date(dt.getTime() + shiftDays * 86400000);
    }
  } catch (e) {
    // Keep the noon-UTC anchor if the timezone lookup fails
  }
  return dt;
}

// Helper to inspect Row 2 headers in the Attendance sheet and return column index maps (0-based offset relative to Column B)
function getAttendanceHeaderMap(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 10);
  const numCols = lastCol - 1; // from Col B (col 2)
  if (numCols < 1) return null;

  const headers = sheet
    .getRange(2, 2, 1, numCols)
    .getValues()[0]
    .map((h) => String(h || "").trim().toLowerCase());

  function findIndex(aliases, fallbackIdx) {
    // 1. Exact match pass
    for (let a = 0; a < aliases.length; a++) {
      for (let i = 0; i < headers.length; i++) {
        if (headers[i] === aliases[a]) return i;
      }
    }
    // 2. Partial match pass (skip short aliases to avoid false positives)
    for (let a = 0; a < aliases.length; a++) {
      const alias = aliases[a];
      if (alias.length < 4) continue;
      for (let i = 0; i < headers.length; i++) {
        if (headers[i].indexOf(alias) !== -1) return i;
      }
    }
    return fallbackIdx;
  }

  return {
    numCols: numCols,
    idxId: findIndex(["attendance id", "attendanceid", "id"], 0),
    idxDate: findIndex(["date", "attendance date"], 1),
    idxYear: findIndex(["academic year", "academicyear", "year"], 2),
    idxClass: findIndex(["class name", "classname", "student class", "class"], 3),
    idxStudentId: findIndex(["student id", "studentid", "student_id", "sid"], 4),
    idxStudentName: findIndex(["student name", "studentname", "full name", "name"], 5),
    idxCourse: findIndex(["course id", "courseid", "course"], 6),
    idxStatus: findIndex(["status", "attendance status"], 7),
    idxTerm: findIndex(["term"], 8),
  };
}

// 9. Fetch Attendance Data (authoritative sheet reader — used by IndexingLayer cache wrapper)
function getAttendanceDataFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  if (!sheet) return [];

  // Auto-verify and enforce Row 2 headers match canonical order
  try {
    const headerCheck = sheet.getRange(2, 2, 1, 9).getValues()[0];
    const expected = ["Attendance ID", "Date", "Academic Year", "Class", "Student ID", "Student Name", "Course ID", "Status", "Term"];
    let needsFix = false;
    for (let i = 0; i < expected.length; i++) {
      if (String(headerCheck[i] || "").trim().toLowerCase() !== expected[i].toLowerCase()) {
        needsFix = true;
        break;
      }
    }
    if (needsFix) {
      const hRange = sheet.getRange(2, 2, 1, expected.length);
      hRange.setValues([expected]);
      hRange.setFontWeight("bold");
      hRange.setBackground("#4A90E2");
      hRange.setFontColor("#FFFFFF");
    }
  } catch (e) {
    Logger.log("Error checking/fixing headers: " + e.message);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const hMap = getAttendanceHeaderMap(sheet);
  if (!hMap) return [];

  const dataRange = sheet.getRange(3, 2, lastRow - 2, hMap.numCols);
  const data = dataRange.getValues();

  const statusesList = typeof getAttendanceStatuses === "function" ? getAttendanceStatuses() : [];
  const defaultStatus = statusesList.length > 0 ? statusesList[0].value : "Present";

  return data
    .filter((row) => {
      const idVal = row[hMap.idxId] !== undefined ? String(row[hMap.idxId] || "").trim() : "";
      const stuIdVal = row[hMap.idxStudentId] !== undefined ? String(row[hMap.idxStudentId] || "").trim() : "";
      const nameVal = row[hMap.idxStudentName] !== undefined ? String(row[hMap.idxStudentName] || "").trim() : "";
      const classVal = row[hMap.idxClass] !== undefined ? String(row[hMap.idxClass] || "").trim() : "";
      return idVal !== "" || stuIdVal !== "" || nameVal !== "" || classVal !== "";
    })
    .map((row) => {
      const getVal = (idx) => {
        return idx >= 0 && idx < row.length ? String(row[idx] || "").trim() : "";
      };

      const rawId = getVal(hMap.idxId);
      let dateVal = hMap.idxDate >= 0 && hMap.idxDate < row.length ? formatAttendanceDate(row[hMap.idxDate], ss) : "";
      const rawYear = getVal(hMap.idxYear);
      const rawClass = getVal(hMap.idxClass);
      const rawStuId = getVal(hMap.idxStudentId);
      const rawStuName = getVal(hMap.idxStudentName);
      const rawCourse = getVal(hMap.idxCourse);
      const rawStatus = getVal(hMap.idxStatus);
      const rawTerm = getVal(hMap.idxTerm);

      let finalStatus = rawStatus;
      if (!finalStatus || finalStatus.toLowerCase() === "undefined" || finalStatus.toLowerCase() === "null") {
        finalStatus = defaultStatus;
      }

      return {
        attendanceId: rawId,
        date: dateVal,
        academicYear: rawYear,
        className: rawClass,
        studentId: rawStuId,
        studentName: rawStuName,
        courseId: rawCourse,
        status: finalStatus,
        term: rawTerm,
      };
    });
}

// 10. Helper to auto-generate the next Attendance ID (e.g. ATT-1001)
function generateNextAttendanceId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "ATT-1001";

  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  let maxIdNum = 1000;
  values.forEach((row) => {
    const idStr = String(row[0]).trim();
    if (idStr.startsWith("ATT-")) {
      const num = parseInt(idStr.substring(4), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return "ATT-" + (maxIdNum + 1);
}

// 11. Add Attendance Logic
function addAttendance(attendanceData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  if (!sheet) throw new Error("Attendance worksheet not found.");

  const studentId = String(attendanceData.studentId || "").trim();
  // Normalize the incoming date string first (the date input sends "yyyy-MM-dd")
  const normalizedDateString = formatAttendanceDate(attendanceData.date, ss);
  const dateValue = normalizedDateString
    ? buildAttendanceSheetDate(normalizedDateString, ss.getSpreadsheetTimeZone())
    : null;
  const dateString = normalizedDateString;

  if (!studentId) {
    return { success: false, message: "Student ID is required." };
  }
  if (!dateString) {
    return { success: false, message: "A valid attendance date is required." };
  }

  const hMap = getAttendanceHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  const incomingCourseId = String(attendanceData.courseId || "").trim();

  if (lastRow >= 3 && hMap) {
    const existingRows = sheet.getRange(3, 2, lastRow - 2, hMap.numCols).getValues();
    for (let i = 0; i < existingRows.length; i++) {
      const row = existingRows[i];
      const existingDate = row[hMap.idxDate];
      const existingStudentId = String(row[hMap.idxStudentId] || "").trim();
      const existingCourseId = String(row[hMap.idxCourse] || "").trim();
      // Format with formatAttendanceDate so both sides use the spreadsheet
      // timezone (toISOString would be UTC and could differ by a day)
      const existingDateString = existingDate
        ? formatAttendanceDate(existingDate, ss)
        : "";

      const sameCourse = !incomingCourseId || !existingCourseId || existingCourseId === incomingCourseId;
      if (
        existingStudentId === studentId &&
        existingDateString === dateString &&
        sameCourse
      ) {
        return {
          success: false,
          duplicate: true,
          message: `Attendance for student ${studentId} in course ${incomingCourseId || 'this course'} on ${dateString} has already been recorded.`,
        };
      }
    }
  }

  const nextId = generateNextAttendanceId(sheet);
  const targetRow = lastRow + 1;

  let rawStatus = String(attendanceData.status || "").trim();
  if (!rawStatus || rawStatus.toLowerCase() === "undefined" || rawStatus.toLowerCase() === "null") {
    // Use first status from settings as default
    const statuses = getAttendanceStatuses();
    rawStatus = statuses.length > 0 ? statuses[0].value : "Present";
  }

  const createdSnapshot = buildAuditSnapshot({
    attendanceId: nextId,
    studentId: studentId,
    studentName: attendanceData.studentName || '',
    date: dateString,
    academicYear: attendanceData.academicYear || '',
    className: attendanceData.className || '',
    courseId: attendanceData.courseId || '',
    status: rawStatus,
    term: attendanceData.term || ''
  });
  const createYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Create',
    module: 'Attendance',
    recordId: nextId,
    academicYear: attendanceData.academicYear,
    newValue: createdSnapshot,
    overrideConfirmed: attendanceData && attendanceData.__adminAcademicYearOverride === true
  });
  if (!createYearGuard.allowed) return createYearGuard;

  function setCellVal(idx, val) {
    if (idx >= 0) {
      sheet.getRange(targetRow, idx + 2).setValue(val);
    }
  }

  if (hMap) {
    setCellVal(hMap.idxId, nextId);
    setCellVal(hMap.idxDate, dateValue);
    setCellVal(hMap.idxYear, attendanceData.academicYear || "");
    setCellVal(hMap.idxClass, attendanceData.className || "");
    setCellVal(hMap.idxStudentId, studentId);
    setCellVal(hMap.idxStudentName, attendanceData.studentName || "");
    setCellVal(hMap.idxCourse, attendanceData.courseId || "");
    setCellVal(hMap.idxStatus, rawStatus);
    setCellVal(hMap.idxTerm, attendanceData.term || "");
  } else {
    sheet.getRange(targetRow, 2, 1, 9).setValues([[
      nextId, dateValue, attendanceData.academicYear || "", attendanceData.className || "",
      studentId, attendanceData.studentName || "", attendanceData.courseId || "", rawStatus, attendanceData.term || ""
    ]]);
  }

  if (typeof invalidateAttendanceCache === "function") {
    invalidateAttendanceCache();
  }

  safeLogAuditEvent(
    'Create',
    'Attendance',
    nextId,
    appendAcademicYearOverrideAuditDetails('Created attendance record', createYearGuard),
    null,
    createdSnapshot
  );

  return { success: true, attendanceId: nextId };
}

// 12a. Helper to find attendance row by ID
function findAttendanceRowById(sheet, attendanceId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return -1;
  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(attendanceId).trim()) {
      return i + 3;
    }
  }
  return -1;
}

// 12b. Update Attendance (Edit)
function updateAttendance(attendanceId, attendanceData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  if (!sheet) throw new Error("Attendance worksheet not found.");

  const row = findAttendanceRowById(sheet, attendanceId);
  if (row === -1)
    return { success: false, message: "Attendance record not found." };

  const hMap = getAttendanceHeaderMap(sheet);
  if (!hMap) return { success: false, message: "Could not read sheet headers." };

  const currentValues = sheet.getRange(row, 2, 1, 9).getValues()[0];
  const oldSnapshot = buildAuditSnapshot({
    attendanceId: String(currentValues[0] || '').trim(),
    date: currentValues[1] instanceof Date ? currentValues[1].toISOString() : String(currentValues[1] || '').trim(),
    academicYear: String(currentValues[2] || '').trim(),
    className: String(currentValues[3] || '').trim(),
    studentId: String(currentValues[4] || '').trim(),
    studentName: String(currentValues[5] || '').trim(),
    courseId: String(currentValues[6] || '').trim(),
    status: String(currentValues[7] || '').trim(),
    term: String(currentValues[8] || '').trim()
  });

  let rawStatus = String(attendanceData.status || "").trim();
  if (!rawStatus || rawStatus.toLowerCase() === "undefined" || rawStatus.toLowerCase() === "null") {
    // Use first status from settings as default
    const statuses = getAttendanceStatuses();
    rawStatus = statuses.length > 0 ? statuses[0].value : "Present";
  }

  const updatedSnapshot = buildAuditSnapshot({
    attendanceId: attendanceId,
    date: attendanceData.date || '',
    academicYear: attendanceData.academicYear || '',
    className: attendanceData.className || '',
    studentId: attendanceData.studentId || '',
    studentName: attendanceData.studentName || '',
    courseId: attendanceData.courseId || '',
    status: rawStatus,
    term: attendanceData.term || ''
  });
  const updateYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Update',
    module: 'Attendance',
    recordId: attendanceId,
    academicYears: [oldSnapshot.academicYear, attendanceData.academicYear],
    oldValue: oldSnapshot,
    newValue: updatedSnapshot,
    overrideConfirmed: attendanceData && attendanceData.__adminAcademicYearOverride === true
  });
  if (!updateYearGuard.allowed) return updateYearGuard;

  function setCellVal(idx, val) {
    if (idx >= 0) {
      sheet.getRange(row, idx + 2).setValue(val);
    }
  }

  const normalizedDateString = formatAttendanceDate(attendanceData.date, ss);
  setCellVal(hMap.idxDate, normalizedDateString
    ? buildAttendanceSheetDate(normalizedDateString, ss.getSpreadsheetTimeZone())
    : "");
  setCellVal(hMap.idxYear, attendanceData.academicYear || "");
  setCellVal(hMap.idxClass, attendanceData.className || "");
  setCellVal(hMap.idxStudentId, attendanceData.studentId || "");
  setCellVal(hMap.idxStudentName, attendanceData.studentName || "");
  setCellVal(hMap.idxCourse, attendanceData.courseId || "");
  setCellVal(hMap.idxStatus, rawStatus);
  setCellVal(hMap.idxTerm, attendanceData.term || "");

  if (typeof invalidateAttendanceCache === "function") {
    invalidateAttendanceCache();
  }

  safeLogAuditEvent(
    'Update',
    'Attendance',
    attendanceId,
    appendAcademicYearOverrideAuditDetails('Updated attendance record', updateYearGuard),
    oldSnapshot,
    updatedSnapshot
  );

  return { success: true, message: "Attendance record updated." };
}

// 12c. Delete Attendance (Delete)
function deleteAttendance(attendanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  if (!sheet) throw new Error("Attendance worksheet not found.");

  const row = findAttendanceRowById(sheet, attendanceId);
  if (row === -1)
    return { success: false, message: "Attendance record not found." };

  const currentValues = sheet.getRange(row, 2, 1, 9).getValues()[0];
  const deletedSnapshot = buildAuditSnapshot({
    attendanceId: String(currentValues[0] || '').trim(),
    date: currentValues[1] instanceof Date ? currentValues[1].toISOString() : String(currentValues[1] || '').trim(),
    academicYear: String(currentValues[2] || '').trim(),
    className: String(currentValues[3] || '').trim(),
    studentId: String(currentValues[4] || '').trim(),
    studentName: String(currentValues[5] || '').trim(),
    courseId: String(currentValues[6] || '').trim(),
    status: String(currentValues[7] || '').trim(),
    term: String(currentValues[8] || '').trim()
  });
  const deleteYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Delete',
    module: 'Attendance',
    recordId: attendanceId,
    academicYear: deletedSnapshot.academicYear,
    oldValue: deletedSnapshot,
    overrideConfirmed: arguments[1] && arguments[1].adminAcademicYearOverride === true
  });
  if (!deleteYearGuard.allowed) return deleteYearGuard;

  sheet.deleteRow(row);
  if (typeof invalidateAttendanceCache === "function") {
    invalidateAttendanceCache();
  }

  safeLogAuditEvent(
    'Delete',
    'Attendance',
    attendanceId,
    appendAcademicYearOverrideAuditDetails('Deleted attendance record', deleteYearGuard),
    deletedSnapshot,
    null
  );

  return { success: true, message: "Attendance record deleted." };
}

// 12c-2. Check and Fix Attendance Sheet Headers
function checkAndFixAttendanceHeaders() {
  requireLogin();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  if (!sheet) return { success: false, message: "Attendance sheet not found!" };

  const expectedHeaders = [
    "Attendance ID", "Date", "Academic Year", "Class",
    "Student ID", "Student Name", "Course ID", "Status", "Term"
  ];
  const currentHeaders = sheet.getRange(2, 2, 1, expectedHeaders.length).getValues()[0].map((header) => String(header || '').trim());

  try {
    sheet.getRange(2, 2, 1, expectedHeaders.length).setValues([expectedHeaders]);
    const headerRange = sheet.getRange(2, 2, 1, expectedHeaders.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4A90E2");
    headerRange.setFontColor("#FFFFFF");

    safeLogAuditEvent(
      'Repair',
      'Attendance',
      'Headers',
      'Repaired attendance sheet headers',
      buildAuditSnapshot({ headers: currentHeaders }),
      buildAuditSnapshot({ headers: expectedHeaders })
    );

    return { success: true, message: "Headers verified and fixed successfully!" };
  } catch (e) {
    return { success: false, message: "Error fixing headers: " + e.message };
  }
}

// 12d. Helper: Get just Student IDs and Names for dropdowns (includes Class field for dynamic filtering)
function getStudentIdNamePairs() {
  const students = getStudentsData();
  return students
    .map((s) => ({
      studentId: s.studentId,
      fullName: `${s.firstName} ${s.lastName}`.trim(),
      studentClass: s.class,
    }))
    .filter((s) => s.studentId !== "");
}

// 13. Helper: Get just Course IDs for dropdowns
function getCourseIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Courses");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const data = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  return data.map((row) => String(row[0]).trim()).filter(Boolean);
}
// 16. Export filtered attendance to PDF and return base64 string
function exportAttendancePdf(filters) {
  requireLogin();
  const records = getAttendanceData();

  // Records come back HTML-escaped from the cache layer, while filter values
  // arrive raw from the page dropdowns. Decode before comparing so years like
  // "2025/2026" match instead of their escaped form "2025&#x2F;2026".
  const decode = typeof decodeSanitizedHtml === 'function'
    ? decodeSanitizedHtml
    : function (v) { return v; };

  // Apply simple filter matching on provided keys (date, status, academicYear, className, term)
  const filtered = records.filter((r) => {
    if (filters.date && filters.date !== r.date) return false;
    if (filters.status && filters.status !== decode(r.status)) return false;
    if (filters.academicYear && filters.academicYear !== decode(r.academicYear))
      return false;
    if (filters.className && filters.className !== decode(r.className)) return false;
    if (filters.term && filters.term !== decode(r.term)) return false;
    if (filters.search) {
      const s = String(filters.search).toLowerCase();
      const hay = [
        r.attendanceId,
        r.studentId,
        r.studentName,
        r.courseId,
        r.status,
        r.className,
      ]
        .join(" ")
        .toLowerCase();
      if (!hay.includes(s)) return false;
    }
    return true;
  });

  // Build a small HTML table for PDF
  let html =
    '<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Arial,Helvetica,sans-serif}table{width:100%;border-collapse:collapse}th,td{border:1px solid #ddd;padding:8px;font-size:12px}th{background:#f4f6f8;font-weight:700}</style></head><body>';
  html += "<h2>Attendance Report</h2>";
  html +=
    "<table><thead><tr><th>Attendance ID</th><th>Date</th><th>Academic Year</th><th>Class</th><th>Student</th><th>Course</th><th>Status</th><th>Term</th></tr></thead><tbody>";

  filtered.forEach((r) => {
    html += `<tr><td>${r.attendanceId}</td><td>${r.date}</td><td>${r.academicYear}</td><td>${r.className}</td><td>${r.studentId} ${r.studentName}</td><td>${r.courseId}</td><td>${r.status}</td><td>${r.term}</td></tr>`;
  });

  html += "</tbody></table></body></html>";

  const blob = HtmlService.createHtmlOutput(html)
    .getBlob()
    .getAs("application/pdf");
  const encoded = Utilities.base64Encode(blob.getBytes());
  return encoded;
}

// Save filtered attendance PDF to Drive and return file URL and id
function exportAttendancePdfToDrive(filters) {
  requireLogin();
  // Optional: filters.folderId may be provided to save into a specific Drive folder
  const folderId =
    filters && filters.folderId ? String(filters.folderId).trim() : null;

  // Reuse existing exporter to get base64 PDF
  const base64 = exportAttendancePdf(filters);
  if (!base64) return { success: false, message: "No PDF generated" };

  const bytes = Utilities.base64Decode(base64);
  const fileName =
    "Attendance_Report_" + new Date().toISOString().slice(0, 10) + ".pdf";
  const blob = Utilities.newBlob(bytes, "application/pdf", fileName);

  var file;
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      file = folder.createFile(blob);
    } catch (e) {
      // fallback to root if folder invalid
      Logger.log("Invalid folderId or access denied: " + e.message);
      file = DriveApp.createFile(blob);
    }
  } else {
    file = DriveApp.createFile(blob);
  }

  // Attempt to set sharing so the download redirect works for anyone with link
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log("Could not set sharing: " + e.message);
  }

  const downloadUrl =
    "https://docs.google.com/uc?export=download&id=" +
    encodeURIComponent(file.getId());
  return {
    success: true,
    url: file.getUrl(),
    id: file.getId(),
    name: file.getName(),
    downloadUrl: downloadUrl,
  };
}
// 14. Fetch Courses Data
function getCoursesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Courses");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const dataRange = sheet.getRange(3, 2, lastRow - 2, 6);
  const data = dataRange.getValues();

  return data
    .map((row) => ({
      courseId: String(row[0]).trim(),
      courseName: String(row[1]).trim(),
      instructor: String(row[2]).trim(),
      credits: String(row[3]).trim(),
      semester: String(row[4]).trim(),
      status: String(row[5]).trim(),
    }))
    .filter((c) => c.courseId !== "");
}

// 15. Helper to auto-generate the next Course ID (e.g. CRS-1001)
function generateNextCourseId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "CRS-1001";

  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  let maxIdNum = 1000;
  values.forEach((row) => {
    const idStr = String(row[0]).trim();
    if (idStr.startsWith("CRS-")) {
      const num = parseInt(idStr.substring(4), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return "CRS-" + (maxIdNum + 1);
}

// 16. Add Course Logic
function addCourse(courseData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Courses");
  if (!sheet) throw new Error("Courses worksheet not found.");

  const nextId = generateNextCourseId(sheet);
  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  // Set values starting from Column 2 (B) leaving Column A empty
  sheet
    .getRange(targetRow, 2, 1, 6)
    .setValues([
      [
        nextId,
        courseData.courseName,
        courseData.instructor,
        courseData.credits,
        courseData.semester,
        courseData.status || "Active",
      ],
    ]);

  safeLogAuditEvent(
    'Create',
    'Courses',
    nextId,
    'Created course record',
    null,
    buildAuditSnapshot({
      courseId: nextId,
      courseName: courseData.courseName,
      instructor: courseData.instructor,
      credits: courseData.credits,
      semester: courseData.semester,
      status: courseData.status || 'Active'
    })
  );

  return { success: true, courseId: nextId };
}

// Update Course
function updateCourse(courseId, courseData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Courses");
  if (!sheet) throw new Error("Courses worksheet not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "No courses found" };

  // Find the row with the course ID (Column B)
  const dataRange = sheet.getRange(2, 2, lastRow - 1, 6);
  const data = dataRange.getValues();
  
  let rowIndex = -1;
  let oldData = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === courseId) {
      rowIndex = i + 2; // +2 because data starts at row 2
      oldData = {
        courseId: String(data[i][0]).trim(),
        courseName: String(data[i][1]).trim(),
        instructor: String(data[i][2]).trim(),
        credits: String(data[i][3]).trim(),
        semester: String(data[i][4]).trim(),
        status: String(data[i][5]).trim()
      };
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: "Course not found" };
  }

  // Update the row (Columns B to G: courseId, courseName, instructor, credits, semester, status)
  sheet.getRange(rowIndex, 2, 1, 6).setValues([[
    courseId, // Keep same course ID
    courseData.courseName,
    courseData.instructor,
    courseData.credits,
    courseData.semester,
    courseData.status || "Active"
  ]]);

  safeLogAuditEvent(
    'Update',
    'Courses',
    courseId,
    'Updated course record',
    buildAuditSnapshot(oldData),
    buildAuditSnapshot({
      courseId: courseId,
      courseName: courseData.courseName,
      instructor: courseData.instructor,
      credits: courseData.credits,
      semester: courseData.semester,
      status: courseData.status || 'Active'
    })
  );

  return { success: true };
}

// Delete Course
function deleteCourse(courseId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Courses");
  if (!sheet) throw new Error("Courses worksheet not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "No courses found" };

  // Find the row with the course ID (Column B)
  const dataRange = sheet.getRange(2, 2, lastRow - 1, 6);
  const data = dataRange.getValues();
  
  let rowIndex = -1;
  let deletedData = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === courseId) {
      rowIndex = i + 2; // +2 because data starts at row 2
      deletedData = {
        courseId: String(data[i][0]).trim(),
        courseName: String(data[i][1]).trim(),
        instructor: String(data[i][2]).trim(),
        credits: String(data[i][3]).trim(),
        semester: String(data[i][4]).trim(),
        status: String(data[i][5]).trim()
      };
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: "Course not found" };
  }

  // Delete the row
  sheet.deleteRow(rowIndex);

  safeLogAuditEvent(
    'Delete',
    'Courses',
    courseId,
    'Deleted course record',
    buildAuditSnapshot(deletedData),
    null
  );

  return { success: true };
}

// 17. Fetch Enrollments Data
function getEnrollmentsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Enrollments");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  // Row 3 to Last, Col B(2) to G(7) => 6 columns
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 6);
  const data = dataRange.getValues();

  return data.map((row) => ({
    enrollmentId: String(row[0]).trim(),
    studentId: String(row[1]).trim(),
    courseId: String(row[2]).trim(),
    enrollmentDate: row[3] ? new Date(row[3]).toISOString().split("T")[0] : "",
    grade: String(row[4]).trim(),
    status: String(row[5]).trim(),
  }));
}

// 18. Helper: Get Course IDs AND Names for enrollment dropdowns
function getCourseIdNamePairs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Courses");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  // Get ID (Col B), Name (Col C)
  const data = sheet.getRange(3, 2, lastRow - 2, 2).getValues();
  return data.map((row) => ({
    courseId: String(row[0]).trim(),
    courseName: String(row[1]).trim(),
  }));
}

// 19. Helper: Generate Enrollment ID (e.g. ENR-1001)
function generateNextEnrollmentId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "ENR-1001";

  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  let maxIdNum = 1000;
  values.forEach((row) => {
    const idStr = String(row[0]).trim();
    if (idStr.startsWith("ENR-")) {
      const num = parseInt(idStr.substring(4), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return "ENR-" + (maxIdNum + 1);
}

// 20. Add Enrollment Logic
function addEnrollment(enrollmentData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Enrollments");
  if (!sheet) throw new Error("Enrollments worksheet not found.");

  const nextId = generateNextEnrollmentId(sheet);
  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  sheet
    .getRange(targetRow, 2, 1, 6)
    .setValues([
      [
        nextId,
        enrollmentData.studentId,
        enrollmentData.courseId,
        new Date(enrollmentData.enrollmentDate),
        enrollmentData.grade,
        enrollmentData.status,
      ],
    ]);

  safeLogAuditEvent(
    'Create',
    'Enrollments',
    nextId,
    'Created enrollment record',
    null,
    buildAuditSnapshot({
      enrollmentId: nextId,
      studentId: enrollmentData.studentId,
      courseId: enrollmentData.courseId,
      enrollmentDate: enrollmentData.enrollmentDate,
      grade: enrollmentData.grade,
      status: enrollmentData.status
    })
  );

  return { success: true, enrollmentId: nextId };
}

// Class helpers
function normalizeClassNameForComparison_(className) {
  return String(className || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function invalidateClassesCache_() {
  try {
    if (typeof invalidateCacheOnModify === 'function') {
      invalidateCacheOnModify('Classes');
    } else if (typeof invalidateCache === 'function') {
      invalidateCache('data_Classes');
    }
  } catch (e) {
    Logger.log('Could not invalidate Classes cache: ' + e.message);
  }
}

function findClassRowByName_(sheet, className, ignoreRow) {
  const targetKey = normalizeClassNameForComparison_(className);
  if (!targetKey) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    if (ignoreRow && rowNumber === Number(ignoreRow)) continue;

    const existingName = String(values[i][0] || '').trim();
    if (existingName && normalizeClassNameForComparison_(existingName) === targetKey) {
      return { row: rowNumber, className: existingName };
    }
  }

  return null;
}

function getStudentsDataForClassCount_() {
  try {
    if (typeof getStudentsDataUncached === 'function') {
      return getStudentsDataUncached();
    }
  } catch (e) {
    Logger.log('Could not read uncached students for class count: ' + e.message);
  }

  try {
    if (typeof getStudentsData === 'function') {
      return getStudentsData();
    }
  } catch (e) {
    Logger.log('Could not read students for class count: ' + e.message);
  }

  return [];
}

function getStudentCountsByNormalizedClass_() {
  const counts = {};
  const studentsData = getStudentsDataForClassCount_();

  studentsData.forEach(function(student) {
    const rawClassName = student && (student.class || student.className || student.studentClass);
    const classKey = normalizeClassNameForComparison_(rawClassName);
    if (classKey) {
      counts[classKey] = (counts[classKey] || 0) + 1;
    }
  });

  return counts;
}

function getStudentCountForClass_(className) {
  const classKey = normalizeClassNameForComparison_(className);
  if (!classKey) return 0;

  const counts = getStudentCountsByNormalizedClass_();
  return counts[classKey] || 0;
}

// 21. Fetch Classes Data (For the Classes page)
function getClassesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Classes");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Get class names from Column A, starting from Row 2
  const classRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const classData = classRange.getValues();
  const studentCounts = getStudentCountsByNormalizedClass_();
  
  return classData.map((row, idx) => {
    const className = String(row[0]).trim();
    return {
      id: idx + 2, // Sheet row number for reference
      className: typeof sanitizeHtml === 'function' ? sanitizeHtml(className) : className,
      studentCount: studentCounts[normalizeClassNameForComparison_(className)] || 0
    };
  }).filter((r) => r.className);
}

// 22. Add Class Logic
function addClass(className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Classes");
  if (!sheet) throw new Error("Classes worksheet not found.");

  className = String(className || '').trim().replace(/\s+/g, ' ');
  if (!className) {
    return { success: false, message: 'Class name is required.' };
  }

  const duplicate = findClassRowByName_(sheet, className);
  if (duplicate) {
    return {
      success: false,
      message: 'A class named "' + duplicate.className + '" already exists.'
    };
  }

  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  // Set value in Column A
  sheet.getRange(targetRow, 1).setValue(className);
  invalidateClassesCache_();

  safeLogAuditEvent(
    'Create',
    'Classes',
    className,
    'Created class record',
    null,
    buildAuditSnapshot({ className: className, rowId: targetRow })
  );

  return { success: true, className: className, rowId: targetRow };
}

// 23. Update Class Logic
function updateClass(id, className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Classes");
  if (!sheet) throw new Error("Classes worksheet not found.");

  id = Number(id);
  // id is the row number in the sheet
  if (!id || id < 2 || id > sheet.getLastRow()) throw new Error("Invalid class ID.");

  className = String(className || '').trim().replace(/\s+/g, ' ');
  if (!className) {
    return { success: false, message: 'Class name is required.' };
  }

  const duplicate = findClassRowByName_(sheet, className, id);
  if (duplicate) {
    return {
      success: false,
      message: 'A class named "' + duplicate.className + '" already exists.'
    };
  }

  const oldClassName = String(sheet.getRange(id, 1).getValue() || '').trim();
  if (!oldClassName) {
    return { success: false, message: 'Class not found.' };
  }

  // Update Column A with the new class name
  sheet.getRange(id, 1).setValue(className);
  invalidateClassesCache_();

  safeLogAuditEvent(
    'Update',
    'Classes',
    String(id),
    'Updated class record',
    buildAuditSnapshot({ className: oldClassName, rowId: id }),
    buildAuditSnapshot({ className: className, rowId: id })
  );

  return { success: true, className: className, rowId: id };
}

// 24. Delete Class Logic
function deleteClass(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Classes");
  if (!sheet) throw new Error("Classes worksheet not found.");

  id = Number(id);
  // id is the row number in the sheet
  if (!id || id < 2 || id > sheet.getLastRow()) throw new Error("Invalid class ID.");

  const className = String(sheet.getRange(id, 1).getValue() || '').trim();
  if (!className) {
    return { success: false, message: 'Class not found.' };
  }

  const studentCount = getStudentCountForClass_(className);
  if (studentCount > 0) {
    return {
      success: false,
      message: 'Cannot delete "' + className + '" because ' + studentCount + ' student' + (studentCount === 1 ? ' is' : 's are') + ' currently assigned to this class.'
    };
  }

  // Delete the entire row
  sheet.deleteRow(id);
  invalidateClassesCache_();

  safeLogAuditEvent(
    'Delete',
    'Classes',
    String(id),
    'Deleted class record',
    buildAuditSnapshot({ className: className, rowId: id }),
    null
  );

  return { success: true };
}

// --- ACADEMIC YEARS FUNCTIONS ---

// Fetch data from Academic Years sheet
function getAcademicYearsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Academic Years");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; // Headers at Row 2

  // Get range from Row 3, Column B to Column C
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 2);
  const data = dataRange.getValues();

  return data
    .filter((row) => String(row[0]).trim() !== "") // Filter out empty rows
    .map((row) => ({
      academicYear: String(row[0]).trim(),
      status: String(row[1]).trim() || "Active",
    }));
}

// Add new Academic Year
function addAcademicYear(yearData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Academic Years");
  if (!sheet) throw new Error("Academic Years worksheet not found.");

  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  // Write to Column B and C
  sheet
    .getRange(targetRow, 2, 1, 2)
    .setValues([[yearData.academicYear, yearData.status || "Active"]]);

  safeLogAuditEvent(
    'Create',
    'Academic Years',
    yearData.academicYear,
    'Created academic year record',
    null,
    buildAuditSnapshot({
      academicYear: yearData.academicYear,
      status: yearData.status || 'Active'
    })
  );

  return { success: true };
}

// Update Academic Year
function updateAcademicYear(originalYear, yearData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Academic Years");
  if (!sheet) throw new Error("Academic Years worksheet not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return { success: false, message: "No academic years found" };

  // Find the row with the original year
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 2);
  const data = dataRange.getValues();
  
  let rowIndex = -1;
  let oldData = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === originalYear) {
      rowIndex = i + 3; // +3 because data starts at row 3
      oldData = {
        academicYear: String(data[i][0]).trim(),
        status: String(data[i][1]).trim()
      };
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: "Academic year not found" };
  }

  // Update the row
  sheet.getRange(rowIndex, 2, 1, 2).setValues([[
    yearData.academicYear,
    yearData.status || "Active"
  ]]);

  safeLogAuditEvent(
    'Update',
    'Academic Years',
    originalYear,
    'Updated academic year record',
    buildAuditSnapshot(oldData),
    buildAuditSnapshot({
      academicYear: yearData.academicYear,
      status: yearData.status || 'Active'
    })
  );

  return { success: true };
}

// Delete Academic Year
function deleteAcademicYear(yearName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Academic Years");
  if (!sheet) throw new Error("Academic Years worksheet not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return { success: false, message: "No academic years found" };

  // Find the row with the year
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 2);
  const data = dataRange.getValues();
  
  let rowIndex = -1;
  let deletedData = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === yearName) {
      rowIndex = i + 3; // +3 because data starts at row 3
      deletedData = {
        academicYear: String(data[i][0]).trim(),
        status: String(data[i][1]).trim()
      };
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: "Academic year not found" };
  }

  // Delete the row
  sheet.deleteRow(rowIndex);

  safeLogAuditEvent(
    'Delete',
    'Academic Years',
    yearName,
    'Deleted academic year record',
    buildAuditSnapshot(deletedData),
    null
  );

  return { success: true };
}

// --- TEACHERS FUNCTIONS ---

// Fetch Teachers Data
function getTeachersData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Teachers");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  // Get range from Row 3, Col B to Col F
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 5);
  const data = dataRange.getValues();

  return data.map((row) => ({
    teacherId: String(row[0]).trim(),
    firstName: String(row[1]).trim(),
    lastName: String(row[2]).trim(),
    class: String(row[3]).trim(),
    academicYear: String(row[4]).trim(),
  }));
}

// Helper to auto-generate the next Teacher ID (e.g. TCH-1001)
function generateNextTeacherId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "TCH-1001";

  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
  let maxIdNum = 1000;
  values.forEach((row) => {
    const idStr = String(row[0]).trim();
    if (idStr.startsWith("TCH-")) {
      const num = parseInt(idStr.substring(4), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return "TCH-" + (maxIdNum + 1);
}

// Add new Teacher
function addTeacher(teacherData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Teachers");
  if (!sheet) throw new Error("Teachers worksheet not found.");

  const nextId = generateNextTeacherId(sheet);
  const teacherSnapshot = buildAuditSnapshot({
    teacherId: nextId,
    firstName: teacherData.firstName,
    lastName: teacherData.lastName,
    class: teacherData.class,
    academicYear: teacherData.academicYear
  });
  const teacherYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Create',
    module: 'Teachers',
    recordId: nextId,
    academicYear: teacherData.academicYear,
    newValue: teacherSnapshot,
    overrideConfirmed: teacherData && teacherData.__adminAcademicYearOverride === true
  });
  if (!teacherYearGuard.allowed) return teacherYearGuard;
  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  // Write to Column B to F
  sheet
    .getRange(targetRow, 2, 1, 5)
    .setValues([
      [
        nextId,
        teacherData.firstName,
        teacherData.lastName,
        teacherData.class,
        teacherData.academicYear,
      ],
    ]);

  safeLogAuditEvent(
    'Create',
    'Teachers',
    nextId,
    appendAcademicYearOverrideAuditDetails('Created teacher record', teacherYearGuard),
    null,
    teacherSnapshot
  );

  return { success: true, teacherId: nextId };
}

// Update Teacher
function updateTeacher(teacherId, teacherData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Teachers");
  if (!sheet) throw new Error("Teachers worksheet not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "No teachers found" };

  // Find the row with the teacher ID (Column B)
  const dataRange = sheet.getRange(2, 2, lastRow - 1, 5);
  const data = dataRange.getValues();
  
  let rowIndex = -1;
  let oldData = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === teacherId) {
      rowIndex = i + 2; // +2 because data starts at row 2
      oldData = {
        teacherId: String(data[i][0]).trim(),
        firstName: String(data[i][1]).trim(),
        lastName: String(data[i][2]).trim(),
        class: String(data[i][3]).trim(),
        academicYear: String(data[i][4]).trim()
      };
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: "Teacher not found" };
  }

  // Build new snapshot for audit
  const newSnapshot = buildAuditSnapshot({
    teacherId: teacherId,
    firstName: teacherData.firstName,
    lastName: teacherData.lastName,
    class: teacherData.class,
    academicYear: teacherData.academicYear
  });

  // Check academic year security
  const teacherYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Update',
    module: 'Teachers',
    recordId: teacherId,
    academicYear: teacherData.academicYear,
    newValue: newSnapshot,
    overrideConfirmed: teacherData && teacherData.__adminAcademicYearOverride === true
  });
  if (!teacherYearGuard.allowed) return teacherYearGuard;

  // Update the row (Columns B to F: teacherId, firstName, lastName, class, academicYear)
  sheet.getRange(rowIndex, 2, 1, 5).setValues([[
    teacherId, // Keep same teacher ID
    teacherData.firstName,
    teacherData.lastName,
    teacherData.class,
    teacherData.academicYear
  ]]);

  safeLogAuditEvent(
    'Update',
    'Teachers',
    teacherId,
    appendAcademicYearOverrideAuditDetails('Updated teacher record', teacherYearGuard),
    buildAuditSnapshot(oldData),
    newSnapshot
  );

  return { success: true };
}

// Delete Teacher
function deleteTeacher(teacherId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Teachers");
  if (!sheet) throw new Error("Teachers worksheet not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "No teachers found" };

  // Find the row with the teacher ID (Column B)
  const dataRange = sheet.getRange(2, 2, lastRow - 1, 5);
  const data = dataRange.getValues();
  
  let rowIndex = -1;
  let deletedData = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === teacherId) {
      rowIndex = i + 2; // +2 because data starts at row 2
      deletedData = {
        teacherId: String(data[i][0]).trim(),
        firstName: String(data[i][1]).trim(),
        lastName: String(data[i][2]).trim(),
        class: String(data[i][3]).trim(),
        academicYear: String(data[i][4]).trim()
      };
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: "Teacher not found" };
  }

  // Delete the row
  sheet.deleteRow(rowIndex);

  safeLogAuditEvent(
    'Delete',
    'Teachers',
    teacherId,
    'Deleted teacher record',
    buildAuditSnapshot(deletedData),
    null
  );

  return { success: true };
}

// Generic helper to get just Class names (for Teacher dropdown)
function getClassNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Classes");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map((row) => String(row[0]).trim()).filter(Boolean);
}

// --- PERFORMANCE FUNCTIONS ---

// Fetch Performance Data with dynamic column mapping (used by DataAccessLayer cache wrapper)
function getPerformanceDataFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Performance");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  // Read headers in Row 2, from Col B (2) to Col M (13) -> 12 columns
  const headers = sheet
    .getRange(2, 2, 1, 12)
    .getValues()[0]
    .map((h) => String(h).trim().toLowerCase());

  // Helper: find header index by prioritized alias lists
  function findHeaderIndex(aliasGroups) {
    // aliasGroups: array of alias strings in priority order
    // 1) exact match
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      for (let a = 0; a < aliasGroups.length; a++) {
        if (h === aliasGroups[a]) return i;
      }
    }
    // 2) contains any alias (less strict)
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      for (let a = 0; a < aliasGroups.length; a++) {
        if (h.indexOf(aliasGroups[a]) !== -1) return i;
      }
    }
    return -1;
  }

  const idxId = findHeaderIndex([
    "performance id",
    "performance",
    "perf id",
    "perf",
  ]);
  const idxStudentId = findHeaderIndex([
    "student id",
    "studentid",
    "student_id",
    "sid",
  ]);
  const idxName = findHeaderIndex(["student name", "name", "student"]);
  const idxClass = findHeaderIndex(["student class", "class name", "class"]);
  const idxTerm = findHeaderIndex(["term"]);
  const idxYear = findHeaderIndex(["academic year", "year"]);
  const idxClassScore = findHeaderIndex(["class score", "classscore"]);
  const idxExam100 = findHeaderIndex([
    "exam score (100%)",
    "exam 100",
    "exam score 100",
    "exam100",
  ]);
  const idxExam60 = findHeaderIndex([
    "exam score (50%)",
    "exam score (50)",
    "exam 50",
    "exam score 50",
    "exam score (60%)",
    "exam 60",
    "exam score 60",
    "exam60",
  ]);
  const idxTotal = findHeaderIndex(["total"]);
  const idxRank = findHeaderIndex(["rank"]);
  const idxCourse = findHeaderIndex(["course"]);

  const data = sheet.getRange(3, 2, lastRow - 2, 12).getValues();

  return data
    .filter((row) => {
      // Filter out empty rows - check if student name exists
      const studentName = idxName !== -1 ? String(row[idxName]).trim() : "";
      return studentName !== "";
    })
    .map((row, index) => {
      return {
        performanceId: `ROW-${index + 3}`, // Use actual row number as ID
        rowNumber: index + 3, // Store actual row number for updates
        studentId: idxStudentId !== -1 ? String(row[idxStudentId]).trim() : "",
        studentName: idxName !== -1 ? String(row[idxName]).trim() : "",
        studentClass: idxClass !== -1 ? String(row[idxClass]).trim() : "",
        term: idxTerm !== -1 ? String(row[idxTerm]).trim() : "",
        academicYear: idxYear !== -1 ? String(row[idxYear]).trim() : "",
        classScore: idxClassScore !== -1 ? Number(row[idxClassScore]) || 0 : 0,
        examScore100: idxExam100 !== -1 ? Number(row[idxExam100]) || 0 : 0,
        examScore60: idxExam60 !== -1 ? Number(row[idxExam60]) || 0 : 0,
        total: idxTotal !== -1 ? Number(row[idxTotal]) || 0 : 0,
        rank: idxRank !== -1 ? String(row[idxRank]).trim() : "",
        course: idxCourse !== -1 ? String(row[idxCourse]).trim() : "",
      };
    });
}

function buildPerformanceAuditSnapshot(perfData) {
  if (!perfData) return null;

  return buildAuditSnapshot({
    performanceId: perfData.performanceId,
    studentId: perfData.studentId,
    studentName: perfData.studentName,
    studentClass: perfData.studentClass,
    term: perfData.term,
    academicYear: perfData.academicYear,
    classScore: Number(perfData.classScore) || 0,
    examScore100: Number(perfData.examScore100) || 0,
    examScore60: Number(perfData.examScore60) || 0,
    total: Number(perfData.total) || 0,
    rank: perfData.rank,
    course: perfData.course
  });
}

function normalizePerformanceText(value) {
  return String(value || '').trim();
}

function normalizePerformanceKey(value) {
  return normalizePerformanceText(value).toLowerCase();
}

function roundPerformanceNumber(value) {
  return Math.round((Number(value) || 0) * 10) / 10;
}

function calculatePerformanceDerivedScores(classScore, examScore100) {
  const percentages = getScorePercentages();
  const parsedClassScore = roundPerformanceNumber(classScore);
  const parsedExamScore100 = roundPerformanceNumber(examScore100);
  const weightedExamScore = roundPerformanceNumber(parsedExamScore100 * Number(percentages.examScoreDecimal || 0.5));
  const total = roundPerformanceNumber(parsedClassScore + weightedExamScore);

  return {
    classScore: parsedClassScore,
    examScore100: parsedExamScore100,
    examScore60: weightedExamScore,
    examScore50: weightedExamScore,
    total: total,
    maxClassScore: Number(percentages.classScorePercentage) || 50
  };
}

function normalizePerformanceTerm(value) {
  const key = normalizePerformanceKey(value).replace(/\s+/g, '');
  const termMap = {
    term1: 'Term 1',
    term2: 'Term 2',
    term3: 'Term 3'
  };
  return termMap[key] || '';
}

function buildPerformanceImportReferenceData() {
  const studentLookup = {};
  getStudentIdNamePairs().forEach(function(student) {
    const studentId = normalizePerformanceText(student && student.studentId);
    if (!studentId) return;
    studentLookup[normalizePerformanceKey(studentId)] = {
      studentId: studentId,
      fullName: normalizePerformanceText(student.fullName),
      studentClass: normalizePerformanceText(student.studentClass)
    };
  });

  const classLookup = {};
  getClassNames().forEach(function(className) {
    const value = normalizePerformanceText(className);
    if (!value) return;
    classLookup[normalizePerformanceKey(value)] = value;
  });

  const courseLookup = {};
  getCoursesData().forEach(function(course) {
    const courseName = normalizePerformanceText(course && course.courseName);
    if (!courseName) return;
    courseLookup[normalizePerformanceKey(courseName)] = courseName;
  });

  const academicYearLookup = {};
  getAcademicYearsData().forEach(function(year) {
    const academicYear = normalizePerformanceText(year && year.academicYear);
    if (!academicYear) return;
    academicYearLookup[normalizePerformanceKey(academicYear)] = academicYear;
  });

  return {
    studentLookup: studentLookup,
    classLookup: classLookup,
    courseLookup: courseLookup,
    academicYearLookup: academicYearLookup
  };
}

function importPerformance(performanceArray) {
  try {
    requireLogin();

    const rows = Array.isArray(performanceArray) ? performanceArray : [];
    if (rows.length === 0) {
      return { success: false, message: 'No performance data provided for import.' };
    }

    const referenceData = buildPerformanceImportReferenceData();
    const uniqueYears = rows
      .map(function(row) {
        const rawYear = normalizePerformanceText(row && row.academicYear);
        return referenceData.academicYearLookup[normalizePerformanceKey(rawYear)] || normalizeAcademicYearValue(rawYear);
      })
      .filter(function(year, index, arr) { return year && arr.indexOf(year) === index; });

    const importGuard = enforceAcademicYearCrudSecurity({
      action: 'Import',
      module: 'Performance',
      recordId: 'PerformanceImport',
      academicYears: uniqueYears,
      newValue: buildAuditSnapshot({ totalRows: rows.length, academicYears: uniqueYears }),
      overrideConfirmed: rows.some(function(row) { return row && row.__adminAcademicYearOverride === true; })
    });
    if (!importGuard.allowed) return importGuard;

    let imported = 0;
    let failed = 0;
    let duplicates = 0;
    const errors = [];
    const seenImportKeys = {};

    rows.forEach(function(row, index) {
      const rowNumber = row && row.rowNumber ? row.rowNumber : index + 2;
      try {
        const studentId = normalizePerformanceText(row && row.studentId);
        const rawStudentName = normalizePerformanceText(row && row.studentName);
        const rawStudentClass = normalizePerformanceText((row && row.studentClass) || (row && row.class));
        const lookupStudent = referenceData.studentLookup[normalizePerformanceKey(studentId)] || null;
        const rawTerm = normalizePerformanceText(row && row.term);
        const rawAcademicYear = normalizePerformanceText(row && row.academicYear);
        const rawCourse = normalizePerformanceText((row && row.course) || (row && row.subject));
        const rawClassScore = row && row.classScore;
        const rawExamScore100 = row && row.examScore100;

        if (!studentId) {
          failed++;
          errors.push(`Row ${rowNumber}: Student ID is required.`);
          return;
        }

        if (!lookupStudent) {
          failed++;
          errors.push(`Row ${rowNumber}: Student ID not found in the system.`);
          return;
        }

        const studentName = normalizePerformanceText(lookupStudent.fullName);
        const studentClass = referenceData.classLookup[normalizePerformanceKey(lookupStudent.studentClass)] || normalizePerformanceText(lookupStudent.studentClass);
        const term = normalizePerformanceTerm(rawTerm);
        const academicYear = referenceData.academicYearLookup[normalizePerformanceKey(rawAcademicYear)] || '';
        const course = referenceData.courseLookup[normalizePerformanceKey(rawCourse)] || '';

        if (rawStudentName && normalizePerformanceKey(rawStudentName) !== normalizePerformanceKey(studentName)) {
          failed++;
          errors.push(`Row ${rowNumber}: Student Name does not match Student ID ${studentId}. Expected ${studentName}.`);
          return;
        }

        if (rawStudentClass && normalizePerformanceKey(rawStudentClass) !== normalizePerformanceKey(studentClass)) {
          failed++;
          errors.push(`Row ${rowNumber}: Class does not match Student ID ${studentId}. Expected ${studentClass}.`);
          return;
        }

        if (!studentName || !studentClass) {
          failed++;
          errors.push(`Row ${rowNumber}: Student record is missing a valid name or class.`);
          return;
        }

        if (!term) {
          failed++;
          errors.push(`Row ${rowNumber}: Term must be Term 1, Term 2, or Term 3.`);
          return;
        }

        if (!academicYear) {
          failed++;
          errors.push(`Row ${rowNumber}: Academic Year was not found in the system.`);
          return;
        }

        if (!course) {
          failed++;
          errors.push(`Row ${rowNumber}: Course / Subject was not found in the system.`);
          return;
        }

        const importKey = [studentId, studentClass, term, academicYear, course]
          .map(normalizePerformanceKey)
          .join('|');
        if (seenImportKeys[importKey]) {
          duplicates++;
          errors.push(`Row ${rowNumber}: Duplicate entry found in the CSV file for this student, class, term, academic year, and course.`);
          return;
        }
        seenImportKeys[importKey] = true;

        if (rawClassScore === '' || rawClassScore == null || isNaN(Number(rawClassScore))) {
          failed++;
          errors.push(`Row ${rowNumber}: Invalid Class Score.`);
          return;
        }

        if (rawExamScore100 === '' || rawExamScore100 == null || isNaN(Number(rawExamScore100))) {
          failed++;
          errors.push(`Row ${rowNumber}: Invalid Exam Score (100).`);
          return;
        }

        const derived = calculatePerformanceDerivedScores(rawClassScore, rawExamScore100);
        if (derived.classScore < 0 || derived.classScore > derived.maxClassScore) {
          failed++;
          errors.push(`Row ${rowNumber}: Class Score must be between 0 and ${derived.maxClassScore}.`);
          return;
        }
        if (derived.examScore100 < 0 || derived.examScore100 > 100) {
          failed++;
          errors.push(`Row ${rowNumber}: Exam Score must be between 0 and 100.`);
          return;
        }

        const payload = {
          studentId: studentId,
          studentName: studentName,
          studentClass: studentClass,
          term: term,
          academicYear: academicYear,
          course: course,
          classScore: derived.classScore,
          examScore100: derived.examScore100,
          examScore60: derived.examScore60,
          examScore50: derived.examScore50,
          total: derived.total
        };
        if (importGuard.adminOverride) payload.__adminAcademicYearOverride = true;

        const result = addPerformance(payload);
        if (result && result.success) {
          imported++;
          return;
        }

        if (result && result.requiresAdminOverride) {
          throw new Error(result.message || 'Admin confirmation required.');
        }

        if (result && result.message && /duplicate/i.test(result.message)) {
          duplicates++;
        } else {
          failed++;
        }
        errors.push(`Row ${rowNumber}: ${(result && result.message) ? result.message : 'Import failed.'}`);
      } catch (error) {
        failed++;
        errors.push(`Row ${rowNumber}: ${error.message}`);
      }
    });

    safeLogAuditEvent(
      'Import',
      'Performance',
      'BulkImport',
      appendAcademicYearOverrideAuditDetails('Imported performance records in bulk', importGuard),
      null,
      buildAuditSnapshot({
        totalRows: rows.length,
        imported: imported,
        failed: failed,
        duplicates: duplicates,
        academicYears: uniqueYears
      })
    );

    return {
      success: true,
      imported: imported,
      failed: failed,
      duplicates: duplicates,
      errors: errors
    };
  } catch (error) {
    Logger.log('ERROR in importPerformance: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function saveBulkPerformanceEntries(payload) {
  try {
    requireLogin();

    const data = payload || {};
    const studentClass = normalizePerformanceText(data.studentClass);
    const term = normalizePerformanceText(data.term);
    const academicYear = normalizePerformanceText(data.academicYear);
    const course = normalizePerformanceText(data.course);
    const records = Array.isArray(data.records) ? data.records : [];

    if (!studentClass || !term || !academicYear || !course) {
      return { success: false, message: 'Class, Term, Academic Year, and Course are required.' };
    }

    if (records.length === 0) {
      return { success: false, message: 'No student performance rows were provided.' };
    }

    const bulkGuard = enforceAcademicYearCrudSecurity({
      action: 'Update',
      module: 'Performance',
      recordId: `${studentClass}|${term}|${academicYear}|${course}`,
      academicYear: academicYear,
      newValue: buildAuditSnapshot({ studentClass: studentClass, term: term, academicYear: academicYear, course: course, rows: records.length }),
      overrideConfirmed: data.__adminAcademicYearOverride === true
    });
    if (!bulkGuard.allowed) return bulkGuard;

    let created = 0;
    let updated = 0;
    let skipped = 0;
    let failed = 0;
    let duplicates = 0;
    const errors = [];

    records.forEach(function(row, index) {
      const rowNumber = index + 1;
      try {
        const studentId = normalizePerformanceText(row && row.studentId);
        const studentName = normalizePerformanceText(row && row.studentName);
        const performanceId = normalizePerformanceText(row && row.performanceId);
        const rawClassScore = row && row.classScore;
        const rawExamScore100 = row && row.examScore100;
        const hasClassScore = rawClassScore !== '' && rawClassScore != null;
        const hasExamScore = rawExamScore100 !== '' && rawExamScore100 != null;

        if (!studentId || !studentName) {
          skipped++;
          return;
        }

        if (!hasClassScore && !hasExamScore) {
          skipped++;
          return;
        }

        if (!hasClassScore || isNaN(Number(rawClassScore))) {
          failed++;
          errors.push(`Row ${rowNumber} (${studentName}): Invalid Class Score.`);
          return;
        }

        if (!hasExamScore || isNaN(Number(rawExamScore100))) {
          failed++;
          errors.push(`Row ${rowNumber} (${studentName}): Invalid Exam Score (100).`);
          return;
        }

        const derived = calculatePerformanceDerivedScores(rawClassScore, rawExamScore100);
        if (derived.classScore < 0 || derived.classScore > derived.maxClassScore) {
          failed++;
          errors.push(`Row ${rowNumber} (${studentName}): Class Score must be between 0 and ${derived.maxClassScore}.`);
          return;
        }
        if (derived.examScore100 < 0 || derived.examScore100 > 100) {
          failed++;
          errors.push(`Row ${rowNumber} (${studentName}): Exam Score must be between 0 and 100.`);
          return;
        }

        const perfPayload = {
          studentId: studentId,
          studentName: studentName,
          studentClass: studentClass,
          term: term,
          academicYear: academicYear,
          course: course,
          classScore: derived.classScore,
          examScore100: derived.examScore100,
          examScore60: derived.examScore60,
          examScore50: derived.examScore50,
          total: derived.total
        };
        if (bulkGuard.adminOverride) perfPayload.__adminAcademicYearOverride = true;

        let result;
        if (performanceId) {
          result = updatePerformance(performanceId, perfPayload);
        } else {
          result = addPerformance(perfPayload);
        }

        if (result && result.success) {
          if (performanceId) updated++;
          else created++;
          return;
        }

        if (result && result.requiresAdminOverride) {
          throw new Error(result.message || 'Admin confirmation required.');
        }

        if (result && result.message && /duplicate/i.test(result.message)) {
          duplicates++;
        } else {
          failed++;
        }
        errors.push(`Row ${rowNumber} (${studentName}): ${(result && result.message) ? result.message : 'Save failed.'}`);
      } catch (error) {
        failed++;
        errors.push(`Row ${rowNumber}: ${error.message}`);
      }
    });

    try {
      calculateAndSetRanks(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Performance'), studentClass, term, academicYear, course);
    } catch (rankError) {
      Logger.log('Bulk performance rank recalculation failed: ' + rankError.message);
    }

    safeLogAuditEvent(
      'Update',
      'Performance',
      `${studentClass}|${term}|${academicYear}|${course}`,
      appendAcademicYearOverrideAuditDetails('Bulk saved performance records', bulkGuard),
      null,
      buildAuditSnapshot({
        studentClass: studentClass,
        term: term,
        academicYear: academicYear,
        course: course,
        created: created,
        updated: updated,
        skipped: skipped,
        failed: failed,
        duplicates: duplicates,
        totalRows: records.length
      })
    );

    return {
      success: true,
      created: created,
      updated: updated,
      skipped: skipped,
      failed: failed,
      duplicates: duplicates,
      errors: errors
    };
  } catch (error) {
    Logger.log('ERROR in saveBulkPerformanceEntries: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Recalculate ranks for ALL performance records
 * Call this once to fix existing records that don't have ranks
 */
function recalculateAllRanks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Performance");
  if (!sheet) {
    Logger.log('Performance sheet not found');
    return { success: false, message: 'Performance sheet not found' };
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    Logger.log('No records to process');
    return { success: true, message: 'No records found' };
  }
  
  // Get headers
  const headers = sheet.getRange(2, 2, 1, 12).getValues()[0].map(h => String(h).trim().toLowerCase());
  
  function findCol(aliases) {
    for (let i = 0; i < headers.length; i++) {
      for (let a of aliases) {
        if (headers[i] === a || headers[i].indexOf(a) !== -1) return i + 2;
      }
    }
    return -1;
  }
  
  const colClass = findCol(['student class', 'class name', 'class']);
  const colTerm = findCol(['term']);
  const colYear = findCol(['academic year', 'year']);
  const colCourse = findCol(['course']);
  const colTotal = findCol(['total']);
  const colRank = findCol(['rank']);
  
  if (colClass === -1 || colTerm === -1 || colYear === -1 || colCourse === -1 || colTotal === -1 || colRank === -1) {
    Logger.log('Cannot recalculate ranks - missing required columns');
    return { success: false, message: 'Missing required columns in Performance sheet' };
  }
  
  // Get all data
  const allData = sheet.getRange(3, 2, lastRow - 2, headers.length).getValues();
  
  // Group records by class/term/year/course
  const groups = {};
  
  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const rowClass = String(row[colClass - 2] || '').trim();
    const rowTerm = String(row[colTerm - 2] || '').trim();
    const rowYear = String(row[colYear - 2] || '').trim();
    const rowCourse = String(row[colCourse - 2] || '').trim();
    const rowTotal = Number(row[colTotal - 2]) || 0;
    
    if (!rowClass || !rowTerm || !rowYear || !rowCourse) continue;
    
    const groupKey = `${rowClass}|${rowTerm}|${rowYear}|${rowCourse}`;
    
    if (!groups[groupKey]) {
      groups[groupKey] = [];
    }
    
    groups[groupKey].push({
      rowNumber: i + 3,
      total: rowTotal
    });
  }
  
  // Calculate ranks for each group
  let totalRecordsRanked = 0;
  
  Object.keys(groups).forEach(function(groupKey) {
    const records = groups[groupKey];
    
    // Sort by total (descending)
    records.sort((a, b) => b.total - a.total);
    
    // Assign ranks (handle ties)
    let currentRank = 1;
    for (let i = 0; i < records.length; i++) {
      if (i > 0 && records[i].total < records[i - 1].total) {
        currentRank = i + 1;
      }
      sheet.getRange(records[i].rowNumber, colRank).setValue(currentRank);
      totalRecordsRanked++;
    }
    
    Logger.log(`Ranked ${records.length} students in group: ${groupKey}`);
  });
  
  Logger.log(`Total records ranked: ${totalRecordsRanked} across ${Object.keys(groups).length} groups`);
  
  return {
    success: true,
    message: `Successfully ranked ${totalRecordsRanked} records across ${Object.keys(groups).length} course groups`,
    totalRecords: totalRecordsRanked,
    totalGroups: Object.keys(groups).length
  };
}

/**
 * Clean up existing duplicate performance records in the Performance sheet.
 * Retains the first occurrence for each (Student + Class + Term + Academic Year + Course) combination
 * and deletes any duplicate rows below it.
 */
function removeDuplicatePerformanceRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Performance");
  if (!sheet) return { success: false, message: 'Performance sheet not found' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return { success: true, message: 'No performance records found', deletedCount: 0 };

  const headers = sheet.getRange(2, 2, 1, 12).getValues()[0].map(h => String(h).trim().toLowerCase());
  
  function findCol(aliases) {
    for (let i = 0; i < headers.length; i++) {
      for (let a of aliases) {
        if (headers[i] === a || headers[i].indexOf(a) !== -1) return i;
      }
    }
    return -1;
  }

  const idxName = findCol(["student name", "name", "student"]);
  const idxStudentId = findCol(["student id", "studentid", "student_id", "sid"]);
  const idxClass = findCol(["student class", "class name", "class"]);
  const idxTerm = findCol(["term"]);
  const idxYear = findCol(["academic year", "year"]);
  const idxCourse = findCol(["course"]);

  const allVals = sheet.getRange(3, 2, lastRow - 2, headers.length).getValues();
  const normalize = (str) => String(str || "").trim().toLowerCase();

  const seenKeys = new Set();
  const rowsToDelete = [];

  for (let i = 0; i < allVals.length; i++) {
    const rowNum = i + 3;
    const row = allVals[i];

    const studentName = idxName !== -1 ? normalize(row[idxName]) : "";
    const studentId = idxStudentId !== -1 ? normalize(row[idxStudentId]) : "";
    const className = idxClass !== -1 ? normalize(row[idxClass]) : "";
    const term = idxTerm !== -1 ? normalize(row[idxTerm]) : "";
    const year = idxYear !== -1 ? normalize(row[idxYear]) : "";
    const course = idxCourse !== -1 ? normalize(row[idxCourse]) : "";

    if (!studentName && !studentId) continue;

    const studentIdentifier = studentId !== "" ? studentId : studentName;
    const key = `${studentIdentifier}|${className}|${term}|${year}|${course}`;

    if (seenKeys.has(key)) {
      rowsToDelete.push(rowNum);
    } else {
      seenKeys.add(key);
    }
  }

  const deletedRecords = rowsToDelete.map(function(rowNum) {
    const row = allVals[rowNum - 3] || [];
    return buildAuditSnapshot({
      performanceId: 'ROW-' + rowNum,
      studentId: idxStudentId !== -1 ? String(row[idxStudentId] || '').trim() : '',
      studentName: idxName !== -1 ? String(row[idxName] || '').trim() : '',
      studentClass: idxClass !== -1 ? String(row[idxClass] || '').trim() : '',
      term: idxTerm !== -1 ? String(row[idxTerm] || '').trim() : '',
      academicYear: idxYear !== -1 ? String(row[idxYear] || '').trim() : '',
      course: idxCourse !== -1 ? String(row[idxCourse] || '').trim() : ''
    });
  });

  // Delete duplicate rows from bottom to top to preserve row indexing
  rowsToDelete.reverse().forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });

  if (rowsToDelete.length > 0) {
    try { recalculateAllRanks(); } catch (e) { /* ignore */ }
    try { invalidatePerformanceCache(); } catch (e) { /* ignore */ }

    safeLogAuditEvent(
      'Cleanup',
      'Performance',
      'Duplicates',
      'Removed duplicate performance records',
      buildAuditSnapshot({ duplicateCount: rowsToDelete.length, deletedRecords: deletedRecords }),
      buildAuditSnapshot({ duplicateCount: 0 })
    );
  }

  return {
    success: true,
    message: `Removed ${rowsToDelete.length} duplicate performance record(s).`,
    deletedCount: rowsToDelete.length
  };
}

/**
 * Calculate and set ranks for students in a specific class/term/year/course
 * Ranks students by total score (highest first)
 */
function calculateAndSetRanks(sheet, className, term, academicYear, courseName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  
  // Get headers
  const headers = sheet.getRange(2, 2, 1, 12).getValues()[0].map(h => String(h).trim().toLowerCase());
  
  function findCol(aliases) {
    for (let i = 0; i < headers.length; i++) {
      for (let a of aliases) {
        if (headers[i] === a || headers[i].indexOf(a) !== -1) return i + 2;
      }
    }
    return -1;
  }
  
  const colClass = findCol(['student class', 'class name', 'class']);
  const colTerm = findCol(['term']);
  const colYear = findCol(['academic year', 'year']);
  const colCourse = findCol(['course']);
  const colTotal = findCol(['total']);
  const colRank = findCol(['rank']);
  
  if (colClass === -1 || colTerm === -1 || colYear === -1 || colCourse === -1 || colTotal === -1 || colRank === -1) {
    Logger.log('Cannot calculate ranks - missing required columns');
    return;
  }
  
  // Get all data
  const allData = sheet.getRange(3, 2, lastRow - 2, headers.length).getValues();
  
  // Filter records matching class/term/year/course
  const matchingRecords = [];
  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const rowClass = String(row[colClass - 2] || '').trim();
    const rowTerm = String(row[colTerm - 2] || '').trim();
    const rowYear = String(row[colYear - 2] || '').trim();
    const rowCourse = String(row[colCourse - 2] || '').trim();
    const rowTotal = Number(row[colTotal - 2]) || 0;
    
    if (rowClass === className && rowTerm === term && rowYear === academicYear && rowCourse === courseName) {
      matchingRecords.push({
        rowNumber: i + 3,
        total: rowTotal
      });
    }
  }
  
  // Sort by total (descending)
  matchingRecords.sort((a, b) => b.total - a.total);
  
  // Assign ranks (handle ties)
  let currentRank = 1;
  for (let i = 0; i < matchingRecords.length; i++) {
    if (i > 0 && matchingRecords[i].total < matchingRecords[i - 1].total) {
      currentRank = i + 1;
    }
    sheet.getRange(matchingRecords[i].rowNumber, colRank).setValue(currentRank);
  }
  
  Logger.log(`Ranks calculated for ${matchingRecords.length} students in ${courseName}`);
}

// Add new Performance Record dynamically looking up header columns
function addPerformance(perfData) {
  try {
    Logger.log('addPerformance called with data: ' + JSON.stringify(perfData));
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Performance");
    if (!sheet) throw new Error("Performance worksheet not found.");

    const lastRow = sheet.getLastRow();

    // Let's get header columns to know where to write what (from B to M)
    const headers = sheet
      .getRange(2, 2, 1, 12)
      .getValues()[0]
      .map((h) => String(h).trim().toLowerCase());

    // Helper to locate header column (1-based sheet column number)
    function findColNum(aliasGroups) {
      // exact match first
      for (let i = 0; i < headers.length; i++) {
        for (let a = 0; a < aliasGroups.length; a++) {
          if (headers[i] === aliasGroups[a]) return i + 2;
        }
      }
      // contains match next
      for (let i = 0; i < headers.length; i++) {
        for (let a = 0; a < aliasGroups.length; a++) {
          if (headers[i].indexOf(aliasGroups[a]) !== -1) return i + 2;
        }
      }
      return -1;
    }

    const colName = findColNum(["student name", "name", "student"]);
    const colStudentId = findColNum([
      "student id",
      "studentid",
      "student_id",
      "sid",
    ]);
    const colClass = findColNum(["student class", "class name", "class"]);
    const colTerm = findColNum(["term"]);
    const colYear = findColNum(["academic year", "year"]);
    const colClassScore = findColNum(["class score", "classscore"]);
    const colExam100 = findColNum([
      "exam score (100%)",
      "exam score (100)",
      "exam 100",
      "exam score 100",
      "exam100",
    ]);
    const colExam60 = findColNum([
      "exam score (50%)",
      "exam score (50)",
      "exam 50",
      "exam score 50",
      "exam score (60%)",
      "exam 60",
      "exam score 60",
      "exam60",
    ]);
    const colTotal = findColNum(["total"]);
    const colCourse = findColNum(["course"]);

    // Server-side duplicate check: prevent same student/course/class/term/year
    if (lastRow >= 3) {
      const allVals = sheet.getRange(3, 2, lastRow - 2, headers.length).getValues();
      
      // Normalize function for consistent comparison
      const normalize = (str) => String(str || "").trim().toLowerCase();
      
      const incomingStudentId = normalize(perfData.studentId);
      const incomingStudentName = normalize(perfData.studentName);
      const incomingClass = normalize(perfData.studentClass);
      const incomingTerm = normalize(perfData.term);
      const incomingYear = normalize(perfData.academicYear);
      const incomingCourse = normalize(perfData.course);
      
      Logger.log('Checking for duplicates with: StudentID=' + incomingStudentId + ', Name=' + incomingStudentName + ', Class=' + incomingClass + ', Term=' + incomingTerm + ', Year=' + incomingYear + ', Course=' + incomingCourse);
      
      for (let i = 0; i < allVals.length; i++) {
        const row = allVals[i];
        const valAt = (col) => {
          if (col === -1) return "";
          const idx = col - 2;
          return String(row[idx] || "").trim();
        };
        
        const existingStudentId = normalize(valAt(colStudentId));
        const existingName = normalize(valAt(colName));
        const existingClass = normalize(valAt(colClass));
        const existingTerm = normalize(valAt(colTerm));
        const existingYear = normalize(valAt(colYear));
        const existingCourse = normalize(valAt(colCourse));

        // Match by student ID first (if both non-empty), otherwise fallback to student name
        const sameStudent = (incomingStudentId !== "" && existingStudentId !== "")
          ? existingStudentId === incomingStudentId
          : existingName === incomingStudentName;

        const sameClass = existingClass === incomingClass;
        const sameTerm = existingTerm === incomingTerm;
        const sameYear = existingYear === incomingYear;
        const sameCourse = existingCourse === incomingCourse;

        if (sameStudent && sameClass && sameTerm && sameYear && sameCourse) {
          Logger.log('DUPLICATE FOUND at row ' + (i + 3) + ': ExistingID=' + existingStudentId + ', ExistingName=' + existingName + ', Course=' + existingCourse);
          return { 
            success: false, 
            message: `Duplicate record: ${perfData.studentName} already has a performance record for ${perfData.course} in ${perfData.term} ${perfData.academicYear}.` 
          };
        }
      }
      Logger.log('No duplicate found - record will be added');
    }

    const targetRow = lastRow + 1;
    const performanceId = `ROW-${targetRow}`;
    const pendingPerformanceSnapshot = buildPerformanceAuditSnapshot({
      performanceId: performanceId,
      studentId: perfData.studentId || perfData.studentName || '',
      studentName: perfData.studentName,
      studentClass: perfData.studentClass,
      term: perfData.term,
      academicYear: perfData.academicYear,
      classScore: perfData.classScore,
      examScore100: perfData.examScore100,
      examScore60: perfData.examScore50 || perfData.examScore60 || 0,
      total: perfData.total,
      course: perfData.course
    });
    const createPerformanceYearGuard = enforceAcademicYearCrudSecurity({
      action: 'Create',
      module: 'Performance',
      recordId: performanceId,
      academicYear: perfData.academicYear,
      newValue: pendingPerformanceSnapshot,
      overrideConfirmed: perfData && perfData.__adminAcademicYearOverride === true
    });
    if (!createPerformanceYearGuard.allowed) return createPerformanceYearGuard;

    // Write values to their respective columns
    if (colName !== -1)
      sheet.getRange(targetRow, colName).setValue(perfData.studentName);
    if (colStudentId !== -1)
      sheet
        .getRange(targetRow, colStudentId)
        .setValue(perfData.studentId || perfData.studentName || "");
    if (colClass !== -1)
      sheet.getRange(targetRow, colClass).setValue(perfData.studentClass);
    if (colTerm !== -1)
      sheet.getRange(targetRow, colTerm).setValue(perfData.term);
    if (colYear !== -1)
      sheet.getRange(targetRow, colYear).setValue(perfData.academicYear);
    if (colClassScore !== -1)
      sheet.getRange(targetRow, colClassScore).setValue(perfData.classScore);
    if (colExam100 !== -1)
      sheet.getRange(targetRow, colExam100).setValue(perfData.examScore100);
    if (colExam60 !== -1)
      sheet.getRange(targetRow, colExam60).setValue(perfData.examScore50 || perfData.examScore60 || 0);
    if (colTotal !== -1)
      sheet.getRange(targetRow, colTotal).setValue(perfData.total);
    if (colCourse !== -1)
      sheet.getRange(targetRow, colCourse).setValue(perfData.course);
    
    // Calculate and set rank after adding the record
    try {
      calculateAndSetRanks(sheet, perfData.studentClass, perfData.term, perfData.academicYear, perfData.course);
    } catch (rankError) {
      Logger.log('Error calculating ranks: ' + rankError.message);
      // Continue anyway - rank can be calculated later
    }

    Logger.log('Performance record added successfully at row: ' + targetRow);
    try { invalidatePerformanceCache(); } catch (e) { /* fail silently */ }

    const createdPerformance = getPerformanceDataFromSheet().find((record) => record.rowNumber === targetRow) || buildPerformanceAuditSnapshot({
      performanceId: performanceId,
      studentId: perfData.studentId || perfData.studentName || '',
      studentName: perfData.studentName,
      studentClass: perfData.studentClass,
      term: perfData.term,
      academicYear: perfData.academicYear,
      classScore: perfData.classScore,
      examScore100: perfData.examScore100,
      examScore60: perfData.examScore50 || perfData.examScore60 || 0,
      total: perfData.total,
      course: perfData.course
    });

    safeLogAuditEvent(
      'Create',
      'Performance',
      performanceId,
      appendAcademicYearOverrideAuditDetails('Created performance record for ' + String(perfData.studentName || performanceId).trim(), createPerformanceYearGuard),
      null,
      buildPerformanceAuditSnapshot(createdPerformance)
    );

    return { success: true, performanceId: performanceId };
    
  } catch (error) {
    Logger.log('ERROR in addPerformance: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Helper: find performance row by performanceId
function findPerformanceRowById(sheet, performanceId) {
  if (!sheet) return -1;
  
  // New format: performanceId is "ROW-X" where X is the actual row number
  if (String(performanceId).startsWith('ROW-')) {
    const rowNum = parseInt(String(performanceId).replace('ROW-', ''), 10);
    if (!isNaN(rowNum) && rowNum >= 3) {
      return rowNum;
    }
  }
  
  // Fallback: old behavior (shouldn't happen with new format)
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return -1;
  
  // Try to find by student name as fallback
  Logger.log('Warning: Could not parse row from performanceId: ' + performanceId);
  return -1;
}

// Update Performance record by ID
function updatePerformance(performanceId, perfData) {
  try {
    Logger.log('updatePerformance - ID: ' + performanceId + ', Data: ' + JSON.stringify(perfData));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Performance");
  if (!sheet)
    return { success: false, message: "Performance sheet not found." };

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return { success: false, message: "No records." };

  // read headers and build helper findColNum (same logic as addPerformance)
  const headers = sheet
    .getRange(2, 2, 1, 12)
    .getValues()[0]
    .map((h) => String(h).trim().toLowerCase());
  function findColNumLocal(aliasGroups) {
    for (let i = 0; i < headers.length; i++) {
      for (let a = 0; a < aliasGroups.length; a++) {
        if (headers[i] === aliasGroups[a]) return i + 2;
      }
    }
    for (let i = 0; i < headers.length; i++) {
      for (let a = 0; a < aliasGroups.length; a++) {
        if (headers[i].indexOf(aliasGroups[a]) !== -1) return i + 2;
      }
    }
    return -1;
  }

  const existingPerformance = getPerformanceDataFromSheet().find((record) => String(record.performanceId).trim() === String(performanceId).trim()) || null;
  const row = findPerformanceRowById(sheet, performanceId);
  if (row === -1) return { success: false, message: "Record not found." };

  const colName = findColNumLocal(["student name", "name", "student"]);
  const colStudentId = findColNumLocal([
    "student id",
    "studentid",
    "student_id",
    "sid",
  ]);
  const colClass = findColNumLocal(["student class", "class name", "class"]);
  const colTerm = findColNumLocal(["term"]);
  const colYear = findColNumLocal(["academic year", "year"]);
  const colClassScore = findColNumLocal(["class score", "classscore"]);
  const colExam100 = findColNumLocal([
    "exam score (100%)",
    "exam score (100)",
    "exam 100",
    "exam score 100",
    "exam100",
  ]);
  const colExam60 = findColNumLocal([
    "exam score (50%)",
    "exam score (50)",
    "exam 50",
    "exam score 50",
    "exam50",
    "exam score (60%)",
    "exam 60",
    "exam score 60",
    "exam60",
  ]);
  const colTotal = findColNumLocal(["total"]);
  const colRank = findColNumLocal(["rank"]);
  const colCourse = findColNumLocal(["course"]);

  const pendingUpdatedPerformance = buildPerformanceAuditSnapshot(Object.assign({}, existingPerformance || {}, {
    performanceId: performanceId,
    studentId: perfData.studentId || perfData.studentName || '',
    studentName: perfData.studentName || '',
    studentClass: perfData.studentClass || '',
    term: perfData.term || '',
    academicYear: perfData.academicYear || '',
    classScore: perfData.classScore || 0,
    examScore100: perfData.examScore100 || 0,
    examScore60: perfData.examScore60 !== undefined ? perfData.examScore60 : perfData.examScore50 || 0,
    total: perfData.total || 0,
    rank: perfData.rank || '',
    course: perfData.course || ''
  }));
  const updatePerformanceYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Update',
    module: 'Performance',
    recordId: performanceId,
    academicYears: [existingPerformance && existingPerformance.academicYear, perfData.academicYear],
    oldValue: buildPerformanceAuditSnapshot(existingPerformance),
    newValue: pendingUpdatedPerformance,
    overrideConfirmed: perfData && perfData.__adminAcademicYearOverride === true
  });
  if (!updatePerformanceYearGuard.allowed) return updatePerformanceYearGuard;

  // Check for duplicate before applying update (ignore current row)
  const dupRow = findPerformanceDuplicateLocal(perfData, row);
  if (dupRow !== -1) {
    return { success: false, message: "Another performance record already exists for this student in the same course, class, term and academic year." };
  }

  // Server-side duplicate check for update: ignore the row being updated
  function findPerformanceDuplicateLocal(data, ignoreRow) {
    const lastRowCheck = sheet.getLastRow();
    if (lastRowCheck < 3) return -1;
    const allVals = sheet.getRange(3, 2, lastRowCheck - 2, headers.length).getValues();
    const normalize = (str) => String(str || "").trim().toLowerCase();

    const incomingStudentId = normalize(data.studentId);
    const incomingStudentName = normalize(data.studentName);
    const incomingClass = normalize(data.studentClass);
    const incomingTerm = normalize(data.term);
    const incomingYear = normalize(data.academicYear);
    const incomingCourse = normalize(data.course);

    for (let i = 0; i < allVals.length; i++) {
      const rowNum = i + 3;
      if (ignoreRow && rowNum === ignoreRow) continue;
      const row = allVals[i];
      const valAt = (col) => {
        if (col === -1) return "";
        const idx = col - 2;
        return String(row[idx] || "").trim();
      };
      const existingStudentId = normalize(valAt(colStudentId));
      const existingName = normalize(valAt(colName));
      const existingClass = normalize(valAt(colClass));
      const existingTerm = normalize(valAt(colTerm));
      const existingYear = normalize(valAt(colYear));
      const existingCourse = normalize(valAt(colCourse));

      const sameStudent = (incomingStudentId !== "" && existingStudentId !== "")
        ? existingStudentId === incomingStudentId
        : existingName === incomingStudentName;

      const sameClass = existingClass === incomingClass;
      const sameTerm = existingTerm === incomingTerm;
      const sameYear = existingYear === incomingYear;
      const sameCourse = existingCourse === incomingCourse;

      if (sameStudent && sameClass && sameTerm && sameYear && sameCourse) {
        return rowNum;
      }
    }
    return -1;
  }

  if (colName !== -1)
    sheet.getRange(row, colName).setValue(perfData.studentName || "");
  if (colStudentId !== -1)
    sheet
      .getRange(row, colStudentId)
      .setValue(perfData.studentId || perfData.studentName || "");
  if (colClass !== -1)
    sheet.getRange(row, colClass).setValue(perfData.studentClass || "");
  if (colTerm !== -1)
    sheet.getRange(row, colTerm).setValue(perfData.term || "");
  if (colYear !== -1)
    sheet.getRange(row, colYear).setValue(perfData.academicYear || "");
  if (colClassScore !== -1)
    sheet.getRange(row, colClassScore).setValue(perfData.classScore || 0);
  if (colExam100 !== -1)
    sheet.getRange(row, colExam100).setValue(perfData.examScore100 || 0);
  if (colExam60 !== -1)
    sheet.getRange(row, colExam60).setValue(
      perfData.examScore60 !== undefined
        ? perfData.examScore60
        : perfData.examScore50 || 0,
    );
  if (colTotal !== -1)
    sheet.getRange(row, colTotal).setValue(perfData.total || 0);
  if (colRank !== -1)
    sheet.getRange(row, colRank).setValue(perfData.rank || "");
  if (colCourse !== -1)
    sheet.getRange(row, colCourse).setValue(perfData.course || "");

  // Recalculate ranks after update
  try {
    calculateAndSetRanks(sheet, perfData.studentClass, perfData.term, perfData.academicYear, perfData.course);
  } catch (rankError) {
    Logger.log('Error calculating ranks: ' + rankError.message);
  }

  Logger.log('Update successful');
  try { invalidatePerformanceCache(); } catch (e) { /* fail silently */ }

  const updatedPerformance = getPerformanceDataFromSheet().find((record) => String(record.performanceId).trim() === String(performanceId).trim()) || buildPerformanceAuditSnapshot(Object.assign({}, existingPerformance || {}, {
    performanceId: performanceId,
    studentId: perfData.studentId || perfData.studentName || '',
    studentName: perfData.studentName || '',
    studentClass: perfData.studentClass || '',
    term: perfData.term || '',
    academicYear: perfData.academicYear || '',
    classScore: perfData.classScore || 0,
    examScore100: perfData.examScore100 || 0,
    examScore60: perfData.examScore60 !== undefined ? perfData.examScore60 : perfData.examScore50 || 0,
    total: perfData.total || 0,
    rank: perfData.rank || '',
    course: perfData.course || ''
  }));

  safeLogAuditEvent(
    'Update',
    'Performance',
    performanceId,
    appendAcademicYearOverrideAuditDetails('Updated performance record for ' + String((updatedPerformance && updatedPerformance.studentName) || performanceId).trim(), updatePerformanceYearGuard),
    buildPerformanceAuditSnapshot(existingPerformance),
    buildPerformanceAuditSnapshot(updatedPerformance)
  );

  return { success: true };
  } catch (error) {
    Logger.log('ERROR in updatePerformance: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Delete Performance record by ID
function deletePerformance(performanceId) {
  try {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Performance");
  if (!sheet)
    return { success: false, message: "Performance sheet not found." };
  const existingPerformance = getPerformanceDataFromSheet().find((record) => String(record.performanceId).trim() === String(performanceId).trim()) || null;
  const row = findPerformanceRowById(sheet, performanceId);
  if (row === -1) return { success: false, message: "Record not found." };
  const deletePerformanceYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Delete',
    module: 'Performance',
    recordId: performanceId,
    academicYear: existingPerformance && existingPerformance.academicYear,
    oldValue: buildPerformanceAuditSnapshot(existingPerformance),
    overrideConfirmed: arguments[1] && arguments[1].adminAcademicYearOverride === true
  });
  if (!deletePerformanceYearGuard.allowed) return deletePerformanceYearGuard;
  sheet.deleteRow(row);
  try { invalidatePerformanceCache(); } catch (e) { /* fail silently */ }

  safeLogAuditEvent(
    'Delete',
    'Performance',
    performanceId,
    appendAcademicYearOverrideAuditDetails('Deleted performance record for ' + String(((existingPerformance && existingPerformance.studentName) || performanceId)).trim(), deletePerformanceYearGuard),
    buildPerformanceAuditSnapshot(existingPerformance),
    null
  );

  return { success: true };
  } catch (error) {
    Logger.log('ERROR in deletePerformance: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}


// ============================================
// TERMINAL REPORT GENERATION
// ============================================

function generateTerminalReports(reportData) {
  requireLogin();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const students = reportData.students || [];
  
  if (students.length === 0) {
    return { success: false, message: 'No students selected' };
  }
  
  try {
    // If single student, generate single PDF
    if (students.length === 1) {
      const result = generateSingleTerminalReport(reportData, students[0]);
      return {
        success: true,
        base64: result.base64,
        fileName: result.fileName,
      };
    }

    // If multiple students, generate ZIP
    const result = generateMultipleTerminalReports(reportData, students);
    return {
      success: true,
      base64: result.base64,
      fileName: result.fileName,
      isZip: true,
    };
  } catch (error) {
    Logger.log('Error generating terminal reports: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function generateSingleTerminalReport(reportData, student) {
  try {
    Logger.log('=== generateSingleTerminalReport START ===');
    Logger.log('Report Data: ' + JSON.stringify(reportData));
    Logger.log('Student: ' + JSON.stringify(student));
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get student performance data
    const performanceSheet = ss.getSheetByName('Performance');
    if (!performanceSheet) throw new Error('Performance sheet not found');
    
    const rawPerformanceData = typeof getPerformanceDataFromSheet === 'function' ? getPerformanceDataFromSheet() : getPerformanceData();
    const performanceData = (rawPerformanceData || []).map(function(p) {
      return {
        performanceId: typeof decodeSanitizedHtml === 'function' ? decodeSanitizedHtml(p.performanceId) : p.performanceId,
        studentId: typeof decodeSanitizedHtml === 'function' ? decodeSanitizedHtml(p.studentId) : p.studentId,
        studentName: typeof decodeSanitizedHtml === 'function' ? decodeSanitizedHtml(p.studentName) : p.studentName,
        studentClass: typeof decodeSanitizedHtml === 'function' ? decodeSanitizedHtml(p.studentClass) : p.studentClass,
        academicYear: typeof decodeSanitizedHtml === 'function' ? decodeSanitizedHtml(p.academicYear) : p.academicYear,
        term: typeof decodeSanitizedHtml === 'function' ? decodeSanitizedHtml(p.term) : p.term,
        course: typeof decodeSanitizedHtml === 'function' ? decodeSanitizedHtml(p.course) : p.course,
        classScore: p.classScore,
        examScore100: p.examScore100,
        examScore60: p.examScore60,
        examScore50: p.examScore50,
        total: p.total,
        rank: typeof decodeSanitizedHtml === 'function' ? decodeSanitizedHtml(p.rank) : p.rank
      };
    });
    Logger.log('Total performance records: ' + performanceData.length);
    
    // Improved matching: normalize IDs and names, require class/term/year, prefer ID, then name equality, then substring fallback
    function normText(s) { return String(s || '').toLowerCase().trim(); }
    function normId(s) { return String(s || '').replace(/[^a-z0-9]/gi, '').toLowerCase().trim(); }

    const studentPerformance = performanceData.filter(p => {
      const classMatch = normText(p.studentClass) === normText(reportData.studentClass);
      const yearMatch = normText(p.academicYear) === normText(reportData.academicYear);
      const termMatch = normText(p.term) === normText(reportData.term);
      if (!classMatch || !yearMatch || !termMatch) return false;

      // Try ID match if available
      if (student.studentId && String(student.studentId).trim() !== '') {
        if (p.studentId && String(p.studentId).trim() !== '') {
          if (normId(p.studentId) === normId(student.studentId)) {
            Logger.log('MATCH FOUND by ID: ' + JSON.stringify(p));
            return true;
          }
        }
        // fall through to name checks
      }

      const perfName = normText(p.studentName);
      const targetName = normText(student.studentName || '');
      if (perfName === targetName) {
        Logger.log('MATCH FOUND by exact name: ' + JSON.stringify(p));
        return true;
      }
      // substring fallback
      if (perfName.indexOf(targetName) !== -1 || targetName.indexOf(perfName) !== -1) {
        Logger.log('MATCH FOUND by substring: ' + JSON.stringify(p));
        return true;
      }

      return false;
    });
    
    Logger.log('Filtered student performance records: ' + studentPerformance.length);
    if (studentPerformance.length === 0) {
      Logger.log('WARNING: No performance records found for student!');
      Logger.log('Search criteria - Name: ' + student.studentName + ', ID: ' + student.studentId + ', Class: ' + reportData.studentClass + ', Term: ' + reportData.term + ', Year: ' + reportData.academicYear);
    }

  // Compute overall position (POS) for this student among peers in same class/term/year
  try {
    const allPerf = performanceData.filter(p => p.studentClass === reportData.studentClass && p.academicYear === reportData.academicYear && p.term === reportData.term);
    const grouped = {};
    allPerf.forEach(p => {
      const key = (p.studentId && String(p.studentId).trim()) || (p.studentName || '').toLowerCase().trim();
      if (!grouped[key]) grouped[key] = { studentName: p.studentName || '', total: 0 };
      grouped[key].total = (grouped[key].total || 0) + (Number(p.total) || 0);
    });
    const standings = Object.keys(grouped).map(k => ({ key: k, studentName: grouped[k].studentName, total: grouped[k].total }));
    standings.sort((a, b) => { const d = b.total - a.total; if (d !== 0) return d; return (a.studentName || '').localeCompare(b.studentName || ''); });
    const studentKey = (student.studentId && String(student.studentId).trim()) || (student.studentName || '').toLowerCase().trim();
    let overallPosition = '-';
    for (let i = 0; i < standings.length; i++) {
      if (standings[i].key === studentKey) { overallPosition = String(i + 1); break; }
    }
    reportData.overallPosition = overallPosition;
  } catch (e) {
    Logger.log('Error computing overall position: ' + e.toString());
    reportData.overallPosition = '-';
  }
  
  // Get school information from Settings sheet
  const schoolName = getSystemParameter('School Name') || 'GLOBAL EVANGELICAL BASIC SCHOOL, TETTEKOPE';
  const schoolAddress = getSystemParameter('School Address') || getSystemParameter('Address') || 'P.O. BOX KW 182, KETA';
  const schoolEmail = getSystemParameter('School Email') || getSystemParameter('Email') || '';
  const schoolPhone = getSystemParameter('School Phone') || getSystemParameter('Phone') || '';
  const rawSchoolLogo = getSystemParameter('School Logo') || getSystemParameter('Logo URL') || 'https://drive.google.com/uc?export=view&id=1MVnH55BHBynLBOD4pLIZE-5Y9gMaJbOe';
  const schoolLogo = getImageAsBase64(rawSchoolLogo);
  
  Logger.log('School Info - Name: ' + schoolName + ', Address: ' + schoolAddress + ', Email: ' + schoolEmail);
  
  // Get student info
  const studentsSheet = ss.getSheetByName('Students');
  let studentInfo = { fullName: student.studentName, firstName: '', lastName: '' };
  if (studentsSheet) {
    const studentsData = getStudentsData();
    let found = null;
    // Prefer lookup by studentId if available
    if (student.studentId && String(student.studentId).trim() !== '') {
      found = studentsData.find(s => String(s.studentId || '').trim() === String(student.studentId).trim());
    }
    // Fallback to name match
    if (!found) {
      found = studentsData.find(s => {
        const fullName = `${s.firstName} ${s.lastName}`.trim();
        return fullName.toLowerCase() === (student.studentName || '').toLowerCase().trim();
      });
    }
    if (found) {
      studentInfo = found;
      studentInfo.fullName = `${found.firstName} ${found.lastName}`.trim();
    }
  }
  
  // Generate HTML report
  const html = buildTerminalReportHTML(schoolName, schoolAddress, schoolEmail, schoolPhone, schoolLogo, studentInfo, reportData, studentPerformance);
  
  Logger.log('HTML generated, length: ' + html.length);
  
  // Convert to PDF and return base64 (avoids Drive permission requirement)
  const fileName = `${student.studentName}_${reportData.term}_${reportData.academicYear}_Report.pdf`;
  const blob = Utilities.newBlob(html, 'text/html', 'report.html').getAs('application/pdf');
  blob.setName(fileName);

  Logger.log('=== generateSingleTerminalReport END ===');
  
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    fileName: fileName,
    blob: blob,
  };
  
  } catch (error) {
    Logger.log('ERROR in generateSingleTerminalReport: ' + error.toString());
    throw error;
  }
}

function generateMultipleTerminalReports(reportData, students) {
  const blobs = students.map(function (student) {
    return generateSingleTerminalReport(reportData, student).blob;
  });

  const fileName = `Terminal_Reports_${reportData.studentClass}_${reportData.term}_${reportData.academicYear}.zip`;
  const zipBlob = Utilities.zip(blobs, fileName);

  return {
    base64: Utilities.base64Encode(zipBlob.getBytes()),
    fileName: fileName,
  };
}

function buildTerminalReportHTML(schoolName, schoolAddress, schoolEmail, schoolPhone, schoolLogo, studentInfo, reportData, performanceData) {
  // Calculate overall statistics
  let totalMarks = 0;
  let subjectCount = 0;
  let overallPosition = (reportData && reportData.overallPosition) ? reportData.overallPosition : '-';
  
  performanceData.forEach(p => {
    totalMarks += (p.total || 0);
    subjectCount++;
  });
  
  const average = subjectCount > 0 ? (totalMarks / subjectCount).toFixed(2) : 0;
  
  // Determine grade based on average
  let grade = 'F FAIL';
  if (average >= 80) grade = 'A EXCELLENT';
  else if (average >= 70) grade = 'B VERY GOOD';
  else if (average >= 60) grade = 'C GOOD';
  else if (average >= 45) grade = 'D CREDIT';
  else if (average >= 35) grade = 'E WEAK';
  
  // Build subject rows dynamically from actual performance data
  let subjectRows = '';
  
  // Sort performance data by course name for consistent ordering
  const sortedPerformance = performanceData.sort((a, b) => {
    const courseA = (a.course || '').toUpperCase();
    const courseB = (b.course || '').toUpperCase();
    return courseA.localeCompare(courseB);
  });
  
  sortedPerformance.forEach(p => {
    const subjectName = p.course || 'N/A';
    const classScore = (p.classScore || 0).toFixed(1);
    const examScore100 = (p.examScore100 || 0).toFixed(1);
    const examScore50 = (p.examScore50 !== undefined ? p.examScore50 : (p.examScore60 || 0)).toFixed(1);
    const total = (p.total || 0).toFixed(1);
    const position = p.rank || '-'; // This is the RANK from performance table
    
    // Determine grade for subject
    let subjectGrade = 'F';
    let remarks = 'Fail';
    const totalNum = parseFloat(total);
    
    if (totalNum >= 80) {
      subjectGrade = 'A';
      remarks = 'Excellent';
    } else if (totalNum >= 70) {
      subjectGrade = 'B';
      remarks = 'Very Good';
    } else if (totalNum >= 60) {
      subjectGrade = 'C';
      remarks = 'Good';
    } else if (totalNum >= 45) {
      subjectGrade = 'D';
      remarks = 'Credit';
    } else if (totalNum >= 35) {
      subjectGrade = 'E';
      remarks = 'Weak';
    } else {
      subjectGrade = 'F';
      remarks = 'Fail';
    }
    
    subjectRows += `
      <tr>
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: left;">${subjectName}</td>
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: center;">${classScore}</td>
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: center;">${examScore100}</td>
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: center;">${examScore50}</td>
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: center;"><strong>${total}</strong></td>
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: center;">${position}</td>
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: center;"><strong>${subjectGrade}</strong></td>
        <td style="padding: 6px 4px; border: 1px solid #000; font-size: 10px;">${remarks}</td>
      </tr>`;
  });
  
  // Add a summary row at the bottom showing average
  if (sortedPerformance.length > 0) {
    // Determine overall remark based on average
    let overallRemark = 'Fail';
    const avgNum = parseFloat(average);
    if (avgNum >= 80) overallRemark = 'Excellent';
    else if (avgNum >= 70) overallRemark = 'Very Good';
    else if (avgNum >= 60) overallRemark = 'Good';
    else if (avgNum >= 45) overallRemark = 'Credit';
    else if (avgNum >= 35) overallRemark = 'Weak';
    
    subjectRows += `
      <tr style="background: #f0f0f0; font-weight: bold;">
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: left;">AVERAGE / TOTAL</td>
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: center;" colspan="3">${subjectCount} Subjects</td>
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: center;">${average}</td>
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: center;">${overallPosition}</td>
        <td style="padding: 6px 4px; border: 1px solid #000; text-align: center;">${grade.split(' ')[0]}</td>
        <td style="padding: 6px 4px; border: 1px solid #000; font-size: 10px;">${overallRemark}</td>
      </tr>`;
  }
  
  const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Terminal Report - ${studentInfo.fullName || studentInfo.firstName + ' ' + studentInfo.lastName}</title>
  <style>
    @page { 
      size: A4 portrait; 
      margin: 10mm 15mm;
    }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { 
      font-family: 'Times New Roman', Times, serif; 
      background: #fff; 
      width: 210mm;
      min-height: 297mm;
      margin: 0 auto;
    }
    .report-container { 
      width: 100%;
      min-height: 297mm;
      border: 3px solid #654321; 
      padding: 12px; 
      background: #f5f5f5;
      display: flex;
      flex-direction: column;
    }
    .report-content {
      flex: 1;
      display: flex;
      flex-direction: column;
    }
    .header { 
      text-align: center; 
      margin-bottom: 12px; 
      border-bottom: 2px solid #000; 
      padding-bottom: 10px;
      position: relative;
    }
    .school-logo {
      width: 60px;
      height: 60px;
      position: absolute;
      left: 10px;
      top: 0;
      object-fit: contain;
    }
    .header h1 { font-size: 16px; margin: 6px 0; text-transform: uppercase; font-weight: bold; }
    .header p { font-size: 11px; margin: 2px 0; }
    .student-info { margin: 10px 0; font-size: 12px; line-height: 1.6; }
    .student-info div { display: flex; justify-content: space-between; }
    .student-info span { flex: 1; }
    .grading-key { margin: 10px 0; padding: 6px; background: #fff; border: 1px solid #000; }
    .grading-key table { width: 100%; font-size: 10px; }
    .grading-key td { padding: 3px; border: 1px solid #000; text-align: center; font-weight: bold; }
    
    .table-wrapper {
      flex: 1;
      display: flex;
      flex-direction: column;
      min-height: 300px;
    }
    
    .performance-table { 
      width: 100%; 
      border-collapse: collapse; 
      background: #fff; 
      font-size: 11px;
    }
    .performance-table th { 
      background: #e0e0e0; 
      padding: 3px 2px; 
      border: 1px solid #000; 
      font-size: 9px; 
      text-align: center; 
      font-weight: bold; 
      line-height: 1; 
    }
    .performance-table td { 
      font-size: 10px; 
      padding: 3px 2px; 
      border: 1px solid #000; 
      vertical-align: middle;
      line-height: 1.1;
    }
    
    .footer-section { 
      margin-top: auto;
      padding-top: 12px;
      font-size: 11px; 
      line-height: 1.6; 
      page-break-inside: avoid;
    }
    .signature-line { border-bottom: 1px solid #000; display: inline-block; min-width: 180px; margin-left: 8px; }
    .motto { text-align: center; font-style: italic; margin-top: 8px; font-size: 10px; color: #333; font-weight: bold; }
  </style>
</head>
<body>
  <div class="report-container">
    <div class="report-content">
      <!-- Header -->
      <div class="header">
        ${schoolLogo ? `<img src="${schoolLogo}" class="school-logo" alt="School Logo" />` : ''}
        <h1>${schoolName}</h1>
        <p>${schoolAddress}</p>
        ${schoolEmail ? `<p style="font-size: 10px;">Email: ${schoolEmail}</p>` : ''}
        ${schoolPhone ? `<p style="font-size: 10px;">Phone: ${schoolPhone}</p>` : ''}
        <p style="font-weight: bold; margin-top: 6px;">PUPIL'S REPORT SHEET</p>
      </div>
      
      <!-- Student Info -->
      <div class="student-info">
        <div>
          <span><strong>NAME:</strong> ${studentInfo.fullName || (studentInfo.firstName + ' ' + studentInfo.lastName)}</span>
          <span><strong>No. of Roll:</strong> .................................</span>
        </div>
        <div>
          <span><strong>FORM / CLASS:</strong> ${reportData.studentClass}</span>
          <span><strong>TERM:</strong> ${reportData.term}</span>
          <span><strong>YEAR:</strong> ${reportData.academicYear}</span>
          <span><strong>VAC. DATE:</strong> .............................</span>
        </div>
        <div>
          <span><strong>NEXT TERM BEGINS:</strong> ....................................</span>
          <span><strong>OVERALL POS:</strong> ${overallPosition}</span>
        </div>
      </div>
      
      <!-- Grading Key -->
      <div class="grading-key">
        <table>
          <tr>
            <td><strong>A EXCELLENT (80% - 100%)</strong></td>
            <td><strong>B VERY GOOD (70% - 79%)</strong></td>
            <td><strong>C GOOD (60% - 69%)</strong></td>
          </tr>
          <tr>
            <td><strong>D CREDIT (45% - 59%)</strong></td>
            <td><strong>E WEAK (35% - 44%)</strong></td>
            <td><strong>F FAIL (00% - 34%)</strong></td>
          </tr>
        </table>
      </div>
      
      <!-- Performance Table -->
      <div class="table-wrapper">
        <table class="performance-table">
          <thead>
            <tr>
              <th rowspan="2">SUBJECTS</th>
              <th>CLASS<br>SCORE<br>50%</th>
              <th>EXAMS<br>SCORE<br>100%</th>
              <th>EXAMS<br>SCORE<br>50%</th>
              <th>TOTAL<br>SCORE<br>100%</th>
              <th rowspan="2">POS</th>
              <th rowspan="2">GRADES</th>
              <th rowspan="2">REMARKS<br><span style="font-weight: normal; font-size: 9px;">Specific Areas of<br>Strength & Weakness</span></th>
            </tr>
          </thead>
          <tbody>
            ${subjectRows}
          </tbody>
        </table>
      </div>
    </div>
    
    <!-- Footer Section -->
    <div class="footer-section">
      <div style="margin-bottom: 6px;">
        <strong>ATTENDANCE:</strong> ...............................
        <span style="margin-left: 40px;"><strong>PROMOTED TO:</strong> .................................</span>
      </div>
      <div style="margin-bottom: 6px;">
        <strong>CONDUCT:</strong> .................................................................................................................................
      </div>
      <div style="margin-bottom: 6px;">
        <strong>ATTITUDE:</strong> ................................................................................................................................
      </div>
      <div style="margin-bottom: 6px;">
        <strong>CLASS TEACHER'S REMARK:</strong> ..........................................................................................................
        <br>
        ....................................................................................................................................................................
      </div>
      <div style="margin-bottom: 8px;">
        <strong>HEAD TEACHER'S SIGNATURE:</strong> <span class="signature-line"></span>
      </div>
      
      <!-- Motto -->
      <div class="motto">
        <strong>NO CROSS NO CROWN</strong>
      </div>
    </div>
  </div>
</body>
</html>`;
  
  return html;
}

function normalizeParamName(name) {
  return String(name || '').trim().toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-9]/g, '');
}

function hashString(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8);
  return bytes.map(function(byte) {
    return (byte + 256).toString(16).slice(-2);
  }).join('');
}

function verifyPassword(candidate, storedValue) {
  if (!storedValue) return false;
  if (String(storedValue).startsWith('sha256:')) {
    return hashString(candidate) === String(storedValue).slice(7);
  }
  return String(candidate) === String(storedValue);
}

function getSystemParameter(paramName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) return null;
  
  const data = settingsSheet.getDataRange().getValues();
  const normalizedTarget = normalizeParamName(paramName);

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    for (let j = 0; j < row.length - 1; j++) {
      const cellVal = String(row[j] || '').trim();
      if (cellVal && normalizeParamName(cellVal) === normalizedTarget) {
        // Find the next non-empty cell in the row
        for (let k = j + 1; k < row.length; k++) {
          const val = String(row[k] || '').trim();
          if (val !== '') {
            return val;
          }
        }
      }
    }
  }
  return null;
}

/**
 * Get attendance status options from Settings sheet
 * Returns array of status objects with value, label, and badge class
 * Falls back to default statuses if not configured
 */
function getAttendanceStatuses() {
  try {
    const statusConfig = getSystemParameter('Attendance Statuses');
    
    if (statusConfig) {
      // Parse comma-separated status list from Settings
      const statuses = statusConfig.split(',').map(s => s.trim()).filter(s => s !== '');
      return statuses.map(status => {
        return {
          value: status,
          label: status,
          badgeClass: 'status-' + status.toLowerCase().replace(/\s+/g, '-')
        };
      });
    }
  } catch (e) {
    Logger.log('Error reading attendance statuses from settings: ' + e.message);
  }
  
  // Default statuses if not configured
  return [
    { value: 'Present', label: 'Present', badgeClass: 'status-present' },
    { value: 'Absent', label: 'Absent', badgeClass: 'status-absent' },
    { value: 'Late', label: 'Late', badgeClass: 'status-late' },
    { value: 'Excused', label: 'Excused', badgeClass: 'status-excused' }
  ];
}


/**
 * Convert a logo URL or Drive ID to a base64 Data URI for reliable PDF rendering
 */
function getImageAsBase64(logoUrl) {
  if (!logoUrl) return '';
  const str = String(logoUrl).trim();
  if (str.startsWith('data:image/')) return str; // Already base64

  try {
    let blob = null;

    // Check if string contains a Google Drive file ID
    let fileId = null;
    const matchId = str.match(/id=([a-zA-Z0-9_-]+)/);
    const matchD = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (matchId) {
      fileId = matchId[1];
    } else if (matchD) {
      fileId = matchD[1];
    } else if (/^[a-zA-Z0-9_-]{25,}$/.test(str)) {
      fileId = str;
    }

    if (fileId) {
      try {
        const file = DriveApp.getFileById(fileId);
        blob = file.getBlob();
      } catch (e) {
        Logger.log('DriveApp failed to get file by ID (' + fileId + '): ' + e.message);
      }
    }

    // Fallback: fetch via UrlFetchApp if DriveApp didn't get a blob
    if (!blob && (str.startsWith('http://') || str.startsWith('https://'))) {
      try {
        const response = UrlFetchApp.fetch(str, { muteHttpExceptions: true });
        if (response.getResponseCode() === 200) {
          blob = response.getBlob();
        }
      } catch (e) {
        Logger.log('UrlFetchApp failed for logo URL: ' + e.message);
      }
    }

    if (blob) {
      const mimeType = blob.getContentType() || 'image/png';
      const base64 = Utilities.base64Encode(blob.getBytes());
      return `data:${mimeType};base64,${base64}`;
    }
  } catch (err) {
    Logger.log('Error converting logo to Base64: ' + err.toString());
  }

  return logoUrl; // Fallback to original string if conversion fails
}

function createDailyBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupName = `SMS_Backup_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}`;
  
  const folder = DriveApp.getFolderById('YOUR_BACKUP_FOLDER_ID');
  ss.copy(backupName).moveTo(folder);
}

// Set up trigger: Edit > Current project's triggers
// Add: createDailyBackup, Time-driven, Day timer, 2am-3am


function lockSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Change "Students" to the name of the sheet you want to lock
  const sheet = ss.getSheetByName("Students"); 

  // Create a protection object for the sheet
  const protection = sheet.protect().setDescription('Locked for script entry only');

  // Get your own email (the person running this script)
  const me = Session.getEffectiveUser().getEmail();
  
  // Add yourself as an editor
  protection.addEditor(me);
  
  // Remove all other editors
  protection.removeEditors(protection.getEditors());
  
  // If your domain allows domain-wide editing, disable that too
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

/**
 * Menu hook for Google Sheets: Adds "🚀 SaaS Migration" menu to the toolbar.
 */
function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu("🚀 SaaS Migration")
      .addItem("📦 Export Data for SaaS (Modal & Download)", "showSaaSMigrationModal")
      .addItem("📁 Save JSON to Google Drive", "exportSaaSMigrationPayload")
      .addToUi();
  } catch (err) {
    Logger.log("onOpen menu registration notice: " + err.message);
  }
}

/**
 * Packages the school's data into a unified JSON structure,
 * saves a complete file in Google Drive, and logs the download URL.
 */
function exportSaaSMigrationPayload() {
  var payload = {
    exportedAt: new Date().toISOString(),
    schoolInfo: typeof getParameters === 'function' ? getParameters() : {},
    students: typeof getStudentsData === 'function' ? getStudentsData() : [],
    courses: typeof getCoursesData === 'function' ? getCoursesData() : [],
    classes: typeof getClassesData === 'function' ? getClassesData() : [],
    teachers: typeof getTeachersData === 'function' ? getTeachersData() : [],
    billingCategories: typeof getBillingCategoriesData === 'function' ? getBillingCategoriesData() : [],
    academicYears: typeof getAcademicYearsData === 'function' ? getAcademicYearsData() : []
  };

  var jsonString = JSON.stringify(payload, null, 2);

  var fileUrl = "";
  try {
    var fileName = "sms_saas_migration_data_" + new Date().toISOString().slice(0, 10) + ".json";
    var file = DriveApp.createFile(fileName, jsonString, MimeType.PLAIN_TEXT);
    fileUrl = file.getUrl();
  } catch (e) {
    Logger.log("Drive file creation error (may lack permissions): " + e.message);
  }

  Logger.log("===============================================================");
  Logger.log("✅ SAAS MIGRATION EXPORT READY!");
  Logger.log("📊 Summary of Exported Records:");
  Logger.log("   • Students:           " + (payload.students ? payload.students.length : 0));
  Logger.log("   • Classes:            " + (payload.classes ? payload.classes.length : 0));
  Logger.log("   • Courses/Subjects:   " + (payload.courses ? payload.courses.length : 0));
  Logger.log("   • Academic Years:     " + (payload.academicYears ? payload.academicYears.length : 0));
  Logger.log("   • Billing Categories: " + (payload.billingCategories ? payload.billingCategories.length : 0));
  Logger.log("   • Teachers:           " + (payload.teachers ? payload.teachers.length : 0));
  Logger.log("---------------------------------------------------------------");
  if (fileUrl) {
    Logger.log("📁 COMPLETE JSON FILE CREATED IN YOUR GOOGLE DRIVE:");
    Logger.log("👉 " + fileUrl);
    Logger.log("---------------------------------------------------------------");
  }
  Logger.log("📋 JSON SNIPPET (First 1500 chars):");
  Logger.log(jsonString.substring(0, 1500));
  Logger.log("===============================================================");

  return payload;
}

/**
 * Returns the raw JSON migration string for the HTML modal dialog.
 */
function getSaaSMigrationJSONString() {
  var payload = {
    exportedAt: new Date().toISOString(),
    schoolInfo: typeof getParameters === 'function' ? getParameters() : {},
    students: typeof getStudentsData === 'function' ? getStudentsData() : [],
    courses: typeof getCoursesData === 'function' ? getCoursesData() : [],
    classes: typeof getClassesData === 'function' ? getClassesData() : [],
    teachers: typeof getTeachersData === 'function' ? getTeachersData() : [],
    billingCategories: typeof getBillingCategoriesData === 'function' ? getBillingCategoriesData() : [],
    academicYears: typeof getAcademicYearsData === 'function' ? getAcademicYearsData() : []
  };
  return JSON.stringify(payload, null, 2);
}

/**
 * Displays a visual modal dialog inside the Google Spreadsheet with 1-click
 * Copy, Download, and Direct-Send options.
 */
function showSaaSMigrationModal() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html><head><base target="_top">' +
    '<style>' +
    'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 15px; margin: 0; background: #f8fafc; color: #1e293b; }' +
    '.card { background: white; border-radius: 12px; padding: 16px; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }' +
    'h2 { margin: 0 0 10px; font-size: 18px; color: #0f172a; }' +
    '.stat-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin-bottom: 12px; }' +
    '.stat { background: #f1f5f9; padding: 8px; border-radius: 8px; text-align: center; }' +
    '.stat-val { font-size: 18px; font-weight: bold; color: #2563eb; }' +
    '.stat-lbl { font-size: 10px; color: #64748b; text-transform: uppercase; }' +
    'textarea { width: 100%; height: 160px; font-family: monospace; font-size: 11px; padding: 8px; border: 1px solid #cbd5e1; border-radius: 8px; box-sizing: border-box; resize: vertical; }' +
    '.btn-row { display: flex; gap: 8px; margin-top: 12px; }' +
    'button { flex: 1; padding: 10px; border-radius: 8px; border: none; font-size: 12px; font-weight: bold; cursor: pointer; transition: background 0.2s; }' +
    '.btn-primary { background: #2563eb; color: white; }' +
    '.btn-primary:hover { background: #1d4ed8; }' +
    '.btn-secondary { background: #e2e8f0; color: #334155; }' +
    '.btn-secondary:hover { background: #cbd5e1; }' +
    '#msg { margin-top: 8px; font-size: 11px; font-weight: bold; text-align: center; min-height: 16px; }' +
    '</style></head><body>' +
    '<div class="card">' +
    '<h2>🚀 Export Data for Cloud SaaS (Neon DB)</h2>' +
    '<div id="loading" style="text-align:center; padding: 20px; font-size: 13px; color: #64748b;">Generating migration data from sheets...</div>' +
    '<div id="content" style="display:none;">' +
    '<div class="stat-grid" id="stats"></div>' +
    '<label style="font-size: 11px; font-weight: bold; color: #475569; display:block; margin-bottom: 4px;">Migration JSON Data:</label>' +
    '<textarea id="jsonArea" readonly></textarea>' +
    '<div class="btn-row">' +
    '<button class="btn-primary" onclick="copyJSON()">📋 Copy to Clipboard</button>' +
    '<button class="btn-secondary" onclick="downloadJSON()">💾 Download File</button>' +
    '</div>' +
    '<div id="msg"></div>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'var rawData = "";' +
    'google.script.run.withSuccessHandler(function(jsonStr) {' +
    '  rawData = jsonStr;' +
    '  document.getElementById("loading").style.display = "none";' +
    '  document.getElementById("content").style.display = "block";' +
    '  document.getElementById("jsonArea").value = jsonStr;' +
    '  try {' +
    '    var data = JSON.parse(jsonStr);' +
    '    var s = document.getElementById("stats");' +
    '    s.innerHTML = "<div class=\'stat\'><div class=\'stat-val\'>" + (data.students ? data.students.length : 0) + "</div><div class=\'stat-lbl\'>Students</div></div>" +' +
    '                  "<div class=\'stat\'><div class=\'stat-val\'>" + (data.classes ? data.classes.length : 0) + "</div><div class=\'stat-lbl\'>Classes</div></div>" +' +
    '                  "<div class=\'stat\'><div class=\'stat-val\'>" + (data.courses ? data.courses.length : 0) + "</div><div class=\'stat-lbl\'>Subjects</div></div>";' +
    '  } catch(e) {}' +
    '}).getSaaSMigrationJSONString();' +
    'function copyJSON() {' +
    '  var ta = document.getElementById("jsonArea");' +
    '  ta.select();' +
    '  document.execCommand("copy");' +
    '  var m = document.getElementById("msg");' +
    '  m.style.color = "#16a34a";' +
    '  m.innerText = "✅ Copied to clipboard! Go to your SaaS Dashboard -> Sheets Migration to paste it.";' +
    '}' +
    'function downloadJSON() {' +
    '  var blob = new Blob([rawData], { type: "application/json" });' +
    '  var a = document.createElement("a");' +
    '  a.href = URL.createObjectURL(blob);' +
    '  a.download = "sms_saas_migration_data.json";' +
    '  a.click();' +
    '  var m = document.getElementById("msg");' +
    '  m.style.color = "#16a34a";' +
    '  m.innerText = "✅ Downloaded sms_saas_migration_data.json!";' +
    '}' +
    '</script></body></html>'
  ).setWidth(520).setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, "SaaS Platform Migration");
}
