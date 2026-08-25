// --- INVOICES FUNCTIONS ---

// Fill invoice student details from the authoritative Students sheet.  The
// browser sends these details for a responsive UI, but this keeps direct/API
// calls correct too and preserves historical details if a student was removed.
function populateInvoiceStudentDetails(invData, existingInvoice) {
  if (!invData) return invData;
  const studentId = String(invData.studentId || '').trim();
  const student = studentId && typeof getStudentIdNamePairs === 'function'
    ? getStudentIdNamePairs().find(function(item) {
        return String(item.studentId || '').trim() === studentId;
      })
    : null;

  if (student) {
    invData.studentName = student.fullName || '';
    invData.studentClass = student.studentClass || '';
  } else if (existingInvoice) {
    invData.studentName = invData.studentName || existingInvoice.studentName || '';
    invData.studentClass = invData.studentClass || existingInvoice.studentClass || '';
  }
  return invData;
}

// 1. Fetch Invoices Data from Sheet (authoritative reader)
function getInvoicesDataFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Invoices");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; 

  // Row 3 to Last, Col B(2) onwards - read all columns to accommodate new structure
  // Updated structure: Invoice ID, Student ID, Student Name, Class, Academic Year, Term, Category, 
  //                    Items, Amount, Issue Date, Due Date, Status, Payment Status (13 columns)
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 13); 
  const data = dataRange.getValues();

  return data
    .filter(row => String(row[0]).trim() !== '') // Filter out empty rows
    .map((row) => {
      // Helper to safely parse dates
      const parseDate = (dateValue) => {
        if (!dateValue) return '';
        try {
          const d = new Date(dateValue);
          if (isNaN(d.getTime())) return ''; // Invalid date
          return d.toISOString().split('T')[0];
        } catch (e) {
          return '';
        }
      };

      return {
        invoiceId: String(row[0]).trim(),
        studentId: String(row[1]).trim(),
        studentName: String(row[2] || '').trim(),
        studentClass: String(row[3] || '').trim(),
        academicYear: String(row[4] || '').trim(),
        term: String(row[5] || '').trim(),
        category: String(row[6] || '').trim(),
        items: String(row[7] || '').trim(),
        amountDue: Number(row[8]) || 0,
        issueDate: parseDate(row[9]),
        dueDate: parseDate(row[10]),
        status: String(row[11] || '').trim(),
        paymentStatus: String(row[12] || '').trim()
      };
    });
}

// 2. Helper to generate Invoice ID (e.g. INV-1001)
function generateNextInvoiceId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "INV-1001";

  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues(); 
  let maxIdNum = 1000;
  values.forEach((row) => {
    const idStr = String(row[0]).trim();
    if (idStr.startsWith("INV-")) {
      const num = parseInt(idStr.substring(4), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return "INV-" + (maxIdNum + 1);
}

// 3. Helper: Find row by Invoice ID
function findInvoiceRowById(sheet, invoiceId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return -1;
  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues(); // Column B
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(invoiceId).trim()) {
      return i + 3; // Index 0 matches Row 3
    }
  }
  return -1;
}

// 4. Add New Invoice (Create)
function addInvoice(invData) {
  invData = invData || {};
  // Default to the active year if this endpoint is called outside the form.
  if (!String(invData.academicYear || '').trim() && typeof getActiveAcademicYearValue === 'function') {
    invData.academicYear = getActiveAcademicYearValue();
  }
  populateInvoiceStudentDetails(invData);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Invoices");
  if (!sheet) throw new Error("Invoices worksheet not found.");

  const nextId = generateNextInvoiceId(sheet);
  const pendingInvoiceSnapshot = buildAuditSnapshot({
    invoiceId: nextId,
    studentId: invData.studentId || '',
    studentName: invData.studentName || '',
    studentClass: invData.studentClass || '',
    academicYear: invData.academicYear || '',
    term: invData.term || '',
    category: invData.category || '',
    items: invData.items || '',
    amountDue: Number(invData.amountDue) || 0,
    issueDate: invData.issueDate || '',
    dueDate: invData.dueDate || '',
    status: invData.status || 'Pending',
    paymentStatus: invData.paymentStatus || 'Unpaid'
  });
  const createInvoiceYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Create',
    module: 'Invoices',
    recordId: nextId,
    academicYear: invData.academicYear,
    newValue: pendingInvoiceSnapshot,
    overrideConfirmed: invData && invData.__adminAcademicYearOverride === true
  });
  if (!createInvoiceYearGuard.allowed) return createInvoiceYearGuard;
  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  // Helper to safely parse date
  const parseDate = (dateValue) => {
    if (!dateValue) return new Date();
    try {
      const d = new Date(dateValue);
      return isNaN(d.getTime()) ? new Date() : d;
    } catch (e) {
      return new Date();
    }
  };

  // Write to Column B onwards (13 columns for new structure)
  sheet.getRange(targetRow, 2, 1, 13).setValues([
    [
      nextId,
      invData.studentId || '',
      invData.studentName || '',
      invData.studentClass || '',
      invData.academicYear || '',
      invData.term || '',
      invData.category || '',
      invData.items || '',
      invData.amountDue || 0,
      parseDate(invData.issueDate),
      parseDate(invData.dueDate),
      invData.status || 'Pending',
      invData.paymentStatus || 'Unpaid'
    ]
  ]);

  const rawInvoices = getInvoicesDataFromSheet();
  const createdInvoice = rawInvoices.find((invoice) => String(invoice.invoiceId).trim() === String(nextId).trim()) || buildAuditSnapshot({
    invoiceId: nextId,
    studentId: invData.studentId || '',
    studentName: invData.studentName || '',
    studentClass: invData.studentClass || '',
    academicYear: invData.academicYear || '',
    term: invData.term || '',
    category: invData.category || '',
    items: invData.items || '',
    amountDue: Number(invData.amountDue) || 0,
    issueDate: invData.issueDate || '',
    dueDate: invData.dueDate || '',
    status: invData.status || 'Pending',
    paymentStatus: invData.paymentStatus || 'Unpaid'
  });

  if (typeof invalidateFinancialCache === 'function') {
    invalidateFinancialCache();
  }

  safeLogAuditEvent(
    'Create',
    'Invoices',
    nextId,
    appendAcademicYearOverrideAuditDetails('Created invoice for ' + String(createdInvoice.studentName || createdInvoice.studentId || nextId).trim(), createInvoiceYearGuard),
    null,
    createdInvoice
  );

  return { success: true, invoiceId: nextId };
}

// 5. Update Invoice (Update)
function updateInvoice(invoiceId, invData) {
  invData = invData || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Invoices");
  if (!sheet) throw new Error("Invoices worksheet not found.");

  const rawInvoices = getInvoicesDataFromSheet();
  const existingInvoice = rawInvoices.find((invoice) => String(invoice.invoiceId).trim() === String(invoiceId).trim()) || null;
  populateInvoiceStudentDetails(invData, existingInvoice);
  const pendingUpdatedInvoice = buildAuditSnapshot(Object.assign({}, existingInvoice || {}, invData, { invoiceId: invoiceId }));
  const updateInvoiceYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Update',
    module: 'Invoices',
    recordId: invoiceId,
    academicYears: [existingInvoice && existingInvoice.academicYear, invData.academicYear],
    oldValue: existingInvoice,
    newValue: pendingUpdatedInvoice,
    overrideConfirmed: invData && invData.__adminAcademicYearOverride === true
  });
  if (!updateInvoiceYearGuard.allowed) return updateInvoiceYearGuard;
  const row = findInvoiceRowById(sheet, invoiceId);
  if (row === -1) throw new Error("Invoice record not found.");

  // Helper to safely parse date
  const parseDate = (dateValue) => {
    if (!dateValue) return new Date();
    try {
      const d = new Date(dateValue);
      return isNaN(d.getTime()) ? new Date() : d;
    } catch (e) {
      return new Date();
    }
  };

  // Update from Col C(3) onwards -> 12 columns (excluding Invoice ID)
  sheet.getRange(row, 3, 1, 12).setValues([
    [
      invData.studentId || '',
      invData.studentName || '',
      invData.studentClass || '',
      invData.academicYear || '',
      invData.term || '',
      invData.category || '',
      invData.items || '',
      invData.amountDue || 0,
      parseDate(invData.issueDate),
      parseDate(invData.dueDate),
      invData.status || 'Pending',
      invData.paymentStatus || 'Unpaid'
    ]
  ]);

  const updatedInvoice = getInvoicesDataFromSheet().find((invoice) => String(invoice.invoiceId).trim() === String(invoiceId).trim()) || buildAuditSnapshot(Object.assign({}, existingInvoice || {}, invData, { invoiceId: invoiceId }));

  if (typeof invalidateFinancialCache === 'function') {
    invalidateFinancialCache();
  }

  safeLogAuditEvent(
    'Update',
    'Invoices',
    invoiceId,
    appendAcademicYearOverrideAuditDetails('Updated invoice for ' + String((updatedInvoice && (updatedInvoice.studentName || updatedInvoice.studentId)) || invoiceId).trim(), updateInvoiceYearGuard),
    existingInvoice,
    updatedInvoice
  );

  return { success: true };
}

// 6. Delete Invoice (Delete)
function deleteInvoice(invoiceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Invoices");
  if (!sheet) throw new Error("Invoices worksheet not found.");

  const rawInvoices = getInvoicesDataFromSheet();
  const existingInvoice = rawInvoices.find((invoice) => String(invoice.invoiceId).trim() === String(invoiceId).trim()) || null;
  const deleteInvoiceYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Delete',
    module: 'Invoices',
    recordId: invoiceId,
    academicYear: existingInvoice && existingInvoice.academicYear,
    oldValue: existingInvoice,
    overrideConfirmed: arguments[1] && arguments[1].adminAcademicYearOverride === true
  });
  if (!deleteInvoiceYearGuard.allowed) return deleteInvoiceYearGuard;
  const row = findInvoiceRowById(sheet, invoiceId);
  if (row === -1) throw new Error("Invoice record not found.");

  sheet.deleteRow(row);

  if (typeof invalidateFinancialCache === 'function') {
    invalidateFinancialCache();
  }

  safeLogAuditEvent(
    'Delete',
    'Invoices',
    invoiceId,
    appendAcademicYearOverrideAuditDetails('Deleted invoice for ' + String(((existingInvoice && (existingInvoice.studentName || existingInvoice.studentId)) || invoiceId)).trim(), deleteInvoiceYearGuard),
    existingInvoice,
    null
  );

  return { success: true };
}
