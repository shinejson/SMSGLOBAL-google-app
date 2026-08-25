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

/**
 * Dynamic Header Mapper for Invoices sheet (Row 2, Column B onwards)
 */
function getInvoicesHeaderInfo(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 14);
  const rawHeaders = sheet.getRange(2, 2, 1, lastCol - 1).getValues()[0];
  const headers = rawHeaders.map(h => String(h || '').trim().toLowerCase());

  function findIndex(aliasGroups) {
    // 1) exact match
    for (let i = 0; i < headers.length; i++) {
      if (!headers[i]) continue;
      for (let a of aliasGroups) {
        if (headers[i] === a) return i;
      }
    }
    // 2) contains match
    for (let i = 0; i < headers.length; i++) {
      if (!headers[i]) continue;
      for (let a of aliasGroups) {
        if (headers[i].indexOf(a) !== -1) return i;
      }
    }
    return -1;
  }

  return {
    headers: headers,
    rawHeaders: rawHeaders,
    numCols: headers.length,
    idxInvoiceId: findIndex(['invoice id', 'invid', 'invoice_id', 'invoice', 'id']),
    idxStudentId: findIndex(['student id', 'studentid', 'student_id', 'sid']),
    idxStudentName: findIndex(['student name', 'name']),
    idxStudentClass: findIndex(['student class', 'class name', 'class']),
    idxAcademicYear: findIndex(['academic year', 'acad. year', 'acad year', 'year']),
    idxTerm: findIndex(['term']),
    idxCategory: findIndex(['category', 'billing category']),
    idxItems: findIndex(['items', 'debit items', 'item', 'description']),
    idxAmountDue: findIndex(['amount due', 'amount', 'total amount', 'total']),
    idxIssueDate: findIndex(['issue date', 'issued date', 'issue', 'date']),
    idxDueDate: findIndex(['due date', 'due']),
    idxStatus: findIndex(['status', 'invoice status']),
    idxPaymentStatus: findIndex(['payment status', 'pay status'])
  };
}

/**
 * Build a row array matched to the sheet's actual column positions
 */
function buildInvoiceRowArray(headerInfo, invoiceId, invData) {
  const row = new Array(headerInfo.numCols).fill('');

  const parseDate = (dateValue) => {
    if (!dateValue) return new Date();
    try {
      const d = new Date(dateValue);
      return isNaN(d.getTime()) ? new Date() : d;
    } catch (e) {
      return new Date();
    }
  };

  if (headerInfo.idxInvoiceId !== -1) row[headerInfo.idxInvoiceId] = invoiceId;
  if (headerInfo.idxStudentId !== -1) row[headerInfo.idxStudentId] = invData.studentId || '';
  if (headerInfo.idxStudentName !== -1) row[headerInfo.idxStudentName] = invData.studentName || '';
  if (headerInfo.idxStudentClass !== -1) row[headerInfo.idxStudentClass] = invData.studentClass || '';
  if (headerInfo.idxAcademicYear !== -1) row[headerInfo.idxAcademicYear] = invData.academicYear || '';
  if (headerInfo.idxTerm !== -1) row[headerInfo.idxTerm] = invData.term || '';
  if (headerInfo.idxCategory !== -1) row[headerInfo.idxCategory] = invData.category || 'Tuition';
  if (headerInfo.idxItems !== -1) row[headerInfo.idxItems] = invData.items || '';
  if (headerInfo.idxAmountDue !== -1) row[headerInfo.idxAmountDue] = Number(invData.amountDue) || 0;
  if (headerInfo.idxIssueDate !== -1) row[headerInfo.idxIssueDate] = parseDate(invData.issueDate);
  if (headerInfo.idxDueDate !== -1) row[headerInfo.idxDueDate] = parseDate(invData.dueDate);
  if (headerInfo.idxStatus !== -1) row[headerInfo.idxStatus] = invData.status || 'Pending';
  if (headerInfo.idxPaymentStatus !== -1) row[headerInfo.idxPaymentStatus] = invData.paymentStatus || 'Unpaid';

  return row;
}

// 1. Fetch Invoices Data from Sheet (authoritative reader)
function getInvoicesDataFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Invoices");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; 

  const headerInfo = getInvoicesHeaderInfo(sheet);
  const data = sheet.getRange(3, 2, lastRow - 2, headerInfo.numCols).getValues();

  const parseDateStr = (dateValue) => {
    if (!dateValue) return '';
    try {
      const d = new Date(dateValue);
      if (isNaN(d.getTime())) return String(dateValue).trim();
      return d.toISOString().split('T')[0];
    } catch (e) {
      return String(dateValue || '').trim();
    }
  };

  return data
    .filter(row => {
      const id = headerInfo.idxInvoiceId !== -1 ? String(row[headerInfo.idxInvoiceId] || '').trim() : String(row[0] || '').trim();
      return id !== '';
    })
    .map((row) => {
      const val = (idx, fallback = '') => (idx !== -1 && row[idx] !== undefined && row[idx] !== '') ? row[idx] : fallback;

      return {
        invoiceId: String(val(headerInfo.idxInvoiceId)).trim(),
        studentId: String(val(headerInfo.idxStudentId)).trim(),
        studentName: String(val(headerInfo.idxStudentName)).trim(),
        studentClass: String(val(headerInfo.idxStudentClass)).trim(),
        academicYear: String(val(headerInfo.idxAcademicYear)).trim(),
        term: String(val(headerInfo.idxTerm)).trim(),
        category: String(val(headerInfo.idxCategory)).trim(),
        items: String(val(headerInfo.idxItems)).trim(),
        amountDue: Number(val(headerInfo.idxAmountDue, 0)) || 0,
        issueDate: parseDateStr(val(headerInfo.idxIssueDate)),
        dueDate: parseDateStr(val(headerInfo.idxDueDate)),
        status: String(val(headerInfo.idxStatus, 'Pending')).trim(),
        paymentStatus: String(val(headerInfo.idxPaymentStatus, 'Unpaid')).trim()
      };
    });
}

// 2. Helper to generate Invoice ID (e.g. INV-1001)
function generateNextInvoiceId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "INV-1001";

  const headerInfo = getInvoicesHeaderInfo(sheet);
  const idCol = headerInfo.idxInvoiceId !== -1 ? headerInfo.idxInvoiceId + 2 : 2;

  const values = sheet.getRange(3, idCol, lastRow - 2, 1).getValues(); 
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
  const headerInfo = getInvoicesHeaderInfo(sheet);
  const idCol = headerInfo.idxInvoiceId !== -1 ? headerInfo.idxInvoiceId + 2 : 2;

  const values = sheet.getRange(3, idCol, lastRow - 2, 1).getValues();
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

  const headerInfo = getInvoicesHeaderInfo(sheet);
  const rowArray = buildInvoiceRowArray(headerInfo, nextId, invData);

  sheet.getRange(targetRow, 2, 1, headerInfo.numCols).setValues([rowArray]);

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

  const headerInfo = getInvoicesHeaderInfo(sheet);
  const rowArray = buildInvoiceRowArray(headerInfo, invoiceId, invData);

  sheet.getRange(row, 2, 1, headerInfo.numCols).setValues([rowArray]);

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
