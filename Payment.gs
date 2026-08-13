// --- PAYMENTS FUNCTIONS ---

function formatSheetDate(value) {
  if (!value) return '';
  try {
    const d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return String(value).trim();
    return d.toISOString().split('T')[0];
  } catch (e) {
    return String(value).trim();
  }
}

// 1. Fetch Payments Data (used by DataAccessLayer cache wrapper)
function getPaymentsDataFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Payments");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; 

  // Row 3 to Last, Col B(2) to Col N(14) => 13 columns (includes Balance + BalanceStatus)
  const dataRange = sheet.getRange(3, 2, lastRow - 2, 13); 
  const data = dataRange.getValues();

  return data
    .filter(row => String(row[0]).trim() !== "") // Skip empty spreadsheet rows
    .map((row) => ({
    transactionId: String(row[0]).trim(),
    invoiceId: String(row[1]).trim(),
    studentId: String(row[2]).trim(),
    paymentDate: formatSheetDate(row[3]),
    amountPaid: Number(row[4]) || 0,
    paymentMethod: String(row[5]).trim(),
    referenceNo: String(row[6]).trim(),
    studentName: String(row[7]).trim(),
    academicYear: String(row[8]).trim(),
    term: String(row[9]).trim(),
    studentClass: String(row[10]).trim(),
    balance: Number(row[11]) || 0,
    balanceStatus: String(row[12]).trim() || ''
  }));
}

// 2. Helper to generate Transaction ID (e.g. PAY-1001)
function generateNextPaymentId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "PAY-1001";

  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues(); 
  let maxIdNum = 1000;
  values.forEach((row) => {
    const idStr = String(row[0]).trim();
    if (idStr.startsWith("PAY-")) {
      const num = parseInt(idStr.substring(4), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return "PAY-" + (maxIdNum + 1);
}

// 3. Helper: Find row by Transaction ID
function findPaymentRowById(sheet, transactionId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return -1;
  const values = sheet.getRange(3, 2, lastRow - 2, 1).getValues(); // Column B
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(transactionId).trim()) {
      return i + 3;
    }
  }
  return -1;
}

// 4. Helper: Get minimal Invoice details (for the smart dropdown)
function getInvoiceDetails() {
  return getInvoicesData().map(inv => ({
    invoiceId: inv.invoiceId,
    studentId: inv.studentId,
    academicYear: inv.academicYear,
    term: inv.term
  }));
}

// 5. Helper: Get Student details by ID (for the smart modal)
function getStudentDetailsById(studentId) {
  const students = getStudentsData();
  const student = students.find(s => String(s.studentId).trim() === String(studentId).trim());
  if (!student) return null;

  return {
    studentId: student.studentId,
    studentName: `${String(student.firstName || '').trim()} ${String(student.lastName || '').trim()}`.trim(),
    studentClass: String(student.class || '').trim()
  };
}

// 6. Add New Payment (Create)
function addPayment(payData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Payments");
  if (!sheet) throw new Error("Payments worksheet not found.");

  const nextId = generateNextPaymentId(sheet);
  const pendingPaymentSnapshot = buildAuditSnapshot({
    transactionId: nextId,
    invoiceId: payData.invoiceId,
    studentId: payData.studentId,
    paymentDate: payData.paymentDate,
    amountPaid: Number(payData.amountPaid) || 0,
    paymentMethod: payData.paymentMethod,
    referenceNo: payData.referenceNo,
    studentName: payData.studentName,
    academicYear: payData.academicYear,
    term: payData.term,
    studentClass: payData.studentClass
  });
  const createPaymentYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Create',
    module: 'Payments',
    recordId: nextId,
    academicYear: payData.academicYear,
    newValue: pendingPaymentSnapshot,
    overrideConfirmed: payData && payData.__adminAcademicYearOverride === true
  });
  if (!createPaymentYearGuard.allowed) return createPaymentYearGuard;
  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  // Write to Column B to Column N (includes Balance + BalanceStatus placeholders)
  sheet.getRange(targetRow, 2, 1, 13).setValues([
    [
      nextId,
      payData.invoiceId,
      payData.studentId,
      new Date(payData.paymentDate),
      payData.amountPaid,
      payData.paymentMethod,
      payData.referenceNo,
      payData.studentName,
      payData.academicYear,
      payData.term,
      payData.studentClass,
      0, // balance placeholder; will be recalculated
      '' // balanceStatus placeholder
    ]
  ]);

  // Recalculate balances after adding
  try { recalcPaymentsBalances(); } catch (e) { /* fail silently to avoid breaking UI */ }
  try { invalidateFinancialCache(); } catch (e) { /* fail silently */ }

  const createdPayment = getPaymentsDataFromSheet().find((payment) => String(payment.transactionId).trim() === String(nextId).trim()) || buildAuditSnapshot({
    transactionId: nextId,
    invoiceId: payData.invoiceId,
    studentId: payData.studentId,
    paymentDate: payData.paymentDate,
    amountPaid: Number(payData.amountPaid) || 0,
    paymentMethod: payData.paymentMethod,
    referenceNo: payData.referenceNo,
    studentName: payData.studentName,
    academicYear: payData.academicYear,
    term: payData.term,
    studentClass: payData.studentClass
  });

  safeLogAuditEvent(
    'Create',
    'Payments',
    nextId,
    appendAcademicYearOverrideAuditDetails('Created payment for ' + String(createdPayment.studentName || createdPayment.studentId || nextId).trim(), createPaymentYearGuard),
    null,
    createdPayment
  );

  return { success: true, transactionId: nextId };
}

// 7. Update Payment (Update)
function updatePayment(transactionId, payData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Payments");
  if (!sheet) throw new Error("Payments worksheet not found.");

  const existingPayment = getPaymentsDataFromSheet().find((payment) => String(payment.transactionId).trim() === String(transactionId).trim()) || null;
  const pendingUpdatedPayment = buildAuditSnapshot(Object.assign({}, existingPayment || {}, payData, { transactionId: transactionId }));
  const updatePaymentYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Update',
    module: 'Payments',
    recordId: transactionId,
    academicYears: [existingPayment && existingPayment.academicYear, payData.academicYear],
    oldValue: existingPayment,
    newValue: pendingUpdatedPayment,
    overrideConfirmed: payData && payData.__adminAcademicYearOverride === true
  });
  if (!updatePaymentYearGuard.allowed) return updatePaymentYearGuard;
  const row = findPaymentRowById(sheet, transactionId);
  if (row === -1) throw new Error("Payment record not found.");

  // Update from Col C(3) to Col N(14) -> 12 columns (includes Balance + BalanceStatus placeholders)
  sheet.getRange(row, 3, 1, 12).setValues([
    [
      payData.invoiceId,
      payData.studentId,
      new Date(payData.paymentDate),
      payData.amountPaid,
      payData.paymentMethod,
      payData.referenceNo,
      payData.studentName,
      payData.academicYear,
      payData.term,
      payData.studentClass,
      0, // balance placeholder
      '' // balanceStatus placeholder
    ]
  ]);

  // Recalculate balances after updating
  try { recalcPaymentsBalances(); } catch (e) { /* fail silently to avoid breaking UI */ }
  try { invalidateFinancialCache(); } catch (e) { /* fail silently */ }

  const updatedPayment = getPaymentsDataFromSheet().find((payment) => String(payment.transactionId).trim() === String(transactionId).trim()) || buildAuditSnapshot(Object.assign({}, existingPayment || {}, payData, { transactionId: transactionId }));

  safeLogAuditEvent(
    'Update',
    'Payments',
    transactionId,
    appendAcademicYearOverrideAuditDetails('Updated payment for ' + String((updatedPayment && (updatedPayment.studentName || updatedPayment.studentId)) || transactionId).trim(), updatePaymentYearGuard),
    existingPayment,
    updatedPayment
  );

  return { success: true };
}

// 8. Delete Payment (Delete)
function deletePayment(transactionId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Payments");
  if (!sheet) throw new Error("Payments worksheet not found.");

  const existingPayment = getPaymentsDataFromSheet().find((payment) => String(payment.transactionId).trim() === String(transactionId).trim()) || null;
  const deletePaymentYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Delete',
    module: 'Payments',
    recordId: transactionId,
    academicYear: existingPayment && existingPayment.academicYear,
    oldValue: existingPayment,
    overrideConfirmed: arguments[1] && arguments[1].adminAcademicYearOverride === true
  });
  if (!deletePaymentYearGuard.allowed) return deletePaymentYearGuard;
  const row = findPaymentRowById(sheet, transactionId);
  if (row === -1) throw new Error("Payment record not found.");

  sheet.deleteRow(row);
  // Recalculate balances after deletion
  try { recalcPaymentsBalances(); } catch (e) { /* fail silently */ }
  try { invalidateFinancialCache(); } catch (e) { /* fail silently */ }

  safeLogAuditEvent(
    'Delete',
    'Payments',
    transactionId,
    appendAcademicYearOverrideAuditDetails('Deleted payment for ' + String(((existingPayment && (existingPayment.studentName || existingPayment.studentId)) || transactionId)).trim(), deletePaymentYearGuard),
    existingPayment,
    null
  );

  return { success: true };
}

// 9. Recalculate running balances per (billingCategory/invoice/student) and persist to sheet
function recalcPaymentsBalances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Payments");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  // Read full payments data including invoice/billing references
  // Columns B..N (2..14) => 13 columns
  const range = sheet.getRange(3, 2, lastRow - 2, 13);
  const rows = range.getValues();

  // Build a map grouping by composite key: invoiceId|studentId
  const groups = {};

  // We need billing category totals to compute balance per invoice/billing.
  // Attempt to fetch billing categories totals by invoiceId via invoices -> billingCategory
  const invoices = getInvoicesData();
  const billingCats = getBillingCategoriesData();

  const invoiceToTotal = {};
  invoices.forEach(inv => {
    const bc = billingCats.find(b => String(b.id).trim() === String(inv.billingCategoryId).trim());
    invoiceToTotal[inv.invoiceId] = bc ? Number(bc.totalAmount) || 0 : Number(inv.amount) || 0;
  });

  // Normalize and group rows
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const transactionId = String(r[0]).trim();
    if (!transactionId) continue;
    const invoiceId = String(r[1]).trim();
    const studentId = String(r[2]).trim();
    const paymentDate = r[3] ? new Date(r[3]) : new Date(2000,0,1);
    const amountPaid = Number(r[4]) || 0;

    const key = invoiceId + '|' + studentId;
    if (!groups[key]) groups[key] = [];
    groups[key].push({ idx: i, transactionId, invoiceId, studentId, paymentDate, amountPaid });
  }

  // Prepare output array to write balances and statuses back into M and N columns
  const out = rows.map(r => ([r[11] || 0, r[12] || '']));

  Object.keys(groups).forEach(key => {
    const list = groups[key];
    // sort by paymentDate asc
    list.sort((a,b) => a.paymentDate - b.paymentDate);

    const invoiceId = list[0].invoiceId;
    const totalAmount = Number(invoiceToTotal[invoiceId]) || 0;

    let cumulative = 0;
    list.forEach(item => {
      cumulative += Number(item.amountPaid) || 0;
      const balance = Math.max(0, totalAmount - cumulative);
      let status = 'No Payment';
      if (cumulative <= 0) status = 'No Payment';
      else if (cumulative >= totalAmount) status = 'Paid';
      else status = 'Part Payment';

      // write into out at index
      out[item.idx] = [balance, status];
    });
  });

  // Write back balances/statuses into columns M(13) and N(14)
  const writeRange = sheet.getRange(3, 13, out.length, 2);
  writeRange.setValues(out);
}