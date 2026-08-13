// --- BILLINGS FUNCTIONS ---

// 1. Fetch Billings Data
function getBillingsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Billings");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // Headers in Row 1

  // Row 2 to Last, Col A(1) to Col E(5) => 5 columns
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 5); 
  const data = dataRange.getValues();

  return data
    .filter(row => String(row[0]).trim() !== "") // Skip empty spreadsheet rows
    .map((row) => ({
    billingId: String(row[0]).trim(),
    item: String(row[1]).trim(),
    amount: Number(row[2]) || 0,
    status: String(row[3]).trim() || "Active",
    description: String(row[4]).trim(),
  }));
}

// 2. Helper to generate Billing ID (e.g. BIL-1001)
function generateNextBillingId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "BIL-1001";

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); 
  let maxIdNum = 1000;
  values.forEach((row) => {
    const idStr = String(row[0]).trim();
    if (idStr.startsWith("BIL-")) {
      const num = parseInt(idStr.substring(4), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return "BIL-" + (maxIdNum + 1);
}

// 3. Helper: Find row by Billing ID (Column A)
function findBillingRowById(sheet, billingId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Column A
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(billingId).trim()) {
      return i + 2; // Index 0 matches Row 2
    }
  }
  return -1;
}

function buildBillingAuditSnapshot(billingData) {
  if (!billingData) return null;

  return buildAuditSnapshot({
    billingId: billingData.billingId,
    item: billingData.item,
    amount: Number(billingData.amount) || 0,
    status: billingData.status,
    description: billingData.description
  });
}

// 4. Add New Billing Item (Create)
function addBilling(billingData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Billings");
  if (!sheet) throw new Error("Billings worksheet not found.");

  const nextId = generateNextBillingId(sheet);
  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  // Write to Column A to Column E
  sheet.getRange(targetRow, 1, 1, 5).setValues([
    [
      nextId,
      billingData.item,
      billingData.amount,
      billingData.status || "Active",
      billingData.description
    ]
  ]);

  const createdBilling = buildBillingAuditSnapshot({
    billingId: nextId,
    item: billingData.item,
    amount: billingData.amount,
    status: billingData.status || 'Active',
    description: billingData.description
  });

  safeLogAuditEvent(
    'Create',
    'Billing',
    nextId,
    'Created billing item ' + String(billingData.item || nextId).trim(),
    null,
    createdBilling
  );

  return { success: true, billingId: nextId };
}

// 5. Update Billing Item (Update)
function updateBilling(billingId, billingData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Billings");
  if (!sheet) throw new Error("Billings worksheet not found.");

  const existingBilling = getBillingsData().find((billing) => String(billing.billingId).trim() === String(billingId).trim()) || null;
  const row = findBillingRowById(sheet, billingId);
  if (row === -1) throw new Error("Billing record not found.");

  // Update from Col B(2) to Col E(5) -> 4 columns
  sheet.getRange(row, 2, 1, 4).setValues([
    [
      billingData.item,
      billingData.amount,
      billingData.status,
      billingData.description
    ]
  ]);

  const updatedBilling = buildBillingAuditSnapshot(Object.assign({}, existingBilling || {}, billingData, { billingId: billingId }));

  safeLogAuditEvent(
    'Update',
    'Billing',
    billingId,
    'Updated billing item ' + String((updatedBilling && updatedBilling.item) || billingId).trim(),
    buildBillingAuditSnapshot(existingBilling),
    updatedBilling
  );

  return { success: true };
}

// 6. Delete Billing Item (Delete)
function deleteBilling(billingId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Billings");
  if (!sheet) throw new Error("Billings worksheet not found.");

  const existingBilling = getBillingsData().find((billing) => String(billing.billingId).trim() === String(billingId).trim()) || null;
  const row = findBillingRowById(sheet, billingId);
  if (row === -1) throw new Error("Billing record not found.");

  sheet.deleteRow(row);

  safeLogAuditEvent(
    'Delete',
    'Billing',
    billingId,
    'Deleted billing item ' + String(((existingBilling && existingBilling.item) || billingId)).trim(),
    buildBillingAuditSnapshot(existingBilling),
    null
  );

  return { success: true };
}
