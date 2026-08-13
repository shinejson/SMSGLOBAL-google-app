// --- BILLING CATEGORIES FUNCTIONS ---

// 1. Fetch Billing Categories Data
function getBillingCategoriesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Billing Categories");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; 

  // Row 2 to Last, Col A(1) to Col J(10) => 10 columns
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 10); 
  const data = dataRange.getValues();

  return data
    .filter(row => {
      const id = String(row[0]).trim();
      // Only include rows with non-empty ID that looks like a valid ID (starts with number or BC-)
      return id !== "" && (id.match(/^\d+$/) || id.startsWith("BC-"));
    })
    .map((row) => ({
    id: String(row[0]).trim(),
    academicYear: String(row[1]).trim(),
    terms: String(row[2]).trim(),
    category: String(row[3]).trim(),
    items: String(row[4]).trim(), // Comma-separated string of items
    totalAmount: Number(row[5]) || 0,
    nursery: Number(row[6]) || 0,
    date: row[7] ? new Date(row[7]).toISOString().split('T')[0] : '',
    descriptions: String(row[8]).trim(),
  }));
}

// 2. Helper to generate Billing Category ID (e.g. BC-1001)
function generateNextBillingCategoryId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "BC-1001";

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); 
  let maxIdNum = 1000;
  values.forEach((row) => {
    const idStr = String(row[0]).trim();
    if (idStr.startsWith("BC-")) {
      const num = parseInt(idStr.substring(3), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return "BC-" + (maxIdNum + 1);
}

// 3. Helper: Find row by Billing Category ID (Column A)
function findBillingCategoryRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(id).trim()) {
      return i + 2; 
    }
  }
  return -1;
}

function buildBillingCategoryAuditSnapshot(categoryData) {
  if (!categoryData) return null;

  return buildAuditSnapshot({
    id: categoryData.id,
    academicYear: categoryData.academicYear,
    terms: categoryData.terms,
    category: categoryData.category,
    items: categoryData.items,
    totalAmount: Number(categoryData.totalAmount) || 0,
    nursery: Number(categoryData.nursery) || 0,
    date: categoryData.date || '',
    descriptions: categoryData.descriptions || ''
  });
}

function buildGeneratedInvoiceAuditSnapshot(fields) {
  return buildAuditSnapshot({
    invoiceId: fields.invoiceId,
    studentId: fields.studentId,
    studentName: fields.studentName,
    studentClass: fields.studentClass,
    academicYear: fields.academicYear,
    term: fields.term,
    category: fields.category,
    items: fields.items,
    amountDue: Number(fields.amountDue) || 0,
    issueDate: fields.issueDate,
    dueDate: fields.dueDate,
    status: fields.status,
    paymentStatus: fields.paymentStatus,
    source: fields.source
  });
}

// 4. Add New Billing Category (Create)
function addBillingCategory(catData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Billing Categories");
  if (!sheet) throw new Error("Billing Categories worksheet not found.");

  const nextId = generateNextBillingCategoryId(sheet);
  const pendingCategorySnapshot = buildBillingCategoryAuditSnapshot({
    id: nextId,
    academicYear: catData.academicYear,
    terms: catData.terms,
    category: catData.category,
    items: catData.items,
    totalAmount: catData.totalAmount,
    nursery: catData.nursery || 0,
    date: catData.date,
    descriptions: catData.descriptions
  });
  const createBillingCategoryYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Create',
    module: 'Billing Categories',
    recordId: nextId,
    academicYear: catData.academicYear,
    newValue: pendingCategorySnapshot,
    overrideConfirmed: catData && catData.__adminAcademicYearOverride === true
  });
  if (!createBillingCategoryYearGuard.allowed) return createBillingCategoryYearGuard;
  const lastRow = sheet.getLastRow();
  const targetRow = lastRow + 1;

  // Write to Column A to Column I
  sheet.getRange(targetRow, 1, 1, 9).setValues([
    [
      nextId,
      catData.academicYear,
      catData.terms,
      catData.category,
      catData.items, // Comma separated string
      catData.totalAmount,
      catData.nursery || 0,
      new Date(catData.date),
      catData.descriptions
    ]
  ]);

  const createdCategory = pendingCategorySnapshot;

  safeLogAuditEvent(
    'Create',
    'Billing Categories',
    nextId,
    appendAcademicYearOverrideAuditDetails('Created billing category ' + String(catData.category || nextId).trim(), createBillingCategoryYearGuard),
    null,
    createdCategory
  );

  return { success: true, id: nextId };
}

// 5. Update Billing Category (Update)
function updateBillingCategory(id, catData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Billing Categories");
  if (!sheet) throw new Error("Billing Categories worksheet not found.");

  const existingCategory = getBillingCategoriesData().find((category) => String(category.id).trim() === String(id).trim()) || null;
  const pendingUpdatedCategory = buildBillingCategoryAuditSnapshot(Object.assign({}, existingCategory || {}, catData, { id: id }));
  const updateBillingCategoryYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Update',
    module: 'Billing Categories',
    recordId: id,
    academicYears: [existingCategory && existingCategory.academicYear, catData.academicYear],
    oldValue: buildBillingCategoryAuditSnapshot(existingCategory),
    newValue: pendingUpdatedCategory,
    overrideConfirmed: catData && catData.__adminAcademicYearOverride === true
  });
  if (!updateBillingCategoryYearGuard.allowed) return updateBillingCategoryYearGuard;
  const row = findBillingCategoryRowById(sheet, id);
  if (row === -1) throw new Error("Billing Category record not found.");

  // Update from Col B(2) to Col I(9) -> 8 columns
  sheet.getRange(row, 2, 1, 8).setValues([
    [
      catData.academicYear,
      catData.terms,
      catData.category,
      catData.items,
      catData.totalAmount,
      catData.nursery || 0,
      new Date(catData.date),
      catData.descriptions
    ]
  ]);

  const updatedCategory = pendingUpdatedCategory;

  safeLogAuditEvent(
    'Update',
    'Billing Categories',
    id,
    appendAcademicYearOverrideAuditDetails('Updated billing category ' + String((updatedCategory && updatedCategory.category) || id).trim(), updateBillingCategoryYearGuard),
    buildBillingCategoryAuditSnapshot(existingCategory),
    updatedCategory
  );

  return { success: true };
}

// 6. Delete Billing Category (Delete)
function deleteBillingCategory(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Billing Categories");
  if (!sheet) throw new Error("Billing Categories worksheet not found.");

  const existingCategory = getBillingCategoriesData().find((category) => String(category.id).trim() === String(id).trim()) || null;
  const deleteBillingCategoryYearGuard = enforceAcademicYearCrudSecurity({
    action: 'Delete',
    module: 'Billing Categories',
    recordId: id,
    academicYear: existingCategory && existingCategory.academicYear,
    oldValue: buildBillingCategoryAuditSnapshot(existingCategory),
    overrideConfirmed: arguments[1] && arguments[1].adminAcademicYearOverride === true
  });
  if (!deleteBillingCategoryYearGuard.allowed) return deleteBillingCategoryYearGuard;
  const row = findBillingCategoryRowById(sheet, id);
  if (row === -1) throw new Error("Billing Category record not found.");

  sheet.deleteRow(row);

  safeLogAuditEvent(
    'Delete',
    'Billing Categories',
    id,
    appendAcademicYearOverrideAuditDetails('Deleted billing category ' + String(((existingCategory && existingCategory.category) || id)).trim(), deleteBillingCategoryYearGuard),
    buildBillingCategoryAuditSnapshot(existingCategory),
    null
  );

  return { success: true };
}


// 7. Generate Student Billings from a Billing Category
function generateStudentBillings(data) {
  try {
    Logger.log('generateStudentBillings called with: ' + JSON.stringify(data));
    
    const generationYearGuard = enforceAcademicYearCrudSecurity({
      action: 'Create',
      module: 'Invoices',
      recordId: String(data.billingCategoryId || data.category || '').trim(),
      academicYear: data.academicYear,
      newValue: buildAuditSnapshot({
        billingCategoryId: data.billingCategoryId,
        academicYear: data.academicYear,
        term: data.term,
        category: data.category,
        studentClass: data.studentClass,
        selectedStudents: (data.studentIds || []).length
      }),
      overrideConfirmed: data && data.__adminAcademicYearOverride === true
    });
    if (!generationYearGuard.allowed) return generationYearGuard;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invoicesSheet = ss.getSheetByName("Invoices");
    
    if (!invoicesSheet) {
      throw new Error("Invoices worksheet not found. Please create the Invoices sheet first.");
    }

    const studentIds = data.studentIds || [];
    if (studentIds.length === 0) {
      return { success: false, message: 'No students selected' };
    }

    // Get student details
    const studentsData = getStudentsData();
    
    // Get school information for bill header
    const schoolName = getSystemParameter('School Name') || 'GLOBAL EVANGELICAL BASIC SCHOOL - TETTEKOPE';
    const schoolAddress = getSystemParameter('School Address') || 'P.O. BOX 182, KETA';
    
    let successCount = 0;
    const errors = [];

    studentIds.forEach(function(studentId) {
      try {
        const student = studentsData.find(s => s.studentId === studentId);
        if (!student) {
          errors.push(`Student ${studentId} not found`);
          return;
        }

        // Check if invoice already exists for this student/year/term/category
        const existingInvoices = getInvoicesData();
        const duplicate = existingInvoices.find(inv => 
          inv.studentId === studentId &&
          inv.academicYear === data.academicYear &&
          inv.term === data.term &&
          inv.category === data.category
        );

        if (duplicate) {
          Logger.log('Skipping duplicate invoice for student: ' + studentId);
          errors.push(`Invoice already exists for ${student.firstName} ${student.lastName}`);
          return;
        }

        // Generate invoice ID
        const nextInvoiceId = generateNextInvoiceId(invoicesSheet);
        const lastRow = invoicesSheet.getLastRow();
        const targetRow = lastRow + 1;

        // Calculate due date (30 days from now)
        const issueDate = new Date();
        const dueDate = new Date(issueDate);
        dueDate.setDate(dueDate.getDate() + 30);

        // Parse items from billing category (comma-separated)
        const items = data.items || '';
        const totalAmount = data.totalAmount + (data.nursery || 0);
        
        // Get student's full name
        const studentName = `${student.firstName} ${student.lastName}`.trim();
        const studentClass = data.studentClass || student.class || '';

        // Write invoice record to sheet
        // Structure to match template: Invoice ID, Student ID, Student Name, Class, Academic Year, Term, 
        //                             Category, Debit Items, Total Amount, Issue Date, Due Date, Status, Payment Status
        invoicesSheet.getRange(targetRow, 2, 1, 13).setValues([
          [
            nextInvoiceId,
            student.studentId,
            studentName,
            studentClass,
            data.academicYear,
            data.term,
            data.category,
            items, // Debit items (comma-separated list)
            totalAmount,
            issueDate,
            dueDate,
            'Pending', // Status
            'Unpaid'   // Payment Status
          ]
        ]);

        const createdInvoice = buildGeneratedInvoiceAuditSnapshot({
          invoiceId: nextInvoiceId,
          studentId: student.studentId,
          studentName: studentName,
          studentClass: studentClass,
          academicYear: data.academicYear,
          term: data.term,
          category: data.category,
          items: items,
          amountDue: totalAmount,
          issueDate: issueDate,
          dueDate: dueDate,
          status: 'Pending',
          paymentStatus: 'Unpaid',
          source: 'Billing Category Generation'
        });

        safeLogAuditEvent(
          'Create',
          'Invoices',
          nextInvoiceId,
          appendAcademicYearOverrideAuditDetails('Generated invoice from billing category ' + String(data.category || '').trim() + ' for ' + studentName, generationYearGuard),
          null,
          createdInvoice
        );

        successCount++;
        Logger.log('Invoice created: ' + nextInvoiceId + ' for student: ' + studentId + ' - ' + studentName);
        
      } catch (err) {
        Logger.log('Error creating invoice for student ' + studentId + ': ' + err.toString());
        errors.push(`Error for ${studentId}: ${err.message}`);
      }
    });

    if (errors.length > 0) {
      Logger.log('Some invoices failed: ' + errors.join('; '));
    }

    safeLogAuditEvent(
      'Generate',
      'Billing Categories',
      String(data.category || '').trim(),
      appendAcademicYearOverrideAuditDetails('Generated ' + successCount + ' invoice(s) from billing category ' + String(data.category || '').trim(), generationYearGuard),
      null,
      buildAuditSnapshot({
        academicYear: data.academicYear,
        term: data.term,
        category: data.category,
        studentClass: data.studentClass,
        selectedStudents: studentIds.length,
        generatedInvoices: successCount,
        errors: errors.length
      })
    );

    return { 
      success: true, 
      count: successCount,
      message: `Successfully generated ${successCount} invoice(s)` + (errors.length > 0 ? `. ${errors.length} failed: ${errors.join(', ')}` : '')
    };
    
  } catch (error) {
    Logger.log('ERROR in generateStudentBillings: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Helper function to generate next Invoice ID
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

// 8. Generate Pupil's Bill PDF (matching template format)
function generatePupilBill(invoiceId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invoicesSheet = ss.getSheetByName("Invoices");
    
    if (!invoicesSheet) {
      throw new Error("Invoices sheet not found");
    }

    // Find invoice row
    const row = findInvoiceRowById(invoicesSheet, invoiceId);
    if (row === -1) {
      throw new Error("Invoice not found");
    }

    // Get invoice data (columns B through N)
    const invoiceData = invoicesSheet.getRange(row, 2, 1, 13).getValues()[0];
    
    const studentId = String(invoiceData[1]).trim();
    const studentName = String(invoiceData[2]).trim();
    const studentClass = String(invoiceData[3]).trim();
    const academicYear = String(invoiceData[4]).trim();
    const term = String(invoiceData[5]).trim();
    const category = String(invoiceData[6]).trim();
    const items = String(invoiceData[7]).trim(); // Comma-separated debit items
    const totalAmount = Number(invoiceData[8]) || 0;

    // Get school information
    const schoolName = getSystemParameter('School Name') || 'GLOBAL EVANGELICAL BASIC SCHOOL - TETTEKOPE';
    const schoolAddress = getSystemParameter('School Address') || 'P.O. BOX 182, KETA';
    const rawSchoolLogo = getSystemParameter('School Logo') || getSystemParameter('Logo URL') || '';
    const schoolLogo = getImageAsBase64(rawSchoolLogo);

    // Parse debit items
    const itemsList = items.split(',').map(i => i.trim()).filter(i => i !== '');
    
    // Get billing items with prices from Billings sheet
    const billingsData = getBillingsData();
    const billingsMap = {};
    billingsData.forEach(b => {
      billingsMap[b.item.trim()] = b.amount;
    });

    // Build debit rows
    let debitRows = '';
    itemsList.forEach(itemName => {
      const amount = billingsMap[itemName] || 0;
      const ghc = Math.floor(amount);
      const pesewas = Math.round((amount - ghc) * 100);
      
      debitRows += `
        <tr>
          <td style="padding: 8px; border: 1px solid #000;">${itemName}</td>
          <td style="padding: 8px; border: 1px solid #000; text-align: center;">Per Term</td>
          <td style="padding: 8px; border: 1px solid #000; text-align: right;">${ghc}</td>
          <td style="padding: 8px; border: 1px solid #000; text-align: right;">${String(pesewas).padStart(2, '0')}</td>
          <td style="padding: 8px; border: 1px solid #000;"></td>
          <td style="padding: 8px; border: 1px solid #000;"></td>
        </tr>`;
    });

    // Calculate total
    const totalGhc = Math.floor(totalAmount);
    const totalPesewas = Math.round((totalAmount - totalGhc) * 100);

    const html = buildPupilBillHTML(schoolName, schoolAddress, schoolLogo, studentName, studentClass, academicYear, debitRows, totalGhc, totalPesewas);
    
    const fileName = `Pupil_Bill_${studentId}_${term}_${academicYear}.pdf`;
    const blob = Utilities.newBlob(html, 'text/html', 'bill.html').getAs('application/pdf');
    blob.setName(fileName);

    return {
      base64: Utilities.base64Encode(blob.getBytes()),
      fileName: fileName
    };
    
  } catch (error) {
    Logger.log('ERROR in generatePupilBill: ' + error.toString());
    throw error;
  }
}

function buildPupilBillHTML(schoolName, schoolAddress, schoolLogo, studentName, studentClass, academicYear, debitRows, totalGhc, totalPesewas) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Pupil's Bill</title>
  <style>
    @page { size: A4 portrait; margin: 15mm; }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: Arial, sans-serif; font-size: 12px; line-height: 1.4; }
    .header { text-align: center; margin-bottom: 20px; position: relative; }
    .logo { width: 80px; height: 80px; position: absolute; }
    .logo-left { left: 20px; top: 0; }
    .logo-right { right: 20px; top: 0; }
    .school-name { font-size: 18px; font-weight: bold; margin: 10px 0; }
    .school-address { font-size: 12px; margin: 5px 0; }
    .bill-title { font-size: 16px; font-weight: bold; margin: 10px 0; }
    .info-row { display: flex; justify-content: space-between; margin: 8px 0; font-size: 13px; }
    .info-row strong { margin-right: 10px; }
    table { width: 100%; border-collapse: collapse; margin: 15px 0; }
    th { background: #f0f0f0; padding: 8px; border: 1px solid #000; font-size: 12px; font-weight: bold; text-align: center; }
    td { padding: 8px; border: 1px solid #000; font-size: 12px; }
    .total-row { font-weight: bold; background: #f8f8f8; }
    .notes { margin-top: 20px; font-size: 11px; line-height: 1.6; }
    .notes strong { display: block; margin-bottom: 8px; }
    .notes ol { margin-left: 20px; }
    .notes li { margin-bottom: 8px; }
  </style>
</head>
<body>
  <div class="header">
    ${schoolLogo ? `<img src="${schoolLogo}" class="logo logo-left" alt="Logo" />` : ''}
    ${schoolLogo ? `<img src="${schoolLogo}" class="logo logo-right" alt="Logo" />` : ''}
    <div class="school-name">${schoolName}</div>
    <div class="school-address">${schoolAddress}</div>
    <div class="bill-title">PUPIL'S BILL</div>
  </div>

  <div class="info-row">
    <div><strong>Name:</strong> ${studentName}</div>
  </div>
  <div class="info-row">
    <div><strong>Class:</strong> ${studentClass}</div>
    <div><strong>Year:</strong> ${academicYear}</div>
  </div>

  <table>
    <thead>
      <tr>
        <th rowspan="2" style="width: 40%;">DEBIT</th>
        <th rowspan="2" style="width: 15%;">TERM</th>
        <th colspan="2" style="width: 22.5%;">CREDIT</th>
        <th colspan="2" style="width: 22.5%;"></th>
      </tr>
      <tr>
        <th>GH¢</th>
        <th>P</th>
        <th>GH¢</th>
        <th>P</th>
      </tr>
    </thead>
    <tbody>
      ${debitRows}
      <tr class="total-row">
        <td colspan="2">Total Amount Due</td>
        <td style="text-align: right;">${totalGhc}</td>
        <td style="text-align: right;">${String(totalPesewas).padStart(2, '0')}</td>
        <td></td>
        <td></td>
      </tr>
    </tbody>
  </table>

  <div class="notes">
    <strong>NOTE:</strong>
    <ol>
      <li>If there is any mistake on your bill, please bring it to the attention of the headmaster/teacher immediately so it can be rectified.</li>
      <li>All bills must be settled in full on or before *deadline date*. Please make special arrangement for *deadline date*, because after that, any coming to settle his/her bill for next term during the school time will be disrupted and not doing will incur penalty charges.</li>
      <li>If you cannot pay your ward school fees within the *deadline date*, please contact the school (Global Evangelical Basic School) for any assistance.</li>
    </ol>
  </div>
</body>
</html>`;
}


// 9. Generate Single Pupil Bill (for preview)
function generateSinglePupilBill(reportData, student) {
  try {
    Logger.log('=== generateSinglePupilBill START ===');
    Logger.log('Report Data: ' + JSON.stringify(reportData));
    Logger.log('Student: ' + JSON.stringify(student));
    
    // Get school information
    const schoolName = getSystemParameter('School Name') || 'GLOBAL EVANGELICAL BASIC SCHOOL - TETTEKOPE';
    const schoolAddress = getSystemParameter('School Address') || 'P.O. BOX 182, KETA';
    const rawSchoolLogo = getSystemParameter('School Logo') || getSystemParameter('Logo URL') || '';
    const schoolLogo = getImageAsBase64(rawSchoolLogo);

    // Parse debit items
    const items = reportData.items || '';
    const itemsList = items.split(',').map(i => i.trim()).filter(i => i !== '');
    
    // Get billing items with prices from Billings sheet
    const billingsData = getBillingsData();
    const billingsMap = {};
    billingsData.forEach(b => {
      billingsMap[b.item.trim()] = b.amount;
    });

    // Build debit rows
    let debitRows = '';
    let calculatedTotal = 0;
    
    itemsList.forEach(itemName => {
      const amount = billingsMap[itemName] || 0;
      calculatedTotal += amount;
      const ghc = Math.floor(amount);
      const pesewas = Math.round((amount - ghc) * 100);
      
      debitRows += `
        <tr>
          <td style="padding: 8px; border: 1px solid #000;">${itemName}</td>
          <td style="padding: 8px; border: 1px solid #000; text-align: center;">Per Term</td>
          <td style="padding: 8px; border: 1px solid #000; text-align: right;">${ghc}</td>
          <td style="padding: 8px; border: 1px solid #000; text-align: right;">${String(pesewas).padStart(2, '0')}</td>
          <td style="padding: 8px; border: 1px solid #000;"></td>
          <td style="padding: 8px; border: 1px solid #000;"></td>
        </tr>`;
    });

    // Calculate total (use calculated total from items)
    const totalAmount = calculatedTotal + (reportData.nursery || 0);
    const totalGhc = Math.floor(totalAmount);
    const totalPesewas = Math.round((totalAmount - totalGhc) * 100);

    const html = buildPupilBillHTML(
      schoolName, 
      schoolAddress, 
      schoolLogo, 
      student.studentName, 
      student.studentClass || reportData.studentClass, 
      reportData.academicYear, 
      debitRows, 
      totalGhc, 
      totalPesewas
    );
    
    const fileName = `Pupil_Bill_${student.studentName.replace(/\s+/g, '_')}_${reportData.term}_${reportData.academicYear}.pdf`;
    const blob = Utilities.newBlob(html, 'text/html', 'bill.html').getAs('application/pdf');
    blob.setName(fileName);

    Logger.log('=== generateSinglePupilBill END ===');
    
    return {
      base64: Utilities.base64Encode(blob.getBytes()),
      fileName: fileName,
      blob: blob
    };
    
  } catch (error) {
    Logger.log('ERROR in generateSinglePupilBill: ' + error.toString());
    throw error;
  }
}

// 10. Generate Multiple Pupil Bills (returns ZIP file)
function generateMultiplePupilBills(reportData, students) {
  try {
    Logger.log('Generating multiple pupil bills for ' + students.length + ' students');
    
    const blobs = [];
    
    students.forEach(function(student) {
      try {
        const result = generateSinglePupilBill(reportData, student);
        blobs.push(result.blob);
      } catch (err) {
        Logger.log('Error generating bill for student ' + student.studentId + ': ' + err.toString());
      }
    });

    if (blobs.length === 0) {
      throw new Error('No bills were generated successfully');
    }

    const zipFileName = `Pupil_Bills_${reportData.studentClass}_${reportData.term}_${reportData.academicYear}.zip`;
    const zipBlob = Utilities.zip(blobs, zipFileName);

    return {
      base64: Utilities.base64Encode(zipBlob.getBytes()),
      fileName: zipFileName
    };
    
  } catch (error) {
    Logger.log('ERROR in generateMultiplePupilBills: ' + error.toString());
    throw error;
  }
}
