// --- LOGIN & SECURITY ---

// 1. Legacy login helper kept for backwards compatibility.
// The active login flow now authenticates against the Users sheet in User.gs.
function legacyLoginUser(username, password) {
  const props = PropertiesService.getScriptProperties();
  const validUsername = props.getProperty('ADMIN_USERNAME') || 'admin';
  const validPassword = props.getProperty('ADMIN_PASSWORD') || 'password123';

  if (username === validUsername && password === validPassword) {
    props.setProperty('IS_LOGGED_IN', 'true');
    props.setProperty('LOGGED_IN_ROLE', 'Admin');
    return { success: true };
  } else {
    return { success: false, message: 'Invalid username or password' };
  }
}

// NOTE: checkSession() and logoutUser() moved to User.gs (primary location)


// --- REPORTS FUNCTIONS ---

// Helper to safely fetch school name from Settings sheet with full fallback
function getSchoolNameFromSettings() {
  try {
    if (typeof getParameters === 'function') {
      const params = getParameters();
      if (params && (params['School Name'] || params['schoolName'])) {
        return params['School Name'] || params['schoolName'];
      }
    }
    if (typeof getSystemParameter === 'function') {
      const name = getSystemParameter('School Name');
      if (name) return name;
    }
  } catch(e) {}
  return 'GLOBAL EVANGELICAL BASIC SCHOOL, TETTEKOPE';
}

// Helper to fetch direct school logo URL from Settings sheet
function getSchoolLogoFromSettings() {
  try {
    if (typeof getParameters === 'function') {
      const params = getParameters();
      if (params && (params['schoolLogoDirectUrl'] || params['School Logo URL'])) {
        return params['schoolLogoDirectUrl'] || params['School Logo URL'];
      }
    }
  } catch(e) {}
  return 'https://lh3.googleusercontent.com/d/1MVnH55BHBynLBOD4pLIZE-5Y9gMaJbOe';
}

// 4. Gather comprehensive data for the Reports page
function getReportData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Student Count
  const studentSheet = ss.getSheetByName("Students");
  const studentCount = studentSheet ? studentSheet.getLastRow() - 2 : 0;

  // 2. Invoice & Payment Summary
  const invoiceSheet = ss.getSheetByName("Invoices");
  let totalInvoiced = 0;
  let totalPaid = 0;
  let totalUnpaid = 0;
  let invoiceCount = 0;

  if (invoiceSheet && invoiceSheet.getLastRow() > 2) {
    const invData = invoiceSheet.getRange(3, 5, invoiceSheet.getLastRow() - 2, 3).getValues(); // Col E(Amount), F(Status)
    invoiceCount = invData.length;
    invData.forEach(row => {
      const amount = Number(row[0]) || 0;
      const status = String(row[1]).trim();
      
      totalInvoiced += amount;
      if (status === 'Paid') totalPaid += amount;
      else totalUnpaid += amount;
    });
  }

  // 3. Attendance Average
  const attSheet = ss.getSheetByName("Attendance");
  let attendanceAvg = 0;
  let attTotal = 0;
  let attPresent = 0;

  if (attSheet && attSheet.getLastRow() > 2) {
    const attData = attSheet.getRange(3, 5, attSheet.getLastRow() - 2, 1).getValues(); // Col E(Status)
    attTotal = attData.length;
    attData.forEach(row => {
      if (String(row[0]).trim() === 'Present') attPresent++;
    });
    attendanceAvg = attTotal > 0 ? Math.round((attPresent / attTotal) * 100) : 0;
  }

  // 4. Top Performers (Based on Performance sheet)
  const perfSheet = ss.getSheetByName("Performance");
  let topStudents = [];
  if (perfSheet && perfSheet.getLastRow() > 2) {
    // Get Student Name (Col C), Total (Col J), Rank (Col K)
    const perfData = perfSheet.getRange(3, 3, perfSheet.getLastRow() - 2, 3).getValues(); 
    // Sort by Total descending
    perfData.sort((a, b) => Number(b[1]) - Number(a[1]));
    // Take top 5
    topStudents = perfData.slice(0, 5).map(row => ({
      name: String(row[0]).trim(),
      total: Number(row[1]) || 0,
      rank: String(row[2]).trim()
    }));
  }

  return {
    studentCount: Math.max(0, studentCount),
    invoiceCount: invoiceCount,
    totalInvoiced: totalInvoiced,
    totalPaid: totalPaid,
    totalUnpaid: totalUnpaid,
    attendanceAvg: attendanceAvg,
    topStudents: topStudents
  };
}


// 5. Generate Reports based on type and filters
function generateReport(reportType, filters) {
  try {
    Logger.log('Generating report: ' + reportType + ' with filters: ' + JSON.stringify(filters));
    
    switch(reportType) {
      case 'students':
        return generateStudentsReport(filters);
      case 'attendance':
        return generateAttendanceReport(filters);
      case 'performance':
        return generatePerformanceReport(filters);
      case 'financial':
        return generateFinancialReport(filters);
      case 'enrollment':
        return generateEnrollmentReport(filters);
      case 'class':
        return generateClassReport(filters);
      default:
        throw new Error('Unknown report type: ' + reportType);
    }
    
  } catch (error) {
    Logger.log('ERROR in generateReport: ' + error.toString());
    throw error;
  }
}

// Students Report Generator
function generateStudentsReport(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentsData = getStudentsData();
  
  // Apply filters
  let filteredData = studentsData;
  if (filters.class) {
    filteredData = filteredData.filter(s => s.class === filters.class);
  }
  if (filters.status) {
    filteredData = filteredData.filter(s => s.status === filters.status);
  }
  
  // Build HTML
  let tableRows = '';
  filteredData.forEach((s, index) => {
    tableRows += `
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${index + 1}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${s.studentId}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${s.firstName} ${s.lastName}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${s.class}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${s.gender}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${s.enrollmentDate}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;"><strong>${s.status}</strong></td>
      </tr>`;
  });
  
  const schoolName = getSchoolNameFromSettings();
  const schoolLogo = getSchoolLogoFromSettings();
  const html = buildReportHTML('Students Report', schoolName, schoolLogo, tableRows, 
    `<tr><th>No.</th><th>Student ID</th><th>Name</th><th>Class</th><th>Gender</th><th>Enrollment Date</th><th>Status</th></tr>`,
    `Total Students: ${filteredData.length}${filters.class ? ' | Class: ' + filters.class : ''}${filters.status ? ' | Status: ' + filters.status : ''}`
  );
  
  const fileName = `Students_Report_${new Date().toISOString().split('T')[0]}.pdf`;
  const blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    fileName: fileName
  };
}

// Attendance Report Generator
function generateAttendanceReport(filters) {
  const attendanceData = getAttendanceData();
  
  // Apply filters
  let filteredData = attendanceData.filter(a => a.academicYear === filters.year && a.term === filters.term);
  if (filters.class) {
    filteredData = filteredData.filter(a => a.className === filters.class);
  }
  
  // Apply date range filter if provided
  if (filters.startDate || filters.endDate) {
    filteredData = filteredData.filter(a => {
      if (!a.date) return false;
      
      const recordDate = new Date(a.date);
      let include = true;
      
      if (filters.startDate) {
        const startDate = new Date(filters.startDate);
        if (recordDate < startDate) include = false;
      }
      
      if (filters.endDate) {
        const endDate = new Date(filters.endDate);
        endDate.setHours(23, 59, 59, 999);
        if (recordDate > endDate) include = false;
      }
      
      return include;
    });
  }
  
  // Sort by date (most recent first)
  filteredData.sort((a, b) => {
    const dateA = new Date(a.date);
    const dateB = new Date(b.date);
    return dateB - dateA;
  });
  
  // Calculate stats
  const totalRecords = filteredData.length;
  const presentCount = filteredData.filter(a => a.status === 'Present').length;
  const absentCount = filteredData.filter(a => a.status === 'Absent').length;
  const attendanceRate = totalRecords > 0 ? ((presentCount / totalRecords) * 100).toFixed(1) : 0;
  
  // Build HTML
  let tableRows = '';
  filteredData.forEach((a, index) => {
    const statusColor = a.status === 'Present' ? '#10b981' : '#ef4444';
    tableRows += `
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${index + 1}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${a.date}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${a.studentName}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${a.className}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${a.courseId}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center; color: ${statusColor}; font-weight: bold;">${a.status}</td>
      </tr>`;
  });
  
  // Build summary text
  let summaryText = `${filters.year} | ${filters.term}`;
  if (filters.class) summaryText += ` | Class: ${filters.class}`;
  if (filters.startDate || filters.endDate) {
    summaryText += ` | Date Range: ${filters.startDate || 'Start'} to ${filters.endDate || 'End'}`;
  }
  summaryText += ` | Total: ${totalRecords} | Present: ${presentCount} | Absent: ${absentCount} | Attendance Rate: ${attendanceRate}%`;
  
  const schoolName = getSchoolNameFromSettings();
  const schoolLogo = getSchoolLogoFromSettings();
  const html = buildReportHTML('Attendance Report', schoolName, schoolLogo, tableRows,
    `<tr><th>No.</th><th>Date</th><th>Student</th><th>Class</th><th>Course</th><th>Status</th></tr>`,
    summaryText
  );
  
  const fileNameSuffix = filters.startDate && filters.endDate 
    ? `_${filters.startDate}_to_${filters.endDate}` 
    : `_${filters.term}`;
  const fileName = `Attendance_Report_${filters.year}${fileNameSuffix}.pdf`;
  const blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    fileName: fileName
  };
}

// Performance Report Generator  
function generatePerformanceReport(filters) {
  const performanceData = getPerformanceData();
  
  // Apply filters
  let filteredData = performanceData.filter(p => p.academicYear === filters.year && p.term === filters.term);
  if (filters.class) {
    filteredData = filteredData.filter(p => p.studentClass === filters.class);
  }
  
  // Build HTML
  let tableRows = '';
  filteredData.forEach((p, index) => {
    tableRows += `
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${index + 1}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${p.studentName}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${p.studentClass}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${p.course}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${p.classScore}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${p.examScore100}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;"><strong>${p.total}</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${p.rank}</td>
      </tr>`;
  });
  
  const schoolName = getSchoolNameFromSettings();
  const schoolLogo = getSchoolLogoFromSettings();
  const html = buildReportHTML('Performance Report', schoolName, schoolLogo, tableRows,
    `<tr><th>No.</th><th>Student</th><th>Class</th><th>Course</th><th>Class (50%)</th><th>Exam (100%)</th><th>Total</th><th>Rank</th></tr>`,
    `${filters.year} | ${filters.term}${filters.class ? ' | Class: ' + filters.class : ''} | Total Records: ${filteredData.length}`
  );
  
  const fileName = `Performance_Report_${filters.year}_${filters.term}.pdf`;
  const blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    fileName: fileName
  };
}

// Financial Report Generator
function generateFinancialReport(filters) {
  const invoicesData = getInvoicesData();
  const paymentsData = getPaymentsData();
  
  // Apply filters
  let filteredInvoices = invoicesData.filter(i => i.academicYear === filters.year);
  if (filters.term) {
    filteredInvoices = filteredInvoices.filter(i => i.term === filters.term);
  }
  if (filters.status) {
    const paidInvoices = paymentsData.filter(p => p.paymentStatus === 'Paid').map(p => p.invoiceId);
    if (filters.status === 'Paid') {
      filteredInvoices = filteredInvoices.filter(i => paidInvoices.includes(i.invoiceId));
    } else if (filters.status === 'Unpaid') {
      filteredInvoices = filteredInvoices.filter(i => !paidInvoices.includes(i.invoiceId) && i.status !== 'Paid');
    } else if (filters.status === 'Pending') {
      filteredInvoices = filteredInvoices.filter(i => i.status === 'Pending');
    }
  }
  
  // Calculate totals
  const totalInvoiced = filteredInvoices.reduce((sum, i) => sum + (i.amountDue || 0), 0);
  const paidAmount = paymentsData.filter(p => filteredInvoices.find(i => i.invoiceId === p.invoiceId) && p.paymentStatus === 'Paid')
    .reduce((sum, p) => sum + (p.amountPaid || 0), 0);
  const unpaidAmount = totalInvoiced - paidAmount;
  
  // Build HTML
  let tableRows = '';
  filteredInvoices.forEach((inv, index) => {
    const payment = paymentsData.find(p => p.invoiceId === inv.invoiceId);
    const payStatus = payment && payment.paymentStatus === 'Paid' ? 'Paid' : 'Unpaid';
    const statusColor = payStatus === 'Paid' ? '#10b981' : '#ef4444';
    
    tableRows += `
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${index + 1}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${inv.invoiceId}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${inv.studentId}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${inv.category}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">GHS ${inv.amountDue.toFixed(2)}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center; color: ${statusColor}; font-weight: bold;">${payStatus}</td>
      </tr>`;
  });
  
  const schoolName = getSchoolNameFromSettings();
  const schoolLogo = getSchoolLogoFromSettings();
  const html = buildReportHTML('Financial Report', schoolName, schoolLogo, tableRows,
    `<tr><th>No.</th><th>Invoice ID</th><th>Student ID</th><th>Category</th><th>Amount</th><th>Status</th></tr>`,
    `${filters.year}${filters.term ? ' | ' + filters.term : ''} | Total Invoiced: GHS ${totalInvoiced.toFixed(2)} | Paid: GHS ${paidAmount.toFixed(2)} | Unpaid: GHS ${unpaidAmount.toFixed(2)}`
  );
  
  const fileName = `Financial_Report_${filters.year}.pdf`;
  const blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    fileName: fileName
  };
}

// Enrollment & Class Reports
function generateEnrollmentReport(filters) {
  return generateStudentsReport(filters);
}

function generateClassReport(filters) {
  const classesData = getClassesData();
  const studentsData = getStudentsData();
  
  // Build HTML showing class sizes
  let tableRows = '';
  classesData.forEach((cls, index) => {
    const studentCount = studentsData.filter(s => s.class === cls.className).length;
    tableRows += `
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${index + 1}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${cls.className}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${cls.classTeacher || '-'}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;"><strong>${studentCount}</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${cls.academicYear}</td>
      </tr>`;
  });
  
  const schoolName = getSchoolNameFromSettings();
  const schoolLogo = getSchoolLogoFromSettings();
  const html = buildReportHTML('Class Report', schoolName, schoolLogo, tableRows,
    `<tr><th>No.</th><th>Class Name</th><th>Class Teacher</th><th>Students</th><th>Academic Year</th></tr>`,
    filters.year ? `Academic Year: ${filters.year}` : 'All Classes'
  );
  
  const fileName = `Class_Report_${new Date().toISOString().split('T')[0]}.pdf`;
  const blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    fileName: fileName
  };
}

// Helper to build report HTML for PDF export
function buildReportHTML(title, schoolName, schoolLogo, tableRows, headerRow, summary) {
  const logoHtml = schoolLogo ? `<div style="margin-bottom:8px;"><img src="${schoolLogo}" style="height:60px; max-width:180px; object-fit:contain;" alt="Logo"/></div>` : '';
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>${title}</title>
  <style>
    @page { size: A4 landscape; margin: 15mm; }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: Arial, sans-serif; font-size: 11px; color: #111; }
    .header { text-align: center; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 2px solid #333; }
    .header h1 { font-size: 18px; margin-bottom: 5px; text-transform: uppercase; letter-spacing: 0.5px; }
    .header h2 { font-size: 14px; margin-bottom: 10px; color: #2563eb; font-weight: bold; }
    .summary { background: #f0f4f8; padding: 10px 14px; margin-bottom: 15px; border-radius: 4px; font-size: 11px; border: 1px solid #cbd5e1; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
    th { background: #0f172a; color: white; padding: 10px 8px; border: 1px solid #0f172a; font-size: 11px; text-align: left; text-transform: uppercase; }
    td { padding: 8px; border: 1px solid #cbd5e1; font-size: 11px; }
    tr:nth-child(even) { background: #f8fafc; }
    .footer { text-align: center; margin-top: 20px; padding-top: 15px; border-top: 1px solid #ddd; font-size: 10px; color: #666; }
  </style>
</head>
<body>
  <div class="header">
    ${logoHtml}
    <h1>${schoolName}</h1>
    <h2>${title}</h2>
    <div style="font-size: 10px; color: #64748b;">Generated on: ${new Date().toLocaleDateString('en-GB')} at ${new Date().toLocaleTimeString()}</div>
  </div>
  <div class="summary"><strong>Report Summary:</strong> ${summary}</div>
  <table>
    <thead>${headerRow}</thead>
    <tbody>${tableRows || '<tr><td colspan="10" style="text-align: center; padding: 20px; color: #999;">No data available</td></tr>'}</tbody>
  </table>
  <div class="footer">
    <p>${schoolName} - Official System Report</p>
  </div>
</body>
</html>`;
}
