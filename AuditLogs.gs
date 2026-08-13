/* ============================================================
   AUDIT LOGS - JAVASCRIPT
   Separate JS file for Audit Trail functionality
============================================================ */

// STATE
var allLogs = [];
var filteredLogs = [];

// INITIALIZATION - Handles both initial load and dynamic/iframe load
function initAuditLogs() {
  setDefaultDates();
  loadAuditLogs();
}

if (typeof document !== 'undefined') {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initAuditLogs);
  } else {
    // DOM is already ready (e.g. embedded iframe or dynamic load)
    initAuditLogs();
  }
}

function setDefaultDates() {
  const today = new Date();
  const thirtyDaysAgo = new Date(today);
  thirtyDaysAgo.setDate(today.getDate() - 30);
  
  const fromInput = document.getElementById('filterDateFrom');
  const toInput = document.getElementById('filterDateTo');
  if (fromInput) fromInput.value = thirtyDaysAgo.toISOString().split('T')[0];
  if (toInput) toInput.value = today.toISOString().split('T')[0];
}

// LOAD DATA
function loadAuditLogs() {
  showLoading();
  
  if (typeof google !== 'undefined' && google.script && google.script.run) {
    google.script.run
      .withSuccessHandler(function(logs) {
        allLogs = logs || [];
        filterLogs(); // Filter and render logs with active default dates
        populateFilters();
      })
      .withFailureHandler(function(error) {
        showToast('Error', 'Failed to load audit logs: ' + (error ? error.message || error : 'Unknown error'), 'error');
        hideLoading();
      })
      .getAuditLogs();
  } else {
    console.warn('google.script.run is not available (running in preview mode).');
    allLogs = [];
    filteredLogs = [];
    renderLogs();
    updateStats();
    populateFilters();
  }
}

function refreshLogs() {
  loadAuditLogs();
  showToast('Success', 'Audit logs refreshed successfully', 'success');
}

// RENDER LOGS
function renderLogs() {
  const tbody = document.getElementById('tableBody');
  const loading = document.getElementById('loading');
  const table = document.getElementById('logsTable');
  const logCount = document.getElementById('logCount');
  
  loading.style.display = 'none';
  table.style.display = 'table';
  logCount.textContent = filteredLogs.length;
  
  tbody.innerHTML = '';
  
  if (filteredLogs.length === 0) {
    tbody.innerHTML = `
      <tr><td colspan="7">
        <div class="empty-state">
          <div class="empty-state-icon">📋</div>
          <h4>No Audit Logs Found</h4>
          <p>No logs match your current filters.</p>
        </div>
      </td></tr>`;
    return;
  }
  
  filteredLogs.forEach(function(log) {
    const tr = document.createElement('tr');
    
    // Format timestamp
    const timestamp = formatTimestamp(log.timestamp);
    
    // Get action badge
    const actionBadge = getActionBadge(log.action);
    
    // Get timeline indicator
    const timelineIndicator = getTimelineIndicator(log.action);
    
    // Format user
    const userDisplay = formatUser(log.userId, log.userName);
    
    // Truncate details
    const details = truncateText(log.details || '-', 50);
    
    tr.innerHTML = `
      <td>
        <span class="timeline-indicator ${timelineIndicator}"></span>
        ${timestamp}
      </td>
      <td>${userDisplay}</td>
      <td><span class="badge ${actionBadge}">${escapeHtml(log.action)}</span></td>
      <td>${escapeHtml(log.module || '-')}</td>
      <td>${escapeHtml(details)}</td>
      <td style="font-family:monospace;font-size:12px;color:#64748b;">${escapeHtml(log.ipAddress || '-')}</td>
      <td style="text-align:center;">
        <button class="btn btn-secondary" style="padding:6px 12px;font-size:12px;" onclick='viewDetails(${JSON.stringify(log).replace(/'/g, "&apos;")})'>
          👁️ View
        </button>
      </td>`;
    
    tbody.appendChild(tr);
  });
}

// FILTER LOGS
function filterLogs() {
  const searchVal = document.getElementById('searchInput').value.toLowerCase();
  const dateFrom = document.getElementById('filterDateFrom').value;
  const dateTo = document.getElementById('filterDateTo').value;
  const actionVal = document.getElementById('filterAction').value;
  const userVal = document.getElementById('filterUser').value;
  const moduleVal = document.getElementById('filterModule').value;
  
  filteredLogs = allLogs.filter(function(log) {
    // Search filter
    const searchMatch = !searchVal || 
      (log.userName || '').toLowerCase().includes(searchVal) ||
      (log.action || '').toLowerCase().includes(searchVal) ||
      (log.module || '').toLowerCase().includes(searchVal) ||
      (log.details || '').toLowerCase().includes(searchVal) ||
      (log.recordId || '').toLowerCase().includes(searchVal);
    
    // Date filter
    let dateMatch = true;
    if (dateFrom || dateTo) {
      const logDate = new Date(log.timestamp);
      if (dateFrom) {
        const from = new Date(dateFrom);
        from.setHours(0, 0, 0, 0);
        dateMatch = dateMatch && logDate >= from;
      }
      if (dateTo) {
        const to = new Date(dateTo);
        to.setHours(23, 59, 59, 999);
        dateMatch = dateMatch && logDate <= to;
      }
    }
    
    // Action filter
    const actionMatch = !actionVal || log.action === actionVal;
    
    // User filter
    const userMatch = !userVal || log.userId === userVal;
    
    // Module filter
    const moduleMatch = !moduleVal || log.module === moduleVal;
    
    return searchMatch && dateMatch && actionMatch && userMatch && moduleMatch;
  });
  
  renderLogs();
  updateStats();
}

// UPDATE STATS
function updateStats() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const todayLogs = allLogs.filter(function(log) {
    const logDate = new Date(log.timestamp);
    logDate.setHours(0, 0, 0, 0);
    return logDate.getTime() === today.getTime();
  });
  
  const uniqueUsers = new Set(allLogs.map(log => log.userId)).size;
  
  const deletions = allLogs.filter(log => log.action === 'Delete').length;
  
  document.getElementById('statTotal').textContent = allLogs.length;
  document.getElementById('statToday').textContent = todayLogs.length;
  document.getElementById('statUsers').textContent = uniqueUsers;
  document.getElementById('statDeletions').textContent = deletions;
}

// POPULATE FILTERS
function populateFilters() {
  // Populate user filter
  const users = new Set();
  allLogs.forEach(log => {
    if (log.userId && log.userName) {
      users.add(JSON.stringify({ id: log.userId, name: log.userName }));
    }
  });
  
  const userFilter = document.getElementById('filterUser');
  userFilter.innerHTML = '<option value="">All Users</option>';
  Array.from(users).map(u => JSON.parse(u)).forEach(user => {
    const option = document.createElement('option');
    option.value = user.id;
    option.textContent = user.name;
    userFilter.appendChild(option);
  });
  
  // Populate module filter
  const modules = new Set();
  allLogs.forEach(log => {
    if (log.module) modules.add(log.module);
  });
  
  const moduleFilter = document.getElementById('filterModule');
  moduleFilter.innerHTML = '<option value="">All Modules</option>';
  Array.from(modules).sort().forEach(module => {
    const option = document.createElement('option');
    option.value = module;
    option.textContent = module;
    moduleFilter.appendChild(option);
  });
}

// VIEW DETAILS
function viewDetails(log) {
  const modal = document.getElementById('detailModal');
  const content = document.getElementById('detailContent');
  
  content.innerHTML = `
    <div class="detail-row">
      <div class="detail-label">Log ID</div>
      <div class="detail-value">${escapeHtml(log.logId || '-')}</div>
    </div>
    <div class="detail-row">
      <div class="detail-label">Timestamp</div>
      <div class="detail-value">${formatTimestamp(log.timestamp, true)}</div>
    </div>
    <div class="detail-row">
      <div class="detail-label">User ID</div>
      <div class="detail-value">${escapeHtml(log.userId || '-')}</div>
    </div>
    <div class="detail-row">
      <div class="detail-label">User Name</div>
      <div class="detail-value">${escapeHtml(log.userName || '-')}</div>
    </div>
    <div class="detail-row">
      <div class="detail-label">Action</div>
      <div class="detail-value"><span class="badge ${getActionBadge(log.action)}">${escapeHtml(log.action)}</span></div>
    </div>
    <div class="detail-row">
      <div class="detail-label">Module</div>
      <div class="detail-value">${escapeHtml(log.module || '-')}</div>
    </div>
    <div class="detail-row">
      <div class="detail-label">Record ID</div>
      <div class="detail-value">${escapeHtml(log.recordId || '-')}</div>
    </div>
    <div class="detail-row">
      <div class="detail-label">Details</div>
      <div class="detail-value">${escapeHtml(log.details || '-')}</div>
    </div>
    <div class="detail-row">
      <div class="detail-label">IP Address</div>
      <div class="detail-value" style="font-family:monospace;">${escapeHtml(log.ipAddress || '-')}</div>
    </div>
    <div class="detail-row">
      <div class="detail-label">User Agent</div>
      <div class="detail-value" style="font-size:12px;color:#64748b;">${escapeHtml(log.userAgent || '-')}</div>
    </div>
    ${log.oldValue ? `
    <div class="detail-row">
      <div class="detail-label">Old Value</div>
      <div class="detail-value"><pre>${escapeHtml(log.oldValue)}</pre></div>
    </div>` : ''}
    ${log.newValue ? `
    <div class="detail-row">
      <div class="detail-label">New Value</div>
      <div class="detail-value"><pre>${escapeHtml(log.newValue)}</pre></div>
    </div>` : ''}
  `;
  
  modal.classList.add('active');
}

function closeDetailModal() {
  document.getElementById('detailModal').classList.remove('active');
}

// EXPORT LOGS
function exportLogs() {
  if (filteredLogs.length === 0) {
    showToast('Warning', 'No logs to export', 'warning');
    return;
  }
  
  showToast('Info', 'Exporting audit logs...', 'info');
  
  // Prepare data for export
  const exportData = filteredLogs.map(log => ({
    'Log ID': log.logId || '',
    'Timestamp': formatTimestamp(log.timestamp, true),
    'User ID': log.userId || '',
    'User Name': log.userName || '',
    'Action': log.action || '',
    'Module': log.module || '',
    'Record ID': log.recordId || '',
    'Details': log.details || '',
    'IP Address': log.ipAddress || '',
    'User Agent': log.userAgent || ''
  }));
  
  // Create worksheet
  const ws = XLSX.utils.json_to_sheet(exportData);
  
  // Create workbook
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Audit Logs');
  
  // Generate filename
  const dateStr = new Date().toISOString().split('T')[0];
  const filename = `AuditLogs_${dateStr}.xlsx`;
  
  // Download
  XLSX.writeFile(wb, filename);
  
  showToast('Success', 'Audit logs exported successfully', 'success');
}

// CLEAR OLD LOGS
function openClearModal() {
  document.getElementById('clearModal').classList.add('active');
}

function closeClearModal() {
  document.getElementById('clearModal').classList.remove('active');
}

function clearOldLogs() {
  const days = parseInt(document.getElementById('clearDays').value);
  
  if (!days || days < 1) {
    showToast('Warning', 'Please enter a valid number of days', 'warning');
    return;
  }
  
  if (!confirm(`Are you sure you want to delete audit logs older than ${days} days? This action cannot be undone.`)) {
    return;
  }
  
  showToast('Info', 'Deleting old audit logs...', 'info');
  
  google.script.run
    .withSuccessHandler(function(result) {
      closeClearModal();
      if (result.success) {
        showToast('Success', result.message, 'success');
        loadAuditLogs();
      } else {
        showToast('Error', result.message, 'error');
      }
    })
    .withFailureHandler(function(error) {
      showToast('Error', 'Failed to delete logs: ' + error.message, 'error');
    })
    .clearOldAuditLogs(days);
}

// UTILITY FUNCTIONS
function formatTimestamp(timestamp, includeFull) {
  if (!timestamp) return '-';
  
  const date = new Date(timestamp);
  const now = new Date();
  const diffMs = now - date;
  const diffMins = Math.floor(diffMs / 60000);
  const diffHours = Math.floor(diffMs / 3600000);
  const diffDays = Math.floor(diffMs / 86400000);
  
  let relative = '';
  if (diffMins < 1) relative = 'Just now';
  else if (diffMins < 60) relative = diffMins + ' min ago';
  else if (diffHours < 24) relative = diffHours + ' hour' + (diffHours > 1 ? 's' : '') + ' ago';
  else if (diffDays < 7) relative = diffDays + ' day' + (diffDays > 1 ? 's' : '') + ' ago';
  else relative = date.toLocaleDateString();
  
  if (includeFull) {
    const full = date.toLocaleString();
    return `${full} (${relative})`;
  }
  
  return relative;
}

function formatUser(userId, userName) {
  if (!userId && !userName) return '-';
  
  const initials = userName ? userName.split(' ').map(n => n.charAt(0).toUpperCase()).join('').substring(0, 2) : 'U';
  const name = userName || userId;
  
  return `
    <div class="user-badge">
      <div class="user-avatar">${initials}</div>
      <span>${escapeHtml(name)}</span>
    </div>`;
}

function getActionBadge(action) {
  const badges = {
    'Create': 'badge-create',
    'Update': 'badge-update',
    'Delete': 'badge-delete',
    'View': 'badge-view',
    'Login': 'badge-login',
    'Logout': 'badge-logout',
    'Export': 'badge-export',
    'Import': 'badge-import',
    'Generate': 'badge-generate',
    'Warning': 'badge-warning',
    'Denied': 'badge-denied'
  };
  return badges[action] || 'badge-other';
}

function getTimelineIndicator(action) {
  const indicators = {
    'Create': 'timeline-create',
    'Update': 'timeline-update',
    'Delete': 'timeline-delete',
    'View': 'timeline-view',
    'Generate': 'timeline-generate',
    'Warning': 'timeline-warning',
    'Denied': 'timeline-denied'
  };
  return indicators[action] || 'timeline-other';
}

function truncateText(text, maxLength) {
  if (!text || text.length <= maxLength) return text;
  return text.substring(0, maxLength) + '...';
}

function escapeHtml(text) {
  if (!text) return '';
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function showLoading() {
  document.getElementById('loading').style.display = 'block';
  document.getElementById('logsTable').style.display = 'none';
}

function hideLoading() {
  document.getElementById('loading').style.display = 'none';
  document.getElementById('logsTable').style.display = 'table';
}

// TOAST NOTIFICATIONS
function showToast(title, message, type) {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = 'toast';
  if (type) toast.classList.add('toast-' + type);
  
  const iconClass = type === 'error' ? 'error' : type === 'warning' ? 'error' : 'success';
  const iconText = type === 'error' ? '✗' : type === 'warning' ? '⚠' : '✓';
  
  toast.innerHTML = `
    <div class="toast-icon ${iconClass}">${iconText}</div>
    <div class="toast-body">
      <div class="toast-title">${escapeHtml(title)}</div>
      <div class="toast-message">${escapeHtml(message)}</div>
    </div>`;
  
  container.appendChild(toast);
  
  setTimeout(() => toast.classList.add('active'), 10);
  
  setTimeout(() => {
    toast.classList.remove('active');
    setTimeout(() => container.removeChild(toast), 350);
  }, 4000);
}
