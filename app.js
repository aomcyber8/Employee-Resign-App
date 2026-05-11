// ============================================
// Employee Resign App — Sena Development
// Monthly Sheets & Sorting Logic
// ============================================
import * as XLSX from 'xlsx';

// --- Constants ---
const STORAGE_KEY = 'sena_resign_data_v2';
const ITEMS_PER_PAGE = 25;

// --- State ---
let data = {
  sheets: [],
  employees: []
};
let currentSheetId = 'all'; // Default to 'all' tab
let currentFilters = {
  search: '',
  resignStatus: '',
  docStatus: '',
  hrStatus: '',
  factor: '',
  startDate: '',
  endDate: ''
};
let sortConfig = {
  column: 'terminationDate',
  direction: 'desc' // 'asc' or 'desc'
};
let currentPage = 1;
let editingId = null;
let deletingId = null;

// --- DOM Elements ---
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

const dom = {
  // Stats
  statTotal: $('#stat-total'),
  statActive: $('#stat-active'),
  statHrClosed: $('#stat-hr-closed'),
  statPending: $('#stat-pending'),
  statCanceled: $('#stat-canceled'),

  // Tabs
  sheetTabs: $('#sheet-tabs'),
  btnAddSheet: $('#btn-add-sheet'),
  emptyState: $('#empty-state'),

  // Table
  tableBody: $('#table-body'),
  dataTable: $('#data-table'),
  pagination: $('#pagination'),

  // Filters
  searchInput: $('#search-input'),
  filterResign: $('#filter-resign-status'),
  filterDoc: $('#filter-doc-status'),
  filterHr: $('#filter-hr-status'),
  filterFactor: $('#filter-factor'),
  filterStartDate: $('#filter-start-date'),
  filterEndDate: $('#filter-end-date'),
  sortHeaders: $$('.sortable'),

  // Modals
  modalOverlay: $('#modal-overlay'),
  addSheetOverlay: $('#add-sheet-overlay'),
  deleteOverlay: $('#delete-modal-overlay'),
  detailOverlay: $('#detail-modal-overlay'),

  // Forms
  form: $('#employee-form'),
  formId: $('#form-id'),
  fCode: $('#form-employeeCode'),
  fName: $('#form-fullName'),
  fPos: $('#form-position'),
  fDept: $('#form-department'),
  fEmail: $('#form-email'),
  fLastWork: $('#form-lastWorkDate'),
  fTerm: $('#form-terminationDate'),
  fResignStatus: $('#form-resignStatus'),
  fDocStatus: $('#form-documentStatus'),
  fHrStatus: $('#form-hrOnlineStatus'),
  fFactor: $('#form-resignFactor'),
  fReason: $('#form-resignReason'),
  fDelEmail: $('#form-emailDeleted'),
  fDelRem: $('#form-remDeleted'),
  fDelCms: $('#form-cmsDeleted'),

  sheetForm: $('#sheet-form'),
  fSheetMonth: $('#form-sheet-month'),

  // Buttons
  btnImport: $('#btn-import'),
  fileImport: $('#file-import'),
  btnAdd: $('#btn-add'),
  btnExport: $('#btn-export'),
  btnCancel: $('#btn-cancel'),
  btnModalClose: $('#btn-modal-close'),
  btnSheetCancel: $('#btn-sheet-cancel'),
  btnSheetClose: $('#btn-sheet-close'),
  btnDeleteCancel: $('#btn-delete-cancel'),
  btnDeleteClose: $('#btn-delete-close'),
  btnDeleteConfirm: $('#btn-delete-confirm'),
  btnDetailClose: $('#btn-detail-close'),

  // Toast
  toastContainer: $('#toast-container'),
};

// ============================================
// DATA LAYER (localStorage)
// ============================================

function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      data = JSON.parse(raw);
    }
  } catch (e) {
    console.error("Failed to load data", e);
  }

  // Auto-select 'all' or first sheet if empty
  if (data.sheets.length > 0) {
    // Sort sheets by date desc
    data.sheets.sort((a, b) => b.id.localeCompare(a.id));
  }
}

function saveData() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}

function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

// Sheet Management
function getSheetIdFromDate(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  if (isNaN(d)) return null;
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`; // YYYY-MM format for easy sorting
}

function getSheetLabel(sheetId) {
  const [yearStr, monthStr] = sheetId.split('-');
  const year = parseInt(yearStr);
  const month = parseInt(monthStr) - 1;
  const d = new Date(year, month, 1);
  return d.toLocaleDateString('th-TH', { month: 'short', year: '2-digit' });
}

function ensureSheetExists(sheetId) {
  if (!sheetId) return;
  const exists = data.sheets.find(s => s.id === sheetId);
  if (!exists) {
    data.sheets.push({
      id: sheetId,
      label: getSheetLabel(sheetId)
    });
    data.sheets.sort((a, b) => b.id.localeCompare(a.id)); // Keep descending
    saveData();
  }
  return sheetId;
}

// ============================================
// CRUD OPERATIONS
// ============================================

function saveEmployee(empData) {
  const now = new Date().toISOString();
  const sheetId = ensureSheetExists(getSheetIdFromDate(empData.terminationDate));

  let emp;
  if (empData.id) {
    // Update
    const idx = data.employees.findIndex(e => e.id === empData.id);
    if (idx !== -1) {
      emp = { ...data.employees[idx], ...empData, updatedAt: now };
      data.employees[idx] = emp;
    }
  } else {
    // Insert
    emp = { ...empData, id: generateId(), createdAt: now, updatedAt: now };
    data.employees.push(emp);
  }

  saveData();
  
  // Switch to the sheet where the employee was saved
  if (sheetId) currentSheetId = sheetId;
  
  return emp;
}

function deleteEmployee(id) {
  data.employees = data.employees.filter(e => e.id !== id);
  saveData();
}

// ============================================
// FILTERING & SORTING
// ============================================

function getFilteredAndSortedEmployees() {
  let result = data.employees;

  // 1. Filter by Sheet (skip if 'all')
  if (currentSheetId !== 'all') {
    result = result.filter(e => getSheetIdFromDate(e.terminationDate) === currentSheetId);
  }

  // 2. Filter by Search
  if (currentFilters.search) {
    const q = currentFilters.search.toLowerCase();
    result = result.filter(e => 
      (e.fullName && e.fullName.toLowerCase().includes(q)) ||
      (e.employeeCode && e.employeeCode.toLowerCase().includes(q))
    );
  }

  // 3. Filter by Dropdowns
  if (currentFilters.resignStatus) {
    if (currentFilters.resignStatus === 'not_canceled') {
      result = result.filter(e => e.resignStatus !== 'ยกเลิกลาออก');
    } else {
      result = result.filter(e => e.resignStatus === currentFilters.resignStatus);
    }
  }
  if (currentFilters.docStatus) {
    result = result.filter(e => e.documentStatus === currentFilters.docStatus);
  }
  if (currentFilters.hrStatus) {
    result = result.filter(e => e.hrOnlineStatus === currentFilters.hrStatus);
  }
  if (currentFilters.factor) {
    result = result.filter(e => e.resignFactor === currentFilters.factor);
  }

  // 4. Filter by Date Range (Termination Date)
  if (currentFilters.startDate) {
    const start = new Date(currentFilters.startDate).getTime();
    result = result.filter(e => {
      if (!e.terminationDate) return false;
      return new Date(e.terminationDate).getTime() >= start;
    });
  }
  if (currentFilters.endDate) {
    // Add 24 hours to include the entire end day
    const end = new Date(currentFilters.endDate).getTime() + (24 * 60 * 60 * 1000) - 1;
    result = result.filter(e => {
      if (!e.terminationDate) return false;
      return new Date(e.terminationDate).getTime() <= end;
    });
  }

  // 5. Sort
  if (sortConfig.column) {
    result.sort((a, b) => {
      let valA = a[sortConfig.column] || '';
      let valB = b[sortConfig.column] || '';

      // Date sorting
      if (sortConfig.column.includes('Date')) {
        valA = valA ? new Date(valA).getTime() : 0;
        valB = valB ? new Date(valB).getTime() : 0;
      }

      if (valA < valB) return sortConfig.direction === 'asc' ? -1 : 1;
      if (valA > valB) return sortConfig.direction === 'asc' ? 1 : -1;
      return 0;
    });
  }

  return result;
}

// ============================================
// RENDERING
// ============================================

function formatDate(dateStr) {
  if (!dateStr) return '—';
  try {
    const d = new Date(dateStr);
    return d.toLocaleDateString('th-TH', { year: '2-digit', month: 'short', day: 'numeric' });
  } catch {
    return dateStr;
  }
}

function getBadgeClass(status) {
  if (!status) return 'tag';
  if (status.includes('ยกเลิก')) return 'badge-cancel';
  if (status.includes('ปลดออก')) return 'badge-fire';
  if (status.includes('ไม่ครบ')) return 'badge-early';
  if (status.includes('ลาออก')) return 'badge-resign';
  return 'tag';
}

function renderTabs() {
  let tabsHtml = '';
  
  // "รวม" Tab
  const allCount = data.employees.length;
  const isAllActive = currentSheetId === 'all' ? 'active' : '';
  tabsHtml += `
    <button class="sheet-tab ${isAllActive}" data-id="all">
      รวม <span class="sheet-tab-count">${allCount}</span>
    </button>
  `;

  // Monthly Tabs
  tabsHtml += data.sheets.map(sheet => {
    const count = data.employees.filter(e => getSheetIdFromDate(e.terminationDate) === sheet.id).length;
    const isActive = sheet.id === currentSheetId ? 'active' : '';
    return `
      <button class="sheet-tab ${isActive}" data-id="${sheet.id}">
        ${sheet.label} <span class="sheet-tab-count">${count}</span>
      </button>
    `;
  }).join('');

  dom.sheetTabs.innerHTML = tabsHtml;

  // Tab click events
  $$('.sheet-tab').forEach(btn => {
    btn.addEventListener('click', (e) => {
      currentSheetId = e.currentTarget.dataset.id;
      currentPage = 1;
      renderAll();
    });
  });
}

function renderStats(filteredList) {
  // Stats calculate based on current sheet OR all employees if 'all' tab is selected
  const sheetEmployees = currentSheetId === 'all' 
    ? data.employees 
    : data.employees.filter(e => getSheetIdFromDate(e.terminationDate) === currentSheetId);
  
  const total = sheetEmployees.length;
  const hrClosed = sheetEmployees.filter(e => e.hrOnlineStatus === 'ปิดลาออกแล้ว').length;
  const canceled = sheetEmployees.filter(e => e.resignStatus === 'ยกเลิกลาออก').length;
  const active = total - canceled;
  const pending = total - hrClosed - canceled;

  animateNumber(dom.statTotal, total);
  animateNumber(dom.statActive, active);
  animateNumber(dom.statHrClosed, hrClosed);
  animateNumber(dom.statPending, pending);
  animateNumber(dom.statCanceled, canceled);
}

function animateNumber(el, target) {
  const current = parseInt(el.textContent) || 0;
  if (current === target) return;

  const duration = 400;
  const startTime = performance.now();

  function update(now) {
    const elapsed = now - startTime;
    const progress = Math.min(elapsed / duration, 1);
    const eased = 1 - Math.pow(1 - progress, 3);
    el.textContent = Math.round(current + (target - current) * eased);
    if (progress < 1) requestAnimationFrame(update);
  }
  requestAnimationFrame(update);
}

function renderSortHeaders() {
  dom.sortHeaders.forEach(th => {
    th.removeAttribute('data-sort-dir');
    if (th.dataset.sort === sortConfig.column) {
      th.setAttribute('data-sort-dir', sortConfig.direction);
    }
  });
}

function renderTable(filteredList) {
  renderSortHeaders();

  if (data.sheets.length === 0) {
    dom.dataTable.style.display = 'none';
    dom.emptyState.style.display = 'flex';
    $('#empty-state-text').textContent = 'ยังไม่มีชีทข้อมูล กดปุ่ม + เพื่อสร้างชีทแรก';
    dom.pagination.innerHTML = '';
    return;
  }

  if (filteredList.length === 0) {
    dom.dataTable.style.display = 'none';
    dom.emptyState.style.display = 'flex';
    $('#empty-state-text').textContent = 'ไม่พบข้อมูลในชีทนี้ หรือตามเงื่อนไขที่ค้นหา';
    dom.pagination.innerHTML = '';
    return;
  }

  dom.dataTable.style.display = 'table';
  dom.emptyState.style.display = 'none';

  // Pagination Logic
  const totalPages = Math.ceil(filteredList.length / ITEMS_PER_PAGE);
  if (currentPage > totalPages) currentPage = totalPages;
  if (currentPage < 1) currentPage = 1;

  const startIdx = (currentPage - 1) * ITEMS_PER_PAGE;
  const pageData = filteredList.slice(startIdx, startIdx + ITEMS_PER_PAGE);

  dom.tableBody.innerHTML = pageData.map((emp, i) => `
    <tr data-action="view" data-id="${emp.id}" title="คลิกเพื่อดูรายละเอียด">
      <td class="td-no">${startIdx + i + 1}</td>
      <td class="td-code">${escapeHtml(emp.employeeCode)}</td>
      <td class="td-name">${escapeHtml(emp.fullName)}</td>
      <td>${escapeHtml(emp.department || '—')}</td>
      <td class="td-date">${formatDate(emp.terminationDate)}</td>
      <td>
        <span class="status-badge ${getBadgeClass(emp.resignStatus)}">${emp.resignStatus || '—'}</span>
      </td>
      <td>
        ${emp.documentStatus ? `<span class="tag tag-${emp.documentStatus.includes('ได้รับ')?'success':'warning'}">${emp.documentStatus}</span><br>` : ''}
        ${emp.hrOnlineStatus ? `<span class="tag tag-${emp.hrOnlineStatus.includes('ปิด')?'success':'warning'}">${emp.hrOnlineStatus}</span>` : ''}
      </td>
      <td onclick="event.stopPropagation()">
        <div class="actions-cell">
          <button class="btn-icon" data-action="edit" data-id="${emp.id}" title="แก้ไข">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
              <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
            </svg>
          </button>
          <button class="btn-icon btn-icon-danger" data-action="delete" data-id="${emp.id}" title="ลบ">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <polyline points="3 6 5 6 21 6"/>
              <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
            </svg>
          </button>
        </div>
      </td>
    </tr>
  `).join('');

  renderPagination(filteredList.length, startIdx + 1, Math.min(startIdx + ITEMS_PER_PAGE, filteredList.length), totalPages);
}

function renderPagination(total, start, end, totalPages) {
  let html = `<div class="page-info">แสดง ${start}-${end} จาก ${total} รายการ</div>`;
  
  if (totalPages > 1) {
    html += `<div class="page-controls">
      <button class="btn-page" ${currentPage === 1 ? 'disabled' : ''} data-page="${currentPage - 1}">‹</button>
    `;
    
    for (let i = 1; i <= totalPages; i++) {
      // Show around current page logic (simple version)
      if (i === 1 || i === totalPages || (i >= currentPage - 1 && i <= currentPage + 1)) {
        html += `<button class="btn-page ${i === currentPage ? 'active' : ''}" data-page="${i}">${i}</button>`;
      } else if (i === currentPage - 2 || i === currentPage + 2) {
        html += `<span>...</span>`;
      }
    }
    
    html += `
      <button class="btn-page" ${currentPage === totalPages ? 'disabled' : ''} data-page="${currentPage + 1}">›</button>
    </div>`;
  } else {
    html += `<div></div>`;
  }
  
  dom.pagination.innerHTML = html;

  $$('.btn-page').forEach(btn => {
    btn.addEventListener('click', (e) => {
      if (!btn.disabled) {
        currentPage = parseInt(btn.dataset.page);
        renderTable(getFilteredAndSortedEmployees());
      }
    });
  });
}

function renderAll() {
  renderTabs();
  const filtered = getFilteredAndSortedEmployees();
  renderStats(filtered); // Pass filtered if you want stats to respect filters, currently it calculates sheet total
  renderTable(filtered);
}

function escapeHtml(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

// ============================================
// MODAL MANAGEMENT
// ============================================

function openModal(overlay) {
  overlay.classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeModal(overlay) {
  overlay.classList.remove('open');
  document.body.style.overflow = '';
}

function openAddModal() {
  editingId = null;
  $('#modal-title').textContent = 'เพิ่มพนักงานลาออก';
  dom.form.reset();
  dom.formId.value = '';
  
  // Set default dates to current sheet month if exists, else today
  let defaultDate = new Date();
  if (currentSheetId !== 'all' && currentSheetId) {
    const [y, m] = currentSheetId.split('-');
    defaultDate = new Date(y, parseInt(m)-1, 15); // middle of month
  }
  const dateStr = defaultDate.toISOString().split('T')[0];
  dom.fLastWork.value = dateStr;
  dom.fTerm.value = dateStr;
  
  clearFormErrors();
  openModal(dom.modalOverlay);
  dom.fCode.focus();
}

function openEditModal(id) {
  const emp = data.employees.find((e) => e.id === id);
  if (!emp) return;
  editingId = id;
  $('#modal-title').textContent = 'แก้ไขข้อมูลพนักงาน';
  
  dom.formId.value = emp.id;
  dom.fCode.value = emp.employeeCode || '';
  dom.fName.value = emp.fullName || '';
  dom.fPos.value = emp.position || '';
  dom.fDept.value = emp.department || '';
  dom.fEmail.value = emp.email || '';
  dom.fLastWork.value = emp.lastWorkDate || '';
  dom.fTerm.value = emp.terminationDate || '';
  dom.fResignStatus.value = emp.resignStatus || '';
  dom.fDocStatus.value = emp.documentStatus || '';
  dom.fHrStatus.value = emp.hrOnlineStatus || '';
  dom.fFactor.value = emp.resignFactor || '';
  dom.fReason.value = emp.resignReason || '';
  dom.fDelEmail.checked = emp.emailDeleted || false;
  dom.fDelRem.checked = emp.remDeleted || false;
  dom.fDelCms.checked = emp.cmsDeleted || false;

  clearFormErrors();
  openModal(dom.modalOverlay);
}

function openDeleteModal(id) {
  const emp = data.employees.find((e) => e.id === id);
  if (!emp) return;
  deletingId = id;
  $('#delete-message').textContent = 'คุณต้องการลบข้อมูลพนักงาน';
  $('#delete-name').textContent = `${emp.employeeCode} - ${emp.fullName}`;
  openModal(dom.deleteOverlay);
}

function openDetailModal(id) {
  const emp = data.employees.find(e => e.id === id);
  if (!emp) return;

  const content = $('#detail-content');
  content.innerHTML = `
    <div class="detail-grid">
      <div>
        <h3 class="section-title">ข้อมูลพนักงาน</h3>
        <div class="detail-item"><div class="detail-label">รหัสพนักงาน</div><div class="detail-value">${escapeHtml(emp.employeeCode)}</div></div>
        <div class="detail-item"><div class="detail-label">ชื่อ - สกุล</div><div class="detail-value">${escapeHtml(emp.fullName)}</div></div>
        <div class="detail-item"><div class="detail-label">ตำแหน่ง</div><div class="detail-value">${escapeHtml(emp.position || '-')}</div></div>
        <div class="detail-item"><div class="detail-label">แผนก</div><div class="detail-value">${escapeHtml(emp.department || '-')}</div></div>
        <div class="detail-item"><div class="detail-label">E-Mail</div><div class="detail-value">${escapeHtml(emp.email || '-')}</div></div>
      </div>
      <div>
        <h3 class="section-title">วันที่ & สถานะ</h3>
        <div class="detail-item"><div class="detail-label">ทำงานวันสุดท้าย</div><div class="detail-value">${formatDate(emp.lastWorkDate)}</div></div>
        <div class="detail-item"><div class="detail-label">วันที่พ้นสภาพ</div><div class="detail-value">${formatDate(emp.terminationDate)}</div></div>
        <div class="detail-item"><div class="detail-label">สถานะลาออก</div><div class="detail-value"><span class="status-badge ${getBadgeClass(emp.resignStatus)}">${emp.resignStatus || '-'}</span></div></div>
        <div class="detail-item"><div class="detail-label">สถานะเอกสาร</div><div class="detail-value">${emp.documentStatus || '-'}</div></div>
        <div class="detail-item"><div class="detail-label">HR Online</div><div class="detail-value">${emp.hrOnlineStatus || '-'}</div></div>
      </div>
      <div style="grid-column: 1 / -1;">
        <h3 class="section-title">ปัจจัยและระบบ</h3>
        <div class="detail-item"><div class="detail-label">ปัจจัยการลาออก</div><div class="detail-value">${emp.resignFactor || '-'}</div></div>
        <div class="detail-item"><div class="detail-label">เหตุผลเพิ่มเติม</div><div class="detail-value">${escapeHtml(emp.resignReason || '-')}</div></div>
        <div class="detail-item">
          <div class="detail-label">ลบข้อมูลระบบ</div>
          <div class="detail-value" style="display:flex; gap:10px; margin-top:5px;">
            <span class="tag ${emp.emailDeleted?'tag-success':''}">${emp.emailDeleted?'✓':'✗'} Email</span>
            <span class="tag ${emp.remDeleted?'tag-success':''}">${emp.remDeleted?'✓':'✗'} REM</span>
            <span class="tag ${emp.cmsDeleted?'tag-success':''}">${emp.cmsDeleted?'✓':'✗'} CMS</span>
          </div>
        </div>
      </div>
    </div>
  `;
  openModal(dom.detailOverlay);
}

function clearFormErrors() {
  $$('.form-group input').forEach((el) => el.classList.remove('error'));
}

// ============================================
// TOAST NOTIFICATIONS
// ============================================

function showToast(message, type = 'success') {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;

  const icon = type === 'success' 
    ? `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>`
    : `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>`;

  toast.innerHTML = `${icon}<span>${message}</span>`;
  dom.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.classList.add('toast-out');
    toast.addEventListener('animationend', () => toast.remove());
  }, 3000);
}

// ============================================
// EXPORT CSV
// ============================================

function exportCSV() {
  const list = getFilteredAndSortedEmployees();
  if (list.length === 0) {
    showToast('ไม่มีข้อมูลให้ Export', 'error');
    return;
  }

  const headers = [
    'รหัสพนักงาน', 'ชื่อ - สกุล', 'ตำแหน่ง', 'แผนก', 'E-Mail', 
    'ทำงานวันสุดท้าย', 'วันที่พ้นสภาพ', 'สถานะลาออก', 'สถานะเอกสาร', 
    'HR Online', 'ปัจจัยการลาออก', 'เหตุผล', 'ลบ Email', 'ลบ REM', 'ลบ CMS'
  ];

  const rows = list.map(e => [
    e.employeeCode, e.fullName, e.position, e.department, e.email,
    e.lastWorkDate, e.terminationDate, e.resignStatus, e.documentStatus,
    e.hrOnlineStatus, e.resignFactor, e.resignReason,
    e.emailDeleted ? 'Y' : '', e.remDeleted ? 'Y' : '', e.cmsDeleted ? 'Y' : ''
  ]);

  const BOM = '\uFEFF';
  const csv = BOM + [
    headers.join(','), 
    ...rows.map(r => r.map(c => `"${String(c || '').replace(/"/g, '""')}"`).join(','))
  ].join('\n');

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const fileNameId = currentSheetId === 'all' ? 'All' : currentSheetId;
  a.download = `Sena_Resign_${fileNameId}_${new Date().toISOString().split('T')[0]}.csv`;
  a.click();
  URL.revokeObjectURL(url);

  showToast('Export CSV สำเร็จ');
}

// ============================================
// IMPORT EXCEL
// ============================================

function parseExcelDate(serial) {
  if (!serial) return '';
  if (typeof serial === 'string') {
    // Attempt to parse string dates if present
    const d = new Date(serial);
    if (!isNaN(d)) return d.toISOString().split('T')[0];
    return serial;
  }
  // Excel uses 1900 date system
  const utc_days  = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;                                        
  const date_info = new Date(utc_value * 1000);

  const fractional_day = serial - Math.floor(serial) + 0.0000001;
  let total_seconds = Math.floor(86400 * fractional_day);
  const seconds = total_seconds % 60;
  total_seconds -= seconds;
  const hours = Math.floor(total_seconds / (60 * 60));
  const minutes = Math.floor(total_seconds / 60) % 60;

  date_info.setHours(hours, minutes, seconds);
  
  // Convert to YYYY-MM-DD
  const year = date_info.getFullYear();
  const month = String(date_info.getMonth() + 1).padStart(2, '0');
  const day = String(date_info.getDate()).padStart(2, '0');
  
  if (isNaN(year) || isNaN(date_info.getMonth())) return '';
  return `${year}-${month}-${day}`;
}

function handleImport(file) {
  const overlay = $('#loading-overlay');
  overlay.classList.add('open');

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const dataStr = new Uint8Array(e.target.result);
      const workbook = XLSX.read(dataStr, { type: 'array' });
      
      // Get target sheet or first sheet
      const targetSheetName = "รายชื่อพนักงานลาออกพ้นสภาพ";
      let sheetName = workbook.SheetNames.includes(targetSheetName) ? targetSheetName : workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      // Convert to JSON (array of arrays)
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      if (jsonData.length < 2) {
        showToast('ไม่พบข้อมูลในไฟล์ Excel', 'error');
        overlay.classList.remove('open');
        return;
      }

      // Skip header row (index 0)
      let importedCount = 0;
      let updatedCount = 0;
      const now = new Date().toISOString();

      for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || row.length === 0 || !row[1]) continue; // Skip empty rows or rows without employeeCode (Col B)

        const employeeCode = String(row[1] || '').trim();
        if (!employeeCode) continue;

        const termDateStr = parseExcelDate(row[7]); // Col H
        
        const empData = {
          employeeCode: employeeCode,
          fullName: String(row[2] || '').trim(), // Col C
          position: String(row[3] || '').trim(), // Col D
          department: String(row[4] || '').trim(), // Col E
          email: String(row[5] || '').trim(), // Col F
          lastWorkDate: parseExcelDate(row[6]), // Col G
          terminationDate: termDateStr,
          resignStatus: String(row[8] || '').trim(), // Col I
          documentStatus: String(row[9] || '').trim(), // Col J
          hrOnlineStatus: String(row[10] || '').trim(), // Col K
          resignFactor: String(row[11] || '').trim(), // Col L
          resignReason: String(row[12] || '').trim(), // Col M
          emailDeleted: !!row[13], // Col N
          remDeleted: !!row[14], // Col O
          cmsDeleted: !!row[15], // Col P
        };

        const sheetId = ensureSheetExists(getSheetIdFromDate(termDateStr));

        // Deduplication Check
        const existingIdx = data.employees.findIndex(e => e.employeeCode === employeeCode);
        if (existingIdx !== -1) {
          // Update
          data.employees[existingIdx] = { 
            ...data.employees[existingIdx], 
            ...empData, 
            updatedAt: now 
          };
          updatedCount++;
        } else {
          // Insert
          data.employees.push({ 
            ...empData, 
            id: generateId(), 
            createdAt: now, 
            updatedAt: now 
          });
          importedCount++;
        }
      }

      saveData();
      
      // Auto select the 'all' sheet after import so user sees everything
      currentSheetId = 'all';

      renderAll();
      showToast(`นำเข้าสำเร็จ: เพิ่มใหม่ ${importedCount} รายการ, อัปเดต ${updatedCount} รายการ`);

    } catch (err) {
      console.error(err);
      showToast('เกิดข้อผิดพลาดในการอ่านไฟล์ Excel', 'error');
    } finally {
      overlay.classList.remove('open');
      dom.fileImport.value = ''; // Reset file input
    }
  };
  
  reader.onerror = function() {
    showToast('ไม่สามารถอ่านไฟล์ได้', 'error');
    overlay.classList.remove('open');
    dom.fileImport.value = '';
  };

  reader.readAsArrayBuffer(file);
}

// ============================================
// EVENT HANDLERS
// ============================================

function initEventListeners() {
  // Add / Export / Import
  dom.btnAdd.addEventListener('click', openAddModal);
  dom.btnExport.addEventListener('click', exportCSV);
  
  dom.btnImport.addEventListener('click', () => dom.fileImport.click());
  dom.fileImport.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleImport(file);
  });

  // Add Sheet
  dom.btnAddSheet.addEventListener('click', () => {
    dom.fSheetMonth.value = '';
    openModal(dom.addSheetOverlay);
  });
  
  dom.sheetForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const val = dom.fSheetMonth.value;
    if (val) {
      ensureSheetExists(val);
      currentSheetId = val;
      renderAll();
      closeModal(dom.addSheetOverlay);
      showToast('สร้างชีทสำเร็จ');
    }
  });

  // Main Form Submit
  dom.form.addEventListener('submit', (e) => {
    e.preventDefault();
    
    // Basic validation
    let valid = true;
    clearFormErrors();
    if (!dom.fCode.value.trim()) { dom.fCode.classList.add('error'); valid = false; }
    if (!dom.fName.value.trim()) { dom.fName.classList.add('error'); valid = false; }
    if (!dom.fLastWork.value) { dom.fLastWork.classList.add('error'); valid = false; }
    if (!dom.fTerm.value) { dom.fTerm.classList.add('error'); valid = false; }
    
    if (!valid) {
      showToast('กรุณากรอกข้อมูลบังคับ (*) ให้ครบถ้วน', 'error');
      return;
    }

    const empData = {
      id: dom.formId.value || undefined,
      employeeCode: dom.fCode.value.trim(),
      fullName: dom.fName.value.trim(),
      position: dom.fPos.value.trim(),
      department: dom.fDept.value.trim(),
      email: dom.fEmail.value.trim(),
      lastWorkDate: dom.fLastWork.value,
      terminationDate: dom.fTerm.value,
      resignStatus: dom.fResignStatus.value,
      documentStatus: dom.fDocStatus.value,
      hrOnlineStatus: dom.fHrStatus.value,
      resignFactor: dom.fFactor.value,
      resignReason: dom.fReason.value.trim(),
      emailDeleted: dom.fDelEmail.checked,
      remDeleted: dom.fDelRem.checked,
      cmsDeleted: dom.fDelCms.checked,
    };

    saveEmployee(empData);
    showToast(editingId ? 'แก้ไขข้อมูลสำเร็จ' : 'บันทึกข้อมูลสำเร็จ');
    closeModal(dom.modalOverlay);
    renderAll();
  });

  // Delete confirm
  dom.btnDeleteConfirm.addEventListener('click', () => {
    if (deletingId) {
      deleteEmployee(deletingId);
      showToast('ลบข้อมูลสำเร็จ');
      closeModal(dom.deleteOverlay);
      renderAll();
      deletingId = null;
    }
  });

  // Table actions delegation
  dom.tableBody.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-action]');
    if (btn) {
      e.stopPropagation(); // prevent row click
      const action = btn.dataset.action;
      const id = btn.dataset.id;
      if (action === 'edit') openEditModal(id);
      if (action === 'delete') openDeleteModal(id);
      return;
    }

    // Row click for detail
    const row = e.target.closest('tr[data-action="view"]');
    if (row) {
      openDetailModal(row.dataset.id);
    }
  });

  // Sorting
  dom.sortHeaders.forEach(th => {
    th.addEventListener('click', () => {
      const column = th.dataset.sort;
      if (sortConfig.column === column) {
        sortConfig.direction = sortConfig.direction === 'asc' ? 'desc' : 'asc';
      } else {
        sortConfig.column = column;
        sortConfig.direction = 'asc';
      }
      renderTable(getFilteredAndSortedEmployees());
    });
  });

  // Filters & Search
  let searchTimeout;
  dom.searchInput.addEventListener('input', (e) => {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => {
      currentFilters.search = e.target.value.trim();
      currentPage = 1;
      renderAll();
    }, 300);
  });

  const applyFilter = (e, key) => {
    currentFilters[key] = e.target.value;
    currentPage = 1;
    renderAll();
  };

  dom.filterResign.addEventListener('change', e => applyFilter(e, 'resignStatus'));
  dom.filterDoc.addEventListener('change', e => applyFilter(e, 'docStatus'));
  dom.filterHr.addEventListener('change', e => applyFilter(e, 'hrStatus'));
  dom.filterFactor.addEventListener('change', e => applyFilter(e, 'factor'));
  dom.filterStartDate.addEventListener('change', e => applyFilter(e, 'startDate'));
  dom.filterEndDate.addEventListener('change', e => applyFilter(e, 'endDate'));

  // Modals closing
  const closes = [
    { btn: dom.btnCancel, modal: dom.modalOverlay },
    { btn: dom.btnModalClose, modal: dom.modalOverlay },
    { btn: dom.btnSheetCancel, modal: dom.addSheetOverlay },
    { btn: dom.btnSheetClose, modal: dom.addSheetOverlay },
    { btn: dom.btnDeleteCancel, modal: dom.deleteOverlay },
    { btn: dom.btnDeleteClose, modal: dom.deleteOverlay },
    { btn: dom.btnDetailClose, modal: dom.detailOverlay },
  ];
  
  closes.forEach(({btn, modal}) => {
    if(btn && modal) btn.addEventListener('click', () => closeModal(modal));
  });

  [dom.modalOverlay, dom.addSheetOverlay, dom.deleteOverlay, dom.detailOverlay].forEach(overlay => {
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) closeModal(overlay);
    });
  });
}

// ============================================
// INITIALIZATION
// ============================================

function init() {
  loadData();
  initEventListeners();
  renderAll();
}

init();
