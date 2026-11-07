/**
 * Heavy Equipment Spare Parts Inventory Dashboard
 * Core Application Logic & Google Sheets Synchronization
 */

// Application State
const state = {
  inventory: [],
  filtered: [],
  currentPage: 1,
  pageSize: 25,
  sortColumn: 'no',
  sortDirection: 'asc',
  filters: {
    search: '',
    prod: 'ALL',
    modelUnit: 'ALL',
    stockStatus: 'ALL',
    serviceStatus: 'ALL',
    quickTab: 'ALL'
  },
  config: {
    gasWebAppUrl: '',
    spreadsheetId: '',
    autoSync: false,
    syncIntervalMinutes: 2,
    theme: 'dark'
  },
  syncTimer: null,
  isSyncing: false,
  charts: {
    statusChart: null,
    brandChart: null,
    criticalChart: null,
    serviceChart: null
  }
};

// DOM Ready initialization
document.addEventListener('DOMContentLoaded', () => {
  loadConfig();
  initTheme();
  initInventoryData();
  setupEventListeners();
  renderAll();

  // If Google Apps Script URL is saved, test or auto-sync
  if (state.config.gasWebAppUrl) {
    updateSyncBadge('online', 'Terkoneksi ke Google Sheets');
    if (state.config.autoSync) {
      startAutoSync();
    }
  } else {
    updateSyncBadge('offline', 'Mode Offline / Lokal');
  }
});

/**
 * Load Configuration from LocalStorage
 */
function loadConfig() {
  try {
    const saved = localStorage.getItem('he_inventory_config');
    if (saved) {
      state.config = { ...state.config, ...JSON.parse(saved) };
    }
  } catch (e) {
    console.error('Error loading config', e);
  }
}

/**
 * Save Configuration to LocalStorage
 */
function saveConfig() {
  try {
    localStorage.setItem('he_inventory_config', JSON.stringify(state.config));
  } catch (e) {
    console.error('Error saving config', e);
  }
}

/**
 * Initialize Inventory Data:
 * Check localStorage first, otherwise fallback to INITIAL_INVENTORY_DATA from data.js
 */
function initInventoryData() {
  try {
    const cachedData = localStorage.getItem('he_inventory_data');
    if (cachedData) {
      const parsed = JSON.parse(cachedData);
      if (Array.isArray(parsed) && parsed.length > 0) {
        state.inventory = parsed;
        return;
      }
    }
  } catch (e) {
    console.warn('Could not read cached inventory, using default data', e);
  }

  // Fallback to data.js
  if (typeof INITIAL_INVENTORY_DATA !== 'undefined' && Array.isArray(INITIAL_INVENTORY_DATA)) {
    state.inventory = JSON.parse(JSON.stringify(INITIAL_INVENTORY_DATA));
    saveInventoryToLocal();
  } else {
    state.inventory = [];
  }
}

/**
 * Save Inventory to LocalStorage
 */
function saveInventoryToLocal() {
  try {
    localStorage.setItem('he_inventory_data', JSON.stringify(state.inventory));
  } catch (e) {
    console.error('Error saving inventory to localStorage', e);
  }
}

/**
 * Setup Theme (Dark / Light)
 */
function initTheme() {
  const theme = state.config.theme || 'dark';
  document.documentElement.setAttribute('data-theme', theme);
  const themeBtn = document.getElementById('themeToggleBtn');
  if (themeBtn) {
    themeBtn.innerHTML = theme === 'dark' ? '<i class="fa-solid fa-sun"></i>' : '<i class="fa-solid fa-moon"></i>';
  }
}

function toggleTheme() {
  const current = document.documentElement.getAttribute('data-theme') || 'dark';
  const newTheme = current === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', newTheme);
  state.config.theme = newTheme;
  saveConfig();

  const themeBtn = document.getElementById('themeToggleBtn');
  if (themeBtn) {
    themeBtn.innerHTML = newTheme === 'dark' ? '<i class="fa-solid fa-sun"></i>' : '<i class="fa-solid fa-moon"></i>';
  }

  // Update chart themes
  updateCharts();
}

/**
 * Stock Status Assessment
 */
function getStockStatus(item) {
  const qty = Number(item.qty) || 0;
  const min = item.min !== null && item.min !== undefined && item.min !== '' ? Number(item.min) : null;
  const max = item.max !== null && item.max !== undefined && item.max !== '' ? Number(item.max) : null;

  if (qty === 0) {
    return { code: 'OUT_OF_STOCK', label: 'Habis (Out of Stock)', class: 'badge-outstock', color: '#ef4444' };
  }
  if (min !== null && qty < min) {
    return { code: 'UNDERSTOCK', label: 'Understock (Kritis)', class: 'badge-understock', color: '#f59e0b' };
  }
  if (max !== null && qty > max) {
    return { code: 'OVERSTOCK', label: 'Overstock (Surplus)', class: 'badge-overstock', color: '#06b6d4' };
  }
  return { code: 'OPTIMAL', label: 'Optimal / Cukup', class: 'badge-optimal', color: '#10b981' };
}

/**
 * Setup Event Listeners
 */
function setupEventListeners() {
  // Theme Toggle
  const themeBtn = document.getElementById('themeToggleBtn');
  if (themeBtn) themeBtn.addEventListener('click', toggleTheme);

  // Search input
  const searchInput = document.getElementById('searchInput');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      state.filters.search = e.target.value.trim().toLowerCase();
      state.currentPage = 1;
      applyFilters();
    });
  }

  // Quick Filter Tabs
  document.querySelectorAll('.tab-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      document.querySelectorAll('.tab-pill').forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      state.filters.quickTab = pill.dataset.filter || 'ALL';
      state.currentPage = 1;
      applyFilters();
    });
  });

  // Dropdown Filters
  const filterProd = document.getElementById('filterProd');
  if (filterProd) {
    filterProd.addEventListener('change', (e) => {
      state.filters.prod = e.target.value;
      state.currentPage = 1;
      applyFilters();
    });
  }

  const filterModel = document.getElementById('filterModel');
  if (filterModel) {
    filterModel.addEventListener('change', (e) => {
      state.filters.modelUnit = e.target.value;
      state.currentPage = 1;
      applyFilters();
    });
  }

  const filterStock = document.getElementById('filterStock');
  if (filterStock) {
    filterStock.addEventListener('change', (e) => {
      state.filters.stockStatus = e.target.value;
      state.currentPage = 1;
      applyFilters();
    });
  }

  const filterService = document.getElementById('filterService');
  if (filterService) {
    filterService.addEventListener('change', (e) => {
      state.filters.serviceStatus = e.target.value;
      state.currentPage = 1;
      applyFilters();
    });
  }

  const pageSizeSelect = document.getElementById('pageSizeSelect');
  if (pageSizeSelect) {
    pageSizeSelect.addEventListener('change', (e) => {
      state.pageSize = e.target.value === 'ALL' ? 999999 : Number(e.target.value);
      state.currentPage = 1;
      renderTable();
    });
  }

  // Action Buttons
  const btnSyncNow = document.getElementById('btnSyncNow');
  if (btnSyncNow) btnSyncNow.addEventListener('click', syncFromGoogleSheets);

  const btnOpenSettings = document.getElementById('btnOpenSettings');
  if (btnOpenSettings) btnOpenSettings.addEventListener('click', openSettingsModal);

  const btnAddNew = document.getElementById('btnAddNew');
  if (btnAddNew) btnAddNew.addEventListener('click', openAddModal);

  const btnExportCsv = document.getElementById('btnExportCsv');
  if (btnExportCsv) btnExportCsv.addEventListener('click', exportToCsv);

  const btnExportJson = document.getElementById('btnExportJson');
  if (btnExportJson) btnExportJson.addEventListener('click', exportToJson);

  const btnPrint = document.getElementById('btnPrint');
  if (btnPrint) btnPrint.addEventListener('click', () => window.print());

  // Modal Closers
  document.querySelectorAll('.modal-close, .btn-modal-cancel').forEach(btn => {
    btn.addEventListener('click', closeAllModals);
  });

  // Close modal when clicking outside
  document.querySelectorAll('.modal-overlay').forEach(overlay => {
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) closeAllModals();
    });
  });

  // Table header sorting
  document.querySelectorAll('.inventory-table th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.sort;
      if (state.sortColumn === col) {
        state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc';
      } else {
        state.sortColumn = col;
        state.sortDirection = 'asc';
      }
      sortAndRenderTable();
    });
  });

  // Part Form Submit (Add/Edit)
  const partForm = document.getElementById('partForm');
  if (partForm) {
    partForm.addEventListener('submit', handlePartFormSubmit);
  }

  // Settings Form Submit
  const settingsForm = document.getElementById('settingsForm');
  if (settingsForm) {
    settingsForm.addEventListener('submit', handleSettingsSubmit);
  }

  // Test GAS Connection Button
  const btnTestGas = document.getElementById('btnTestGas');
  if (btnTestGas) {
    btnTestGas.addEventListener('click', testGoogleAppsScript);
  }

  // Bulk Push to Google Sheets Button
  const btnBulkPush = document.getElementById('btnBulkPush');
  if (btnBulkPush) {
    btnBulkPush.addEventListener('click', bulkPushToGoogleSheets);
  }

  // Copy Code Button
  const btnCopyCode = document.getElementById('btnCopyCode');
  if (btnCopyCode) {
    btnCopyCode.addEventListener('click', copyGasScript);
  }
}

/**
 * Filter and Sort Data
 */
function applyFilters() {
  const { search, prod, modelUnit, stockStatus, serviceStatus, quickTab } = state.filters;

  state.filtered = state.inventory.filter(item => {
    // Search
    if (search) {
      const matchPn = (item.pnFinal || '').toLowerCase().includes(search);
      const matchAlt = (item.alternatePn || '').toLowerCase().includes(search);
      const matchDesc = (item.desc || '').toLowerCase().includes(search);
      const matchModel = (item.modelUnit || '').toLowerCase().includes(search);
      if (!matchPn && !matchAlt && !matchDesc && !matchModel) return false;
    }

    // Producer / Brand
    if (prod !== 'ALL' && item.prod !== prod) return false;

    // Model Unit
    if (modelUnit !== 'ALL' && item.modelUnit !== modelUnit) return false;

    // Service Status (PS vs Non-PS)
    if (serviceStatus !== 'ALL') {
      const isPs = (item.status || '').toUpperCase() === 'PS';
      if (serviceStatus === 'PS' && !isPs) return false;
      if (serviceStatus === 'NON_PS' && isPs) return false;
    }

    // Stock Status
    const statusObj = getStockStatus(item);
    if (stockStatus !== 'ALL' && statusObj.code !== stockStatus) return false;

    // Quick Tab filter
    if (quickTab === 'KRITIS' && statusObj.code !== 'OUT_OF_STOCK' && statusObj.code !== 'UNDERSTOCK') return false;
    if (quickTab === 'HABIS' && statusObj.code !== 'OUT_OF_STOCK') return false;
    if (quickTab === 'UNDERSTOCK' && statusObj.code !== 'UNDERSTOCK') return false;
    if (quickTab === 'OPTIMAL' && statusObj.code !== 'OPTIMAL') return false;
    if (quickTab === 'OVERSTOCK' && statusObj.code !== 'OVERSTOCK') return false;
    if (quickTab === 'PS' && (item.status || '').toUpperCase() !== 'PS') return false;

    return true;
  });

  sortAndRenderTable();
}

/**
 * Sort and Render Table
 */
function sortAndRenderTable() {
  const col = state.sortColumn;
  const dir = state.sortDirection === 'asc' ? 1 : -1;

  state.filtered.sort((a, b) => {
    let valA = a[col];
    let valB = b[col];

    if (col === 'qty' || col === 'min' || col === 'max' || col === 'no' || col === 'doi') {
      valA = Number(valA) || 0;
      valB = Number(valB) || 0;
      return (valA - valB) * dir;
    }

    valA = (valA || '').toString().toLowerCase();
    valB = (valB || '').toString().toLowerCase();
    if (valA < valB) return -1 * dir;
    if (valA > valB) return 1 * dir;
    return 0;
  });

  renderTable();
  updateSortIndicators();
}

/**
 * Update Sort Indicators on headers
 */
function updateSortIndicators() {
  document.querySelectorAll('.inventory-table th[data-sort]').forEach(th => {
    const col = th.dataset.sort;
    const icon = th.querySelector('i');
    if (!icon) return;

    if (state.sortColumn === col) {
      icon.className = state.sortDirection === 'asc' ? 'fa-solid fa-sort-up' : 'fa-solid fa-sort-down';
    } else {
      icon.className = 'fa-solid fa-sort';
    }
  });
}

/**
 * Populate Dropdowns (Prod & Model Unit)
 */
function populateFilterDropdowns() {
  const prodSelect = document.getElementById('filterProd');
  const modelSelect = document.getElementById('filterModel');

  if (prodSelect) {
    const prods = [...new Set(state.inventory.map(i => i.prod).filter(Boolean))].sort();
    const currentVal = prodSelect.value;
    prodSelect.innerHTML = '<option value="ALL">Semua Produsen</option>' +
      prods.map(p => `<option value="${escapeHtml(p)}" ${currentVal === p ? 'selected' : ''}>${escapeHtml(p)}</option>`).join('');
  }

  if (modelSelect) {
    const models = [...new Set(state.inventory.map(i => i.modelUnit).filter(Boolean))].sort();
    const currentVal = modelSelect.value;
    modelSelect.innerHTML = '<option value="ALL">Semua Model Unit</option>' +
      models.map(m => `<option value="${escapeHtml(m)}" ${currentVal === m ? 'selected' : ''}>${escapeHtml(m)}</option>`).join('');
  }
}

/**
 * Render All (KPIs, Charts, Dropdowns, Table)
 */
function renderAll() {
  populateFilterDropdowns();
  updateKPIs();
  updateCharts();
  applyFilters();
}

/**
 * Update KPI Statistics
 */
function updateKPIs() {
  const totalItems = state.inventory.length;
  let totalStock = 0;
  let outOfStock = 0;
  let understock = 0;
  let optimal = 0;
  let overstock = 0;

  state.inventory.forEach(item => {
    const qty = Number(item.qty) || 0;
    totalStock += qty;
    const status = getStockStatus(item);
    if (status.code === 'OUT_OF_STOCK') outOfStock++;
    else if (status.code === 'UNDERSTOCK') understock++;
    else if (status.code === 'OPTIMAL') optimal++;
    else if (status.code === 'OVERSTOCK') overstock++;
  });

  const brands = new Set(state.inventory.map(i => i.prod).filter(Boolean)).size;

  setText('kpiTotalItems', totalItems.toLocaleString());
  setText('kpiTotalStock', totalStock.toLocaleString());
  setText('kpiOutOfStock', outOfStock.toLocaleString());
  setText('kpiUnderstock', understock.toLocaleString());
  setText('kpiOptimal', optimal.toLocaleString());
  setText('kpiOverstock', overstock.toLocaleString());
  setText('kpiBrands', brands.toString());

  // Update quick tab counts
  setText('countAll', totalItems.toString());
  setText('countKritis', (outOfStock + understock).toString());
  setText('countHabis', outOfStock.toString());
  setText('countUnderstock', understock.toString());
  setText('countOptimal', optimal.toString());
  setText('countOverstock', overstock.toString());
}

/**
 * Render Inventory Table with Pagination
 */
function renderTable() {
  const tbody = document.getElementById('inventoryTableBody');
  if (!tbody) return;

  const total = state.filtered.length;
  const totalPages = Math.ceil(total / state.pageSize) || 1;

  if (state.currentPage > totalPages) state.currentPage = totalPages;
  if (state.currentPage < 1) state.currentPage = 1;

  const start = (state.currentPage - 1) * state.pageSize;
  const end = Math.min(start + state.pageSize, total);
  const pagedItems = state.filtered.slice(start, end);

  if (pagedItems.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="9" style="text-align: center; padding: 3rem 1rem; color: var(--text-muted);">
          <i class="fa-solid fa-box-open" style="font-size: 2.5rem; margin-bottom: 0.75rem; display: block; opacity: 0.5;"></i>
          Tidak ada data spare part yang cocok dengan filter.
        </td>
      </tr>
    `;
    renderPagination(0, 0, 0, 1);
    return;
  }

  tbody.innerHTML = pagedItems.map((item, index) => {
    const status = getStockStatus(item);
    const min = item.min !== null && item.min !== undefined && item.min !== '' ? Number(item.min) : 0;
    const max = item.max !== null && item.max !== undefined && item.max !== '' ? Number(item.max) : (min * 3 || 10);
    const qty = Number(item.qty) || 0;

    // Health percentage calculation
    let healthPct = 0;
    if (max > 0) {
      healthPct = Math.min(Math.round((qty / max) * 100), 100);
    } else if (qty > 0) {
      healthPct = 100;
    }

    const isPs = (item.status || '').toUpperCase() === 'PS';

    return `
      <tr>
        <td style="color: var(--text-muted); font-weight: 500;">${item.no || (start + index + 1)}</td>
        <td>
          <span class="badge badge-prod">${escapeHtml(item.prod || '-')}</span>
        </td>
        <td>
          <span class="badge badge-model">${escapeHtml(item.modelUnit || '-')}</span>
        </td>
        <td>
          <div class="pn-cell">
            <span class="pn-final" title="Part Number">${escapeHtml(item.pnFinal || '-')}</span>
            ${item.alternatePn ? `<span class="pn-alt" title="Alternate Part Number"><i class="fa-solid fa-code-fork"></i> ${escapeHtml(item.alternatePn)}</span>` : ''}
          </div>
        </td>
        <td style="max-width: 320px;">
          <div style="font-weight: 600; color: var(--text-primary); margin-bottom: 2px;">${escapeHtml(item.desc || '-')}</div>
          ${item.remarks ? `<div style="font-size: 0.75rem; color: var(--text-secondary);"><i class="fa-regular fa-note-sticky"></i> ${escapeHtml(item.remarks)}</div>` : ''}
        </td>
        <td>
          <span class="badge ${isPs ? 'badge-ps' : 'badge-non-ps'}">${isPs ? 'PS' : 'NON-PS'}</span>
        </td>
        <td>
          <div class="stock-cell">
            <div class="stock-value-group">
              <span class="stock-num" style="color: ${status.color};">${qty}</span>
              <span class="stock-thresholds">Min: ${item.min ?? '-'} | Max: ${item.max ?? '-'}</span>
            </div>
            <div class="health-bar" title="Kesehatan Stok: ${healthPct}%">
              <div class="health-fill" style="width: ${healthPct}%; background-color: ${status.color};"></div>
            </div>
          </div>
        </td>
        <td>
          <span class="badge ${status.class}">${status.label}</span>
          ${item.doi ? `<div style="font-size: 0.72rem; color: var(--text-muted); margin-top: 4px;">DOI: ${item.doi} hari</div>` : ''}
        </td>
        <td>
          <div class="row-actions">
            <button class="btn-action btn-minus" title="Kurangi Stok (-1)" onclick="quickAdjustStock('${escapeQuote(item.pnFinal)}', -1)">
              <i class="fa-solid fa-minus"></i>
            </button>
            <button class="btn-action btn-plus" title="Tambah Stok (+1)" onclick="quickAdjustStock('${escapeQuote(item.pnFinal)}', 1)">
              <i class="fa-solid fa-plus"></i>
            </button>
            <button class="btn-action" title="Edit Detail" onclick="openEditModal('${escapeQuote(item.pnFinal)}')">
              <i class="fa-regular fa-pen-to-square"></i>
            </button>
            <button class="btn-action btn-delete" title="Hapus Part" onclick="confirmDeleteItem('${escapeQuote(item.pnFinal)}')">
              <i class="fa-regular fa-trash-can"></i>
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join('');

  renderPagination(start + 1, end, total, totalPages);
}

/**
 * Render Pagination Controls
 */
function renderPagination(start, end, total, totalPages) {
  const info = document.getElementById('paginationInfo');
  if (info) {
    info.textContent = total > 0 
      ? `Menampilkan ${start} - ${end} dari ${total} spare part`
      : '0 spare part';
  }

  const container = document.getElementById('paginationControls');
  if (!container) return;

  if (totalPages <= 1) {
    container.innerHTML = '';
    return;
  }

  let html = `
    <button class="page-btn" ${state.currentPage === 1 ? 'disabled' : ''} onclick="changePage(${state.currentPage - 1})">
      <i class="fa-solid fa-chevron-left"></i>
    </button>
  `;

  // Dynamic page numbering
  const maxButtons = 5;
  let startPage = Math.max(1, state.currentPage - 2);
  let endPage = Math.min(totalPages, startPage + maxButtons - 1);

  if (endPage - startPage < maxButtons - 1) {
    startPage = Math.max(1, endPage - maxButtons + 1);
  }

  if (startPage > 1) {
    html += `<button class="page-btn" onclick="changePage(1)">1</button>`;
    if (startPage > 2) html += `<span style="padding: 0 4px; color: var(--text-muted);">...</span>`;
  }

  for (let p = startPage; p <= endPage; p++) {
    html += `<button class="page-btn ${p === state.currentPage ? 'active' : ''}" onclick="changePage(${p})">${p}</button>`;
  }

  if (endPage < totalPages) {
    if (endPage < totalPages - 1) html += `<span style="padding: 0 4px; color: var(--text-muted);">...</span>`;
    html += `<button class="page-btn" onclick="changePage(${totalPages})">${totalPages}</button>`;
  }

  html += `
    <button class="page-btn" ${state.currentPage === totalPages ? 'disabled' : ''} onclick="changePage(${state.currentPage + 1})">
      <i class="fa-solid fa-chevron-right"></i>
    </button>
  `;

  container.innerHTML = html;
}

function changePage(page) {
  state.currentPage = page;
  renderTable();
  const tableEl = document.querySelector('.table-panel');
  if (tableEl) tableEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/**
 * Quick Adjust Stock (+1 or -1 or custom)
 */
async function quickAdjustStock(pn, delta) {
  const item = state.inventory.find(i => i.pnFinal === pn);
  if (!item) return;

  const currentQty = Number(item.qty) || 0;
  const newQty = Math.max(0, currentQty + delta);

  if (currentQty === newQty) return;

  item.qty = newQty;
  saveInventoryToLocal();
  updateKPIs();
  updateCharts();
  renderTable();

  showToast(`Stok ${pn} diubah menjadi ${newQty}`, 'info');

  // Push update to Google Sheets if connected
  if (state.config.gasWebAppUrl) {
    try {
      updateSyncBadge('syncing', 'Menyimpan ke Google Sheets...');
      const res = await fetch(state.config.gasWebAppUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'updateStock',
          pnFinal: pn,
          qty: newQty,
          note: `Quick Adjust ${delta > 0 ? '+' : ''}${delta}`
        })
      });
      const json = await res.json();
      if (json.status === 'success') {
        updateSyncBadge('online', 'Tersinkron ke Google Sheets');
      } else {
        showToast('Gagal update di Google Sheets: ' + json.message, 'warning');
        updateSyncBadge('online', 'Terkoneksi');
      }
    } catch (err) {
      console.error('Error updating stock to GAS:', err);
      updateSyncBadge('offline', 'Koneksi Sheets Terputus');
    }
  }
}

/**
 * Open Modal for Adding New Item
 */
function openAddModal() {
  document.getElementById('partModalTitle').innerHTML = '<i class="fa-solid fa-plus-circle"></i> Tambah Spare Part Baru';
  document.getElementById('partForm').reset();
  document.getElementById('editOriginalPn').value = '';
  document.getElementById('formNo').value = state.inventory.length + 1;
  openModal('partModal');
}

/**
 * Open Modal for Editing Item
 */
function openEditModal(pn) {
  const item = state.inventory.find(i => i.pnFinal === pn);
  if (!item) return;

  document.getElementById('partModalTitle').innerHTML = '<i class="fa-solid fa-pen-to-square"></i> Edit Spare Part';
  document.getElementById('editOriginalPn').value = item.pnFinal;
  document.getElementById('formNo').value = item.no || '';
  document.getElementById('formProd').value = item.prod || '';
  document.getElementById('formModelUnit').value = item.modelUnit || '';
  document.getElementById('formPnFinal').value = item.pnFinal || '';
  document.getElementById('formDesc').value = item.desc || '';
  document.getElementById('formMin').value = item.min !== null ? item.min : '';
  document.getElementById('formMax').value = item.max !== null ? item.max : '';
  document.getElementById('formStatus').value = item.status || 'PS';
  document.getElementById('formQty').value = item.qty !== null ? item.qty : 0;
  document.getElementById('formDoi').value = item.doi !== null ? item.doi : '';
  document.getElementById('formAltPn').value = item.alternatePn || '';
  document.getElementById('formRemarks').value = item.remarks || '';

  openModal('partModal');
}

/**
 * Handle Part Form Submit (Add or Edit)
 */
async function handlePartFormSubmit(e) {
  e.preventDefault();

  const originalPn = document.getElementById('editOriginalPn').value;
  const isEditing = Boolean(originalPn);

  const itemData = {
    no: Number(document.getElementById('formNo').value) || (state.inventory.length + 1),
    prod: document.getElementById('formProd').value.trim(),
    modelUnit: document.getElementById('formModelUnit').value.trim(),
    pnFinal: document.getElementById('formPnFinal').value.trim(),
    desc: document.getElementById('formDesc').value.trim(),
    min: document.getElementById('formMin').value !== '' ? Number(document.getElementById('formMin').value) : null,
    max: document.getElementById('formMax').value !== '' ? Number(document.getElementById('formMax').value) : null,
    status: document.getElementById('formStatus').value,
    qty: Number(document.getElementById('formQty').value) || 0,
    doi: document.getElementById('formDoi').value !== '' ? Number(document.getElementById('formDoi').value) : null,
    alternatePn: document.getElementById('formAltPn').value.trim(),
    remarks: document.getElementById('formRemarks').value.trim()
  };

  if (!itemData.pnFinal || !itemData.desc) {
    showToast('Part Number dan Deskripsi wajib diisi!', 'warning');
    return;
  }

  if (isEditing) {
    const index = state.inventory.findIndex(i => i.pnFinal === originalPn);
    if (index !== -1) {
      state.inventory[index] = { ...state.inventory[index], ...itemData };
      showToast(`Part ${itemData.pnFinal} berhasil diperbarui!`, 'success');
    }
  } else {
    // Check duplicate
    if (state.inventory.some(i => i.pnFinal.toLowerCase() === itemData.pnFinal.toLowerCase())) {
      showToast(`Part Number ${itemData.pnFinal} sudah ada!`, 'warning');
      return;
    }
    state.inventory.unshift(itemData);
    showToast(`Part ${itemData.pnFinal} berhasil ditambahkan!`, 'success');
  }

  saveInventoryToLocal();
  closeAllModals();
  renderAll();

  // Push to Google Sheets if connected
  if (state.config.gasWebAppUrl) {
    try {
      updateSyncBadge('syncing', 'Menyimpan ke Google Sheets...');
      const action = isEditing ? 'editItem' : 'addItem';
      const res = await fetch(state.config.gasWebAppUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: action,
          originalPn: originalPn,
          item: itemData
        })
      });
      const json = await res.json();
      if (json.status === 'success') {
        updateSyncBadge('online', 'Tersinkron ke Google Sheets');
      } else {
        showToast('Gagal sinkron ke Sheets: ' + json.message, 'warning');
      }
    } catch (err) {
      console.error('Error syncing part change to GAS', err);
      updateSyncBadge('offline', 'Koneksi Sheets Terputus');
    }
  }
}

/**
 * Confirm and Delete Item
 */
async function confirmDeleteItem(pn) {
  if (!confirm(`Apakah Anda yakin ingin menghapus part "${pn}"?`)) return;

  state.inventory = state.inventory.filter(i => i.pnFinal !== pn);
  saveInventoryToLocal();
  renderAll();
  showToast(`Part ${pn} telah dihapus.`, 'info');

  if (state.config.gasWebAppUrl) {
    try {
      updateSyncBadge('syncing', 'Menghapus dari Google Sheets...');
      const res = await fetch(state.config.gasWebAppUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'deleteItem', pnFinal: pn })
      });
      const json = await res.json();
      if (json.status === 'success') {
        updateSyncBadge('online', 'Tersinkron ke Google Sheets');
      }
    } catch (err) {
      console.error('Error deleting item in GAS', err);
    }
  }
}

/**
 * Sync Data FROM Google Sheets (Pull)
 */
async function syncFromGoogleSheets() {
  if (!state.config.gasWebAppUrl) {
    openSettingsModal();
    showToast('Silakan masukkan Google Apps Script Web App URL terlebih dahulu.', 'warning');
    return;
  }

  if (state.isSyncing) return;
  state.isSyncing = true;
  updateSyncBadge('syncing', 'Menarik data Google Sheets...');

  try {
    const url = `${state.config.gasWebAppUrl}?action=getData&_t=${Date.now()}`;
    const res = await fetch(url);
    const json = await res.json();

    if (json.status === 'success' && Array.isArray(json.data) && json.data.length > 0) {
      state.inventory = json.data;
      saveInventoryToLocal();
      renderAll();
      showToast(`Berhasil menarik ${json.data.length} item dari Google Sheets!`, 'success');
      updateSyncBadge('online', `Terakhir sync: ${new Date().toLocaleTimeString('id-ID')}`);
    } else if (json.status === 'success' && json.count === 0) {
      showToast('Sheet terhubung tapi masih kosong. Klik "Kirim Seluruh Data ke Sheets" untuk mengisi data awal.', 'info');
      updateSyncBadge('online', 'Terkoneksi (Sheet Kosong)');
    } else {
      showToast('Gagal memuat data dari Google Sheets: ' + (json.message || 'Format tidak valid'), 'error');
      updateSyncBadge('offline', 'Error Sinkronisasi');
    }
  } catch (err) {
    console.error('Error fetching from GAS:', err);
    showToast('Koneksi ke Google Sheets gagal. Pastikan deployment Web App diset ke "Anyone" (Siapa saja).', 'error');
    updateSyncBadge('offline', 'Koneksi Gagal');
  } finally {
    state.isSyncing = false;
  }
}

/**
 * Bulk Push seluruh data lokal ke Google Sheets
 */
async function bulkPushToGoogleSheets() {
  if (!state.config.gasWebAppUrl) {
    showToast('Isi Web App URL terlebih dahulu!', 'warning');
    return;
  }

  if (!confirm(`Kirim seluruh ${state.inventory.length} item ke Google Sheets? Data di sheet saat ini akan diperbarui.`)) {
    return;
  }

  const btn = document.getElementById('btnBulkPush');
  if (btn) btn.disabled = true;
  updateSyncBadge('syncing', 'Mengunggah seluruh data ke Google Sheets...');

  try {
    const res = await fetch(state.config.gasWebAppUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        action: 'bulkSync',
        items: state.inventory
      })
    });
    const json = await res.json();
    if (json.status === 'success') {
      showToast(`Sukses! ${json.message}`, 'success');
      updateSyncBadge('online', 'Tersinkron');
    } else {
      showToast(`Gagal: ${json.message}`, 'error');
    }
  } catch (err) {
    console.error('Error bulk syncing to GAS:', err);
    showToast('Gagal mengunggah ke Google Sheets: ' + err.message, 'error');
    updateSyncBadge('offline', 'Koneksi Gagal');
  } finally {
    if (btn) btn.disabled = false;
  }
}

/**
 * Test GAS Connection
 */
async function testGoogleAppsScript() {
  const urlInput = document.getElementById('cfgGasUrl').value.trim();
  if (!urlInput) {
    showToast('Masukkan Web App URL terlebih dahulu.', 'warning');
    return;
  }

  const btn = document.getElementById('btnTestGas');
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Menghubungkan...';
  }

  try {
    const testUrl = `${urlInput}?action=ping&_t=${Date.now()}`;
    const res = await fetch(testUrl);
    const json = await res.json();

    if (json.status === 'success') {
      showToast(`Koneksi Berhasil! Terhubung ke spreadsheet: "${json.spreadsheetTitle || 'Spreadsheet'}"`, 'success');
      updateSyncBadge('online', 'Terkoneksi ke Google Sheets');
    } else {
      showToast('Koneksi ditolak: ' + json.message, 'error');
    }
  } catch (err) {
    console.error('Ping test error:', err);
    showToast('Gagal terhubung. Pastikan Web App disetting "Execute as: Me" dan "Who has access: Anyone".', 'error');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<i class="fa-solid fa-plug"></i> Test Koneksi';
    }
  }
}

/**
 * Auto Sync Management
 */
function startAutoSync() {
  stopAutoSync();
  const minutes = Math.max(1, state.config.syncIntervalMinutes || 2);
  state.syncTimer = setInterval(() => {
    syncFromGoogleSheets();
  }, minutes * 60 * 1000);
}

function stopAutoSync() {
  if (state.syncTimer) {
    clearInterval(state.syncTimer);
    state.syncTimer = null;
  }
}

/**
 * Settings Modal Form Submit
 */
function handleSettingsSubmit(e) {
  e.preventDefault();

  const url = document.getElementById('cfgGasUrl').value.trim();
  const sheetId = document.getElementById('cfgSheetId').value.trim();
  const autoSync = document.getElementById('cfgAutoSync').checked;
  const interval = Number(document.getElementById('cfgSyncInterval').value) || 2;

  state.config.gasWebAppUrl = url;
  state.config.spreadsheetId = sheetId;
  state.config.autoSync = autoSync;
  state.config.syncIntervalMinutes = interval;

  saveConfig();

  if (autoSync && url) {
    startAutoSync();
  } else {
    stopAutoSync();
  }

  updateEmbedSheet();
  closeAllModals();
  showToast('Pengaturan Google Sheets berhasil disimpan!', 'success');

  if (url) {
    updateSyncBadge('online', 'Terkoneksi ke Google Sheets');
    syncFromGoogleSheets();
  } else {
    updateSyncBadge('offline', 'Mode Offline / Lokal');
  }
}

/**
 * Open Settings Modal
 */
function openSettingsModal() {
  document.getElementById('cfgGasUrl').value = state.config.gasWebAppUrl || '';
  document.getElementById('cfgSheetId').value = state.config.spreadsheetId || '';
  document.getElementById('cfgAutoSync').checked = Boolean(state.config.autoSync);
  document.getElementById('cfgSyncInterval').value = state.config.syncIntervalMinutes || 2;

  const codeEl = document.getElementById('gasCodePreview');
  if (codeEl) {
    codeEl.innerText = GAS_CODE_SNIPPET;
  }

  updateEmbedSheet();
  openModal('settingsModal');
}

/**
 * Update Embedded Google Sheet Preview
 */
function updateEmbedSheet() {
  const container = document.getElementById('embedSheetContainer');
  if (!container) return;

  const sheetId = state.config.spreadsheetId;
  if (sheetId) {
    container.style.display = 'block';
    const iframe = container.querySelector('iframe');
    if (iframe) {
      iframe.src = `https://docs.google.com/spreadsheets/d/${sheetId}/edit?usp=sharing&rm=minimal`;
    }
  } else {
    container.style.display = 'none';
  }
}

/**
 * Copy GAS Script to Clipboard
 */
const GAS_CODE_SNIPPET = `const SHEET_NAME = "INVENTORY";
const HEADERS = ["No", "Prod", "Model Unit", "Part Number", "Description", "Min", "Max", "Status", "Qty", "Month", "Week", "Day", "DOI", "Remarks", "Alternate PN", "Last Updated"];

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getData";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME) || setupSheet();

    if (action === "ping") {
      return jsonResponse({ status: "success", message: "Google Apps Script API is active & connected!", spreadsheetTitle: ss.getName() });
    }
    if (action === "getData") {
      const data = getSheetData(sheet);
      return jsonResponse({ status: "success", count: data.length, data: data });
    }
    return jsonResponse({ status: "error", message: "Unknown action: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    if (!e || !e.postData || !e.postData.contents) return jsonResponse({ status: "error", message: "No post data" });
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME) || setupSheet();

    if (action === "updateStock") {
      const pn = payload.pnFinal;
      const newQty = Number(payload.qty);
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][3]).trim() === String(pn).trim()) {
          sheet.getRange(i + 1, 9).setValue(newQty);
          sheet.getRange(i + 1, 16).setValue(new Date().toLocaleString("id-ID"));
          return jsonResponse({ status: "success", message: "Stok berhasil diperbarui" });
        }
      }
      return jsonResponse({ status: "error", message: "Part tidak ditemukan" });
    }

    if (action === "addItem") {
      const item = payload.item;
      const lastRow = sheet.getLastRow();
      sheet.appendRow([
        item.no || (lastRow), item.prod || "", item.modelUnit || "", item.pnFinal || "", item.desc || "",
        item.min ?? "", item.max ?? "", item.status || "PS", Number(item.qty) || 0,
        item.month ?? "", item.week ?? "", item.day ?? "", item.doi ?? "",
        item.remarks || "", item.alternatePn || "", new Date().toLocaleString("id-ID")
      ]);
      return jsonResponse({ status: "success", message: "Part ditambahkan" });
    }

    if (action === "editItem") {
      const item = payload.item;
      const pn = item.pnFinal;
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][3]).trim() === String(pn).trim() || (item.originalPn && String(rows[i][3]).trim() === String(item.originalPn).trim())) {
          const rowIdx = i + 1;
          sheet.getRange(rowIdx, 2).setValue(item.prod || "");
          sheet.getRange(rowIdx, 3).setValue(item.modelUnit || "");
          sheet.getRange(rowIdx, 4).setValue(item.pnFinal || "");
          sheet.getRange(rowIdx, 5).setValue(item.desc || "");
          sheet.getRange(rowIdx, 6).setValue(item.min ?? "");
          sheet.getRange(rowIdx, 7).setValue(item.max ?? "");
          sheet.getRange(rowIdx, 8).setValue(item.status || "PS");
          sheet.getRange(rowIdx, 9).setValue(Number(item.qty) || 0);
          sheet.getRange(rowIdx, 13).setValue(item.doi ?? "");
          sheet.getRange(rowIdx, 14).setValue(item.remarks || "");
          sheet.getRange(rowIdx, 15).setValue(item.alternatePn || "");
          sheet.getRange(rowIdx, 16).setValue(new Date().toLocaleString("id-ID"));
          return jsonResponse({ status: "success", message: "Part diperbarui" });
        }
      }
      return jsonResponse({ status: "error", message: "Part tidak ditemukan" });
    }

    if (action === "deleteItem") {
      const pn = payload.pnFinal;
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][3]).trim() === String(pn).trim()) {
          sheet.deleteRow(i + 1);
          return jsonResponse({ status: "success", message: "Part dihapus" });
        }
      }
      return jsonResponse({ status: "error", message: "Part tidak ditemukan" });
    }

    if (action === "bulkSync" && Array.isArray(payload.items)) {
      sheet.clearContents();
      sheet.appendRow(HEADERS);
      const items = payload.items;
      const rowsToInsert = items.map((it, idx) => [
        it.no || (idx + 1), it.prod || "", it.modelUnit || "", it.pnFinal || "", it.desc || "",
        it.min ?? "", it.max ?? "", it.status || "PS", Number(it.qty) || 0,
        it.month ?? "", it.week ?? "", it.day ?? "", it.doi ?? "",
        it.remarks || "", it.alternatePn || "", new Date().toLocaleString("id-ID")
      ]);
      if (rowsToInsert.length > 0) {
        sheet.getRange(2, 1, rowsToInsert.length, HEADERS.length).setValues(rowsToInsert);
      }
      formatSheetDesign(sheet);
      return jsonResponse({ status: "success", message: "Bulk sync berhasil: " + items.length + " item" });
    }

    return jsonResponse({ status: "error", message: "Action tidak valid" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function getSheetData(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const items = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[3] && !row[4]) continue;
    items.push({
      no: row[0] ? Number(row[0]) : i,
      prod: String(row[1] || "").trim(),
      modelUnit: String(row[2] || "").trim(),
      pnFinal: String(row[3] || "").trim(),
      desc: String(row[4] || "").trim(),
      min: row[5] !== "" && row[5] !== null ? Number(row[5]) : null,
      max: row[6] !== "" && row[6] !== null ? Number(row[6]) : null,
      status: String(row[7] || "PS").trim(),
      qty: row[8] !== "" && row[8] !== null ? Number(row[8]) : 0,
      month: row[9] !== "" && row[9] !== null ? Number(row[9]) : null,
      week: row[10] !== "" && row[10] !== null ? Number(row[10]) : null,
      day: row[11] !== "" && row[11] !== null ? Number(row[11]) : null,
      doi: row[12] !== "" && row[12] !== null ? Number(row[12]) : null,
      remarks: String(row[13] || "").trim(),
      alternatePn: String(row[14] || "").trim()
    });
  }
  return items;
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  formatSheetDesign(sheet);
  return sheet;
}

function formatSheetDesign(sheet) {
  const header = sheet.getRange(1, 1, 1, HEADERS.length);
  header.setBackground("#1e293b").setFontColor("#f8fafc").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}`;

function copyGasScript() {
  navigator.clipboard.writeText(GAS_CODE_SNIPPET)
    .then(() => showToast('Kode Google Apps Script berhasil disalin ke clipboard!', 'success'))
    .catch(() => showToast('Gagal menyalin kode.', 'error'));
}

/**
 * Update Visual Sync Status Badge
 */
function updateSyncBadge(status, text) {
  const dot = document.getElementById('syncDot');
  const label = document.getElementById('syncLabel');

  if (dot) {
    dot.className = 'sync-dot';
    if (status === 'online') dot.classList.add('online');
    else if (status === 'syncing') dot.classList.add('syncing');
  }

  if (label) {
    label.textContent = text;
  }
}

/**
 * Chart.js Visualizations Setup & Dynamic Updates
 */
function updateCharts() {
  if (typeof Chart === 'undefined') return;

  const isDark = (document.documentElement.getAttribute('data-theme') || 'dark') === 'dark';
  const textColor = isDark ? '#94a3b8' : '#475569';
  const gridColor = isDark ? 'rgba(148, 163, 184, 0.1)' : 'rgba(203, 213, 225, 0.5)';

  // Chart 1: Stock Status Distribution (Doughnut)
  renderStatusChart(textColor);

  // Chart 2: Inventory by Brand / Producer (Polar/Bar)
  renderBrandChart(textColor, gridColor);

  // Chart 3: Critical Low Stock Items (Horizontal Bar)
  renderCriticalChart(textColor, gridColor);

  // Chart 4: Service Distribution (PS vs Non-PS)
  renderServiceChart(textColor);
}

function renderStatusChart(textColor) {
  const canvas = document.getElementById('statusChartCanvas');
  if (!canvas) return;

  let outStock = 0, understock = 0, optimal = 0, overstock = 0;
  state.inventory.forEach(i => {
    const s = getStockStatus(i);
    if (s.code === 'OUT_OF_STOCK') outStock++;
    else if (s.code === 'UNDERSTOCK') understock++;
    else if (s.code === 'OPTIMAL') optimal++;
    else if (s.code === 'OVERSTOCK') overstock++;
  });

  const data = {
    labels: ['Habis (0)', 'Understock (< Min)', 'Optimal', 'Overstock (> Max)'],
    datasets: [{
      data: [outStock, understock, optimal, overstock],
      backgroundColor: ['#ef4444', '#f59e0b', '#10b981', '#06b6d4'],
      borderWidth: 2,
      borderColor: (document.documentElement.getAttribute('data-theme') || 'dark') === 'dark' ? '#1e293b' : '#ffffff'
    }]
  };

  if (state.charts.statusChart) {
    state.charts.statusChart.data = data;
    state.charts.statusChart.options.plugins.legend.labels.color = textColor;
    state.charts.statusChart.update();
  } else {
    state.charts.statusChart = new Chart(canvas, {
      type: 'doughnut',
      data: data,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: textColor, font: { size: 11, family: 'Inter' }, padding: 12 }
          }
        },
        cutout: '68%'
      }
    });
  }
}

function renderBrandChart(textColor, gridColor) {
  const canvas = document.getElementById('brandChartCanvas');
  if (!canvas) return;

  const brandCounts = {};
  state.inventory.forEach(i => {
    const b = i.prod || 'Other';
    brandCounts[b] = (brandCounts[b] || 0) + 1;
  });

  const labels = Object.keys(brandCounts);
  const data = Object.values(brandCounts);

  const chartData = {
    labels: labels,
    datasets: [{
      label: 'Jumlah Part',
      data: data,
      backgroundColor: [
        '#3b82f6', '#8b5cf6', '#10b981', '#f59e0b', '#ec4899', '#06b6d4'
      ],
      borderRadius: 6
    }]
  };

  if (state.charts.brandChart) {
    state.charts.brandChart.data = chartData;
    state.charts.brandChart.options.scales.x.ticks.color = textColor;
    state.charts.brandChart.options.scales.y.ticks.color = textColor;
    state.charts.brandChart.options.scales.y.grid.color = gridColor;
    state.charts.brandChart.update();
  } else {
    state.charts.brandChart = new Chart(canvas, {
      type: 'bar',
      data: chartData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false }
        },
        scales: {
          x: { ticks: { color: textColor, font: { family: 'Inter' } }, grid: { display: false } },
          y: { ticks: { color: textColor, font: { family: 'Inter' } }, grid: { color: gridColor } }
        }
      }
    });
  }
}

function renderCriticalChart(textColor, gridColor) {
  const canvas = document.getElementById('criticalChartCanvas');
  if (!canvas) return;

  // Filter top 10 items with lowest stock compared to minimum requirement
  const criticalItems = state.inventory
    .filter(i => {
      const min = Number(i.min) || 0;
      const qty = Number(i.qty) || 0;
      return min > 0 && qty <= min;
    })
    .sort((a, b) => {
      const defA = (Number(a.qty) || 0) - (Number(a.min) || 0);
      const defB = (Number(b.qty) || 0) - (Number(b.min) || 0);
      return defA - defB;
    })
    .slice(0, 10);

  const labels = criticalItems.map(i => (i.pnFinal || '').slice(0, 15));
  const currentQtys = criticalItems.map(i => Number(i.qty) || 0);
  const minQtys = criticalItems.map(i => Number(i.min) || 0);

  const chartData = {
    labels: labels,
    datasets: [
      {
        label: 'Stok Saat Ini',
        data: currentQtys,
        backgroundColor: '#ef4444',
        borderRadius: 4
      },
      {
        label: 'Minimum Requirement',
        data: minQtys,
        backgroundColor: 'rgba(148, 163, 184, 0.4)',
        borderRadius: 4
      }
    ]
  };

  if (state.charts.criticalChart) {
    state.charts.criticalChart.data = chartData;
    state.charts.criticalChart.options.scales.x.ticks.color = textColor;
    state.charts.criticalChart.options.scales.y.ticks.color = textColor;
    state.charts.criticalChart.options.scales.x.grid.color = gridColor;
    state.charts.criticalChart.update();
  } else {
    state.charts.criticalChart = new Chart(canvas, {
      type: 'bar',
      data: chartData,
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'top',
            labels: { color: textColor, font: { size: 11, family: 'Inter' } }
          }
        },
        scales: {
          x: { ticks: { color: textColor, font: { family: 'Inter' } }, grid: { color: gridColor } },
          y: { ticks: { color: textColor, font: { family: 'Inter', size: 10 } }, grid: { display: false } }
        }
      }
    });
  }
}

function renderServiceChart(textColor) {
  const canvas = document.getElementById('serviceChartCanvas');
  if (!canvas) return;

  let psCount = 0;
  let nonPsCount = 0;

  state.inventory.forEach(i => {
    if ((i.status || '').toUpperCase() === 'PS') psCount++;
    else nonPsCount++;
  });

  const chartData = {
    labels: ['Periodic Service (PS)', 'Non-PS'],
    datasets: [{
      data: [psCount, nonPsCount],
      backgroundColor: ['#10b981', '#64748b'],
      borderWidth: 2,
      borderColor: (document.documentElement.getAttribute('data-theme') || 'dark') === 'dark' ? '#1e293b' : '#ffffff'
    }]
  };

  if (state.charts.serviceChart) {
    state.charts.serviceChart.data = chartData;
    state.charts.serviceChart.options.plugins.legend.labels.color = textColor;
    state.charts.serviceChart.update();
  } else {
    state.charts.serviceChart = new Chart(canvas, {
      type: 'doughnut',
      data: chartData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: textColor, font: { size: 11, family: 'Inter' }, padding: 12 }
          }
        },
        cutout: '65%'
      }
    });
  }
}

/**
 * Export to CSV
 */
function exportToCsv() {
  const headers = ['No', 'Prod', 'Model Unit', 'Part Number', 'Description', 'Min', 'Max', 'Status', 'Qty', 'Month', 'Week', 'Day', 'DOI', 'Remarks', 'Alternate PN'];
  const rows = state.filtered.map(item => [
    item.no || '',
    `"${(item.prod || '').replace(/"/g, '""')}"`,
    `"${(item.modelUnit || '').replace(/"/g, '""')}"`,
    `"${(item.pnFinal || '').replace(/"/g, '""')}"`,
    `"${(item.desc || '').replace(/"/g, '""')}"`,
    item.min ?? '',
    item.max ?? '',
    item.status || '',
    item.qty ?? '',
    item.month ?? '',
    item.week ?? '',
    item.day ?? '',
    item.doi ?? '',
    `"${(item.remarks || '').replace(/"/g, '""')}"`,
    `"${(item.alternatePn || '').replace(/"/g, '""')}"`
  ]);

  const csvContent = 'data:text/csv;charset=utf-8,\uFEFF' + [headers.join(','), ...rows.map(e => e.join(','))].join('\n');
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement('a');
  link.setAttribute('href', encodedUri);
  link.setAttribute('download', `Inventory_SpareParts_${new Date().toISOString().slice(0, 10)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast('File CSV berhasil diunduh!', 'success');
}

/**
 * Export to JSON
 */
function exportToJson() {
  const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(state.filtered, null, 2));
  const link = document.createElement('a');
  link.setAttribute('href', dataStr);
  link.setAttribute('download', `Inventory_SpareParts_${new Date().toISOString().slice(0, 10)}.json`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast('File JSON berhasil diunduh!', 'success');
}

/**
 * Toast Notification Helper
 */
function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const icons = {
    success: 'fa-solid fa-circle-check',
    error: 'fa-solid fa-circle-xmark',
    warning: 'fa-solid fa-triangle-exclamation',
    info: 'fa-solid fa-circle-info'
  };

  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `
    <i class="${icons[type] || icons.info}"></i>
    <span>${escapeHtml(message)}</span>
  `;

  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(100%)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

/**
 * Modal Open / Close Helpers
 */
function openModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.add('active');
}

function closeAllModals() {
  document.querySelectorAll('.modal-overlay').forEach(el => el.classList.remove('active'));
}

/**
 * Utility: Set text content safely
 */
function setText(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

/**
 * Utility: Escape HTML
 */
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function escapeQuote(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/'/g, "\\'").replace(/"/g, '&quot;');
}
