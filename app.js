/**
 * ============================================================================
 * PT. STAR WAGEN INDONESIA - VIB Spare Parts Inventory Dashboard - Main Controller (app.js)
 * ============================================================================
 * - Friction-Free Team Real-Time Updates (Edit PN, QTY, Min/Max, PS/NON PS, LIB/NON LIB)
 * - Interactive Inline Editing: Click-to-toggle PS/LIB, Quick QTY Steppers (+/-)
 * - Floating Multi-Select Bulk Actions Bar (Update tens or hundreds of parts in 1 click)
 * - Universal Bulk Excel Parser (Pasting ANY column combination from office Excel)
 * - Real-Time Google Sheets Cloud Sync (Multi-Sheet) & Automated Recalculation
 * ============================================================================
 */

// Global Application State
const state = {
  currentView: 'dashboard',
  master: [],
  filteredMaster: [],
  daily: [],
  filteredDaily: [],
  audit: [],
  filteredAudit: [],
  
  // Multi-Select for Mass Bulk Editing
  selectedMasterCodes: new Set(),

  // Master Pagination & Sorting
  masterPage: 1,
  masterPageSize: 25,
  masterSortCol: 'no',
  masterSortDir: 'asc',
  masterFilters: {
    search: '',
    lokasi: 'ALL',
    service: 'ALL',
    lib: 'ALL',
    statusTab: 'ALL'
  },

  // Daily Pagination & Sorting
  dailyPage: 1,
  dailyPageSize: 25,
  dailySortCol: 'tanggal',
  dailySortDir: 'desc',
  dailyFilters: {
    search: '',
    date: '',
    month: 'ALL',
    year: '2026',
    hour: 'ALL'
  },

  // Bulk Import Preview Cache
  parsedImport: null,

  // History Tab & Filter
  historyTab: 'ledger', // 'ledger' or 'audit'
  historyPage: 1,
  historyPageSize: 25,
  historyFilters: {
    search: '',
    date: '',
    hour: 'ALL',
    lokasi: 'ALL'
  },

  // Reports State
  reportType: 'as-of-date', // 'as-of-date' or 'valuation'
  reportTargetDate: '',
  reportData: [],
  reportPage: 1,
  reportPageSize: 25,

  // Config & Sync
  config: {
    gasWebAppUrl: '',
    spreadsheetId: '',
    autoSync: true,
    syncIntervalMinutes: 1,
    theme: 'dark'
  },
  syncTimer: null,
  isSyncing: false,

  // Chart Instances
  charts: {
    statusChart: null,
    brandChart: null,
    criticalChart: null,
    serviceChart: null
  }
};

/**
 * ============================================================================
 * ROLE-BASED ACCESS CONTROL (RBAC) - ADMIN & PENGUNJUNG (VISITOR)
 * ============================================================================
 * - Admin: Full access (Edit, add, delete, steppers, toggle PS/LIB, bulk edit, Excel import, sync)
 * - Pengunjung: Read-only (Melihat data terupdate, grafik, riwayat, dan ekspor laporan)
 * ============================================================================
 */
let pendingAdminAction = null;

function isAdminLoggedIn() {
  return sessionStorage.getItem('he_inventory_admin_session') === 'true';
}

function getAdminPassword() {
  return localStorage.getItem('he_inventory_admin_pass') || 'admin123';
}

function openAdminLoginModal(callback = null, actionDescription = '') {
  if (isAdminLoggedIn()) {
    if (typeof callback === 'function') callback();
    return;
  }

  pendingAdminAction = callback;
  const passInput = document.getElementById('adminPassInput') || document.getElementById('adminPasswordInput');
  if (passInput) {
    passInput.value = '';
    passInput.type = 'password';
  }

  const eyeIcon = document.getElementById('adminPassEye') || document.querySelector('#btnToggleAdminPass i, #btnTogglePassword i');
  if (eyeIcon) {
    eyeIcon.classList.remove('fa-eye-slash');
    eyeIcon.classList.add('fa-eye');
  }

  const promptEl = document.getElementById('adminLoginPromptText');
  if (promptEl) {
    if (actionDescription) {
      promptEl.innerHTML = `Silakan masukkan password admin untuk <strong>${escapeHtml(actionDescription)}</strong>.`;
    } else {
      promptEl.textContent = 'Masukkan password admin untuk mengakses fitur ubah stok, tambah/edit data, dan konfigurasi Google Sheets.';
    }
  }

  if (actionDescription) {
    showToast(`Silakan login sebagai Admin untuk ${actionDescription}.`, 'warning');
  }
  openModal('adminLoginModal');
  setTimeout(() => passInput?.focus(), 250);
}

function toggleAdminPassVisibility() {
  const passInput = document.getElementById('adminPassInput') || document.getElementById('adminPasswordInput');
  const eyeIcon = document.getElementById('adminPassEye') || document.querySelector('#btnToggleAdminPass i, #btnTogglePassword i');
  if (!passInput) return;

  if (passInput.type === 'password') {
    passInput.type = 'text';
    if (eyeIcon) {
      eyeIcon.classList.remove('fa-eye');
      eyeIcon.classList.add('fa-eye-slash');
    }
  } else {
    passInput.type = 'password';
    if (eyeIcon) {
      eyeIcon.classList.remove('fa-eye-slash');
      eyeIcon.classList.add('fa-eye');
    }
  }
}

function requireAdmin(actionDescription, callback) {
  if (isAdminLoggedIn()) {
    return callback();
  }
  openAdminLoginModal(callback, actionDescription);
}

function handleAdminLoginSubmit(e) {
  if (e) e.preventDefault();
  const passInput = document.getElementById('adminPassInput') || document.getElementById('adminPasswordInput');
  const entered = (passInput ? passInput.value : '').trim();
  const correct = getAdminPassword();

  if (entered === correct) {
    sessionStorage.setItem('he_inventory_admin_session', 'true');
    closeModal('adminLoginModal');
    showToast('Login berhasil! Anda sekarang dalam Mode Admin (Akses Penuh).', 'success');
    renderAuthStatus();
    renderCurrentView();

    if (typeof pendingAdminAction === 'function') {
      const act = pendingAdminAction;
      pendingAdminAction = null;
      act();
    }
  } else {
    showToast('Password Admin salah! Silakan periksa kembali.', 'danger');
    if (passInput) {
      passInput.focus();
      passInput.select();
    }
  }
}

function handleAdminLogout() {
  sessionStorage.removeItem('he_inventory_admin_session');
  state.selectedMasterCodes.clear();
  showToast('Anda telah keluar. Sekarang dalam Mode Pengunjung (Read-Only).', 'info');
  renderAuthStatus();
  renderCurrentView();
}

function openChangePasswordModal() {
  requireAdmin('mengubah password admin', () => {
    document.getElementById('changePasswordForm')?.reset();
    openModal('changePasswordModal');
    setTimeout(() => {
      document.getElementById('currentPassInput')?.focus();
    }, 200);
  });
}

async function handleChangePasswordSubmit(e) {
  if (e) e.preventDefault();
  const currentPassInput = document.getElementById('currentPassInput');
  const newPassInput = document.getElementById('newPassInput');
  const confirmPassInput = document.getElementById('confirmPassInput');

  const currentPass = (currentPassInput?.value || '').trim();
  const newPass = (newPassInput?.value || '').trim();
  const confirmPass = (confirmPassInput?.value || '').trim();

  if (currentPass !== getAdminPassword()) {
    showToast('Password lama tidak sesuai!', 'danger');
    currentPassInput?.focus();
    return;
  }
  if (!newPass || newPass.length < 4) {
    showToast('Password baru minimal 4 karakter!', 'warning');
    newPassInput?.focus();
    return;
  }
  if (newPass !== confirmPass) {
    showToast('Konfirmasi password baru tidak cocok!', 'warning');
    confirmPassInput?.focus();
    return;
  }

  localStorage.setItem('he_inventory_admin_pass', newPass);
  closeModal('changePasswordModal');
  showToast('Password Admin berhasil diperbarui!', 'success');
  if (window.InventoryDB?.addAuditLog) {
    await window.InventoryDB.addAuditLog('CHANGE_PASSWORD', 'Admin', 'ALL', 'Password admin berhasil diubah');
  }
}

function renderAuthStatus() {
  const container = document.getElementById('adminAuthControls');
  if (!container) return;

  const isAdmin = isAdminLoggedIn();

  if (isAdmin) {
    container.innerHTML = `
      <div class="admin-auth-wrapper">
        <span class="badge badge-admin" title="Mode Administrator: Hak akses penuh edit & import">
          <i class="fa-solid fa-user-shield"></i> Mode Admin (Akses Penuh)
        </span>
        <button class="btn btn-secondary btn-sm" id="btnAdminChangePass" title="Ganti Password Admin">
          <i class="fa-solid fa-key"></i>
        </button>
        <button class="btn btn-danger btn-sm" id="btnAdminLogout" title="Keluar ke Mode Pengunjung (Read-Only)">
          <i class="fa-solid fa-right-from-bracket"></i> Keluar
        </button>
      </div>
    `;
    document.getElementById('btnAdminChangePass')?.addEventListener('click', openChangePasswordModal);
    document.getElementById('btnAdminLogout')?.addEventListener('click', handleAdminLogout);
  } else {
    container.innerHTML = `
      <div class="admin-auth-wrapper">
        <span class="badge badge-guest" title="Mode Pengunjung: Hanya dapat melihat data (Read-Only)">
          <i class="fa-solid fa-eye"></i> Mode Pengunjung (Read-Only)
        </span>
        <button class="btn btn-warning btn-sm" id="btnAdminLoginTrigger" title="Login Admin untuk mengedit & update data">
          <i class="fa-solid fa-lock"></i> Login Admin
        </button>
      </div>
    `;
    document.getElementById('btnAdminLoginTrigger')?.addEventListener('click', () => openAdminLoginModal(null, 'Mode Admin'));
  }
}

/**
 * ============================================================================
 * NAVIGATION & VIEW SWITCHING
 * ============================================================================
 */
function switchAppView(viewName) {
  state.currentView = viewName;

  document.querySelectorAll('.nav-tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.getAttribute('data-view') === viewName);
  });

  document.querySelectorAll('.app-view').forEach(view => {
    const isTarget = view.id === ('view' + viewName.charAt(0).toUpperCase() + viewName.slice(1));
    view.classList.toggle('active', isTarget);
  });

  renderCurrentView();
}

function renderCurrentView() {
  const isAdmin = isAdminLoggedIn();

  // Daily View Gating
  const dailyNotice = document.getElementById('dailyGuestNotice');
  const dailyForm = document.getElementById('dailyQuickFormCard');
  if (dailyNotice && dailyForm) {
    if (isAdmin) {
      dailyNotice.style.display = 'none';
      dailyForm.style.display = 'block';
    } else {
      dailyNotice.style.display = 'flex';
      dailyForm.style.display = 'none';
    }
  }

  // Import View Gating
  const importNotice = document.getElementById('importGuestNotice');
  const importWorkspace = document.getElementById('importAdminWorkspace');
  if (importNotice && importWorkspace) {
    if (isAdmin) {
      importNotice.style.display = 'none';
      importWorkspace.style.display = 'grid';
    } else {
      importNotice.style.display = 'block';
      importWorkspace.style.display = 'none';
    }
  }

  // Subtitle updates for role clarity
  const subTitle = document.getElementById('masterViewSubTitle');
  if (subTitle) {
    subTitle.innerHTML = isAdmin
      ? 'Katalog tunggal suku cadang alat berat dengan Kode Material unik, lokasi, harga, dan stok. <strong style="color: #34d399;">(Mode Admin: Akses Edit Aktif)</strong>'
      : 'Katalog tunggal suku cadang alat berat dengan Kode Material unik, lokasi, harga, dan stok. <span style="color: #94a3b8;">(Mode Pengunjung: Hanya Lihat Data)</span>';
  }

  switch (state.currentView) {
    case 'dashboard':
      updateDashboard();
      break;
    case 'master':
      applyMasterFilters();
      break;
    case 'daily':
      applyDailyFilters();
      break;
    case 'import':
      break;
    case 'history':
      applyHistoryFilters();
      break;
    case 'reports':
      generateReport();
      break;
  }
}

/**
 * ============================================================================
 * DATA REFRESH & SYNCHRONIZATION FROM DB
 * ============================================================================
 */
async function loadDataFromDB() {
  try {
    const [master, daily, audit] = await Promise.all([
      window.InventoryDB.getAllMasterMaterials(),
      window.InventoryDB.getAllDailyStock(),
      window.InventoryDB.getAllAuditLogs()
    ]);

    state.master = master || [];
    state.daily = daily || [];
    state.audit = audit || [];

    populateMasterDropdowns();
    populateDailyDropdowns();
    populateDatalist();

    renderCurrentView();
  } catch (err) {
    console.error('Error loading DB data:', err);
    showToast('Gagal memuat data dari database: ' + err.message, 'danger');
  }
}

function populateMasterDropdowns() {
  const selLokasi = document.getElementById('filterMasterLokasi');
  const selHistoryLokasi = document.getElementById('filterHistoryLokasi');

  const lokSet = new Set();
  state.master.forEach(m => {
    if (m.lokasi) lokSet.add(m.lokasi.trim());
  });

  const lokOptions = '<option value="ALL">Semua Lokasi</option>' +
    Array.from(lokSet).sort().map(l => `<option value="${escapeQuote(l)}">${escapeHtml(l)}</option>`).join('');

  if (selLokasi) {
    const curVal = selLokasi.value;
    selLokasi.innerHTML = lokOptions;
    selLokasi.value = curVal || 'ALL';
  }

  if (selHistoryLokasi) {
    const curVal = selHistoryLokasi.value;
    selHistoryLokasi.innerHTML = lokOptions;
    selHistoryLokasi.value = curVal || 'ALL';
  }
}

function populateDailyDropdowns() {
  const selDailyHour = document.getElementById('filterDailyHour');
  const selHistoryHour = document.getElementById('filterHistoryHour');

  let hourOptions = '<option value="ALL">Semua Jam</option>';
  for (let h = 0; h < 24; h++) {
    const hh = String(h).padStart(2, '0');
    hourOptions += `<option value="${hh}">${hh}:00 - ${hh}:59</option>`;
  }

  if (selDailyHour) selDailyHour.innerHTML = hourOptions;
  if (selHistoryHour) selHistoryHour.innerHTML = hourOptions;
}

function populateDatalist() {
  const datalist = document.getElementById('masterKodeList');
  if (!datalist) return;
  datalist.innerHTML = state.master.map(m => 
    `<option value="${escapeQuote(m.kodeMaterial)}">${escapeHtml(m.namaMaterial)} [${m.status || 'PS'}] [${m.libCategory || 'NON LIB'}] (${escapeHtml(m.lokasi)})</option>`
  ).join('');
}

/**
 * ============================================================================
 * MODULE 1: DASHBOARD & ANALYTICS
 * ============================================================================
 */
function updateDashboard() {
  const master = state.master;
  const totalItems = master.length;
  let totalStock = 0;
  let understock = 0;
  let outstock = 0;
  let overstock = 0;
  let totalValuation = 0;

  master.forEach(item => {
    const qty = Number(item.stokSaatIni) || 0;
    const min = item.min !== null && item.min !== undefined ? Number(item.min) : 1;
    const max = item.max !== null && item.max !== undefined ? Number(item.max) : 5;
    const price = Number(item.price) || 0;

    totalStock += qty;
    totalValuation += (qty * price);

    if (qty === 0) {
      outstock++;
    } else if (qty < min) {
      understock++;
    } else if (qty > max) {
      overstock++;
    }
  });

  setText('kpiTotalItems', totalItems.toLocaleString('id-ID'));
  setText('kpiTotalStock', totalStock.toLocaleString('id-ID'));
  setText('kpiUnderstock', understock.toLocaleString('id-ID'));
  setText('kpiOutOfStock', outstock.toLocaleString('id-ID'));
  setText('kpiOverstock', overstock.toLocaleString('id-ID'));
  setText('kpiTotalValuation', formatRupiah(totalValuation));

  updateCharts();
}

function updateCharts() {
  if (typeof Chart === 'undefined') return;

  const isLight = document.documentElement.getAttribute('data-theme') === 'light';
  const textColor = isLight ? '#1e293b' : '#f8fafc';
  const gridColor = isLight ? 'rgba(0, 0, 0, 0.06)' : 'rgba(255, 255, 255, 0.06)';

  renderStatusChart(textColor);
  renderBrandChart(textColor, gridColor);
  renderServiceChart(textColor);
  renderCriticalChart(textColor, gridColor);
}

function renderStatusChart(textColor) {
  const canvas = document.getElementById('statusChartCanvas');
  if (!canvas) return;

  let optimal = 0, understock = 0, outstock = 0, overstock = 0;
  state.master.forEach(m => {
    const qty = Number(m.stokSaatIni) || 0;
    const min = m.min !== null ? Number(m.min) : 1;
    const max = m.max !== null ? Number(m.max) : 5;

    if (qty === 0) outstock++;
    else if (qty < min) understock++;
    else if (qty > max) overstock++;
    else optimal++;
  });

  if (state.charts.statusChart) state.charts.statusChart.destroy();

  state.charts.statusChart = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels: ['Optimal', 'Understock', 'Out of Stock', 'Overstock'],
      datasets: [{
        data: [optimal, understock, outstock, overstock],
        backgroundColor: ['#10b981', '#f59e0b', '#ef4444', '#06b6d4'],
        borderWidth: 2,
        borderColor: 'transparent'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: textColor, boxWidth: 12, padding: 12, font: { family: 'Inter', size: 11 } }
        }
      },
      cutout: '68%'
    }
  });
}

function renderBrandChart(textColor, gridColor) {
  const canvas = document.getElementById('brandChartCanvas');
  if (!canvas) return;

  const locMap = {};
  state.master.forEach(m => {
    const key = (m.lokasi || 'GUDANG UTAMA').trim();
    locMap[key] = (locMap[key] || 0) + 1;
  });

  const sorted = Object.entries(locMap).sort((a, b) => b[1] - a[1]).slice(0, 7);
  const labels = sorted.map(s => s[0]);
  const data = sorted.map(s => s[1]);

  if (state.charts.brandChart) state.charts.brandChart.destroy();

  state.charts.brandChart = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Jumlah Part per Lokasi',
        data: data,
        backgroundColor: '#8b5cf6',
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: textColor, font: { size: 10 } }, grid: { display: false } },
        y: { ticks: { color: textColor }, grid: { color: gridColor } }
      }
    }
  });
}

function renderServiceChart(textColor) {
  const canvas = document.getElementById('serviceChartCanvas');
  if (!canvas) return;

  let ps = 0, nonPs = 0, lib = 0, nonLib = 0;
  state.master.forEach(m => {
    if (String(m.status).toUpperCase().includes('PS') && !String(m.status).toUpperCase().includes('NON')) ps++;
    else nonPs++;

    if (String(m.libCategory).toUpperCase() === 'LIB') lib++;
    else nonLib++;
  });

  if (state.charts.serviceChart) state.charts.serviceChart.destroy();

  state.charts.serviceChart = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels: ['PS (Periodic)', 'NON PS', 'Kategori LIB', 'NON LIB'],
      datasets: [{
        data: [ps, nonPs, lib, nonLib],
        backgroundColor: ['#10b981', '#64748b', '#06b6d4', '#94a3b8'],
        borderWidth: 2,
        borderColor: 'transparent'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: textColor, boxWidth: 12, padding: 10, font: { family: 'Inter', size: 11 } }
        }
      },
      cutout: '65%'
    }
  });
}

function renderCriticalChart(textColor, gridColor) {
  const canvas = document.getElementById('criticalChartCanvas');
  if (!canvas) return;

  const critical = state.master
    .filter(m => (m.min !== null && m.min !== undefined) && ((Number(m.stokSaatIni) || 0) < Number(m.min)))
    .sort((a, b) => {
      const defA = (Number(a.min) || 0) - (Number(a.stokSaatIni) || 0);
      const defB = (Number(b.min) || 0) - (Number(b.stokSaatIni) || 0);
      return defB - defA;
    })
    .slice(0, 10);

  const labels = critical.map(c => (c.kodeMaterial + ' - ' + (c.namaMaterial || '')).substring(0, 30));
  const currentQty = critical.map(c => Number(c.stokSaatIni) || 0);
  const minQty = critical.map(c => Number(c.min) || 1);

  if (state.charts.criticalChart) state.charts.criticalChart.destroy();

  state.charts.criticalChart = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'Stok Fisik Saat Ini',
          data: currentQty,
          backgroundColor: '#ef4444',
          borderRadius: 4
        },
        {
          label: 'Batas Minimum Aman (Min)',
          data: minQty,
          backgroundColor: 'rgba(148, 163, 184, 0.4)',
          borderRadius: 4
        }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'top',
          labels: { color: textColor, font: { family: 'Inter', size: 11 } }
        }
      },
      scales: {
        x: { ticks: { color: textColor }, grid: { color: gridColor } },
        y: { ticks: { color: textColor, font: { size: 10 } }, grid: { display: false } }
      }
    }
  });
}

/**
 * ============================================================================
 * MODULE 2: MASTER DATA MATERIAL & FAST BULK EDITING
 * ============================================================================
 */
function applyMasterFilters() {
  const { search, lokasi, service, lib, statusTab } = state.masterFilters;
  const q = search.toLowerCase().trim();

  let totalAll = state.master.length;
  let cntOptimal = 0, cntUnder = 0, cntHabis = 0, cntOver = 0;

  state.master.forEach(m => {
    const qty = Number(m.stokSaatIni) || 0;
    const min = m.min !== null ? Number(m.min) : 1;
    const max = m.max !== null ? Number(m.max) : 5;
    if (qty === 0) cntHabis++;
    else if (qty < min) cntUnder++;
    else if (qty > max) cntOver++;
    else cntOptimal++;
  });

  setText('countMasterAll', totalAll);
  setText('countMasterOptimal', cntOptimal);
  setText('countMasterUnderstock', cntUnder);
  setText('countMasterHabis', cntHabis);
  setText('countMasterOverstock', cntOver);

  state.filteredMaster = state.master.filter(m => {
    if (q) {
      const match = (m.kodeMaterial && m.kodeMaterial.toLowerCase().includes(q)) ||
                    (m.namaMaterial && m.namaMaterial.toLowerCase().includes(q)) ||
                    (m.lokasi && m.lokasi.toLowerCase().includes(q)) ||
                    (m.modelUnit && m.modelUnit.toLowerCase().includes(q));
      if (!match) return false;
    }

    if (lokasi !== 'ALL' && m.lokasi !== lokasi) return false;

    if (service !== 'ALL') {
      const isPs = String(m.status).toUpperCase().includes('PS') && !String(m.status).toUpperCase().includes('NON');
      if (service === 'PS' && !isPs) return false;
      if (service === 'NON PS' && isPs) return false;
    }

    if (lib !== 'ALL') {
      const isLib = String(m.libCategory).toUpperCase() === 'LIB';
      if (lib === 'LIB' && !isLib) return false;
      if (lib === 'NON LIB' && isLib) return false;
    }

    if (statusTab !== 'ALL') {
      const qty = Number(m.stokSaatIni) || 0;
      const min = m.min !== null ? Number(m.min) : 1;
      const max = m.max !== null ? Number(m.max) : 5;

      if (statusTab === 'HABIS' && qty !== 0) return false;
      if (statusTab === 'UNDERSTOCK' && (qty === 0 || qty >= min)) return false;
      if (statusTab === 'OPTIMAL' && (qty < min || qty > max)) return false;
      if (statusTab === 'OVERSTOCK' && qty <= max) return false;
    }

    return true;
  });

  state.filteredMaster.sort((a, b) => {
    let valA = a[state.masterSortCol];
    let valB = b[state.masterSortCol];
    if (state.masterSortCol === 'stokSaatIni' || state.masterSortCol === 'price' || state.masterSortCol === 'min' || state.masterSortCol === 'max') {
      const numA = Number(valA) || 0;
      const numB = Number(valB) || 0;
      return state.masterSortDir === 'asc' ? numA - numB : numB - numA;
    }
    if (typeof valA === 'string') valA = valA.toLowerCase();
    if (typeof valB === 'string') valB = valB.toLowerCase();
    if (valA < valB) return state.masterSortDir === 'asc' ? -1 : 1;
    if (valA > valB) return state.masterSortDir === 'asc' ? 1 : -1;
    return 0;
  });

  state.masterPage = 1;
  renderMasterTable();
}

function setupMasterTableSorting() {
  document.querySelectorAll('#masterTableHead th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (!col) return;
      if (state.masterSortCol === col) {
        state.masterSortDir = state.masterSortDir === 'asc' ? 'desc' : 'asc';
      } else {
        state.masterSortCol = col;
        state.masterSortDir = 'asc';
      }
      updateMasterSortIcons();
      applyMasterFilters();
    });
  });
  updateMasterSortIcons();
}

function updateMasterSortIcons() {
  document.querySelectorAll('#masterTableHead th[data-sort]').forEach(th => {
    const col = th.getAttribute('data-sort');
    const icon = th.querySelector('i');
    if (!icon) return;
    if (state.masterSortCol === col) {
      icon.className = state.masterSortDir === 'asc' ? 'fa-solid fa-sort-up' : 'fa-solid fa-sort-down';
      icon.style.color = '#38bdf8';
    } else {
      icon.className = 'fa-solid fa-sort';
      icon.style.color = '';
    }
  });
}

function renderMasterTable() {
  const tbody = document.getElementById('masterTableBody');
  if (!tbody) return;

  const isAdmin = isAdminLoggedIn();
  const total = state.filteredMaster.length;
  const pageSize = state.masterPageSize === 'ALL' ? total : Number(state.masterPageSize);
  const totalPages = Math.ceil(total / pageSize) || 1;

  if (state.masterPage > totalPages) state.masterPage = totalPages;

  const start = (state.masterPage - 1) * pageSize;
  const end = Math.min(start + pageSize, total);
  const pageItems = state.filteredMaster.slice(start, end);

  updateMasterBulkBar();

  if (pageItems.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="12" style="text-align: center; padding: 2.5rem; color: var(--text-muted);">
          <i class="fa-solid fa-box-open" style="font-size: 2rem; margin-bottom: 0.5rem; display: block;"></i>
          Tidak ada data material yang cocok dengan filter.
        </td>
      </tr>
    `;
    setText('masterPaginationInfo', 'Menampilkan 0 data');
    document.getElementById('masterPaginationControls').innerHTML = '';
    return;
  }

  tbody.innerHTML = pageItems.map((item, idx) => {
    const qty = Number(item.stokSaatIni) || 0;
    const min = item.min !== null ? Number(item.min) : 1;
    const max = item.max !== null ? Number(item.max) : 5;
    const isChecked = isAdmin && state.selectedMasterCodes.has(item.kodeMaterial);

    const isPs = String(item.status || 'PS').toUpperCase().includes('PS') && !String(item.status).toUpperCase().includes('NON');
    const isLib = String(item.libCategory || 'NON LIB').toUpperCase() === 'LIB';

    let healthBadge = '';
    if (qty === 0) {
      healthBadge = '<span class="status-badge status-outstock"><i class="fa-solid fa-ban"></i> Habis</span>';
    } else if (qty < min) {
      healthBadge = `<span class="status-badge status-understock"><i class="fa-solid fa-triangle-exclamation"></i> Kritis</span>`;
    } else if (qty > max) {
      healthBadge = `<span class="status-badge status-overstock"><i class="fa-solid fa-arrow-trend-up"></i> Over</span>`;
    } else {
      healthBadge = '<span class="status-badge status-optimal"><i class="fa-solid fa-circle-check"></i> Optimal</span>';
    }

    const priceFormatted = item.price ? formatRupiah(item.price) : 'Rp 0';

    // Role-based elements
    const checkboxHtml = isAdmin
      ? `<input type="checkbox" class="master-row-check" data-kode="${escapeQuote(item.kodeMaterial)}" ${isChecked ? 'checked' : ''}>`
      : `<input type="checkbox" disabled style="opacity: 0.3; cursor: not-allowed;" title="Khusus Admin">`;

    const psBadgeHtml = isAdmin
      ? `<span class="badge-toggle-ps ${isPs ? 'badge-ps' : 'badge-non-ps'}" onclick="togglePartService('${escapeQuote(item.kodeMaterial)}')" title="Klik untuk ubah status PS / NON PS"><i class="fa-solid ${isPs ? 'fa-wrench' : 'fa-ban'}"></i> ${isPs ? 'PS' : 'NON PS'}</span>`
      : `<span class="badge-static-ps ${isPs ? 'badge-ps' : 'badge-non-ps'}" title="Status Servis Part (Mode Pengunjung)"><i class="fa-solid ${isPs ? 'fa-wrench' : 'fa-ban'}"></i> ${isPs ? 'PS' : 'NON PS'}</span>`;

    const libBadgeHtml = isAdmin
      ? `<span class="badge-toggle-lib ${isLib ? 'badge-lib' : 'badge-non-lib'}" onclick="togglePartLibCategory('${escapeQuote(item.kodeMaterial)}')" title="Klik untuk ubah kategori LIB / NON LIB"><i class="fa-solid ${isLib ? 'fa-oil-can' : 'fa-cube'}"></i> ${isLib ? 'LIB' : 'NON LIB'}</span>`
      : `<span class="badge-static-lib ${isLib ? 'badge-lib' : 'badge-non-lib'}" title="Kategori LIB Part (Mode Pengunjung)"><i class="fa-solid ${isLib ? 'fa-oil-can' : 'fa-cube'}"></i> ${isLib ? 'LIB' : 'NON LIB'}</span>`;

    const qtyHtml = isAdmin
      ? `<div class="qty-stepper">
          <button class="qty-stepper-btn" onclick="stepPartQty('${escapeQuote(item.kodeMaterial)}', -1)" title="Kurangi 1">-</button>
          <span class="qty-val-display" onclick="inlineEditQty(this, '${escapeQuote(item.kodeMaterial)}')" title="Klik untuk ketik nilai langsung">${qty}</span>
          <button class="qty-stepper-btn" onclick="stepPartQty('${escapeQuote(item.kodeMaterial)}', 1)" title="Tambah 1">+</button>
        </div>`
      : `<strong style="font-family: 'JetBrains Mono', monospace; font-size: 0.95rem; color: var(--text-primary);">${qty}</strong>`;

    const actionsHtml = isAdmin
      ? `<div style="display: inline-flex; gap: 0.35rem;">
          <button class="btn btn-icon btn-sm" onclick="quickDailyForMaterial('${escapeQuote(item.kodeMaterial)}')" title="Catat Transaksi Harian untuk Part Ini">
            <i class="fa-solid fa-calendar-plus" style="color: #38bdf8;"></i>
          </button>
          <button class="btn btn-icon btn-sm" onclick="openEditMasterModal('${escapeQuote(item.kodeMaterial)}')" title="Edit Detail Master">
            <i class="fa-solid fa-pen-to-square" style="color: var(--primary);"></i>
          </button>
          <button class="btn btn-icon btn-sm" onclick="confirmDeleteMaster('${escapeQuote(item.kodeMaterial)}')" title="Hapus Master Material">
            <i class="fa-solid fa-trash" style="color: var(--danger);"></i>
          </button>
        </div>`
      : `<span class="badge badge-guest" style="cursor: pointer;" onclick="openAdminLoginModal(null, 'mengedit material ini')" title="Klik untuk Login Admin">
          <i class="fa-solid fa-eye"></i> Hanya Lihat
        </span>`;

    return `
      <tr class="${isChecked ? 'selected-row' : ''}">
        <td style="text-align: center;">${checkboxHtml}</td>
        <td style="color: var(--text-muted); font-size: 0.85rem;">${start + idx + 1}</td>
        <td>
          <strong style="font-family: 'JetBrains Mono', monospace; color: #93c5fd;">${escapeHtml(item.kodeMaterial)}</strong>
        </td>
        <td>
          <span class="badge" style="background: rgba(148, 163, 184, 0.15); color: var(--text-primary); font-size: 0.8rem;">
            ${escapeHtml(item.lokasi || '-')}
          </span>
        </td>
        <td>
          <div style="font-weight: 600; color: var(--text-primary);">${escapeHtml(item.namaMaterial)}</div>
          ${item.modelUnit ? `<div style="font-size: 0.75rem; color: var(--text-muted);">${escapeHtml(item.modelUnit)}</div>` : ''}
        </td>
        <td style="text-align: center;">${psBadgeHtml}</td>
        <td style="text-align: center;">${libBadgeHtml}</td>
        <td style="font-family: 'JetBrains Mono', monospace; font-size: 0.85rem; color: var(--text-secondary);">
          <span title="Min" style="color: var(--warning);">${min}</span> - <span title="Max" style="color: var(--info);">${max}</span>
        </td>
        <td>${qtyHtml}</td>
        <td>${healthBadge}</td>
        <td style="font-family: 'JetBrains Mono', monospace; font-size: 0.82rem;">
          ${priceFormatted}
        </td>
        <td style="text-align: center;">${actionsHtml}</td>
      </tr>
    `;
  }).join('');

  setText('masterPaginationInfo', `Menampilkan ${start + 1} - ${end} dari ${total} material`);
  renderPaginationControls('masterPaginationControls', state.masterPage, totalPages, (p) => {
    state.masterPage = p;
    renderMasterTable();
  });

  // Attach Checkbox Listeners for Admin
  if (isAdmin) {
    document.querySelectorAll('.master-row-check').forEach(cb => {
      cb.addEventListener('change', () => {
        const kode = cb.getAttribute('data-kode');
        if (cb.checked) {
          state.selectedMasterCodes.add(kode);
        } else {
          state.selectedMasterCodes.delete(kode);
        }
        updateMasterBulkBar();
        cb.closest('tr')?.classList.toggle('selected-row', cb.checked);
      });
    });
  }

  // Select all checkbox state
  const checkAll = document.getElementById('checkSelectAllMaster');
  if (checkAll) {
    if (isAdmin) {
      checkAll.disabled = false;
      checkAll.style.opacity = '1';
      checkAll.style.cursor = 'pointer';
      const allChecked = pageItems.length > 0 && pageItems.every(it => state.selectedMasterCodes.has(it.kodeMaterial));
      checkAll.checked = allChecked;
    } else {
      checkAll.disabled = true;
      checkAll.style.opacity = '0.3';
      checkAll.style.cursor = 'not-allowed';
      checkAll.checked = false;
    }
  }
}

// Update Floating Bulk Bar visibility and count
function updateMasterBulkBar() {
  const bar = document.getElementById('masterBulkBar');
  const countEl = document.getElementById('selectedMasterCount');

  if (!isAdminLoggedIn()) {
    state.selectedMasterCodes.clear();
    if (bar) bar.style.display = 'none';
    return;
  }

  const count = state.selectedMasterCodes.size;
  if (countEl) countEl.textContent = count;
  if (bar) {
    bar.style.display = count > 0 ? 'flex' : 'none';
  }
}

// Select All Handler
function handleSelectAllMaster(e) {
  if (!isAdminLoggedIn()) {
    e.target.checked = false;
    openAdminLoginModal(null, 'melakukan aksi massal');
    return;
  }

  const checked = e.target.checked;
  state.filteredMaster.forEach(m => {
    if (checked) state.selectedMasterCodes.add(m.kodeMaterial);
    else state.selectedMasterCodes.delete(m.kodeMaterial);
  });
  renderMasterTable();
}

function deselectAllMaster() {
  state.selectedMasterCodes.clear();
  renderMasterTable();
}

/**
 * ============================================================================
 * INTERACTIVE FAST EDITS (PS/NON PS, LIB/NON LIB, QTY STEPPER, INLINE EDIT)
 * ============================================================================
 */
async function togglePartService(kode) {
  const item = state.master.find(m => m.kodeMaterial === kode);
  if (!item) return;

  const currentIsPs = String(item.status || 'PS').toUpperCase().includes('PS') && !String(item.status).toUpperCase().includes('NON');
  const nextStatus = currentIsPs ? 'NON PS' : 'PS';

  try {
    await window.InventoryDB.upsertMasterMaterial({
      ...item,
      status: nextStatus
    }, 'Team');

    item.status = nextStatus;
    renderMasterTable();
    showToast(`Status ${kode} diubah ke: ${nextStatus}`, 'info');
  } catch (err) {
    showToast('Gagal mengubah status: ' + err.message, 'danger');
  }
}

async function togglePartLibCategory(kode) {
  const item = state.master.find(m => m.kodeMaterial === kode);
  if (!item) return;

  const currentIsLib = String(item.libCategory || 'NON LIB').toUpperCase() === 'LIB';
  const nextCategory = currentIsLib ? 'NON LIB' : 'LIB';

  try {
    await window.InventoryDB.upsertMasterMaterial({
      ...item,
      libCategory: nextCategory
    }, 'Team');

    item.libCategory = nextCategory;
    renderMasterTable();
    showToast(`Kategori ${kode} diubah ke: ${nextCategory}`, 'info');
  } catch (err) {
    showToast('Gagal mengubah kategori: ' + err.message, 'danger');
  }
}

async function stepPartQty(kode, delta) {
  const item = state.master.find(m => m.kodeMaterial === kode);
  if (!item) return;

  const current = Number(item.stokSaatIni) || 0;
  const nextQty = Math.max(0, current + delta);

  try {
    await window.InventoryDB.upsertMasterMaterial({
      ...item,
      stokSaatIni: nextQty,
      stokAwal: nextQty
    }, 'Team');

    item.stokSaatIni = nextQty;
    item.stokAwal = nextQty;
    renderMasterTable();
    updateDashboard();
  } catch (err) {
    showToast('Gagal update stok: ' + err.message, 'danger');
  }
}

function inlineEditQty(el, kode) {
  const currentVal = el.textContent.trim();
  const input = document.createElement('input');
  input.type = 'number';
  input.className = 'inline-edit-input';
  input.value = currentVal;
  input.min = '0';

  el.replaceWith(input);
  input.focus();
  input.select();

  const save = async () => {
    const val = parseInt(input.value, 10);
    const item = state.master.find(m => m.kodeMaterial === kode);
    if (!isNaN(val) && item && val !== item.stokSaatIni) {
      const nextQty = Math.max(0, val);
      try {
        await window.InventoryDB.upsertMasterMaterial({
          ...item,
          stokSaatIni: nextQty,
          stokAwal: nextQty
        }, 'Team');

        item.stokSaatIni = nextQty;
        item.stokAwal = nextQty;
        showToast(`Stok ${kode} diperbarui menjadi: ${nextQty}`, 'success');
      } catch (err) {
        showToast('Gagal simpan stok: ' + err.message, 'danger');
      }
    }
    renderMasterTable();
    updateDashboard();
  };

  input.addEventListener('blur', save);
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      input.removeEventListener('blur', save);
      save();
    } else if (e.key === 'Escape') {
      renderMasterTable();
    }
  });
}

/**
 * ============================================================================
 * MASS BULK ACTIONS (APPLY TO ALL SELECTED PARTS)
 * ============================================================================
 */
async function executeBulkStatusChange(targetStatus) {
  const codes = Array.from(state.selectedMasterCodes);
  if (codes.length === 0) return;

  const updates = codes.map(k => ({ kodeMaterial: k, status: targetStatus }));

  try {
    await window.InventoryDB.bulkUpdateMasterMaterials(updates, 'Team');
    codes.forEach(k => {
      const m = state.master.find(it => it.kodeMaterial === k);
      if (m) m.status = targetStatus;
    });

    showToast(`Berhasil! ${codes.length} part diubah statusnya menjadi ${targetStatus}.`, 'success');
    deselectAllMaster();
  } catch (err) {
    showToast('Gagal ubah status massal: ' + err.message, 'danger');
  }
}

async function executeBulkLibChange(targetLib) {
  const codes = Array.from(state.selectedMasterCodes);
  if (codes.length === 0) return;

  const updates = codes.map(k => ({ kodeMaterial: k, libCategory: targetLib }));

  try {
    await window.InventoryDB.bulkUpdateMasterMaterials(updates, 'Team');
    codes.forEach(k => {
      const m = state.master.find(it => it.kodeMaterial === k);
      if (m) m.libCategory = targetLib;
    });

    showToast(`Berhasil! ${codes.length} part diubah kategorinya menjadi ${targetLib}.`, 'success');
    deselectAllMaster();
  } catch (err) {
    showToast('Gagal ubah kategori massal: ' + err.message, 'danger');
  }
}

function openBulkLocationModal() {
  const count = state.selectedMasterCodes.size;
  setText('bulkLocationPrompt', `Masukkan lokasi baru untuk <strong>${count}</strong> part terpilih:`);
  document.getElementById('bulkLocationInput').value = '';
  openModal('bulkLocationModal');
}

async function handleBulkLocationSubmit(e) {
  e.preventDefault();
  const newLocation = document.getElementById('bulkLocationInput').value.trim();
  if (!newLocation) return;

  const codes = Array.from(state.selectedMasterCodes);
  const updates = codes.map(k => ({ kodeMaterial: k, lokasi: newLocation }));

  try {
    await window.InventoryDB.bulkUpdateMasterMaterials(updates, 'Team');
    codes.forEach(k => {
      const m = state.master.find(it => it.kodeMaterial === k);
      if (m) m.lokasi = newLocation;
    });

    closeModal('bulkLocationModal');
    showToast(`Lokasi ${codes.length} part berhasil diubah ke: ${newLocation}`, 'success');
    populateMasterDropdowns();
    deselectAllMaster();
  } catch (err) {
    showToast('Gagal pindah lokasi: ' + err.message, 'danger');
  }
}

function openBulkQtyModal() {
  document.getElementById('bulkQtyInput').value = '';
  openModal('bulkQtyModal');
}

async function handleBulkQtySubmit(e) {
  e.preventDefault();
  const mode = document.querySelector('input[name="qtyMode"]:checked')?.value || 'set';
  const val = Number(document.getElementById('bulkQtyInput').value);
  if (isNaN(val)) return;

  const codes = Array.from(state.selectedMasterCodes);
  const updates = codes.map(k => {
    if (mode === 'set') {
      return { kodeMaterial: k, stokSaatIni: Math.max(0, val) };
    } else {
      return { kodeMaterial: k, qtyDelta: val };
    }
  });

  try {
    const updated = await window.InventoryDB.bulkUpdateMasterMaterials(updates, 'Team');
    const updateMap = new Map();
    updated.forEach(u => updateMap.set(u.kodeMaterial, u.stokSaatIni));

    state.master.forEach(m => {
      if (updateMap.has(m.kodeMaterial)) {
        m.stokSaatIni = updateMap.get(m.kodeMaterial);
      }
    });

    closeModal('bulkQtyModal');
    showToast(`QTY ${codes.length} part berhasil diperbarui!`, 'success');
    deselectAllMaster();
    updateDashboard();
  } catch (err) {
    showToast('Gagal update QTY massal: ' + err.message, 'danger');
  }
}

function openBulkMinMaxModal() {
  document.getElementById('bulkMinInput').value = '1';
  document.getElementById('bulkMaxInput').value = '5';
  openModal('bulkMinMaxModal');
}

async function handleBulkMinMaxSubmit(e) {
  e.preventDefault();
  const min = Number(document.getElementById('bulkMinInput').value);
  const max = Number(document.getElementById('bulkMaxInput').value);

  const codes = Array.from(state.selectedMasterCodes);
  const updates = codes.map(k => ({ kodeMaterial: k, min: min, max: max }));

  try {
    await window.InventoryDB.bulkUpdateMasterMaterials(updates, 'Team');
    codes.forEach(k => {
      const m = state.master.find(it => it.kodeMaterial === k);
      if (m) { m.min = min; m.max = max; }
    });

    closeModal('bulkMinMaxModal');
    showToast(`Batas Min (${min}) & Max (${max}) berhasil diterapkan ke ${codes.length} part!`, 'success');
    deselectAllMaster();
    updateDashboard();
  } catch (err) {
    showToast('Gagal update Min/Max: ' + err.message, 'danger');
  }
}

async function handleBulkDeleteSelected() {
  const codes = Array.from(state.selectedMasterCodes);
  if (codes.length === 0) return;

  if (!confirm(`Apakah Anda yakin ingin menghapus ${codes.length} part terpilih sekaligus?\nRiwayat transaksi harian dari part-part ini juga akan dihapus.`)) {
    return;
  }

  try {
    for (const k of codes) {
      await window.InventoryDB.deleteMasterMaterial(k, 'Team');
    }
    showToast(`${codes.length} part berhasil dihapus.`, 'info');
    deselectAllMaster();
    await loadDataFromDB();
  } catch (err) {
    showToast('Gagal menghapus massal: ' + err.message, 'danger');
  }
}

/**
 * ============================================================================
 * SINGLE MASTER CRUD (MODAL)
 * ============================================================================
 */
function openAddMasterModal() {
  requireAdmin('menambah master material baru', () => {
    document.getElementById('partForm').reset();
    document.getElementById('editOriginalKode').value = '';
    document.getElementById('formKodeMaterial').readOnly = false;
    setText('partModalTitle', '<i class="fa-solid fa-plus-circle"></i> Tambah Master Material Baru');
    openModal('partModal');
  });
}

async function openEditMasterModal(kode) {
  requireAdmin('mengedit master material', async () => {
    const item = await window.InventoryDB.getMasterMaterial(kode);
    if (!item) {
      showToast('Material tidak ditemukan.', 'danger');
      return;
    }

    document.getElementById('editOriginalKode').value = item.kodeMaterial;
    const kodeInput = document.getElementById('formKodeMaterial');
    kodeInput.value = item.kodeMaterial;
    kodeInput.readOnly = true;

    document.getElementById('formLokasi').value = item.lokasi || '';
    document.getElementById('formNamaMaterial').value = item.namaMaterial || '';
    document.getElementById('formStatus').value = item.status || 'PS';
    document.getElementById('formLibCategory').value = item.libCategory || 'NON LIB';
    document.getElementById('formPrice').value = item.price || 0;
    document.getElementById('formStokAwal').value = item.stokSaatIni !== undefined ? item.stokSaatIni : 0;
    document.getElementById('formMin').value = item.min !== undefined ? item.min : 1;
    document.getElementById('formMax').value = item.max !== undefined ? item.max : 5;
    document.getElementById('formProd').value = item.prod || '';
    document.getElementById('formModelUnit').value = item.modelUnit || '';
    document.getElementById('formRemarks').value = item.remarks || '';

    setText('partModalTitle', `<i class="fa-solid fa-pen-to-square"></i> Edit Master Material (${item.kodeMaterial})`);
    openModal('partModal');
  });
}

async function handleMasterFormSubmit(e) {
  e.preventDefault();
  if (!isAdminLoggedIn()) {
    openAdminLoginModal(null, 'menyimpan master material');
    return;
  }

  const kode = document.getElementById('formKodeMaterial').value.trim();
  const originalKode = document.getElementById('editOriginalKode').value.trim();

  if (!kode) {
    showToast('Kode Material wajib diisi!', 'warning');
    return;
  }

  if (!originalKode) {
    const existing = await window.InventoryDB.getMasterMaterial(kode);
    if (existing) {
      showToast(`Kode Material "${kode}" sudah terdaftar! Sistem tidak mengizinkan duplikasi.`, 'warning');
      return;
    }
  }

  const materialData = {
    kodeMaterial: kode,
    lokasi: document.getElementById('formLokasi').value.trim() || 'GUDANG UTAMA',
    namaMaterial: document.getElementById('formNamaMaterial').value.trim(),
    status: document.getElementById('formStatus').value,
    libCategory: document.getElementById('formLibCategory').value,
    price: Number(document.getElementById('formPrice').value) || 0,
    stokAwal: Number(document.getElementById('formStokAwal').value) || 0,
    stokSaatIni: Number(document.getElementById('formStokAwal').value) || 0,
    min: Number(document.getElementById('formMin').value) || 1,
    max: Number(document.getElementById('formMax').value) || 5,
    prod: document.getElementById('formProd').value.trim() || 'GENERAL',
    modelUnit: document.getElementById('formModelUnit').value.trim(),
    remarks: document.getElementById('formRemarks').value.trim()
  };

  try {
    await window.InventoryDB.upsertMasterMaterial(materialData, 'Team');
    closeModal('partModal');
    showToast(`Master material ${kode} berhasil disimpan!`, 'success');
    await loadDataFromDB();
  } catch (err) {
    showToast('Gagal menyimpan master material: ' + err.message, 'danger');
  }
}

function confirmDeleteMaster(kode) {
  requireAdmin('menghapus master material', () => {
    document.getElementById('confirmDeleteMessage').innerHTML = `
      Apakah Anda yakin ingin menghapus material <strong>${escapeHtml(kode)}</strong>?<br>
      <span style="color: var(--danger); font-size: 0.85rem;">Seluruh riwayat transaksi harian material ini juga akan dihapus.</span>
    `;
    const btnExecute = document.getElementById('btnExecuteDelete');
    btnExecute.onclick = async () => {
      try {
        await window.InventoryDB.deleteMasterMaterial(kode, 'Team');
        closeModal('confirmDeleteModal');
        showToast(`Material ${kode} berhasil dihapus.`, 'info');
        await loadDataFromDB();
      } catch (err) {
        showToast('Gagal menghapus material: ' + err.message, 'danger');
      }
    };
    openModal('confirmDeleteModal');
  });
}

/**
 * ============================================================================
 * MODULE 3: DAILY STOCK UPDATE & CASCADE RECALCULATION
 * ============================================================================
 */
function quickDailyForMaterial(kode) {
  switchAppView('daily');
  const kodeInput = document.getElementById('dailyFormKode');
  if (kodeInput) {
    kodeInput.value = kode;
    syncDailyFormMaterialName();
    setTimeout(() => {
      document.getElementById('dailyFormIn').focus();
    }, 150);
  }
}

function syncDailyFormMaterialName() {
  const kode = (document.getElementById('dailyFormKode').value || '').trim();
  const namaInput = document.getElementById('dailyFormNama');
  if (!kode || !namaInput) return;

  const mat = state.master.find(m => m.kodeMaterial.toLowerCase() === kode.toLowerCase());
  if (mat) {
    namaInput.value = `${mat.namaMaterial} [${mat.status || 'PS'}] [${mat.libCategory || 'NON LIB'}]`;
  } else {
    namaInput.value = '';
  }
}

async function handleDailyQuickSubmit(e) {
  e.preventDefault();
  const tanggal = document.getElementById('dailyFormDate').value;
  const kode = document.getElementById('dailyFormKode').value.trim();
  const inQty = Number(document.getElementById('dailyFormIn').value) || 0;
  const outQty = Number(document.getElementById('dailyFormOut').value) || 0;
  const utd = document.getElementById('dailyFormUtd').value.trim();

  if (!tanggal) {
    showToast('Pilih tanggal transaksi terlebih dahulu.', 'warning');
    return;
  }
  if (!kode) {
    showToast('Kode Material wajib diisi.', 'warning');
    return;
  }

  const mat = await window.InventoryDB.getMasterMaterial(kode);
  if (!mat) {
    showToast(`Material ${kode} belum terdaftar di Master Data. Silakan daftarkan di Master Data atau Bulk Import.`, 'warning');
    return;
  }

  try {
    const result = await window.InventoryDB.upsertDailyStock({
      tanggal,
      kodeMaterial: mat.kodeMaterial,
      inQty,
      outQty,
      utd,
      notes: 'Input Daily Mutasi'
    }, 'Team');

    showToast(`Transaksi ${mat.kodeMaterial} tanggal ${tanggal} disimpan! Saldo SISA: ${result.record.sisa}`, 'success');
    
    document.getElementById('dailyFormIn').value = '0';
    document.getElementById('dailyFormOut').value = '0';
    document.getElementById('dailyFormUtd').value = '';

    await loadDataFromDB();
  } catch (err) {
    showToast('Gagal menyimpan transaksi: ' + err.message, 'danger');
  }
}

function applyDailyFilters() {
  const { search, date, month, year, hour } = state.dailyFilters;
  const q = search.toLowerCase().trim();

  state.filteredDaily = state.daily.filter(t => {
    if (q) {
      const match = (t.kodeMaterial && t.kodeMaterial.toLowerCase().includes(q)) ||
                    (t.namaMaterial && t.namaMaterial.toLowerCase().includes(q)) ||
                    (t.notes && t.notes.toLowerCase().includes(q));
      if (!match) return false;
    }

    if (date && t.tanggal !== date) return false;

    if (t.tanggal) {
      const parts = t.tanggal.split('-');
      if (year !== 'ALL' && parts[0] !== year) return false;
      if (month !== 'ALL' && parseInt(parts[1], 10) !== parseInt(month, 10)) return false;
    }

    if (hour !== 'ALL' && t.jam) {
      const hStr = t.jam.substring(0, 2);
      if (hStr !== hour) return false;
    }

    return true;
  });

  state.filteredDaily.sort((a, b) => {
    if (a.tanggal === b.tanggal) {
      return (b.jam || '').localeCompare(a.jam || '');
    }
    return b.tanggal.localeCompare(a.tanggal);
  });

  state.dailyPage = 1;
  renderDailyTable();
}

function renderDailyTable() {
  const tbody = document.getElementById('dailyTableBody');
  if (!tbody) return;

  const total = state.filteredDaily.length;
  const pageSize = state.dailyPageSize === 'ALL' ? total : Number(state.dailyPageSize);
  const totalPages = Math.ceil(total / pageSize) || 1;

  if (state.dailyPage > totalPages) state.dailyPage = totalPages;

  const start = (state.dailyPage - 1) * pageSize;
  const end = Math.min(start + pageSize, total);
  const pageItems = state.filteredDaily.slice(start, end);

  if (pageItems.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="11" style="text-align: center; padding: 2.5rem; color: var(--text-muted);">
          <i class="fa-solid fa-calendar-xmark" style="font-size: 2rem; margin-bottom: 0.5rem; display: block;"></i>
          Belum ada transaksi harian yang sesuai filter.
        </td>
      </tr>
    `;
    setText('dailyPaginationInfo', 'Menampilkan 0 transaksi');
    document.getElementById('dailyPaginationControls').innerHTML = '';
    return;
  }

  const isAdmin = isAdminLoggedIn();

  tbody.innerHTML = pageItems.map((item, idx) => {
    const actionsHtml = isAdmin
      ? `<div style="display: inline-flex; gap: 0.35rem;">
          <button class="btn btn-icon btn-sm" onclick="openEditDailyModal('${escapeQuote(item.id)}')" title="Edit Transaksi (Hitung Ulang Saldo)">
            <i class="fa-solid fa-pen-to-square" style="color: #38bdf8;"></i>
          </button>
          <button class="btn btn-icon btn-sm" onclick="confirmDeleteDaily('${escapeQuote(item.id)}')" title="Hapus Transaksi">
            <i class="fa-solid fa-trash" style="color: var(--danger);"></i>
          </button>
        </div>`
      : `<span class="badge badge-guest" style="cursor: pointer;" onclick="openAdminLoginModal(null, 'mengedit transaksi ini')" title="Klik untuk Login Admin">
          <i class="fa-solid fa-eye"></i> Hanya Lihat
        </span>`;

    return `
      <tr>
        <td style="color: var(--text-muted); font-size: 0.85rem;">${start + idx + 1}</td>
        <td>
          <span style="font-weight: 600; color: var(--text-primary);">${item.tanggal}</span>
        </td>
        <td style="font-size: 0.82rem; color: var(--text-muted);">${item.jam || '-'}</td>
        <td>
          <strong style="font-family: 'JetBrains Mono', monospace; color: #93c5fd;">${escapeHtml(item.kodeMaterial)}</strong>
        </td>
        <td>
          <div style="font-weight: 500; color: var(--text-primary);">${escapeHtml(item.namaMaterial || '-')}</div>
        </td>
        <td>
          <span class="badge" style="background: rgba(148, 163, 184, 0.12);">${escapeHtml(item.lokasi || '-')}</span>
        </td>
        <td>
          <span class="badge-in">+${item.inQty || 0}</span>
        </td>
        <td>
          <span class="badge-out">-${item.outQty || 0}</span>
        </td>
        <td style="color: var(--text-secondary); font-size: 0.85rem;">
          ${escapeHtml(item.utd || '-')}
        </td>
        <td>
          <span class="badge-sisa">${item.sisa !== undefined ? item.sisa : 0}</span>
        </td>
        <td style="text-align: center;">${actionsHtml}</td>
      </tr>
    `;
  }).join('');

  setText('dailyPaginationInfo', `Menampilkan ${start + 1} - ${end} dari ${total} transaksi`);
  renderPaginationControls('dailyPaginationControls', state.dailyPage, totalPages, (p) => {
    state.dailyPage = p;
    renderDailyTable();
  });
}

async function openEditDailyModal(txId) {
  requireAdmin('mengedit transaksi harian', async () => {
    const tx = await window.InventoryDB.getDailyStockById(txId);
    if (!tx) {
      showToast('Transaksi tidak ditemukan.', 'danger');
      return;
    }

    document.getElementById('editDailyTxId').value = tx.id;
    document.getElementById('editDailyKode').value = tx.kodeMaterial;
    document.getElementById('editDailyNama').value = tx.namaMaterial || '';
    document.getElementById('editDailyTanggal').value = tx.tanggal;
    document.getElementById('editDailyJam').value = tx.jam || '';
    document.getElementById('editDailyIn').value = tx.inQty !== undefined ? tx.inQty : 0;
    document.getElementById('editDailyOut').value = tx.outQty !== undefined ? tx.outQty : 0;
    document.getElementById('editDailyUtd').value = tx.utd || '';
    document.getElementById('editDailyNotes').value = tx.notes || '';

    openModal('editDailyModal');
  });
}

async function handleEditDailySubmit(e) {
  e.preventDefault();
  if (!isAdminLoggedIn()) {
    openAdminLoginModal(null, 'mengedit transaksi harian');
    return;
  }

  const txId = document.getElementById('editDailyTxId').value;
  const kode = document.getElementById('editDailyKode').value;
  const tanggal = document.getElementById('editDailyTanggal').value;
  const jam = document.getElementById('editDailyJam').value;
  const inQty = Number(document.getElementById('editDailyIn').value) || 0;
  const outQty = Number(document.getElementById('editDailyOut').value) || 0;
  const utd = document.getElementById('editDailyUtd').value.trim();
  const notes = document.getElementById('editDailyNotes').value.trim();

  try {
    const result = await window.InventoryDB.upsertDailyStock({
      id: txId,
      kodeMaterial: kode,
      tanggal,
      jam,
      inQty,
      outQty,
      utd,
      notes
    }, 'Team');

    closeModal('editDailyModal');
    showToast(`Transaksi diperbarui! Saldo SISA material ${kode} telah dihitung ulang: ${result.record.sisa}`, 'success');
    await loadDataFromDB();
  } catch (err) {
    showToast('Gagal memperbarui transaksi: ' + err.message, 'danger');
  }
}

async function confirmDeleteDaily(txId) {
  requireAdmin('menghapus transaksi harian', async () => {
    const tx = await window.InventoryDB.getDailyStockById(txId);
    if (!tx) return;

    document.getElementById('confirmDeleteMessage').innerHTML = `
      Apakah Anda yakin ingin menghapus transaksi tanggal <strong>${tx.tanggal}</strong> untuk material <strong>${escapeHtml(tx.kodeMaterial)}</strong>?<br>
      <span style="color: #93c5fd; font-size: 0.85rem;">Saldo stok akan otomatis dihitung ulang kembali dari stok awal.</span>
    `;
    const btnExecute = document.getElementById('btnExecuteDelete');
    btnExecute.onclick = async () => {
      try {
        await window.InventoryDB.deleteDailyStock(txId, 'Team');
        closeModal('confirmDeleteModal');
        showToast('Transaksi dihapus dan saldo berhasil dihitung ulang.', 'info');
        await loadDataFromDB();
      } catch (err) {
        showToast('Gagal menghapus transaksi: ' + err.message, 'danger');
      }
    };
    openModal('confirmDeleteModal');
  });
}

/**
 * ============================================================================
 * MODULE 4: UNIVERSAL BULK DATA IMPORT (EXCEL COPY-PASTE MASSAL)
 * ============================================================================
 */
async function pasteFromClipboard() {
  try {
    const text = await navigator.clipboard.readText();
    if (!text) {
      showToast('Clipboard kosong atau tidak berisi teks.', 'warning');
      return;
    }
    const area = document.getElementById('excelPasteArea');
    area.value = text;
    showToast('Data dari clipboard berhasil ditempel! Klik tombol Proses & Tampilkan Preview.', 'info');
  } catch (err) {
    showToast('Gagal membaca clipboard browser: ' + err.message, 'warning');
  }
}

function parseExcelText(rawText) {
  const lines = rawText.trim().split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length === 0) return { error: 'Teks kosong' };

  const firstLine = lines[0];
  let delimiter = '\t';
  if (firstLine.includes('\t')) delimiter = '\t';
  else if (firstLine.includes('|')) delimiter = '|';
  else if (firstLine.includes(';') && !firstLine.includes(',')) delimiter = ';';
  else if (firstLine.includes(',')) delimiter = ',';

  const rows = lines.map(line => line.split(delimiter).map(cell => cell.trim().replace(/^["']|["']$/g, '')));

  const headerMap = {};
  const firstRow = rows[0].map(c => c.toUpperCase());
  let hasHeader = false;

  firstRow.forEach((col, idx) => {
    if (col.includes('KODE') || col.includes('PART NUMBER') || col === 'PN' || col.includes('ITEM CODE') || col.includes('FINAL PN')) {
      headerMap.kodeMaterial = idx;
      hasHeader = true;
    } else if (col.includes('LOKASI') || col.includes('LOCATION') || col.includes('PALET') || col.includes('RAK')) {
      headerMap.lokasi = idx;
      hasHeader = true;
    } else if (col.includes('NAMA') || col.includes('DESC') || col.includes('DESKRIPSI') || col.includes('DESCRIPTION')) {
      headerMap.namaMaterial = idx;
      hasHeader = true;
    } else if (col === 'QTY' || col.includes('STOK') || col.includes('STOCK') || col === 'JUMLAH') {
      headerMap.qty = idx;
      hasHeader = true;
    } else if (col === 'MIN' || col.includes('MINIMUM')) {
      headerMap.min = idx;
      hasHeader = true;
    } else if (col === 'MAX' || col.includes('MAKSIMUM')) {
      headerMap.max = idx;
      hasHeader = true;
    } else if (col === 'PS' || col.includes('SERVIS') || col === 'STATUS') {
      headerMap.status = idx;
      hasHeader = true;
    } else if (col === 'LIB' || col.includes('KATEGORI') || col.includes('CATEGORY')) {
      headerMap.libCategory = idx;
      hasHeader = true;
    } else if (col.includes('PRICE') || col.includes('HARGA') || col.includes('COST')) {
      headerMap.price = idx;
      hasHeader = true;
    } else if (col === 'IN' || col.includes('MASUK') || col.includes('QTY IN')) {
      headerMap.inQty = idx;
      hasHeader = true;
    } else if (col === 'OUT' || col.includes('KELUAR') || col.includes('QTY OUT')) {
      headerMap.outQty = idx;
      hasHeader = true;
    } else if (col === 'UTD' || col.includes('UNIT TO DATE') || col.includes('USAGE')) {
      headerMap.utd = idx;
      hasHeader = true;
    } else if (col.includes('TANGGAL') || col.includes('DATE') || col === 'TGL') {
      headerMap.tanggal = idx;
      hasHeader = true;
    }
  });

  if (!hasHeader) {
    headerMap.kodeMaterial = 0;
    headerMap.namaMaterial = 1;
    headerMap.lokasi = 2;
    headerMap.qty = 3;
    headerMap.min = 4;
    headerMap.max = 5;
    headerMap.status = 6;
    headerMap.libCategory = 7;
  }

  const cleanNum = (val) => {
    if (val === undefined || val === null || val === '') return null;
    const cleaned = String(val).replace(/[^0-9.-]/g, '');
    const parsed = parseFloat(cleaned);
    return isNaN(parsed) ? null : parsed;
  };

  const dataRows = hasHeader ? rows.slice(1) : rows;
  const parsedItems = dataRows.map((row, rIdx) => {
    const kode = (headerMap.kodeMaterial !== undefined && row[headerMap.kodeMaterial]) ? row[headerMap.kodeMaterial].trim() : '';
    const nama = (headerMap.namaMaterial !== undefined && row[headerMap.namaMaterial]) ? row[headerMap.namaMaterial].trim() : '';
    const lokasi = (headerMap.lokasi !== undefined && row[headerMap.lokasi]) ? row[headerMap.lokasi].trim() : '';

    const qty = cleanNum(headerMap.qty !== undefined ? row[headerMap.qty] : null);
    const min = cleanNum(headerMap.min !== undefined ? row[headerMap.min] : null);
    const max = cleanNum(headerMap.max !== undefined ? row[headerMap.max] : null);
    const price = cleanNum(headerMap.price !== undefined ? row[headerMap.price] : null) || 0;
    const inQty = cleanNum(headerMap.inQty !== undefined ? row[headerMap.inQty] : null) || 0;
    const outQty = cleanNum(headerMap.outQty !== undefined ? row[headerMap.outQty] : null) || 0;

    let status = headerMap.status !== undefined && row[headerMap.status] ? row[headerMap.status].trim().toUpperCase() : null;
    if (status) {
      if (status.includes('NON')) status = 'NON PS';
      else if (status.includes('PS')) status = 'PS';
    }

    let lib = headerMap.libCategory !== undefined && row[headerMap.libCategory] ? row[headerMap.libCategory].trim().toUpperCase() : null;
    if (lib) {
      if (lib.includes('NON')) lib = 'NON LIB';
      else if (lib.includes('LIB')) lib = 'LIB';
    }

    const utd = (headerMap.utd !== undefined && row[headerMap.utd] !== undefined) ? row[headerMap.utd].trim() : '';
    const tanggal = (headerMap.tanggal !== undefined && row[headerMap.tanggal]) ? row[headerMap.tanggal].trim() : null;

    return {
      rowIndex: rIdx + 1,
      kodeMaterial: kode,
      namaMaterial: nama,
      lokasi: lokasi,
      qty: qty,
      min: min,
      max: max,
      status: status,
      libCategory: lib,
      price: price,
      inQty: inQty,
      outQty: outQty,
      utd: utd,
      tanggal: tanggal,
      hasTransaction: Boolean(inQty > 0 || outQty > 0 || utd),
      isValid: Boolean(kode)
    };
  });

  return { hasHeader, headerMap, items: parsedItems };
}

async function processImportPreview() {
  const area = document.getElementById('excelPasteArea');
  const text = (area.value || '').trim();

  if (!text) {
    showToast('Tempelkan teks data dari Excel terlebih dahulu!', 'warning');
    return;
  }

  const parseResult = parseExcelText(text);
  if (parseResult.error || !parseResult.items || parseResult.items.length === 0) {
    showToast('Gagal mengenali data: ' + (parseResult.error || 'Format tidak valid'), 'danger');
    return;
  }

  const defaultDate = document.getElementById('importDefaultDate').value || new Date().toISOString().split('T')[0];

  const masterMap = new Map();
  state.master.forEach(m => masterMap.set(m.kodeMaterial.toLowerCase(), m));

  const seenInBatch = new Set();
  let cntValid = 0, cntNew = 0, cntExisting = 0, cntError = 0;

  const previewItems = parseResult.items.map(item => {
    if (!item.isValid) {
      cntError++;
      return {
        ...item,
        statusLabel: 'GAGAL',
        badgeClass: 'error',
        note: 'Kode material (PN) kosong'
      };
    }

    const mat = masterMap.get(item.kodeMaterial.toLowerCase());
    const isDuplicateInBatch = seenInBatch.has(item.kodeMaterial.toLowerCase());
    seenInBatch.add(item.kodeMaterial.toLowerCase());

    if (isDuplicateInBatch) {
      return {
        ...item,
        namaMaterial: item.namaMaterial || (mat ? mat.namaMaterial : 'Part ' + item.kodeMaterial),
        statusLabel: 'DUPLIKAT DI BATCH',
        badgeClass: 'duplicate',
        note: 'Kode muncul lebih dari 1x di tabel paste'
      };
    }

    if (mat) {
      cntExisting++;
      cntValid++;
      return {
        ...item,
        namaMaterial: item.namaMaterial || mat.namaMaterial,
        lokasi: item.lokasi || mat.lokasi,
        qty: item.qty !== null ? item.qty : mat.stokSaatIni,
        min: item.min !== null ? item.min : mat.min,
        max: item.max !== null ? item.max : mat.max,
        status: item.status || mat.status,
        libCategory: item.libCategory || mat.libCategory,
        statusLabel: 'UPDATE MASTER',
        badgeClass: 'existing',
        note: 'Material terdaftar -> kolom yang diisi akan diupdate massal'
      };
    } else {
      cntNew++;
      cntValid++;
      return {
        ...item,
        namaMaterial: item.namaMaterial || ('Spare Part ' + item.kodeMaterial),
        lokasi: item.lokasi || 'GUDANG UTAMA',
        qty: item.qty !== null ? item.qty : 0,
        min: item.min !== null ? item.min : 1,
        max: item.max !== null ? item.max : 5,
        status: item.status || 'PS',
        libCategory: item.libCategory || 'NON LIB',
        statusLabel: 'MATERIAL BARU',
        badgeClass: 'new',
        note: 'Part baru -> didaftarkan otomatis ke Master Data'
      };
    }
  });

  state.parsedImport = {
    items: previewItems,
    defaultDate: defaultDate
  };

  setText('prevMetricTotal', previewItems.length);
  setText('prevMetricNew', cntNew);
  setText('prevMetricExisting', cntExisting);
  setText('prevMetricValid', cntValid);
  setText('prevMetricError', cntError);

  const tbody = document.getElementById('importPreviewTableBody');
  tbody.innerHTML = previewItems.map((it, idx) => {
    return `
      <tr>
        <td style="color: var(--text-muted); font-size: 0.85rem;">${idx + 1}</td>
        <td>
          <span class="badge-status ${it.badgeClass}">
            ${it.statusLabel}
          </span>
        </td>
        <td>
          <strong style="font-family: 'JetBrains Mono', monospace; color: #93c5fd;">${escapeHtml(it.kodeMaterial || '-')}</strong>
        </td>
        <td>${escapeHtml(it.namaMaterial || '-')}</td>
        <td><span class="badge" style="background: rgba(148, 163, 184, 0.15);">${escapeHtml(it.lokasi || '-')}</span></td>
        <td><strong style="color: #34d399; font-size: 1rem;">${it.qty !== null ? it.qty : '-'}</strong></td>
        <td style="font-size: 0.85rem;">${it.min !== null ? it.min : 1} - ${it.max !== null ? it.max : 5}</td>
        <td><span class="badge" style="background: rgba(16, 185, 129, 0.15); color: #34d399;">${it.status || 'PS'}</span></td>
        <td><span class="badge" style="background: rgba(6, 182, 212, 0.15); color: #22d3ee;">${it.libCategory || 'NON LIB'}</span></td>
        <td style="font-size: 0.85rem;">+${it.inQty || 0} / -${it.outQty || 0}</td>
        <td style="font-size: 0.8rem; color: var(--text-secondary);">${escapeHtml(it.note || '-')}</td>
      </tr>
    `;
  }).join('');

  document.getElementById('importPreviewContainer').style.display = 'block';
  document.getElementById('importPreviewContainer').scrollIntoView({ behavior: 'smooth' });
}

function cancelImportPreview() {
  state.parsedImport = null;
  document.getElementById('importPreviewContainer').style.display = 'none';
  showToast('Proses import dibatalkan. Tidak ada data yang diubah.', 'info');
}

async function confirmAndExecuteImport() {
  if (!isAdminLoggedIn()) {
    openAdminLoginModal(confirmAndExecuteImport, 'menyimpan pembaruan massal');
    return;
  }

  if (!state.parsedImport || !state.parsedImport.items || state.parsedImport.items.length === 0) {
    showToast('Tidak ada data preview untuk diimport.', 'warning');
    return;
  }

  const validRows = state.parsedImport.items.filter(it => it.isValid);
  if (validRows.length === 0) {
    showToast('Tidak ada baris data yang valid untuk disimpan.', 'warning');
    return;
  }

  const btnConfirm = document.getElementById('btnConfirmImport');
  btnConfirm.disabled = true;
  btnConfirm.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Menyimpan Pembaruan Massal...';

  try {
    const result = await window.InventoryDB.processBulkImport(
      validRows,
      state.parsedImport.defaultDate,
      'Team'
    );

    btnConfirm.disabled = false;
    btnConfirm.innerHTML = '<i class="fa-solid fa-cloud-arrow-up"></i> Konfirmasi & Simpan Pembaruan Massal';

    document.getElementById('excelPasteArea').value = '';
    document.getElementById('importPreviewContainer').style.display = 'none';
    state.parsedImport = null;

    showToast(`Pembaruan Massal Berhasil! ${result.totalRows} baris diproses (+${result.insertedMaterials} material baru, ~${result.updatedMaterials} master update, ${result.insertedTx + result.updatedTx} mutasi dicatat).`, 'success');

    await loadDataFromDB();
    switchAppView('master');
  } catch (err) {
    btnConfirm.disabled = false;
    btnConfirm.innerHTML = '<i class="fa-solid fa-cloud-arrow-up"></i> Konfirmasi & Simpan Pembaruan Massal';
    showToast('Gagal memproses import massal: ' + err.message, 'danger');
  }
}

/**
 * ============================================================================
 * MODULE 5: STOCK HISTORY & AUDIT LOG
 * ============================================================================
 */
function applyHistoryFilters() {
  const { search, date, hour, lokasi } = state.historyFilters;
  const q = search.toLowerCase().trim();

  if (state.historyTab === 'ledger') {
    state.filteredDaily = state.daily.filter(t => {
      if (q) {
        const match = (t.kodeMaterial && t.kodeMaterial.toLowerCase().includes(q)) ||
                      (t.namaMaterial && t.namaMaterial.toLowerCase().includes(q));
        if (!match) return false;
      }
      if (date && t.tanggal !== date) return false;
      if (hour !== 'ALL' && t.jam && t.jam.substring(0, 2) !== hour) return false;
      if (lokasi !== 'ALL' && t.lokasi !== lokasi) return false;
      return true;
    });

    state.filteredDaily.sort((a, b) => {
      if (a.tanggal === b.tanggal) return (b.jam || '').localeCompare(a.jam || '');
      return b.tanggal.localeCompare(a.tanggal);
    });
  } else {
    state.filteredAudit = state.audit.filter(a => {
      if (q) {
        const match = (a.action && a.action.toLowerCase().includes(q)) ||
                      (a.user && a.user.toLowerCase().includes(q)) ||
                      (a.kodeMaterial && a.kodeMaterial.toLowerCase().includes(q)) ||
                      (a.details && a.details.toLowerCase().includes(q));
        if (!match) return false;
      }
      if (date && a.timestamp && !a.timestamp.startsWith(date)) return false;
      if (hour !== 'ALL' && a.timestamp) {
        const timePart = a.timestamp.split('T')[1] || '';
        if (timePart.substring(0, 2) !== hour) return false;
      }
      return true;
    });
  }

  state.historyPage = 1;
  renderHistoryTable();
}

function renderHistoryTable() {
  const thead = document.getElementById('historyTableHead');
  const tbody = document.getElementById('historyTableBody');
  if (!thead || !tbody) return;

  if (state.historyTab === 'ledger') {
    thead.innerHTML = `
      <tr>
        <th style="width: 50px;">No</th>
        <th>Tanggal & Jam</th>
        <th>Kode Material</th>
        <th>Nama Material</th>
        <th>Lokasi</th>
        <th>IN</th>
        <th>OUT</th>
        <th>UTD</th>
        <th>SISA Saldo</th>
        <th>Catatan</th>
        <th>Updated By</th>
      </tr>
    `;

    const total = state.filteredDaily.length;
    const pageSize = state.historyPageSize;
    const totalPages = Math.ceil(total / pageSize) || 1;
    const start = (state.historyPage - 1) * pageSize;
    const end = Math.min(start + pageSize, total);
    const pageItems = state.filteredDaily.slice(start, end);

    if (pageItems.length === 0) {
      tbody.innerHTML = `<tr><td colspan="11" style="text-align: center; padding: 2rem; color: var(--text-muted);">Tidak ada riwayat transaksi.</td></tr>`;
      setText('historyPaginationInfo', 'Menampilkan 0 riwayat');
      document.getElementById('historyPaginationControls').innerHTML = '';
      return;
    }

    tbody.innerHTML = pageItems.map((item, idx) => `
      <tr>
        <td style="color: var(--text-muted); font-size: 0.85rem;">${start + idx + 1}</td>
        <td><strong>${item.tanggal}</strong> <span style="font-size: 0.8rem; color: var(--text-muted);">${item.jam || ''}</span></td>
        <td style="font-family: 'JetBrains Mono', monospace; color: #93c5fd;">${escapeHtml(item.kodeMaterial)}</td>
        <td>${escapeHtml(item.namaMaterial || '-')}</td>
        <td><span class="badge" style="background: rgba(148, 163, 184, 0.15);">${escapeHtml(item.lokasi || '-')}</span></td>
        <td><span class="badge-in">+${item.inQty || 0}</span></td>
        <td><span class="badge-out">-${item.outQty || 0}</span></td>
        <td style="color: var(--text-secondary);">${escapeHtml(item.utd || '-')}</td>
        <td><span class="badge-sisa">${item.sisa !== undefined ? item.sisa : 0}</span></td>
        <td style="font-size: 0.82rem; color: var(--text-muted);">${escapeHtml(item.notes || '-')}</td>
        <td style="font-size: 0.82rem; color: var(--text-secondary);">${escapeHtml(item.updatedBy || 'Team')}</td>
      </tr>
    `).join('');

    setText('historyPaginationInfo', `Menampilkan ${start + 1} - ${end} dari ${total} riwayat`);
    renderPaginationControls('historyPaginationControls', state.historyPage, totalPages, p => {
      state.historyPage = p;
      renderHistoryTable();
    });

  } else {
    thead.innerHTML = `
      <tr>
        <th style="width: 50px;">No</th>
        <th>Waktu Audit</th>
        <th>User</th>
        <th>Aksi</th>
        <th>Target Material</th>
        <th>Detail Riwayat Perubahan</th>
      </tr>
    `;

    const total = state.filteredAudit.length;
    const pageSize = state.historyPageSize;
    const totalPages = Math.ceil(total / pageSize) || 1;
    const start = (state.historyPage - 1) * pageSize;
    const end = Math.min(start + pageSize, total);
    const pageItems = state.filteredAudit.slice(start, end);

    if (pageItems.length === 0) {
      tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; padding: 2rem; color: var(--text-muted);">Tidak ada audit log.</td></tr>`;
      setText('historyPaginationInfo', 'Menampilkan 0 log');
      document.getElementById('historyPaginationControls').innerHTML = '';
      return;
    }

    tbody.innerHTML = pageItems.map((log, idx) => {
      const dateFormatted = log.timestamp ? new Date(log.timestamp).toLocaleString('id-ID') : '-';
      return `
        <tr>
          <td style="color: var(--text-muted); font-size: 0.85rem;">${start + idx + 1}</td>
          <td style="font-size: 0.82rem; white-space: nowrap; color: var(--text-secondary);">${dateFormatted}</td>
          <td>
            <span class="badge badge-admin" style="font-size: 0.75rem;">
              <i class="fa-solid fa-user-check"></i> ${escapeHtml(log.user || 'Team')}
            </span>
          </td>
          <td>
            <span class="audit-action-tag ${log.action}">
              ${escapeHtml(log.action)}
            </span>
          </td>
          <td style="font-family: 'JetBrains Mono', monospace; font-size: 0.85rem; color: #93c5fd;">
            ${escapeHtml(log.kodeMaterial || 'ALL')}
          </td>
          <td style="font-size: 0.85rem; color: var(--text-primary);">
            ${escapeHtml(log.details || '-')}
          </td>
        </tr>
      `;
    }).join('');

    setText('historyPaginationInfo', `Menampilkan ${start + 1} - ${end} dari ${total} log`);
    renderPaginationControls('historyPaginationControls', state.historyPage, totalPages, p => {
      state.historyPage = p;
      renderHistoryTable();
    });
  }
}

/**
 * ============================================================================
 * MODULE 6: REPORTS & VALUATION
 * ============================================================================
 */
async function generateReport() {
  const targetDate = document.getElementById('reportTargetDate').value || new Date().toISOString().split('T')[0];
  state.reportTargetDate = targetDate;
  setText('reportDateIndicator', `Menampilkan data posisi stok per: <strong>${targetDate}</strong>`);

  const reportItems = await window.InventoryDB.getStockAsOfDate(targetDate);
  state.reportData = reportItems;
  state.reportPage = 1;
  renderReportTable();
}

function renderReportTable() {
  const thead = document.getElementById('reportTableHead');
  const tbody = document.getElementById('reportTableBody');
  if (!thead || !tbody) return;

  if (state.reportType === 'as-of-date') {
    thead.innerHTML = `
      <tr>
        <th style="width: 50px;">No</th>
        <th>Kode Material</th>
        <th>Lokasi</th>
        <th>Nama Material</th>
        <th>Servis</th>
        <th>Kategori</th>
        <th>Harga Satuan (Rp)</th>
        <th>Posisi Stok pada Tanggal</th>
        <th>Total Nilai Aset (Rp)</th>
      </tr>
    `;

    const total = state.reportData.length;
    const pageSize = state.reportPageSize;
    const totalPages = Math.ceil(total / pageSize) || 1;
    const start = (state.reportPage - 1) * pageSize;
    const end = Math.min(start + pageSize, total);
    const pageItems = state.reportData.slice(start, end);

    if (pageItems.length === 0) {
      tbody.innerHTML = `<tr><td colspan="9" style="text-align: center; padding: 2rem;">Tidak ada data laporan.</td></tr>`;
      setText('reportPaginationInfo', '0 data');
      document.getElementById('reportPaginationControls').innerHTML = '';
      return;
    }

    tbody.innerHTML = pageItems.map((item, idx) => `
      <tr>
        <td style="color: var(--text-muted); font-size: 0.85rem;">${start + idx + 1}</td>
        <td style="font-family: 'JetBrains Mono', monospace; color: #93c5fd;">${escapeHtml(item.kodeMaterial)}</td>
        <td><span class="badge" style="background: rgba(148, 163, 184, 0.15);">${escapeHtml(item.lokasi || '-')}</span></td>
        <td><strong>${escapeHtml(item.namaMaterial)}</strong></td>
        <td><span class="badge" style="background: rgba(16, 185, 129, 0.15); color: #34d399;">${item.status || 'PS'}</span></td>
        <td><span class="badge" style="background: rgba(6, 182, 212, 0.15); color: #22d3ee;">${item.libCategory || 'NON LIB'}</span></td>
        <td style="font-family: 'JetBrains Mono', monospace;">${formatRupiah(item.price)}</td>
        <td><strong style="font-size: 1.05rem; color: #93c5fd;">${item.stokOnDate}</strong></td>
        <td style="font-family: 'JetBrains Mono', monospace; font-weight: 600;">${formatRupiah(item.totalValuation)}</td>
      </tr>
    `).join('');

    setText('reportPaginationInfo', `Menampilkan ${start + 1} - ${end} dari ${total} item laporan`);
    renderPaginationControls('reportPaginationControls', state.reportPage, totalPages, p => {
      state.reportPage = p;
      renderReportTable();
    });

  } else {
    thead.innerHTML = `
      <tr>
        <th style="width: 50px;">No</th>
        <th>Lokasi Penyimpanan</th>
        <th>Total Jenis Part</th>
        <th>Total Fisik Unit Stok</th>
        <th>Total Valuasi Nilai (IDR)</th>
      </tr>
    `;

    const locMap = {};
    state.reportData.forEach(item => {
      const loc = item.lokasi || 'GUDANG UTAMA';
      if (!locMap[loc]) {
        locMap[loc] = { count: 0, units: 0, valuation: 0 };
      }
      locMap[loc].count++;
      locMap[loc].units += Number(item.stokOnDate) || 0;
      locMap[loc].valuation += Number(item.totalValuation) || 0;
    });

    const locEntries = Object.entries(locMap).sort((a, b) => b[1].valuation - a[1].valuation);

    tbody.innerHTML = locEntries.map(([loc, stats], idx) => `
      <tr>
        <td style="color: var(--text-muted);">${idx + 1}</td>
        <td><strong style="color: #93c5fd; font-size: 1rem;">${escapeHtml(loc)}</strong></td>
        <td>${stats.count} material</td>
        <td><strong>${stats.units.toLocaleString('id-ID')} unit</strong></td>
        <td style="font-family: 'JetBrains Mono', monospace; font-weight: 700; color: #a78bfa; font-size: 1.05rem;">
          ${formatRupiah(stats.valuation)}
        </td>
      </tr>
    `).join('');

    setText('reportPaginationInfo', `Total ${locEntries.length} lokasi penyimpanan`);
    document.getElementById('reportPaginationControls').innerHTML = '';
  }
}

/**
 * ============================================================================
 * BACKUP & RESTORE DATABASE
 * ============================================================================
 */
async function downloadDatabaseBackup() {
  try {
    const backup = await window.InventoryDB.createDatabaseBackup();
    const jsonStr = JSON.stringify(backup, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const dateStr = new Date().toISOString().split('T')[0];
    a.href = url;
    a.download = `heavy_equipment_inventory_backup_${dateStr}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('File backup database berhasil diunduh!', 'success');
  } catch (err) {
    showToast('Gagal membuat backup: ' + err.message, 'danger');
  }
}

function triggerRestoreDatabase() {
  requireAdmin('memulihkan database dari backup', () => {
    const input = document.getElementById('backupFileInput');
    if (input) input.click();
  });
}

async function handleRestoreFileSelected(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async (event) => {
    try {
      const json = JSON.parse(event.target.result);
      if (!confirm(`Apakah Anda yakin ingin memulihkan database dari file cadangan ini?\nSeluruh data saat ini akan digantikan oleh isi backup.`)) {
        return;
      }

      await window.InventoryDB.restoreDatabaseBackup(json, 'Team');
      closeModal('backupModal');
      showToast('Database berhasil dipulihkan dari file backup!', 'success');
      await loadDataFromDB();
    } catch (err) {
      showToast('Gagal memulihkan database: ' + err.message, 'danger');
    }
  };
  reader.readAsText(file);
}

/**
 * ============================================================================
 * GOOGLE SHEETS LIVE SYNC
 * ============================================================================
 */
function loadConfig() {
  try {
    const saved = localStorage.getItem('he_inventory_config');
    if (saved) {
      state.config = { ...state.config, ...JSON.parse(saved) };
    }
  } catch (e) {}

  const urlInput = document.getElementById('cfgGasUrl');
  const sheetIdInput = document.getElementById('cfgSheetId');
  const intervalSelect = document.getElementById('cfgSyncInterval');
  const autoSyncCheck = document.getElementById('cfgAutoSync');

  if (urlInput) urlInput.value = state.config.gasWebAppUrl || '';
  if (sheetIdInput) sheetIdInput.value = state.config.spreadsheetId || '';
  if (intervalSelect) intervalSelect.value = state.config.syncIntervalMinutes || '1';
  if (autoSyncCheck) autoSyncCheck.checked = Boolean(state.config.autoSync);

  updateSyncBadge(state.config.gasWebAppUrl ? 'connected' : 'idle', state.config.gasWebAppUrl ? 'Tersambung Sheets' : 'Sheets Belum Disetel');

  if (state.config.gasWebAppUrl && state.config.autoSync) {
    startAutoSync();
  }
}

function saveConfig() {
  localStorage.setItem('he_inventory_config', JSON.stringify(state.config));
}

function updateSyncBadge(status, text) {
  const dot = document.getElementById('syncDot');
  const label = document.getElementById('syncLabel');

  if (dot) {
    dot.className = 'sync-dot';
    if (status === 'connected') dot.classList.add('online');
    else if (status === 'syncing') dot.classList.add('syncing');
    else if (status === 'error') dot.classList.add('offline');
  }

  if (label) label.textContent = text;
}

async function testGasConnection() {
  const url = (document.getElementById('cfgGasUrl').value || '').trim();
  if (!url) {
    showToast('Masukkan URL Web App Google Apps Script terlebih dahulu.', 'warning');
    return;
  }

  const btn = document.getElementById('btnTestGas');
  btn.disabled = true;
  btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Mengetes...';

  try {
    const res = await fetch(`${url}?action=ping`, { method: 'GET', mode: 'cors' });
    const data = await res.json();
    btn.disabled = false;
    btn.innerHTML = '<i class="fa-solid fa-plug"></i> Test Koneksi';

    if (data.status === 'success') {
      showToast(`Koneksi Berhasil! Terhubung ke spreadsheet: "${data.spreadsheetTitle}"`, 'success');
    } else {
      showToast('Koneksi Gagal: ' + (data.message || 'Error tidak diketahui'), 'danger');
    }
  } catch (err) {
    btn.disabled = false;
    btn.innerHTML = '<i class="fa-solid fa-plug"></i> Test Koneksi';
    showToast('Gagal menghubungi Web App URL: ' + err.message, 'danger');
  }
}

async function syncNow() {
  if (!state.config.gasWebAppUrl) {
    openModal('settingsModal');
    showToast('Silakan masukkan Web App URL Google Apps Script terlebih dahulu.', 'info');
    return;
  }

  if (state.isSyncing) return;
  state.isSyncing = true;
  updateSyncBadge('syncing', 'Sinkronisasi Cloud...');

  try {
    const res = await fetch(`${state.config.gasWebAppUrl}?action=getAllData`, { method: 'GET', mode: 'cors' });
    const json = await res.json();

    if (json.status === 'success' && json.data) {
      const { master, daily } = json.data;
      if (Array.isArray(master) && master.length > 0) {
        for (const m of master) {
          await window.InventoryDB.upsertMasterMaterial(m, 'GoogleSheets');
        }
        for (const d of (daily || [])) {
          await window.InventoryDB.upsertDailyStock(d, 'GoogleSheets');
        }
        await loadDataFromDB();
        showToast(`Sinkronisasi real-time berhasil! ${master.length} material sinkron dengan Google Sheets.`, 'success');
      } else {
        showToast('Sheet di Google Sheets masih kosong. Gunakan "Kirim Semua Data ke Sheets" untuk upload awal.', 'info');
      }
      updateSyncBadge('connected', 'Tersinkronisasi Real-Time');
    } else {
      updateSyncBadge('error', 'Gagal Sinkron');
      showToast('Gagal sinkronisasi: ' + (json.message || 'Error'), 'danger');
    }
  } catch (err) {
    updateSyncBadge('error', 'Error Koneksi');
    showToast('Gagal sinkronisasi dengan Google Sheets: ' + err.message, 'danger');
  } finally {
    state.isSyncing = false;
  }
}

async function bulkPushToSheets() {
  if (!state.config.gasWebAppUrl) {
    showToast('Masukkan URL Google Apps Script terlebih dahulu di kolom pengaturan.', 'warning');
    return;
  }

  const btn = document.getElementById('btnBulkPush');
  btn.disabled = true;
  btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Mengirim Seluruh Data...';

  try {
    const master = await window.InventoryDB.getAllMasterMaterials();
    const daily = await window.InventoryDB.getAllDailyStock();
    const audit = await window.InventoryDB.getAllAuditLogs();

    const payload = {
      action: 'syncAll',
      master: master,
      daily: daily,
      audit: audit
    };

    const res = await fetch(state.config.gasWebAppUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();

    btn.disabled = false;
    btn.innerHTML = '<i class="fa-solid fa-upload"></i> Kirim Semua Data ke Sheets';

    if (data.status === 'success') {
      showToast(data.message || 'Semua data Master (PS & LIB), Transaksi Harian, dan Audit berhasil terkirim ke Google Sheets!', 'success');
      updateSyncBadge('connected', 'Tersinkronisasi');
    } else {
      showToast('Gagal bulk push: ' + (data.message || 'Error'), 'danger');
    }
  } catch (err) {
    btn.disabled = false;
    btn.innerHTML = '<i class="fa-solid fa-upload"></i> Kirim Semua Data ke Sheets';
    showToast('Gagal mengirim data ke Sheets: ' + err.message, 'danger');
  }
}

function startAutoSync() {
  stopAutoSync();
  const intervalMs = (Number(state.config.syncIntervalMinutes) || 1) * 60 * 1000;
  state.syncTimer = setInterval(() => {
    if (state.config.gasWebAppUrl && !state.isSyncing) {
      syncNow();
    }
  }, intervalMs);
}

function stopAutoSync() {
  if (state.syncTimer) {
    clearInterval(state.syncTimer);
    state.syncTimer = null;
  }
}

function handleSettingsSubmit(e) {
  e.preventDefault();
  state.config.gasWebAppUrl = (document.getElementById('cfgGasUrl').value || '').trim();
  state.config.spreadsheetId = (document.getElementById('cfgSheetId').value || '').trim();
  state.config.syncIntervalMinutes = document.getElementById('cfgSyncInterval').value;
  state.config.autoSync = document.getElementById('cfgAutoSync').checked;

  saveConfig();
  closeModal('settingsModal');
  showToast('Pengaturan Google Sheets berhasil disimpan!', 'success');

  if (state.config.autoSync) {
    startAutoSync();
  } else {
    stopAutoSync();
  }
  updateSyncBadge(state.config.gasWebAppUrl ? 'connected' : 'idle', state.config.gasWebAppUrl ? 'Tersambung Sheets' : 'Sheets Belum Disetel');
}

/**
 * ============================================================================
 * UTILITIES & HELPERS
 * ============================================================================
 */
function initTheme() {
  const saved = localStorage.getItem('he_inventory_theme') || 'dark';
  document.documentElement.setAttribute('data-theme', saved);
  updateThemeIcon(saved);
}

function toggleTheme() {
  const current = document.documentElement.getAttribute('data-theme') || 'dark';
  const next = current === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  localStorage.setItem('he_inventory_theme', next);
  updateThemeIcon(next);
  updateCharts();
}

function updateThemeIcon(theme) {
  const btn = document.getElementById('themeToggleBtn');
  if (btn) {
    btn.innerHTML = theme === 'dark' ? '<i class="fa-solid fa-sun"></i>' : '<i class="fa-solid fa-moon"></i>';
  }
}

function formatRupiah(num) {
  const n = Number(num) || 0;
  return 'Rp ' + n.toLocaleString('id-ID');
}

function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;

  let icon = 'info-circle';
  if (type === 'success') icon = 'check-circle';
  if (type === 'warning') icon = 'triangle-exclamation';
  if (type === 'danger') icon = 'circle-exclamation';

  toast.innerHTML = `
    <i class="fa-solid fa-${icon} toast-icon"></i>
    <div class="toast-content">${escapeHtml(message)}</div>
    <button class="toast-close" title="Tutup">&times;</button>
  `;

  toast.querySelector('.toast-close').addEventListener('click', () => toast.remove());
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    setTimeout(() => toast.remove(), 300);
  }, 4500);
}

function openModal(id) {
  const modal = document.getElementById(id);
  if (modal) modal.classList.add('active');
}

function closeModal(id) {
  const modal = document.getElementById(id);
  if (modal) modal.classList.remove('active');
}

function closeAllModals() {
  document.querySelectorAll('.modal-overlay').forEach(m => m.classList.remove('active'));
}

function setText(id, text) {
  const el = document.getElementById(id);
  if (el) el.innerHTML = text;
}

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeQuote(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/'/g, "\\'").replace(/"/g, '&quot;');
}

function renderPaginationControls(containerId, currentPage, totalPages, onPageClick) {
  const container = document.getElementById(containerId);
  if (!container) return;

  if (totalPages <= 1) {
    container.innerHTML = '';
    return;
  }

  let html = `
    <button class="page-btn" ${currentPage === 1 ? 'disabled' : ''} data-page="${currentPage - 1}">
      <i class="fa-solid fa-chevron-left"></i>
    </button>
  `;

  for (let p = 1; p <= totalPages; p++) {
    if (p === 1 || p === totalPages || (p >= currentPage - 2 && p <= currentPage + 2)) {
      html += `<button class="page-btn ${p === currentPage ? 'active' : ''}" data-page="${p}">${p}</button>`;
    } else if (p === currentPage - 3 || p === currentPage + 3) {
      html += `<span class="page-ellipsis">...</span>`;
    }
  }

  html += `
    <button class="page-btn" ${currentPage === totalPages ? 'disabled' : ''} data-page="${currentPage + 1}">
      <i class="fa-solid fa-chevron-right"></i>
    </button>
  `;

  container.innerHTML = html;
  container.querySelectorAll('.page-btn:not([disabled])').forEach(btn => {
    btn.addEventListener('click', () => {
      const p = parseInt(btn.getAttribute('data-page'), 10);
      if (p) onPageClick(p);
    });
  });
}

function exportToCsv(items, filename) {
  if (!items || items.length === 0) {
    showToast('Tidak ada data untuk diekspor.', 'warning');
    return;
  }

  const keys = Object.keys(items[0]);
  const csvRows = [];
  csvRows.push(keys.join(','));

  items.forEach(item => {
    const values = keys.map(k => {
      const val = item[k] === null || item[k] === undefined ? '' : String(item[k]);
      return `"${val.replace(/"/g, '""')}"`;
    });
    csvRows.push(values.join(','));
  });

  const blob = new Blob([csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${filename}_${new Date().toISOString().split('T')[0]}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('File CSV berhasil diunduh.', 'success');
}

function exportToJson(items, filename) {
  if (!items || items.length === 0) {
    showToast('Tidak ada data untuk diekspor.', 'warning');
    return;
  }
  const blob = new Blob([JSON.stringify(items, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${filename}_${new Date().toISOString().split('T')[0]}.json`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('File JSON berhasil diunduh.', 'success');
}

function copyGasScript() {
  fetch('Code.gs')
    .then(r => r.text())
    .then(text => {
      navigator.clipboard.writeText(text);
      showToast('Skrip Google Apps Script (Code.gs) berhasil disalin ke clipboard!', 'success');
    })
    .catch(() => {
      showToast('File Code.gs ada di dalam folder proyek Anda.', 'info');
    });
}

/**
 * ============================================================================
 * EVENT LISTENERS SETUP
 * ============================================================================
 */
function setupEventListeners() {
  // Navigation Tabs
  document.querySelectorAll('.nav-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const view = btn.getAttribute('data-view');
      if (view) switchAppView(view);
    });
  });

  // Top Bar Actions
  document.getElementById('btnSyncNow')?.addEventListener('click', syncNow);
  document.getElementById('btnOpenSettings')?.addEventListener('click', () => openModal('settingsModal'));
  document.getElementById('btnOpenBackup')?.addEventListener('click', () => openModal('backupModal'));
  document.getElementById('themeToggleBtn')?.addEventListener('click', toggleTheme);

  // Admin Authentication & Password Forms
  document.getElementById('adminLoginForm')?.addEventListener('submit', handleAdminLoginSubmit);
  document.getElementById('changePasswordForm')?.addEventListener('submit', handleChangePasswordSubmit);
  document.getElementById('btnToggleAdminPass')?.addEventListener('click', toggleAdminPassVisibility);
  document.getElementById('btnOpenChangePassword')?.addEventListener('click', () => {
    closeModal('settingsModal');
    openChangePasswordModal();
  });

  // Master Table Column Sorting
  setupMasterTableSorting();

  // Corporate Hero Banner Lightbox Modal
  const openHeroLightbox = () => openModal('imageLightboxModal');
  document.getElementById('btnOpenHeroLightbox')?.addEventListener('click', openHeroLightbox);
  document.getElementById('heroBannerClickable')?.addEventListener('click', openHeroLightbox);

  // Modal Closers
  document.querySelectorAll('.modal-close, .btn-modal-cancel').forEach(btn => {
    btn.addEventListener('click', closeAllModals);
  });
  document.querySelectorAll('.modal-overlay').forEach(overlay => {
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) closeAllModals();
    });
  });

  // Master Data Multi-Select & Bulk Bar Actions
  document.getElementById('checkSelectAllMaster')?.addEventListener('change', handleSelectAllMaster);
  document.getElementById('btnBulkDeselect')?.addEventListener('click', deselectAllMaster);
  document.getElementById('btnBulkSetPS')?.addEventListener('click', () => executeBulkStatusChange('PS'));
  document.getElementById('btnBulkSetNonPS')?.addEventListener('click', () => executeBulkStatusChange('NON PS'));
  document.getElementById('btnBulkSetLIB')?.addEventListener('click', () => executeBulkLibChange('LIB'));
  document.getElementById('btnBulkSetNonLIB')?.addEventListener('click', () => executeBulkLibChange('NON LIB'));
  document.getElementById('btnBulkOpenLocation')?.addEventListener('click', openBulkLocationModal);
  document.getElementById('bulkLocationForm')?.addEventListener('submit', handleBulkLocationSubmit);
  document.getElementById('btnBulkOpenQty')?.addEventListener('click', openBulkQtyModal);
  document.getElementById('bulkQtyForm')?.addEventListener('submit', handleBulkQtySubmit);
  document.getElementById('btnBulkOpenMinMax')?.addEventListener('click', openBulkMinMaxModal);
  document.getElementById('bulkMinMaxForm')?.addEventListener('submit', handleBulkMinMaxSubmit);
  document.getElementById('btnBulkDeleteSelected')?.addEventListener('click', handleBulkDeleteSelected);

  // Master Data Filters & Actions
  document.getElementById('btnAddNewMaster')?.addEventListener('click', openAddMasterModal);
  document.getElementById('partForm')?.addEventListener('submit', handleMasterFormSubmit);
  document.getElementById('searchMasterInput')?.addEventListener('input', (e) => {
    state.masterFilters.search = e.target.value;
    applyMasterFilters();
  });
  document.getElementById('filterMasterLokasi')?.addEventListener('change', (e) => {
    state.masterFilters.lokasi = e.target.value;
    applyMasterFilters();
  });
  document.getElementById('filterMasterService')?.addEventListener('change', (e) => {
    state.masterFilters.service = e.target.value;
    applyMasterFilters();
  });
  document.getElementById('filterMasterLib')?.addEventListener('change', (e) => {
    state.masterFilters.lib = e.target.value;
    applyMasterFilters();
  });
  document.getElementById('pageSizeMaster')?.addEventListener('change', (e) => {
    state.masterPageSize = e.target.value;
    renderMasterTable();
  });
  document.querySelectorAll('[data-master-tab]').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('[data-master-tab]').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      state.masterFilters.statusTab = tab.getAttribute('data-master-tab');
      applyMasterFilters();
    });
  });
  document.getElementById('btnExportMasterCsv')?.addEventListener('click', () => {
    exportToCsv(state.filteredMaster, 'master_material');
  });
  document.getElementById('btnExportMasterJson')?.addEventListener('click', () => {
    exportToJson(state.filteredMaster, 'master_material');
  });

  // Daily Stock Actions
  document.getElementById('dailyQuickForm')?.addEventListener('submit', handleDailyQuickSubmit);
  document.getElementById('dailyFormKode')?.addEventListener('input', syncDailyFormMaterialName);
  document.getElementById('editDailyForm')?.addEventListener('submit', handleEditDailySubmit);
  document.getElementById('searchDailyInput')?.addEventListener('input', (e) => {
    state.dailyFilters.search = e.target.value;
    applyDailyFilters();
  });
  document.getElementById('filterDailyDate')?.addEventListener('change', (e) => {
    state.dailyFilters.date = e.target.value;
    applyDailyFilters();
  });
  document.getElementById('filterDailyMonth')?.addEventListener('change', (e) => {
    state.dailyFilters.month = e.target.value;
    applyDailyFilters();
  });
  document.getElementById('filterDailyYear')?.addEventListener('change', (e) => {
    state.dailyFilters.year = e.target.value;
    applyDailyFilters();
  });
  document.getElementById('filterDailyHour')?.addEventListener('change', (e) => {
    state.dailyFilters.hour = e.target.value;
    applyDailyFilters();
  });
  document.getElementById('pageSizeDaily')?.addEventListener('change', (e) => {
    state.dailyPageSize = e.target.value;
    renderDailyTable();
  });
  document.getElementById('btnDailyTodayFilter')?.addEventListener('click', () => {
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('filterDailyDate').value = today;
    state.dailyFilters.date = today;
    applyDailyFilters();
  });
  document.getElementById('btnResetDailyFilter')?.addEventListener('click', () => {
    document.getElementById('searchDailyInput').value = '';
    document.getElementById('filterDailyDate').value = '';
    document.getElementById('filterDailyMonth').value = 'ALL';
    document.getElementById('filterDailyYear').value = 'ALL';
    document.getElementById('filterDailyHour').value = 'ALL';
    state.dailyFilters = { search: '', date: '', month: 'ALL', year: 'ALL', hour: 'ALL' };
    applyDailyFilters();
  });
  document.getElementById('btnExportDailyCsv')?.addEventListener('click', () => {
    exportToCsv(state.filteredDaily, 'daily_stock');
  });

  // Bulk Data Import Actions
  document.getElementById('btnPasteClipboard')?.addEventListener('click', pasteFromClipboard);
  document.getElementById('btnClearPaste')?.addEventListener('click', () => {
    document.getElementById('excelPasteArea').value = '';
  });
  document.getElementById('btnProcessPreview')?.addEventListener('click', processImportPreview);
  document.getElementById('btnCancelImport')?.addEventListener('click', cancelImportPreview);
  document.getElementById('btnConfirmImport')?.addEventListener('click', confirmAndExecuteImport);

  // Stock History Actions
  document.getElementById('tabHistoryLedger')?.addEventListener('click', () => {
    document.getElementById('tabHistoryLedger').classList.add('active');
    document.getElementById('tabHistoryAudit').classList.remove('active');
    state.historyTab = 'ledger';
    applyHistoryFilters();
  });
  document.getElementById('tabHistoryAudit')?.addEventListener('click', () => {
    document.getElementById('tabHistoryAudit').classList.add('active');
    document.getElementById('tabHistoryLedger').classList.remove('active');
    state.historyTab = 'audit';
    applyHistoryFilters();
  });
  document.getElementById('searchHistoryInput')?.addEventListener('input', (e) => {
    state.historyFilters.search = e.target.value;
    applyHistoryFilters();
  });
  document.getElementById('filterHistoryDate')?.addEventListener('change', (e) => {
    state.historyFilters.date = e.target.value;
    applyHistoryFilters();
  });
  document.getElementById('filterHistoryHour')?.addEventListener('change', (e) => {
    state.historyFilters.hour = e.target.value;
    applyHistoryFilters();
  });
  document.getElementById('filterHistoryLokasi')?.addEventListener('change', (e) => {
    state.historyFilters.lokasi = e.target.value;
    applyHistoryFilters();
  });
  document.getElementById('btnResetHistoryFilter')?.addEventListener('click', () => {
    document.getElementById('searchHistoryInput').value = '';
    document.getElementById('filterHistoryDate').value = '';
    document.getElementById('filterHistoryHour').value = 'ALL';
    document.getElementById('filterHistoryLokasi').value = 'ALL';
    state.historyFilters = { search: '', date: '', hour: 'ALL', lokasi: 'ALL' };
    applyHistoryFilters();
  });

  // Reports Actions
  document.querySelectorAll('[data-report]').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('[data-report]').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.reportType = btn.getAttribute('data-report');
      renderReportTable();
    });
  });
  document.getElementById('btnGenerateReport')?.addEventListener('click', generateReport);
  document.getElementById('btnPrintReport')?.addEventListener('click', () => window.print());
  document.getElementById('btnExportReportCsv')?.addEventListener('click', () => {
    exportToCsv(state.reportData, `report_stock_${state.reportTargetDate}`);
  });

  // Backup & Restore Actions
  document.getElementById('btnDownloadBackup')?.addEventListener('click', downloadDatabaseBackup);
  document.getElementById('btnTriggerRestore')?.addEventListener('click', triggerRestoreDatabase);
  document.getElementById('backupFileInput')?.addEventListener('change', handleRestoreFileSelected);

  // Settings & Sheets Sync
  document.getElementById('settingsForm')?.addEventListener('submit', handleSettingsSubmit);
  document.getElementById('btnTestGas')?.addEventListener('click', testGasConnection);
  document.getElementById('btnBulkPush')?.addEventListener('click', bulkPushToSheets);
  document.getElementById('btnCopyCode')?.addEventListener('click', copyGasScript);

  // Default Dates
  const today = new Date().toISOString().split('T')[0];
  const dailyDateEl = document.getElementById('dailyFormDate');
  const importDateEl = document.getElementById('importDefaultDate');
  const reportDateEl = document.getElementById('reportTargetDate');
  if (dailyDateEl) dailyDateEl.value = today;
  if (importDateEl) importDateEl.value = today;
  if (reportDateEl) reportDateEl.value = today;
}

/**
 * ============================================================================
 * APPLICATION ENTRY POINT
 * ============================================================================
 */
document.addEventListener('DOMContentLoaded', async () => {
  initTheme();
  renderAuthStatus();
  setupEventListeners();

  try {
    updateSyncBadge('syncing', 'Inisialisasi Database...');
    await window.InventoryDB.init();
    loadConfig();
    await loadDataFromDB();
    updateSyncBadge('connected', 'Sistem Siap Real-Time');
  } catch (err) {
    console.error('Initialization error:', err);
    showToast('Error inisialisasi: ' + err.message, 'danger');
  }
});
