/**
 * ============================================================================
 * PT. STAR WAGEN INDONESIA - VIB - IndexedDB Database Engine (db.js)
 * ============================================================================
 * - Primary Key / Unique Constraint on Kode Material
 * - Flexible Multi-Field Editing: QTY, Min, Max, PS/NON PS, LIB/NON LIB, Lokasi
 * - Rapid Mass Bulk Updates & Universal Excel Table Parsing
 * - Cascade Stock Balance Recalculation (Previous Day SISA -> Next Day Saldo Awal)
 * - Offline-First with Google Sheets Cloud Sync Integration
 * ============================================================================
 */

const DB_NAME = 'HeavyEquipmentInventoryDB';
const DB_VERSION = 3;

let dbInstance = null;

// Initialize & Upgrade IndexedDB
function initDatabase() {
  return new Promise((resolve, reject) => {
    if (dbInstance) {
      return resolve(dbInstance);
    }

    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = (event) => {
      const db = event.target.result;

      // 1. Master Material Store (Key: kodeMaterial UNIQUE)
      let masterStore;
      if (!db.objectStoreNames.contains('master_material')) {
        masterStore = db.createObjectStore('master_material', { keyPath: 'kodeMaterial' });
        masterStore.createIndex('namaMaterial', 'namaMaterial', { unique: false });
        masterStore.createIndex('lokasi', 'lokasi', { unique: false });
        masterStore.createIndex('prod', 'prod', { unique: false });
      } else {
        masterStore = event.target.transaction.objectStore('master_material');
      }

      // Add indexes for status and libCategory if not exist
      if (!masterStore.indexNames.contains('status')) {
        masterStore.createIndex('status', 'status', { unique: false });
      }
      if (!masterStore.indexNames.contains('libCategory')) {
        masterStore.createIndex('libCategory', 'libCategory', { unique: false });
      }

      // 2. Daily Stock Transactions Store (Key: id UNIQUE)
      if (!db.objectStoreNames.contains('daily_stock')) {
        const dailyStore = db.createObjectStore('daily_stock', { keyPath: 'id' });
        dailyStore.createIndex('kodeMaterial', 'kodeMaterial', { unique: false });
        dailyStore.createIndex('tanggal', 'tanggal', { unique: false });
        dailyStore.createIndex('kodeMaterial_tanggal', ['kodeMaterial', 'tanggal'], { unique: false });
      }

      // 3. Audit Logs Store (Key: id UNIQUE)
      if (!db.objectStoreNames.contains('audit_logs')) {
        const auditStore = db.createObjectStore('audit_logs', { keyPath: 'id' });
        auditStore.createIndex('timestamp', 'timestamp', { unique: false });
        auditStore.createIndex('kodeMaterial', 'kodeMaterial', { unique: false });
        auditStore.createIndex('action', 'action', { unique: false });
      }

      // 4. Recovery Snapshots Store (Key: id)
      if (!db.objectStoreNames.contains('snapshots')) {
        const snapStore = db.createObjectStore('snapshots', { keyPath: 'id' });
        snapStore.createIndex('timestamp', 'timestamp', { unique: false });
      }

      // 5. Config Store (Key: key)
      if (!db.objectStoreNames.contains('config')) {
        db.createObjectStore('config', { keyPath: 'key' });
      }
    };

    request.onsuccess = async (event) => {
      dbInstance = event.target.result;
      try {
        await checkAndMigrateInitialData();
        resolve(dbInstance);
      } catch (err) {
        console.warn('Migration warning:', err);
        resolve(dbInstance);
      }
    };

    request.onerror = (event) => {
      console.error('IndexedDB error:', event.target.error);
      reject(event.target.error);
    };
  });
}

// Auto-migration for existing 338 parts with LIB & PS categorization
async function checkAndMigrateInitialData() {
  const count = await getCount('master_material');
  if (count === 0) {
    console.log('Master material store is empty. Migrating initial dataset...');
    let sourceData = [];

    // Check localStorage first
    try {
      const local = localStorage.getItem('he_inventory_data');
      if (local) {
        sourceData = JSON.parse(local);
      }
    } catch (e) {}

    // Fallback to INITIAL_INVENTORY_DATA in data.js
    if ((!sourceData || sourceData.length === 0) && typeof INITIAL_INVENTORY_DATA !== 'undefined' && Array.isArray(INITIAL_INVENTORY_DATA)) {
      sourceData = INITIAL_INVENTORY_DATA;
    }

    if (sourceData && sourceData.length > 0) {
      const tx = dbInstance.transaction(['master_material', 'audit_logs'], 'readwrite');
      const store = tx.objectStore('master_material');
      const nowIso = new Date().toISOString();

      for (let i = 0; i < sourceData.length; i++) {
        const it = sourceData[i];
        const kode = String(it.pnFinal || it.kodeMaterial || ('MAT-' + (i + 1))).trim();
        if (!kode) continue;

        const descUpper = String(it.desc || it.namaMaterial || '').toUpperCase();
        // Intelligent auto-detection of LIB (Lubricant, Oil, Pelumas, Grease)
        const isAutoLib = descUpper.includes('OLEO') || descUpper.includes('OIL') || descUpper.includes('LUB') || descUpper.includes('PELUMAS') || descUpper.includes('GREASE');
        const libCategory = it.libCategory || (isAutoLib ? 'LIB' : 'NON LIB');

        const record = {
          no: it.no || (i + 1),
          kodeMaterial: kode,
          lokasi: (it.remarks && it.remarks.includes('PALET')) ? it.remarks : (it.lokasi || 'GUDANG UTAMA'),
          namaMaterial: it.desc || it.namaMaterial || 'Spare Part ' + kode,
          price: Number(it.price) || 0,
          stokAwal: Number(it.qty !== undefined ? it.qty : (it.stokAwal || 0)),
          stokSaatIni: Number(it.qty !== undefined ? it.qty : (it.stokSaatIni || 0)),
          prod: it.prod || 'GENERAL',
          modelUnit: it.modelUnit || '',
          min: it.min !== undefined && it.min !== null ? Number(it.min) : 1,
          max: it.max !== undefined && it.max !== null ? Number(it.max) : 5,
          status: it.status || 'PS',
          libCategory: libCategory,
          doi: it.doi !== undefined ? it.doi : null,
          remarks: it.remarks || '',
          alternatePn: it.alternatePn || '',
          createdAt: nowIso,
          updatedAt: nowIso
        };

        store.put(record);
      }

      // Add audit entry for initial migration
      const auditStore = tx.objectStore('audit_logs');
      auditStore.put({
        id: 'LOG-INIT-' + Date.now(),
        timestamp: nowIso,
        user: 'System',
        action: 'INITIAL_MIGRATION',
        kodeMaterial: 'ALL',
        details: `Migrasi awal berhasil: ${sourceData.length} spare part didaftarkan ke Master Data Material (dengan status PS/NON PS dan kategori LIB/NON LIB).`
      });

      return new Promise((res, rej) => {
        tx.oncomplete = () => {
          console.log(`Migrasi berhasil: ${sourceData.length} item dimasukkan ke master_material.`);
          res();
        };
        tx.onerror = () => rej(tx.error);
      });
    }
  } else {
    // If database already existed, ensure every record has a valid libCategory
    const all = await getAllMasterMaterials();
    const needsPatch = all.some(m => !m.libCategory);
    if (needsPatch) {
      const tx = dbInstance.transaction('master_material', 'readwrite');
      const store = tx.objectStore('master_material');
      for (const m of all) {
        if (!m.libCategory) {
          const descUpper = String(m.namaMaterial || '').toUpperCase();
          const isAutoLib = descUpper.includes('OLEO') || descUpper.includes('OIL') || descUpper.includes('LUB') || descUpper.includes('PELUMAS') || descUpper.includes('GREASE');
          m.libCategory = isAutoLib ? 'LIB' : 'NON LIB';
          store.put(m);
        }
      }
    }
  }
}

// Generic count helper
function getCount(storeName) {
  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction(storeName, 'readonly');
    const store = tx.objectStore(storeName);
    const req = store.count();
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

/**
 * ============================================================================
 * MASTER DATA MATERIAL OPERATIONS
 * ============================================================================
 */

// Get all master materials
function getAllMasterMaterials() {
  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction('master_material', 'readonly');
    const store = tx.objectStore('master_material');
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

// Get single master material by unique Kode Material
function getMasterMaterial(kodeMaterial) {
  return new Promise((resolve, reject) => {
    if (!kodeMaterial) return resolve(null);
    const tx = dbInstance.transaction('master_material', 'readonly');
    const store = tx.objectStore('master_material');
    const req = store.get(String(kodeMaterial).trim());
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

// Upsert Master Material (Create if not exist, update if already registered)
async function upsertMasterMaterial(materialData, user = 'Team') {
  const kode = String(materialData.kodeMaterial || '').trim();
  if (!kode) {
    throw new Error('Kode Material wajib diisi.');
  }

  const existing = await getMasterMaterial(kode);
  const nowIso = new Date().toISOString();

  const record = {
    no: materialData.no || (existing ? existing.no : (await getCount('master_material')) + 1),
    kodeMaterial: kode,
    lokasi: String(materialData.lokasi !== undefined ? materialData.lokasi : (existing ? existing.lokasi : 'GUDANG UTAMA')).trim(),
    namaMaterial: String(materialData.namaMaterial !== undefined ? materialData.namaMaterial : (existing ? existing.namaMaterial : '')).trim(),
    price: materialData.price !== undefined ? Number(materialData.price) : (existing ? existing.price : 0),
    stokAwal: materialData.stokAwal !== undefined ? Number(materialData.stokAwal) : (existing ? existing.stokAwal : 0),
    stokSaatIni: materialData.stokSaatIni !== undefined ? Number(materialData.stokSaatIni) : (existing ? existing.stokSaatIni : (Number(materialData.stokAwal) || 0)),
    prod: materialData.prod || (existing ? existing.prod : 'GENERAL'),
    modelUnit: materialData.modelUnit !== undefined ? materialData.modelUnit : (existing ? existing.modelUnit : ''),
    min: materialData.min !== undefined && materialData.min !== null ? Number(materialData.min) : (existing ? existing.min : 1),
    max: materialData.max !== undefined && materialData.max !== null ? Number(materialData.max) : (existing ? existing.max : 5),
    status: materialData.status || (existing ? existing.status : 'PS'),
    libCategory: materialData.libCategory || (existing ? existing.libCategory : 'NON LIB'),
    doi: materialData.doi !== undefined ? materialData.doi : (existing ? existing.doi : null),
    remarks: materialData.remarks !== undefined ? materialData.remarks : (existing ? existing.remarks : ''),
    alternatePn: materialData.alternatePn !== undefined ? materialData.alternatePn : (existing ? existing.alternatePn : ''),
    createdAt: existing ? existing.createdAt : nowIso,
    updatedAt: nowIso
  };

  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction(['master_material', 'audit_logs'], 'readwrite');
    const store = tx.objectStore('master_material');
    store.put(record);

    const auditStore = tx.objectStore('audit_logs');
    auditStore.put({
      id: 'LOG-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
      timestamp: nowIso,
      user: user,
      action: existing ? 'UPDATE_MATERIAL' : 'CREATE_MATERIAL',
      kodeMaterial: kode,
      details: existing 
        ? `Update material: ${record.namaMaterial} (Qty: ${record.stokSaatIni}, Min: ${record.min}, Max: ${record.max}, Status: ${record.status}, LIB: ${record.libCategory}, Lokasi: ${record.lokasi})`
        : `Material baru terdaftar: ${record.namaMaterial} (Stok Awal: ${record.stokAwal}, Status: ${record.status}, LIB: ${record.libCategory})`
    });

    tx.oncomplete = () => resolve(record);
    tx.onerror = () => reject(tx.error);
  });
}

/**
 * Mass Bulk Update Master Materials (Update multiple parts in a single transaction in milliseconds)
 * Supports updating any combination of: QTY, min, max, status (PS/NON PS), libCategory (LIB/NON LIB), lokasi, price
 */
async function bulkUpdateMasterMaterials(updates, user = 'Team') {
  if (!Array.isArray(updates) || updates.length === 0) return [];
  const nowIso = new Date().toISOString();
  const updatedRecords = [];

  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction(['master_material', 'audit_logs'], 'readwrite');
    const store = tx.objectStore('master_material');
    const auditStore = tx.objectStore('audit_logs');

    const getReq = store.getAll();
    getReq.onsuccess = () => {
      const all = getReq.result || [];
      const map = new Map();
      all.forEach(m => map.set(m.kodeMaterial, m));

      let updateCount = 0;
      for (const u of updates) {
        const kode = String(u.kodeMaterial || '').trim();
        const existing = map.get(kode);
        if (!existing) continue;

        if (u.stokSaatIni !== undefined && u.stokSaatIni !== null) {
          existing.stokSaatIni = Number(u.stokSaatIni);
        }
        if (u.qtyDelta !== undefined && u.qtyDelta !== null) {
          existing.stokSaatIni = Math.max(0, (Number(existing.stokSaatIni) || 0) + Number(u.qtyDelta));
        }
        if (u.min !== undefined && u.min !== null) existing.min = Number(u.min);
        if (u.max !== undefined && u.max !== null) existing.max = Number(u.max);
        if (u.status !== undefined && u.status !== null) existing.status = String(u.status).trim();
        if (u.libCategory !== undefined && u.libCategory !== null) existing.libCategory = String(u.libCategory).trim();
        if (u.lokasi !== undefined && u.lokasi !== null) existing.lokasi = String(u.lokasi).trim();
        if (u.namaMaterial !== undefined && u.namaMaterial !== null) existing.namaMaterial = String(u.namaMaterial).trim();
        if (u.price !== undefined && u.price !== null) existing.price = Number(u.price);
        if (u.remarks !== undefined && u.remarks !== null) existing.remarks = String(u.remarks).trim();

        existing.updatedAt = nowIso;
        store.put(existing);
        updatedRecords.push(existing);
        updateCount++;
      }

      auditStore.put({
        id: 'LOG-' + Date.now(),
        timestamp: nowIso,
        user: user,
        action: 'BULK_UPDATE_MASTER',
        kodeMaterial: `${updateCount} Parts`,
        details: `Pembaruan massal ${updateCount} part (QTY / Min-Max / PS-NON PS / LIB-NON LIB / Lokasi).`
      });
    };

    tx.oncomplete = () => resolve(updatedRecords);
    tx.onerror = () => reject(tx.error);
  });
}

// Delete Master Material and its daily transactions
async function deleteMasterMaterial(kodeMaterial, user = 'Team') {
  const kode = String(kodeMaterial).trim();
  const existing = await getMasterMaterial(kode);
  if (!existing) {
    throw new Error(`Material dengan kode ${kode} tidak ditemukan.`);
  }

  // Get all transactions for this material
  const transactions = await getDailyStockByMaterial(kode);

  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction(['master_material', 'daily_stock', 'audit_logs'], 'readwrite');
    const masterStore = tx.objectStore('master_material');
    const dailyStore = tx.objectStore('daily_stock');
    const auditStore = tx.objectStore('audit_logs');

    masterStore.delete(kode);
    for (const t of transactions) {
      dailyStore.delete(t.id);
    }

    auditStore.put({
      id: 'LOG-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
      timestamp: new Date().toISOString(),
      user: user,
      action: 'DELETE_MATERIAL',
      kodeMaterial: kode,
      details: `Menghapus master material ${existing.namaMaterial} beserta ${transactions.length} riwayat transaksi harian.`
    });

    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
  });
}

/**
 * ============================================================================
 * DAILY STOCK TRANSACTIONS OPERATIONS & CASCADE RECALCULATION
 * ============================================================================
 */

// Get all daily transactions
function getAllDailyStock() {
  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction('daily_stock', 'readonly');
    const store = tx.objectStore('daily_stock');
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

// Get daily transactions for a specific material
function getDailyStockByMaterial(kodeMaterial) {
  return new Promise((resolve, reject) => {
    const kode = String(kodeMaterial).trim();
    const tx = dbInstance.transaction('daily_stock', 'readonly');
    const store = tx.objectStore('daily_stock');
    const index = store.index('kodeMaterial');
    const req = index.getAll(kode);
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

// Get single transaction by ID
function getDailyStockById(id) {
  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction('daily_stock', 'readonly');
    const store = tx.objectStore('daily_stock');
    const req = store.get(id);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

/**
 * Cascade Recalculation for a material across all transactions
 */
async function recalculateStockForMaterial(kodeMaterial) {
  const kode = String(kodeMaterial).trim();
  const material = await getMasterMaterial(kode);
  if (!material) return null;

  const rawTransactions = await getDailyStockByMaterial(kode);

  if (rawTransactions.length === 0) {
    return { updatedMaterial: material, updatedTransactions: [] };
  }

  // Sort ascending by date, then time
  const sorted = [...rawTransactions].sort((a, b) => {
    if (a.tanggal === b.tanggal) {
      return (a.jam || a.createdAt || '').localeCompare(b.jam || b.createdAt || '');
    }
    return a.tanggal.localeCompare(b.tanggal);
  });

  let currentBalance = Number(material.stokAwal) || 0;
  const updatedTransactions = [];

  for (let i = 0; i < sorted.length; i++) {
    const tx = sorted[i];
    const inQty = Number(tx.inQty) || 0;
    const outQty = Number(tx.outQty) || 0;
    const sisa = currentBalance + inQty - outQty;

    updatedTransactions.push({
      ...tx,
      saldoAwal: currentBalance,
      sisa: sisa
    });

    currentBalance = sisa;
  }

  // Persist updated transactions and updated material stock
  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction(['daily_stock', 'master_material'], 'readwrite');
    const dailyStore = tx.objectStore('daily_stock');
    const masterStore = tx.objectStore('master_material');

    for (const t of updatedTransactions) {
      dailyStore.put(t);
    }

    const updatedMaterial = {
      ...material,
      stokSaatIni: currentBalance,
      updatedAt: new Date().toISOString()
    };
    masterStore.put(updatedMaterial);

    tx.oncomplete = () => {
      resolve({ updatedMaterial, updatedTransactions });
    };
    tx.onerror = () => reject(tx.error);
  });
}

// Upsert Daily Stock Transaction
async function upsertDailyStock(txData, user = 'Team') {
  const kode = String(txData.kodeMaterial || '').trim();
  if (!kode) throw new Error('Kode Material wajib diisi.');
  if (!txData.tanggal) throw new Error('Tanggal transaksi wajib diisi.');

  const material = await getMasterMaterial(kode);
  if (!material) {
    throw new Error(`Material ${kode} belum terdaftar di Master Data.`);
  }

  const existingTxs = await getDailyStockByMaterial(kode);
  
  let existingForDate = null;
  if (txData.id) {
    existingForDate = existingTxs.find(t => t.id === txData.id);
  }
  if (!existingForDate) {
    existingForDate = existingTxs.find(t => t.tanggal === txData.tanggal);
  }

  const now = new Date();
  const timeStr = txData.jam || now.toTimeString().split(' ')[0].substring(0, 5); // "HH:MM"
  const nowIso = now.toISOString();

  const txId = existingForDate ? existingForDate.id : ('TX-' + Date.now() + '-' + Math.random().toString(36).substr(2, 6));

  const record = {
    id: txId,
    tanggal: txData.tanggal, // Format: "YYYY-MM-DD"
    jam: timeStr,
    kodeMaterial: kode,
    namaMaterial: material.namaMaterial,
    lokasi: material.lokasi,
    inQty: Number(txData.inQty !== undefined ? txData.inQty : (existingForDate ? existingForDate.inQty : 0)) || 0,
    outQty: Number(txData.outQty !== undefined ? txData.outQty : (existingForDate ? existingForDate.outQty : 0)) || 0,
    utd: txData.utd !== undefined ? String(txData.utd).trim() : (existingForDate ? existingForDate.utd : ''),
    sisa: 0,
    notes: String(txData.notes || (existingForDate ? existingForDate.notes : '')).trim(),
    createdAt: existingForDate ? existingForDate.createdAt : nowIso,
    updatedAt: nowIso,
    updatedBy: user
  };

  await new Promise((resolve, reject) => {
    const tx = dbInstance.transaction(['daily_stock', 'audit_logs'], 'readwrite');
    const store = tx.objectStore('daily_stock');
    store.put(record);

    const auditStore = tx.objectStore('audit_logs');
    auditStore.put({
      id: 'LOG-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
      timestamp: nowIso,
      user: user,
      action: existingForDate ? 'EDIT_DAILY_STOCK' : 'ADD_DAILY_STOCK',
      kodeMaterial: kode,
      details: `${existingForDate ? 'Edit' : 'Tambah'} transaksi ${record.tanggal}: IN=${record.inQty}, OUT=${record.outQty}, UTD=${record.utd}`
    });

    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });

  const recalcResult = await recalculateStockForMaterial(kode);
  const finalRecord = recalcResult.updatedTransactions.find(t => t.id === txId) || record;

  return { record: finalRecord, material: recalcResult.updatedMaterial };
}

// Delete Daily Stock Transaction
async function deleteDailyStock(txId, user = 'Team') {
  const existing = await getDailyStockById(txId);
  if (!existing) {
    throw new Error('Transaksi tidak ditemukan.');
  }

  const kode = existing.kodeMaterial;

  await new Promise((resolve, reject) => {
    const tx = dbInstance.transaction(['daily_stock', 'audit_logs'], 'readwrite');
    const store = tx.objectStore('daily_stock');
    store.delete(txId);

    const auditStore = tx.objectStore('audit_logs');
    auditStore.put({
      id: 'LOG-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
      timestamp: new Date().toISOString(),
      user: user,
      action: 'DELETE_DAILY_STOCK',
      kodeMaterial: kode,
      details: `Hapus transaksi ${existing.tanggal}: IN=${existing.inQty}, OUT=${existing.outQty}`
    });

    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });

  await recalculateStockForMaterial(kode);
  return true;
}

/**
 * ============================================================================
 * BULK DATA IMPORT & UNIVERSAL EXCEL UPSERT ENGINE
 * ============================================================================
 */

// Create pre-import snapshot for recovery
async function createSnapshot(note = 'Pre-import snapshot') {
  const master = await getAllMasterMaterials();
  const daily = await getAllDailyStock();
  const snapId = 'SNAP-' + Date.now();

  const snapshot = {
    id: snapId,
    timestamp: new Date().toISOString(),
    note: note,
    masterCount: master.length,
    dailyCount: daily.length,
    data: { master, daily }
  };

  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction('snapshots', 'readwrite');
    const store = tx.objectStore('snapshots');
    store.put(snapshot);
    tx.oncomplete = () => resolve(snapshot);
    tx.onerror = () => reject(tx.error);
  });
}

/**
 * Universal Bulk Import Rows parsed from Excel
 * - Updates/Inserts Master Data: PN, QTY, Min, Max, PS/NON PS, LIB/NON LIB, Lokasi, Nama, Price
 * - Records Daily Transactions if IN or OUT are provided
 * - Realtime cascade recalculation
 */
async function processBulkImport(rows, defaultDate, user = 'Team') {
  if (!rows || rows.length === 0) {
    return { success: false, message: 'Tidak ada baris data untuk diimport.' };
  }

  const snapshot = await createSnapshot(`Otomatis sebelum import ${rows.length} baris`);

  const nowIso = new Date().toISOString();
  const affectedCodes = new Set();
  let insertedMaterials = 0;
  let updatedMaterials = 0;
  let insertedTx = 0;
  let updatedTx = 0;
  const errors = [];

  const allMaster = await getAllMasterMaterials();
  const masterMap = new Map();
  allMaster.forEach(m => masterMap.set(m.kodeMaterial.toUpperCase(), m));

  const allDaily = await getAllDailyStock();
  const dailyMap = new Map();
  allDaily.forEach(t => dailyMap.set(`${t.kodeMaterial.toUpperCase()}_${t.tanggal}`, t));

  const tx = dbInstance.transaction(['master_material', 'daily_stock', 'audit_logs'], 'readwrite');
  const masterStore = tx.objectStore('master_material');
  const dailyStore = tx.objectStore('daily_stock');
  const auditStore = tx.objectStore('audit_logs');

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const kode = String(row.kodeMaterial || '').trim();
    if (!kode) {
      errors.push({ row: i + 1, error: 'Kode Material kosong' });
      continue;
    }

    const kodeUpper = kode.toUpperCase();
    affectedCodes.add(kode);
    const rowDate = row.tanggal || defaultDate || new Date().toISOString().split('T')[0];
    const inQty = Number(row.inQty) || 0;
    const outQty = Number(row.outQty) || 0;
    const utd = row.utd !== undefined ? String(row.utd).trim() : '';

    // Direct QTY or initial stock
    const directQty = (row.qty !== undefined && row.qty !== null && row.qty !== '') ? Number(row.qty) : null;
    const minVal = (row.min !== undefined && row.min !== null && row.min !== '') ? Number(row.min) : null;
    const maxVal = (row.max !== undefined && row.max !== null && row.max !== '') ? Number(row.max) : null;
    const statusVal = row.status || row.ps || null;
    const libVal = row.libCategory || row.lib || null;

    // A. Handle Master Material (UPSERT)
    let mat = masterMap.get(kodeUpper);
    if (!mat) {
      // New Material
      const initStock = directQty !== null ? directQty : (inQty - outQty > 0 ? (inQty - outQty) : 0);
      mat = {
        no: masterMap.size + 1,
        kodeMaterial: kode,
        lokasi: row.lokasi ? String(row.lokasi).trim() : 'GUDANG UTAMA',
        namaMaterial: row.namaMaterial ? String(row.namaMaterial).trim() : 'Part ' + kode,
        price: Number(row.price) || 0,
        stokAwal: initStock,
        stokSaatIni: initStock,
        prod: row.prod || 'GENERAL',
        modelUnit: row.modelUnit || '',
        min: minVal !== null ? minVal : 1,
        max: maxVal !== null ? maxVal : 5,
        status: statusVal ? String(statusVal).trim().toUpperCase() : 'PS',
        libCategory: libVal ? String(libVal).trim().toUpperCase() : 'NON LIB',
        doi: null,
        remarks: row.remarks || '',
        alternatePn: '',
        createdAt: nowIso,
        updatedAt: nowIso
      };
      masterMap.set(kodeUpper, mat);
      masterStore.put(mat);
      insertedMaterials++;
    } else {
      // Existing Material: update any provided fields
      let modified = false;
      if (row.lokasi && row.lokasi !== mat.lokasi) { mat.lokasi = String(row.lokasi).trim(); modified = true; }
      if (row.namaMaterial && row.namaMaterial !== mat.namaMaterial) { mat.namaMaterial = String(row.namaMaterial).trim(); modified = true; }
      if (row.price !== undefined && Number(row.price) > 0 && Number(row.price) !== mat.price) { mat.price = Number(row.price); modified = true; }
      if (minVal !== null && minVal !== mat.min) { mat.min = minVal; modified = true; }
      if (maxVal !== null && maxVal !== mat.max) { mat.max = maxVal; modified = true; }
      if (statusVal && String(statusVal).trim().toUpperCase() !== mat.status) { mat.status = String(statusVal).trim().toUpperCase(); modified = true; }
      if (libVal && String(libVal).trim().toUpperCase() !== mat.libCategory) { mat.libCategory = String(libVal).trim().toUpperCase(); modified = true; }
      
      // If direct QTY is given in Excel, update physical stock directly
      if (directQty !== null && directQty !== mat.stokSaatIni) {
        mat.stokSaatIni = directQty;
        mat.stokAwal = directQty;
        modified = true;
      }

      if (modified) {
        mat.updatedAt = nowIso;
        masterStore.put(mat);
        updatedMaterials++;
      }
    }

    // B. Handle Daily Stock Transaction if IN or OUT > 0 or explicitly provided
    if (inQty > 0 || outQty > 0 || utd || row.hasTransaction) {
      const txKey = `${kodeUpper}_${rowDate}`;
      const existingTx = dailyMap.get(txKey);

      if (existingTx) {
        existingTx.inQty = inQty;
        existingTx.outQty = outQty;
        if (utd) existingTx.utd = utd;
        existingTx.updatedAt = nowIso;
        existingTx.updatedBy = user;
        dailyStore.put(existingTx);
        updatedTx++;
      } else {
        const newTx = {
          id: 'TX-' + Date.now() + '-' + Math.random().toString(36).substr(2, 6) + '-' + i,
          tanggal: rowDate,
          jam: new Date().toTimeString().split(' ')[0].substring(0, 5),
          kodeMaterial: kode,
          namaMaterial: mat.namaMaterial,
          lokasi: mat.lokasi,
          inQty: inQty,
          outQty: outQty,
          utd: utd,
          sisa: 0,
          notes: row.notes || 'Bulk Import Excel',
          createdAt: nowIso,
          updatedAt: nowIso,
          updatedBy: user
        };
        dailyMap.set(txKey, newTx);
        dailyStore.put(newTx);
        insertedTx++;
      }
    }
  }

  // Audit log for bulk import
  auditStore.put({
    id: 'LOG-' + Date.now(),
    timestamp: nowIso,
    user: user,
    action: 'BULK_IMPORT',
    kodeMaterial: `${affectedCodes.size} Materials`,
    details: `Universal bulk import ${rows.length} baris: +${insertedMaterials} material baru, ~${updatedMaterials} material update, +${insertedTx} transaksi baru, ~${updatedTx} transaksi update.`
  });

  await new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });

  // Run cascade recalculation for materials with transactions
  for (const code of affectedCodes) {
    try {
      await recalculateStockForMaterial(code);
    } catch (e) {}
  }

  return {
    success: true,
    totalRows: rows.length,
    insertedMaterials,
    updatedMaterials,
    insertedTx,
    updatedTx,
    affectedMaterials: affectedCodes.size,
    errors,
    snapshotId: snapshot.id
  };
}

/**
 * ============================================================================
 * REPORTING & STOCK AS OF DATE
 * ============================================================================
 */
async function getStockAsOfDate(targetDate) {
  if (!targetDate) {
    const today = new Date().toISOString().split('T')[0];
    targetDate = today;
  }

  const allMaster = await getAllMasterMaterials();
  const allDaily = await getAllDailyStock();

  const txByMaterial = new Map();
  for (const t of allDaily) {
    if (t.tanggal <= targetDate) {
      const k = t.kodeMaterial.toUpperCase();
      if (!txByMaterial.has(k)) {
        txByMaterial.set(k, []);
      }
      txByMaterial.get(k).push(t);
    }
  }

  const report = allMaster.map(mat => {
    const txs = txByMaterial.get(mat.kodeMaterial.toUpperCase());
    let stockOnDate = Number(mat.stokAwal) || 0;
    let totalIn = 0;
    let totalOut = 0;
    let lastTxDate = null;

    if (txs && txs.length > 0) {
      txs.sort((a, b) => a.tanggal.localeCompare(b.tanggal));
      for (const t of txs) {
        totalIn += Number(t.inQty) || 0;
        totalOut += Number(t.outQty) || 0;
      }
      const lastTx = txs[txs.length - 1];
      stockOnDate = Number(lastTx.sisa);
      lastTxDate = lastTx.tanggal;
    } else {
      stockOnDate = Number(mat.stokSaatIni) || 0;
    }

    const price = Number(mat.price) || 0;
    const totalValuation = stockOnDate * price;

    return {
      kodeMaterial: mat.kodeMaterial,
      namaMaterial: mat.namaMaterial,
      lokasi: mat.lokasi,
      price: price,
      stokAwal: mat.stokAwal,
      totalInToDate: totalIn,
      totalOutToDate: totalOut,
      stokOnDate: stockOnDate,
      status: mat.status || 'PS',
      libCategory: mat.libCategory || 'NON LIB',
      lastTxDate: lastTxDate || 'Belum ada transaksi',
      totalValuation: totalValuation,
      prod: mat.prod,
      modelUnit: mat.modelUnit
    };
  });

  return report;
}

/**
 * ============================================================================
 * AUDIT LOGS & BACKUP RESTORE
 * ============================================================================
 */
function getAllAuditLogs() {
  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction('audit_logs', 'readonly');
    const store = tx.objectStore('audit_logs');
    const req = store.getAll();
    req.onsuccess = () => {
      const logs = req.result || [];
      logs.sort((a, b) => (b.timestamp || '').localeCompare(a.timestamp || ''));
      resolve(logs);
    };
    req.onerror = () => reject(req.error);
  });
}

async function createDatabaseBackup() {
  const master = await getAllMasterMaterials();
  const daily = await getAllDailyStock();
  const audit = await getAllAuditLogs();

  return {
    version: DB_VERSION,
    appName: 'PT. STAR WAGEN INDONESIA - VIB',
    exportedAt: new Date().toISOString(),
    stats: {
      masterCount: master.length,
      dailyCount: daily.length,
      auditCount: audit.length
    },
    data: {
      master_material: master,
      daily_stock: daily,
      audit_logs: audit
    }
  };
}

async function restoreDatabaseBackup(backupJson, user = 'Team') {
  if (!backupJson || !backupJson.data || !Array.isArray(backupJson.data.master_material)) {
    throw new Error('Format file backup tidak valid.');
  }

  await createSnapshot('Otomatis sebelum restore database');

  const { master_material, daily_stock, audit_logs } = backupJson.data;

  const tx = dbInstance.transaction(['master_material', 'daily_stock', 'audit_logs'], 'readwrite');
  const masterStore = tx.objectStore('master_material');
  const dailyStore = tx.objectStore('daily_stock');
  const auditStore = tx.objectStore('audit_logs');

  masterStore.clear();
  dailyStore.clear();
  auditStore.clear();

  for (const m of (master_material || [])) {
    masterStore.put(m);
  }
  for (const d of (daily_stock || [])) {
    dailyStore.put(d);
  }
  for (const a of (audit_logs || [])) {
    auditStore.put(a);
  }

  auditStore.put({
    id: 'LOG-' + Date.now(),
    timestamp: new Date().toISOString(),
    user: user,
    action: 'RESTORE_BACKUP',
    kodeMaterial: 'ALL',
    details: `Database dipulihkan: ${master_material.length} master items, ${(daily_stock || []).length} transaksi harian.`
  });

  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
  });
}

// Global Export
window.InventoryDB = {
  init: initDatabase,
  getAllMasterMaterials,
  getMasterMaterial,
  upsertMasterMaterial,
  bulkUpdateMasterMaterials,
  deleteMasterMaterial,
  getAllDailyStock,
  getDailyStockByMaterial,
  getDailyStockById,
  upsertDailyStock,
  deleteDailyStock,
  recalculateStockForMaterial,
  processBulkImport,
  getStockAsOfDate,
  getAllAuditLogs,
  createDatabaseBackup,
  restoreDatabaseBackup,
  createSnapshot
};
