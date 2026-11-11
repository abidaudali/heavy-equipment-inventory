/**
 * ============================================================================
 * GOOGLE APPS SCRIPT: PT. STAR WAGEN INDONESIA - VIB Cloud Database Backend
 * ============================================================================
 * Mendukung 3 Sheet Tersinkronisasi Otomatis:
 * 1. MASTER_MATERIAL (Master Data Spare Part, QTY, Min/Max, PS/NON PS, LIB/NON LIB)
 * 2. DAILY_STOCK (Pencatatan Transaksi Stok Harian IN / OUT / UTD / SISA)
 * 3. AUDIT_LOGS (Riwayat Lengkap Perubahan Data & Keamanan)
 * ============================================================================
 */

const SHEET_MASTER = "MASTER_MATERIAL";
const SHEET_DAILY = "DAILY_STOCK";
const SHEET_AUDIT = "AUDIT_LOGS";

// Header Kolom Standar
const HEADERS_MASTER = [
  "No",
  "Kode Material",
  "Lokasi",
  "Nama Material",
  "Price",
  "Stok Awal",
  "Stok Saat Ini",
  "Min",
  "Max",
  "Status Servis (PS)",
  "Kategori LIB",
  "Prod",
  "Model Unit",
  "Remarks",
  "Last Updated"
];

const HEADERS_DAILY = [
  "ID Transaksi",
  "Tanggal",
  "Jam",
  "Kode Material",
  "Nama Material",
  "Lokasi",
  "IN",
  "OUT",
  "UTD",
  "SISA",
  "Catatan",
  "Updated By",
  "Last Updated"
];

const HEADERS_AUDIT = [
  "Log ID",
  "Timestamp",
  "User",
  "Aksi",
  "Kode Material",
  "Detail Perubahan"
];

/**
 * Handle GET Requests
 */
function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);

  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getAllData";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    ensureAllSheets(ss);

    if (action === "ping") {
      return jsonResponse({
        status: "success",
        message: "Google Apps Script Inventory API is online & active!",
        spreadsheetTitle: ss.getName(),
        spreadsheetUrl: ss.getUrl()
      });
    }

    if (action === "getAllData" || action === "getData") {
      const masterData = getMasterSheetData(ss.getSheetByName(SHEET_MASTER));
      const dailyData = getDailySheetData(ss.getSheetByName(SHEET_DAILY));
      const auditData = getAuditSheetData(ss.getSheetByName(SHEET_AUDIT));

      return jsonResponse({
        status: "success",
        counts: {
          master: masterData.length,
          daily: dailyData.length,
          audit: auditData.length
        },
        data: {
          master: masterData,
          daily: dailyData,
          audit: auditData
        }
      });
    }

    if (action === "getMasterData") {
      const data = getMasterSheetData(ss.getSheetByName(SHEET_MASTER));
      return jsonResponse({ status: "success", count: data.length, data: data });
    }

    if (action === "getDailyStock") {
      const data = getDailySheetData(ss.getSheetByName(SHEET_DAILY));
      return jsonResponse({ status: "success", count: data.length, data: data });
    }

    return jsonResponse({ status: "error", message: "Action tidak dikenali: " + action });

  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handle POST Requests
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(20000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: "error", message: "Tidak ada data POST yang diterima." });
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureAllSheets(ss);

    const nowStr = new Date().toLocaleString("id-ID");

    // 1. SYNC ALL (Bulk Push dari website ke Google Sheets)
    if (action === "syncAll" || action === "bulkSync") {
      const masterItems = payload.master || payload.items || [];
      const dailyItems = payload.daily || [];
      const auditItems = payload.audit || [];

      // A. Sync Master Material
      const sheetM = ss.getSheetByName(SHEET_MASTER);
      sheetM.clearContents();
      sheetM.appendRow(HEADERS_MASTER);

      if (masterItems.length > 0) {
        const rowsM = masterItems.map((m, idx) => [
          m.no || (idx + 1),
          String(m.kodeMaterial || m.pnFinal || "").trim(),
          String(m.lokasi || "").trim(),
          String(m.namaMaterial || m.desc || "").trim(),
          Number(m.price) || 0,
          Number(m.stokAwal !== undefined ? m.stokAwal : (m.qty || 0)),
          Number(m.stokSaatIni !== undefined ? m.stokSaatIni : (m.qty || 0)),
          m.min !== undefined && m.min !== null ? Number(m.min) : 1,
          m.max !== undefined && m.max !== null ? Number(m.max) : 5,
          String(m.status || "PS").trim().toUpperCase(),
          String(m.libCategory || "NON LIB").trim().toUpperCase(),
          m.prod || "GENERAL",
          m.modelUnit || "",
          m.remarks || "",
          nowStr
        ]);
        sheetM.getRange(2, 1, rowsM.length, HEADERS_MASTER.length).setValues(rowsM);
      }
      formatHeader(sheetM, HEADERS_MASTER.length, "#1e293b");

      // B. Sync Daily Stock
      const sheetD = ss.getSheetByName(SHEET_DAILY);
      sheetD.clearContents();
      sheetD.appendRow(HEADERS_DAILY);

      if (dailyItems.length > 0) {
        const rowsD = dailyItems.map(d => [
          d.id || ("TX-" + Date.now()),
          d.tanggal || "",
          d.jam || "",
          String(d.kodeMaterial || "").trim(),
          String(d.namaMaterial || "").trim(),
          String(d.lokasi || "").trim(),
          Number(d.inQty) || 0,
          Number(d.outQty) || 0,
          String(d.utd || ""),
          Number(d.sisa) || 0,
          String(d.notes || ""),
          d.updatedBy || "Team",
          nowStr
        ]);
        sheetD.getRange(2, 1, rowsD.length, HEADERS_DAILY.length).setValues(rowsD);
      }
      formatHeader(sheetD, HEADERS_DAILY.length, "#065f46");

      // C. Sync Audit Logs
      const sheetA = ss.getSheetByName(SHEET_AUDIT);
      sheetA.clearContents();
      sheetA.appendRow(HEADERS_AUDIT);

      if (auditItems.length > 0) {
        const rowsA = auditItems.slice(0, 500).map(a => [
          a.id || ("LOG-" + Date.now()),
          a.timestamp || nowStr,
          a.user || "Team",
          a.action || "UPDATE",
          a.kodeMaterial || "",
          a.details || ""
        ]);
        sheetA.getRange(2, 1, rowsA.length, HEADERS_AUDIT.length).setValues(rowsA);
      }
      formatHeader(sheetA, HEADERS_AUDIT.length, "#4c1d95");

      return jsonResponse({
        status: "success",
        message: `Sinkronisasi berhasil! ${masterItems.length} Master Material (PS & LIB), ${dailyItems.length} Transaksi Harian, dan ${auditItems.length} Audit Logs tersimpan.`
      });
    }

    // 2. UPSERT MASTER MATERIAL
    if (action === "upsertMaster") {
      const item = payload.item;
      const sheetM = ss.getSheetByName(SHEET_MASTER);
      const rows = sheetM.getDataRange().getValues();
      const kodeTarget = String(item.kodeMaterial || item.pnFinal || "").trim();
      let foundIndex = -1;

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][1]).trim() === kodeTarget) {
          foundIndex = i + 1;
          break;
        }
      }

      const rowData = [
        item.no || (foundIndex > 0 ? rows[foundIndex - 1][0] : sheetM.getLastRow()),
        kodeTarget,
        item.lokasi || "",
        item.namaMaterial || item.desc || "",
        Number(item.price) || 0,
        Number(item.stokAwal) || 0,
        Number(item.stokSaatIni !== undefined ? item.stokSaatIni : (item.qty || 0)),
        item.min !== undefined ? Number(item.min) : 1,
        item.max !== undefined ? Number(item.max) : 5,
        String(item.status || "PS").trim().toUpperCase(),
        String(item.libCategory || "NON LIB").trim().toUpperCase(),
        item.prod || "GENERAL",
        item.modelUnit || "",
        item.remarks || "",
        nowStr
      ];

      if (foundIndex > 0) {
        sheetM.getRange(foundIndex, 1, 1, HEADERS_MASTER.length).setValues([rowData]);
        return jsonResponse({ status: "success", message: `Material ${kodeTarget} berhasil diperbarui.` });
      } else {
        sheetM.appendRow(rowData);
        return jsonResponse({ status: "success", message: `Material baru ${kodeTarget} berhasil ditambahkan.` });
      }
    }

    // 3. UPSERT DAILY STOCK TRANSACTION
    if (action === "upsertDailyStock") {
      const tx = payload.tx;
      const sheetD = ss.getSheetByName(SHEET_DAILY);
      const rows = sheetD.getDataRange().getValues();
      let foundIndex = -1;

      for (let i = 1; i < rows.length; i++) {
        if (tx.id && String(rows[i][0]).trim() === String(tx.id).trim()) {
          foundIndex = i + 1;
          break;
        }
        if (String(rows[i][3]).trim() === String(tx.kodeMaterial).trim() && String(rows[i][1]).trim() === String(tx.tanggal).trim()) {
          foundIndex = i + 1;
          break;
        }
      }

      const rowData = [
        tx.id || ("TX-" + Date.now()),
        tx.tanggal || "",
        tx.jam || "",
        String(tx.kodeMaterial || "").trim(),
        String(tx.namaMaterial || "").trim(),
        String(tx.lokasi || "").trim(),
        Number(tx.inQty) || 0,
        Number(tx.outQty) || 0,
        String(tx.utd || ""),
        Number(tx.sisa) || 0,
        String(tx.notes || ""),
        tx.updatedBy || "Team",
        nowStr
      ];

      if (foundIndex > 0) {
        sheetD.getRange(foundIndex, 1, 1, HEADERS_DAILY.length).setValues([rowData]);
        return jsonResponse({ status: "success", message: `Transaksi ${tx.kodeMaterial} tanggal ${tx.tanggal} berhasil diperbarui.` });
      } else {
        sheetD.appendRow(rowData);
        return jsonResponse({ status: "success", message: `Transaksi baru ${tx.kodeMaterial} tanggal ${tx.tanggal} berhasil dicatat.` });
      }
    }

    // 4. LOG AUDIT
    if (action === "logAudit") {
      const log = payload.log;
      const sheetA = ss.getSheetByName(SHEET_AUDIT);
      sheetA.appendRow([
        log.id || ("LOG-" + Date.now()),
        log.timestamp || nowStr,
        log.user || "Team",
        log.action || "INFO",
        log.kodeMaterial || "",
        log.details || ""
      ]);
      return jsonResponse({ status: "success" });
    }

    return jsonResponse({ status: "error", message: "Action POST tidak dikenali: " + action });

  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Pastikan 3 Sheet Tersedia & Terformat
 */
function ensureAllSheets(ss) {
  let sheetM = ss.getSheetByName(SHEET_MASTER);
  if (!sheetM) {
    sheetM = ss.insertSheet(SHEET_MASTER);
    sheetM.appendRow(HEADERS_MASTER);
    formatHeader(sheetM, HEADERS_MASTER.length, "#1e293b");
  }

  let sheetD = ss.getSheetByName(SHEET_DAILY);
  if (!sheetD) {
    sheetD = ss.insertSheet(SHEET_DAILY);
    sheetD.appendRow(HEADERS_DAILY);
    formatHeader(sheetD, HEADERS_DAILY.length, "#065f46");
  }

  let sheetA = ss.getSheetByName(SHEET_AUDIT);
  if (!sheetA) {
    sheetA = ss.insertSheet(SHEET_AUDIT);
    sheetA.appendRow(HEADERS_AUDIT);
    formatHeader(sheetA, HEADERS_AUDIT.length, "#4c1d95");
  }
}

/**
 * Format Header Row
 */
function formatHeader(sheet, numCols, bgColor) {
  const range = sheet.getRange(1, 1, 1, numCols);
  range.setBackground(bgColor);
  range.setFontColor("#ffffff");
  range.setFontWeight("bold");
  range.setHorizontalAlignment("center");
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
}

/**
 * Helper: Ambil data Sheet Master
 */
function getMasterSheetData(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return [];

  const items = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[1] && !row[3]) continue;

    items.push({
      no: row[0] ? Number(row[0]) : i,
      kodeMaterial: String(row[1] || "").trim(),
      lokasi: String(row[2] || "").trim(),
      namaMaterial: String(row[3] || "").trim(),
      price: Number(row[4]) || 0,
      stokAwal: Number(row[5]) || 0,
      stokSaatIni: Number(row[6]) || 0,
      min: row[7] !== "" ? Number(row[7]) : 1,
      max: row[8] !== "" ? Number(row[8]) : 5,
      status: String(row[9] || "PS").trim().toUpperCase(),
      libCategory: String(row[10] || "NON LIB").trim().toUpperCase(),
      prod: String(row[11] || "GENERAL").trim(),
      modelUnit: String(row[12] || "").trim(),
      remarks: String(row[13] || "").trim(),
      lastUpdated: row[14] || ""
    });
  }
  return items;
}

/**
 * Helper: Ambil data Sheet Daily Stock
 */
function getDailySheetData(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return [];

  const items = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[3]) continue;

    let tglStr = String(row[1] || "").trim();
    if (row[1] instanceof Date) {
      tglStr = Utilities.formatDate(row[1], "GMT+7", "yyyy-MM-dd");
    }

    items.push({
      id: String(row[0] || ("TX-" + i)).trim(),
      tanggal: tglStr,
      jam: String(row[2] || "").trim(),
      kodeMaterial: String(row[3] || "").trim(),
      namaMaterial: String(row[4] || "").trim(),
      lokasi: String(row[5] || "").trim(),
      inQty: Number(row[6]) || 0,
      outQty: Number(row[7]) || 0,
      utd: String(row[8] || "").trim(),
      sisa: Number(row[9]) || 0,
      notes: String(row[10] || "").trim(),
      updatedBy: String(row[11] || "").trim(),
      lastUpdated: row[12] || ""
    });
  }
  return items;
}

/**
 * Helper: Ambil data Sheet Audit
 */
function getAuditSheetData(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return [];

  const items = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    items.push({
      id: String(row[0] || "").trim(),
      timestamp: String(row[1] || "").trim(),
      user: String(row[2] || "").trim(),
      action: String(row[3] || "").trim(),
      kodeMaterial: String(row[4] || "").trim(),
      details: String(row[5] || "").trim()
    });
  }
  return items;
}

/**
 * Helper: Output JSON Response
 */
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
