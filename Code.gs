/**
 * ============================================================================
 * GOOGLE APPS SCRIPT: Heavy Equipment Spare Parts Inventory Backend
 * ============================================================================
 * Petunjuk Instalasi Singkat:
 * 1. Buat Google Sheet baru di Google Drive (https://sheets.new)
 * 2. Klik menu "Extensions" (Ekstensi) > "Apps Script"
 * 3. Hapus kode bawaan, lalu paste seluruh isi file ini ke dalamnya.
 * 4. Jalankan fungsi "setupSheet" sekali untuk membuat header dan format otomatis.
 * 5. Klik "Deploy" (Terapkan) > "New deployment" (Penerapan baru).
 *    - Select type: Web app (Aplikasi Web)
 *    - Description: Heavy Equipment Inventory API
 *    - Execute as: Me (Email Anda)
 *    - Who has access: Anyone (Siapa saja)  <-- PENTING agar web dapat mengakses
 * 6. Klik "Deploy", izinkan akses akun Google Anda.
 * 7. Salin "Web App URL" yang diberikan, lalu paste di menu pengaturan Web Dashboard!
 * ============================================================================
 */

const SHEET_NAME = "INVENTORY";

// Header kolom standar
const HEADERS = [
  "No",
  "Prod",
  "Model Unit",
  "Part Number",
  "Description",
  "Min",
  "Max",
  "Status",
  "Qty",
  "Month",
  "Week",
  "Day",
  "DOI",
  "Remarks",
  "Alternate PN",
  "Last Updated"
];

/**
 * Handle GET Requests
 * Contoh: 
 * - ?action=getData
 * - ?action=ping
 * - ?action=getStats
 */
function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getData";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = setupSheet();
    }

    if (action === "ping") {
      return jsonResponse({
        status: "success",
        message: "Google Apps Script API is active & connected!",
        spreadsheetTitle: ss.getName(),
        spreadsheetUrl: ss.getUrl()
      });
    }

    if (action === "getData") {
      const data = getSheetData(sheet);
      return jsonResponse({
        status: "success",
        count: data.length,
        data: data
      });
    }

    return jsonResponse({ status: "error", message: "Unknown action: " + action });

  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handle POST Requests
 * Digunakan untuk:
 * - updateStock (menambah / mengurangi stok)
 * - addItem (menambah part baru)
 * - editItem (mengubah data part)
 * - deleteItem (menghapus part)
 * - bulkSync (mengirim seluruh data dari web ke Google Sheet)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: "error", message: "No post data received" });
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = setupSheet();
    }

    // 1. UPDATE STOCK
    if (action === "updateStock") {
      const pn = payload.pnFinal;
      const newQty = Number(payload.qty);
      const note = payload.note || "Quick Stock Update";

      const rows = sheet.getDataRange().getValues();
      let found = false;

      for (let i = 1; i < rows.length; i++) {
        const rowPn = String(rows[i][3]).trim();
        if (rowPn === String(pn).trim()) {
          sheet.getRange(i + 1, 9).setValue(newQty); // Kolom Qty (kolom ke-9)
          sheet.getRange(i + 1, 16).setValue(new Date().toLocaleString("id-ID")); // Last Updated
          found = true;
          break;
        }
      }

      if (!found) {
        return jsonResponse({ status: "error", message: "Part number " + pn + " tidak ditemukan di Google Sheet." });
      }

      return jsonResponse({ status: "success", message: "Stok part " + pn + " berhasil diperbarui menjadi " + newQty });
    }

    // 2. ADD ITEM
    if (action === "addItem") {
      const item = payload.item;
      const lastRow = sheet.getLastRow();
      const newNo = lastRow > 1 ? Number(sheet.getRange(lastRow, 1).getValue()) + 1 : 1;

      sheet.appendRow([
        item.no || newNo,
        item.prod || "",
        item.modelUnit || "",
        item.pnFinal || "",
        item.desc || "",
        item.min !== undefined && item.min !== null ? item.min : "",
        item.max !== undefined && item.max !== null ? item.max : "",
        item.status || "PS",
        item.qty !== undefined && item.qty !== null ? Number(item.qty) : 0,
        item.month !== undefined ? item.month : "",
        item.week !== undefined ? item.week : "",
        item.day !== undefined ? item.day : "",
        item.doi !== undefined ? item.doi : "",
        item.remarks || "",
        item.alternatePn || "",
        new Date().toLocaleString("id-ID")
      ]);

      return jsonResponse({ status: "success", message: "Part baru berhasil ditambahkan!", item: item });
    }

    // 3. EDIT ITEM
    if (action === "editItem") {
      const item = payload.item;
      const pn = item.pnFinal;
      const rows = sheet.getDataRange().getValues();
      let foundIndex = -1;

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][3]).trim() === String(pn).trim() || (item.originalPn && String(rows[i][3]).trim() === String(item.originalPn).trim())) {
          foundIndex = i + 1;
          break;
        }
      }

      if (foundIndex === -1) {
        return jsonResponse({ status: "error", message: "Part tidak ditemukan untuk diedit." });
      }

      sheet.getRange(foundIndex, 2).setValue(item.prod || "");
      sheet.getRange(foundIndex, 3).setValue(item.modelUnit || "");
      sheet.getRange(foundIndex, 4).setValue(item.pnFinal || "");
      sheet.getRange(foundIndex, 5).setValue(item.desc || "");
      sheet.getRange(foundIndex, 6).setValue(item.min !== null ? item.min : "");
      sheet.getRange(foundIndex, 7).setValue(item.max !== null ? item.max : "");
      sheet.getRange(foundIndex, 8).setValue(item.status || "PS");
      sheet.getRange(foundIndex, 9).setValue(item.qty !== null ? Number(item.qty) : 0);
      sheet.getRange(foundIndex, 10).setValue(item.month !== null ? item.month : "");
      sheet.getRange(foundIndex, 11).setValue(item.week !== null ? item.week : "");
      sheet.getRange(foundIndex, 12).setValue(item.day !== null ? item.day : "");
      sheet.getRange(foundIndex, 13).setValue(item.doi !== null ? item.doi : "");
      sheet.getRange(foundIndex, 14).setValue(item.remarks || "");
      sheet.getRange(foundIndex, 15).setValue(item.alternatePn || "");
      sheet.getRange(foundIndex, 16).setValue(new Date().toLocaleString("id-ID"));

      return jsonResponse({ status: "success", message: "Data part berhasil diperbarui!" });
    }

    // 4. DELETE ITEM
    if (action === "deleteItem") {
      const pn = payload.pnFinal;
      const rows = sheet.getDataRange().getValues();
      let foundIndex = -1;

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][3]).trim() === String(pn).trim()) {
          foundIndex = i + 1;
          break;
        }
      }

      if (foundIndex === -1) {
        return jsonResponse({ status: "error", message: "Part tidak ditemukan untuk dihapus." });
      }

      sheet.deleteRow(foundIndex);
      return jsonResponse({ status: "success", message: "Part " + pn + " berhasil dihapus dari Google Sheet." });
    }

    // 5. BULK SYNC (Import seluruh data dari website ke Google Sheet)
    if (action === "bulkSync" && Array.isArray(payload.items)) {
      sheet.clearContents();
      sheet.appendRow(HEADERS);

      const items = payload.items;
      const rowsToInsert = [];
      const nowStr = new Date().toLocaleString("id-ID");

      for (let i = 0; i < items.length; i++) {
        const it = items[i];
        rowsToInsert.push([
          it.no || (i + 1),
          it.prod || "",
          it.modelUnit || "",
          it.pnFinal || "",
          it.desc || "",
          it.min !== undefined && it.min !== null ? it.min : "",
          it.max !== undefined && it.max !== null ? it.max : "",
          it.status || "PS",
          it.qty !== undefined && it.qty !== null ? Number(it.qty) : 0,
          it.month !== undefined && it.month !== null ? it.month : "",
          it.week !== undefined && it.week !== null ? it.week : "",
          it.day !== undefined && it.day !== null ? it.day : "",
          it.doi !== undefined && it.doi !== null ? it.doi : "",
          it.remarks || "",
          it.alternatePn || "",
          nowStr
        ]);
      }

      if (rowsToInsert.length > 0) {
        sheet.getRange(2, 1, rowsToInsert.length, HEADERS.length).setValues(rowsToInsert);
      }

      formatSheetDesign(sheet);
      return jsonResponse({ status: "success", message: "Bulk sync berhasil: " + items.length + " item disinkronkan!" });
    }

    return jsonResponse({ status: "error", message: "Action tidak dikenali: " + action });

  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper: Ambil seluruh data sheet menjadi array objek JSON
 */
function getSheetData(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return [];

  const items = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[3] && !row[4]) continue; // Lewati baris kosong

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
      alternatePn: String(row[14] || "").trim(),
      lastUpdated: row[15] || ""
    });
  }

  return items;
}

/**
 * Format & Buat Sheet Baru jika belum ada
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  formatSheetDesign(sheet);
  return sheet;
}

/**
 * Styling tampilan Spreadsheet agar elegan dan mudah dibaca
 */
function formatSheetDesign(sheet) {
  // Header Style
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setBackground("#1e293b");
  headerRange.setFontColor("#f8fafc");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);

  // Auto resize column widths
  for (let c = 1; c <= HEADERS.length; c++) {
    sheet.autoResizeColumn(c);
  }
}

/**
 * Response Helper dengan header CORS JSON
 */
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
