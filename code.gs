/**
 * Google Apps Script – Rohis Registration Backend (Versi Kaka Kelas)
 *
 * Spreadsheet layout:
 *   • Sheet "Pendaftar"          – Kolom: Nama, Kelas, WA, Kaka Kelas
 *   • Sheet "Daftar Kaka Kelas"  – Kolom A = nama-nama Kaka Kelas
 *
 * Deploy as Web App (Execute as Me, Access: Anyone, even anonymous).
 * Endpoints:
 *   GET  ?action=kaka_kelas  → { kaka_kelas: ["Ali","Budi",...] }
 *   GET  ?action=summary     → [{ kaka_kelas: "Ali", count: 12 }, ...]
 *   POST (JSON)              → append row to "Pendaftar"
 */

/** Helper: fetch a sheet by name */
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet "${name}" not found`);
  return sheet;
}

/** GET handler – returns kaka kelas list, summary, or serves the HTML homepage */
function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();

  if (action === 'kaka_kelas') {
    return ContentService.createTextOutput(JSON.stringify({ kaka_kelas: getKakaKelas() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'summary') {
    return ContentService.createTextOutput(JSON.stringify(getKakaKelasSummary()))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  // Serve the index.html file as the web app homepage
  return HtmlService.createHtmlOutputFromFile('index')
                    .setTitle('Pendaftaran & Dashboard Rohis SMAN 27')
                    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
                    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** POST handler – receives registration data */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    // basic validation
    if (!payload.nama || !payload.kelas || !payload.wa || !payload.kaka_kelas) {
      throw new Error('Data tidak lengkap');
    }
    const sheet = getSheet('Pendaftar');
    sheet.appendRow([payload.nama, payload.kelas, payload.wa, payload.kaka_kelas]);
    // plain‑text response works with fetch(no‑cors)
    return ContentService.createTextOutput('OK')
                         .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput('Error: ' + err.message)
                         .setMimeType(ContentService.MimeType.TEXT);
  }
}

/** -----------------------------------------------------------------
 *  Helper: get array of kaka kelas names from sheet "Daftar Kaka Kelas"
 * ----------------------------------------------------------------- */
function getKakaKelas() {
  const sheet = getSheet('Daftar Kaka Kelas');
  const values = sheet.getRange('A2:A' + sheet.getLastRow()).getValues(); // skip header
  return values.map(r => r[0]).filter(name => name);
}

/** -----------------------------------------------------------------
 *  Helper: compute referral counts per kaka kelas from "Pendaftar"
 *  Returns sorted descending array [{ kaka_kelas, count }, …]
 * ----------------------------------------------------------------- */
function getKakaKelasSummary() {
  const sheet = getSheet('Pendaftar');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // Nama, Kelas, WA, Kaka Kelas
  const map = {};
  data.forEach(row => {
    const nama = row[0];
    const kelas = row[1];
    const name = row[3];
    if (!name) return;
    if (!map[name]) {
      map[name] = { kaka_kelas: name, count: 0, siswa_list: [] };
    }
    map[name].count++;
    map[name].siswa_list.push({ nama: nama, kelas: kelas });
  });
  const result = Object.keys(map).map(k => map[k]);
  result.sort((a, b) => b.count - a.count);
  return result;
}
