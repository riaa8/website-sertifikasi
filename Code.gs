/**
 * CODE.GS - RESTORED & MERGED VERSION
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistem Informasi Sertifikasi')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

var SPREADSHEET_ID = '1Zd4plVMj7Z_UczDz8enSMYKI3AgD505AuuQdhNdPGqo'; 
var PESERTA_SPREADSHEET_ID = '1sALSUSLTYruo-F2wyAeoRlRqSpZe1a2ECjsTT9LG9wM';
var MAIN_SHEET_NAME_CAP = 'Perencanaan';
var MAIN_SHEET_NAME_LOWER = 'perencanaan';

function connect() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME_CAP) || ss.getSheetByName(MAIN_SHEET_NAME_LOWER);
  return sheet;
}

/* --- 1. DATA SERTIFIKASI (KIRI) --- */
function getData() {
  try {
    var sheet = connect();
    if (!sheet) return []; // Safety check
    
    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    var lastSAP = "";
    var lastNama = "";

    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (String(r[0]).toUpperCase() === "NO") continue;
        
        // Update last SAP/Nama if present
        if (r[1]) lastSAP = cleanString(r[1]);
        if (r[2]) lastNama = String(r[2]);

        // SKIP: Jika baris dianggap kosong total (tidak ada SAP/Nama dan tidak ada Item ID)
        if (!r[1] && !r[2] && !r[3]) continue;

        try {
            var certItemId = r[3]; // Kolom D
            
            if ((certItemId && String(certItemId).trim() !== "") || (r[1] && r[2])) {
                var id = r[0] ? String(r[0]) : "ROW_" + i;

                data.push({
                    id: id,          
                    sap: lastSAP, 
                    nama: lastNama,        
                    itemId: String(r[3]),      
                    judul: String(r[4]),       
                    periode: safeParseDate(r[5]), 
                    jumlah: String(r[6]),           
                    statusAnggaran: String(r[7]),   
                    mandatory: String(r[8]),        
                    resiko: String(r[9]),           
                    type: 'cert'
                });
            }
        } catch (rowErr) {
            Logger.log("Error processing CERT row " + i + ": " + rowErr);
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getData: ' + e.message);
    return []; // Return empty array to keep frontend running
  }
}

/* --- 2. DATA LAT (KANAN - Kolom L ke kanan) --- */
function getLATData() {
  try {
    var sheet = connect();
    if (!sheet) return [];

    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    var lastSAP = "";
    var lastNama = "";

    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (String(r[0]).toUpperCase() === "NO") continue;

        if (r[1]) lastSAP = cleanString(r[1]);
        if (r[2]) lastNama = String(r[2]);

        if (!r[1] && !r[2] && !r[11]) continue; 

        try {
            var latItemId = r[11];
            
            if ((latItemId && String(latItemId).trim() !== "") || (r[1] && r[2] && r[11])) {
                 var id = r[0] ? String(r[0]) : "ROW_" + i;
                 
                 data.push({
                    id: id + "_LAT", 
                    originalId: id,
                    sap: lastSAP,
                    nama: lastNama,
                    itemId: String(r[11]),     
                    judul: String(r[12]),      
                    instruktur: String(r[13]), 
                    periode: safeParseDate(r[14]),
                    resiko: String(r[15]),     
                    type: 'lat'
                });
            }
        } catch (rowErr) {
             Logger.log("Error processing LAT row " + i + ": " + rowErr);
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getLATData: ' + e.message);
    return [];
  }
}

// HELPER
function cleanString(val) {
  if (!val) return "";
  return String(val).trim().toUpperCase(); 
}

// SAFE PARSE DATE - Handles Indonesian format and returns YYYY-MM-DD
function safeParseDate(dateVal) {
  try {
      if (!dateVal) return "";
      
      // 1. Jika object Date (dari Excel date cell)
      if (Object.prototype.toString.call(dateVal) === '[object Date]') {
        var yyyy = dateVal.getFullYear();
        var mm = String(dateVal.getMonth() + 1).padStart(2, '0');
        var dd = String(dateVal.getDate()).padStart(2, '0');
        return yyyy + "-" + mm + "-" + dd;
      }
      
      var str = String(dateVal).trim();

      // 2. Handle Format "Bulan Tahun" (Contoh: "Maret 2026")
      var monthMap = {
        'JANUARI': '01', 'FEBRUARI': '02', 'MARET': '03', 'APRIL': '04', 'MEI': '05', 'JUNI': '06',
        'JULI': '07', 'AGUSTUS': '08', 'SEPTEMBER': '09', 'OKTOBER': '10', 'NOVEMBER': '11', 'DESEMBER': '12',
        'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'JUN': '06', 'JUL': '07', 'AGU': '08', 'SEP': '09', 'OKT': '10', 'NOV': '11', 'DES': '12'
      };
      
      // Cek apakah format "NamaBulan Tahun"
      var parts = str.split(' ');
      if (parts.length === 2) {
        var mName = parts[0].toUpperCase();
        var yName = parts[1];
        if (monthMap[mName] && !isNaN(yName)) {
           return yName + "-" + monthMap[mName] + "-01";
        }
      }
      
      // 3. Handle Format "D/M/YYYY" atau "M/D/YYYY" (Excel text format kadang begini)
      // Asumsi default Spreadsheet Indonesia: DD/MM/YYYY
      if (str.includes('/')) {
         var p = str.split('/');
         if (p.length === 3) {
            // Cek mana yang tahun (biasanya 4 digit)
            if (p[2].length === 4) return p[2] + "-" + String(p[1]).padStart(2,'0') + "-" + String(p[0]).padStart(2,'0');
            // Jika format english M/D/Y
            if (p[2].length === 2 && p[0].length === 4) return p[0] + "-" + String(p[1]).padStart(2,'0') + "-" + String(p[2]).padStart(2,'0'); 
         }
      }

      return str; 
  } catch (e) {
      return String(dateVal);
  }
}

function parseDate(d) { return safeParseDate(d); }

/* --- 3. CRUD (Update Mapping Save) --- */
/* --- 3. CRUD (DIPERBAIKI AGAR NOMOR BERURUTAN) --- */

// Helper untuk mendapatkan nomor urut selanjutnya
function getNextId(sheet) {
  var lastRow = sheet.getLastRow();
  
  // Jika baris hanya 1 (hanya header), mulai dari 1
  if (lastRow <= 1) return 1;

  // Ambil nilai dari kolom A baris terakhir
  var lastVal = sheet.getRange(lastRow, 1).getValue();

  // Pastikan nilainya angka, jika tidak (misal error), gunakan nomor baris
  var nextNum = parseInt(lastVal);
  if (isNaN(nextNum)) {
    return lastRow; // Fallback jika data berantakan
  }
  
  return nextNum + 1; // Nomor terakhir + 1
}

function addData(formObject) {
  var sheet = connect();
  
  var id = "=ROW()-1"; 
  
  var newRow = [
      id, 
      formObject.sap, 
      formObject.nama,
      formObject.itemId, 
      formObject.judul, 
      formObject.periode, 
      formObject.jumlah,         
      formObject.statusAnggaran, 
      formObject.mandatory,      
      formObject.resiko,         
      "", "", "", "", "", "" 
  ];
  sheet.appendRow(newRow);
  copyRowFormat(sheet, sheet.getLastRow() - 1, sheet.getLastRow());
  return { success: true };
}

function addLATData(formObject) {
    var sheet = connect();
    
    var id = "=ROW()-1";

    var newRow = [
        id, formObject.sap, formObject.nama,
        "", "", "", "", "", "", "", 
        "", 
        formObject.itemId, formObject.judul, formObject.instruktur, 
        formObject.periode, formObject.resiko
    ];
    sheet.appendRow(newRow);
    copyRowFormat(sheet, sheet.getLastRow() - 1, sheet.getLastRow());
    return { success: true };
}

/* --- UPDATE DATA SERTIFIKASI --- */
function updateData(formData) {
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(formData.id)) {
        rowIndex = i + 1; // 1-indexed untuk getRange
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data tidak ditemukan dengan ID: ' + formData.id };

    // Update kolom Cert: B(SAP), C(Nama), D(ItemId), E(Judul), F(Periode), G(Jumlah), H(StatusAnggaran), I(Mandatory), J(Resiko)
    sheet.getRange(rowIndex, 2).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 3).setValue(formData.nama || '');
    sheet.getRange(rowIndex, 4).setValue(formData.itemId || '');
    sheet.getRange(rowIndex, 5).setValue(formData.judul || '');
    sheet.getRange(rowIndex, 6).setValue(formData.periode || '');
    sheet.getRange(rowIndex, 7).setValue(formData.jumlah || '');
    sheet.getRange(rowIndex, 8).setValue(formData.statusAnggaran || '');
    sheet.getRange(rowIndex, 9).setValue(formData.mandatory || '');
    sheet.getRange(rowIndex, 10).setValue(formData.resiko || '');

    return { success: true };
  } catch (e) {
    Logger.log('Error in updateData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- UPDATE DATA LAT --- */
function updateLATData(formData) {
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    // ID LAT format: "X_LAT" — ambil original ID dengan hapus "_LAT"
    var originalId = String(formData.id).replace('_LAT', '');

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === originalId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data LAT tidak ditemukan dengan ID: ' + originalId };

    // Update kolom LAT: B(SAP), C(Nama), L(ItemId), M(Judul), N(Instruktur), O(Periode), P(Resiko)
    sheet.getRange(rowIndex, 2).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 3).setValue(formData.nama || '');
    sheet.getRange(rowIndex, 12).setValue(formData.itemId || '');
    sheet.getRange(rowIndex, 13).setValue(formData.judul || '');
    sheet.getRange(rowIndex, 14).setValue(formData.instruktur || '');
    sheet.getRange(rowIndex, 15).setValue(formData.periode || '');
    sheet.getRange(rowIndex, 16).setValue(formData.resiko || '');

    return { success: true };
  } catch (e) {
    Logger.log('Error in updateLATData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- DELETE DATA SERTIFIKASI --- */
function deleteData(id) {
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data tidak ditemukan dengan ID: ' + id };

    // Cek apakah baris ini juga punya data LAT (kolom L / index 11)
    var hasLAT = rows[rowIndex - 1][11] && String(rows[rowIndex - 1][11]).trim() !== '';

    if (hasLAT) {
      // Baris punya LAT juga — hanya kosongkan kolom Cert (D-J) agar data LAT aman
      sheet.getRange(rowIndex, 4, 1, 7).clearContent(); // D=4 sampai J=10 (7 kolom)
    } else {
      // Baris hanya Cert — hapus seluruh baris
      sheet.deleteRow(rowIndex);
    }

    return { success: true };
  } catch (e) {
    Logger.log('Error in deleteData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- DELETE DATA LAT --- */
function deleteLATData(id) {
  try {
    var sheet = connect();
    if (!sheet) return { success: false, error: 'Sheet Perencanaan tidak ditemukan' };

    // ID LAT format: "X_LAT"
    var originalId = String(id).replace('_LAT', '');

    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === originalId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Data LAT tidak ditemukan dengan ID: ' + originalId };

    // Cek apakah baris ini juga punya data Cert (kolom D / index 3)
    var hasCert = rows[rowIndex - 1][3] && String(rows[rowIndex - 1][3]).trim() !== '';

    if (hasCert) {
      // Baris punya Cert juga — hanya kosongkan kolom LAT (L-P) agar data Cert aman
      sheet.getRange(rowIndex, 12, 1, 5).clearContent(); // L=12 sampai P=16 (5 kolom)
    } else {
      // Baris hanya LAT — hapus seluruh baris
      sheet.deleteRow(rowIndex);
    }

    return { success: true };
  } catch (e) {
    Logger.log('Error in deleteLATData: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* --- DATA PELAKSANAAN (UPDATED: SESUAI USER HEADERS) --- */
/**
 * Membaca data dari sheet Pelaksanaan.
 * Kolom: NO, SAP, Start, End, Bulan, Tahun, Item ID, Sap Instruktur, Nama Instruktur, 
 * Course Title, SAP, Nama Partisipan, Room, Pesona, Kel, Departemen, Unit Kerja, 
 * Jumlah Hadir, Count Pelatihan, Durasi, Kehadiran, Durasi Peserta, Durasi Instruktur
 */
/**
 * Membaca data dari sheet peserta pada Spreadsheet baru.
 * Kolom: umn, Start, End, Bulan, Tahun, Item ID, Sap Instruktur, Name Instruktur, Corse Title, SAP, Name Participant, Room, Presensi, Ket, Departemen
 */
function getRealizationData() {
  try {
    var ss = SpreadsheetApp.openById(PESERTA_SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Peserta'); 
    
    if (!sheet) {
      Logger.log('Sheet peserta tidak ditemukan di Spreadsheet baru');
      return [];
    }

    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        
        try {
            // Skip jika baris kosong (cek SAP atau Course Title)
            if (!r[9] && !r[8]) continue;

            var dateStart = safeParseDate(r[1]);
            var dateEnd = safeParseDate(r[2]);
            
            var tahun = r[4] ? String(r[4]).trim() : "";
            
            // Fallback tahun jika kosong
            if (!tahun && dateStart) {
                var d = new Date(dateStart);
                if (!isNaN(d.getTime())) tahun = String(d.getFullYear());
            }

            data.push({
                id: r[0] ? String(r[0]) : "P_" + i,
                rowIndex: i + 1,
                sapStart: dateStart,         
                end: dateEnd,              
                bulan: r[3] ? String(r[3]) : "",               
                tahun: tahun,                                   
                itemId: r[5] ? String(r[5]) : "",              
                sapInstruktur: r[6] ? String(r[6]) : "",       
                namaInstruktur: r[7] ? String(r[7]) : "",      
                courseTitle: r[8] ? String(r[8]) : "",
                judulPelatihan: r[8] ? String(r[8]) : "", 
                sapPeserta: r[9] ? String(r[9]) : "",        
                namaPeserta: r[10] ? String(r[10]) : "",       
                room: r[11] ? String(r[11]) : "",              
                presensi: r[12] ? String(r[12]) : "",
                ket: r[13] ? String(r[13]) : "",
                departemen: r[14] ? String(r[14]) : "",
                
                // NEW COLUMNS FROM SCREENSHOT
                unitKerja: r[15] ? String(r[15]) : "",
                jumlahHadir: r[16] ? String(r[16]) : "",
                countPelatihan: r[17] ? String(r[17]) : "",
                durasi: r[18] ? String(r[18]) : "",
                kehadiran: r[19] ? String(r[19]) : "",
                durasiPelatihan: r[20] ? String(r[20]) : "",
                durasiIndividu: r[21] ? String(r[21]) : "",
                
                // Compatibility Fields
                sap: r[9] ? String(r[9]) : "NO_SAP",
                nama: r[10] ? String(r[10]) : "No Name"
            });
        } catch (errRow) {
            Logger.log("Error processing row " + i + ": " + errRow.message);
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getRealizationData: ' + e.message);
    return []; 
  }
}

/* ─────────────────────────────────────────────────────────────────────────────
 * REALIZATION DATA CRUD OPERATIONS
 * ───────────────────────────────────────────────────────────────────────────── */

/**
 * Add new realization data to Pelaksanaan sheet
 */
// Removed realization CRUD functions (addRealizationData, updateRealizationData, deleteRealizationData)

/* ─────────────────────────────────────────────────────────────────────────────
 * EVALUASI L1 DATA OPERATIONS
 * ───────────────────────────────────────────────────────────────────────────── */

/**
 * Get all L1 evaluation data from sheet "L1"
 */
function getEvaluasiL1Data() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      Logger.log('Sheet L1 not found');
      return [];
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) {
      Logger.log('No data in L1 sheet');
      return [];
    }

    var data = [];
    
    // Start from row 2 (skip header)
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      
      // Skip empty rows
      if (!r[0] && !r[1]) continue;
      
        data.push({
        id: r[0] ? String(r[0]) : '',                           
        judulPelatihan: r[1] ? String(r[1]) : '',               
        pelaksanaanId: r[2] ? safeParseDate(r[2]) : '',               
        sap: r[3] ? String(r[3]) : '',                          
        namaPeserta: r[4] ? String(r[4]) : '',                  
        tempatPembelajaran: r[5] ? String(r[5]) : '',           
        fasilitasMedia: r[6] ? String(r[6]) : '',               
        pelayananUmum: r[7] ? String(r[7]) : '',                
        ratapenyelenggaraan: r[8] ? String(r[8]) : '',         
        materi: r[9] ? String(r[9]) : '',                       
        tujuanTercapai: r[10] ? String(r[10]) : '',              
        penyajian: r[11] ? String(r[11]) : '',                   
        disiplin: r[12] ? String(r[12]) : '',                    
        rataPembelajaran: r[13] ? String(r[13]) : '',            
        pengetahuan: r[14] ? String(r[14]) : '',                 
        presentasi: r[15] ? String(r[15]) : '',                  
        perilaku: r[16] ? String(r[16]) : '',                    
        waktu: r[17] ? String(r[17]) : '',                       
        rataInstruktur: r[18] ? String(r[18]) : '',              
        rataKeseluruhan: r[19] ? String(r[19]) : '',             
        komentarPeserta: r[20] ? String(r[20]) : ''              
      });
    }
    
    Logger.log('L1 data loaded: ' + data.length + ' records');
    return data;
    
  } catch (e) {
    Logger.log('ERROR getEvaluasiL1Data: ' + e.message);
    return [];
  }
}

/**
 * Add new L1 evaluation
 */
function addEvaluasiL1(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      return { success: false, error: 'Sheet L1 not found' };
    }

    var nextId = "=ROW()-1";
    
    var newRow = [
      nextId,                                     // A
      formData.judulPelatihan || '',              // B
      formData.pelaksanaanId || '',               // C
      formData.sap || '',                         // D
      formData.namaPeserta || '',                 // E
      formData.tempatPembelajaran || '',          // F
      formData.fasilitasMedia || '',              // G
      formData.pelayananUmum || '',               // H
      formData.ratapenyelenggaraan || '',         // I (Manual Input)
      formData.materi || '',                      // J
      formData.tujuanTercapai || '',              // K
      formData.penyajian || '',                   // L
      formData.disiplin || '',                    // M
      formData.rataPembelajaran || '',            // N (Manual Input)
      formData.pengetahuan || '',                 // O
      formData.presentasi || '',                  // P
      formData.perilaku || '',                    // Q
      formData.waktu || '',                       // R
      formData.rataInstruktur || '',              // S (Manual Input)
      formData.rataKeseluruhan || '',             // T (Manual Input)
      formData.komentarPeserta || ''              // U
    ];
    
    sheet.appendRow(newRow);
    copyRowFormat(sheet, sheet.getLastRow() - 1, sheet.getLastRow());
    
    var updatedData = getEvaluasiL1Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in addEvaluasiL1: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Update L1 evaluation
 */
function updateEvaluasiL1(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      return { success: false, error: 'Sheet L1 not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Data not found' };
    }
    
    // Update all fields
    sheet.getRange(rowIndex, 2).setValue(formData.judulPelatihan || '');
    sheet.getRange(rowIndex, 3).setValue(formData.pelaksanaanId || '');
    sheet.getRange(rowIndex, 4).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 5).setValue(formData.namaPeserta || '');
    sheet.getRange(rowIndex, 6).setValue(formData.tempatPembelajaran || '');
    sheet.getRange(rowIndex, 7).setValue(formData.fasilitasMedia || '');
    sheet.getRange(rowIndex, 8).setValue(formData.pelayananUmum || '');
    sheet.getRange(rowIndex, 9).setValue(formData.ratapenyelenggaraan || ''); // Manual Input
    sheet.getRange(rowIndex, 10).setValue(formData.materi || '');
    sheet.getRange(rowIndex, 11).setValue(formData.tujuanTercapai || '');
    sheet.getRange(rowIndex, 12).setValue(formData.penyajian || '');
    sheet.getRange(rowIndex, 13).setValue(formData.disiplin || '');
    sheet.getRange(rowIndex, 14).setValue(formData.rataPembelajaran || ''); // Manual Input
    sheet.getRange(rowIndex, 15).setValue(formData.pengetahuan || '');
    sheet.getRange(rowIndex, 16).setValue(formData.presentasi || '');
    sheet.getRange(rowIndex, 17).setValue(formData.perilaku || '');
    sheet.getRange(rowIndex, 18).setValue(formData.waktu || '');
    sheet.getRange(rowIndex, 19).setValue(formData.rataInstruktur || ''); // Manual Input
    sheet.getRange(rowIndex, 20).setValue(formData.rataKeseluruhan || ''); // Manual Input
    sheet.getRange(rowIndex, 21).setValue(formData.komentarPeserta || '');
    
    var updatedData = getEvaluasiL1Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in updateEvaluasiL1: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Delete L1 evaluation
 */
function deleteEvaluasiL1(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L1') || ss.getSheetByName('l1');
    
    if (!sheet) {
      return { success: false, error: 'Sheet L1 not found' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Data not found' };
    }
    
    sheet.deleteRow(rowIndex);
    
    var updatedData = getEvaluasiL1Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in deleteEvaluasiL1: ' + e.message);
    return { success: false, error: e.message };
  }
}

/** 
 * =================================================================================
 * EVALUASI L2 (LEARNING) - CRUD
 * =================================================================================
 */

function getEvaluasiL2Data() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    
    if (!sheet) {
      // Auto-create if not exists
      sheet = ss.insertSheet('L2');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan', 'SAP', 'Nama Peserta', 
        'Pre Test', 'Post Test', 'Increase', 'Ket.'
      ]);
      return [];
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) return [];

    var data = [];
    
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!r[0] && !r[1] && !r[4]) continue; // Check ID, Judul, or Nama
      
      data.push({
        id: r[0] ? String(r[0]) : '',
        judulPelatihan: r[1] ? String(r[1]) : '',
        pelaksanaanId: r[2] ? safeParseDate(r[2]) : '',
        sap: r[3] ? String(r[3]) : '',
        namaPeserta: r[4] ? String(r[4]) : '',
        preTest: r[5] ? String(r[5]) : '0',
        postTest: r[6] ? String(r[6]) : '0',
        increase: r[7] ? String(r[7]) : '0',       
        ket: r[8] ? String(r[8]) : ''            
      });
    }
    
    return data;
    
  } catch (e) {
    Logger.log('ERROR getEvaluasiL2Data: ' + e.message);
    return [];
  }
}

function addEvaluasiL2(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    
    if (!sheet) {
      sheet = ss.insertSheet('L2');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan', 'SAP', 'Nama Peserta', 
        'Pre Test', 'Post Test', 'Increase', 'Ket.'
      ]);
    }

    var nextId = "=ROW()-1";
    var increase = (parseFloat(formData.postTest) || 0) - (parseFloat(formData.preTest) || 0);
    
    var newRow = [
      nextId,
      formData.judulPelatihan || '',
      formData.pelaksanaanId || '',
      formData.sap || '',
      formData.namaPeserta || '',
      formData.preTest || 0,
      formData.postTest || 0,
      increase.toFixed(2),
      formData.ket || ''
    ];
    
    sheet.appendRow(newRow);
    copyRowFormat(sheet, sheet.getLastRow() - 1, sheet.getLastRow());
    
    var updatedData = getEvaluasiL2Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in addEvaluasiL2: ' + e.message);
    return { success: false, error: e.message };
  }
}

function updateEvaluasiL2(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    if (!sheet) return { success: false, error: 'Sheet L2 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    var increase = (parseFloat(formData.postTest) || 0) - (parseFloat(formData.preTest) || 0);

    // Update fields (Columns 2-9)
    sheet.getRange(rowIndex, 2).setValue(formData.judulPelatihan || '');
    sheet.getRange(rowIndex, 3).setValue(formData.pelaksanaanId || '');
    sheet.getRange(rowIndex, 4).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 5).setValue(formData.namaPeserta || '');
    sheet.getRange(rowIndex, 6).setValue(formData.preTest || 0);
    sheet.getRange(rowIndex, 7).setValue(formData.postTest || 0);
    sheet.getRange(rowIndex, 8).setValue(increase.toFixed(2));
    sheet.getRange(rowIndex, 9).setValue(formData.ket || '');
    
    var updatedData = getEvaluasiL2Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in updateEvaluasiL2: ' + e.message);
    return { success: false, error: e.message };
  }
}

function deleteEvaluasiL2(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L2');
    if (!sheet) return { success: false, error: 'Sheet L2 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    sheet.deleteRow(rowIndex);
    
    var updatedData = getEvaluasiL2Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in deleteEvaluasiL2: ' + e.message);
    return { success: false, error: e.message };
  }
}

/** 
 * =================================================================================
 * EVALUASI L3 (BEHAVIOR) - CRUD
 *Headers: No, Judul Pelatihan, Pelaksanaan Learning, SAP, Nama Peserta, Nilai Evaluasi, Ket., Key Behaviour, Tanggal Eval
 * =================================================================================
 */

function getEvaluasiL3Data() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    
    if (!sheet) {
      sheet = ss.insertSheet('L3');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan Learning', 'SAP', 'Nama Peserta', 
        'Nilai Evaluasi', 'Ket.', 'Key Behaviour', 'Tanggal Eval'
      ]);
      return [];
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) return [];

    var data = [];
    
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!r[0] && !r[1] && !r[4]) continue;
      
      data.push({
        id: r[0] ? String(r[0]) : '',
        judulPelatihan: r[1] ? String(r[1]) : '',
        pelaksanaanId: r[2] ? String(r[2]) : '',
        sap: r[3] ? String(r[3]) : '',
        namaPeserta: r[4] ? String(r[4]) : '',
        nilaiEvaluasi: r[5] ? String(r[5]) : '',
        ket: r[6] ? String(r[6]) : '',
        keyBehaviour: r[7] ? String(r[7]) : '',
        tanggalEval: r[8] ? Utilities.formatDate(new Date(r[8]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : ''
      });
    }
    
    return data;
    
  } catch (e) {
    Logger.log('ERROR getEvaluasiL3Data: ' + e.message);
    return [];
  }
}

function addEvaluasiL3(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    
    if (!sheet) {
      sheet = ss.insertSheet('L3');
      sheet.appendRow([
        'No', 'Judul Pelatihan', 'Pelaksanaan Learning', 'SAP', 'Nama Peserta', 
        'Nilai Evaluasi', 'Ket.', 'Key Behaviour', 'Tanggal Eval'
      ]);
    }

    var nextId = "=ROW()-1";
    
    var newRow = [
      nextId,
      formData.judulPelatihan || '',
      formData.pelaksanaanId || '',
      formData.sap || '',
      formData.namaPeserta || '',
      formData.nilaiEvaluasi || '',
      formData.ket || '',
      formData.keyBehaviour || '',
      formData.tanggalEval || ''
    ];
    
    sheet.appendRow(newRow);
    copyRowFormat(sheet, sheet.getLastRow() - 1, sheet.getLastRow());
    
    var updatedData = getEvaluasiL3Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in addEvaluasiL3: ' + e.message);
    return { success: false, error: e.message };
  }
}

function updateEvaluasiL3(formData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    if (!sheet) return { success: false, error: 'Sheet L3 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    // Update fields (Columns 2-9)
    sheet.getRange(rowIndex, 2).setValue(formData.judulPelatihan || '');
    sheet.getRange(rowIndex, 3).setValue(formData.pelaksanaanId || '');
    sheet.getRange(rowIndex, 4).setValue(formData.sap || '');
    sheet.getRange(rowIndex, 5).setValue(formData.namaPeserta || '');
    sheet.getRange(rowIndex, 6).setValue(formData.nilaiEvaluasi || '');
    sheet.getRange(rowIndex, 7).setValue(formData.ket || '');
    sheet.getRange(rowIndex, 8).setValue(formData.keyBehaviour || '');
    sheet.getRange(rowIndex, 9).setValue(formData.tanggalEval || '');
    
    var updatedData = getEvaluasiL3Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in updateEvaluasiL3: ' + e.message);
    return { success: false, error: e.message };
  }
}

function deleteEvaluasiL3(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('L3') || ss.getSheetByName('l3');
    if (!sheet) return { success: false, error: 'Sheet L3 not found' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: 'Data not found' };
    
    sheet.deleteRow(rowIndex);
    
    var updatedData = getEvaluasiL3Data();
    return { success: true, data: updatedData };
    
  } catch (e) {
    Logger.log('Error in deleteEvaluasiL3: ' + e.message);
    return { success: false, error: e.message };
  }
}

/* ─────────────────────────────────────────────────────────────────────────────
 * VENDOR DATA OPERATIONS (NEW)
 * ───────────────────────────────────────────────────────────────────────────── */

function getVendorData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Ajuan Vendor");
    if (!sheet) {
      Logger.log("Sheet 'Ajuan Vendor' not found");
      return [];
    }
    
    var rows = sheet.getDataRange().getValues();
    var richTextValues = sheet.getDataRange().getRichTextValues(); // Get links
    var data = [];
    
    // Skip header row (i=1)
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r || r.length < 2) continue;
      
      var vendorName = String(r[1] || "").trim();
      if (!vendorName) continue; // Skip if vendor name is empty
      
      // Extract link from Column L (index 11)
      var fileUrl = "-";
      var richTextCell = richTextValues[i][11];
      if (richTextCell) {
        fileUrl = richTextCell.getLinkUrl() || r[11] || "-";
      }

      // Process Biaya (Column H / index 7)
      var biayaRaw = r[7];
      var biayaFormatted = "-";
      if (biayaRaw) {
        if (typeof biayaRaw === "number") {
          biayaFormatted = "Rp " + biayaRaw.toLocaleString('id-ID');
        } else {
          // If it's a string like "1.000.000", try to clean it
          var cleanNum = Number(String(biayaRaw).replace(/[^\d]/g, ''));
          if (!isNaN(cleanNum) && cleanNum > 0) {
            biayaFormatted = "Rp " + cleanNum.toLocaleString('id-ID');
          } else {
            biayaFormatted = String(biayaRaw);
          }
        }
      }
      
      data.push({
        rowIndex: i + 1,
        timestamp: r[0] instanceof Date ? r[0].getTime() : 0,
        namaVendor: vendorName,
        namaLsp: String(r[2] || "-"),
        picVendor: String(r[3] || "-"),
        namaSertifikasi: String(r[4] || "-"),
        jenisSertifikasi: String(r[5] || "-"),
        silabus: String(r[6] || "-"),
        biaya: biayaFormatted,
        metode: String(r[8] || "-"),
        tempat: String(r[9] || "-"),
        tanggal: r[10] ? safeParseDate(r[10]) : "-",
        file: fileUrl,
        status: String(r[12] || "Pending").trim() || "Pending"
      });
    }
    Logger.log("getVendorData found " + data.length + " valid rows");
    return data;
  } catch (e) {
    Logger.log('ERROR getVendorData: ' + e.message);
    return [];
  }
}

function updateVendorStatus(rowIndex, status) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Ajuan Vendor");
    
    // Update Status (Column M / 13)
    sheet.getRange(rowIndex, 13).setValue(status);
    
    // Update Audit Trail (Column N/O)
    var adminEmail = Session.getActiveUser().getEmail() || "Admin Dashboard";
    var timestamp = new Date();
    sheet.getRange(rowIndex, 14).setValue(adminEmail);
    sheet.getRange(rowIndex, 15).setValue(timestamp);
    
    var rowData = sheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
    var vendorName = rowData[1];
    var sertifikasi = rowData[4];
    var picName = rowData[3];
    
    // --- Notification Logic ---
    // Note: You can enable real email sending here if needed
    // MailApp.sendEmail(recipient, subject, body);

    return { success: true, data: getVendorData() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function debugHeaders() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var p = ss.getSheetByName("Pelaksanaan");
    var l1 = ss.getSheetByName("L1");
    var res = "";
    if (p) res += "PELAKSANAAN HEADERS: " + JSON.stringify(p.getRange(1, 1, 1, 20).getValues()[0]) + "\n";
    if (l1) res += "L1 HEADERS: " + JSON.stringify(l1.getRange(1, 1, 1, 30).getValues()[0]) + "\n";
    return res;
  } catch (e) { return e.message; }
}

/**
 * Helper to copy format from one row to another
 */
function copyRowFormat(sheet, sourceRow, targetRow) {
  try {
    if (sourceRow < 1) return;
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    var sourceRange = sheet.getRange(sourceRow, 1, 1, lastCol);
    var targetRange = sheet.getRange(targetRow, 1, 1, lastCol);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  } catch (e) {
    Logger.log("Error copying format: " + e.message);
  }
}
