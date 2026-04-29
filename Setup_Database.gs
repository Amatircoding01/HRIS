/**
 * AUTO-SETUP DATABASE KLAIM - HRIS KOBE 2026
 * Script ini hanya perlu dijalankan 1x untuk membuat struktur tabel secara otomatis
 */

function setupDatabaseKlaimOtomatis() {
  try {
    // Memanggil ID Spreadsheet dari Config.gs
    const klaimSSId = CONFIG.DATABASE.KLAIM_SS_ID; 
    
    if (!klaimSSId || klaimSSId === "ID_SPREADSHEET_KLAIM_ANDA_DISINI") {
      throw new Error("Mohon masukkan ID Spreadsheet Klaim yang valid di dalam file Config.gs terlebih dahulu!");
    }

    const ss = SpreadsheetApp.openById(klaimSSId);
    
    // Daftar Sheet yang akan dibuat otomatis
    const sheetsToCreate = [
      "Klaim_Pemeriksaan_Kehamilan",
      "Klaim_Kacamata",
      "Klaim_Gigi",
      "Klaim_KB",
      "Maternity"
    ];
    
    // Struktur Header Standar Enterprise untuk Pengajuan Klaim
    const headers = [
      "ID_Klaim", 
      "NRPP", 
      "Nama_Karyawan", 
      "Departement",
      "Tanggal_Pengajuan", 
      "Jenis_Klaim",
      "Keterangan_Medis", 
      "Nominal_Kwitansi", 
      "Nominal_Disetujui", 
      "Status_Approval", 
      "Keterangan_HR",
      "Timestamp_Update"
    ];
    
    // Looping untuk membuat masing-masing sheet
    sheetsToCreate.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      
      // Jika sheet belum ada, maka sistem akan membuatkannya
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      } else {
        // Jika sudah ada, bersihkan dulu agar tidak tumpang tindih
        sheet.clear();
      }
      
      // Inject Header ke Baris Pertama (A1:L1)
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Desain Tampilan Header (Modern Corporate Style)
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight("bold")
           .setBackground("#1e293b") // Warna Deep Blue khas UI KOBE
           .setFontColor("white")
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle");
           
      // Kunci (Freeze) baris pertama agar saat di-scroll header tetap di atas
      sheet.setFrozenRows(1);
      
      // Otomatis menyesuaikan lebar kolom dengan teks
      sheet.autoResizeColumns(1, headers.length);
    });
    
    Logger.log("SUKSES: 5 Sheet Klaim beserta Header berhasil dibuat otomatis!");
    
  } catch (err) {
    Logger.log("ERROR: " + err.message);
  }
}
