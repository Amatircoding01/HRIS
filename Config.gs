/**
 * CONFIGURATION FILE - HRIS KOBE 2026
 */
const CONFIG = {
  PROJECT_NAME: "HRIS KOBE 2026",
  VERSION: "1.4.0",
  DEVELOPER: "Enterprise Architect",
  
  // KREDENSIAL MASTER LOGIN
  AUTH: {
    USERNAME: "admin",      
    PASSWORD: "12345" 
  },

  // SENTRALISASI DATABASE (MULTIPLE SPREADSHEETS)
  DATABASE: {
    // 1. Database Master Karyawan
    KARYAWAN_SS_ID: "1vCRZ1ex4LSlAMui_VIqlcRxiRwirT0VfAMtJMyLr_3s", // Tetap biarkan ID Karyawan Anda di sini
    SHEET_NAME: "Karyawan",

    // 2. Database Klaim Karyawan (BARU)
    KLAIM_SS_ID: "1kOVpm0b1L_afb5qtcFK3Ynrd6tdU-7EisY5g4UNYtUk", // Masukkan ID file Spreadsheet khusus Klaim nantinya
    SHEET_KLAIM: {
        PERSALINAN: "Klaim_Persalinan",
        PEMERIKSAAN_KEHAMILAN: "Klaim_Pemeriksaan_Kehamilan",
        KACAMATA: "Klaim_Kacamata",
        GIGI: "Klaim_Gigi",
        KB: "Klaim_KB",
        MATERNITY: "Maternity"
    }
  },

  // PENGATURAN TEMA UI
  UI: {
    THEME_COLOR: "#1e293b",
    ACCENT_COLOR: "#10b981"
  }
};

/**
 * Koneksi ke Database Karyawan
 */
function getKaryawanSheet() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATABASE.KARYAWAN_SS_ID);
    const sheet = ss.getSheetByName(CONFIG.DATABASE.SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${CONFIG.DATABASE.SHEET_NAME}" tidak ditemukan.`);
    return sheet;
  } catch (err) {
    throw err;
  }
}

/**
 * Koneksi ke Database Klaim (BARU)
 * Memungkinkan pemanggilan sheet spesifik berdasarkan jenis klaim
 */
function getKlaimSheet(namaSheetKlaim) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATABASE.KLAIM_SS_ID);
    const sheet = ss.getSheetByName(namaSheetKlaim);
    if (!sheet) throw new Error(`Sheet Klaim "${namaSheetKlaim}" tidak ditemukan di database Klaim.`);
    return sheet;
  } catch (err) {
    throw err;
  }
}
