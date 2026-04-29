/**
 * =========================================================
 * KONTROLER MODUL ABSENSI
 * Terhubung dengan Database: 15HSyULRuhrsq86TMogPfahY-KZsCuxM60k4TvQJu6Z4
 * =========================================================
 */

const DB_ABSENSI_ID = '15HSyULRuhrsq86TMogPfahY-KZsCuxM60k4TvQJu6Z4';

function getDashboardAbsensiData() {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
    if (!sheetRaw) return JSON.stringify({ status: 'error', message: 'Sheet Raw_Kehadiran tidak ditemukan!' });
    
    const data = sheetRaw.getDataRange().getValues();
    data.shift(); // Buang header
    
    // Ambil info hari ini
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = today.getMonth(); // Bulan dimulai dari 0
    const dd = today.getDate();
    
    let statHadir = 0, statCuti = 0, statIzinSakit = 0, statTerlambat = 0;
    let logHariIni = [];
    let logSemua = []; // Array penampung seluruh data (untuk Fallback)
    
    for (let i = 0; i < data.length; i++) {
      let row = data[i];
      let tglRow = row[0]; // Kolom A: Tanggal
      
      if (!tglRow || tglRow === "") continue; 
      
      // LOGIKA TANGGAL YANG LEBIH FLEKSIBEL
      let isToday = false;
      let rowDate = new Date(tglRow); // Paksa jadikan objek Date
      
      if (!isNaN(rowDate.getTime())) {
        // Bandingkan secara absolut (Tahun, Bulan, Tanggal) -> anti meleset
        if (rowDate.getFullYear() === yyyy && rowDate.getMonth() === mm && rowDate.getDate() === dd) {
          isToday = true;
        }
      } else {
        // Jika formatnya teks murni dan gagal di-parse, cek string secara kasar
        let tglStr = String(tglRow);
        let strToday1 = Utilities.formatDate(today, "Asia/Jakarta", "yyyy-MM-dd");
        let strToday2 = Utilities.formatDate(today, "Asia/Jakarta", "dd/MM/yyyy");
        let strToday3 = Utilities.formatDate(today, "Asia/Jakarta", "dd-MM-yyyy");
        if (tglStr.includes(strToday1) || tglStr.includes(strToday2) || tglStr.includes(strToday3)) {
          isToday = true;
        }
      }
      
      // Ambil sisa kolom
      let nrpp = row[1] || "-";
      let nama = row[2] || "-";
      let jamMasuk = row[3] || "";
      let jamKeluar = row[4] || "";
      let status = row[5] || "";
      let statusLower = String(status).toLowerCase();
      
      // Normalisasi Jam dan Keterlambatan
      let jmFormat = "-", jkFormat = "-", keterangan = "-";
      if (jamMasuk instanceof Date) {
        jmFormat = Utilities.formatDate(jamMasuk, "Asia/Jakarta", "HH:mm");
        if (jamMasuk.getHours() > 8 || (jamMasuk.getHours() === 8 && jamMasuk.getMinutes() > 0)) {
          keterangan = "Terlambat";
        }
      } else if (jamMasuk !== "") jmFormat = String(jamMasuk);
      
      if (jamKeluar instanceof Date) {
        jkFormat = Utilities.formatDate(jamKeluar, "Asia/Jakarta", "HH:mm");
      } else if (jamKeluar !== "") jkFormat = String(jamKeluar);
      
      let itemLog = {
        NRPP: nrpp, Nama: nama, JamMasuk: jmFormat, JamKeluar: jkFormat, Status: status, Keterangan: keterangan, TglAsli: String(tglRow)
      };
      
      // Kumpulkan ke keranjang semua data
      logSemua.push(itemLog);
      
      // Jika memang datanya HARI INI, masukkan ke statistik
      if (isToday) {
        logHariIni.push(itemLog);
        if (statusLower.includes('hadir')) statHadir++;
        else if (statusLower.includes('cuti')) statCuti++;
        else if (statusLower.includes('sakit') || statusLower.includes('izin')) statIzinSakit++;
        if (keterangan === "Terlambat") statTerlambat++;
      }
    }
    
    // --- MODE FALLBACK CERDAS ---
    // Jika hari ini 0 data, kita JANGAN tampilkan tabel kosong.
    // Kita paksa kembalikan 5 data terakhir dari sheet agar kamu bisa melihat isinya.
    let isFallback = false;
    if (logHariIni.length === 0 && logSemua.length > 0) {
      isFallback = true;
      logHariIni = logSemua.slice(-5).reverse(); // Ambil 5 data paling bawah, balik urutannya
    }
    
    return JSON.stringify({
      status: 'success',
      stats: { hadir: statHadir, cuti: statCuti, izin: statIzinSakit, terlambat: statTerlambat },
      logs: logHariIni,
      fallbackMode: isFallback // Penanda untuk Frontend
    });
    
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}
