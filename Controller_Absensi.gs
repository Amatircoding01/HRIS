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

/**
 * =========================================================
 * FUNGSI: KELOLA HARI LIBUR
 * =========================================================
 */

// 1. Mengambil daftar hari libur dari sheet Master_Libur
function getLiburData() {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    const sheet = ss.getSheetByName('Master_Libur');
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Sheet Master_Libur tidak ditemukan' });
    
    const data = sheet.getDataRange().getValues();
    data.shift(); // Buang baris header
    
    let result = [];
    for (let i = 0; i < data.length; i++) {
      let tgl = data[i][0];
      if (!tgl || tgl === "") continue; // Lewati baris kosong
      
      // Format tanggal agar rapi saat ditampilkan di tabel
      let tglStr = (tgl instanceof Date) ? Utilities.formatDate(tgl, "Asia/Jakarta", "dd-MM-yyyy") : String(tgl);
      
      result.push({
        Tanggal: tglStr,
        Status: data[i][1] || '-',
        Keterangan: data[i][2] || '-'
      });
    }
    
    return JSON.stringify({ status: 'success', data: result });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

// 2. Menyimpan hari libur baru ke sheet Master_Libur
function saveLiburData(tglInput, statusInput, ketInput) {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    const sheet = ss.getSheetByName('Master_Libur');
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Sheet Master_Libur tidak ditemukan' });
    
    // Konversi input string (YYYY-MM-DD dari form HTML) ke format Date untuk Spreadsheet
    let parts = tglInput.split('-');
    let dateObj = new Date(parts[0], parts[1] - 1, parts[2]); // Tahun, Bulan (0-index), Hari
    
    // Tambahkan baris baru ke paling bawah
    sheet.appendRow([dateObj, statusInput, ketInput]);
    
    return JSON.stringify({ status: 'success' });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * =========================================================
 * FUNGSI: EDIT / TAMBAH ABSENSI (VERSI FLEKSIBEL)
 * =========================================================
 */

function getEditAbsensiData(pt, tglInput) {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    
    // 1. Pecah input tanggal dari UI (YYYY-MM-DD)
    let parts = tglInput.split('-');
    let tgtY = parseInt(parts[0]);
    let tgtM = parseInt(parts[1]) - 1; // Index bulan di JS mulai dari 0
    let tgtD = parseInt(parts[2]);

    // 2. AMBIL MASTER KARYAWAN
    const sheetKar = ss.getSheetByName('Master_Karyawan');
    if (!sheetKar) throw new Error("Sheet Master_Karyawan tidak ditemukan.");
    const dataKar = sheetKar.getDataRange().getValues();
    dataKar.shift(); // Buang header
    
    let listKaryawan = [];
    let ptLower = String(pt).toLowerCase(); // misal "kobe"
    
    for (let i = 0; i < dataKar.length; i++) {
      let nrppKar = String(dataKar[i][0]).trim();
      let namaKar = dataKar[i][1];
      let ptKar = String(dataKar[i][7] || "").toLowerCase(); // Kolom H
      
      if (!nrppKar) continue;
      
      // STRATEGI TOLERANSI PT: Gunakan 'includes' bukan '==='
      // Jika UI mengirim "kobe", dan di sheet tertulis "pt kobexindo", "kobe" ada di dalamnya!
      if (ptKar.includes(ptLower) || ptLower === 'all' || ptLower === '') {
        listKaryawan.push({ NRPP: nrppKar, Nama: namaKar });
      }
    }

    // Jika karyawan tidak ditemukan sama sekali, kembalikan error spesifik
    if (listKaryawan.length === 0) {
        return JSON.stringify({ status: 'error', message: `Tidak ada karyawan ditemukan untuk kata kunci PT: "${pt}" di Master_Karyawan.` });
    }

    // 3. AMBIL RAW KEHADIRAN
    const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
    if (!sheetRaw) throw new Error("Sheet Raw_Kehadiran tidak ditemukan.");
    const dataRaw = sheetRaw.getDataRange().getValues();
    
    let mapAbsensi = {};
    
    for (let i = 1; i < dataRaw.length; i++) {
      let tglRow = dataRaw[i][0];
      if (!tglRow || tglRow === "") continue;
      
      // LOGIKA TANGGAL KEBAL
      let isMatch = false;
      let rowDate = new Date(tglRow);
      
      if (!isNaN(rowDate.getTime())) {
        if (rowDate.getFullYear() === tgtY && rowDate.getMonth() === tgtM && rowDate.getDate() === tgtD) {
          isMatch = true;
        }
      } else {
        // Fallback string manual
        let tglStr = String(tglRow);
        let strTgt1 = tglInput; // "YYYY-MM-DD"
        let strTgt2 = `${("0"+tgtD).slice(-2)}/${("0"+(tgtM+1)).slice(-2)}/${tgtY}`; // "DD/MM/YYYY"
        let strTgt3 = `${("0"+tgtD).slice(-2)}-${("0"+(tgtM+1)).slice(-2)}-${tgtY}`; // "DD-MM-YYYY"
        if (tglStr.includes(strTgt1) || tglStr.includes(strTgt2) || tglStr.includes(strTgt3)) {
          isMatch = true;
        }
      }

      if (isMatch) {
        let nrppRow = String(dataRaw[i][1]).trim();
        mapAbsensi[nrppRow] = {
          rowIndex: i + 1, // +1 karena array mulai dari 0 tapi sheet mulai dari 1
          JamMasuk: dataRaw[i][3] instanceof Date ? Utilities.formatDate(dataRaw[i][3], "Asia/Jakarta", "HH:mm") : (dataRaw[i][3] || ""),
          JamKeluar: dataRaw[i][4] instanceof Date ? Utilities.formatDate(dataRaw[i][4], "Asia/Jakarta", "HH:mm") : (dataRaw[i][4] || ""),
          Status: dataRaw[i][5] || ""
        };
      }
    }

    // 4. GABUNGKAN DATA
    let result = [];
    for (let k of listKaryawan) {
      let absen = mapAbsensi[k.NRPP];
      result.push({
        NRPP: k.NRPP,
        Nama: k.Nama,
        JamMasuk: absen ? absen.JamMasuk : "",
        JamKeluar: absen ? absen.JamKeluar : "",
        Status: absen ? absen.Status : "",
        RowIndex: absen ? absen.rowIndex : null
      });
    }

    return JSON.stringify({ status: 'success', data: result });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

function saveSingleAbsensi(nrpp, nama, tglInput, jamMasuk, jamKeluar, status, rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
    
    // Normalisasi tanggal agar konsisten di spreadsheet
    let dateObj = tglInput; // Default biarkan format "YYYY-MM-DD"
    let parts = tglInput.split('-');
    if (parts.length === 3) {
      dateObj = new Date(parts[0], parts[1] - 1, parts[2]); // Convert ke objek tanggal
    }

    if (rowIndex) {
      // MODE EDIT
      sheetRaw.getRange(rowIndex, 4).setValue(jamMasuk);
      sheetRaw.getRange(rowIndex, 5).setValue(jamKeluar);
      sheetRaw.getRange(rowIndex, 6).setValue(status);
    } else {
      // MODE TAMBAH BARU
      sheetRaw.appendRow([dateObj, nrpp, nama, jamMasuk, jamKeluar, status]);
    }

    return JSON.stringify({ status: 'success' });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * Mengambil daftar perusahaan unik dari Master_Karyawan Kolom H
 */
function getUniquePerusahaan() {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    const sheet = ss.getSheetByName('Master_Karyawan');
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Sheet Master_Karyawan tidak ditemukan' });
    
    const data = sheet.getDataRange().getValues();
    data.shift(); // Buang header
    
    let listPT = [];
    for (let i = 0; i < data.length; i++) {
      let pt = data[i][7]; // Kolom H (index 7)
      if (pt && pt !== "" && !listPT.includes(pt)) {
        listPT.push(pt);
      }
    }
    
    // Urutkan berdasarkan abjad
    return JSON.stringify({ status: 'success', data: listPT.sort() });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

