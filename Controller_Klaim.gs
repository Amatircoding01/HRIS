/**
 * CONTROLLER KLAIM - HRIS KOBE 2026 FINAL STABLE
 */

function getKaryawanForKlaim(nrpp) {
  try {
    const s = getKaryawanSheet(); const d = s.getDataRange().getValues(); const h = d[0];
    const idxNRPP = h.indexOf("NRPP"); const idxNama = h.indexOf("Nama"); const idxDept = h.indexOf("Departement"); const idxGol = h.indexOf("Golongan");
    const rIdx = d.findIndex(r => r[idxNRPP] == nrpp);
    if (rIdx === -1) throw new Error("Karyawan tidak ditemukan.");
    let total = 0;
    try {
      const sk = getKlaimSheet(CONFIG.DATABASE.SHEET_KLAIM.PEMERIKSAAN_KEHAMILAN); const dk = sk.getDataRange().getValues(); const hk = dk[0];
      const kNrppIdx = hk.indexOf("NRPP"); const kSetujuIdx = hk.indexOf("Nominal_Disetujui"); const kStatusIdx = hk.indexOf("Status_Approval");
      for (let i = 1; i < dk.length; i++) { if (dk[i][kNrppIdx] == nrpp && !dk[i][kStatusIdx].toString().toLowerCase().includes("ditolak")) total += parseFloat(dk[i][kSetujuIdx]) || 0; }
    } catch (e) {}
    return Utils.success({ NRPP: d[rIdx][idxNRPP], Nama: d[rIdx][idxNama], Departement: d[rIdx][idxDept], Golongan: d[rIdx][idxGol], TotalTerpakai: total });
  } catch (err) { return Utils.error(err.message); }
}

function submitKlaim(f) { try { getKlaimSheet(CONFIG.DATABASE.SHEET_KLAIM.PEMERIKSAAN_KEHAMILAN).appendRow([ "KLM-KEH-"+Date.now(), f.NRPP, f.Nama, f.Departement, f.Tanggal_Pengajuan, "Pemeriksaan Kehamilan", f.Keterangan_Medis, f.Nominal_Kwitansi, f.Nominal_Disetujui, "Menunggu Approval HR", "", new Date().toLocaleString('id-ID') ]); return Utils.success(null, "Sukses"); } catch (e) { return Utils.error(e.message); } }

function getKaryawanForKlaimKacamata(nrpp) {
  try {
    const s = getKaryawanSheet(); const d = s.getDataRange().getValues(); const h = d[0];
    const idxNRPP = h.indexOf("NRPP"); const idxNama = h.indexOf("Nama"); const idxDept = h.indexOf("Departement"); const idxGol = h.indexOf("Golongan");
    const rIdx = d.findIndex(r => r[idxNRPP] == nrpp);
    if (rIdx === -1) throw new Error("Karyawan tidak ditemukan.");
    let lL = null; let lF = null;
    try {
      const sk = getKlaimSheet(CONFIG.DATABASE.SHEET_KLAIM.KACAMATA); const dk = sk.getDataRange().getValues(); const hk = dk[0];
      const kNrppIdx = hk.indexOf("NRPP"); const kJenisIdx = hk.indexOf("Jenis_Klaim"); const kTglIdx = hk.indexOf("Tanggal_Pengajuan"); const kStatusIdx = hk.indexOf("Status_Approval");
      for (let i = 1; i < dk.length; i++) {
        if (dk[i][kNrppIdx] == nrpp && !dk[i][kStatusIdx].toString().toLowerCase().includes("ditolak")) {
          let j = dk[i][kJenisIdx].toString().toLowerCase(); let t = new Date(dk[i][kTglIdx]);
          if (j.includes("lensa") && (!lL || t > lL)) lL = t; if (j.includes("frame") && (!lF || t > lF)) lF = t;
        }
      }
    } catch (e) {}
    const now = new Date(); let lT = true, fT = true, iL = "Tersedia", iF = "Tersedia";
    if (lL && (now-lL)/(864e5) < 365) { lT = false; iL = "Terkunci"; } if (lF && (now-lF)/(864e5) < 730) { fT = false; iF = "Terkunci"; }
    return Utils.success({ NRPP: d[rIdx][idxNRPP], Nama: d[rIdx][idxNama], Departement: d[rIdx][idxDept], Golongan: d[rIdx][idxGol], LensaTersedia: lT, FrameTersedia: fT, InfoLensa: iL, InfoFrame: iF });
  } catch (err) { return Utils.error(err.message); }
}

function submitKlaimKacamata(items) { try { const s = getKlaimSheet(CONFIG.DATABASE.SHEET_KLAIM.KACAMATA); items.forEach(i => s.appendRow([ "KLM-KCM-"+Date.now(), i.NRPP, i.Nama, i.Departement, i.Tanggal_Pengajuan, i.Jenis_Klaim, i.Keterangan_Medis, i.Nominal_Kwitansi, i.Nominal_Disetujui, "Menunggu Approval HR", "", new Date().toLocaleString() ])); return Utils.success(null, "Sukses"); } catch (e) { return Utils.error(e.message); } }

function getKaryawanForKlaimGigi(nrpp) {
  try {
    const s = getKaryawanSheet(); const d = s.getDataRange().getValues(); const h = d[0];
    const idxNRPP = h.indexOf("NRPP"); const idxNama = h.indexOf("Nama"); const idxDept = h.indexOf("Departement"); const idxGol = h.indexOf("Golongan");
    const rIdx = d.findIndex(r => r[idxNRPP] == nrpp);
    if (rIdx === -1) throw new Error("Karyawan tidak ditemukan.");
    let total = 0; 
    try { 
      const sk = getKlaimSheet(CONFIG.DATABASE.SHEET_KLAIM.GIGI); const dk = sk.getDataRange().getValues(); const hk = dk[0];
      const kNrppIdx = hk.indexOf("NRPP"); const kSetujuIdx = hk.indexOf("Nominal_Disetujui"); const kStatusIdx = hk.indexOf("Status_Approval");
      for (let i = 1; i < dk.length; i++) { if (dk[i][kNrppIdx] == nrpp && !dk[i][kStatusIdx].toString().toLowerCase().includes("ditolak")) total += parseFloat(dk[i][kSetujuIdx]) || 0; } 
    } catch (e) {}
    return Utils.success({ NRPP: d[rIdx][idxNRPP], Nama: d[rIdx][idxNama], Departement: d[rIdx][idxDept], Golongan: d[rIdx][idxGol], TotalTerpakai: total });
  } catch (err) { return Utils.error(err.message); }
}

function submitKlaimGigi(f) { try { getKlaimSheet(CONFIG.DATABASE.SHEET_KLAIM.GIGI).appendRow([ "KLM-GGI-"+Date.now(), f.NRPP, f.Nama, f.Departement, f.Tanggal_Pengajuan, "Perawatan Gigi", f.Keterangan_Medis, f.Nominal_Kwitansi, f.Nominal_Disetujui, "Menunggu Approval HR", "", new Date().toLocaleString() ]); return Utils.success(null, "Sukses"); } catch (e) { return Utils.error(e.message); } }

function getKaryawanForKlaimPersalinan(nrpp) { 
  try { 
    const s = getKaryawanSheet(); const d = s.getDataRange().getValues(); const h = d[0];
    const idxNRPP = h.indexOf("NRPP"); const idxNama = h.indexOf("Nama"); const idxDept = h.indexOf("Departement"); const idxGol = h.indexOf("Golongan");
    const rIdx = d.findIndex(r => r[idxNRPP] == nrpp);
    if (rIdx === -1) throw new Error("Karyawan tidak ditemukan.");
    return Utils.success({ NRPP: d[rIdx][idxNRPP], Nama: d[rIdx][idxNama], Departement: d[rIdx][idxDept], Golongan: d[rIdx][idxGol] }); 
  } catch (e) { return Utils.error(e.message); } 
}

function submitKlaimPersalinan(f) { try { getKlaimSheet(CONFIG.DATABASE.SHEET_KLAIM.PERSALINAN).appendRow([ "KLM-PRS-"+Date.now(), f.NRPP, f.Nama, f.Departement, f.Tanggal_Pengajuan, "Persalinan - "+f.Jenis_Persalinan, f.Keterangan_Medis, f.Nominal_Kwitansi, f.Nominal_Disetujui, "Menunggu Approval HR", "", new Date().toLocaleString() ]); return Utils.success(null, "Sukses"); } catch (e) { return Utils.error(e.message); } }

function submitKlaimKB(f) { try { getKlaimSheet(CONFIG.DATABASE.SHEET_KLAIM.KB).appendRow([ "KLM-KBB-"+Date.now(), f.NRPP, f.Nama, f.Departement, f.Tanggal_Pengajuan, "KB - "+f.Jenis_KB, f.Keterangan_Medis, f.Nominal_Kwitansi, f.Nominal_Disetujui, "Menunggu Approval HR", "", new Date().toLocaleString() ]); return Utils.success(null, "Sukses"); } catch (e) { return Utils.error(e.message); } }

function getSemuaKlaim() {
  try {
    const list = [ CONFIG.DATABASE.SHEET_KLAIM.PERSALINAN, CONFIG.DATABASE.SHEET_KLAIM.PEMERIKSAAN_KEHAMILAN, CONFIG.DATABASE.SHEET_KLAIM.KACAMATA, CONFIG.DATABASE.SHEET_KLAIM.GIGI, CONFIG.DATABASE.SHEET_KLAIM.KB ];
    let all = []; list.forEach(s => { try { const v = getKlaimSheet(s).getDataRange().getValues(); const h = v[0]; v.slice(1).forEach(r => { let o = { sheetSource: s }; h.forEach((k, i) => o[k.replace(/\s+/g, '_')] = r[i]); all.push(o); }); } catch (e) {} });
    all.reverse(); return Utils.success(all);
  } catch (e) { return Utils.error(e.message); }
}

function updateStatusKlaim(id, s, st, k) { try { const sh = getKlaimSheet(s); const d = sh.getDataRange().getValues(); const rIdx = d.findIndex(row => row[0] == id); sh.getRange(rIdx + 1, 10).setValue(st); sh.getRange(rIdx + 1, 11).setValue(k); return Utils.success(null, "Sukses"); } catch (e) { return Utils.error(e.message); } }

function getDataForLaporan(s) { try { if(s==="Data Karyawan") return Utils.success(getKaryawanSheet().getDataRange().getValues()); let sh = s==="Klaim KB"?CONFIG.DATABASE.SHEET_KLAIM.KB:(s==="Klaim Persalinan"?CONFIG.DATABASE.SHEET_KLAIM.PERSALINAN:(s==="Klaim Kehamilan"?CONFIG.DATABASE.SHEET_KLAIM.PEMERIKSAAN_KEHAMILAN:(s==="Klaim Kacamata"?CONFIG.DATABASE.SHEET_KLAIM.KACAMATA:CONFIG.DATABASE.SHEET_KLAIM.GIGI))); return Utils.success(getKlaimSheet(sh).getDataRange().getValues()); } catch (e) { return Utils.error(e.message); } }
