/**
 * CONTROLLER KARYAWAN - HRIS KOBE
 */
function getAllKaryawan() {
  try {
    const sheet = getKaryawanSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return Utils.success([], "Data kosong");

    const headers = data[0];
    const rows = data.slice(1);
    const formattedData = rows.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        const key = header.toString().trim().replace(/\s+/g, '_');
        obj[key] = row[index];
      });
      return obj;
    });
    return Utils.success(formattedData);
  } catch (err) {
    return Utils.error("Gagal mengambil data: " + err.message);
  }
}

function createKaryawan(formData) {
  try {
    const sheet = getKaryawanSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(header => {
      const key = header.toString().trim().replace(/\s+/g, '_');
      if (key === 'No') return sheet.getLastRow();
      return formData[key] || "";
    });
    sheet.appendRow(newRow);
    return Utils.success(null, "Data karyawan berhasil ditambahkan.");
  } catch (err) {
    return Utils.error("Gagal menambah data: " + err.message);
  }
}

function updateKaryawan(formData) {
  try {
    const sheet = getKaryawanSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const nrppIndex = headers.indexOf("NRPP");
    if (nrppIndex === -1) throw new Error("Kolom NRPP tidak ditemukan");

    const rowIndex = data.findIndex(row => row[nrppIndex] == formData.NRPP);
    if (rowIndex === -1) throw new Error("Data tidak ditemukan");

    const updatedRow = headers.map((header, index) => {
      const key = header.toString().trim().replace(/\s+/g, '_');
      const readOnlyFields = ['No', 'Masa_Kerja_All', 'Masa_Kerja_All_(Tahun)', 'Usia', 'Rentang_Usia'];
      if (readOnlyFields.includes(key)) return data[rowIndex][index]; 
      return formData[key] !== undefined ? formData[key] : data[rowIndex][index];
    });

    sheet.getRange(rowIndex + 1, 1, 1, updatedRow.length).setValues([updatedRow]);
    return Utils.success(null, "Data berhasil diperbarui.");
  } catch (err) {
    return Utils.error("Gagal update data: " + err.message);
  }
}
