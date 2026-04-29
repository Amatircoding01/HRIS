/**
 * MAIN ENTRY POINT - HRIS KOBE
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
      .setTitle(CONFIG.PROJECT_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * BACKEND GENERATOR PDF
 * Lokasi: Main.gs
 */
function generateAbsensiPDF() {
  try {
    const htmlTemplate = HtmlService.createTemplateFromFile('View_Print');
    
    // Kirim data ke template
    htmlTemplate.tanggalCetak = Utilities.formatDate(new Date(), "GMT+7", "dd-MM-yyyy HH:mm");
    
    const htmlContent = htmlTemplate.evaluate().getContent();
    const blob = Utilities.newBlob(htmlContent, 'text/html', 'temp.html');
    const pdfBlob = blob.getAs('application/pdf').setName('Laporan_HRIS_KOBE.pdf');
    
    // --- BAGIAN KRUSIAL: MENGIRIM DATA BALIK KE BROWSER ---
    return {
      status: 'success',
      base64Data: Utilities.base64Encode(pdfBlob.getBytes()),
      fileName: 'Laporan_Absensi_KOBE.pdf'
    };

  } catch (err) {
    return {
      status: 'error',
      message: err.toString()
    };
  }
}
