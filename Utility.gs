/**
 * UTILITIES & HELPERS
 */
const Utils = {
  success: function(data, message = "Success") {
    return JSON.stringify({ status: "success", message: message, data: data });
  },
  error: function(message = "Internal Server Error") {
    return JSON.stringify({ status: "error", message: message });
  }
};

/**
 * Fungsi agar file HTML bisa memanggil file HTML lainnya
 * Lokasi: Utility.gs
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
