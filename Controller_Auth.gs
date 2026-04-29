/**
 * CONTROLLER AUTHENTICATION - HRIS KOBE 2026
 */
function verifyLogin(username, password) {
  try {
    if (username === CONFIG.AUTH.USERNAME && password === CONFIG.AUTH.PASSWORD) {
      return Utils.success({
        role: 'Super Admin',
        name: 'Admin KOBE'
      }, "Login Berhasil");
    } else {
      return Utils.error("Username atau Password salah!");
    }
  } catch (err) {
    return Utils.error("Terjadi kesalahan server: " + err.message);
  }
}
