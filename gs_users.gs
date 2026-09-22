/**
 * System Users Directory & Access Control (Super Admin CRUD Engine)
 * System: e-LMS Data Center KSPS
 */

function getUsersList() {
  try {
    var rows = getSheetDisplayValues("USERS");
    var users = [];
    for (var i = 1; i < rows.length; i++) {
      users.push({
        user_id: rows[i][0],
        nama_lengkap: rows[i][1],
        email: rows[i][2],
        role: rows[i][3],
        departemen: rows[i][4],
        status_aktif: rows[i][5],
        password: rows[i][6] || "", // Password akun
        is_verified: rows[i][7],
        created_at: rows[i][9],
        site: rows[i][10] || "CGK1"
      });
    }
    return { success: true, data: users };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Memperbarui data profil, peran, lokasi site, dan password akun oleh Super Admin
 * Sinkronisasi otomatis ke sheet USERS dan EMPLOYEES
 */
function updateUserByAdmin(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("USERS");
    var empSheet = ss.getSheetByName("EMPLOYEES");

    if (!userSheet) return { success: false, error: "Sheet USERS tidak ditemukan!" };

    var targetUid = (data.user_id || "").trim();
    if (!targetUid) return { success: false, error: "User ID tidak valid!" };

    var uRows = userSheet.getDataRange().getDisplayValues();
    var userRowIdx = -1;

    for (var i = 1; i < uRows.length; i++) {
      if (uRows[i][0] === targetUid) {
        userRowIdx = i + 1;
        break;
      }
    }

    if (userRowIdx === -1) {
      return { success: false, error: "Pengguna dengan ID " + targetUid + " tidak ditemukan!" };
    }

    // Update kolom di sheet USERS
    if (data.nama_lengkap) userSheet.getRange(userRowIdx, 2).setValue(data.nama_lengkap);
    if (data.email) userSheet.getRange(userRowIdx, 3).setValue(data.email);
    if (data.role) userSheet.getRange(userRowIdx, 4).setValue(data.role);
    if (data.departemen) userSheet.getRange(userRowIdx, 5).setValue(data.departemen);
    if (data.site) userSheet.getRange(userRowIdx, 11).setValue(data.site);

    // Jika Super Admin menginput password baru, perbarui kolom password (kolom 7)
    if (data.new_password && String(data.new_password).trim() !== "") {
      userSheet.getRange(userRowIdx, 7).setValue(String(data.new_password).trim());
    }

    if (empSheet && empSheet.getLastRow() > 1) {
      var empRows = empSheet.getDataRange().getDisplayValues();
      for (var e = 1; e < empRows.length; e++) {
        if (empRows[e][1] === targetUid) {
          var empRowIdx = e + 1;
          if (data.nama_lengkap) empSheet.getRange(empRowIdx, 4).setValue(data.nama_lengkap);
          if (data.email) empSheet.getRange(empRowIdx, 5).setValue(data.email);
          if (data.departemen) empSheet.getRange(empRowIdx, 7).setValue(data.departemen);
          if (data.site) empSheet.getRange(empRowIdx, 11).setValue(data.site);
          break;
        }
      }
    }

    SpreadsheetApp.flush();
    return { 
      success: true, 
      message: "Data akun " + (data.nama_lengkap || targetUid) + " berhasil diperbarui oleh Super Admin!" 
    };
  } catch (err) {
    return { success: false, error: "Gagal memperbarui pengguna: " + err.toString() };
  }
}

/**
 * Menghapus akun pengguna dari sistem (Proteksi akun Super Admin tidak boleh terhapus)
 */
function deleteUserByAdmin(targetUid) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("USERS");
    var empSheet = ss.getSheetByName("EMPLOYEES");

    if (!userSheet) return { success: false, error: "Sheet USERS tidak ditemukan!" };

    var uRows = userSheet.getDataRange().getDisplayValues();
    var userRowIdx = -1;
    var targetRole = "";

    for (var i = 1; i < uRows.length; i++) {
      if (uRows[i][0] === targetUid) {
        userRowIdx = i + 1;
        targetRole = uRows[i][3];
        break;
      }
    }

    if (userRowIdx === -1) {
      return { success: false, error: "Pengguna tidak ditemukan!" };
    }

    // Proteksi: Super Admin tidak boleh menghapus akun Super Admin utama
    if (targetRole === "Super Admin" && targetUid === "USR-0001") {
      return { success: false, error: "Akun Super Admin Utama terlindungi dan tidak dapat dihapus!" };
    }

    userSheet.deleteRow(userRowIdx);

    if (empSheet && empSheet.getLastRow() > 1) {
      var empRows = empSheet.getDataRange().getDisplayValues();
      for (var e = 1; e < empRows.length; e++) {
        if (empRows[e][1] === targetUid) {
          empSheet.deleteRow(e + 1);
          break;
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Pengguna " + targetUid + " berhasil dihapus dari sistem." };
  } catch (err) {
    return { success: false, error: "Gagal menghapus pengguna: " + err.toString() };
  }
}
