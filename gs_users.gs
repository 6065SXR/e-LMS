/**
 * System Users Directory & Access Control
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
