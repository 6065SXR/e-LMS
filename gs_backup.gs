/**
 * Backup, Restore & Safe Maintenance Controller
 * System: e-LMS Data Center KSPS
 * 
 * PROTEKSI MUTLAK:
 * Sheet COURSES dan QUIZZES dikunci dari operasi pembersihan / penghapusan massal!
 */

var PROTECTED_CORE_SHEETS = ["COURSES", "QUIZZES"];

/**
 * Mengambil seluruh data sheet aktif untuk diunduh sebagai file backup JSON
 */
function exportFullDatabaseBackup() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var backupData = {};
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd_HH-mm-ss');

    for (var i = 0; i < sheets.length; i++) {
      var s = sheets[i];
      var sheetName = s.getName();
      var data = s.getDataRange().getDisplayValues();
      backupData[sheetName] = data;
    }

    return {
      success: true,
      filename: "eLMS_KSPS_Backup_" + timestamp + ".json",
      timestamp: timestamp,
      sheetsCount: Object.keys(backupData).length,
      payload: backupData
    };
  } catch (err) {
    return { success: false, error: "Gagal membuat cadangan database: " + err.toString() };
  }
}

/**
 * Memulihkan data dari file cadangan JSON yang diunggah Super Admin
 */
function restoreDatabaseFromJson(jsonString) {
  try {
    if (!jsonString) {
      return { success: false, error: "Data backup kosong atau tidak terbaca." };
    }

    var parsed = JSON.parse(jsonString);
    var backupSheets = parsed.payload || parsed;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    for (var sheetName in backupSheets) {
      if (!backupSheets.hasOwnProperty(sheetName)) continue;
      
      var rows = backupSheets[sheetName];
      if (!rows || rows.length === 0) continue;

      var targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet) {
        targetSheet = ss.insertSheet(sheetName);
      }

      targetSheet.clearContents();
      var numRows = rows.length;
      var numCols = rows[0].length;
      if (numRows > 0 && numCols > 0) {
        targetSheet.getRange(1, 1, numRows, numCols).setValues(rows);
      }
    }

    SpreadsheetApp.flush();
    return { 
      success: true, 
      message: "Database e-LMS KSPS berhasil dipulihkan dari berkas cadangan!" 
    };
  } catch (err) {
    return { success: false, error: "Gagal memulihkan database: " + err.toString() };
  }
}

/**
 * Melakukan pembersihan data operasional secara selektif dan aman
 * SHEET COURSES DAN QUIZZES TIDAK AKAN PERNAH DIHAPUS OLEH FUNGSI INI!
 */
function executeSafeCleanDatabase(cleanType, confirmSecret) {
  try {
    if (confirmSecret !== "HAPUS-DATA") {
      return { success: false, error: "Kata konfirmasi keamanan tidak valid!" };
    }

    // PROTEKSI GANDA: Menolak jika ada permintaan menghapus materi atau kuis
    if (cleanType === "COURSES" || cleanType === "QUIZZES" || cleanType === "ALL_MATERIALS") {
      return { 
        success: false, 
        error: "DITOLAK: Materi Pelatihan (COURSES) dan Bank Soal Kuis (QUIZZES) dikunci permanen dan dilindungi dari penghapusan massal!" 
      };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var clearedInfo = [];

    if (cleanType === "OVERTIME_AND_KJK") {
      // Hanya menghapus transaksi lembur dan KJK
      var ovtSheet = ss.getSheetByName("OVERTIME");
      if (ovtSheet && ovtSheet.getLastRow() > 1) {
        ovtSheet.getRange(2, 1, ovtSheet.getLastRow() - 1, ovtSheet.getLastColumn()).clearContent();
        clearedInfo.push("Riwayat Transaksi OVERTIME");
      }
      var kjkSheet = ss.getSheetByName("KJK");
      if (kjkSheet && kjkSheet.getLastRow() > 1) {
        kjkSheet.getRange(2, 1, kjkSheet.getLastRow() - 1, kjkSheet.getLastColumn()).clearContent();
        clearedInfo.push("Rekapitulasi KJK");
      }
    } else if (cleanType === "PROGRESS_AND_MISSIONS") {
      // Hanya menghapus progress pengerjaan belajar dan misi bulanan (Kuis & Materi tetap 100% aman)
      var progSheet = ss.getSheetByName("PROGRESS");
      if (progSheet && progSheet.getLastRow() > 1) {
        progSheet.getRange(2, 1, progSheet.getLastRow() - 1, progSheet.getLastColumn()).clearContent();
        clearedInfo.push("Progress & Transkrip Belajar Karyawan");
      }
      var msnSheet = ss.getSheetByName("MISSIONS");
      if (msnSheet && msnSheet.getLastRow() > 1) {
        msnSheet.getRange(2, 1, msnSheet.getLastRow() - 1, msnSheet.getLastColumn()).clearContent();
        clearedInfo.push("Riwayat Penugasan Misi Bulanan");
      }
    } else if (cleanType === "EMPLOYEES_ONLY") {
      // Menghapus data karyawan staf dan user non-admin
      var empSheet = ss.getSheetByName("EMPLOYEES");
      if (empSheet && empSheet.getLastRow() > 1) {
        empSheet.getRange(2, 1, empSheet.getLastRow() - 1, empSheet.getLastColumn()).clearContent();
        clearedInfo.push("Data Master EMPLOYEES");
      }
      
      // Bersihkan user non-admin di sheet USERS
      var userSheet = ss.getSheetByName("USERS");
      if (userSheet && userSheet.getLastRow() > 1) {
        var uRows = userSheet.getDataRange().getDisplayValues();
        var adminRows = [uRows[0]]; // Simpan header
        for (var u = 1; u < uRows.length; u++) {
          if (uRows[u][3] === "Super Admin") {
            adminRows.push(uRows[u]); // Simpan Super Admin
          }
        }
        userSheet.clearContents();
        if (adminRows.length > 0) {
          userSheet.getRange(1, 1, adminRows.length, adminRows[0].length).setValues(adminRows);
        }
        clearedInfo.push("Akun Pengguna Non-Admin");
      }
    } else {
      return { success: false, error: "Tipe pembersihan data tidak dikenali." };
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: "Pembersihan berhasil untuk: " + clearedInfo.join(", ") + ". Materi Training & Kuis Anda tetap aman 100%!"
    };
  } catch (err) {
    return { success: false, error: "Gagal membersihkan data: " + err.toString() };
  }
}
