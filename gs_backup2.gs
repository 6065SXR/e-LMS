/**
 * Backup, Restore, Direct Drive Cloning & Safe Maintenance Controller
 * System: e-LMS Data Center KSPS
 * 
 * PROTEKSI MUTLAK:
 * Sheet COURSES dan QUIZZES dikunci dari operasi pembersihan / penghapusan massal!
 */

var PROTECTED_CORE_SHEETS = ["COURSES", "QUIZZES"];
var BACKUP_DRIVE_FOLDER_ID = "1IdDGJn52BxcK8783RUxswaM_8Gsmya_H";
var MAIN_SPREADSHEET_ID = "14qUWKn95f6NYt8UWuTZsTFNoLaosAB5VkQMyvApKcMc";

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
 * Menduplikasi spreadsheet database utama secara langsung ke Folder Google Drive Cadangan
 */
function backupDatabaseToGoogleDrive() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssId = ss.getId() || MAIN_SPREADSHEET_ID;
    var file = DriveApp.getFileById(ssId);
    var targetFolder = DriveApp.getFolderById(BACKUP_DRIVE_FOLDER_ID);
    
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd_HH-mm-ss');
    var backupName = "eLMS_KSPS_Database_Backup_" + timestamp;
    
    var copiedFile = file.makeCopy(backupName, targetFolder);
    
    return {
      success: true,
      message: "Database e-LMS KSPS berhasil diduplikasi ke Google Drive!",
      filename: backupName,
      fileUrl: copiedFile.getUrl(),
      fileId: copiedFile.getId(),
      timestamp: timestamp
    };
  } catch (err) {
    return { success: false, error: "Gagal menduplikasi database ke Google Drive: " + err.toString() };
  }
}

/**
 * Penutupan Buku Akhir Bulan: Duplikasi Arsip ke Drive & Auto-Purge Sheet Transaksi Bulanan
 */
function executeMonthlyArchiveAndReset(confirmSecret) {
  try {
    if (confirmSecret !== "RESET-BULANAN") {
      return { success: false, error: "Kata konfirmasi keamanan salah! Wajib ketik persis 'RESET-BULANAN'." };
    }

    // Step 1: Buat Duplikasi Berkas Arsip ke Folder Google Drive Target
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssId = ss.getId() || MAIN_SPREADSHEET_ID;
    var file = DriveApp.getFileById(ssId);
    var targetFolder = DriveApp.getFolderById(BACKUP_DRIVE_FOLDER_ID);
    
    var currentPeriod = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var archiveName = "eLMS_Database_Archive_" + currentPeriod;
    
    var copiedFile = file.makeCopy(archiveName, targetFolder);
    
    if (!copiedFile) {
      return { success: false, error: "Gagal membuat berkas cadangan arsip di Google Drive. Pembersihan dibatalkan!" };
    }

    // Step 2: Auto-Purge Sheet Transaksi Bulanan (Baris ke-2 ke bawah, Header tetap utuh)
    var sheetsToPurge = ["OVERTIME", "KJK", "REKAP_LEMBUR", "REKAP_KJK", "PROGRESS", "MONTHLY_MISSIONS"];
    var clearedSheets = [];

    sheetsToPurge.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
        clearedSheets.push(sheetName);
      }
    });

    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Penutupan buku bulanan & reset data transaksi berhasil dijalankan!",
      archiveUrl: copiedFile.getUrl(),
      archiveName: archiveName,
      clearedSheets: clearedSheets
    };
  } catch (err) {
    return { success: false, error: "Gagal mengeksekusi arsip & reset bulanan: " + err.toString() };
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

/**
 * Smart Historical Archive Reader:
 * Mencari dan membuka berkas arsip di Google Drive Folder jika data bulan yang dicari
 * tidak ada di Spreadsheet Utama (sudah di-reset saat penutupan buku bulanan).
 */
function getSpreadsheetForPeriod(periodStr, sampleSheetName) {
  var activeSs = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = sampleSheetName || "OVERTIME";
  
  if (!periodStr || periodStr === "ALL") {
    return activeSs;
  }

  // 1. Cek apakah Spreadsheet Utama memiliki data untuk periode ini
  var sheet = activeSs.getSheetByName(sheetName);
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      var rowStr = data[i].join(" ");
      if (rowStr.indexOf(periodStr) !== -1) {
        return activeSs; // Data ditemukan di Spreadsheet Utama
      }
    }
  }

  // 2. Jika data di Spreadsheet Utama kosong/tidak ditemukan, cari file arsip di Folder Drive Cadangan
  try {
    var targetFolderId = BACKUP_DRIVE_FOLDER_ID;
    if (!targetFolderId) return activeSs;

    var folder = DriveApp.getFolderById(targetFolderId);
    var files = folder.getFiles();

    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
      
      // Cocokkan nama file arsip yang mengandung periode (contoh: "2026-09" atau "eLMS_Database_Archive_2026-09")
      if (fileName.indexOf(periodStr) !== -1) {
        return SpreadsheetApp.openById(file.getId());
      }
    }
  } catch (err) {
    Logger.log("Smart Archive Reader Note: " + err.toString());
  }

  return activeSs;
}
