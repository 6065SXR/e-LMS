/**
 * Overtime / Form Lembur, Multi-Slot Attendance Form & Dedicated KJK Sheet Backend Controller
 * System: e-LMS Data Center KSPS
 * Update:
 * 1. Mendukung 3 slot berkas upload (Slot 1: Form Absensi Fix, Slot 2 & 3: Custom Dokumen misal Surat Sakit / Surat Cuti)
 * 2. Fungsi hapus berkas per-slot jika karyawan salah upload file
 * 3. Helper penarikan seluruh berkas PDF Base64 untuk Cetak Gabungan
 */

// ID Folder Google Drive target untuk menyimpan berkas upload karyawan
var ATTENDANCE_FOLDER_ID = "1n289pzwIJ93e-3cTc-w2EazTGc6w_BcY";

/**
 * Memastikan sheet database KJK tersedia dengan struktur kolom presisi
 */
function ensureKjkSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var kjkSheet = ss.getSheetByName("KJK");
  if (!kjkSheet) {
    kjkSheet = ss.insertSheet("KJK");
    kjkSheet.appendRow([
      "KJK ID", "User ID", "NIK", "Nama Lengkap", "Site", "Total Jam KJK", "Bulan", "Created At"
    ]);
    var headerRange = kjkSheet.getRange(1, 1, 1, 8);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#1b2559");
    headerRange.setFontColor("#ffffff");
    kjkSheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  }
  return kjkSheet;
}

/**
 * Memastikan sheet database ATTENDANCE siap mendukung 3 slot dokumen upload
 */
function ensureAttendanceSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ATTENDANCE");
  if (!sheet) {
    sheet = ss.insertSheet("ATTENDANCE");
    sheet.appendRow([
      "Attendance ID", "User ID", "Periode", "Slot Key", "Doc Name", "File Name", "File URL", "Created At"
    ]);
    var headerRange = sheet.getRange(1, 1, 1, 8);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#1b2559");
    headerRange.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  } else {
    // Memastikan jika ada sheet ATTENDANCE versi lama (6 kolom), dinaikkan menjadi 8 kolom
    var lastCol = sheet.getLastColumn();
    if (lastCol < 8 && sheet.getLastRow() > 0) {
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (headers.indexOf("Slot Key") === -1) {
        sheet.getRange(1, 4).setValue("Slot Key");
        sheet.getRange(1, 5).setValue("Doc Name");
        sheet.getRange(1, 6).setValue("File Name");
        sheet.getRange(1, 7).setValue("File URL");
        sheet.getRange(1, 8).setValue("Created At");
      }
    }
  }
  return sheet;
}

/**
 * Memastikan kolom database OVERTIME, USERS, EMPLOYEES, KJK & ATTENDANCE siap digunakan
 */
function ensureDatabaseColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Pastikan kolom di tab OVERTIME
  var ovtSheet = ss.getSheetByName("OVERTIME");
  if (ovtSheet && ovtSheet.getLastRow() > 0) {
    var headers = ovtSheet.getRange(1, 1, 1, ovtSheet.getLastColumn()).getValues()[0];
    var requiredCols = [
      "Overtime ID", "User ID", "Tanggal", "Jam Mulai", "Jam Selesai",
      "Total Jam", "Deskripsi", "Catatan", "Evidence URL", "Status",
      "Created At", "Catatan Revisi", "Approved By", "Approver Signature", "User Signature"
    ];
    
    if (headers.length < requiredCols.length) {
      for (var c = headers.length; c < requiredCols.length; c++) {
        ovtSheet.getRange(1, c + 1).setValue(requiredCols[c]);
      }
    }
  }

  // 2. Pastikan kolom signature di tab USERS
  var userSheet = ss.getSheetByName("USERS");
  if (userSheet && userSheet.getLastRow() > 0) {
    var userHeaders = userSheet.getRange(1, 1, 1, userSheet.getLastColumn()).getValues()[0];
    if (userHeaders.length < 12) {
      userSheet.getRange(1, 12).setValue("signature_data");
    }
  }

  // 3. Pastikan kolom signature di tab EMPLOYEES
  var empSheet = ss.getSheetByName("EMPLOYEES");
  if (empSheet && empSheet.getLastRow() > 0) {
    var empHeaders = empSheet.getRange(1, 1, 1, empSheet.getLastColumn()).getValues()[0];
    if (empHeaders.length < 12) {
      empSheet.getRange(1, 12).setValue("signature_data");
    }
  }

  // 4. Pastikan tab KJK dan ATTENDANCE sudah ada
  ensureKjkSheet();
  ensureAttendanceSheet();
}

function getOvertimeData(userId, monthPeriod) {
  try {
    ensureDatabaseColumns();
    var period = monthPeriod || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var overtimeRows = getSheetDisplayValues("OVERTIME");
    var userRows = getSheetDisplayValues("USERS");
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var kjkRows = getSheetDisplayValues("KJK");
    var attendanceRows = getSheetDisplayValues("ATTENDANCE");

    var list = [];
    var totalHours = 0;
    var kjkMap = {};
    var attendanceMap = {};

    // 1. Baca data dari sheet khusus KJK
    for (var k = 1; k < kjkRows.length; k++) {
      var kUid = kjkRows[k][1];
      var kMonthKey = String(kjkRows[k][6] || "").trim().toLowerCase();
      var kHours = parseFloat(kjkRows[k][5]) || 0;
      kjkMap[kUid + "_" + kMonthKey] = kHours;
    }

    // 2. Baca data Form Absensi & Dokumen Pendukung (3 Slot) dari sheet ATTENDANCE
    for (var a = 1; a < attendanceRows.length; a++) {
      var aUid = attendanceRows[a][1];
      var aPer = attendanceRows[a][2];
      if (!period || aPer === period) {
        if (!attendanceMap[aUid]) {
          attendanceMap[aUid] = {};
        }

        // Cek struktur kolom versi baru (8 kolom) vs versi lama (6 kolom)
        var slotKey = attendanceRows[a][3] || "slot1";
        var docName = attendanceRows[a][4] || "Form Absensi";
        var fileName = attendanceRows[a][5] || attendanceRows[a][3] || "Dokumen.pdf";
        var fileUrl = attendanceRows[a][6] || attendanceRows[a][4] || "#";

        // Jika terdeteksi baris lama di mana kolom 3 adalah nama file
        if (slotKey.indexOf(".pdf") !== -1 || slotKey.indexOf("Form_") === 0) {
          fileName = attendanceRows[a][3];
          fileUrl = attendanceRows[a][4];
          slotKey = "slot1";
          docName = "Form Absensi";
        }

        attendanceMap[aUid][slotKey] = {
          attendance_id: attendanceRows[a][0],
          slot_key: slotKey,
          doc_name: docName,
          fileName: fileName,
          fileUrl: fileUrl
        };
      }
    }

    // Cache user signatures
    var userSigMap = {};
    for (var u = 1; u < userRows.length; u++) {
      userSigMap[userRows[u][0]] = userRows[u][11] || "";
    }

    // Pemetaan NIK berdasarkan Nama & User ID untuk mencari NIK Approver/PM
    var empMapByName = {};
    var empMapByUid = {};
    for (var ep = 1; ep < empRows.length; ep++) {
      var epUid = empRows[ep][1];
      var epNik = empRows[ep][2];
      var epName = String(empRows[ep][3] || "").trim();
      if (epUid) empMapByUid[epUid] = epNik;
      if (epName) empMapByName[epName.toLowerCase()] = epNik;
    }

    // 3. Baca data dari sheet OVERTIME (hanya lembur shift harian murni)
    for (var i = 1; i < overtimeRows.length; i++) {
      var r = overtimeRows[i];
      var uid = r[1];
      var tgl = r[2]; // YYYY-MM-DD
      var rPeriod = tgl ? tgl.substring(0, 7) : "";

      // Abaikan jika masih ada entri lama bertanda KJK di sheet OVERTIME
      if (r[0].indexOf("KJK-") === 0 || r[6] === "REKAPITULASI KELEBIHAN JAM KERJA (KJK)") {
        continue;
      }

      if ((userId === "ALL" || uid === userId) && (!period || rPeriod === period)) {
        var hours = parseFloat(r[5]) || 0;
        totalHours += hours;

        var empName = "-", empNik = "-", site = "-", dept = "-";
        for (var e = 1; e < empRows.length; e++) {
          if (empRows[e][1] === uid) {
            empNik = empRows[e][2];
            empName = empRows[e][3];
            dept = empRows[e][6];
            site = empRows[e][10] || "CGK3";
            break;
          }
        }

        if (empName === "-") {
          for (var u2 = 1; u2 < userRows.length; u2++) {
            if (userRows[u2][0] === uid) {
              empName = userRows[u2][1];
              dept = userRows[u2][4] || "Operations";
              site = userRows[u2][10] || "CGK3";
              break;
            }
          }
        }

        var userSig = r[14] || userSigMap[uid] || "";
        var approverName = r[12] || "";
        var approverNik = "7268900080"; // Default NIK Project Manager

        if (approverName) {
          var cleanApp = approverName.trim().toLowerCase();
          if (empMapByName[cleanApp]) {
            approverNik = empMapByName[cleanApp];
          } else if (cleanApp.indexOf("tommy") !== -1) {
            approverNik = "7268900080";
          } else if (cleanApp.indexOf("budi") !== -1) {
            approverNik = "7268900081";
          }
        }

        list.push({
          overtime_id: r[0],
          user_id: uid,
          nik: empNik,
          nama_lengkap: empName,
          site: site,
          departemen: dept,
          tanggal: r[2],
          jam_mulai: r[3],
          jam_selesai: r[4],
          total_jam: hours,
          deskripsi: r[6],
          catatan: r[7],
          status_approval: r[9] || "Pending",
          created_at: r[10],
          catatan_revisi: r[11] || "",
          approved_by: approverName,
          approver_nik: approverNik,
          approver_signature: r[13] || "",
          user_signature: userSig
        });
      }
    }

    return {
      success: true,
      data: list,
      totalHours: Math.round(totalHours * 10) / 10,
      period: period,
      kjkMap: kjkMap,
      attendanceMap: attendanceMap
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function saveOvertimeEntry(data) {
  try {
    ensureDatabaseColumns();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("OVERTIME");

    if (!sheet) {
      sheet = ss.insertSheet("OVERTIME");
      sheet.appendRow([
        "Overtime ID", "User ID", "Tanggal", "Jam Mulai", "Jam Selesai",
        "Total Jam", "Deskripsi", "Catatan", "Evidence URL", "Status", "Created At",
        "Catatan Revisi", "Approved By", "Approver Signature", "User Signature"
      ]);
    }

    var userSig = data.user_signature || "";
    if (!userSig) {
      var userRows = getSheetDisplayValues("USERS");
      for (var u = 1; u < userRows.length; u++) {
        if (userRows[u][0] === data.user_id) {
          userSig = userRows[u][11] || "";
          break;
        }
      }
    }

    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    if (data.overtime_id && String(data.overtime_id).trim() !== "") {
      var rows = sheet.getDataRange().getDisplayValues();
      var foundRowIdx = -1;
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.overtime_id) {
          foundRowIdx = i + 1;
          break;
        }
      }

      if (foundRowIdx > -1) {
        sheet.getRange(foundRowIdx, 3).setValue(data.tanggal);
        sheet.getRange(foundRowIdx, 4).setValue(data.jam_mulai);
        sheet.getRange(foundRowIdx, 5).setValue(data.jam_selesai);
        sheet.getRange(foundRowIdx, 6).setValue(data.total_jam);
        sheet.getRange(foundRowIdx, 7).setValue(data.deskripsi);
        sheet.getRange(foundRowIdx, 8).setValue(data.catatan || "-");
        sheet.getRange(foundRowIdx, 10).setValue("Pending");
        sheet.getRange(foundRowIdx, 11).setValue(dateStr);
        if (userSig) sheet.getRange(foundRowIdx, 15).setValue(userSig);

        SpreadsheetApp.flush();
        return { success: true, message: "Koreksi lembur berhasil disimpan & diajukan kembali ke PM!" };
      }
    }

    var nextId = generateSequentialId("OVERTIME", "OVT");
    var row = [
      nextId,
      data.user_id,
      data.tanggal,
      data.jam_mulai,
      data.jam_selesai,
      data.total_jam,
      data.deskripsi,
      data.catatan || "-",
      "-",
      "Pending",
      dateStr,
      "",
      "",
      "",
      userSig
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    return { success: true, message: "Pengajuan lembur berhasil disimpan!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function deleteOvertimeEntry(overtimeId, pmUserId) {
  try {
    ensureDatabaseColumns();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("OVERTIME");
    if (!sheet) return { success: false, error: "Sheet OVERTIME tidak ditemukan!" };

    var rows = sheet.getDataRange().getDisplayValues();
    var targetIdx = -1;

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === overtimeId) {
        targetIdx = i + 1;
        break;
      }
    }

    if (targetIdx === -1) {
      return { success: false, error: "Data lembur tidak ditemukan atau sudah dihapus!" };
    }

    sheet.deleteRow(targetIdx);
    SpreadsheetApp.flush();

    return { 
      success: true, 
      message: "Data lembur (" + overtimeId + ") berhasil dihapus!" 
    };
  } catch (err) {
    return { success: false, error: "Gagal menghapus lembur: " + err.toString() };
  }
}

/**
 * Menyimpan Form Absensi & Dokumen Pendukung (3 Slot) ke Google Drive & Database ATTENDANCE
 */
function saveAttendanceForm(userId, period, slotKey, docName, fileData) {
  try {
    var sheet = ensureAttendanceSheet();
    var uid = userId || "USR-0003";
    var per = period || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var sKey = slotKey || "slot1";
    var dName = docName || (sKey === "slot1" ? "Form Absensi" : "Dokumen Pendukung");
    
    if (!fileData || !fileData.base64) {
      return { success: false, error: "Berkas PDF tidak ditemukan atau kosong!" };
    }

    var cleanDocName = dName.replace(/[^a-zA-Z0-9_\-\s]/g, "").trim().replace(/\s+/g, "_");
    var fileName = fileData.name || (cleanDocName + "_" + uid + "_" + per + ".pdf");
    
    // Menyimpan file langsung ke folder target spesifik Google Drive
    var fileUrl = uploadFileToDrive(fileData.base64, fileName, "application/pdf", ATTENDANCE_FOLDER_ID);
    
    if (!fileUrl || fileUrl === "#" || fileUrl.indexOf("http") === -1) {
      return { success: false, error: "Gagal menyimpan berkas ke folder Google Drive target!" };
    }

    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    var rows = sheet.getDataRange().getDisplayValues();
    var foundIdx = -1;
    for (var i = 1; i < rows.length; i++) {
      var rSlot = rows[i][3] || "slot1";
      if (rows[i][1] === uid && rows[i][2] === per && rSlot === sKey) {
        foundIdx = i + 1;
        break;
      }
    }

    if (foundIdx > -1) {
      sheet.getRange(foundIdx, 4).setValue(sKey);
      sheet.getRange(foundIdx, 5).setValue(dName);
      sheet.getRange(foundIdx, 6).setValue(fileName);
      sheet.getRange(foundIdx, 7).setValue(fileUrl);
      sheet.getRange(foundIdx, 8).setValue(dateStr);
    } else {
      var nextId = generateSequentialId("ATTENDANCE", "ATT");
      sheet.appendRow([nextId, uid, per, sKey, dName, fileName, fileUrl, dateStr]);
    }
    SpreadsheetApp.flush();

    return {
      success: true,
      message: dName + " berhasil disimpan ke sistem dan Google Drive!",
      slotKey: sKey,
      docName: dName,
      fileUrl: fileUrl,
      fileName: fileName
    };
  } catch (err) {
    return { success: false, error: "Gagal mengunggah berkas: " + err.toString() };
  }
}

/**
 * Menghapus berkas upload form absensi atau dokumen pendukung per-slot
 */
function deleteAttendanceFile(userId, period, slotKey) {
  try {
    var sheet = ensureAttendanceSheet();
    var uid = userId || "USR-0003";
    var per = period || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var sKey = slotKey || "slot1";

    var rows = sheet.getDataRange().getDisplayValues();
    var targetIdx = -1;
    var fileUrl = "";
    var docName = "";

    for (var i = 1; i < rows.length; i++) {
      var rSlot = rows[i][3] || "slot1";
      if (rows[i][1] === uid && rows[i][2] === per && rSlot === sKey) {
        targetIdx = i + 1;
        docName = rows[i][4] || "Dokumen";
        fileUrl = rows[i][6] || rows[i][4] || "";
        break;
      }
    }

    if (targetIdx === -1) {
      return { success: false, error: "Berkas tidak ditemukan atau sudah dihapus sebelumnya." };
    }

    // Coba hapus file dari Google Drive jika ada fileId valid
    if (fileUrl && fileUrl.indexOf("drive.google.com") !== -1) {
      try {
        var matchId = fileUrl.match(/\/d\/([a-zA-Z0-9_-]+)/) || fileUrl.match(/id=([a-zA-Z0-9_-]+)/);
        if (matchId && matchId[1]) {
          DriveApp.getFileById(matchId[1]).setTrashed(true);
        }
      } catch (driveErr) {
        Logger.log("Info penghapusan file Drive: " + driveErr.toString());
      }
    }

    sheet.deleteRow(targetIdx);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: (docName || "Berkas") + " berhasil dihapus dari sistem!",
      slotKey: sKey
    };
  } catch (err) {
    return { success: false, error: "Gagal menghapus berkas: " + err.toString() };
  }
}

/**
 * Helper mengambil seluruh stream Base64 dokumen terupload (3 Slot) untuk Cetak Gabungan
 */
function getAllAttendancePdfsBase64(userId, period) {
  try {
    ensureAttendanceSheet();
    var attendanceRows = getSheetDisplayValues("ATTENDANCE");
    var resultFiles = [];

    for (var a = 1; a < attendanceRows.length; a++) {
      var aUid = attendanceRows[a][1];
      var aPer = attendanceRows[a][2];

      if (aUid === userId && aPer === period) {
        var slotKey = attendanceRows[a][3] || "slot1";
        var docName = attendanceRows[a][4] || "Form Absensi";
        var fileUrl = attendanceRows[a][6] || attendanceRows[a][4] || "";

        // Backward compatibility jika baris lama
        if (slotKey.indexOf(".pdf") !== -1 || slotKey.indexOf("Form_") === 0) {
          fileUrl = attendanceRows[a][4];
          slotKey = "slot1";
          docName = "Form Absensi";
        }

        if (fileUrl && fileUrl !== "#" && fileUrl.indexOf("http") !== -1) {
          var pdfRes = getPdfBase64(fileUrl);
          if (pdfRes && pdfRes.success && pdfRes.base64) {
            resultFiles.push({
              slotKey: slotKey,
              docName: docName,
              fileName: attendanceRows[a][5] || "Dokumen.pdf",
              fileUrl: fileUrl,
              base64: pdfRes.base64
            });
          }
        }
      }
    }

    return {
      success: true,
      files: resultFiles
    };
  } catch (err) {
    return { success: false, error: err.toString(), files: [] };
  }
}

/**
 * Menyimpan data KJK langsung ke sheet khusus KJK
 */
function saveKjkEntry(data) {
  try {
    var kjkSheet = ensureKjkSheet();
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var userRows = getSheetDisplayValues("USERS");

    var uid = data.user_id || "USR-0003";
    var monthName = String(data.bulan || "September").trim();
    var totalKjk = parseFloat(data.total_kjk) || 0;

    var empNik = "-", empName = "-", empSite = "CGK1";
    for (var e = 1; e < empRows.length; e++) {
      if (empRows[e][1] === uid) {
        empNik = empRows[e][2];
        empName = empRows[e][3];
        empSite = empRows[e][10] || "CGK1";
        break;
      }
    }

    if (empName === "-") {
      for (var u = 1; u < userRows.length; u++) {
        if (userRows[u][0] === uid) {
          empName = userRows[u][1];
          empSite = userRows[u][10] || "CGK1";
          break;
        }
      }
    }

    var rows = kjkSheet.getDataRange().getDisplayValues();
    var foundRowIdx = -1;
    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1] === uid && String(rows[i][6]).trim().toLowerCase() === monthName.toLowerCase()) {
        foundRowIdx = i + 1;
        break;
      }
    }

    if (foundRowIdx > -1) {
      kjkSheet.getRange(foundRowIdx, 3).setValue(empNik);
      kjkSheet.getRange(foundRowIdx, 4).setValue(empName);
      kjkSheet.getRange(foundRowIdx, 5).setValue(empSite);
      kjkSheet.getRange(foundRowIdx, 6).setValue(totalKjk);
      kjkSheet.getRange(foundRowIdx, 7).setValue(monthName);
      kjkSheet.getRange(foundRowIdx, 8).setValue(dateStr);

      SpreadsheetApp.flush();
      return { success: true, message: "Data KJK periode " + monthName + " berhasil diperbarui di sheet KJK!" };
    } else {
      var nextId = generateSequentialId("KJK", "KJK");
      var newRow = [
        nextId,
        uid,
        empNik,
        empName,
        empSite,
        totalKjk,
        monthName,
        dateStr
      ];

      kjkSheet.appendRow(newRow);
      SpreadsheetApp.flush();
      return { success: true, message: "Data KJK periode " + monthName + " berhasil disimpan ke sheet KJK!" };
    }
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getKjkExportData(site, monthName) {
  try {
    ensureKjkSheet();
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var kjkRows = getSheetDisplayValues("KJK");
    var targetSite = (site || "ALL").trim();
    var targetMonth = String(monthName || "").trim().toLowerCase();

    var kjkMapByUid = {};
    for (var k = 1; k < kjkRows.length; k++) {
      var rMonth = String(kjkRows[k][6] || "").trim().toLowerCase();
      if (!targetMonth || rMonth === targetMonth) {
        var rUid = kjkRows[k][1];
        var rAmount = parseFloat(kjkRows[k][5]) || 0;
        kjkMapByUid[rUid] = rAmount;
      }
    }

    var exportList = [];
    for (var e = 1; e < empRows.length; e++) {
      var eUid = empRows[e][1];
      var eNik = empRows[e][2] || "-";
      var eName = String(empRows[e][3] || "").trim().toUpperCase();
      var eSite = empRows[e][10] || "CGK1";

      if (targetSite === "ALL" || eSite === targetSite) {
        var amount = (kjkMapByUid[eUid] !== undefined) ? kjkMapByUid[eUid] : 0;
        exportList.push({
          nik: eNik,
          employee: eName,
          amount: amount,
          site: eSite
        });
      }
    }

    return { 
      success: true, 
      data: exportList, 
      site: targetSite, 
      month: monthName 
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function approveOvertime(overtimeId, pmUserId, pmSignature, pmName) {
  try {
    ensureDatabaseColumns();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("OVERTIME");
    if (!sheet) return { success: false, error: "Sheet OVERTIME tidak ditemukan!" };

    var rows = sheet.getDataRange().getDisplayValues();
    var targetIdx = -1;

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === overtimeId) {
        targetIdx = i + 1;
        break;
      }
    }

    if (targetIdx === -1) return { success: false, error: "Data lembur tidak ditemukan!" };

    sheet.getRange(targetIdx, 10).setValue("Approved");
    sheet.getRange(targetIdx, 13).setValue(pmName || "Project Manager");
    sheet.getRange(targetIdx, 14).setValue(pmSignature || "");

    SpreadsheetApp.flush();
    return { success: true, message: "Pengajuan lembur berhasil di-approve dan ditandatangani!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function requestOvertimeRevision(overtimeId, revisionNote, pmUserId) {
  try {
    ensureDatabaseColumns();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("OVERTIME");
    if (!sheet) return { success: false, error: "Sheet OVERTIME tidak ditemukan!" };

    var rows = sheet.getDataRange().getDisplayValues();
    var targetIdx = -1;

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === overtimeId) {
        targetIdx = i + 1;
        break;
      }
    }

    if (targetIdx === -1) return { success: false, error: "Data lembur tidak ditemukan!" };

    sheet.getRange(targetIdx, 10).setValue("Perlu Revisi");
    sheet.getRange(targetIdx, 12).setValue(revisionNote || "Mohon periksa kembali detail lembur.");

    SpreadsheetApp.flush();
    return { success: true, message: "Instruksi revisi berhasil dikirim ke karyawan!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function saveUserSignature(userId, signatureBase64) {
  try {
    ensureDatabaseColumns();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("USERS");
    var empSheet = ss.getSheetByName("EMPLOYEES");

    var uRows = getSheetDisplayValues("USERS");
    for (var i = 1; i < uRows.length; i++) {
      if (uRows[i][0] === userId) {
        userSheet.getRange(i + 1, 12).setValue(signatureBase64);
        break;
      }
    }

    var eRows = getSheetDisplayValues("EMPLOYEES");
    for (var j = 1; j < eRows.length; j++) {
      if (eRows[j][1] === userId) {
        empSheet.getRange(j + 1, 12).setValue(signatureBase64);
        break;
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Tanda tangan digital Anda berhasil disimpan!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
