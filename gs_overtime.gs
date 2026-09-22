/**
 * Overtime / Form Lembur & KJK Backend Controller & Digital Signature Engine
 * System: e-LMS Data Center KSPS
 */

/**
 * Memastikan kolom database OVERTIME, USERS, dan EMPLOYEES mendukung TTD & Revisi
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
}

function getOvertimeData(userId, monthPeriod) {
  try {
    ensureDatabaseColumns();
    var period = monthPeriod || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var overtimeRows = getSheetDisplayValues("OVERTIME");
    var userRows = getSheetDisplayValues("USERS");
    var empRows = getSheetDisplayValues("EMPLOYEES");

    var list = [];
    var totalHours = 0;
    var kjkMap = {};

    // Cache user signatures
    var userSigMap = {};
    for (var u = 1; u < userRows.length; u++) {
      userSigMap[userRows[u][0]] = userRows[u][11] || "";
    }

    for (var i = 1; i < overtimeRows.length; i++) {
      var r = overtimeRows[i];
      var uid = r[1];
      var tgl = r[2]; // YYYY-MM-DD atau Nama Bulan untuk KJK
      var rPeriod = tgl ? tgl.substring(0, 7) : "";

      // Deteksi entri khusus KJK pada sheet OVERTIME
      if (r[0].indexOf("KJK-") === 0 || r[6] === "REKAPITULASI KELEBIHAN JAM KERJA (KJK)") {
        var kjkMonthKey = String(r[2] || "").trim().toLowerCase();
        kjkMap[uid + "_" + kjkMonthKey] = parseFloat(r[5]) || 0;
        continue; // Tidak dimasukkan ke daftar tabel lembur shift harian biasa
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
          approved_by: r[12] || "",
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
      kjkMap: kjkMap
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

    // Ambil TTD default user jika tidak dikirim spesifik
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

    // LOGIKA UPDATE / EDIT LEMBUR (JIKA MEMILIKI overtime_id)
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

    // JIKA PENGAJUAN BARU
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

/**
 * Menyimpan data Rekapitulasi Kelebihan Jam Kerja (KJK) ke Sheet OVERTIME
 * Format: hanya menyimpan nama bulan (misal: "July" / "September") pada kolom Tanggal.
 */
function saveKjkEntry(data) {
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

    var uid = data.user_id || "USR-0003";
    var monthName = String(data.bulan || "September").trim();
    var totalKjk = parseFloat(data.total_kjk) || 0;
    var userSig = data.user_signature || "";

    if (!userSig) {
      var userRows = getSheetDisplayValues("USERS");
      for (var u = 1; u < userRows.length; u++) {
        if (userRows[u][0] === uid) {
          userSig = userRows[u][11] || "";
          break;
        }
      }
    }

    var rows = sheet.getDataRange().getDisplayValues();
    var foundRowIdx = -1;
    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1] === uid && 
         (rows[i][0].indexOf("KJK-") === 0 || rows[i][6] === "REKAPITULASI KELEBIHAN JAM KERJA (KJK)") && 
          String(rows[i][2]).trim().toLowerCase() === monthName.toLowerCase()) {
        foundRowIdx = i + 1;
        break;
      }
    }

    if (foundRowIdx > -1) {
      sheet.getRange(foundRowIdx, 3).setValue(monthName); // Simpan nama bulan saja
      sheet.getRange(foundRowIdx, 6).setValue(totalKjk);
      sheet.getRange(foundRowIdx, 11).setValue(dateStr);
      if (userSig) sheet.getRange(foundRowIdx, 15).setValue(userSig);

      SpreadsheetApp.flush();
      return { success: true, message: "Data KJK periode " + monthName + " berhasil diperbarui!" };
    } else {
      var nextId = generateSequentialId("OVERTIME", "KJK");
      var row = [
        nextId,
        uid,
        monthName, // Simpan nama bulan saja
        "-",
        "-",
        totalKjk,
        "REKAPITULASI KELEBIHAN JAM KERJA (KJK)",
        "KJK",
        "-",
        "Approved",
        dateStr,
        "",
        "",
        "",
        userSig
      ];

      sheet.appendRow(row);
      SpreadsheetApp.flush();
      return { success: true, message: "Data KJK periode " + monthName + " berhasil disimpan!" };
    }
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
