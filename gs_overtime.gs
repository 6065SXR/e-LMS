/**
 * Overtime / Form Lembur & Dedicated KJK Sheet Backend Controller
 * System: e-LMS Data Center KSPS
 */

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

  // 4. Pastikan tab KJK sudah ada
  ensureKjkSheet();
}

function getOvertimeData(userId, monthPeriod) {
  try {
    ensureDatabaseColumns();
    var period = monthPeriod || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var overtimeRows = getSheetDisplayValues("OVERTIME");
    var userRows = getSheetDisplayValues("USERS");
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var kjkRows = getSheetDisplayValues("KJK");

    var list = [];
    var totalHours = 0;
    var kjkMap = {};

    // 1. Baca data dari sheet khusus KJK
    for (var k = 1; k < kjkRows.length; k++) {
      var kUid = kjkRows[k][1];
      var kMonthKey = String(kjkRows[k][6] || "").trim().toLowerCase();
      var kHours = parseFloat(kjkRows[k][5]) || 0;
      kjkMap[kUid + "_" + kMonthKey] = kHours;
    }

    // Cache user signatures
    var userSigMap = {};
    for (var u = 1; u < userRows.length; u++) {
      userSigMap[userRows[u][0]] = userRows[u][11] || "";
    }

    // 2. Baca data dari sheet OVERTIME (hanya lembur shift harian murni)
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

/**
 * Menyimpan data KJK langsung ke sheet khusus KJK
 * Kolom: KJK ID, User ID, NIK, Nama Lengkap, Site, Total Jam KJK, Bulan, Created At
 */
function saveKjkEntry(data) {
  try {
    var kjkSheet = ensureKjkSheet();
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var userRows = getSheetDisplayValues("USERS");

    var uid = data.user_id || "USR-0003";
    var monthName = String(data.bulan || "September").trim();
    var totalKjk = parseFloat(data.total_kjk) || 0;

    // Cari NIK, Nama Lengkap, dan Site karyawan dari sheet EMPLOYEES
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

/**
 * Mengambil data KJK per Site dan per Bulan untuk diekspor Excel oleh Project Manager
 * Format output: NIK, Employee (Huruf KAPITAL SEMUA), Amount (Total KJK)
 */
function getKjkExportData(site, monthName) {
  try {
    ensureKjkSheet();
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var kjkRows = getSheetDisplayValues("KJK");
    var targetSite = (site || "ALL").trim();
    var targetMonth = String(monthName || "").trim().toLowerCase();

    // Petakan nilai KJK berdasarkan user_id untuk bulan yang dipilih
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
      var eName = String(empRows[e][3] || "").trim().toUpperCase(); // Format KAPITAL SEMUA
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
