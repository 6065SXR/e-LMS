/**
 * ============================================================================
 * MODUL: KALKULASI & REKAP UPAH LEMBUR (PP NO. 35 TAHUN 2021)
 * System: e-LMS Data Center KSPS
 * Responsibility:
 * 1. Otomasi murni penentuan Tipe Hari (Workday vs Day Off) berdasarkan Kalender Resmi Indonesia.
 * 2. Deteksi Hari Sabtu, Minggu & Tanggal Merah Libur Nasional (Fix/Read-Only).
 * 3. Menyiapkan tab "REKAP_LEMBUR" secara otomatis di Spreadsheet jika belum ada.
 * 4. OTOMATISASI REKAP: Setiap data lembur berstatus 'Approved' langsung otomatis ditulis/disinkronkan ke tab REKAP_LEMBUR.
 * ============================================================================
 */

/**
 * Mengecek apakah tanggal tertentu jatuh pada Hari Libur Nasional Indonesia (Tanggal Merah)
 * @param {string} dateStr Format YYYY-MM-DD
 * @returns {boolean} True jika tanggal merah / libur nasional
 */
function isIndonesianHoliday(dateStr) {
  if (!dateStr) return false;
  
  var holidays = [
    // Libur Nasional Tahunan Tetap (Fixed Date)
    "-01-01", // Tahun Baru Masehi
    "-05-01", // Hari Buruh Internasional
    "-06-01", // Hari Lahir Pancasila
    "-08-17", // Proklamasi Kemerdekaan RI
    "-12-25", // Hari Raya Natal

    // Libur Nasional & Cuti Bersama 2025
    "2025-01-27", "2025-01-29", "2025-03-29", "2025-03-31", "2025-04-01",
    "2025-04-18", "2025-05-12", "2025-05-29", "2025-06-06", "2025-06-27", "2025-09-05",

    // Libur Nasional & Cuti Bersama 2026
    "2026-01-16", "2026-02-17", "2026-03-19", "2026-03-20", "2026-03-21",
    "2026-04-03", "2026-05-14", "2026-05-27", "2026-05-31", "2026-06-16", "2026-08-25",

    // Libur Nasional & Cuti Bersama 2027
    "2027-02-05", "2027-02-06", "2027-03-09", "2027-03-10", "2027-03-26",
    "2027-05-06", "2027-05-16", "2027-05-20", "2027-06-06"
  ];

  var datePart = dateStr.length >= 10 ? dateStr.substring(0, 10) : dateStr;
  var monthDay = datePart.length >= 10 ? datePart.substring(4) : ""; // Format -MM-DD

  return holidays.indexOf(datePart) !== -1 || holidays.indexOf(monthDay) !== -1;
}

/**
 * Penentuan otomatis Tipe Hari (Workday / Day Off) berdasarkan Tanggal Actual (Sabtu/Minggu & Tanggal Merah)
 * @param {string} dateStr Format YYYY-MM-DD
 * @returns {string} "Workday" atau "Day Off"
 */
function determineAutoDayType(dateStr) {
  if (!dateStr) return "Workday";
  
  var d = new Date(dateStr + "T00:00:00");
  var dayOfWeek = d.getDay(); // 0 = Minggu, 6 = Sabtu

  if (dayOfWeek === 0 || dayOfWeek === 6 || isIndonesianHoliday(dateStr)) {
    return "Day Off";
  }
  return "Workday";
}

/**
 * Memastikan sheet database REKAP_LEMBUR tersedia dengan struktur header presisi
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Sheet REKAP_LEMBUR
 */
function ensureRekapLemburSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("REKAP_LEMBUR");
  if (!sheet) {
    sheet = ss.insertSheet("REKAP_LEMBUR");
    var headers = [
      "rekap_id", "overtime_id", "user_id", "nik", "nama_lengkap",
      "site", "periode", "tanggal", "overtime_type", "wage_base",
      "raw_hours", "rest_deduction", "effective_hours", "rate_hours",
      "hourly_wage", "overtime_amount", "calculated_at", "status_sync"
    ];
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#1C1B8E");
    headerRange.setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  }
  return sheet;
}

/**
 * Menghitung matematika PP 35/2021 secara persis di backend
 */
function calculatePP35Backend(rawHours, overtimeType, wageBase) {
  var hourlyWage = Math.round(wageBase / 173);
  var restDeduction = rawHours >= 4 ? 0.5 : 0;
  var effectiveHours = Math.max(0, rawHours - restDeduction);
  effectiveHours = Math.round(effectiveHours * 10) / 10;

  var rateHours = 0;
  if (overtimeType === 'Workday') {
    if (effectiveHours <= 1) {
      rateHours = effectiveHours * 1.5;
    } else {
      rateHours = 1.5 + ((effectiveHours - 1) * 2.0);
    }
  } else {
    if (effectiveHours <= 8) {
      rateHours = effectiveHours * 2.0;
    } else if (effectiveHours <= 9) {
      rateHours = 16.0 + ((effectiveHours - 8) * 3.0);
    } else {
      rateHours = 19.0 + ((effectiveHours - 9) * 4.0);
    }
  }

  rateHours = Math.round(rateHours * 10) / 10;
  var overtimeAmount = Math.round(rateHours * hourlyWage);

  return {
    hourly_wage: hourlyWage,
    rest_deduction: restDeduction,
    effective_hours: effectiveHours,
    rate_hours: rateHours,
    overtime_amount: overtimeAmount
  };
}

/**
 * Menarik data lembur berstatus "Approved" untuk periode dan site tertentu.
 * Mengikutsertakan FITUR OTOMATISASI PENULISAN LANGSUNG KE SHEET REKAP_LEMBUR.
 * @param {string} monthPeriod Format YYYY-MM
 * @param {string} targetSite Nama Site (CGK1-CGK4 atau "ALL")
 * @returns {Object} Result object berisi list data kalkulasi lembur
 */
function getApprovedOvertimeForCalculation(monthPeriod, targetSite) {
  try {
    var rekapSheet = ensureRekapLemburSheet();
    var period = monthPeriod || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var siteFilter = targetSite || "ALL";

    var ovtRows = getSheetDisplayValues("OVERTIME");
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var rekapRows = rekapSheet.getDataRange().getDisplayValues();

    // Map data rekap yang sudah ada di sheet
    var syncedMap = {};
    for (var r = 1; r < rekapRows.length; r++) {
      var ovtId = rekapRows[r][1]; // overtime_id
      if (ovtId) {
        syncedMap[ovtId] = {
          rowIdx: r + 1,
          rekap_id: rekapRows[r][0]
        };
      }
    }

    function getSiteDefaultWage(siteName) {
      var s = String(siteName || "").toUpperCase().trim();
      if (s === "CGK4") return 5783676;
      return 5729876;
    }

    var empMap = {};
    for (var e = 1; e < empRows.length; e++) {
      var uId = empRows[e][1];
      var eNik = empRows[e][2] || "-";
      var eName = empRows[e][3] || "-";
      var eSite = empRows[e][10] || "CGK1";
      var customWage = parseFloat(empRows[e][12]) || 0; 

      empMap[uId] = {
        nik: eNik,
        nama_lengkap: eName,
        site: eSite,
        wage: customWage > 0 ? customWage : getSiteDefaultWage(eSite)
      };
    }

    var list = [];
    var newRowsToAppend = [];
    var nowTimestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    for (var i = 1; i < ovtRows.length; i++) {
      var r = ovtRows[i];
      var ovtId = r[0];
      var uId = r[1];
      var tgl = r[2]; // YYYY-MM-DD
      var statusApp = r[9]; // Status Approval

      if (statusApp !== "Approved") continue;

      var rPeriod = tgl ? tgl.substring(0, 7) : "";
      if (period && rPeriod !== period) continue;

      var empInfo = empMap[uId] || {
        nik: "-",
        nama_lengkap: "Karyawan",
        site: "CGK1",
        wage: 5729876
      };

      if (siteFilter !== "ALL" && empInfo.site !== siteFilter) continue;

      var rawHours = parseFloat(r[5]) || 0;
      var autoType = determineAutoDayType(tgl);
      var calcRes = calculatePP35Backend(rawHours, autoType, empInfo.wage);

      // OTOMATISASI SIMPAN DRAFT KE SHEET JIKA BELUM ADA DI TAB REKAP_LEMBUR
      if (!syncedMap[ovtId]) {
        var generatedRekapId = "RKL-" + Math.floor(100000 + Math.random() * 900000);
        var rowValues = [
          generatedRekapId, ovtId, uId, empInfo.nik, empInfo.nama_lengkap,
          empInfo.site, rPeriod, tgl, autoType, empInfo.wage,
          rawHours, calcRes.rest_deduction, calcRes.effective_hours, calcRes.rate_hours,
          calcRes.hourly_wage, calcRes.overtime_amount, nowTimestamp, "Tersimpan"
        ];
        newRowsToAppend.push(rowValues);
      }

      list.push({
        overtime_id: ovtId,
        user_id: uId,
        nik: empInfo.nik,
        nama_lengkap: empInfo.nama_lengkap,
        site: empInfo.site,
        periode: rPeriod,
        tanggal: tgl,
        jam_mulai: r[3],
        jam_selesai: r[4],
        deskripsi: r[6],
        wage_base: empInfo.wage,
        raw_hours: rawHours,
        overtime_type: autoType,
        is_synced: true,
        hourly_wage: calcRes.hourly_wage,
        rest_deduction: calcRes.rest_deduction,
        effective_hours: calcRes.effective_hours,
        rate_hours: calcRes.rate_hours,
        overtime_amount: calcRes.overtime_amount
      });
    }

    // Jika ada data approved baru yang belum masuk ke REKAP_LEMBUR, langsung append
    if (newRowsToAppend.length > 0) {
      newRowsToAppend.forEach(function(row) {
        rekapSheet.appendRow(row);
      });
      SpreadsheetApp.flush();
    }

    return {
      success: true,
      data: list,
      period: period,
      site: siteFilter
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * HELPER OTOMATISASI:
 * Dipanggil langsung ketika PM menekan tombol Approve pada pengajuan lembur.
 * Langsung membuat baris kalkulasi resmi di tab REKAP_LEMBUR tanpa perlu tombol manual.
 */
function autoSyncApprovedOvertimeEntry(ovtId) {
  try {
    if (!ovtId) return;
    var rekapSheet = ensureRekapLemburSheet();
    var ovtRows = getSheetDisplayValues("OVERTIME");
    var empRows = getSheetDisplayValues("EMPLOYEES");

    var targetOvtRow = null;
    for (var i = 1; i < ovtRows.length; i++) {
      if (ovtRows[i][0] === ovtId) {
        targetOvtRow = ovtRows[i];
        break;
      }
    }

    if (!targetOvtRow) return;

    var uId = targetOvtRow[1];
    var tgl = targetOvtRow[2];
    var rawHours = parseFloat(targetOvtRow[5]) || 0;
    var rPeriod = tgl ? tgl.substring(0, 7) : "";

    function getSiteDefaultWage(siteName) {
      var s = String(siteName || "").toUpperCase().trim();
      if (s === "CGK4") return 5783676;
      return 5729876;
    }

    var empInfo = { nik: "-", nama_lengkap: "Karyawan", site: "CGK1", wage: 5729876 };
    for (var e = 1; e < empRows.length; e++) {
      if (empRows[e][1] === uId) {
        var eSite = empRows[e][10] || "CGK1";
        var customWage = parseFloat(empRows[e][12]) || 0;
        empInfo = {
          nik: empRows[e][2] || "-",
          nama_lengkap: empRows[e][3] || "-",
          site: eSite,
          wage: customWage > 0 ? customWage : getSiteDefaultWage(eSite)
        };
        break;
      }
    }

    var autoType = determineAutoDayType(tgl);
    var calcRes = calculatePP35Backend(rawHours, autoType, empInfo.wage);
    var nowTimestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    // Cek apakah sudah ada di sheet
    var existingRows = rekapSheet.getDataRange().getDisplayValues();
    var targetRowIdx = null;
    for (var r = 1; r < existingRows.length; r++) {
      if (existingRows[r][1] === ovtId) {
        targetRowIdx = r + 1;
        break;
      }
    }

    var rekapId = targetRowIdx ? existingRows[targetRowIdx - 1][0] : "RKL-" + Math.floor(100000 + Math.random() * 900000);
    var rowValues = [
      rekapId, ovtId, uId, empInfo.nik, empInfo.nama_lengkap,
      empInfo.site, rPeriod, tgl, autoType, empInfo.wage,
      rawHours, calcRes.rest_deduction, calcRes.effective_hours, calcRes.rate_hours,
      calcRes.hourly_wage, calcRes.overtime_amount, nowTimestamp, "Tersimpan"
    ];

    if (targetRowIdx) {
      rekapSheet.getRange(targetRowIdx, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      rekapSheet.appendRow(rowValues);
    }
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log("Error autoSyncApprovedOvertimeEntry: " + err.toString());
  }
}
