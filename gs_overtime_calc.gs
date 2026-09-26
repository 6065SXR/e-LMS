/**
 * ============================================================================
 * MODUL: KALKULASI & REKAP UPAH LEMBUR (PP 35/2021), KJK (FLAT 2X) & SUMMARY THP
 * System: e-LMS Data Center KSPS
 * Responsibility:
 * 1. Otomasi murni penentuan Tipe Hari (Workday vs Day Off) berdasarkan Kalender Resmi RI.
 * 2. Deteksi Hari Sabtu, Minggu & Tanggal Merah Libur Nasional (Fix/Read-Only).
 * 3. Otomatisasi pemotongan istirahat 0.5 Jam jika Jam Kotor >= 4 Jam.
 * 4. Menyiapkan tab "REKAP_LEMBUR" dan "REKAP_KJK" secara otomatis di Spreadsheet.
 * 5. OTOMATISASI REKAP & SUMMARY PAYROLL: Menghitung total THP per karyawan (Gaji + Lembur + KJK).
 * 6. Matematika Presisi 100% Identik dengan Acuan Excel (PP35, KJK Flat 2x & Summary THP).
 * ============================================================================
 */

function isIndonesianHoliday(dateStr) {
  if (!dateStr) return false;
  
  var holidays = [
    "-01-01", "-05-01", "-06-01", "-08-17", "-12-25",
    "2025-01-27", "2025-01-29", "2025-03-29", "2025-03-31", "2025-04-01",
    "2025-04-18", "2025-05-12", "2025-05-29", "2025-06-06", "2025-06-27", "2025-09-05",
    "2026-01-16", "2026-02-17", "2026-03-19", "2026-03-20", "2026-03-21",
    "2026-04-03", "2026-05-14", "2026-05-27", "2026-05-31", "2026-06-16", "2026-08-25",
    "2027-02-05", "2027-02-06", "2027-03-09", "2027-03-10", "2027-03-26",
    "2027-05-06", "2027-05-16", "2027-05-20", "2027-06-06"
  ];

  var datePart = dateStr.length >= 10 ? dateStr.substring(0, 10) : dateStr;
  var monthDay = datePart.length >= 10 ? datePart.substring(4) : "";

  return holidays.indexOf(datePart) !== -1 || holidays.indexOf(monthDay) !== -1;
}

function determineAutoDayType(dateStr) {
  if (!dateStr) return "Workday";
  
  var d = new Date(dateStr + "T00:00:00");
  var dayOfWeek = d.getDay();

  if (dayOfWeek === 0 || dayOfWeek === 6 || isIndonesianHoliday(dateStr)) {
    return "Day Off";
  }
  return "Workday";
}

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

function ensureRekapKjkSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("REKAP_KJK");
  if (!sheet) {
    sheet = ss.insertSheet("REKAP_KJK");
    var headers = [
      "rekap_kjk_id", "kjk_id", "user_id", "nik", "nama_lengkap",
      "site", "periode", "bulan", "hours", "hours_index", "wage_base",
      "hourly_wage", "kjk_amount", "calculated_at", "status_sync"
    ];
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#D97706");
    headerRange.setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  }
  return sheet;
}

function calculatePP35Backend(rawHours, overtimeType, wageBase) {
  var exactHourlyWage = wageBase / 173;
  var displayHourlyWage = Math.round(exactHourlyWage);

  // Otomatis potongan istirahat 0.5 Jam jika jam kotor >= 4 Jam
  var restDeduction = 0;
  if (rawHours >= 4) {
    restDeduction = 0.5;
  }

  var effectiveHours = Math.round((rawHours - restDeduction) * 10) / 10;
  if (effectiveHours < 0) effectiveHours = 0;

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
  var overtimeAmount = Math.round(rateHours * exactHourlyWage);

  return {
    hourly_wage: displayHourlyWage,
    rest_deduction: restDeduction,
    effective_hours: effectiveHours,
    rate_hours: rateHours,
    overtime_amount: overtimeAmount
  };
}

function calculateKjkBackend(hours, wageBase) {
  var exactHourlyWage = wageBase / 173;
  var displayHourlyWage = Math.round(exactHourlyWage);

  var hoursIndex = Math.round(hours * 2 * 10) / 10;
  var kjkAmount = Math.round(hoursIndex * exactHourlyWage);

  return {
    hours_index: hoursIndex,
    hourly_wage: displayHourlyWage,
    kjk_amount: kjkAmount
  };
}

function parseMonthPeriod(dateStr, createdAtStr) {
  if (!dateStr) return "";
  var s = String(dateStr).trim();
  
  var matchYMD = s.match(/(\d{4})[-/.](\d{1,2})/);
  if (matchYMD) {
    return matchYMD[1] + "-" + ("0" + matchYMD[2]).slice(-2);
  }
  
  var matchMY = s.match(/(\d{1,2})[-/.](\d{4})/);
  if (matchMY) {
    return matchMY[2] + "-" + ("0" + matchMY[1]).slice(-2);
  }
  
  var monthNames = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04", "mei": "05", "may": "05",
    "jun": "06", "jul": "07", "agu": "08", "aug": "08", "sep": "09", "okt": "10",
    "oct": "10", "nov": "11", "des": "12", "dec": "12"
  };
  
  var lower = s.toLowerCase();
  var yearMatch = lower.match(/\d{4}/);
  var year = yearMatch ? yearMatch[0] : "";

  if (!year && createdAtStr) {
    var createdYearMatch = String(createdAtStr).match(/\d{4}/);
    if (createdYearMatch) {
      year = createdYearMatch[0];
    }
  }

  if (!year) {
    year = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy');
  }
  
  for (var key in monthNames) {
    if (lower.indexOf(key) !== -1) {
      return year + "-" + monthNames[key];
    }
  }
  
  if (s.length >= 7) return s.substring(0, 7);
  return s;
}

function getApprovedOvertimeForCalculation(monthPeriod, targetSite) {
  try {
    var rekapSheet = ensureRekapLemburSheet();
    var period = monthPeriod || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var siteFilter = targetSite || "ALL";

    var ovtRows = getSheetDisplayValues("OVERTIME");
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var rekapRows = rekapSheet.getDataRange().getDisplayValues();

    var syncedMap = {};
    for (var r = 1; r < rekapRows.length; r++) {
      var ovtId = rekapRows[r][1];
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
      var tgl = r[2];
      var statusApp = r[9];

      if (statusApp !== "Approved") continue;

      var rPeriod = tgl ? tgl.substring(0, 7) : "";
      if (period && period !== "ALL" && rPeriod !== period) continue;

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

function getApprovedKjkForCalculation(monthPeriod, targetSite) {
  try {
    var rekapKjkSheet = ensureRekapKjkSheet();
    var period = monthPeriod || "";
    var siteFilter = targetSite || "ALL";

    var kjkRows = getSheetDisplayValues("KJK");
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var rekapKjkRows = rekapKjkSheet.getDataRange().getDisplayValues();

    var syncedMap = {};
    for (var r = 1; r < rekapKjkRows.length; r++) {
      var kjkId = rekapKjkRows[r][1];
      if (kjkId) {
        syncedMap[kjkId] = {
          rowIdx: r + 1,
          rekap_kjk_id: rekapKjkRows[r][0]
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

    var headerRow = kjkRows.length > 0 ? kjkRows[0] : [];
    var colKjkId = -1, colUserId = -1, colNik = -1, colName = -1, colSite = -1, colHours = -1, colBulan = -1, colCreatedAt = -1;

    for (var c = 0; c < headerRow.length; c++) {
      var h = String(headerRow[c] || "").toLowerCase().trim();
      if (h.indexOf("kjk id") !== -1 || h === "kjk_id" || h === "id") colKjkId = c;
      else if (h.indexOf("user id") !== -1 || h === "user_id" || h === "user") colUserId = c;
      else if (h === "nik") colNik = c;
      else if (h.indexOf("nama") !== -1 || h === "name") colName = c;
      else if (h === "site" || h === "lokasi") colSite = c;
      else if (h.indexOf("jam") !== -1 || h.indexOf("hour") !== -1 || h === "kjk") colHours = c;
      else if (h.indexOf("bulan") !== -1 || h.indexOf("periode") !== -1) colBulan = c;
      else if (h.indexOf("created") !== -1 || h.indexOf("tanggal") !== -1 || h.indexOf("date") !== -1) colCreatedAt = c;
    }

    if (colKjkId === -1) colKjkId = 0;
    if (colUserId === -1) colUserId = 1;
    if (colNik === -1) colNik = 2;
    if (colName === -1) colName = 3;
    if (colSite === -1) colSite = 4;
    if (colHours === -1) colHours = 5;
    if (colBulan === -1) colBulan = 6;
    if (colCreatedAt === -1) colCreatedAt = 7;

    var list = [];
    var newRowsToAppend = [];
    var nowTimestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    for (var i = 1; i < kjkRows.length; i++) {
      var r = kjkRows[i];
      var kjkId = r[colKjkId] || ("KJK-RAW-" + i);
      var uId = r[colUserId] || "";
      var kjkNik = r[colNik] || "-";
      var kjkName = r[colName] || "Karyawan";
      var kjkSite = r[colSite] || "CGK1";
      var kjkHours = parseFloat(r[colHours]) || 0;
      var kjkBulanRaw = String(r[colBulan] || "").trim();
      var kjkCreatedAtRaw = String(r[colCreatedAt] || "").trim();

      if (!kjkHours || kjkHours <= 0) continue;

      var kjkPeriod = parseMonthPeriod(kjkBulanRaw, kjkCreatedAtRaw);
      if (!kjkPeriod) {
        kjkPeriod = period || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
      }

      if (period && period !== "ALL" && period.trim() !== "") {
        if (kjkPeriod !== period) {
          continue;
        }
      }

      var empInfo = empMap[uId] || {
        nik: kjkNik,
        nama_lengkap: kjkName,
        site: kjkSite,
        wage: getSiteDefaultWage(kjkSite)
      };

      var effectiveSite = empInfo.site || kjkSite;
      if (siteFilter !== "ALL" && effectiveSite !== siteFilter) continue;

      var calcRes = calculateKjkBackend(kjkHours, empInfo.wage);

      if (!syncedMap[kjkId]) {
        var generatedRekapKjkId = "RKJ-" + Math.floor(100000 + Math.random() * 900000);
        var rowValues = [
          generatedRekapKjkId, kjkId, uId, empInfo.nik, empInfo.nama_lengkap,
          effectiveSite, period || kjkPeriod, kjkPeriod, kjkHours, calcRes.hours_index, empInfo.wage,
          calcRes.hourly_wage, calcRes.kjk_amount, nowTimestamp, "Tersimpan"
        ];
        newRowsToAppend.push(rowValues);
      }

      list.push({
        kjk_id: kjkId,
        user_id: uId,
        nik: empInfo.nik,
        nama_lengkap: empInfo.nama_lengkap,
        site: effectiveSite,
        periode: period || kjkPeriod,
        bulan: kjkPeriod,
        hours: kjkHours,
        hours_index: calcRes.hours_index,
        wage_base: empInfo.wage,
        hourly_wage: calcRes.hourly_wage,
        kjk_amount: calcRes.kjk_amount,
        is_synced: true
      });
    }

    if (newRowsToAppend.length > 0) {
      newRowsToAppend.forEach(function(row) {
        rekapKjkSheet.appendRow(row);
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

function getSummaryPayrollForCalculation(monthPeriod, targetSite) {
  try {
    var period = monthPeriod || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var siteFilter = targetSite || "ALL";

    // 1. Get Overtime Data
    var ovtRes = getApprovedOvertimeForCalculation(period, siteFilter);
    var ovtList = (ovtRes && ovtRes.success) ? ovtRes.data : [];
    
    var ovtAggMap = {};
    for (var i = 0; i < ovtList.length; i++) {
      var item = ovtList[i];
      var uId = item.user_id;
      if (!ovtAggMap[uId]) ovtAggMap[uId] = 0;
      ovtAggMap[uId] += (parseFloat(item.overtime_amount) || 0);
    }

    // 2. Get KJK Data
    var kjkRes = getApprovedKjkForCalculation(period, siteFilter);
    var kjkList = (kjkRes && kjkRes.success) ? kjkRes.data : [];

    var kjkAggMap = {};
    for (var j = 0; j < kjkList.length; j++) {
      var kItem = kjkList[j];
      var kUid = kItem.user_id;
      if (!kjkAggMap[kUid]) kjkAggMap[kUid] = 0;
      kjkAggMap[kUid] += (parseFloat(kItem.kjk_amount) || 0);
    }

    // 3. Get All Active Employees
    var empRows = getSheetDisplayValues("EMPLOYEES");
    function getSiteDefaultWage(siteName) {
      var s = String(siteName || "").toUpperCase().trim();
      if (s === "CGK4") return 5783676;
      return 5729876;
    }

    var list = [];

    for (var e = 1; e < empRows.length; e++) {
      var r = empRows[e];
      var uId = r[1];
      if (!uId) continue;
      
      var eNik = r[2] || "-";
      var eName = r[3] || "-";
      var eSite = r[10] || "CGK1";
      var eStatus = r[8] || "Aktif";
      
      if (eStatus.toLowerCase() === 'non-aktif' || eStatus.toLowerCase() === 'non aktif') continue;
      if (siteFilter !== "ALL" && eSite !== siteFilter) continue;

      var customWage = parseFloat(r[12]) || 0;
      var wageBase = customWage > 0 ? customWage : getSiteDefaultWage(eSite);

      var ovtAmount = ovtAggMap[uId] || 0;
      var kjkAmount = kjkAggMap[uId] || 0;
      var thp = wageBase + ovtAmount + kjkAmount;

      list.push({
        user_id: uId,
        nik: eNik,
        nama_lengkap: eName,
        site: eSite,
        periode: period,
        wage_base: wageBase,
        overtime_amount: ovtAmount,
        kjk_amount: kjkAmount,
        take_home_pay: thp
      });
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
