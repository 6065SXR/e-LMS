/**
 * ============================================================================
 * MODUL: KALKULASI & REKAP UPAH LEMBUR (PP NO. 35 TAHUN 2021)
 * System: e-LMS Data Center KSPS
 * Responsibility:
 * 1. Otomasi murni penentuan Tipe Hari (Workday vs Day Off) berdasarkan Kalender Resmi Indonesia.
 * 2. Deteksi Hari Sabtu, Minggu & Tanggal Merah Libur Nasional (Fix/Read-Only).
 * 3. Menyiapkan tab "REKAP_LEMBUR" secara otomatis di Spreadsheet jika belum ada.
 * 4. Menarik data lembur berstatus 'Approved' dari sheet OVERTIME.
 * 5. Pembukuan & sinkronisasi hasil kalkulasi upah lembur ke tab REKAP_LEMBUR.
 * ============================================================================
 */

/**
 * Mengecek apakah tanggal tertentu jatuh pada Hari Libur Nasional Indonesia (Tanggal Merah)
 * @param {string} dateStr Format YYYY-MM-DD
 * @returns {boolean} True jika tanggal merah / libur nasional
 */
function isIndonesianHoliday(dateStr) {
  if (!dateStr) return false;
  
  // Daftar Tanggal Merah Libur Nasional Resmi Kalender Indonesia
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
  
  // Parsing tanggal akurat dengan waktu lokal 00:00:00
  var d = new Date(dateStr + "T00:00:00");
  var dayOfWeek = d.getDay(); // 0 = Minggu, 6 = Sabtu

  // Jika Sabtu, Minggu, atau Tanggal Merah Libur Nasional -> Day Off
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
 * Menarik data lembur berstatus "Approved" untuk periode dan site tertentu,
 * dipadankan dengan data master karyawan & ketentuan gaji per site.
 * Tipe hari dikunci murni secara otomatis dari kalender Indonesia.
 * @param {string} monthPeriod Format YYYY-MM
 * @param {string} targetSite Nama Site (CGK1-CGK4 atau "ALL")
 * @returns {Object} Result object berisi list data kalkulasi lembur
 */
function getApprovedOvertimeForCalculation(monthPeriod, targetSite) {
  try {
    ensureRekapLemburSheet();
    var period = monthPeriod || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var siteFilter = targetSite || "ALL";

    var ovtRows = getSheetDisplayValues("OVERTIME");
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var rekapRows = getSheetDisplayValues("REKAP_LEMBUR");

    // Mapping data rekap yang sudah tersinkronisasi sebelumnya
    var syncedMap = {};
    for (var r = 1; r < rekapRows.length; r++) {
      var ovtId = rekapRows[r][1]; // overtime_id
      if (ovtId) {
        syncedMap[ovtId] = {
          rekap_id: rekapRows[r][0],
          overtime_type: rekapRows[r][8],
          wage_base: parseFloat(rekapRows[r][9]) || 0,
          raw_hours: parseFloat(rekapRows[r][10]) || 0,
          rest_deduction: parseFloat(rekapRows[r][11]) || 0,
          effective_hours: parseFloat(rekapRows[r][12]) || 0,
          rate_hours: parseFloat(rekapRows[r][13]) || 0,
          hourly_wage: parseFloat(rekapRows[r][14]) || 0,
          overtime_amount: parseFloat(rekapRows[r][15]) || 0,
          calculated_at: rekapRows[r][16],
          status_sync: rekapRows[r][17] || "Tersimpan"
        };
      }
    }

    // Standard Wage Base per Site: CGK1, CGK2, CGK3, CGK3A = Rp 5.729.876 | CGK4 = Rp 5.783.676
    function getSiteDefaultWage(siteName) {
      var s = String(siteName || "").toUpperCase().trim();
      if (s === "CGK4") return 5783676;
      return 5729876;
    }

    // Master Data Karyawan Map
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
    for (var i = 1; i < ovtRows.length; i++) {
      var r = ovtRows[i];
      var ovtId = r[0];
      var uId = r[1];
      var tgl = r[2]; // YYYY-MM-DD
      var statusApp = r[9]; // Status Approval

      // Hanya proses lembur yang disetujui (Approved)
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
      var existingSync = syncedMap[ovtId];

      // OTOMATISASI & TERKUNCI MUTLAK: Tipe Hari Murni Dari Tanggal Actual Kalender Indonesia
      var autoType = determineAutoDayType(tgl);

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
        overtime_type: autoType, // Selalu murni otomatis dari tanggal
        auto_detected_type: autoType,
        is_synced: !!existingSync,
        synced_data: existingSync || null
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

/**
 * Menyimpan / memperbarui pembukuan data kalkulasi lembur ke sheet REKAP_LEMBUR.
 * @param {Array<Object>} calculatedDataList List objek kalkulasi lembur
 * @returns {Object} Output status eksekusi
 */
function saveCalculatedOvertimeToSheet(calculatedDataList) {
  try {
    var sheet = ensureRekapLemburSheet();
    if (!calculatedDataList || !Array.isArray(calculatedDataList) || calculatedDataList.length === 0) {
      return { success: false, error: "Tidak ada data kalkulasi untuk disimpan." };
    }

    var existingRows = sheet.getDataRange().getDisplayValues();
    var ovtRowMap = {};
    for (var i = 1; i < existingRows.length; i++) {
      var ovtId = existingRows[i][1];
      if (ovtId) {
        ovtRowMap[ovtId] = i + 1; // 1-based index
      }
    }

    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var savedCount = 0;

    calculatedDataList.forEach(function(item) {
      var targetRowIdx = ovtRowMap[item.overtime_id];
      var rekapId = targetRowIdx ? existingRows[targetRowIdx - 1][0] : generateSequentialId("REKAP_LEMBUR", "RKL", savedCount);

      var rowValues = [
        rekapId,
        item.overtime_id,
        item.user_id,
        item.nik,
        item.nama_lengkap,
        item.site,
        item.periode,
        item.tanggal,
        item.overtime_type,
        item.wage_base,
        item.raw_hours,
        item.rest_deduction,
        item.effective_hours,
        item.rate_hours,
        item.hourly_wage,
        item.overtime_amount,
        dateStr,
        "Tersimpan"
      ];

      if (targetRowIdx) {
        sheet.getRange(targetRowIdx, 1, 1, rowValues.length).setValues([rowValues]);
      } else {
        sheet.appendRow(rowValues);
        savedCount++;
      }
    });

    SpreadsheetApp.flush();
    return {
      success: true,
      message: "Berhasil menyinkronkan " + calculatedDataList.length + " data rekapitulasi upah lembur ke spreadsheet!"
    };
  } catch (err) {
    return { success: false, error: "Gagal menyimpan rekap lembur: " + err.toString() };
  }
}
