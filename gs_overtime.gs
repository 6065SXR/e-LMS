/**
 * Overtime / Form Lembur Backend Controller (Tanpa Evidence File)
 * System: e-LMS Data Center KSPS
 */

function getOvertimeData(userId, monthPeriod) {
  try {
    var period = monthPeriod || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var overtimeRows = getSheetDisplayValues("OVERTIME");
    var userRows = getSheetDisplayValues("USERS");
    var empRows = getSheetDisplayValues("EMPLOYEES");

    var list = [];
    var totalHours = 0;

    for (var i = 1; i < overtimeRows.length; i++) {
      var r = overtimeRows[i];
      var uid = r[1];
      var tgl = r[2]; // YYYY-MM-DD
      var rPeriod = tgl ? tgl.substring(0, 7) : "";

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
          for (var u = 1; u < userRows.length; u++) {
            if (userRows[u][0] === uid) {
              empName = userRows[u][1];
              dept = userRows[u][4] || "Operations";
              site = userRows[u][10] || "CGK3";
              break;
            }
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
          created_at: r[10]
        });
      }
    }

    return {
      success: true,
      data: list,
      totalHours: Math.round(totalHours * 10) / 10,
      period: period
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function saveOvertimeEntry(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("OVERTIME");

    if (!sheet) {
      sheet = ss.insertSheet("OVERTIME");
      sheet.appendRow([
        "Overtime ID", "User ID", "Tanggal", "Jam Mulai", "Jam Selesai",
        "Total Jam", "Deskripsi", "Catatan", "Evidence URL", "Status", "Created At"
      ]);
    }

    var nextId = generateSequentialId("OVERTIME", "OVT");
    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    var row = [
      nextId,
      data.user_id,
      data.tanggal,
      data.jam_mulai,
      data.jam_selesai,
      data.total_jam,
      data.deskripsi,
      data.catatan || "-",
      "-", // Evidence URL (Kosong / Tanpa File)
      "Pending",
      dateStr
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    return { success: true, message: "Pengajuan lembur berhasil disimpan!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
