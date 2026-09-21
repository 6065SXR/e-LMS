/**
 * Employee Data Controller (Kalkulasi Rata-Rata Skor 3-Pilar & Rata-Rata Komposit Multi-Materi)
 * System: e-LMS Data Center KSPS
 */

function getEmployeesList() {
  try {
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var progressRows = getSheetDisplayValues("PROGRESS");

    var list = [];
    for (var i = 1; i < empRows.length; i++) {
      var uid = empRows[i][1];
      
      // Mengelompokkan skor per course_id unik untuk memastikan pembagi (denominator) presisi
      var courseMap = {};

      for (var p = 1; p < progressRows.length; p++) {
        if (progressRows[p][1] === uid) {
          var courseId = progressRows[p][2];
          var pdfVal = parseInt(progressRows[p][9], 10) || (progressRows[p][4] === "TRUE" ? 100 : 0);
          var videoVal = parseInt(progressRows[p][10], 10) || (progressRows[p][5] === "TRUE" ? 100 : 0);
          var quizVal = parseFloat(progressRows[p][6]) || 0;

          // Mengambil pencapaian skor tertinggi untuk setiap course yang diakses
          if (!courseMap[courseId]) {
            courseMap[courseId] = { pdf: pdfVal, video: videoVal, quiz: quizVal };
          } else {
            courseMap[courseId].pdf = Math.max(courseMap[courseId].pdf, pdfVal);
            courseMap[courseId].video = Math.max(courseMap[courseId].video, videoVal);
            courseMap[courseId].quiz = Math.max(courseMap[courseId].quiz, quizVal);
          }
        }
      }

      var totalCourses = Object.keys(courseMap).length;
      var pdfScoreSum = 0, videoScoreSum = 0, quizScoreSum = 0;
      var compositeScoreSum = 0;

      if (totalCourses > 0) {
        for (var cId in courseMap) {
          pdfScoreSum += courseMap[cId].pdf;
          videoScoreSum += courseMap[cId].video;
          quizScoreSum += courseMap[cId].quiz;

          // Hitung Nilai Komposit per-materi (sesuai nilai komposit yang tampil di Katalog Materi)
          var courseComposite = Math.round((0.20 * courseMap[cId].pdf) + (0.30 * courseMap[cId].video) + (0.50 * courseMap[cId].quiz));
          compositeScoreSum += courseComposite;
        }
      }

      // Perhitungan Rata-Rata (Mean Average) dengan pembagi total materi yang dikerjakan
      var pdfAvg = totalCourses > 0 ? Math.round(pdfScoreSum / totalCourses) : 0;
      var videoAvg = totalCourses > 0 ? Math.round(videoScoreSum / totalCourses) : 0;
      var quizAvg = totalCourses > 0 ? Math.round(quizScoreSum / totalCourses) : 0;

      // Nilai Komposit Karyawan diambil dari Rata-Rata Nilai Komposit seluruh materi di Katalog Materi
      var compositeAvg = totalCourses > 0 ? Math.round(compositeScoreSum / totalCourses) : 0;

      list.push({
        employee_id: empRows[i][0],
        user_id: uid,
        nik: empRows[i][2],
        nama_lengkap: empRows[i][3],
        email: empRows[i][4],
        jabatan: empRows[i][5],
        departemen: empRows[i][6],
        no_hp: empRows[i][7],
        status_karyawan: empRows[i][8],
        created_at: empRows[i][9],
        site: empRows[i][10] || "CGK1",
        pdf_score_avg: pdfAvg,
        video_score_avg: videoAvg,
        quiz_score_avg: quizAvg,
        composite_avg: compositeAvg
      });
    }

    return { success: true, data: list };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function saveEmployee(empData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var empSheet = ss.getSheetByName("EMPLOYEES");
    var userSheet = ss.getSheetByName("USERS");

    var nextEmpId = generateSequentialId("EMPLOYEES", "EMP");
    var nextUserId = generateSequentialId("USERS", "USR");
    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var empSite = empData.site || "CGK1";

    empSheet.appendRow([
      nextEmpId, nextUserId, empData.nik, empData.nama_lengkap, empData.email,
      empData.jabatan, empData.departemen, empData.no_hp, "Aktif", dateStr, empSite
    ]);

    userSheet.appendRow([
      nextUserId, empData.nama_lengkap, empData.email, "Karyawan", empData.departemen,
      "Aktif", "dc123456", "TRUE", "ADMIN_CREATED", dateStr, empSite
    ]);

    SpreadsheetApp.flush();
    return { success: true, message: "Data Karyawan baru berhasil ditambahkan!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function bulkUploadEmployees(csvText) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var empSheet = ss.getSheetByName("EMPLOYEES");
    var userSheet = ss.getSheetByName("USERS");

    var lines = csvText.split("\n");
    if (lines.length <= 1) {
      return { success: false, error: "Berkas CSV kosong atau format tidak sesuai!" };
    }

    var addedCount = 0;
    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    for (var i = 1; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;

      var cols = line.split(",");
      if (cols.length >= 6) {
        var nik = cols[0].trim();
        var nama = cols[1].trim();
        var email = cols[2].trim();
        var jabatan = cols[3].trim();
        var dept = cols[4].trim();
        var site = cols[5].trim() || "CGK1";
        var noHp = cols[6] ? cols[6].trim() : "-";

        var nextEmpId = generateSequentialId("EMPLOYEES", "EMP", addedCount);
        var nextUserId = generateSequentialId("USERS", "USR", addedCount);

        empSheet.appendRow([
          nextEmpId, nextUserId, nik, nama, email, jabatan, dept, noHp, "Aktif", dateStr, site
        ]);

        userSheet.appendRow([
          nextUserId, nama, email, "Karyawan", dept, "Aktif", "dc123456", "TRUE", "BULK_UPLOAD", dateStr, site
        ]);

        addedCount++;
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Berhasil mengunggah " + addedCount + " data karyawan baru!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
