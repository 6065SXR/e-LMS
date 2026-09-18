/**
 * ZettBOT Apps Script Assistant - Safe Migrate & Initial Database Setup
 * System: e-LMS Data Center Zettbos
 */

var DB_SHEETS = {
  USERS: ['user_id', 'nama_lengkap', 'email', 'role', 'departemen', 'status_aktif', 'password', 'is_verified', 'verification_token', 'created_at', 'site'],
  EMPLOYEES: ['employee_id', 'user_id', 'nik', 'nama_lengkap', 'email', 'jabatan', 'departemen', 'no_hp', 'status_karyawan', 'created_at', 'site'],
  COURSES: ['course_id', 'judul_materi', 'kategori_dc', 'deskripsi_singkat', 'pdf_link', 'video_url', 'created_by', 'created_at'],
  QUIZZES: ['quiz_id', 'course_id', 'pertanyaan', 'opsi_a', 'opsi_b', 'opsi_c', 'opsi_d', 'jawaban_benar', 'passing_score'],
  PROGRESS: ['progress_id', 'user_id', 'course_id', 'periode_bulan', 'is_pdf_downloaded', 'is_video_watched', 'quiz_score', 'status_penyelesaian', 'last_updated', 'pdf_score', 'video_score', 'pages_read_count', 'total_pages', 'video_seconds_watched'],
  MONTHLY_MISSIONS: ['mission_id', 'user_id', 'periode_bulan', 'target_materi', 'materi_selesai', 'status_misi', 'assigned_course_ids', 'last_updated']
};

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Object.keys(DB_SHEETS).forEach(function(sheetName) {
    var expectedHeaders = DB_SHEETS[sheetName];
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(expectedHeaders);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight("bold").setBackground("#1b2559").setFontColor("#ffffff");
    } else {
      // Safe Migrate Header Check
      var lastRow = sheet.getLastRow();
      var currentHeaders = lastRow > 0 ? sheet.getDataRange().getDisplayValues()[0] : [];
      
      if (currentHeaders.length === 0) {
        sheet.appendRow(expectedHeaders);
        sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight("bold").setBackground("#1b2559").setFontColor("#ffffff");
      } else {
        expectedHeaders.forEach(function(header) {
          if (currentHeaders.indexOf(header) === -1) {
            sheet.getRange(1, currentHeaders.length + 1).setValue(header).setFontWeight("bold").setBackground("#1b2559").setFontColor("#ffffff");
            currentHeaders.push(header);
          }
        });
      }
    }
  });

  seedOrUpdateDummyData(ss);
  SpreadsheetApp.flush();
  Logger.log("Inisialisasi Database e-LMS Zettbos Berhasil!");
}

function seedOrUpdateDummyData(ss) {
  var today = new Date();
  var dateStr = Utilities.formatDate(today, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');

  // Seed / Repair Users Table
  var userSheet = ss.getSheetByName("USERS");
  if (userSheet) {
    var userRows = userSheet.getLastRow() > 0 ? userSheet.getDataRange().getDisplayValues() : [];
    
    var dummyUsers = [
      { id: "USR-0001", name: "Admin Data Center", email: "admin@datacenter.com", role: "Super Admin", dept: "IT Infrastructure", pass: "admin123", site: "CGK1" },
      { id: "USR-0002", name: "Budi PM Infrastructure", email: "pm.budi@datacenter.com", role: "Project Manager", dept: "Power & Cooling", pass: "pm123", site: "CGK2" },
      { id: "USR-0003", name: "Eko Tech Operator", email: "eko.tech@datacenter.com", role: "Karyawan", dept: "Operations", pass: "eko123", site: "CGK3" },
      { id: "USR-0004", name: "Rina DC Technician", email: "rina.tech@datacenter.com", role: "Karyawan", dept: "Network & Cabling", pass: "rina123", site: "CGK3A" }
    ];

    dummyUsers.forEach(function(dUser) {
      var foundIndex = -1;
      for (var i = 1; i < userRows.length; i++) {
        if (userRows[i][2] && userRows[i][2].toLowerCase().trim() === dUser.email.toLowerCase()) {
          foundIndex = i + 1;
          break;
        }
      }

      if (foundIndex > -1) {
        userSheet.getRange(foundIndex, 6).setValue("Aktif");
        userSheet.getRange(foundIndex, 7).setValue(dUser.pass);
        userSheet.getRange(foundIndex, 8).setValue("TRUE");
        userSheet.getRange(foundIndex, 11).setValue(dUser.site);
      } else {
        userSheet.appendRow([
          dUser.id, dUser.name, dUser.email, dUser.role, dUser.dept, "Aktif", dUser.pass, "TRUE", "VERIFIED_INIT", dateStr, dUser.site
        ]);
      }
    });
  }

  // Seed Employees Table
  var empSheet = ss.getSheetByName("EMPLOYEES");
  if (empSheet && empSheet.getLastRow() <= 1) {
    var initialEmployees = [
      ["EMP-0001", "USR-0003", "31710123001", "Eko Tech Operator", "eko.tech@datacenter.com", "Junior DC Technician", "Operations", "081234567890", "Aktif", dateStr, "CGK3"],
      ["EMP-0002", "USR-0004", "31710123002", "Rina DC Technician", "rina.tech@datacenter.com", "Cabling Specialist", "Network & Cabling", "081298765432", "Aktif", dateStr, "CGK3A"]
    ];
    empSheet.getRange(2, 1, initialEmployees.length, initialEmployees[0].length).setValues(initialEmployees);
  }

  // Seed / Update Courses Table with Valid Video Links
  var courseSheet = ss.getSheetByName("COURSES");
  if (courseSheet) {
    var sampleVideoUrl = "https://drive.google.com/file/d/15YCCYywuOOPVhr_J_0DUdx4p-hA6uKF8/view?usp=sharing";
    var cRows = courseSheet.getLastRow() > 0 ? courseSheet.getDataRange().getDisplayValues() : [];

    if (cRows.length <= 1) {
      var initialCourses = [
        ["CRS-0001", "Cooling System CRAH & Chiller SOP", "Introduction", "SOP dan panduan pemeliharaan rutin sistem pendingin Precision Cooling CRAH Data Center.", "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf", sampleVideoUrl, "Project Manager", dateStr],
        ["CRS-0002", "Redundansi UPS N+1 Maintenance", "Data Center Fundamental (Basic)", "Prosedur operasional standar pemeliharaan baterai dan modul UPS tanpa mengganggu SLA.", "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf", sampleVideoUrl, "Project Manager", dateStr]
      ];
      courseSheet.getRange(2, 1, initialCourses.length, initialCourses[0].length).setValues(initialCourses);
    } else {
      for (var k = 1; k < cRows.length; k++) {
        var vUrl = cRows[k][5];
        if (!vUrl || vUrl === '#' || vUrl.trim() === '') {
          courseSheet.getRange(k + 1, 6).setValue(sampleVideoUrl);
        }
      }
    }
  }
}
