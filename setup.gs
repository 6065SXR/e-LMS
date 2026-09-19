/**
 * Script Setup Inisialisasi Database Google Sheets
 * System: e-LMS Data Center Zettbos
 * 
 * Jalankan fungsi setupDatabase() ini sekali di Google Apps Script Editor 
 * untuk mereset/membuat 6 tab sheet dengan struktur header & tipe data presisi.
 */

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetsMaster = [
    {
      name: "USERS",
      headers: [
        "user_id", "nama_lengkap", "email", "role", "departemen", 
        "status_aktif", "password", "is_verified", "token", "created_at", "site"
      ],
      sampleData: [
        ["USR-0001", "Admin Data Center", "admin@datacenter.com", "Super Admin", "Operations", "Aktif", "admin123", "TRUE", "TOKEN-ADMIN", "2026-09-12 08:00:00", "CGK1"],
        ["USR-0002", "Budi PM Infrastructure", "pm.budi@datacenter.com", "Project Manager", "Operations", "Aktif", "pm123", "TRUE", "TOKEN-PM", "2026-09-12 08:30:00", "CGK1"],
        ["USR-0003", "Eko Tech Operator", "eko.tech@datacenter.com", "Karyawan", "Operations", "Aktif", "eko123", "TRUE", "TOKEN-EKO", "2026-09-12 09:00:00", "CGK3"],
        ["USR-0004", "Rina DC Technician", "rina.tech@datacenter.com", "Karyawan", "Network & Cabling", "Aktif", "dc123456", "TRUE", "TOKEN-RINA", "2026-09-12 09:15:00", "CGK3A"]
      ]
    },
    {
      name: "EMPLOYEES",
      headers: [
        "employee_id", "user_id", "nik", "nama_lengkap", "email", 
        "jabatan", "departemen", "no_hp", "status_karyawan", "created_at", "site"
      ],
      sampleData: [
        ["EMP-0001", "USR-0003", "31710123001", "Eko Tech Operator", "eko.tech@datacenter.com", "Junior DC Technician", "Operations", "081234567890", "Aktif", "2026-09-12 09:00:00", "CGK3"],
        ["EMP-0002", "USR-0004", "31710123002", "Rina DC Technician", "rina.tech@datacenter.com", "Cabling Specialist", "Network & Cabling", "081298765432", "Aktif", "2026-09-12 09:15:00", "CGK3A"]
      ]
    },
    {
      name: "COURSES",
      headers: [
        "course_id", "judul_materi", "kategori_dc", "deskripsi_singkat", 
        "pdf_link", "video_url", "created_by", "created_at"
      ],
      sampleData: [
        [
          "CRS-0001", 
          "Cooling System CRAH & Chiller SOP", 
          "Introduction", 
          "SOP dan modul pelatihan teknis operasional Data Center mengenai Cooling System CRAH & Chiller SOP", 
          "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf", 
          "https://drive.google.com/file/d/15YCCYywuOOPVhr_J_0DUdx4p-hA6uKF8/view?usp=sharing", 
          "Project Manager", 
          "2026-09-12 10:00:00"
        ],
        [
          "CRS-0002", 
          "Redundansi UPS N+1 Maintenance", 
          "Data Center Fundamental (Basic)", 
          "SOP dan modul pelatihan teknis operasional Data Center mengenai Redundansi UPS N+1 Maintenance", 
          "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf", 
          "https://drive.google.com/file/d/15YCCYywuOOPVhr_J_0DUdx4p-hA6uKF8/view?usp=sharing", 
          "Project Manager", 
          "2026-09-12 10:30:00"
        ]
      ]
    },
    {
      name: "QUIZZES",
      headers: [
        "quiz_id", "course_id", "pertanyaan", "opsi_a", "opsi_b", 
        "opsi_c", "opsi_d", "jawaban_benar", "passing_score"
      ],
      sampleData: [
        ["QUIZ-0001", "CRS-0001", "Berapa temperatur standar inlet CRAH untuk lingkungan Server Room menurut ASHRAE?", "10 - 15°C", "18 - 27°C", "28 - 35°C", "36 - 40°C", "B", 80],
        ["QUIZ-0002", "CRS-0001", "Apa langkah pertama prosedur pertukaran (switchover) Chiller saat alarm kegagalan kompresor berbunyi?", "Matikan seluruh listrik Data Center", "Aktivasi unit Chiller cadangan Redundant N+1", "Buka seluruh pintu server room", "Diamkan hingga suhu stabil secara alami", "B", 80]
      ]
    },
    {
      name: "PROGRESS",
      headers: [
        "progress_id", "user_id", "course_id", "periode", "is_pdf_downloaded", 
        "is_video_watched", "quiz_score", "is_completed", "last_updated", "pdf_score", 
        "video_score", "pages_read_count", "total_pages", "video_seconds_watched"
      ],
      sampleData: [
        ["PRG-0001", "USR-0003", "CRS-0001", "2026-09", "TRUE", "TRUE", 85, "Lulus", "2026-09-12 14:00:00", 10, 10, 2, 2, 750],
        ["PRG-0002", "USR-0004", "CRS-0001", "2026-09", "TRUE", "FALSE", 0, "Dalam Proses", "2026-09-12 15:00:00", 5, 0, 1, 2, 320]
      ]
    },
    {
      name: "MONTHLY_MISSIONS",
      headers: [
        "mission_id", "user_id", "periode", "target_materi", "materi_selesai", 
        "status", "assigned_course_ids", "last_updated"
      ],
      sampleData: [
        ["MSN-0001", "USR-0003", "2026-09", 2, 1, "In Progress", "CRS-0001,CRS-0002", "2026-09-12 14:00:00"],
        ["MSN-0002", "USR-0004", "2026-09", 2, 0, "In Progress", "CRS-0001,CRS-0002", "2026-09-12 15:00:00"]
      ]
    }
  ];

  sheetsMaster.forEach(function(master) {
    var sheet = ss.getSheetByName(master.name);
    if (!sheet) {
      sheet = ss.insertSheet(master.name);
    } else {
      sheet.clearContents();
    }

    // Set Header Row
    sheet.getRange(1, 1, 1, master.headers.length).setValues([master.headers]);
    
    // Format Header Style
    var headerRange = sheet.getRange(1, 1, 1, master.headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#1b2559");
    headerRange.setFontColor("#ffffff");

    // Insert Sample Data if available
    if (master.sampleData && master.sampleData.length > 0) {
      sheet.getRange(2, 1, master.sampleData.length, master.headers.length).setValues(master.sampleData);
    }

    // Set Specific Column Formats
    if (master.name === "PROGRESS") {
      sheet.getRange("G2:G").setNumberFormat("0");       // quiz_score
      sheet.getRange("J2:L").setNumberFormat("0");       // pdf_score, video_score, pages_read_count
      sheet.getRange("M2:N").setNumberFormat("0");       // total_pages, video_seconds_watched
      sheet.getRange("I2:I").setNumberFormat("yyyy-mm-dd hh:mm:ss"); // last_updated
    } else if (master.name === "COURSES" || master.name === "USERS" || master.name === "EMPLOYEES") {
      sheet.getRange("H2:H").setNumberFormat("yyyy-mm-dd hh:mm:ss");
      sheet.getRange("J2:J").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    }

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, master.headers.length);
  });

  SpreadsheetApp.flush();
  Logger.log("Inisialisasi Database e-LMS Zettbos Berhasil Diperbarui!");
}
