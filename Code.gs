/**
 * ZettBOT Apps Script Assistant - Backend Controller & API (Complete Version)
 * System: e-LMS Data Center Zettbos
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.verify) {
    var token = e.parameter.verify;
    var verifyResult = verifyEmailToken(token);
    
    var htmlContent = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Verifikasi Akun e-LMS</title>' +
      '<script src="https://cdn.tailwindcss.com"></script>' +
      '<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700&display=swap" rel="stylesheet">' +
      '<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">' +
      '</head><body class="min-h-screen bg-slate-900 text-white flex items-center justify-center p-4 font-sans">' +
      '<div class="bg-slate-800 border border-slate-700/80 p-8 rounded-3xl max-w-md w-full text-center shadow-2xl space-y-4">' +
      (verifyResult.success ? 
        '<div class="w-16 h-16 bg-emerald-500/20 text-emerald-400 rounded-full flex items-center justify-center mx-auto text-3xl"><i class="fa-solid fa-circle-check"></i></div>' +
        '<h2 class="text-2xl font-bold text-white">Email Terverifikasi!</h2>' +
        '<p class="text-slate-300 text-sm">' + verifyResult.message + '</p>' +
        '<a href="' + ScriptApp.getService().getUrl() + '" class="mt-4 inline-block w-full py-3 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-xl text-sm transition shadow-lg">Masuk ke Portal e-LMS</a>'
        :
        '<div class="w-16 h-16 bg-rose-500/20 text-rose-400 rounded-full flex items-center justify-center mx-auto text-3xl"><i class="fa-solid fa-circle-xmark"></i></div>' +
        '<h2 class="text-2xl font-bold text-white">Verifikasi Gagal</h2>' +
        '<p class="text-slate-300 text-sm">' + verifyResult.message + '</p>') +
      '</div></body></html>';
      
    return HtmlService.createHtmlOutput(htmlContent);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('e-LMS Data Center Zettbos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheetDisplayValues(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() === 0) return [];
  return sheet.getDataRange().getDisplayValues();
}

function generateSequentialId(sheetName, prefix, extraOffset) {
  var offset = extraOffset || 0;
  var data = getSheetDisplayValues(sheetName);
  if (data.length <= 1) return prefix + "-" + ("0000" + (1 + offset)).slice(-4);
  
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var currentId = data[i][0];
    if (currentId && currentId.indexOf(prefix + "-") === 0) {
      var numPart = parseInt(currentId.replace(prefix + "-", ""), 10);
      if (!isNaN(numPart) && numPart > maxNum) {
        maxNum = numPart;
      }
    }
  }
  var nextNum = maxNum + 1 + offset;
  var formatted = ("0000" + nextNum).slice(-4);
  return prefix + "-" + formatted;
}

function registerUser(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("USERS");
    var empSheet = ss.getSheetByName("EMPLOYEES");

    var users = getSheetDisplayValues("USERS");
    for (var i = 1; i < users.length; i++) {
      if (users[i][2] && users[i][2].toLowerCase().trim() === data.email.toLowerCase().trim()) {
        return { success: false, error: "Email sudah terdaftar! Silakan gunakan email lain atau login." };
      }
    }

    var nextUserId = generateSequentialId("USERS", "USR");
    var nextEmpId = generateSequentialId("EMPLOYEES", "EMP");
    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var token = Utilities.getUuid();
    var userSite = data.site || "CGK1";

    userSheet.appendRow([
      nextUserId, data.nama, data.email, "Karyawan", data.departemen || "Operations",
      "Pending", data.password, "FALSE", token, dateStr, userSite
    ]);

    empSheet.appendRow([
      nextEmpId, nextUserId, data.nik || ("NIK-" + Math.floor(100000 + Math.random() * 900000)),
      data.nama, data.email, "Karyawan Operasional", data.departemen || "Operations",
      data.no_hp || "-", "Pending Verification", dateStr, userSite
    ]);

    SpreadsheetApp.flush();

    var appUrl = ScriptApp.getService().getUrl();
    var verifyLink = appUrl + "?verify=" + token;

    var emailSubject = "[e-LMS Data Center Zettbos] Otentikasi & Verifikasi Akun Baru";
    var emailBody = "Halo " + data.nama + ",\n\n" +
      "Terima kasih telah mendaftar di e-LMS Data Center Zettbos.\n" +
      "Untuk mengaktifkan akun Anda (Site: " + userSite + "), silakan klik link otentikasi di bawah ini:\n\n" +
      verifyLink + "\n\nSalam,\nAdmin e-LMS Data Center Zettbos";

    GmailApp.sendEmail(data.email, emailSubject, emailBody);

    return { 
      success: true, 
      message: "Registrasi berhasil! Link otentikasi telah dikirim ke email " + data.email + "." 
    };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function verifyEmailToken(token) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("USERS");
    var empSheet = ss.getSheetByName("EMPLOYEES");

    var uRows = getSheetDisplayValues("USERS");
    var userRowIdx = -1, userId = "", userName = "";

    for (var i = 1; i < uRows.length; i++) {
      if (uRows[i][8] === token) {
        userRowIdx = i + 1;
        userId = uRows[i][0];
        userName = uRows[i][1];
        break;
      }
    }

    if (userRowIdx === -1) {
      return { success: false, message: "Token verifikasi tidak valid atau sudah kadaluarsa." };
    }

    userSheet.getRange(userRowIdx, 6).setValue("Aktif");
    userSheet.getRange(userRowIdx, 8).setValue("TRUE");

    var eRows = getSheetDisplayValues("EMPLOYEES");
    for (var j = 1; j < eRows.length; j++) {
      if (eRows[j][1] === userId) {
        empSheet.getRange(j + 1, 9).setValue("Aktif");
        break;
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Selamat " + userName + ", akun Anda berhasil terverifikasi!" };

  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function loginUser(email, password) {
  try {
    var users = getSheetDisplayValues("USERS");
    var inputEmail = String(email || "").toLowerCase().trim();
    var inputPass = String(password || "").trim();

    for (var i = 1; i < users.length; i++) {
      var uEmail = String(users[i][2] || "").toLowerCase().trim();
      var uPass = String(users[i][6] || "").trim();
      var isVerified = String(users[i][7] || "").trim().toUpperCase();

      if (uEmail === inputEmail) {
        if (uPass !== inputPass) {
          return { success: false, error: "Password yang Anda masukkan salah!" };
        }
        if (isVerified !== "TRUE") {
          return { success: false, error: "Akun Anda belum diverifikasi! Periksa email Anda." };
        }

        return {
          success: true,
          user: {
            user_id: users[i][0],
            name: users[i][1],
            email: users[i][2],
            role: users[i][3],
            departemen: users[i][4],
            site: users[i][10] || "CGK1"
          }
        };
      }
    }
    return { success: false, error: "Email belum terdaftar di sistem e-LMS!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getEmployeesList() {
  try {
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var progressRows = getSheetDisplayValues("PROGRESS");

    var list = [];
    for (var i = 1; i < empRows.length; i++) {
      var uid = empRows[i][1];
      var pdfCount = 0, quizPassedCount = 0;

      for (var p = 1; p < progressRows.length; p++) {
        if (progressRows[p][1] === uid) {
          if (progressRows[p][4] === "TRUE" || parseInt(progressRows[p][11], 10) > 0) pdfCount++;
          if (progressRows[p][7] === "Lulus" || parseFloat(progressRows[p][6]) >= 80) quizPassedCount++;
        }
      }

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
        materi_download: pdfCount,
        materi_soal: quizPassedCount
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

function uploadFileToDrive(base64Data, fileName, mimeType, targetFolderId) {
  try {
    var rawBase64 = base64Data.indexOf(",") !== -1 ? base64Data.split(",")[1] : base64Data;
    var bytes = Utilities.base64Decode(rawBase64);
    var blob = Utilities.newBlob(bytes, mimeType, fileName);
    
    var file;
    if (targetFolderId) {
      var folder = DriveApp.getFolderById(targetFolderId);
      file = folder.createFile(blob);
    } else {
      file = DriveApp.createFile(blob);
    }
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    Logger.log("Drive Upload Error: " + err.toString());
    return "#";
  }
}

function saveCourseWithFiles(courseData, pdfObject, rawVideoUrl, rawQuizText) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("COURSES");
    if (!sheet) return { success: false, error: "Sheet COURSES tidak ditemukan" };

    var pdfUrl = "#";
    if (pdfObject && pdfObject.base64) {
      pdfUrl = uploadFileToDrive(pdfObject.base64, pdfObject.name, pdfObject.type, null);
    }

    var videoUrl = (rawVideoUrl && String(rawVideoUrl).trim() !== "") ? String(rawVideoUrl).trim() : "#";
    var nextId = generateSequentialId("COURSES", "CRS");
    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var desc = "SOP dan modul pelatihan teknis operasional Data Center mengenai " + courseData.judul_materi;

    var row = [
      nextId,
      courseData.judul_materi,
      courseData.kategori_dc,
      desc,
      pdfUrl,
      videoUrl,
      courseData.created_by || "Project Manager",
      dateStr
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    if (rawQuizText && String(rawQuizText).trim() !== "") {
      parseAndSaveQuizText(nextId, rawQuizText);
    }

    return { success: true, message: "Materi baru & kuis berhasil disimpan!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getCourseDetailsForEdit(courseId) {
  try {
    var courseRows = getSheetDisplayValues("COURSES");
    var quizRows = getSheetDisplayValues("QUIZZES");

    var foundCourse = null;
    for (var i = 1; i < courseRows.length; i++) {
      if (courseRows[i][0] === courseId) {
        foundCourse = {
          course_id: courseRows[i][0],
          judul_materi: courseRows[i][1],
          kategori_dc: courseRows[i][2],
          deskripsi_singkat: courseRows[i][3],
          pdf_link: courseRows[i][4],
          video_url: courseRows[i][5],
          created_by: courseRows[i][6]
        };
        break;
      }
    }

    if (!foundCourse) return { success: false, error: "Materi tidak ditemukan" };

    var quizFormattedLines = [];
    var qNum = 1;
    for (var q = 1; q < quizRows.length; q++) {
      if (quizRows[q][1] === courseId) {
        var question = quizRows[q][2];
        var optA = quizRows[q][3];
        var optB = quizRows[q][4];
        var optC = quizRows[q][5];
        var optD = quizRows[q][6];
        var correct = quizRows[q][7];

        quizFormattedLines.push(qNum + ". " + question);
        quizFormattedLines.push("A. " + optA + (correct === "A" ? " *" : ""));
        quizFormattedLines.push("B. " + optB + (correct === "B" ? " *" : ""));
        quizFormattedLines.push("C. " + optC + (correct === "C" ? " *" : ""));
        quizFormattedLines.push("D. " + optD + (correct === "D" ? " *" : ""));
        quizFormattedLines.push("");
        qNum++;
      }
    }

    foundCourse.quiz_text = quizFormattedLines.join("\n");
    return { success: true, data: foundCourse };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function updateCourseWithFiles(courseId, courseData, pdfObject, rawVideoUrl, rawQuizText) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var courseSheet = ss.getSheetByName("COURSES");
    var quizSheet = ss.getSheetByName("QUIZZES");

    if (!courseSheet) return { success: false, error: "Sheet COURSES tidak ditemukan" };

    var courseRows = getSheetDisplayValues("COURSES");
    var targetRowIdx = -1;

    for (var i = 1; i < courseRows.length; i++) {
      if (courseRows[i][0] === courseId) {
        targetRowIdx = i + 1;
        break;
      }
    }

    if (targetRowIdx === -1) return { success: false, error: "Data materi tidak ditemukan!" };

    courseSheet.getRange(targetRowIdx, 2).setValue(courseData.judul_materi);
    courseSheet.getRange(targetRowIdx, 3).setValue(courseData.kategori_dc);

    if (pdfObject && pdfObject.base64) {
      var newPdfUrl = uploadFileToDrive(pdfObject.base64, pdfObject.name, pdfObject.type, null);
      courseSheet.getRange(targetRowIdx, 5).setValue(newPdfUrl);
    }

    if (rawVideoUrl && String(rawVideoUrl).trim() !== "") {
      courseSheet.getRange(targetRowIdx, 6).setValue(String(rawVideoUrl).trim());
    }

    SpreadsheetApp.flush();

    if (rawQuizText && String(rawQuizText).trim() !== "") {
      var quizRows = quizSheet.getLastRow() > 0 ? quizSheet.getDataRange().getDisplayValues() : [];
      for (var q = quizRows.length - 1; q >= 1; q--) {
        if (quizRows[q][1] === courseId) {
          quizSheet.deleteRow(q + 1);
        }
      }
      SpreadsheetApp.flush();
      parseAndSaveQuizText(courseId, rawQuizText);
    }

    return { success: true, message: "Materi & kuis berhasil diperbarui!" };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function deleteCourse(courseId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var courseSheet = ss.getSheetByName("COURSES");
    if (courseSheet) {
      var cData = courseSheet.getDataRange().getDisplayValues();
      for (var i = cData.length - 1; i >= 1; i--) {
        if (cData[i][0] === courseId) {
          var pdfUrl = cData[i][4];
          if (pdfUrl && pdfUrl.indexOf("drive.google.com") !== -1) {
            try {
              var fileIdMatch = pdfUrl.match(/\/d\/([a-zA-Z0-9_-]+)/) || pdfUrl.match(/id=([a-zA-Z0-9_-]+)/);
              if (fileIdMatch && fileIdMatch[1]) {
                DriveApp.getFileById(fileIdMatch[1]).setTrashed(true);
              }
            } catch (e) {
              Logger.log("Error deleting Drive file: " + e.toString());
            }
          }
          courseSheet.deleteRow(i + 1);
          break;
        }
      }
    }

    var quizSheet = ss.getSheetByName("QUIZZES");
    if (quizSheet) {
      var qData = quizSheet.getDataRange().getDisplayValues();
      for (var q = qData.length - 1; q >= 1; q--) {
        if (qData[q][1] === courseId) {
          quizSheet.deleteRow(q + 1);
        }
      }
    }

    var progressSheet = ss.getSheetByName("PROGRESS");
    if (progressSheet) {
      var pData = progressSheet.getDataRange().getDisplayValues();
      for (var p = pData.length - 1; p >= 1; p--) {
        if (pData[p][2] === courseId) {
          progressSheet.deleteRow(p + 1);
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Materi berhasil dihapus secara permanen!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function parseAndSaveQuizText(courseId, rawText) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("QUIZZES");
    if (!sheet) return;

    var lines = String(rawText || "").split("\n");
    var currentQuestion = null;
    var qIndex = 0;

    lines.forEach(function(line) {
      var trimmed = line.trim();
      if (!trimmed) return;

      if (/^\d+[\.\)]/.test(trimmed)) {
        if (currentQuestion && currentQuestion.q) {
          saveSingleQuizRow(sheet, courseId, currentQuestion, qIndex);
          qIndex++;
        }
        currentQuestion = {
          q: trimmed.replace(/^\d+[\.\)]\s*/, ""),
          a: "", b: "", c: "", d: "",
          correct: "A",
          score: "80"
        };
      } else if (/^[A-Da-d][\.\)]/.test(trimmed) && currentQuestion) {
        var optionChar = trimmed.substring(0, 1).toUpperCase();
        var optionText = trimmed.replace(/^[A-Da-d][\.\)]\s*/, "");
        var isCorrect = false;

        if (optionText.indexOf("*") !== -1) {
          isCorrect = true;
          optionText = optionText.replace(/\*/g, "").trim();
        } else {
          optionText = optionText.trim();
        }

        if (optionChar === "A") currentQuestion.a = optionText;
        if (optionChar === "B") currentQuestion.b = optionText;
        if (optionChar === "C") currentQuestion.c = optionText;
        if (optionChar === "D") currentQuestion.d = optionText;

        if (isCorrect) {
          currentQuestion.correct = optionChar;
        }
      }
    });

    if (currentQuestion && currentQuestion.q) {
      saveSingleQuizRow(sheet, courseId, currentQuestion, qIndex);
    }

    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log("Error parseAndSaveQuizText: " + err.toString());
  }
}

function saveSingleQuizRow(sheet, courseId, qObj, extraOffset) {
  var nextId = generateSequentialId("QUIZZES", "QUIZ", extraOffset);
  sheet.appendRow([
    nextId,
    courseId,
    qObj.q,
    qObj.a || "-",
    qObj.b || "-",
    qObj.c || "-",
    qObj.d || "-",
    qObj.correct || "A",
    qObj.score || "80"
  ]);
}

function recordUserProgress(progressData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PROGRESS");
    if (!sheet) return { success: false, error: "Sheet PROGRESS tidak ditemukan" };
    
    var rows = sheet.getLastRow() > 0 ? sheet.getDataRange().getDisplayValues() : [];
    var currentPeriod = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    var foundIndex = -1;
    var currentPdf = progressData.is_pdf_downloaded ? "TRUE" : "FALSE";
    var currentVideo = progressData.is_video_watched ? "TRUE" : "FALSE";
    var currentScore = (progressData.quiz_score !== undefined && progressData.quiz_score !== null) ? Number(progressData.quiz_score) : 0;
    var currentPdfScore = Number(progressData.pdf_score || 0);
    var currentVideoScore = Number(progressData.video_score || 0);
    var currentPagesRead = (progressData.pages_read_count !== undefined && progressData.pages_read_count !== null) ? Number(progressData.pages_read_count) : 0;
    var currentTotalPages = (progressData.total_pages !== undefined && progressData.total_pages !== null) ? Number(progressData.total_pages) : 1;
    var currentVideoSeconds = (progressData.video_seconds_watched !== undefined && progressData.video_seconds_watched !== null) ? Number(progressData.video_seconds_watched) : 0;

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1] === progressData.user_id && rows[i][2] === progressData.course_id && rows[i][3] === currentPeriod) {
        foundIndex = i + 1;
        if (rows[i][4] === "TRUE") currentPdf = "TRUE";
        if (rows[i][5] === "TRUE") currentVideo = "TRUE";
        if ((progressData.quiz_score === undefined || progressData.quiz_score === null || progressData.quiz_score === 0) && rows[i][6] && parseFloat(rows[i][6]) > 0) {
          currentScore = Number(rows[i][6]);
        }
        if (!progressData.pdf_score && rows[i][9]) currentPdfScore = Number(rows[i][9]);
        if (!progressData.video_score && rows[i][10]) currentVideoScore = Number(rows[i][10]);

        if ((progressData.pages_read_count === undefined || progressData.pages_read_count === null || progressData.pages_read_count === 0) && rows[i][11]) {
          currentPagesRead = Number(rows[i][11]);
        }
        if ((progressData.total_pages === undefined || progressData.total_pages === null || progressData.total_pages === 0) && rows[i][12]) {
          currentTotalPages = Number(rows[i][12]);
        }
        if ((progressData.video_seconds_watched === undefined || progressData.video_seconds_watched === null || progressData.video_seconds_watched === 0) && rows[i][13]) {
          currentVideoSeconds = Number(rows[i][13]);
        }
        break;
      }
    }

    var scoreNum = parseFloat(currentScore) || 0;
    var isCompleted = (scoreNum >= 80 || (currentPdf === "TRUE" && currentVideo === "TRUE")) ? "Lulus" : "Dalam Proses";

    if (foundIndex > -1) {
      sheet.getRange(foundIndex, 5).setValue(currentPdf);
      sheet.getRange(foundIndex, 6).setValue(currentVideo);
      sheet.getRange(foundIndex, 7).setValue(currentScore);
      sheet.getRange(foundIndex, 8).setValue(isCompleted);
      sheet.getRange(foundIndex, 9).setValue(dateStr);
      sheet.getRange(foundIndex, 10).setValue(currentPdfScore);
      sheet.getRange(foundIndex, 11).setValue(currentVideoScore);
      sheet.getRange(foundIndex, 12).setValue(currentPagesRead);
      sheet.getRange(foundIndex, 13).setValue(currentTotalPages);
      sheet.getRange(foundIndex, 14).setValue(currentVideoSeconds);
    } else {
      var nextId = generateSequentialId("PROGRESS", "PRG");
      sheet.appendRow([
        nextId, progressData.user_id, progressData.course_id, currentPeriod,
        currentPdf, currentVideo, currentScore, isCompleted, dateStr,
        currentPdfScore, currentVideoScore, currentPagesRead, currentTotalPages, currentVideoSeconds
      ]);
    }

    SpreadsheetApp.flush();
    syncMonthlyMission(progressData.user_id, currentPeriod);

    return { 
      success: true, 
      status: isCompleted, 
      is_pdf: currentPdf === "TRUE", 
      is_video: currentVideo === "TRUE",
      pages_read_count: currentPagesRead,
      total_pages: currentTotalPages,
      video_seconds_watched: currentVideoSeconds,
      message: "Progress & Analytics berhasil diperbarui!"
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function syncMonthlyMission(userId, periodeBulan) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var progressSheet = ss.getSheetByName("PROGRESS");
  var missionSheet = ss.getSheetByName("MONTHLY_MISSIONS");

  var pRows = (progressSheet && progressSheet.getLastRow() > 0) ? progressSheet.getDataRange().getDisplayValues() : [];
  var completedCount = 0;

  for (var i = 1; i < pRows.length; i++) {
    if (pRows[i][1] === userId && pRows[i][3] === periodeBulan && (pRows[i][7] === "Lulus" || parseFloat(pRows[i][6]) >= 80)) {
      completedCount++;
    }
  }

  var mRows = (missionSheet && missionSheet.getLastRow() > 0) ? missionSheet.getDataRange().getDisplayValues() : [];
  var missionRowIdx = -1, target = 2;

  for (var j = 1; j < mRows.length; j++) {
    if (mRows[j][1] === userId && mRows[j][2] === periodeBulan) {
      missionRowIdx = j + 1;
      target = parseInt(mRows[j][3], 10) || 2;
      break;
    }
  }

  var statusMisi = completedCount >= target ? "Completed" : "In Progress";
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

  if (missionRowIdx > -1) {
    missionSheet.getRange(missionRowIdx, 5).setValue(completedCount);
    missionSheet.getRange(missionRowIdx, 6).setValue(statusMisi);
    missionSheet.getRange(missionRowIdx, 8).setValue(dateStr);
  } else {
    var nextMsnId = generateSequentialId("MONTHLY_MISSIONS", "MSN");
    missionSheet.appendRow([nextMsnId, userId, periodeBulan, target, completedCount, statusMisi, "", dateStr]);
  }
  SpreadsheetApp.flush();
}

function getCoursesPaginated(page, limit, search, category, userId) {
  try {
    var pageNum = parseInt(page, 10) || 1;
    var limitNum = parseInt(limit, 10) || 10;
    var searchTerm = String(search || "").toLowerCase().trim();
    var catFilter = String(category || "all").trim();

    var courseRows = getSheetDisplayValues("COURSES");
    var progressRows = getSheetDisplayValues("PROGRESS");
    var empRows = getSheetDisplayValues("EMPLOYEES");

    var currentPeriod = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var totalEmployees = Math.max(1, empRows.length - 1);

    var filtered = [];
    for (var c = 1; c < courseRows.length; c++) {
      var cId = courseRows[c][0];
      var title = courseRows[c][1];
      var cat = courseRows[c][2];

      if (searchTerm && title.toLowerCase().indexOf(searchTerm) === -1) continue;
      if (catFilter !== "all" && cat !== catFilter) continue;

      var totalScoreSum = 0, completedCount = 0, totalQuizAttempts = 0;
      var userCompleted = false, userPdfRead = 0, userTotalPages = 1, userVidSec = 0, userStatus = "Belum dipelajari";

      for (var p = 1; p < progressRows.length; p++) {
        if (progressRows[p][2] === cId) {
          var score = parseFloat(progressRows[p][6]) || 0;
          if (score > 0) {
            totalScoreSum += score;
            totalQuizAttempts++;
          }
          if (progressRows[p][7] === "Lulus" || score >= 80) {
            completedCount++;
          }

          if (progressRows[p][1] === userId && progressRows[p][3] === currentPeriod) {
            userPdfRead = parseInt(progressRows[p][11], 10) || 0;
            userTotalPages = parseInt(progressRows[p][12], 10) || 1;
            userVidSec = parseInt(progressRows[p][13], 10) || 0;

            var isPdf = progressRows[p][4] === "TRUE";
            var isVid = progressRows[p][5] === "TRUE";

            if (progressRows[p][7] === "Lulus" || score >= 80) {
              userCompleted = true;
              userStatus = "Completed";
            } else if (isPdf || isVid || userPdfRead > 0 || userVidSec > 0) {
              userStatus = "In Progress";
            }
          }
        }
      }

      var avgScore = totalQuizAttempts > 0 ? Math.round(totalScoreSum / totalQuizAttempts) : 0;

      filtered.push({
        course_id: cId,
        judul_materi: title,
        kategori_dc: cat,
        deskripsi: courseRows[c][3],
        pdf_link: courseRows[c][4],
        video_url: courseRows[c][5],
        created_by: courseRows[c][6],
        peserta_selesai: completedCount,
        total_peserta: totalEmployees,
        avg_score: avgScore,
        is_user_completed: userCompleted,
        user_status: userStatus,
        user_pages_read: userPdfRead,
        user_total_pages: userTotalPages,
        user_video_seconds: userVidSec
      });
    }

    var totalItems = filtered.length;
    var totalPages = Math.ceil(totalItems / limitNum) || 1;
    var startIndex = (pageNum - 1) * limitNum;
    var paginated = filtered.slice(startIndex, startIndex + limitNum);

    return {
      success: true,
      currentPage: pageNum,
      totalPages: totalPages,
      totalItems: totalItems,
      data: paginated
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getCourseEmployeeDetails(courseId) {
  try {
    var userRows = getSheetDisplayValues("USERS");
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var progressRows = getSheetDisplayValues("PROGRESS");

    var empMap = {};
    for (var e = 1; e < empRows.length; e++) {
      var uId = empRows[e][1];
      empMap[uId] = {
        nik: empRows[e][2],
        nama_lengkap: empRows[e][3],
        departemen: empRows[e][6],
        site: empRows[e][10] || "CGK1"
      };
    }

    for (var u = 1; u < userRows.length; u++) {
      var uid = userRows[u][0];
      if (!empMap[uid]) {
        empMap[uid] = {
          nik: "NIK-" + uid,
          nama_lengkap: userRows[u][1],
          departemen: userRows[u][4] || "Operations",
          site: userRows[u][10] || "CGK1"
        };
      }
    }

    var list = [];
    for (var p = 1; p < progressRows.length; p++) {
      if (progressRows[p][2] === courseId) {
        var pUid = progressRows[p][1];
        var empInfo = empMap[pUid] || { nik: "-", nama_lengkap: pUid, departemen: "Operations", site: "CGK1" };

        list.push({
          user_id: pUid,
          nik: empInfo.nik,
          nama_lengkap: empInfo.nama_lengkap,
          departemen: empInfo.departemen,
          site: empInfo.site,
          pages_read_count: parseInt(progressRows[p][11], 10) || 0,
          total_pages: parseInt(progressRows[p][12], 10) || 1,
          video_seconds_watched: parseInt(progressRows[p][13], 10) || 0,
          quiz_score: parseFloat(progressRows[p][6]) || 0,
          is_completed: progressRows[p][7] || "Dalam Proses",
          last_updated: progressRows[p][8] || "-"
        });
      }
    }

    return { success: true, data: list };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getDashboardData(userId, userRole, periodeBulan) {
  try {
    var currentPeriod = periodeBulan || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var progressRows = getSheetDisplayValues("PROGRESS");
    var courseRows = getSheetDisplayValues("COURSES");
    var empRows = getSheetDisplayValues("EMPLOYEES");

    var completedCourses = 0, totalQuizSum = 0, quizCount = 0;
    var missionDetail = [], completedDetail = [], scoreDetail = [];

    for (var p = 1; p < progressRows.length; p++) {
      if (progressRows[p][1] === userId) {
        var cId = progressRows[p][2];
        var foundCourse = courseRows.find(function(r) { return r[0] === cId; });
        var title = foundCourse ? foundCourse[1] : cId;
        var cat = foundCourse ? foundCourse[2] : "General";
        var score = parseFloat(progressRows[p][6]) || 0;

        if (progressRows[p][7] === "Lulus" || score >= 80) {
          completedCourses++;
          completedDetail.push({ course_id: cId, title: title, category: cat, updated: progressRows[p][8], score: score });
        }

        if (score > 0) {
          totalQuizSum += score;
          quizCount++;
          scoreDetail.push({ course_id: cId, title: title, score: score, status: score >= 80 ? "Lulus" : "Selesai", updated: progressRows[p][8] });
        }

        if (progressRows[p][3] === currentPeriod) {
          missionDetail.push({ course_id: cId, title: title, category: cat, status: progressRows[p][7] === "Lulus" ? "Selesai" : "In Progress", score: score > 0 ? score : "-" });
        }
      }
    }

    var avgScore = quizCount > 0 ? Math.round(totalQuizSum / quizCount) : 0;

    var siteStats = {
      CGK1: { total: 0, completed: 0, scoreSum: 0, scoreCount: 0 },
      CGK2: { total: 0, completed: 0, scoreSum: 0, scoreCount: 0 },
      CGK3: { total: 0, completed: 0, scoreSum: 0, scoreCount: 0 },
      CGK3A: { total: 0, completed: 0, scoreSum: 0, scoreCount: 0 },
      CGK4: { total: 0, completed: 0, scoreSum: 0, scoreCount: 0 }
    };

    for (var e = 1; e < empRows.length; e++) {
      var site = empRows[e][10] || "CGK1";
      var eUid = empRows[e][1];
      if (siteStats[site]) {
        siteStats[site].total++;
        var empCompletedCount = 0;
        for (var p2 = 1; p2 < progressRows.length; p2++) {
          if (progressRows[p2][1] === eUid && progressRows[p2][3] === currentPeriod) {
            var sVal = parseFloat(progressRows[p2][6]) || 0;
            if (sVal > 0) {
              siteStats[site].scoreSum += sVal;
              siteStats[site].scoreCount++;
            }
            if (progressRows[p2][7] === "Lulus" || sVal >= 80) empCompletedCount++;
          }
        }
        if (empCompletedCount >= 2) siteStats[site].completed++;
      }
    }

    var siteChartLabels = ["Site CGK1", "Site CGK2", "Site CGK3", "Site CGK3A", "Site CGK4"];
    var siteAvgScores = [], siteCompliance = [], siteDetail = [];

    ["CGK1", "CGK2", "CGK3", "CGK3A", "CGK4"].forEach(function(sKey) {
      var st = siteStats[sKey];
      var avgS = st.scoreCount > 0 ? Math.round(st.scoreSum / st.scoreCount) : 80;
      var compRate = st.total > 0 ? Math.round((st.completed / st.total) * 100) : 100;

      siteAvgScores.push(avgS);
      siteCompliance.push(compRate);
      siteDetail.push({
        site: sKey,
        total_karyawan: st.total,
        misi_tuntas: st.completed,
        avg_score: avgS,
        compliance_rate: compRate + "%"
      });
    });

    return {
      success: true,
      data: {
        completedCourses: completedCourses,
        avgScore: avgScore,
        mission: { target: 2, completed: Math.min(2, completedCourses), status: completedCourses >= 2 ? "Completed" : "In Progress" },
        siteChart: { labels: siteChartLabels, avgScores: siteAvgScores, complianceRates: siteCompliance },
        details: {
          missionDetail: missionDetail,
          completedDetail: completedDetail,
          scoreDetail: scoreDetail,
          siteDetail: siteDetail
        }
      }
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getCompetencyHeatmapData(periodeBulan) {
  try {
    return {
      success: true,
      data: {
        sites: ["CGK1", "CGK2", "CGK3", "CGK3A", "CGK4"],
        pillars: [
          { id: "cooling", name: "Cooling System (Chiller/CRAH)" },
          { id: "power", name: "Power Distribution (UPS/Genset)" },
          { id: "safety", name: "Fire & Safety (K3)" },
          { id: "network", name: "Network & Cabling" }
        ],
        matrix: {
          CGK1: { cooling: { avg: 90, risk: "green" }, power: { avg: 88, risk: "green" }, safety: { avg: 92, risk: "green" }, network: { avg: 85, risk: "green" } },
          CGK2: { cooling: { avg: 72, risk: "yellow" }, power: { avg: 65, risk: "red" }, safety: { avg: 80, risk: "yellow" }, network: { avg: 78, risk: "yellow" } },
          CGK3: { cooling: { avg: 88, risk: "green" }, power: { avg: 90, risk: "green" }, safety: { avg: 86, risk: "green" }, network: { avg: 82, risk: "yellow" } },
          CGK3A: { cooling: { avg: 80, risk: "yellow" }, power: { avg: 75, risk: "yellow" }, safety: { avg: 88, risk: "green" }, network: { avg: 80, risk: "yellow" } },
          CGK4: { cooling: { avg: 62, risk: "red" }, power: { avg: 68, risk: "red" }, safety: { avg: 72, risk: "yellow" }, network: { avg: 60, risk: "red" } }
        }
      }
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function deployRemedialToSite(site, pillarCategory, periodeBulan) {
  try {
    var currentPeriod = periodeBulan || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userRows = getSheetDisplayValues("USERS");
    var courseRows = getSheetDisplayValues("COURSES");
    var missionSheet = ss.getSheetByName("MONTHLY_MISSIONS");
    var missionRows = getSheetDisplayValues("MONTHLY_MISSIONS");

    var remedialCourseId = "CRS-0001";
    for (var c = 1; c < courseRows.length; c++) {
      var cCat = courseRows[c][2] || "";
      if ((pillarCategory === "power" && cCat.indexOf("Fundamental") !== -1) ||
          (pillarCategory === "cooling" && cCat.indexOf("Introduction") !== -1) ||
          (pillarCategory === "safety" && cCat.indexOf("Elementary") !== -1) ||
          (pillarCategory === "network" && cCat.indexOf("Tools") !== -1)) {
        remedialCourseId = courseRows[c][0];
        break;
      }
    }

    var updatedCount = 0;
    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    for (var u = 1; u < userRows.length; u++) {
      var uId = userRows[u][0];
      var uSite = userRows[u][10] || "CGK1";
      var uRole = userRows[u][3];

      if (uSite === site && uRole === "Karyawan") {
        var foundIdx = -1;
        var existingAssigned = [];

        for (var m = 1; m < missionRows.length; m++) {
          if (missionRows[m][1] === uId && missionRows[m][2] === currentPeriod) {
            foundIdx = m + 1;
            existingAssigned = missionRows[m][6] ? missionRows[m][6].split(",") : [];
            break;
          }
        }

        if (existingAssigned.indexOf(remedialCourseId) === -1) {
          existingAssigned.push(remedialCourseId);
        }

        var newAssignedStr = existingAssigned.slice(0, 2).join(",");

        if (foundIdx > -1) {
          missionSheet.getRange(foundIdx, 7).setValue(newAssignedStr);
          missionSheet.getRange(foundIdx, 8).setValue(dateStr);
        } else {
          var nextId = generateSequentialId("MONTHLY_MISSIONS", "MSN", updatedCount);
          missionSheet.appendRow([
            nextId, uId, currentPeriod, 2, 0, "In Progress", newAssignedStr, dateStr
          ]);
        }
        updatedCount++;
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Modul Remedial berhasil dikirim ke " + updatedCount + " teknisi di Site " + site + "!" };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function dispatchCustomMission(dispatchData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var period = dispatchData.periode || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var targetSite = dispatchData.targetSite || "ALL";
    var targetDept = dispatchData.targetDept || "ALL";
    var assignedCourseIds = [dispatchData.course1_id, dispatchData.course2_id].filter(Boolean).join(",");

    if (!assignedCourseIds) {
      return { success: false, error: "Harap pilih minimal 1 materi untuk ditugaskan!" };
    }

    var userRows = getSheetDisplayValues("USERS");
    var missionSheet = ss.getSheetByName("MONTHLY_MISSIONS");
    var missionRows = getSheetDisplayValues("MONTHLY_MISSIONS");

    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var dispatchedCount = 0;

    for (var u = 1; u < userRows.length; u++) {
      var uId = userRows[u][0];
      var uRole = userRows[u][3];
      var uDept = userRows[u][4];
      var uSite = userRows[u][10] || "CGK1";

      if (uRole === "Karyawan") {
        var matchSite = (targetSite === "ALL" || uSite === targetSite);
        var matchDept = (targetDept === "ALL" || uDept === targetDept);

        if (matchSite && matchDept) {
          var foundIdx = -1;
          for (var m = 1; m < missionRows.length; m++) {
            if (missionRows[m][1] === uId && missionRows[m][2] === period) {
              foundIdx = m + 1;
              break;
            }
          }

          if (foundIdx > -1) {
            missionSheet.getRange(foundIdx, 7).setValue(assignedCourseIds);
            missionSheet.getRange(foundIdx, 8).setValue(dateStr);
          } else {
            var nextId = generateSequentialId("MONTHLY_MISSIONS", "MSN", dispatchedCount);
            missionSheet.appendRow([
              nextId, uId, period, 2, 0, "In Progress", assignedCourseIds, dateStr
            ]);
          }
          dispatchedCount++;
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Smart Mission berhasil dikirim ke " + dispatchedCount + " Karyawan/Teknisi!" };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getMonthlyMissionsData(userId, userRole, periodeBulan) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var currentPeriod = periodeBulan || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    
    var courseRows = getSheetDisplayValues("COURSES");
    var userRows = getSheetDisplayValues("USERS");
    var empRows = getSheetDisplayValues("EMPLOYEES");
    var progressRows = getSheetDisplayValues("PROGRESS");
    var missionSheet = ss.getSheetByName("MONTHLY_MISSIONS");
    var missionRows = getSheetDisplayValues("MONTHLY_MISSIONS");

    var availableCourses = [];
    for (var c = 1; c < courseRows.length; c++) {
      availableCourses.push({
        course_id: courseRows[c][0],
        judul_materi: courseRows[c][1],
        kategori_dc: courseRows[c][2],
        pdf_link: courseRows[c][4],
        video_url: courseRows[c][5]
      });
    }

    var dateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    var employeeList = [];
    for (var u = 1; u < userRows.length; u++) {
      var uRole = userRows[u][3];
      if (uRole !== "Super Admin" && uRole !== "Project Manager") {
        var empUid = userRows[u][0];
        var empName = userRows[u][1];
        var empSite = userRows[u][10] || "CGK1";
        var empDept = userRows[u][4] || "Operations";
        var empNik = "-";

        for (var e = 1; e < empRows.length; e++) {
          if (empRows[e][1] === empUid) {
            empNik = empRows[e][2];
            break;
          }
        }

        employeeList.push({ user_id: empUid, nama_lengkap: empName, nik: empNik, site: empSite, departemen: empDept });
      }
    }

    var newMissionsCount = 0;
    employeeList.forEach(function(emp) {
      var foundMission = false;
      for (var m = 1; m < missionRows.length; m++) {
        if (missionRows[m][1] === emp.user_id && missionRows[m][2] === currentPeriod) {
          foundMission = true;
          break;
        }
      }

      if (!foundMission && availableCourses.length > 0) {
        var shuffled = availableCourses.slice().sort(function() { return 0.5 - Math.random(); });
        var assigned = shuffled.slice(0, 2).map(function(item) { return item.course_id; });
        var assignedStr = assigned.join(",");

        var nextId = generateSequentialId("MONTHLY_MISSIONS", "MSN", newMissionsCount);
        newMissionsCount++;
        missionSheet.appendRow([nextId, emp.user_id, currentPeriod, 2, 0, "In Progress", assignedStr, dateStr]);
      }
    });

    SpreadsheetApp.flush();
    missionRows = getSheetDisplayValues("MONTHLY_MISSIONS");

    if (userRole === "Karyawan") {
      var myMissionRow = null;
      for (var m2 = 1; m2 < missionRows.length; m2++) {
        if (missionRows[m2][1] === userId && missionRows[m2][2] === currentPeriod) {
          myMissionRow = missionRows[m2];
          break;
        }
      }

      var assignedIds = (myMissionRow && myMissionRow[6]) ? myMissionRow[6].split(",") : [];
      var assignedCoursesDetails = [];
      var completedMissions = 0;

      assignedIds.forEach(function(cId) {
        var foundC = availableCourses.find(function(item) { return item.course_id === cId; });
        if (foundC) {
          var isPdf = false, isVid = false, qScore = 0, statusLulus = false;

          for (var p = 1; p < progressRows.length; p++) {
            if (progressRows[p][1] === userId && progressRows[p][2] === cId && progressRows[p][3] === currentPeriod) {
              if (progressRows[p][4] === "TRUE") isPdf = true;
              if (progressRows[p][5] === "TRUE") isVid = true;
              qScore = parseFloat(progressRows[p][6]) || 0;
              if (progressRows[p][7] === "Lulus" || qScore >= 80) statusLulus = true;
              break;
            }
          }

          if (statusLulus || qScore >= 80 || (isPdf && isVid)) completedMissions++;

          assignedCoursesDetails.push({
            course_id: foundC.course_id,
            judul_materi: foundC.judul_materi,
            kategori_dc: foundC.kategori_dc,
            pdf_link: foundC.pdf_link,
            video_url: foundC.video_url,
            is_pdf_downloaded: isPdf,
            is_video_watched: isVid,
            quiz_score: qScore,
            is_completed: statusLulus
          });
        }
      });

      return {
        success: true,
        userRole: "Karyawan",
        periode: currentPeriod,
        misiStatus: { target: 2, completed: completedMissions, status: completedMissions >= 2 ? "Completed" : "In Progress" },
        assignedCourses: assignedCoursesDetails
      };

    } else {
      var adminTable = [];
      var totalSelesai = 0, totalDalamProses = 0, totalBelum = 0;

      employeeList.forEach(function(emp) {
        var mRow = null;
        for (var m3 = 1; m3 < missionRows.length; m3++) {
          if (missionRows[m3][1] === emp.user_id && missionRows[m3][2] === currentPeriod) {
            mRow = missionRows[m3];
            break;
          }
        }

        var assignedStr = mRow ? mRow[6] : "";
        var assignedIds = assignedStr ? assignedStr.split(",") : [];
        var assignedTitles = [];
        var completedCount = 0;

        assignedIds.forEach(function(cId) {
          var foundC = availableCourses.find(function(item) { return item.course_id === cId; });
          if (foundC) {
            assignedTitles.push(foundC.judul_materi);
            for (var p = 1; p < progressRows.length; p++) {
              if (progressRows[p][1] === emp.user_id && progressRows[p][2] === cId && progressRows[p][3] === currentPeriod) {
                if (progressRows[p][7] === "Lulus" || parseFloat(progressRows[p][6]) >= 80) completedCount++;
                break;
              }
            }
          }
        });

        if (completedCount >= 2) totalSelesai++;
        else if (completedCount === 1) totalDalamProses++;
        else totalBelum++;

        adminTable.push({
          user_id: emp.user_id,
          nik: emp.nik,
          nama_lengkap: emp.nama_lengkap,
          site: emp.site,
          departemen: emp.departemen,
          materi_assigned: assignedTitles.join(" | ") || "Belum Direset",
          materi_selesai: completedCount,
          target_materi: 2,
          status: completedCount >= 2 ? "Completed" : (completedCount === 1 ? "In Progress" : "Belum Dimulai")
        });
      });

      return {
        success: true,
        userRole: userRole,
        periode: currentPeriod,
        stats: { totalKaryawan: employeeList.length, selesai: totalSelesai, dalamProses: totalDalamProses, belumStart: totalBelum },
        tableData: adminTable
      };
    }

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getUsersList() {
  try {
    var rows = getSheetDisplayValues("USERS");
    var users = [];
    for (var i = 1; i < rows.length; i++) {
      users.push({
        user_id: rows[i][0],
        nama_lengkap: rows[i][1],
        email: rows[i][2],
        role: rows[i][3],
        departemen: rows[i][4],
        status_aktif: rows[i][5],
        is_verified: rows[i][7],
        created_at: rows[i][9],
        site: rows[i][10] || "CGK1"
      });
    }
    return { success: true, data: users };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getQuizForCourse(courseId) {
  try {
    var rows = getSheetDisplayValues("QUIZZES");
    var quizzes = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1] === courseId) {
        quizzes.push({
          quiz_id: rows[i][0],
          course_id: rows[i][1],
          pertanyaan: rows[i][2],
          opsi_a: rows[i][3],
          opsi_b: rows[i][4],
          opsi_c: rows[i][5],
          opsi_d: rows[i][6],
          jawaban_benar: rows[i][7],
          passing_score: rows[i][8]
        });
      }
    }
    return { success: true, data: quizzes };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
