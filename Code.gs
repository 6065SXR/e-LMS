/**
 * e-LMS Data Center KSPS - Main Controller & Shared Helpers
 * System: e-LMS Data Center KSPS
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.verify) {
    var token = e.parameter.verify;
    var verifyResult = verifyEmailToken(token);
    
    var htmlContent = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Verifikasi Akun e-LMS KSPS</title>' +
      '<script src="https://cdn.tailwindcss.com"></script>' +
      '<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700&display=swap" rel="stylesheet">' +
      '<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">' +
      '</head><body class="min-h-screen bg-slate-900 text-white flex items-center justify-center p-4 font-sans">' +
      '<div class="bg-slate-800 border border-slate-700/80 p-8 rounded-3xl max-w-md w-full text-center shadow-2xl space-y-4">' +
      (verifyResult.success ? 
        '<div class="w-16 h-16 bg-emerald-500/20 text-emerald-400 rounded-full flex items-center justify-center mx-auto text-3xl"><i class="fa-solid fa-circle-check"></i></div>' +
        '<h2 class="text-2xl font-bold text-white">Email Terverifikasi!</h2>' +
        '<p class="text-slate-300 text-sm">' + verifyResult.message + '</p>' +
        '<a href="' + ScriptApp.getService().getUrl() + '" class="mt-4 inline-block w-full py-3 bg-[#1C1B8E] hover:bg-[#121163] text-white font-bold rounded-xl text-sm transition shadow-lg">Masuk ke Portal e-LMS KSPS</a>'
        :
        '<div class="w-16 h-16 bg-rose-500/20 text-rose-400 rounded-full flex items-center justify-center mx-auto text-3xl"><i class="fa-solid fa-circle-xmark"></i></div>' +
        '<h2 class="text-2xl font-bold text-white">Verifikasi Gagal</h2>' +
        '<p class="text-slate-300 text-sm">' + verifyResult.message + '</p>') +
      '</div></body></html>';
      
    return HtmlService.createHtmlOutput(htmlContent);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('e-LMS Data Center KSPS')
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

/**
 * Konversi PDF Google Drive / URL External menjadi Base64 Stream
 * untuk di-render Native oleh PDF.js di Canvas Browser
 */
function getPdfBase64(pdfUrl) {
  try {
    if (!pdfUrl || pdfUrl === '#' || pdfUrl.trim() === '') {
      return { success: false, error: "URL PDF kosong" };
    }

    var fileId = "";
    var match = pdfUrl.match(/\/d\/([a-zA-Z0-9_-]+)/) || pdfUrl.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      fileId = match[1];
    }

    var blob;
    if (fileId) {
      blob = DriveApp.getFileById(fileId).getBlob();
    } else {
      blob = UrlFetchApp.fetch(pdfUrl).getBlob();
    }

    var bytes = blob.getBytes();
    var base64 = Utilities.base64Encode(bytes);
    return { success: true, base64: base64 };
  } catch (err) {
    Logger.log("getPdfBase64 Error: " + err.toString());
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

/**
 * ============================================================================
 * BACKEND EXTENSION: UNDER CONSTRUCTION MANAGER & WHATSAPP INTEGRATION
 * Sheet Target: "UNDER_CONSTRUCTION"
 * ============================================================================
 */
function getUnderConstructionConfig() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("UNDER_CONSTRUCTION");

    var defaultConfig = {
      whatsapp_number: "081234567890",
      default_message: "Halo Admin KSPS, saya ingin menanyakan pembaruan fitur",
      menus: [
        { id: "dashboard", name: "Dashboard", lock_karyawan: false, lock_pm: false, note: "Dashboard sedang sinkronisasi data analitik." },
        { id: "katalog", name: "Katalog Materi", lock_karyawan: false, lock_pm: false, note: "Silabus materi baru sedang dipersiapkan oleh tim operasional." },
        { id: "karyawan", name: "Data Karyawan", lock_karyawan: false, lock_pm: false, note: "Sinkronisasi profil data personel sedang berlangsung." },
        { id: "misi", name: "Misi Bulanan", lock_karyawan: false, lock_pm: false, note: "Penyesuaian target kompetensi bulanan sedang dioptimalkan." },
        { id: "overtime", name: "Form Lembur", lock_karyawan: false, lock_pm: false, note: "Modul pengajuan lembur sedang dalam pemeliharaan berkala." },
        { id: "users", name: "Manajemen User", lock_karyawan: false, lock_pm: false, note: "Sistem autentikasi dan manajemen hak akses sedang dimutakhirkan." },
        { id: "backup", name: "Backup & Restore", lock_karyawan: false, lock_pm: false, note: "Modul backup database sedang dalam pemeliharaan berkala." }
      ]
    };

    // Buat otomatis sheet jika belum ada di spreadsheet
    if (!sheet) {
      sheet = ss.insertSheet("UNDER_CONSTRUCTION");
      sheet.appendRow(["menu_id", "menu_name", "lock_karyawan", "lock_pm", "custom_note", "whatsapp_number", "default_message"]);
      
      sheet.getRange(1, 1, 1, 7)
           .setFontWeight("bold")
           .setBackground("#1C1B8E")
           .setFontColor("#FFFFFF")
           .setHorizontalAlignment("center");
      
      defaultConfig.menus.forEach(function(m, idx) {
        var wa = (idx === 0) ? "'" + defaultConfig.whatsapp_number : "";
        var msg = (idx === 0) ? defaultConfig.default_message : "";
        sheet.appendRow([m.id, m.name, m.lock_karyawan, m.lock_pm, m.note, wa, msg]);
      });

      sheet.autoResizeColumns(1, 7);
      return { success: true, data: defaultConfig };
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, data: defaultConfig };
    }

    var config = {
      whatsapp_number: (data[1][5] !== undefined && data[1][5] !== "") ? String(data[1][5]) : defaultConfig.whatsapp_number,
      default_message: (data[1][6] !== undefined && data[1][6] !== "") ? String(data[1][6]) : defaultConfig.default_message,
      menus: []
    };

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;

      config.menus.push({
        id: String(row[0]).trim(),
        name: String(row[1] || row[0]).trim(),
        lock_karyawan: (row[2] === true || String(row[2]).toLowerCase() === "true"),
        lock_pm: (row[3] === true || String(row[3]).toLowerCase() === "true"),
        note: String(row[4] || "")
      });
    }

    return { success: true, data: config };
  } catch (err) {
    Logger.log("getUnderConstructionConfig Error: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function saveUnderConstructionConfig(config) {
  try {
    if (!config) {
      return { success: false, error: "Konfigurasi tidak valid." };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("UNDER_CONSTRUCTION");

    if (!sheet) {
      sheet = ss.insertSheet("UNDER_CONSTRUCTION");
    }

    sheet.clear();
    sheet.appendRow(["menu_id", "menu_name", "lock_karyawan", "lock_pm", "custom_note", "whatsapp_number", "default_message"]);
    
    sheet.getRange(1, 1, 1, 7)
         .setFontWeight("bold")
         .setBackground("#1C1B8E")
         .setFontColor("#FFFFFF")
         .setHorizontalAlignment("center");

    var menus = config.menus || [];
    var waNumber = config.whatsapp_number || "081234567890";
    var defaultMsg = config.default_message || "Halo Admin KSPS, saya ingin menanyakan pembaruan fitur";

    menus.forEach(function(m, idx) {
      var wa = (idx === 0) ? "'" + waNumber : "";
      var msg = (idx === 0) ? defaultMsg : "";
      
      sheet.appendRow([
        m.id,
        m.name,
        Boolean(m.lock_karyawan),
        Boolean(m.lock_pm),
        m.note || "",
        wa,
        msg
      ]);
    });

    sheet.autoResizeColumns(1, 7);
    return { success: true, message: "Konfigurasi Under Construction berhasil disimpan ke database!" };
  } catch (err) {
    Logger.log("saveUnderConstructionConfig Error: " + err.toString());
    return { success: false, error: err.toString() };
  }
}
