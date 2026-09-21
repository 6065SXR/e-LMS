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
