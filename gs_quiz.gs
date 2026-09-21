/**
 * Quiz Questions & User Progress Recording Engine (Anti-Cheating & 3-Pilar Composite Scoring)
 * System: e-LMS Data Center KSPS
 * 
 * Composite Score Formula: (20% * PDF Score) + (30% * Video Score) + (50% * Quiz Score)
 */

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
    
    var currentPagesRead = (progressData.pages_read_count !== undefined && progressData.pages_read_count !== null) ? Number(progressData.pages_read_count) : 0;
    var currentTotalPages = (progressData.total_pages !== undefined && progressData.total_pages !== null) ? Number(progressData.total_pages) : 1;
    var currentVideoSeconds = (progressData.video_seconds_watched !== undefined && progressData.video_seconds_watched !== null) ? Number(progressData.video_seconds_watched) : 0;

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1] === progressData.user_id && rows[i][2] === progressData.course_id && rows[i][3] === currentPeriod) {
        foundIndex = i + 1;

        if (rows[i][4] === "TRUE") currentPdf = "TRUE";
        if (rows[i][5] === "TRUE") currentVideo = "TRUE";

        if ((progressData.quiz_score === undefined || progressData.quiz_score === null) && rows[i][6]) {
          currentScore = Number(rows[i][6]);
        } else if (progressData.quiz_score !== undefined && progressData.quiz_score !== null) {
          currentScore = Math.max(Number(rows[i][6] || 0), currentScore);
        }

        if ((progressData.pages_read_count === undefined || progressData.pages_read_count === null || progressData.pages_read_count === 0) && rows[i][11]) {
          currentPagesRead = Number(rows[i][11]);
        } else {
          currentPagesRead = Math.max(Number(rows[i][11] || 0), currentPagesRead);
        }

        if ((progressData.total_pages === undefined || progressData.total_pages === null || progressData.total_pages === 0) && rows[i][12]) {
          currentTotalPages = Number(rows[i][12]);
        }

        if ((progressData.video_seconds_watched === undefined || progressData.video_seconds_watched === null || progressData.video_seconds_watched === 0) && rows[i][13]) {
          currentVideoSeconds = Number(rows[i][13]);
        } else {
          currentVideoSeconds = Math.max(Number(rows[i][13] || 0), currentVideoSeconds);
        }
        break;
      }
    }

    if (currentPdf === "TRUE" && currentPagesRead === 0) {
      currentPagesRead = currentTotalPages;
    }
    if (currentPagesRead >= currentTotalPages && currentTotalPages > 0) {
      currentPdf = "TRUE";
    }

    if (currentVideo === "TRUE" && currentVideoSeconds === 0) {
      currentVideoSeconds = 300;
    }
    if (currentVideoSeconds >= 300) {
      currentVideo = "TRUE";
    }

    // Kalkulasi Skor 3-Pilar Presisi
    var pdfScore = currentTotalPages > 0 ? Math.min(100, Math.round((currentPagesRead / currentTotalPages) * 100)) : (currentPdf === "TRUE" ? 100 : 0);
    var videoScore = currentVideo === "TRUE" ? 100 : (currentVideoSeconds > 0 ? Math.min(100, Math.round((currentVideoSeconds / 300) * 100)) : 0);
    var quizScore = parseFloat(currentScore) || 0;

    // Nilai Komposit Akhir (20% PDF + 30% Video + 50% Kuis)
    var compositeScore = Math.round((0.20 * pdfScore) + (0.30 * videoScore) + (0.50 * quizScore));
    var isCompleted = (quizScore >= 80 || compositeScore >= 80) ? "Lulus" : "Dalam Proses";

    if (foundIndex > -1) {
      sheet.getRange(foundIndex, 5).setValue(currentPdf);
      sheet.getRange(foundIndex, 6).setValue(currentVideo);
      sheet.getRange(foundIndex, 7).setValue(quizScore);
      sheet.getRange(foundIndex, 8).setValue(isCompleted);
      sheet.getRange(foundIndex, 9).setValue(dateStr);
      sheet.getRange(foundIndex, 10).setValue(pdfScore);
      sheet.getRange(foundIndex, 11).setValue(videoScore);
      sheet.getRange(foundIndex, 12).setValue(currentPagesRead);
      sheet.getRange(foundIndex, 13).setValue(currentTotalPages);
      sheet.getRange(foundIndex, 14).setValue(currentVideoSeconds);
    } else {
      var nextId = generateSequentialId("PROGRESS", "PRG");
      sheet.appendRow([
        nextId, progressData.user_id, progressData.course_id, currentPeriod,
        currentPdf, currentVideo, quizScore, isCompleted, dateStr,
        pdfScore, videoScore, currentPagesRead, currentTotalPages, currentVideoSeconds
      ]);
    }

    SpreadsheetApp.flush();
    syncMonthlyMission(progressData.user_id, currentPeriod);

    return { 
      success: true, 
      status: isCompleted, 
      is_pdf: currentPdf === "TRUE", 
      is_video: currentVideo === "TRUE",
      pdf_score: pdfScore,
      video_score: videoScore,
      quiz_score: quizScore,
      composite_score: compositeScore,
      pages_read_count: currentPagesRead,
      total_pages: currentTotalPages,
      video_seconds_watched: currentVideoSeconds,
      message: "Progress & Transkrip 3-Pilar berhasil diperbarui!"
    };
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
