/**
 * Course Catalog & Media Attachment Controller (Transkrip 3-Pilar)
 * System: e-LMS Data Center KSPS
 */

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
      var userCompleted = false, userPdfScore = 0, userVideoScore = 0, userQuizScore = 0, userComposite = 0, userStatus = "Belum dipelajari";

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
            userPdfScore = parseInt(progressRows[p][9], 10) || (progressRows[p][4] === "TRUE" ? 100 : 0);
            userVideoScore = parseInt(progressRows[p][10], 10) || (progressRows[p][5] === "TRUE" ? 100 : 0);
            userQuizScore = score;
            userComposite = Math.round((0.2 * userPdfScore) + (0.3 * userVideoScore) + (0.5 * userQuizScore));

            if (progressRows[p][7] === "Lulus" || userComposite >= 80) {
              userCompleted = true;
              userStatus = "Completed";
            } else if (userPdfScore > 0 || userVideoScore > 0) {
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
        user_pdf_score: userPdfScore,
        user_video_score: userVideoScore,
        user_quiz_score: userQuizScore,
        user_composite_score: userComposite
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

        var pdfScore = parseInt(progressRows[p][9], 10) || (progressRows[p][4] === "TRUE" ? 100 : 0);
        var videoScore = parseInt(progressRows[p][10], 10) || (progressRows[p][5] === "TRUE" ? 100 : 0);
        var quizScore = parseFloat(progressRows[p][6]) || 0;
        var composite = Math.round((0.2 * pdfScore) + (0.3 * videoScore) + (0.5 * quizScore));

        list.push({
          user_id: pUid,
          nik: empInfo.nik,
          nama_lengkap: empInfo.nama_lengkap,
          departemen: empInfo.departemen,
          site: empInfo.site,
          pdf_score: pdfScore,
          video_score: videoScore,
          quiz_score: quizScore,
          composite_score: composite,
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
