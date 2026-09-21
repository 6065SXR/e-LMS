/**
 * Monthly Missions & Smart Mission Dispatcher Engine (Dual-Group Support)
 * System: e-LMS Data Center KSPS
 */

function dispatchCustomMission(dispatchData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var period = dispatchData.periode || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    var targetSite = dispatchData.targetSite || "ALL";
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
      var uSite = userRows[u][10] || "CGK1";

      if (uRole === "Karyawan") {
        var matchSite = (targetSite === "ALL" || uSite === targetSite);

        if (matchSite) {
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
          var pdfScore = 0, videoScore = 0, compositeScore = 0;

          for (var p = 1; p < progressRows.length; p++) {
            if (progressRows[p][1] === userId && progressRows[p][2] === cId && progressRows[p][3] === currentPeriod) {
              if (progressRows[p][4] === "TRUE") isPdf = true;
              if (progressRows[p][5] === "TRUE") isVid = true;
              qScore = parseFloat(progressRows[p][6]) || 0;
              pdfScore = parseInt(progressRows[p][9], 10) || (isPdf ? 100 : 0);
              videoScore = parseInt(progressRows[p][10], 10) || (isVid ? 100 : 0);
              compositeScore = Math.round((0.2 * pdfScore) + (0.3 * videoScore) + (0.5 * qScore));

              if (progressRows[p][7] === "Lulus" || qScore >= 80 || compositeScore >= 80) statusLulus = true;
              break;
            }
          }

          if (statusLulus || qScore >= 80 || compositeScore >= 80) completedMissions++;

          assignedCoursesDetails.push({
            course_id: foundC.course_id,
            judul_materi: foundC.judul_materi,
            kategori_dc: foundC.kategori_dc,
            pdf_link: foundC.pdf_link,
            video_url: foundC.video_url,
            is_pdf_downloaded: isPdf,
            is_video_watched: isVid,
            pdf_score: pdfScore,
            video_score: videoScore,
            quiz_score: qScore,
            composite_score: compositeScore,
            is_completed: statusLulus
          });
        }
      });

      return {
        success: true,
        userRole: "Karyawan",
        periode: currentPeriod,
        availableCourses: availableCourses,
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
                var qVal = parseFloat(progressRows[p][6]) || 0;
                var pdfVal = parseInt(progressRows[p][9], 10) || (progressRows[p][4] === "TRUE" ? 100 : 0);
                var vidVal = parseInt(progressRows[p][10], 10) || (progressRows[p][5] === "TRUE" ? 100 : 0);
                var compVal = Math.round((0.2 * pdfVal) + (0.3 * vidVal) + (0.5 * qVal));

                if (progressRows[p][7] === "Lulus" || qVal >= 80 || compVal >= 80) completedCount++;
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
        availableCourses: availableCourses,
        stats: { totalKaryawan: employeeList.length, selesai: totalSelesai, dalamProses: totalDalamProses, belumStart: totalBelum },
        tableData: adminTable
      };
    }

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
