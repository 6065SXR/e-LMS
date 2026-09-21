/**
 * Dashboard, Analytics & Risk Command Engine (Realtime Calculation)
 * System: e-LMS Data Center KSPS
 */

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

    var siteKeys = ["CGK1", "CGK2", "CGK3", "CGK3A", "CGK4"];
    var siteStats = {};
    siteKeys.forEach(function(s) {
      siteStats[s] = { total: 0, completed: 0, scoreSum: 0, scoreCount: 0 };
    });

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

    var siteChartLabels = siteKeys.map(function(s) { return "Site " + s; });
    var siteAvgScores = [], siteCompliance = [], siteDetail = [];

    siteKeys.forEach(function(sKey) {
      var st = siteStats[sKey];
      var avgS = st.scoreCount > 0 ? Math.round(st.scoreSum / st.scoreCount) : 0;
      var compRate = st.total > 0 ? Math.round((st.completed / st.total) * 100) : 0;

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
    var progressRows = getSheetDisplayValues("PROGRESS");
    var courseRows = getSheetDisplayValues("COURSES");
    var empRows = getSheetDisplayValues("EMPLOYEES");

    var sites = ["CGK1", "CGK2", "CGK3", "CGK3A", "CGK4"];

    // Mengelompokkan langsung ke 4 Group / Kategori materi asli
    function getPillarForCourse(cat) {
      var category = String(cat || "").trim();
      if (category === "Introduction") return "cooling"; // Col 1: Introduction
      if (category === "Elementary") return "power"; // Col 2: Elementary
      if (category === "Data Center Fundamental (Basic)") return "safety"; // Col 3: DC Fundamental
      if (category === "Tools") return "network"; // Col 4: Tools
      return "cooling";
    }

    var userSiteMap = {};
    for (var e = 1; e < empRows.length; e++) {
      userSiteMap[empRows[e][1]] = empRows[e][10] || "CGK1";
    }

    var coursePillarMap = {};
    for (var c = 1; c < courseRows.length; c++) {
      coursePillarMap[courseRows[c][0]] = getPillarForCourse(courseRows[c][2]);
    }

    var matrixScores = {};
    sites.forEach(function(s) {
      matrixScores[s] = {
        cooling: { sum: 0, count: 0 },
        power: { sum: 0, count: 0 },
        safety: { sum: 0, count: 0 },
        network: { sum: 0, count: 0 }
      };
    });

    for (var p = 1; p < progressRows.length; p++) {
      var uId = progressRows[p][1];
      var cId = progressRows[p][2];
      var score = parseFloat(progressRows[p][6]) || 0;

      var site = userSiteMap[uId] || "CGK1";
      var pillar = coursePillarMap[cId] || "cooling";

      if (matrixScores[site] && matrixScores[site][pillar] && score > 0) {
        matrixScores[site][pillar].sum += score;
        matrixScores[site][pillar].count++;
      }
    }

    var matrix = {};
    sites.forEach(function(s) {
      matrix[s] = {};
      ["cooling", "power", "safety", "network"].forEach(function(p) {
        var cell = matrixScores[s][p];
        var avg = cell.count > 0 ? Math.round(cell.sum / cell.count) : 0;
        var risk = "red";
        if (avg >= 85) risk = "green";
        else if (avg >= 70) risk = "yellow";
        else risk = "red";

        matrix[s][p] = { avg: avg, risk: risk };
      });
    });

    return {
      success: true,
      data: {
        sites: sites,
        pillars: [
          { id: "cooling", name: "Introduction" },
          { id: "power", name: "Elementary" },
          { id: "safety", name: "Data Center Fundamental (Basic)" },
          { id: "network", name: "Tools" }
        ],
        matrix: matrix
      }
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
