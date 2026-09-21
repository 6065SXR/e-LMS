/**
 * Authentication & Verification Controller
 * System: e-LMS Data Center KSPS
 */

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

    var emailSubject = "[e-LMS Data Center KSPS] Otentikasi & Verifikasi Akun Baru";
    var emailBody = "Halo " + data.nama + ",\n\n" +
      "Terima kasih telah mendaftar di e-LMS Data Center KSPS.\n" +
      "Untuk mengaktifkan akun Anda (Site: " + userSite + "), silakan klik link otentikasi di bawah ini:\n\n" +
      verifyLink + "\n\nSalam,\nAdmin e-LMS Data Center KSPS";

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
    return { success: false, error: "Email belum terdaftar di sistem e-LMS KSPS!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
