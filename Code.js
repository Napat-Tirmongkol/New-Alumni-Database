// =================================================================
// ส่วนที่ 1: การตั้งค่าและตัวแปรหลัก (CONFIGURATION)
// =================================================================
const SHEET_ID = "1yXsIRJB56XDGPANqMTD1tbAgPPZhnRVj5rwCg1F3Ew8";
const PROFILE_PICTURE_FOLDER_ID = "12VC1rdV1sRXxFP0M-QnZgytokZHwwx26";

const COLS_DB = {
  USER_ID: 1, EMAIL: 2, PASSWORD: 3, ROLE: 4, FIRST_NAME: 5, LAST_NAME: 6,
  REG_DATE: 7, LAST_LOGIN: 8, IS_ACTIVE: 9, RESET_TOKEN: 10, TOKEN_EXPIRY: 11
};
const COLS_PROFILE = {
  USER_ID: 1, EMAIL: 2, PICTURE_ID: 3, PREFIX: 4, STUDENT_ID: 5, FNAME_TH: 6, LNAME_TH: 7,
  FNAME_EN: 8, LNAME_EN: 9, GRAD_CLASS: 10, ADVISOR: 11, DOB: 12, GENDER: 13,
  BIRTH_COUNTRY: 14, NATIONALITY: 15, RACE: 16, PHONE: 17, GPAX: 18, ADDRESS: 19,
  EMERGENCY_NAME: 20, EMERGENCY_RELATION: 21, EMERGENCY_PHONE: 22, AWARDS: 23,
  FUTURE_PLAN: 24, EDUCATION_PLAN: 25, INTERNATIONAL_PLAN: 26, THAI_LICENSE: 27
};

// =================================================================
// ส่วนที่ 2: ฟังก์ชัน Setup (สำหรับสร้าง/ซ่อมแซมชีต)
// =================================================================
function setupProjectSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const userDbHeaders = ['UserID', 'Email', 'Password', 'Role', 'FirstNameTH', 'LastNameTH', 'RegistrationDate', 'LastLogin', 'IsActive', 'ResetToken', 'TokenExpiry'];
    const userProfileHeaders = ['UserID','Email', 'ProfilePictureID', 'Prefix', 'StudentID', 'FirstNameTH', 'LastNameTH', 'FirstNameEN', 'LastNameEN', 'GraduationClass', 'Advisor', 'DateOfBirth', 'Gender', 'BirthCountry', 'Nationality', 'Race', 'PhoneNumber', 'GPAX', 'CurrentAddress', 'EmergencyContactName', 'EmergencyContactRelation', 'EmergencyContactPhone', 'Awards', 'FutureWorkPlan', 'EducationPlan', 'InternationalWorkPlan', 'WillTakeThaiLicense'];
    const licenseExamHeaders = ['UserID', 'Email', 'ExamRound', 'ExamSession', 'ExamYear', 'Subject1_Maternity', 'Subject2_Pediatric', 'Subject3_Adult', 'Subject4_Geriatric', 'Subject5_Psychiatric', 'Subject6_Community', 'Subject7_Law', 'Subject8_Surgical', 'EvidenceFileID'];
    
    checkAndCreateSheet(spreadsheet, 'User_Database', userDbHeaders);
    checkAndCreateSheet(spreadsheet, 'User_Profiles', userProfileHeaders);
    checkAndCreateSheet(spreadsheet, 'License_Exam_Records', licenseExamHeaders);
    Browser.msgBox("ตั้งค่าชีตสำเร็จ!");
  } catch (e) {
    Browser.msgBox("เกิดข้อผิดพลาด: " + e.message);
  }
}

function checkAndCreateSheet(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

// =================================================================
// ส่วนที่ 3: ฟังก์ชันหลักของเว็บแอป (CORE WEB APP FUNCTIONS)
// =================================================================
function doGet(e) {
  const page = e ? e.parameter.page : null;
  const token = e ? e.parameter.token : null;
  if (page === 'resetpassword' && token) {
    const verification = verifyResetToken(token);
    if (verification.isValid) {
      const template = HtmlService.createTemplateFromFile('ResetPassword');
      template.token = token;
      return template.evaluate().setTitle('ตั้งรหัสผ่านใหม่');
    } else {
      return HtmlService.createTemplateFromFile('InvalidToken').evaluate().setTitle('ลิงก์ไม่ถูกต้อง');
    }
  }
  const pageMappings = {'forgotpassword':'ForgotPassword','admindashboard':'AdminDashboard','dashboard':'Dashboard','register':'Register'};
  if (page && pageMappings[page]) {
    return HtmlService.createTemplateFromFile(pageMappings[page]).evaluate().setTitle('ระบบติดตามศิษย์');
  }
  return HtmlService.createTemplateFromFile('Login').evaluate().setTitle('ระบบติดตามศิษย์');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPageContent(pageName) {
  const pageMap = {
    // หน้าสำหรับผู้ใช้ทั่วไป
    'profile': 'Profile',
    'profileSummary': 'ProfileSummary',
    'license': 'LicenseTracking',
    'donations': 'Donations',
    'virtual_id': 'VirtualID',

    // หน้าสำหรับ Admin
    'manageUsers': 'AdminUserManagement', 
    'licenseForAdmin': 'LicenseTracking' 
  };

  if (pageMap[pageName]) {
    return include(pageMap[pageName]);
  }

  return '<h1>ไม่พบหน้าเว็บ</h1>';
}

function getWebAppUrl() {
 return ScriptApp.getService().getUrl(); 
}

// =================================================================
// ส่วนที่ 4: ฟังก์ชันเกี่ยวกับสมาชิกและการยืนยันตัวตน (AUTHENTICATION)
// =================================================================
function validateLogin(credentials) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    if (!sheet) { throw new Error("ไม่พบชีต 'User_Database'"); }
    const data = sheet.getRange(2, 1, sheet.getLastRow(), COLS_DB.IS_ACTIVE).getValues();
    for (let i = 0; i < data.length; i++) {
      const rowData = data[i];
      const email = rowData[COLS_DB.EMAIL - 1];
      const password = String(rowData[COLS_DB.PASSWORD - 1]);
      const isActive = rowData[COLS_DB.IS_ACTIVE - 1];
      if (email === credentials.email && password === credentials.password) {
        if (isActive === false) { return { success: false, message: 'บัญชีของคุณถูกระงับการใช้งาน' }; }
        const userRow = i + 2;
        sheet.getRange(userRow, COLS_DB.LAST_LOGIN).setValue(new Date());
        const user = {firstName:rowData[COLS_DB.FIRST_NAME-1],lastName:rowData[COLS_DB.LAST_NAME-1],role:rowData[COLS_DB.ROLE-1]};
        return { success: true, user: user }; 
      }
    }
    return { success: false, message: 'อีเมลหรือรหัสผ่านไม่ถูกต้อง' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.message };
  }
}

function registerUser(userData) {
  try {
    const dbSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    if (!dbSheet) { throw new Error("ไม่พบชีต 'User_Database'"); }
    const emails = dbSheet.getRange(2, COLS_DB.EMAIL, dbSheet.getLastRow(), 1).getValues().flat();
    if (emails.includes(userData.email)) { return { success: false, message: 'อีเมลนี้ถูกใช้งานแล้ว' }; }
    const newUserId = "UID-" + new Date().getTime();
    const newUserDbRow = [newUserId,userData.email,userData.password,userData.userType,userData.firstName,userData.lastName,new Date(),null,true,null,null];
    dbSheet.appendRow(newUserDbRow);
    const profileSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (profileSheet) {
      const newUserProfileRow = [newUserId,userData.email,'','',userData.firstName,userData.lastName];
      profileSheet.appendRow(newUserProfileRow);
    }
    const licenseSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('License_Exam_Records');
    if (licenseSheet) {
      const newLicenseRow = [newUserId, userData.email];
      licenseSheet.appendRow(newLicenseRow);
    }
    return { success: true, message: 'สมัครสมาชิกสำเร็จ!' };
  } catch(error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.message };
  }
}

function processForgotPassword(email) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    if (!sheet) { throw new Error("ไม่พบชีต 'User_Database'"); }
    const data = sheet.getRange(2, 1, sheet.getLastRow(), COLS_DB.EMAIL).getValues();
    let userRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][COLS_DB.EMAIL - 1] === email) {
        userRow = i + 2;
        break;
      }
    }
    if (userRow === -1) { return { success: false, message: 'ไม่พบอีเมลนี้ในระบบ' }; }
    const token = Utilities.getUuid();
    const expiryDate = new Date(new Date().getTime() + 15 * 60 * 1000);
    sheet.getRange(userRow, COLS_DB.RESET_TOKEN).setValue(token);
    sheet.getRange(userRow, COLS_DB.TOKEN_EXPIRY).setValue(expiryDate);
    const resetLink = getWebAppUrl() + '?page=resetpassword&token=' + token;
    // MailApp.sendEmail(email, subject, body); // ต้องใส่ subject, body
    return { success: true, message: 'ส่งลิงก์สำหรับรีเซ็ตรหัสผ่านไปที่อีเมลของคุณแล้ว' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.message };
  }
}

function verifyResetToken(token) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    if (!sheet) return { isValid: false };
    const data = sheet.getRange(2, COLS_DB.RESET_TOKEN, sheet.getLastRow(), 2).getValues();
    const now = new Date();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === token && new Date(data[i][1]) > now) {
        return { isValid: true, row: i + 2 };
      }
    }
    return { isValid: false };
  } catch (e) {
    return { isValid: false };
  }
}

function updatePassword(token, newPassword) {
  const verification = verifyResetToken(token);
  if (!verification.isValid) { return { success: false, message: "Token ไม่ถูกต้องหรือหมดอายุแล้ว" }; }
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    sheet.getRange(verification.row, COLS_DB.PASSWORD).setValue(newPassword);
    sheet.getRange(verification.row, COLS_DB.RESET_TOKEN).setValue(null);
    sheet.getRange(verification.row, COLS_DB.TOKEN_EXPIRY).setValue(null);
    return { success: true, message: "เปลี่ยนรหัสผ่านสำเร็จแล้ว!" };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาดในการอัปเดตรหัสผ่าน" };
  }
}

// =================================================================
// ส่วนที่ 5: ฟังก์ชันเกี่ยวกับข้อมูลโปรไฟล์ผู้ใช้ (USER PROFILE)
// =================================================================
function getUserProfile(email) {
  if (!email) return null;
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (!sheet || sheet.getLastRow() < 2) return null;
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    for (const row of data) {
      if (row[COLS_PROFILE.EMAIL - 1] === email) {
        const profile = {};
        headers.forEach((header, index) => {
          profile[header] = (row[index] instanceof Date) ? row[index].toISOString() : row[index];
        });
        if (profile.ProfilePictureID) {
          profile.profilePictureUrl = `https://lh3.googleusercontent.com/d/${profile.ProfilePictureID}`;
        }
        return profile;
      }
    }
    return null;
  } catch (e) {
    Logger.log(`Error in getUserProfile: ${e.message}`);
    return null;
  }
}

function saveUserProfile(profileData) {
  try {
    const userEmail = profileData.email;
    if (!userEmail) throw new Error("ไม่สามารถระบุตัวตนผู้ใช้ได้ (Email not provided)");
    const profileSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (!profileSheet) throw new Error("ไม่พบชีต 'User_Profiles'");
    const dbSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    if (!dbSheet) throw new Error("ไม่พบชีต 'User_Database'");
    const dbData = dbSheet.getRange(2, COLS_DB.USER_ID, dbSheet.getLastRow(), 2).getValues();
    let userId = null;
    for (const row of dbData) {
      if (row[1] === userEmail) {
        userId = row[0];
        break;
      }
    }
    if (!userId) throw new Error(`ไม่พบ UserID สำหรับอีเมล: ${userEmail}`);
    let pictureFileId = profileData.existingPictureId || null;
    if (profileData.profilePicture && profileData.mimeType) {
        const decodedImage = Utilities.base64Decode(profileData.profilePicture);
        const blob = Utilities.newBlob(decodedImage, profileData.mimeType, `${userId}_profile.jpg`);
        const folder = DriveApp.getFolderById(PROFILE_PICTURE_FOLDER_ID);
        const oldFiles = folder.getFilesByName(`${userId}_profile.jpg`);
        while(oldFiles.hasNext()) { oldFiles.next().setTrashed(true); }
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        pictureFileId = file.getId();
    }
    const profileRowData = [userId,userEmail,pictureFileId,profileData.prefix,profileData.studentId,profileData.firstNameTH,profileData.lastNameTH,profileData.firstNameEN,profileData.lastNameEN,profileData.gradClass,profileData.advisor,profileData.dob?new Date(profileData.dob):null,profileData.gender,profileData.birthCountry,profileData.nationality,profileData.race,profileData.phone,profileData.gpax,profileData.address,profileData.emergencyName,profileData.emergencyRelation,profileData.emergencyPhone,profileData.awards,profileData.futurePlan,profileData.educationPlan,profileData.internationalPlan,profileData.willTakeThaiLicense];
    const profileUserIds = profileSheet.getRange(2, COLS_PROFILE.USER_ID, profileSheet.getLastRow(), 1).getValues().flat();
    const rowIndex = profileUserIds.indexOf(userId);
    if (rowIndex !== -1) {
        profileSheet.getRange(rowIndex + 2, 1, 1, profileRowData.length).setValues([profileRowData]);
    } else {
        profileSheet.appendRow(profileRowData);
    }
    return "บันทึกข้อมูลส่วนตัวสำเร็จ!";
  } catch (e) {
    Logger.log("เกิดข้อผิดพลาดใน saveUserProfile: " + e.message);
    return `เกิดข้อผิดพลาด: ${e.message}`;
  }
}

// =================================================================
// ส่วนที่ 6: ฟังก์ชันเกี่ยวกับใบประกอบวิชาชีพ (LICENSE TRACKING)
// =================================================================
function saveLicenseExamRecord(recordData) {
  try {
    const email = recordData.userEmail;
    if (!email) throw new Error("ไม่พบอีเมลผู้ใช้");
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('License_Exam_Records');
    if (!sheet) throw new Error("ไม่พบชีต License_Exam_Records");
    const dbSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    const dbData = dbSheet.getRange(2, 1, dbSheet.getLastRow(), 2).getValues();
    let userId = null;
    for (const row of dbData) {
      if (row[1] === email) {
        userId = row[0];
        break;
      }
    }
    if (!userId) throw new Error("ไม่พบ UserID สำหรับอีเมลนี้");
    const newRecordRow = [userId,email,recordData.examRound,recordData.examSession,recordData.examYear,recordData.Subject1_Maternity,recordData.Subject2_Pediatric,recordData.Subject3_Adult,recordData.Subject4_Geriatric,recordData.Subject5_Psychiatric,recordData.Subject6_Community,recordData.Subject7_Law,recordData.Subject8_Surgical,null];
    sheet.appendRow(newRecordRow);
    return { success: true, message: 'บันทึกข้อมูลการสอบสำเร็จ!' };
  } catch (e) {
    Logger.log("Error in saveLicenseExamRecord: " + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึก: ' + e.message };
  }
}

function getLicenseExamStatus(email) {
  const subjects = ['Subject1_Maternity','Subject2_Pediatric','Subject3_Adult','Subject4_Geriatric','Subject5_Psychiatric','Subject6_Community','Subject7_Law','Subject8_Surgical'];
  const status = {};
  subjects.forEach(subject => status[subject] = false);
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('License_Exam_Records'); 
    if (!sheet || sheet.getLastRow() < 2) { return status; }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    for (const row of data) {
      const rowEmail = row[headers.indexOf('Email')];
      if (rowEmail === email) {
        subjects.forEach(subjectKey => {
          const subjectIndex = headers.indexOf(subjectKey);
          if (row[subjectIndex] === 'ผ่าน') {
            status[subjectKey] = true;
          }
        });
      }
    }
    return status;
  } catch (e) {
    Logger.log('Error in getLicenseExamStatus: ' + e.toString());
    return status;
  }
}

function getLicenseExamHistory(email) {
  try {
    if (!email) return [];
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('License_Exam_Records');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const history = [];
    for (const row of data) {
      const record = {};
      const rowEmail = row[headers.indexOf('Email')];
      if (rowEmail === email) {
        headers.forEach((header, index) => {
          record[header] = row[index];
        });
        history.push(record);
      }
    }
    return history;
  } catch(e) {
    Logger.log("Error in getLicenseExamHistory: " + e.message);
    return [];
  }
}

// =================================================================
// ส่วนที่ 7: ฟังก์ชันสำหรับ Admin (ADMIN FUNCTIONS)
// =================================================================
function getExamYears() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Config');
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
  } catch (e) {
    Logger.log("Error in getExamYears: " + e.message);
    return [];
  }
}

function addExamYear(year) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Config');
    if (!sheet) throw new Error("ไม่พบชีต Config");
    const existingYears = getExamYears();
    if (existingYears.includes(year)) {
      return { success: false, message: 'ปี ' + year + ' มีอยู่แล้วในระบบ' };
    }
    sheet.appendRow([year]);
    return { success: true, message: 'เพิ่มปี ' + year + ' สำเร็จ' };
  } catch (e) {
    Logger.log("Error in addExamYear: " + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

/**
 * ค้นหาผู้ใช้จากชื่อ, นามสกุล, หรืออีเมล
 * @param {string} searchText - คำที่ใช้ค้นหา
 * @returns {object[]} อาร์เรย์ของข้อมูลผู้ใช้ที่ค้นเจอ
 */
function searchUsers(searchText) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    if (!sheet || sheet.getLastRow() < 2) return [];

    const lowerCaseSearch = searchText.toLowerCase();
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, COLS_DB.LAST_NAME).getValues();

    const results = [];
    data.forEach(row => {
      const firstName = row[COLS_DB.FIRST_NAME - 1].toLowerCase();
      const lastName = row[COLS_DB.LAST_NAME - 1].toLowerCase();
      const email = row[COLS_DB.EMAIL - 1].toLowerCase();

      if (firstName.includes(lowerCaseSearch) || lastName.includes(lowerCaseSearch) || email.includes(lowerCaseSearch)) {
        results.push({
          userId: row[COLS_DB.USER_ID - 1],
          firstName: row[COLS_DB.FIRST_NAME - 1],
          lastName: row[COLS_DB.LAST_NAME - 1],
          email: row[COLS_DB.EMAIL - 1],
          role: row[COLS_DB.ROLE - 1]
        });
      }
    });
    return results;
  } catch(e) {
    Logger.log("Error in searchUsers: " + e.message);
    return [];
  }
}