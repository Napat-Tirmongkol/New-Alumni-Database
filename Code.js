// @ts-nocheck
// =================================================================
// ส่วนที่ 1: การตั้งค่าและตัวแปรหลัก (CONFIGURATION)
// =================================================================
const SHEET_ID = "1yXsIRJB56XDGPANqMTD1tbAgPPZhnRVj5rwCg1F3Ew8";
const PROFILE_PICTURE_FOLDER_ID = "12VC1rdV1sRXxFP0M-QnZgytokZHwwx26";
const EVIDENCE_FOLDER_ID = "1JZBj33PHAtRcPsMi_URIB0s9Y4VDPFeI";
const DONATION_SLIPS_FOLDER_ID = "1okFXJVYbVAwrQviGcStjEsElsGuFOknK";

const COLS_DB = {
  USER_ID: 1, EMAIL: 2, PASSWORD: 3, ROLE: 4, FIRST_NAME: 5, LAST_NAME: 6,
  REG_DATE: 7, LAST_LOGIN: 8, IS_ACTIVE: 9, RESET_TOKEN: 10, TOKEN_EXPIRY: 11
};
const COLS_PROFILE = {
  USER_ID: 1, EMAIL: 2, PICTURE_ID: 3, PREFIX: 4, STUDENT_ID: 5, FNAME_TH: 6, LNAME_TH: 7,
  FNAME_EN: 8, LNAME_EN: 9, GRAD_CLASS: 10, ADVISOR: 11, DOB: 12, GENDER: 13,
  BIRTH_COUNTRY: 14, NATIONALITY: 15, RACE: 16, PHONE: 17, GPAX: 18, EMPLOYMENT_STATUS: 19,
  ADDRESS: 20, EMERGENCY_NAME: 21, EMERGENCY_RELATION: 22, EMERGENCY_PHONE: 23, AWARDS: 24,
  FUTURE_PLAN: 25, EDUCATION_PLAN: 26, INTERNATIONAL_PLAN: 27, THAI_LICENSE: 28
};

// =================================================================
// ส่วนที่ 2: ฟังก์ชัน Setup (สำหรับสร้าง/ซ่อมแซมชีต)
// =================================================================
function setupProjectSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const userDbHeaders = ['UserID', 'Email', 'Password', 'Role', 'RegistrationDate', 'LastLogin', 'IsActive', 'ResetToken', 'TokenExpiry'];
    const userProfileHeaders = ['UserID','Email', 'ProfilePictureID', 'Prefix', 'StudentID', 'FirstNameTH', 'LastNameTH', 'FirstNameEN', 'LastNameEN', 'GraduationClass', 'Advisor', 'DateOfBirth', 'Gender', 'BirthCountry', 'Nationality', 'Race', 'PhoneNumber', 'GPAX','EmploymentStatus', 'CurrentAddress', 'EmergencyContactName', 'EmergencyContactRelation', 'EmergencyContactPhone', 'Awards', 'FutureWorkPlan', 'EducationPlan', 'InternationalWorkPlan', 'WillTakeThaiLicense'];
    const licenseExamHeaders = ['RecordID', 'UserID', 'Email', 'ExamRound', 'ExamSession', 'ExamYear', 'Subject1_Maternity', 'Subject2_Pediatric', 'Subject3_Adult', 'Subject4_Geriatric', 'Subject5_Psychiatric', 'Subject6_Community', 'Subject7_Law', 'Subject8_Surgical', 'EvidenceFileID'];
    const donationLogHeaders = ['DonationID', 'UserID', 'Email', 'Amount', 'Purpose', 'Message', 'DonorName', 'DonorAddress', 'donorTaxid','SlipFileID', 'Timestamp', 'Status'];


    checkAndCreateSheet(spreadsheet, 'User_Database', userDbHeaders);
    checkAndCreateSheet(spreadsheet, 'User_Profiles', userProfileHeaders);
    checkAndCreateSheet(spreadsheet, 'License_Exam_Records', licenseExamHeaders);
    checkAndCreateSheet(spreadsheet, 'Donations_Log', donationLogHeaders);
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
    'profile': 'Profile',
    'profileSummary': 'ProfileSummary',
    'license': 'LicenseTracking',
    'donations': 'Donations',
    'virtual_id': 'VirtualID',
    'adminDashboard': 'AdminStatsDashboard',
    'manageUsers': 'AdminUserManagement', 
    'licenseForAdmin': 'LicenseTracking',
    'manageBackend': 'AdminBackendManagement' 
  };
  if (pageMap[pageName]) { return include(pageMap[pageName]); }
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

    // ดึงข้อมูลตามโครงสร้างชีตใหม่
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    for (let i = 0; i < data.length; i++) {
      const rowData = data[i];
      // ตรวจสอบเพื่อป้องกันการเข้าถึงแถวว่างเปล่า
      if (!rowData || rowData.length < 4) {
        continue;
      }
      const email = rowData[1]; // คอลัมน์ B
      const password = String(rowData[2]); // คอลัมน์ C
      const role = rowData[3]; // คอลัมน์ D
      const isActive = rowData[6]; // คอลัมน์ G

      if (email === credentials.email && password === credentials.password) {
        if (isActive === false) { return { success: false, message: 'บัญชีของคุณถูกระงับการใช้งาน' }; }

        const userRow = i + 2;
        sheet.getRange(userRow, 6).setValue(new Date()); // อัปเดต LastLogin ในคอลัมน์ F

        // ดึงชื่อ-นามสกุลจากชีต User_Profiles แทน
        const profileSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
        const profileData = profileSheet.getRange(2, 1, profileSheet.getLastRow() - 1, profileSheet.getLastColumn()).getValues();
        const profileRow = profileData.find(row => row[1] === email);

        if (!profileRow) {
          throw new Error("ไม่พบข้อมูลโปรไฟล์สำหรับผู้ใช้นี้");
        }

        const user = {
          firstName: profileRow[5], // Index 5 คือ FirstNameTH ในชีต User_Profiles
          lastName: profileRow[6],  // Index 6 คือ LastNameTH ในชีต User_Profiles
          role: role
        };

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

    // สร้างแถวข้อมูลสำหรับชีต User_Database
    const newUserDbRow = [
      newUserId, 
      userData.email, 
      userData.password, 
      userData.userType,
      new Date(), // RegistrationDate
      null,      // LastLogin
      true,      // IsActive
      null,      // ResetToken
      null       // TokenExpiry
    ];
    dbSheet.appendRow(newUserDbRow);

    // สร้างแถวข้อมูลสำหรับชีต User_Profiles
    const profileSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (profileSheet) {
      const newUserProfileRow = [
        newUserId,
        userData.email,
        '', // ProfilePictureID
        '', // Prefix (จาก Register.html ไม่มีฟิลด์นี้)
        '', // StudentID (จาก Register.html ไม่มีฟิลด์นี้)
        userData.firstName, // FirstNameTH
        userData.lastName,  // LastNameTH
      ];
      profileSheet.appendRow(newUserProfileRow);
    }
    
    // สร้างแถวข้อมูลสำหรับชีต License_Exam_Records
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
    const data = sheet.getRange(2, 1, sheet.getLastRow(), COLS_DB.IS_ACTIVE).getValues();
for (let i = 0; i < data.length; i++) {
  const rowData = data[i];
  // เพิ่มการตรวจสอบความยาวของแถวข้อมูล
  if (rowData.length < COLS_DB.FIRST_NAME) {
    continue; // ถ้าแถวสั้นเกินไป (เป็นแถวเปล่า) ให้ข้ามไป
  }
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

/**
 * ฟังก์ชันใหม่: อัปเดตบทบาทผู้ใช้
 * @param {string} email - อีเมลของผู้ใช้ที่ต้องการอัปเดต
 * @returns {object} ผลลัพธ์การดำเนินการ
 */
function updateUserRole(email) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    if (!sheet) {
      throw new Error("ไม่พบชีต 'User_Database'");
    }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const emailIndex = headers.indexOf('Email');
    const roleIndex = headers.indexOf('Role');
    
    if (emailIndex === -1 || roleIndex === -1) {
      throw new Error("ไม่พบคอลัมน์ที่จำเป็นในชีต User_Database");
    }

    let userRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][emailIndex] === email) {
        userRow = i + 2; // +2 เพราะแถวเริ่มที่ 1 และข้าม header
        break;
      }
    }

    if (userRow === -1) {
      return { success: false, message: 'ไม่พบอีเมลผู้ใช้ในระบบ' };
    }

    sheet.getRange(userRow, roleIndex + 1).setValue('Alumni');
    
    return { success: true, message: 'เปลี่ยนสถานะเป็นศิษย์เก่าสำเร็จ' };

  } catch (e) {
    Logger.log("Error in updateUserRole: " + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาดในการอัปเดตบทบาท: ' + e.message };
  }
}

// =================================================================
// ส่วนที่ 5: ฟังก์ชันเกี่ยวกับข้อมูลโปรไฟล์ผู้ใช้ (USER PROFILE)
// =================================================================
function getProfilePageAndData(email) {
  const html = include('Profile.html');
  const profileData = getUserProfile(email);
  return { html: html, profileData: profileData };
}

// และแก้ไขฟังก์ชัน getUserProfile ให้เป็นแบบนี้
function getUserProfile(email) {
  if (!email) {
    return null;
  }
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (!sheet || sheet.getLastRow() < 2) {
      return null;
    }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    for (const row of data) {
      const sheetEmail = row[COLS_PROFILE.EMAIL - 1] ? String(row[COLS_PROFILE.EMAIL - 1]).trim() : '';
      if (sheetEmail === email.trim()) {
        const profile = {};
        headers.forEach((header, index) => {
          profile[header] = (row[index] instanceof Date) ? row[index].toISOString() : row[index];
        });
        if (profile.ProfilePictureID) {
          profile.profilePictureUrl = `https://lh3.googleusercontent.com/u/0/d/${profile.ProfilePictureID}`;
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
    
    // ค้นหา UserID จากชีต User_Database
    const dbData = dbSheet.getRange(2, COLS_DB.USER_ID, dbSheet.getLastRow() - 1, 2).getValues();
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
    
    // สร้างข้อมูลแถวใหม่สำหรับ User_Profiles
    const profileRowData = [userId,userEmail,pictureFileId,profileData.prefix,profileData.studentId,profileData.firstNameTH,profileData.lastNameTH,profileData.firstNameEN,profileData.lastNameEN,profileData.gradClass,profileData.advisor,profileData.dob?new Date(profileData.dob):null,profileData.gender,profileData.birthCountry,profileData.nationality,profileData.race,profileData.phone,profileData.gpax,profileData.employmentStatus,profileData.address,profileData.emergencyName,profileData.emergencyRelation,profileData.emergencyPhone,profileData.awards,profileData.futurePlan,profileData.educationPlan,profileData.internationalPlan,profileData.willTakeThaiLicense];

    // ค้นหาแถวที่ต้องอัปเดต
    const profileUserIds = profileSheet.getRange(2, COLS_PROFILE.USER_ID, profileSheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = profileUserIds.indexOf(userId);

    if (rowIndex !== -1) {
        // ถ้าพบ ให้อัปเดตข้อมูลในแถวนั้น
        profileSheet.getRange(rowIndex + 2, 1, 1, profileRowData.length).setValues([profileRowData]);
    } else {
        // ถ้าไม่พบ ให้เพิ่มแถวใหม่
        profileSheet.appendRow(profileRowData);
    }
    
    return "บันทึกข้อมูลสำเร็จ!";
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

    let evidenceFileId = null;

    if (recordData.file) {
      const decodedImage = Utilities.base64Decode(recordData.file.base64);
      const blob = Utilities.newBlob(decodedImage, recordData.file.mimeType, recordData.file.fileName);
      const folder = DriveApp.getFolderById(EVIDENCE_FOLDER_ID);
      const file = folder.createFile(blob);
      evidenceFileId = file.getId();
    }

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
    
    // ----- จุดที่แก้ไข: เพิ่มการสร้าง RecordID ที่นี่ -----
    const recordId = "LID-" + new Date().getTime(); 

    const newRecordRow = [
      recordId, // เพิ่ม RecordID เข้าไปเป็นค่าแรกของแถว
      userId, 
      email,
      recordData.examRound, 
      recordData.examSession, 
      recordData.examYear,
      recordData.Subject1_Maternity, 
      recordData.Subject2_Pediatric,
      recordData.Subject3_Adult, 
      recordData.Subject4_Geriatric,
      recordData.Subject5_Psychiatric, 
      recordData.Subject6_Community,
      recordData.Subject7_Law, 
      recordData.Subject8_Surgical,
      evidenceFileId
    ];

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
    const emailIndex = headers.indexOf('Email');

    if(emailIndex === -1) return [];

    for (const row of data) {
      if (row[emailIndex] === email) {
        const record = {};
        headers.forEach((header, index) => {
          record[header] = row[index];
        });
        history.push(record);
      }
    }
    return history.sort((a, b) => b.ExamYear - a.ExamYear || b.ExamRound - a.ExamRound);

  } catch(e) {
    Logger.log("Error in getLicenseExamHistory: " + e.message);
    return [];
  }
}

function updateLicenseExamRecord(recordData) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('License_Exam_Records');
    if (!sheet) throw new Error("ไม่พบชีต License_Exam_Records");

    const recordIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = recordIds.indexOf(recordData.recordId);

    if (rowIndex === -1) {
      return { success: false, message: 'ไม่พบข้อมูลการสอบที่ต้องการแก้ไข' };
    }
    
    const rowToUpdate = rowIndex + 2;
    // [RecordID, UserID, Email, ExamRound, ExamSession, ExamYear, Subject1, ...]
    const updatedRowData = [
      recordData.recordId,
      recordData.userId, // ต้องแน่ใจว่าส่ง UserID มาด้วย
      recordData.email,   // ต้องแน่ใจว่าส่ง Email มาด้วย
      recordData.examRound,
      recordData.examSession,
      recordData.examYear,
      recordData.Subject1_Maternity,
      recordData.Subject2_Pediatric,
      recordData.Subject3_Adult,
      recordData.Subject4_Geriatric,
      recordData.Subject5_Psychiatric,
      recordData.Subject6_Community,
      recordData.Subject7_Law,
      recordData.Subject8_Surgical
      // ไม่รวมการอัปเดตไฟล์หลักฐานในเวอร์ชันนี้เพื่อความง่าย
    ];
    
    // อัปเดตข้อมูล 14 คอลัมน์ (ไม่รวม file ID)
    sheet.getRange(rowToUpdate, 1, 1, 14).setValues([updatedRowData]);

    return { success: true, message: 'อัปเดตข้อมูลการสอบสำเร็จ!' };
  } catch (e) {
    Logger.log("Error in updateLicenseExamRecord: " + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

// ******** ฟังก์ชันใหม่สำหรับลบข้อมูล ********
function deleteLicenseExamRecord(recordId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('License_Exam_Records');
    if (!sheet) throw new Error("ไม่พบชีต License_Exam_Records");

    const recordIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = recordIds.indexOf(recordId);

    if (rowIndex === -1) {
      return { success: false, message: 'ไม่พบข้อมูลการสอบที่ต้องการลบ' };
    }

    sheet.deleteRow(rowIndex + 2);
    
    return { success: true, message: 'ลบข้อมูลการสอบสำเร็จ!' };
  } catch (e) {
    Logger.log("Error in deleteLicenseExamRecord: " + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
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
 * Retrieves a list of users for the admin dashboard with pagination and search.
 * @param {number} page The page number to retrieve.
 * @param {string} searchText The text to search for.
 * @return {object} An object containing the list of users and the total count.
 */
function getUsers(page = 1, searchText = '') {
    try {
        const dbSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
        if (!dbSheet) throw new Error("ไม่พบชีต 'User_Database'");

        const profileSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
        if (!profileSheet) throw new Error("ไม่พบชีต 'User_Profiles'");

        const dbData = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, dbSheet.getLastColumn()).getValues();
        const profileData = profileSheet.getRange(2, 1, profileSheet.getLastRow() - 1, profileSheet.getLastColumn()).getValues();

        const usersWithProfile = dbData.map(dbRow => {
            const userEmail = dbRow[COLS_DB.EMAIL - 1];
            const profileRow = profileData.find(pRow => pRow[COLS_PROFILE.EMAIL - 1] === userEmail);
            
            const profile = profileRow ? {
                firstName: profileRow[COLS_PROFILE.FNAME_TH - 1],
                lastName: profileRow[COLS_PROFILE.LNAME_TH - 1],
                email: profileRow[COLS_PROFILE.EMAIL - 1]
            } : {
                firstName: '',
                lastName: '',
                email: userEmail
            };

            return {
                id: dbRow[COLS_DB.USER_ID - 1],
                email: userEmail,
                role: dbRow[COLS_DB.ROLE - 1],
                firstName: profile.firstName,
                lastName: profile.lastName,
            };
        });

        const filteredUsers = searchText
            ? usersWithProfile.filter(user =>
                user.firstName.toLowerCase().includes(searchText.toLowerCase()) ||
                user.lastName.toLowerCase().includes(searchText.toLowerCase()) ||
                user.email.toLowerCase().includes(searchText.toLowerCase())
            )
            : usersWithProfile;

        const pageSize = 10;
        const start = (page - 1) * pageSize;
        const end = start + pageSize;
        const paginatedUsers = filteredUsers.slice(start, end);

        return {
            users: paginatedUsers,
            totalUsers: filteredUsers.length
        };
    } catch (e) {
        Logger.log("Error in getUsers: " + e.message);
        return { users: [], totalUsers: 0 };
    }
}

/**
 * Deletes a user and all their associated data from all sheets.
 * @param {string} email The email of the user to delete.
 * @return {object} An object indicating the success of the operation.
 */
function deleteUser(email) {
  try {
    if (!email) {
      throw new Error("ไม่ระบุอีเมลผู้ใช้ที่ต้องการลบ");
    }

    const sheetsToDeleteFrom = [
      'User_Database', 
      'User_Profiles', 
      'License_Exam_Records', 
      'Donations_Log'
    ];

    sheetsToDeleteFrom.forEach(sheetName => {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        const emailColumnIndex = data[0].indexOf('Email');
        
        if (emailColumnIndex !== -1) {
          // Iterate backwards to avoid issues with changing row indices
          for (let i = data.length - 1; i >= 1; i--) {
            if (data[i][emailColumnIndex] === email) {
              sheet.deleteRow(i + 1);
            }
          }
        }
      }
    });

    return { success: true, message: 'ลบผู้ใช้สำเร็จ' };
  } catch (e) {
    Logger.log("Error in deleteUser: " + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาดในการลบผู้ใช้: ' + e.message };
  }
}

/**
 * Fetches the Profile.html template and a specific user's profile data.
 * @param {string} email The email of the user to fetch.
 * @return {object} An object with the HTML template and the user's profile data.
 */
function getProfilePageAndData(email) {
    const htmlTemplate = HtmlService.createTemplateFromFile('Profile');
    const profileData = getUserProfile(email);
    
    return {
        html: htmlTemplate.evaluate().getContent(),
        profileData: profileData
    };
}

/**
 * คำนวณและดึงข้อมูลสรุปสำหรับหน้าแดชบอร์ด (ฉบับแก้ไข)
 * @returns {object} อ็อบเจกต์ข้อมูลสถิติ
 */
function getDashboardStats() {
  try {
    const dbSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    const licenseSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('License_Exam_Records');

    if (!dbSheet) throw new Error("ไม่พบชีต User_Database");

    // --- คำนวณจำนวนผู้ใช้ทั้งหมด ---
    const roles = dbSheet.getRange(2, COLS_DB.ROLE, dbSheet.getLastRow() - 1, 1).getValues().flat();
    const alumniCount = roles.filter(role => role === 'Alumni').length;
    const studentCount = roles.filter(role => role === 'Student').length;
    const totalUsers = alumniCount + studentCount;

    // --- ตรรกะสำหรับนับจำนวนผู้ที่สอบผ่านครบทุกวิชา ---
    let passedAllCount = 0;
    if (licenseSheet && licenseSheet.getLastRow() > 1) {
      const licenseData = licenseSheet.getDataRange().getValues();
      const headers = licenseData.shift();
      const emailIndex = headers.indexOf('Email');
      const subjectKeys = ['Subject1_Maternity', 'Subject2_Pediatric', 'Subject3_Adult', 'Subject4_Geriatric', 'Subject5_Psychiatric', 'Subject6_Community', 'Subject7_Law', 'Subject8_Surgical'];
      const subjectIndices = subjectKeys.map(key => headers.indexOf(key));
      const userPassStatus = {};

      licenseData.forEach(row => {
        const email = row[emailIndex];
        if (!email) return;
        if (!userPassStatus[email]) {
          userPassStatus[email] = {};
        }
        subjectIndices.forEach((subjIndex, i) => {
          if (row[subjIndex] === 'ผ่าน') {
            userPassStatus[email][subjectKeys[i]] = true;
          }
        });
      });

      for (const email in userPassStatus) {
        const subjectsPassedCount = Object.values(userPassStatus[email]).filter(status => status === true).length;
        if (subjectsPassedCount === 8) {
          passedAllCount++;
        }
      }
    }
    
    // --- ส่วนที่เพิ่มเข้ามา: คำนวณอัตราการสอบผ่าน ---
    let passRate = 0;
    if (totalUsers > 0) {
      passRate = Math.round((passedAllCount / totalUsers) * 100);
    }

    return {
      totalUsers: totalUsers,
      alumniCount: alumniCount,
      studentCount: studentCount,
      passedAllCount: passedAllCount, // เปลี่ยนชื่อ Key เพื่อความชัดเจน
      passRate: passRate // เพิ่ม Key ใหม่สำหรับอัตราผ่าน
    };
  } catch (e) {
    Logger.log("Error in getDashboardStats: " + e.message);
    return { totalUsers: 0, alumniCount: 0, studentCount: 0, passedAllCount: 0, passRate: 0 };
  }
}

function saveDonationRecord(donationData) {
  try {
    const email = donationData.userEmail;
    if (!email) throw new Error("ไม่พบอีเมลผู้ใช้");

    let slipFileId = null;
    if (donationData.file) {
      const decodedSlip = Utilities.base64Decode(donationData.file.base64);
      const blob = Utilities.newBlob(decodedSlip, donationData.file.mimeType, donationData.file.fileName);
      const folder = DriveApp.getFolderById(DONATION_SLIPS_FOLDER_ID);
      const file = folder.createFile(blob);
      slipFileId = file.getId();
    }

    const logSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Donations_Log');
    if (!logSheet) throw new Error("ไม่พบชีต Donations_Log");

    const dbSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    const dbData = dbSheet.getRange(2, 1, dbSheet.getLastRow(), 2).getValues();
    let userId = null;
    for (const row of dbData) {
      if (row[1] === email) { userId = row[0]; break; }
    }

    const newDonationRow = [
      "DN-" + new Date().getTime(), // DonationID
      userId,
      email,
      donationData.amount,
      donationData.purpose,
      donationData.message,
      donationData.donorName,
      donationData.donorAddress,
      donationData.donorTaxId,
      slipFileId,
      new Date(), // Timestamp
      "Pending Verification" // Status
    ];

    logSheet.appendRow(newDonationRow);
    return { success: true, message: 'บันทึกข้อมูลการบริจาคสำเร็จ! เจ้าหน้าที่จะทำการตรวจสอบและติดต่อกลับโดยเร็วที่สุด' };
  } catch (e) {
    Logger.log("Error in saveDonationRecord: " + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึก: ' + e.message };
  }
}

/**
 * ดึงข้อมูลประวัติการบริจาคของผู้ใช้ (ฉบับแก้ไข)
 * @param {string} email - อีเมลของผู้ใช้ที่ต้องการดึงข้อมูล
 * @returns {object[]} อาร์เรย์ของข้อมูลการบริจาค
 */
function getDonationHistory(email) {
  try {
    if (!email) return [];
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Donations_Log');
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const history = [];
    
    // หา Index ของคอลัมน์ทั้งหมดที่เราต้องการ
    const emailIndex = headers.indexOf('Email');
    if (emailIndex === -1) return []; // ถ้าไม่พบคอลัมน์ Email ให้หยุดทำงาน

    for (const row of data) {
      if (row[emailIndex] === email) {
        const record = {};
        headers.forEach((header, index) => {
          // แปลงวันที่ให้เป็น ISO string เพื่อให้จัดการง่ายในฝั่ง Client
          record[header] = row[index] instanceof Date ? row[index].toISOString() : row[index];
        });
        history.push(record);
      }
    }
    // เรียงลำดับจากล่าสุดไปเก่าสุด
    return history.sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));
  } catch(e) {
    Logger.log("Error in getDonationHistory: " + e.message);
    return [];
  }
}


/**
 * ฟังก์ชันใหม่: อัปเดตข้อมูลการบริจาค
 * @param {object} data - ข้อมูลการบริจาคที่ต้องการอัปเดต
 */
function updateDonationRecord(data) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Donations_Log');
    if (!sheet) throw new Error("ไม่พบชีต Donations_Log");

    const dataValues = sheet.getDataRange().getValues();
    const headers = dataValues.shift();
    const idColumnIndex = headers.indexOf('DonationID');

    let rowIndex = -1;
    for (let i = 0; i < dataValues.length; i++) {
      if (dataValues[i][idColumnIndex] === data.DonationID) {
        rowIndex = i + 2; // +2 เพราะแถวเริ่มที่ 1 และเราข้าม header ไป 1 แถว
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'ไม่พบข้อมูลการบริจาคที่ต้องการแก้ไข' };
    }

    // อัปเดตค่าในแต่ละคอลัมน์
    headers.forEach((header, index) => {
      if (data[header] !== undefined) {
        sheet.getRange(rowIndex, index + 1).setValue(data[header]);
      }
    });

    return { success: true, message: 'อัปเดตข้อมูลการบริจาคสำเร็จ!' };
  } catch (e) {
    Logger.log("Error in updateDonationRecord: " + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

/**
 * ฟังก์ชันใหม่: ลบข้อมูลการบริจาค
 * @param {string} donationId - ID ของการบริจาคที่ต้องการลบ
 */
function deleteDonationRecord(donationId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Donations_Log');
    if (!sheet) throw new Error("ไม่พบชีต Donations_Log");

    const dataValues = sheet.getDataRange().getValues();
    const headers = dataValues.shift();
    const idColumnIndex = headers.indexOf('DonationID');

    let rowIndex = -1;
    for (let i = 0; i < dataValues.length; i++) {
      if (dataValues[i][idColumnIndex] === donationId) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'ไม่พบข้อมูลการบริจาคที่ต้องการลบ' };
    }

    sheet.deleteRow(rowIndex);
    
    return { success: true, message: 'ลบข้อมูลการบริจาคสำเร็จ!' };
  } catch (e) {
    Logger.log("Error in deleteDonationRecord: " + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

/**
 * ดึงข้อมูลการกระจายตัวตามเพศจากชีต User_Profiles
 * @returns {object} อ็อบเจกต์ข้อมูลจำนวนแต่ละเพศ
 */
function getGenderDistribution() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (!sheet || sheet.getLastRow() < 2) return {};

    const genderData = sheet.getRange(2, COLS_PROFILE.GENDER, sheet.getLastRow() - 1, 1).getValues().flat();

    const distribution = {
      male: 0,
      female: 0,
      other: 0
    };

    genderData.forEach(gender => {
      if (gender === 'ชาย') {
        distribution.male++;
      } else if (gender === 'หญิง') {
        distribution.female++;
      } else if (gender) { // นับ "ไม่ระบุ" หรือค่าอื่นๆ ที่ไม่ใช่ค่าว่าง
        distribution.other++;
      }
    });

    return distribution;
  } catch (e) {
    Logger.log("Error in getGenderDistribution: " + e.message);
    return {};
  }
}

/**
 * ฟังก์ชันใหม่: ดึงข้อมูลการกระจายตัวตามอายุจากชีต User_Profiles
 * @returns {object} อ็อบเจกต์ข้อมูลจำนวนผู้ใช้ในแต่ละช่วงอายุ
 */
function getAgeDistribution() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (!sheet || sheet.getLastRow() < 2) return {};

    // กำหนดช่วงอายุที่ต้องการ
    const ageGroups = {
      'ต่ำกว่า 20': 0,
      '20-25 ปี': 0,
      '26-30 ปี': 0,
      '31-40 ปี': 0,
      '41-50 ปี': 0,
      'มากกว่า 50': 0
    };

    const dobData = sheet.getRange(2, COLS_PROFILE.DOB, sheet.getLastRow() - 1, 1).getValues().flat();
    const currentYear = new Date().getFullYear();

    dobData.forEach(dob => {
      if (dob instanceof Date) {
        const birthYear = dob.getFullYear();
        const age = currentYear - birthYear;

        if (age < 20) ageGroups['ต่ำกว่า 20']++;
        else if (age >= 20 && age <= 25) ageGroups['20-25 ปี']++;
        else if (age >= 26 && age <= 30) ageGroups['26-30 ปี']++;
        else if (age >= 31 && age <= 40) ageGroups['31-40 ปี']++;
        else if (age >= 41 && age <= 50) ageGroups['41-50 ปี']++;
        else ageGroups['มากกว่า 50']++;
      }
    });

    return ageGroups;
  } catch (e) {
    Logger.log("Error in getAgeDistribution: " + e.message);
    return {};
  }
}

/**
 * ฟังก์ชันใหม่: ดึงข้อมูลการกระจายตัวตามสัญชาติ
 * @returns {object} อ็อบเจกต์ข้อมูลจำนวนผู้ใช้ในแต่ละสัญชาติ
 */
function getNationalityDistribution() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (!sheet || sheet.getLastRow() < 2) return {};

    const distribution = {};
    const nationalityData = sheet.getRange(2, COLS_PROFILE.NATIONALITY, sheet.getLastRow() - 1, 1).getValues().flat();

    nationalityData.forEach(nationality => {
      if (nationality) { // ตรวจสอบว่าค่าไม่เป็นค่าว่าง
        if (distribution[nationality]) {
          distribution[nationality]++;
        } else {
          distribution[nationality] = 1;
        }
      }
    });

    return distribution;
  } catch (e) {
    Logger.log("Error in getNationalityDistribution: " + e.message);
    return {};
  }
}

/**
 * ฟังก์ชันใหม่: ดึงข้อมูลการกระจายตัวตามรุ่นที่จบ
 * @returns {object} อ็อบเจกต์ข้อมูลจำนวนผู้ใช้ในแต่ละรุ่น
 */
function getClassDistribution() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (!sheet || sheet.getLastRow() < 2) return {};

    const distribution = {};
    const classData = sheet.getRange(2, COLS_PROFILE.GRAD_CLASS, sheet.getLastRow() - 1, 1).getValues().flat();

    classData.forEach(gradClass => {
      if (gradClass) { // ตรวจสอบว่าค่าไม่เป็นค่าว่าง
        const className = String(gradClass).trim(); // Trim whitespace and convert to string
        if (distribution[className]) {
          distribution[className]++;
        } else {
          distribution[className] = 1;
        }
      }
    });

    return distribution;
  } catch (e) {
    Logger.log("Error in getClassDistribution: " + e.message);
    return {};
  }
}

/**
 * ดึงข้อมูลการกระจายตัวของสถานะการทำงาน (เวอร์ชันแก้ไข: รวมและเปลี่ยนชื่อ 'ทำงานแล้ว' เป็น 'กำลังทำงาน')
 * @returns {object} อ็อบเจกต์ข้อมูลจำนวนในแต่ละสถานะ
 */
function getEmploymentStatusDistribution() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (!sheet || sheet.getLastRow() < 2) return {};
    
    const statuses = sheet.getRange(2, COLS_PROFILE.EMPLOYMENT_STATUS, sheet.getLastRow() - 1, 1).getValues().flat();
    
    const distribution = {};

    statuses.forEach(status => {
      if (!status) return; // ข้ามเซลล์ที่ว่าง

      // ตรวจสอบถ้าในข้อความมีคำว่า 'ทำงานแล้ว'
      if (status.includes('ทำงานแล้ว')) {
        // ให้นับรวมใน key 'กำลังทำงาน' เสมอ
        distribution['กำลังทำงาน'] = (distribution['กำลังทำงาน'] || 0) + 1;
      } else {
        // สำหรับสถานะอื่นๆ ให้นับตามปกติ
        distribution[status] = (distribution[status] || 0) + 1;
      }
    });
    
    return distribution;
  } catch (e) {
    Logger.log("Error in getEmploymentStatusDistribution: " + e.message);
    return {};
  }
}

/**
 * ดึงข้อมูลการกระจายตัวของแผนการไปทำงานต่างประเทศ
 * @returns {object} อ็อบเจกต์ข้อมูลจำนวนคนที่มีแผนและไม่มีแผน
 */
function getInternationalWorkPlanDistribution() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (!sheet || sheet.getLastRow() < 2) return {};
    // ดึงข้อมูลจากคอลัมน์ InternationalWorkPlan (คอลัมน์ที่ 27)
    const plans = sheet.getRange(2, COLS_PROFILE.INTERNATIONAL_PLAN, sheet.getLastRow() - 1, 1).getValues().flat();
    const distribution = {
      'ไม่มีแผนไปทำงานต่างประเทศ': 0,
      'มีแผนไปทำงานต่างประเทศ': 0
    };
    plans.forEach(plan => {
      // หากค่าเป็นค่าว่าง หรือ 'no_plan' ให้นับเป็น "ไม่มีแผน"
      if (!plan || plan === 'no_plan') {
        distribution['ไม่มีแผนไปทำงานต่างประเทศ']++;
      } else {
        distribution['มีแผนไปทำงานต่างประเทศ']++;
      }
    });
    return distribution;
  } catch (e) {
    Logger.log("Error in getInternationalWorkPlanDistribution: " + e.message);
    return {};
  }
}

/**
 * ดึงข้อมูลการกระจายตัวตามประเภทสถานที่ทำงาน (FutureWorkPlan)
 * @returns {object} อ็อบเจกต์ข้อมูลจำนวนในแต่ละประเภท
 */
function getFutureWorkPlanDistribution() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (!sheet || sheet.getLastRow() < 2) return {};
    // ดึงข้อมูลจากคอลัมน์ FutureWorkPlan (คอลัมน์ที่ 25)
    const plans = sheet.getRange(2, COLS_PROFILE.FUTURE_PLAN, sheet.getLastRow() - 1, 1).getValues().flat();
    const distribution = {};
    plans.forEach(plan => {
      if (plan) { // นับเฉพาะค่าที่ไม่ใช่เซลล์ว่าง
        distribution[plan] = (distribution[plan] || 0) + 1;
      }
    });
    return distribution;
  } catch (e) {
    Logger.log("Error in getFutureWorkPlanDistribution: " + e.message);
    return {};
  }
}

/**
 * คำนวณอัตราการสอบผ่านของแต่ละรายวิชาในใบประกอบวิชาชีพ
 * @returns {object} อ็อบเจกต์ที่ประกอบด้วยชื่อวิชา (labels) และข้อมูลอัตราการผ่าน (data)
 */
function getSubjectPassRates() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('License_Exam_Records');
    if (!sheet || sheet.getLastRow() < 2) {
      return { labels: [], data: [] };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    const subjectKeys = [
      'Subject1_Maternity', 'Subject2_Pediatric', 'Subject3_Adult', 
      'Subject4_Geriatric', 'Subject5_Psychiatric', 'Subject6_Community', 
      'Subject7_Law', 'Subject8_Surgical'
    ];

    // ชื่อย่อสำหรับแสดงบนกราฟ
    const thaiLabels = [
      'การผดุงครรภ์', 'พยาบาลมารดาฯ', 'พยาบาลเด็กฯ', 'พยาบาลผู้ใหญ่',
      'พยาบาลผู้สูงอายุ', 'สุขภาพจิตฯ', 'อนามัยชุมชนฯ', 'กฎหมายวิชาชีพฯ'
    ];

    const subjectIndices = subjectKeys.map(key => headers.indexOf(key));
    const passCounts = Array(8).fill(0);
    const totalTakers = {}; // ใช้อีเมลเพื่อนับจำนวนผู้เข้าสอบที่ไม่ซ้ำกันในแต่ละวิชา

    data.forEach(row => {
      const email = row[headers.indexOf('Email')];
      if (!email) return;

      subjectIndices.forEach((subjIndex, i) => {
        // ตรวจสอบว่ามีผลสอบ (ไม่เป็นค่าว่าง)
        if (subjIndex !== -1 && row[subjIndex]) { 
          const subjectKey = subjectKeys[i];
          if (!totalTakers[subjectKey]) {
            totalTakers[subjectKey] = new Set();
          }
          totalTakers[subjectKey].add(email);

          if (row[subjIndex] === 'ผ่าน') {
            passCounts[i]++;
          }
        }
      });
    });

    // คำนวณเป็นเปอร์เซ็นต์
    const passRates = passCounts.map((count, i) => {
        const subjectKey = subjectKeys[i];
        const total = totalTakers[subjectKey] ? totalTakers[subjectKey].size : 0;
        return total > 0 ? Math.round((count / total) * 100) : 0;
    });

    return {
      labels: thaiLabels,
      data: passRates
    };

  } catch (e) {
    Logger.log("Error in getSubjectPassRates: " + e.message);
    return { labels: [], data: [] };
  }
}

/**
 * คำนวณอัตราการสอบผ่านโดยแยกตามปีและรอบสอบ
 * @returns {object} อ็อบเจกต์ที่ประกอบด้วยป้ายชื่อรอบสอบ (labels) และข้อมูลอัตราการผ่าน (data)
 */
function getPassRateByRound() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('License_Exam_Records');
    if (!sheet || sheet.getLastRow() < 2) {
      return { labels: [], data: [] };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    const yearIndex = headers.indexOf('ExamYear');
    const sessionIndex = headers.indexOf('ExamSession');
    const subjectKeys = [
      'Subject1_Maternity', 'Subject2_Pediatric', 'Subject3_Adult', 
      'Subject4_Geriatric', 'Subject5_Psychiatric', 'Subject6_Community', 
      'Subject7_Law', 'Subject8_Surgical'
    ];
    const subjectIndices = subjectKeys.map(key => headers.indexOf(key));

    const roundStats = {}; // object สำหรับเก็บข้อมูล เช่น { "2023 รอบ 1": { passed: 0, total: 0 } }

    data.forEach(row => {
      const year = row[yearIndex];
      const session = row[sessionIndex];
      if (!year || !session) return; // ข้ามแถวที่ไม่มีข้อมูลปี/รอบ

      const key = `${year} รอบ ${session}`;
      if (!roundStats[key]) {
        roundStats[key] = { passed: 0, total: 0 };
      }

      subjectIndices.forEach(index => {
        if (index !== -1) {
          const result = row[index];
          if (result === 'ผ่าน') {
            roundStats[key].passed++;
            roundStats[key].total++;
          } else if (result === 'ไม่ผ่าน') {
            roundStats[key].total++;
          }
        }
      });
    });

    // แปลง object เป็น array, เรียงลำดับ, และคำนวณเปอร์เซ็นต์
    const sortedRounds = Object.keys(roundStats).sort((a, b) => {
      const [yearA, , sessionA] = a.split(' ');
      const [yearB, , sessionB] = b.split(' ');
      if (yearA !== yearB) return yearA - yearB;
      return sessionA - sessionB;
    });

    const labels = sortedRounds;
    const rates = sortedRounds.map(key => {
      const stats = roundStats[key];
      return stats.total > 0 ? Math.round((stats.passed / stats.total) * 100) : 0;
    });

    return { labels: labels, data: rates };

  } catch (e) {
    Logger.log("Error in getPassRateByRound: " + e.message);
    return { labels: [], data: [] };
  }
}

/**
 * คำนวณสถิติเปรียบเทียบจำนวนครั้งที่เข้าสอบ
 * @returns {object} อ็อบเจกต์ที่ประกอบด้วยจำนวนผู้ที่สอบครั้งแรก, สอบซ้ำ 1 ครั้ง, และสอบซ้ำ 2 ครั้งขึ้นไป
 */
function getExamAttemptStats() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('License_Exam_Records');
    if (!sheet || sheet.getLastRow() < 2) {
      return { firstTime: 0, retakeOnce: 0, retakeMultiple: 0 };
    }

    // ดึงข้อมูลเฉพาะคอลัมน์ Email (คอลัมน์ C) เพื่อประสิทธิภาพที่ดีขึ้น
    const emailData = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues().flat();

    const userAttempts = {}; // object สำหรับนับจำนวนครั้งที่สอบของแต่ละ email
    emailData.forEach(email => {
      if (email) {
        userAttempts[email] = (userAttempts[email] || 0) + 1;
      }
    });

    const stats = {
      firstTime: 0,
      retakeOnce: 0,
      retakeMultiple: 0
    };

    // วนลูปเพื่อจัดกลุ่มผู้ใช้ตามจำนวนครั้งที่สอบ
    for (const email in userAttempts) {
      const count = userAttempts[email];
      if (count === 1) {
        stats.firstTime++;
      } else if (count === 2) {
        stats.retakeOnce++;
      } else { // count >= 3
        stats.retakeMultiple++;
      }
    }

    return stats;
  } catch (e) {
    Logger.log("Error in getExamAttemptStats: " + e.message);
    return { firstTime: 0, retakeOnce: 0, retakeMultiple: 0 };
  }
}

/**
 * คำนวณข้อมูลสรุปสำหรับการบริจาคทั้งหมด
 * @returns {object} อ็อบเจกต์ข้อมูลสรุป 4 ส่วน: ยอดรวม, จำนวนผู้บริจาค, วัตถุประสงค์, และแนวโน้มรายเดือน
 */
function getDonationDashboardStats() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Donations_Log');
    if (!sheet || sheet.getLastRow() < 2) {
      return { totalAmount: 0, donorCount: 0, purposes: {}, monthlyTrend: { labels: [], data: [] } };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const amountIndex = headers.indexOf('Amount');
    const emailIndex = headers.indexOf('Email');
    const purposeIndex = headers.indexOf('Purpose');
    const timestampIndex = headers.indexOf('Timestamp');

    let totalAmount = 0;
    const donors = new Set();
    const purposes = {};
    const monthlyTrend = {};

    data.forEach(row => {
      const amount = parseFloat(row[amountIndex]) || 0;
      const email = row[emailIndex];
      const purpose = row[purposeIndex] || 'ไม่ระบุ';
      const timestamp = new Date(row[timestampIndex]);

      totalAmount += amount;
      if (email) donors.add(email);
      purposes[purpose] = (purposes[purpose] || 0) + 1; // นับจำนวนครั้งของการบริจาคในแต่ละวัตถุประสงค์

      if (!isNaN(timestamp.getTime())) {
        const monthKey = `${timestamp.getFullYear()}-${String(timestamp.getMonth() + 1).padStart(2, '0')}`;
        monthlyTrend[monthKey] = (monthlyTrend[monthKey] || 0) + amount;
      }
    });

    // ประมวลผลข้อมูลแนวโน้มรายเดือน
    const sortedMonths = Object.keys(monthlyTrend).sort();
    const trendLabels = sortedMonths.map(key => {
      const [year, month] = key.split('-');
      const date = new Date(year, month - 1, 1);
      return date.toLocaleDateString('th-TH', { month: 'short', year: '2-digit' }); // ผลลัพธ์เช่น "ต.ค. 68"
    });
    const trendData = sortedMonths.map(key => monthlyTrend[key]);
    
    return {
      totalAmount: totalAmount,
      donorCount: donors.size,
      purposes: purposes,
      monthlyTrend: { labels: trendLabels, data: trendData }
    };
  } catch (e) {
    Logger.log("Error in getDonationDashboardStats: " + e.message);
    return { totalAmount: 0, donorCount: 0, purposes: {}, monthlyTrend: { labels: [], data: [] } };
  }
}