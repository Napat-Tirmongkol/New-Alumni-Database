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
  EMAIL: 1, PICTURE_ID: 2, PREFIX: 3, STUDENT_ID: 4, FNAME_TH: 5, LNAME_TH: 6,
  FNAME_EN: 7, LNAME_EN: 8, GRAD_CLASS: 9, ADVISOR: 10, DOB: 11, GENDER: 12,
  BIRTH_COUNTRY: 13, NATIONALITY: 14, RACE: 15, PHONE: 16, GPAX: 17, ADDRESS: 18,
  EMERGENCY_NAME: 19, EMERGENCY_RELATION: 20, EMERGENCY_PHONE: 21, AWARDS: 22,
  FUTURE_PLAN: 23, EDUCATION_PLAN: 24, INTERNATIONAL_PLAN: 25, THAI_LICENSE: 26
};

// =================================================================
// ส่วนที่ 2: ฟังก์ชัน Setup (สำหรับสร้าง/ซ่อมแซมชีต)
// =================================================================

/**
 * @description ตรวจสอบและสร้างชีตพร้อมคอลัมน์ที่ถูกต้องทั้งหมด
 */
function setupProjectSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    
    const userDbHeaders = [
      'UserID', 'Email', 'Password', 'Role', 'FirstNameTH', 'LastNameTH', 
      'RegistrationDate', 'LastLogin', 'IsActive', 'ResetToken', 'TokenExpiry'
    ];
    const userProfileHeaders = [
      'UserID','Email', 'ProfilePictureID', 'Prefix', 'StudentID', 'FirstNameTH', 'LastNameTH', 
      'FirstNameEN', 'LastNameEN', 'GraduationClass', 'Advisor', 'DateOfBirth', 
      'Gender', 'BirthCountry', 'Nationality', 'Race', 'PhoneNumber', 'GPAX', 'CurrentAddress',
      'EmergencyContactName', 'EmergencyContactRelation', 'EmergencyContactPhone', 'Awards',
      'FutureWorkPlan', 'EducationPlan', 'InternationalWorkPlan', 'WillTakeThaiLicense'
    ];
    // --- ↓↓↓ เพิ่ม Header สำหรับชีตใหม่ ↓↓↓ ---
    const licenseExamHeaders = [
      'UserID', 'Email', 'ExamRound', 'ExamSession', 'ExamYear', 
      'Subject1_Maternity', 'Subject2_Pediatric', 'Subject3_Adult', 'Subject4_Geriatric',
      'Subject5_Psychiatric', 'Subject6_Community', 'Subject7_Law', 'Subject8_Surgical',
      'EvidenceFileID' // สำหรับเก็บ ID ไฟล์หลักฐานผลสอบ
    ];
    
    checkAndCreateSheet(spreadsheet, 'User_Database', userDbHeaders);
    checkAndCreateSheet(spreadsheet, 'User_Profiles', userProfileHeaders);
    checkAndCreateSheet(spreadsheet, 'License_Exam_Records', licenseExamHeaders);

    Browser.msgBox("ตั้งค่าชีตสำเร็จ!");
  } 
  catch (e) {
    Browser.msgBox("เกิดข้อผิดพลาด: " + e.message);
  }
}

/**
 * @description ฟังก์ชันผู้ช่วย: ตรวจสอบชีต ถ้าไม่มีจะสร้างใหม่พร้อม Header
 */
function checkAndCreateSheet(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    Logger.log(`Created sheet: ${sheetName}`);
  } else {
    Logger.log(`Sheet "${sheetName}" already exists.`);
  }
}


// =================================================================
// ส่วนที่ 3: ฟังก์ชันหลักของเว็บแอป (CORE WEB APP FUNCTIONS)
// =================================================================

/**
 * @description ทำหน้าที่เป็น "Router" หลัก คอยส่งผู้ใช้ไปยังหน้า HTML ที่ถูกต้อง
 * ตามพารามิเตอร์ 'page' ที่ส่งมากับ URL
 * @param {Object} e - อ็อบเจกต์เหตุการณ์ที่ Apps Script ส่งมาให้
 * @returns {HtmlOutput} - หน้าเว็บ HTML ที่จะแสดงผล
 */
function doGet(e) {
  const page = e ? e.parameter.page : null;
  const token = e ? e.parameter.token : null;

  // จัดการกรณีกดลิงก์รีเซ็ตรหัสผ่าน
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
  
  // แผนที่สำหรับจับคู่ 'page' parameter กับชื่อไฟล์ HTML
  const pageMappings = {
    'forgotpassword': 'ForgotPassword',
    'admindashboard': 'AdminDashboard',
    'dashboard': 'Dashboard',
    'register': 'Register'
  };

  if (page && pageMappings[page]) {
    return HtmlService.createTemplateFromFile(pageMappings[page]).evaluate().setTitle('ระบบติดตามศิษย์');
  }
  
  // ถ้าไม่มี 'page' หรือไม่ตรงกับเงื่อนไขใดๆ ให้แสดงหน้า Login
  return HtmlService.createTemplateFromFile('Login').evaluate().setTitle('ระบบติดตามศิษย์');
}

/**
 * @description ฟังก์ชันผู้ช่วยสำหรับ "ดึง" เนื้อหาจากไฟล์ HTML อื่นเข้ามาแปะในไฟล์หลัก
 * @param {string} filename - ชื่อไฟล์ HTML ที่ต้องการ (ไม่ต้องมี .html)
 * @returns {string} - เนื้อหาดิบของไฟล์ HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * @description ดึงเนื้อหาของ "หน้าย่อย" เพื่อนำไปแสดงใน Content Area ของ Dashboard
 * @param {string} pageName - ชื่อย่อของหน้าที่ต้องการ (เช่น 'profile')
 * @returns {string} - เนื้อหาดิบของไฟล์ HTML ที่เกี่ยวข้อง
 */
function getPageContent(pageName) {
  const pageMap = {
    'profile': 'Profile',
    'profileSummary': 'ProfileSummary',
    'license': 'LicenseTracking',
    'donations': 'Donations',
    'virtual_id': 'VirtualID'
  };

  if (pageMap[pageName]) {
    return include(pageMap[pageName]);
  }
  
  return '<h1>ไม่พบหน้าเว็บ</h1>';
}

/**
 * @description ฟังก์ชันผู้ช่วยสำหรับให้ Frontend เรียกใช้เพื่อขอ URL ของเว็บแอปตัวเอง
 * @returns {string} - URL ของเว็บแอปที่ Deploy ไว้
 */
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
      const email = rowData[COLS_DB.EMAIL - 1]; // <-- แก้ไขเป็น COLS_DB
      const password = String(rowData[COLS_DB.PASSWORD - 1]); // <-- แก้ไขเป็น COLS_DB
      const isActive = rowData[COLS_DB.IS_ACTIVE - 1]; // <-- แก้ไขเป็น COLS_DB
      
      if (email === credentials.email && password === credentials.password) {
        if (isActive === false) { return { success: false, message: 'บัญชีของคุณถูกระงับการใช้งาน' }; }
        
        const userRow = i + 2;
        sheet.getRange(userRow, COLS_DB.LAST_LOGIN).setValue(new Date()); // <-- แก้ไขเป็น COLS_DB
        
        const user = {
          firstName: rowData[COLS_DB.FIRST_NAME - 1], // <-- แก้ไขเป็น COLS_DB
          lastName: rowData[COLS_DB.LAST_NAME - 1], // <-- แก้ไขเป็น COLS_DB
          role: rowData[COLS_DB.ROLE - 1] // <-- แก้ไขเป็น COLS_DB
        };
        
        return { success: true, user: user }; 
      }
    }
    return { success: false, message: 'อีเมลหรือรหัสผ่านไม่ถูกต้อง' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.message };
  }
}

// --- ↓↓↓ อัปเดตฟังก์ชัน registerUser ↓↓↓ ---
function registerUser(userData) {
  try {
    const dbSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    if (!dbSheet) { throw new Error("ไม่พบชีต 'User_Database'"); }
    
    const emails = dbSheet.getRange(2, COLS_DB.EMAIL, dbSheet.getLastRow(), 1).getValues().flat();
    if (emails.includes(userData.email)) { return { success: false, message: 'อีเมลนี้ถูกใช้งานแล้ว' }; }
    
    const newUserId = "UID-" + new Date().getTime(); // สร้าง UserID ใหม่

    // 1. เพิ่มข้อมูลลงใน User_Database
    const newUserDbRow = [
      newUserId, userData.email, userData.password, userData.userType,
      userData.firstName, userData.lastName, new Date(), null, true, null, null
    ];
    dbSheet.appendRow(newUserDbRow);

    // 2. เพิ่มข้อมูลเริ่มต้นลงใน User_Profiles
    const profileSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    if (profileSheet) {
      const newUserProfileRow = [
        newUserId, userData.email, '', '', userData.firstName, userData.lastName
        // คอลัมน์ที่เหลือปล่อยให้เป็นค่าว่างไปก่อน
      ];
      profileSheet.appendRow(newUserProfileRow);
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

    const data = sheet.getRange(2, 1, sheet.getLastRow(), COLS_DB.EMAIL).getValues(); // <-- แก้ไขเป็น COLS_DB
    let userRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][COLS_DB.EMAIL - 1] === email) { // <-- แก้ไขเป็น COLS_DB
        userRow = i + 2;
        break;
      }
    }
    if (userRow === -1) { return { success: false, message: 'ไม่พบอีเมลนี้ในระบบ' }; }
    
    const token = Utilities.getUuid();
    const expiryDate = new Date(new Date().getTime() + 15 * 60 * 1000);
    sheet.getRange(userRow, COLS_DB.RESET_TOKEN).setValue(token); // <-- แก้ไขเป็น COLS_DB
    sheet.getRange(userRow, COLS_DB.TOKEN_EXPIRY).setValue(expiryDate); // <-- แก้ไขเป็น COLS_DB
    
    const resetLink = getWebAppUrl() + '?page=resetpassword&token=' + token;
    // ... (ส่วนของเนื้อหาอีเมล) ...
    MailApp.sendEmail(email, subject, body);
    return { success: true, message: 'ส่งลิงก์สำหรับรีเซ็ตรหัสผ่านไปที่อีเมลของคุณแล้ว' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.message };
  }
}

function verifyResetToken(token) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    if (!sheet) return { isValid: false };
    const data = sheet.getRange(2, COLS_DB.RESET_TOKEN, sheet.getLastRow(), 2).getValues(); // <-- แก้ไขเป็น COLS_DB
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
    sheet.getRange(verification.row, COLS_DB.PASSWORD).setValue(newPassword); // <-- แก้ไขเป็น COLS_DB
    sheet.getRange(verification.row, COLS_DB.RESET_TOKEN).setValue(null); // <-- แก้ไขเป็น COLS_DB
    sheet.getRange(verification.row, COLS_DB.TOKEN_EXPIRY).setValue(null); // <-- แก้ไขเป็น COLS_DB
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
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) throw new Error("ไม่สามารถระบุตัวตนผู้ใช้ได้");

    const dbSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Database');
    const profileSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('User_Profiles');
    
    // --- หา UserID จาก User_Database ---
    const dbData = dbSheet.getRange("A:B").getValues();
    let userId = null;
    for (let i = 0; i < dbData.length; i++) {
      if (dbData[i][1] === userEmail) { // คอลัมน์ B คือ Email
        userId = dbData[i][0]; // คอลัมน์ A คือ UserID
        break;
      }
    }
    if (!userId) throw new Error("ไม่พบ UserID สำหรับผู้ใช้นี้");

    // --- ส่วนจัดการอัปโหลดรูปภาพ ---
    if (profileData.profilePicture && profileData.mimeType) {
      Logger.log("กำลังดำเนินการอัปโหลดรูปภาพใหม่...");
      const decodedImage = Utilities.base64Decode(profileData.profilePicture);
      const blob = Utilities.newBlob(decodedImage, profileData.mimeType, `${userEmail}_profile.jpg`);
      const folder = DriveApp.getFolderById(PROFILE_PICTURE_FOLDER_ID);
      
      const oldFiles = folder.getFilesByName(`${userEmail}_profile.jpg`);
      while(oldFiles.hasNext()) {
        oldFiles.next().setTrashed(true);
        Logger.log("ลบรูปโปรไฟล์เก่าแล้ว");
      }

      const file = folder.createFile(blob);
       file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pictureFileId = file.getId();
      Logger.log("อัปโหลดรูปใหม่สำเร็จ! File ID: " + pictureFileId);
    } else {
      Logger.log("ไม่มีการอัปโหลดรูปภาพใหม่ในครั้งนี้");
    }
    
    const profileRowData = [
      userEmail, pictureFileId, profileData.prefix, profileData.studentId, profileData.firstNameTH,
      profileData.lastNameTH, profileData.firstNameEN, profileData.lastNameEN, profileData.gradClass,
      profileData.advisor, profileData.dob ? new Date(profileData.dob) : null, profileData.gender, 
      profileData.birthCountry, profileData.nationality, profileData.race, 
      profileData.phone, profileData.gpax, profileData.address,
      profileData.emergencyName, profileData.emergencyRelation, profileData.emergencyPhone,
      profileData.awards, profileData.futurePlan, profileData.educationPlan,
      profileData.internationalPlan, profileData.willTakeThaiLicense
    ];

    const data = sheet.getRange("A:A").getValues();
    let rowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === userEmail) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, profileRowData.length).setValues([profileRowData]);
    } else {
      sheet.appendRow(profileRowData);
    }
    
    Logger.log("บันทึกข้อมูลลงชีตเรียบร้อย");
    return "บันทึกข้อมูลส่วนตัวสำเร็จ!";
  } catch (e) {
    Logger.log("เกิดข้อผิดพลาดร้ายแรงใน saveUserProfile: " + e.message);
    return `เกิดข้อผิดพลาด: ${e.message}`;
  }
}

function getLicenseExamStatus(email) {
  const subjects = [
    'Subject1_Maternity', 'Subject2_Pediatric', 'Subject3_Adult', 'Subject4_Geriatric',
    'Subject5_Psychiatric', 'Subject6_Community', 'Subject7_Law', 'Subject8_Surgical'
  ];
  
  // สร้างสถานะเริ่มต้น ทุกวิชาคือยังไม่ผ่าน
  const status = {};
  subjects.forEach(subject => status[subject] = false);

  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(CONFIG.SHEETS.LICENSE_EXAMS);
    if (!sheet || sheet.getLastRow() < 2) {
      return status; // ถ้าไม่มีข้อมูลเลย ให้คืนค่าสถานะเริ่มต้น
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // ดึงหัวข้อออก

    // วนลูปดูข้อมูลการสอบทุกครั้งที่ผู้ใช้เคยบันทึกไว้
    for (const row of data) {
      const rowEmail = row[headers.indexOf('Email')];
      
      if (rowEmail === email) {
        // ตรวจสอบผลสอบของแต่ละวิชาในแถวนี้
        subjects.forEach(subjectKey => {
          const subjectIndex = headers.indexOf(subjectKey);
          // ถ้าวิชานี้เคยสอบผ่านแล้ว ให้เปลี่ยนสถานะเป็น true
          if (row[subjectIndex] === 'ผ่าน') {
            status[subjectKey] = true;
          }
        });
      }
    }
    return status;
  } catch (e) {
    Logger.log('Error in getLicenseExamStatus: ' + e.toString());
    return status; // กรณีเกิด Error ให้คืนค่าสถานะเริ่มต้น
  }
}
