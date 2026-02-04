const CONFIG = {
  SPREADSHEET_ID: '1dTu9T9xBPKMvY2DFfuHPkYR1Bgl19My7QPVbHXqW4Aw', // ตรวจสอบ ID Sheet ของท่านให้ถูกต้อง
  SHEET_NAME_DATA: 'ข้อมูลทั้งหมด',
  SHEET_NAME_PLANS_SPECIAL: 'แผนการเรียนห้องเรียนพิเศษ',
  SHEET_NAME_PLANS_GENERAL: 'แผนการเรียนห้องเรียนปกติ',
  SHEET_NAME_ADDRESS: 'อ้างอิงที่อยู่',
  SHEET_NAME_ADMIN: 'แอดมิน',
  FOLDER_ID_PHOTO: '1xfbsjSx_o6jwVqG6ypjFoYTwdSoRgvAU',     // ตรวจสอบ ID Folder รูป
  FOLDER_ID_TRANSCRIPT: '1IUh9NAE64cPGCq0MluT8oqOdD2ntvR9G', // ตรวจสอบ ID Folder ปพ.1
  FOLDER_ID_CONDUCT: '1iysw3WTrUr6NH2T3teyf56wpG1-xH_RZ' // ตรวจสอบ ID Folder ปพ.1
};


function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบรับสมัครนักเรียน 2569 - โรงเรียนอรัญประเทศ') // แก้ปี 2569
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function getSheet(name) {
  try { return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(name); } catch(e) { throw new Error("ไม่พบแผ่นงาน: " + name); }
}


function formatDate(d) {
  if (!d) return "";
  return (d instanceof Date) ? Utilities.formatDate(d, "GMT+7", "yyyy-MM-dd") : String(d);
}


function getStudyPlans(level, type) {
  try {
    const sheetName = (type.includes('พิเศษ')) ? CONFIG.SHEET_NAME_PLANS_SPECIAL : CONFIG.SHEET_NAME_PLANS_GENERAL;
    const sheet = getSheet(sheetName);
    const col = (level === 'ชั้นมัธยมศึกษาปีที่ 1') ? 1 : 2;
    return sheet.getRange(2, col, sheet.getLastRow()-1, 1).getValues().flat().filter(String);
  } catch(e) { return []; }
}


function getAddressData() {
  try { return getSheet(CONFIG.SHEET_NAME_ADDRESS).getRange(2, 1, getSheet(CONFIG.SHEET_NAME_ADDRESS).getLastRow()-1, 4).getValues(); } catch(e) { return []; }
}


function adminLogin(u, p) {
  try {
    const sheet = getSheet(CONFIG.SHEET_NAME_ADMIN);
    const data = sheet.getDataRange().getDisplayValues();
    for(let i=1; i<data.length; i++) {
      let dbUser = String(data[i][0]).trim();
      let dbPass = String(data[i][1]).trim();
      if(dbUser !== "" && dbUser === String(u).trim() && dbPass === String(p).trim()) {
        return { success: true, role: parseInt(data[i][2]) || 1, name: data[i][3] || "เจ้าหน้าที่" };
      }
    }
    return { success: false, message: "Username หรือ Password ไม่ถูกต้อง" };
  } catch(e) { return { success: false, message: "Server Error: " + e.message }; }
}

function submitApplication(fd) {
  var status = getRecruitStatus();
  var type = fd.applyType; // รับค่าประเภทการสมัครจากฟอร์ม (special/general)

  // ตรวจสอบห้องเรียนพิเศษ
  if (type === 'special' && !status.special) {
    throw new Error("ขออภัย ระบบรับสมัคร 'ห้องเรียนพิเศษ' ปิดทำการแล้ว");
  } 
  // ตรวจสอบห้องเรียนทั่วไป
  else if (type === 'general' && !status.general) {
    throw new Error("ขออภัย ระบบรับสมัคร 'ห้องเรียนทั่วไป' ปิดทำการแล้ว");
  }
  
  
 const lock = LockService.getScriptLock();
  // พยายามล็อกระบบเป็นเวลา 10 วินาที หากคนใช้เยอะมากจะรอคิวตรงนี้
  if(!lock.tryLock(10000)) throw new Error("ระบบกำลังประมวลผลในส่วนอื่นอยู่ กรุณารอ 10 วินาทีแล้วลองใหม่อีกครั้ง");
 
  try {
    const sheet = getSheet(CONFIG.SHEET_NAME_DATA);
    let rowIndex = null;
    let appId = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMddHHmmss");
    let timestamp = new Date();
    
    // --- [ปรับปรุง 1] ล้างค่าเลขบัตรประชาชน (ตัดช่องว่าง, เครื่องหมาย ' และทำเป็นตัวพิมพ์ใหญ่) ---
    // .toUpperCase() ช่วยให้ 'g' กลายเป็น 'G' เสมอ
    const idCardClean = String(fd.idCard).replace(/'/g, '').trim().toUpperCase();

    // --- [ปรับปรุง 2] เพิ่มการตรวจสอบเงื่อนไข SMTE ฝั่ง Server ---
    const strictPlans = [
       "SMTE (วิทย์-คณิต-เทคโนโลยี-สิ่งแวดล้อม)",
       "ห้องเรียนพิเศษวิทยาศาสตร์ คณิตศาสตร์ เทคโนโลยีและสิ่งแวดล้อม (SMTE)"
    ];

    // ถ้าแผนที่เลือกอยู่ในกลุ่มบังคับ และ พบตัวอักษรที่ไม่ใช่ตัวเลข (0-9)
    // /[^0-9]/ จะเป็น true ถ้าเจอตัวอักษรอื่นที่ไม่ใช่เลข
    if (strictPlans.includes(fd.plan) && /[^0-9]/.test(idCardClean)) {
        throw new Error("แผนการเรียน " + fd.plan + " อนุญาตให้ใช้เฉพาะเลขบัตรประชาชน (ตัวเลขล้วน) เท่านั้น");
    }
    // --- [ส่วนที่ 1] ตรวจสอบการสมัครซ้ำ ---
    // จะตรวจสอบเฉพาะกรณี "สมัครใหม่" เท่านั้น (isEditMode เป็น false)
    if (!fd.isEditMode) {
      const allData = sheet.getDataRange().getValues();
      // คอลัมน์เลขบัตรประชาชนนักเรียนคือ Index 13 (คอลัมน์ N)
      const isDuplicate = allData.some(row => String(row[13]).replace(/'/g, '').trim() === idCardClean);
      
      if (isDuplicate) {
        throw new Error("เลขบัตรประชาชนนี้ (" + idCardClean + ") ได้ลงทะเบียนในระบบเรียบร้อยแล้ว ไม่สามารถสมัครซ้ำได้");
      }
    }

    // --- [ส่วนที่ 2] กรณีโหมดแก้ไขข้อมูลเดิม ---
    if(fd.isEditMode && fd.editIdCard) {
       const data = sheet.getDataRange().getValues();
       // ค้นหาจากล่างขึ้นบนเพื่อความรวดเร็ว
       for(let i=data.length-1; i>=1; i--) {
         if(String(data[i][13]).replace(/'/g,'') === String(fd.editIdCard)) {
            // อนุญาตให้แก้เฉพาะกรณีสถานะเป็น 'ให้ปรับปรุงข้อมูล' เท่านั้น
            if(data[i][3] !== 'ให้ปรับปรุงข้อมูล') throw new Error("สถานะปัจจุบันของคุณคือ '" + data[i][3] + "' จึงไม่อนุญาตให้แก้ไขข้อมูลผ่านระบบ");
            rowIndex = i+1; 
            appId = data[i][1]; // ใช้ appId เดิม
            break;
         }
       }
       if(!rowIndex) throw new Error("ไม่พบข้อมูลเดิมในระบบที่ต้องการแก้ไข");
    }

    // --- [ส่วนที่ 3] จัดการไฟล์แนบ ---
    let photoUrl = rowIndex ? sheet.getRange(rowIndex, 39).getValue() : "-";
    let transUrl = rowIndex ? sheet.getRange(rowIndex, 40).getValue() : "-";
    let conductUrl = rowIndex ? sheet.getRange(rowIndex, 41).getValue() : "-"; // <<< เพิ่มบรรทัดนี้ (ใบรับรองความประพฤติ อยู่คอลัมน์ 41)
   
    // อัปโหลดไฟล์ใหม่เฉพาะเมื่อมีการเลือกไฟล์เข้ามา
    if(fd.photoFile && fd.photoFile.data) photoUrl = uploadFile(fd.photoFile, CONFIG.FOLDER_ID_PHOTO, appId+"_Photo");
    if(fd.transcriptFile && fd.transcriptFile.data) transUrl = uploadFile(fd.transcriptFile, CONFIG.FOLDER_ID_TRANSCRIPT, appId+"_Transcript");
    if(fd.conductFile && fd.conductFile.data) { // <<< เพิ่มบล็อกนี้
      conductUrl = uploadFile(fd.conductFile, CONFIG.FOLDER_ID_CONDUCT, appId+"_Conduct"); 
    }

    // --- [ส่วนที่ 4] เตรียมข้อมูลสำหรับบันทึก ---
    const addr = `${fd.addrNo} หมู่ ${fd.addrMoo} ซอย ${fd.addrSoi} ถนน ${fd.addrRoad} จ.${fd.province} อ.${fd.district} ต.${fd.subdistrict} ${fd.zipcode}`;
    const status = rowIndex ? "รอตรวจสอบ (แก้ไขแล้ว)" : "รอตรวจสอบ";
    const f = fd.father || {}; 
    const m = fd.mother || {}; 
    const g = fd.guardian || {};

    // ข้อมูลที่อยู่แบบ JSON สำหรับใช้โหลดกลับมาแก้ไข
    const rawAddr = JSON.stringify({
        no: fd.addrNo, moo: fd.addrMoo, soi: fd.addrSoi, road: fd.addrRoad,
        prov: fd.province, dist: fd.district, sub: fd.subdistrict, zip: fd.zipcode
    });

    const rowData = [
      fd.applyType, status, "", fd.level, fd.plan,
      fd.prefix, fd.firstname, fd.lastname, "'"+fd.dob, fd.nationality, fd.religion, "'"+fd.idCard, "'"+fd.phone,
      addr, fd.famStatus,
      f.prefix, f.name, f.lname, f.job, f.age, "'"+f.phone, f.addr,
      m.prefix, m.name, m.lname, m.job, m.age, "'"+m.phone, m.addr,
      g.prefix, g.name, g.lname, g.rel, g.job, g.age, "'"+g.phone, g.addr, 
      photoUrl, transUrl, conductUrl,
      rawAddr 
    ];

    // --- [ส่วนที่ 5] ทำการบันทึกข้อมูล ---
    if(rowIndex) {
      // แก้ไขที่แถวเดิม (เริ่มบันทึกที่คอลัมน์ 3 เพราะคอลัมน์ 1,2 คือ timestamp และ appId ซึ่งไม่เปลี่ยน)
      sheet.getRange(rowIndex, 3, 1, rowData.length).setValues([rowData]);
      return { success: true, message: "อัปเดตข้อมูลการสมัครเรียบร้อยแล้ว", appId: appId };
    } else {
      // สมัครใหม่ บันทึกแถวใหม่ต่อท้าย
      sheet.appendRow([timestamp, appId, ...rowData]);
      return { success: true, message: "ส่งใบสมัครเรียบร้อยแล้ว เลขที่อ้างอิง: " + appId, appId: appId };
    }

  } catch(e) { 
    // ส่งข้อความ Error กลับไปแสดงที่ Swal หน้าบ้าน
    return { success: false, message: e.message }; 
  } finally { 
    // สำคัญมาก: ปลดล็อกระบบเพื่อให้คนถัดไปใช้งานได้
    lock.releaseLock(); 
  }
}


function checkStatus(idCard) {
  const sheet = getSheet(CONFIG.SHEET_NAME_DATA);
  const data = sheet.getDataRange().getValues();
  
  // ค้นหาจากล่างขึ้นบน
  for(let i=data.length-1; i>=1; i--) {
     if(String(data[i][13]).replace(/'/g,'').trim() === String(idCard).trim()) {
        const safeData = data[i].map(c => (c instanceof Date) ? formatDate(c) : String(c));
        
        return {
           found: true, 
           name: data[i][8]+" "+data[i][9], 
           status: data[i][3], 
           reason: data[i][4], 
           applyType: data[i][2],
           seatNo: data[i][43], // <--- เพิ่มบรรทัดนี้ (Index 43 คือ Column 44/AR)
           fullData: (data[i][3]==='ให้ปรับปรุงข้อมูล') ? safeData : null
        };
     }
  }
  return { found: false };
}


function getAdminData() {
  try {
    // 1. เชื่อมต่อ Sheet
    var sheet = getSheet(CONFIG.SHEET_NAME_DATA); 
    
    // ดึงข้อมูลทั้งหมดเป็น Text
    var data = sheet.getDataRange().getDisplayValues();
    
    if (data.length <= 1) {
      return { 
        success: true, 
        students: [], 
        stats: { total:0, approved:0, pending:0, rejected:0 } 
      };
    }

    data.shift(); // ตัดหัวตารางออก

    // 2. แปลงข้อมูล (จุดสำคัญอยู่ตรงนี้!)
    var students = data.map(function(row, i) {
      
      return {
        rowIndex: i + 2,
        timestamp: row[0],
        appId: row[1],
        status: row[3],
        level: row[5],
        plan: row[6],
        name: row[7] + row[8] + " " + row[9], 
        idCard: String(row[13]).replace(/'/g, ''),
        
        phone: String(row[14]).replace(/'/g, ''),
      
        photo: row[39],      
        transcript: row[40], 
        conduct: row[41],
        // ----------------------------------------

        fullData: row
      };
    }).reverse();

    // 3. คำนวณสถิติ
    var stats = {
      total: students.length,
      approved: students.filter(function(s) { return s.status === 'อนุมัติ'; }).length,
      pending: students.filter(function(s) { return s.status && s.status.includes('รอตรวจสอบ'); }).length,
      rejected: students.filter(function(s) { return s.status && s.status.includes('ไม่ผ่าน'); }).length
    };

    return { success: true, students: students, stats: stats };

  } catch (e) {
    Logger.log("Error getAdminData: " + e.toString());
    return { success: false, message: "เกิดข้อผิดพลาดที่ Server: " + e.toString() };
  }
}


function updateStudentStatus(ri, st, re, by) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      // --- [จุดสำคัญ] เรียกใช้ไฟล์จาก ID ใน CONFIG (แก้ปัญหา Error: null) ---
      const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); 
      // -------------------------------------------------------------

      const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_DATA);
      if (!sheet) throw new Error("ไม่พบแผ่นชีต: " + CONFIG.SHEET_NAME_DATA);

      // 1. บันทึกสถานะลงแผ่นหลัก
      sheet.getRange(ri, 4).setValue(st); 
      sheet.getRange(ri, 5).setValue(re + " (" + by + ")");
      SpreadsheetApp.flush(); // บันทึกทันที

      // 2. ระบบรันเลขที่นั่งสอบ (ทำงานเฉพาะเมื่อกด 'อนุมัติ')
      if (st === 'อนุมัติ') {
         const COL_EXAM_ID = 44; // คอลัมน์ AR (เลขที่นั่งสอบ)
         
         // ตรวจสอบและสร้างคอลัมน์เพิ่มอัตโนมัติ (ถ้าไม่พอ)
         if (sheet.getMaxColumns() < COL_EXAM_ID) {
           sheet.insertColumnsAfter(sheet.getMaxColumns(), COL_EXAM_ID - sheet.getMaxColumns());
         }

         const rowValues = sheet.getRange(ri, 1, 1, sheet.getLastColumn()).getValues()[0];
         const currentExamId = (rowValues.length >= COL_EXAM_ID) ? rowValues[COL_EXAM_ID - 1] : "";

         // ถ้ายังไม่มีเลขสอบ ให้สร้างใหม่
         if (!currentExamId) {
            // ดึงข้อมูล (Index: 0=Time, 1=AppID, ... 5=Level, 6=Plan)
            const appId = rowValues[1];
            const sLevel = String(rowValues[5] || ""); 
            const sPlan = String(rowValues[6] || "");
            const sPrefix = rowValues[7] || "";
            const sName = rowValues[8] || "";
            const sLname = rowValues[9] || "";

            let codePrefix = "";
            let targetSheetName = "";

            // ตรวจสอบเงื่อนไข (เพิ่มคำค้นหาภาษาไทยให้แล้ว)
            if (sLevel.includes("1")) {
               targetSheetName = "เลขที่นั่งสอบห้องเรียนพิเศษ ม.1";
               if (sPlan.includes("SMTE") || sPlan.includes("วิทย์") || sPlan.includes("คณิต")) codePrefix = "11";
               else if (sPlan.includes("IEP") || sPlan.includes("อังกฤษ")) codePrefix = "12";
            } else if (sLevel.includes("4")) {
               targetSheetName = "เลขที่นั่งสอบห้องเรียนพิเศษ ม.4";
               if (sPlan.includes("SMTE") || sPlan.includes("วิทย์") || sPlan.includes("คณิต")) codePrefix = "41";
               else if (sPlan.includes("IEP") || sPlan.includes("อังกฤษ")) codePrefix = "42";
            }

            // ถ้าเข้าเงื่อนไข ให้ดำเนินการบันทึก
            if (codePrefix && targetSheetName) {
               const targetSheet = ss.getSheetByName(targetSheetName);
               if (!targetSheet) throw new Error("ไม่พบแผ่นชีตชื่อ: " + targetSheetName);

               // หาเลขลำดับล่าสุด
               const allExamData = targetSheet.getDataRange().getValues();
               let maxNum = 0;
               for (let i = 1; i < allExamData.length; i++) {
                  let eid = String(allExamData[i][0] || "");
                  if (eid.startsWith(codePrefix)) {
                     let num = parseInt(eid.substring(2)) || 0;
                     if (num > maxNum) maxNum = num;
                  }
               }
               
               // สร้างเลขใหม่
               let newId = codePrefix + String(maxNum + 1).padStart(3, '0');
               
               // บันทึกลงแผ่นหลัก
               sheet.getRange(ri, COL_EXAM_ID).setValue(newId);
               // บันทึกลงแผ่นแยก (ตามระดับชั้น)
               targetSheet.appendRow([newId, appId, sPrefix, sName, sLname, sPlan]);
               
               return { success: true, message: "บันทึกและออกเลขสอบ " + newId + " เรียบร้อย" };
            }
         }
      }
      return { success: true, message: "บันทึกสถานะเรียบร้อย (ไม่ได้ออกเลขสอบ)" };

    } catch (e) {
      throw new Error("Error: " + e.message);
    } finally {
      lock.releaseLock();
    }
  }
  throw new Error("ระบบทำงานหนัก กรุณาลองใหม่");
}

function uploadFile(d, fid, fname) {
  try {
    var folder = DriveApp.getFolderById(fid);
    var blob = Utilities.newBlob(Utilities.base64Decode(d.data), d.mimeType, fname);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/uc?export=view&id=" + file.getId();
  } catch(e) {
    return "Error Uploading";
  }
}

// --- ส่วนที่เพิ่มใหม่สำหรับหน้าตรวจสอบรายชื่อ ---
function getPublicReport() {
  try {
    const sheet = getSheet(CONFIG.SHEET_NAME_DATA);
    // ดึงข้อมูลทั้งหมด
    const data = sheet.getDataRange().getDisplayValues();
    
    // ถ้ามีแต่หัวข้อ (ข้อมูลน้อยกว่า 2 บรรทัด) ให้ส่งค่าว่างกลับไป
    if (data.length < 2) return { list: [] };

    let list = [];
    
    // เริ่มวนลูปตั้งแต่แถวที่ 2 (Index 1) เป็นต้นไป
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Index 5 = ระดับชั้น (Col F), Index 6 = แผน (Col G), Index 8 = ชื่อ (Col I)
      if (!row[8]) continue; // ถ้าไม่มีชื่อให้ข้าม

      list.push({
        level: row[5] || "-",       
        plan: row[6] || "-",        
        fullname: `${row[7] || ''}${row[8]} ${row[9]}`, // คำนำหน้า+ชื่อ+สกุล
        status: row[3] || "รอตรวจสอบ" // สถานะ (Col D)
      });
    }

    return { list: list };

  } catch (e) {
    return { error: "เกิดข้อผิดพลาด: " + e.message };
  }
}

// --- ส่วนจัดการเปิด-ปิดระบบรับสมัคร ---

// ฟังก์ชันสำหรับแอดมินกดเปิด/ปิด
function setRecruitStatus(type, isOpen) {
  var props = PropertiesService.getScriptProperties();
  // ตั้งชื่อตัวแปรแยกกัน: STATUS_SPECIAL และ STATUS_GENERAL
  var key = (type === 'special') ? 'STATUS_SPECIAL' : 'STATUS_GENERAL';
  props.setProperty(key, isOpen ? 'true' : 'false');
  return { success: true };
}

// ฟังก์ชันดึงสถานะปัจจุบัน
function getRecruitStatus() {
  var props = PropertiesService.getScriptProperties();
  return {
    // ถ้าไม่เคยตั้งค่า (null) ให้ถือว่าเปิด (true) เป็นค่าเริ่มต้น
    special: props.getProperty('STATUS_SPECIAL') !== 'false', 
    general: props.getProperty('STATUS_GENERAL') !== 'false'
  };
}

// --- เพิ่มฟังก์ชันเช็คเลขบัตรซ้ำ (สำหรับเรียกตรวจสอบทันที) ---
function checkDuplicateID(idCard) {
  const sheet = getSheet(CONFIG.SHEET_NAME_DATA);
  const data = sheet.getDataRange().getValues();
  
  // วนลูปเช็ค (สมมติว่าเลขบัตรอยู่คอลัมน์ N หรือ Index 13)
  // ตัดแถวหัวตารางออก และเช็คเฉพาะคนที่สถานะไม่ใช่ 'ยกเลิก' หรืออื่นๆ ตามต้องการ
  const isDuplicate = data.slice(1).some(row => String(row[13]).replace(/'/g, '').trim() === String(idCard));
  
  return isDuplicate; // ส่งค่า true (ซ้ำ) หรือ false (ไม่ซ้ำ) กลับไป
}
