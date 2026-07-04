// กำหนดค่าเริ่มต้นของระบบเชื่อมต่อ Google Sheet ID ที่ระบุไว้
const TARGET_SHEET_ID = "16ZdZtakZGis9mMZp7cCCVQ28WNJ_yvm_SyJt8fUTQfM";
const SHEET_NAME = "บันทึกข้อมูลรื้อถอน";

/**
 * ฟังก์ชัน doGet สำหรับแสดงผลหน้าเว็บสวยงามแบบ Full Web App
 */
function doGet(e) {
  // ตรวจสอบสถานะการเชื่อมต่อแบบด่วน
  if (e.parameter.action === 'test') {
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  // โหลดดึงข้อมูลจากตาราง
  if (e.parameter.action === 'read') {
    var data = readDataFromSheet();
    return ContentService.createTextOutput(JSON.stringify(data))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  // ส่งออกหน้า UI
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ระบบบันทึกข้อมูลสถานะที่ราชพัสดุและพื้นที่ป่าไม้ฯ แม่ฮ่องสอน เขต 1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ฟังก์ชัน POST สำหรับรับค่า JSON จาก Web App มาจัดเก็บลงใน แผ่นงาน
 */
function doPost(e) {
  try {
    var jsonString = e.postData.contents;
    var data = JSON.parse(jsonString);
    
    // ตรวจสอบและดึงชีตอัตโนมัติ (จะสร้างทันทีถ้าไม่มีชีตนี้อยู่)
    var sheet = getOrCreateTargetSheet();
    
    // จัดรูปแบบเวลาประทับ (Timestamp) ของประเทศไทน
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    
    // บันทึกแถวข้อมูลเรียงตามโครงสร้างข้อมูล
    sheet.appendRow([
      timestamp,
      data.school,
      data.parcelNo,
      data.buildings,
      data.status,
      data.letterNo,
      data.letterDate,
      data.office,
      data.admin
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data saved successfully" }))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันการค้นหาความถูกต้องของโครงสร้างชีต หากยังไม่มีจะทำการอินิทชีตและใส่หัวข้อให้อัตโนมัติ
 */
function getOrCreateTargetSheet() {
  var ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    // ถ้ายังไม่มีหน้าบันทึกข้อมูลดังกล่าว ระบบจะทำการสร้างหน้าชีตนี้ขึ้นมาให้อัตโนมัติทันที
    sheet = ss.insertSheet(SHEET_NAME);
    
    // สร้างหัวตารางข้อกำหนด Column Header
    var headers = [
      "วันเวลาทำรายการ", 
      "ชื่อโรงเรียน", 
      "แปลงหมายเลขที่ มส.", 
      "ลำดับอาคารที่ขอรื้อถอน", 
      "สถานะการส่งหนังสือมาที่เขต", 
      "เลขที่หนังสือศธ.", 
      "วันที่ลงนามในหนังสือเอกสาร",
      "สำนักงานเขตพื้นที่การศึกษา",
      "แอดมินผู้ลงบันทึก"
    ];
    
    sheet.appendRow(headers);
    
    // ตกแต่งหัวตารางของชีตแบบสวยงามให้อัตโนมัติด้วย
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#064e3b") // สีเขียวเข้มสอดคล้องธีมเว็บแอป
               .setFontColor("#ffffff")
               .setFontWeight("bold")
               .setHorizontalAlignment("center");
               
    sheet.setFrozenRows(1); // ตรึงแถวแรกไว้
  }
  return sheet;
}

/**
 * ฟังก์ชันอ่านข้อมูลทั้งหมดจากชีตเพื่อส่งกลับไปเรนเดอร์ในตาราง Data Table บนหน้าเว็บ
 */
function readDataFromSheet() {
  var sheet = getOrCreateTargetSheet();
  var rows = sheet.getDataRange().getValues();
  
  // หากพบเฉพาะหัวข้อตาราง (ยังไม่มีการบันทึกข้อมูลใดๆ)
  if (rows.length <= 1) {
    return [];
  }
  
  var dataResult = [];
  
  // วนลูปเริ่มจากแถวที่ 2 เป็นต้นไป (ไม่นับ Header แถวแรก)
  // ใช้การแสดงผลแบบถอยหลัง (นับจากอันล่าสุดขึ้นมาแสดงด้านบนสุดในตาราง)
  for (var i = rows.length - 1; i >= 1; i--) {
    var row = rows[i];
    dataResult.push({
      timestamp: row[0],
      school: row[1],
      parcelNo: row[2],
      buildings: row[3],
      status: row[4],
      letterNo: row[5],
      letterDate: row[6],
      office: row[7],
      admin: row[8]
    });
  }
  
  return dataResult;
}
