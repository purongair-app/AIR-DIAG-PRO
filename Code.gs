const SPREADSHEET_ID = "1ZbfnZ3DJ2NAbusB55rU0PJNvH8xUOEowPIXqGd99vgc";
const ADMIN_PIN = "1234"; // รหัสผ่านยืนยันตัวตน (Admin PIN)

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('AIR DIAG PRO - ผู้ช่วยวิเคราะห์แอร์มืออาชีพ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. ฟังก์ชันดึงประวัติงานซ่อมจาก Google Sheets
function getRepairHistoryFromSheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("บันทึกหน้างาน");
    if (!sheet) {
      sheet = ss.getSheets()[0]; 
    }
    
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) return [];
    
    var historyList = [];
    for (var i = values.length - 1; i >= 1; i--) {
      var row = values[i];
      if (row.some(function(cell) { return cell !== "" && cell !== null; })) {
        historyList.push({
          rowIndex: i + 1,
          date: row[0] !== undefined && row[0] !== "" ? String(row[0]) : "-",
          customer: row[1] !== undefined && row[1] !== "" ? String(row[1]) : "-",
          brand: row[2] !== undefined && row[2] !== "" ? String(row[2]) : "-",
          btu: row[3] !== undefined && row[3] !== "" ? String(row[3]) : "-",
          symptom: row[4] !== undefined && row[4] !== "" ? String(row[4]) : "-",
          model: row[5] !== undefined && row[5] !== "" ? String(row[5]) : "-",
          ref: row[6] !== undefined && row[6] !== "" ? String(row[6]) : "-",
          psi: row[7] !== undefined && row[7] !== "" ? String(row[7]) : "-",
          amp: row[8] !== undefined && row[8] !== "" ? String(row[8]) : "-",
          volt: row[9] !== undefined && row[9] !== "" ? String(row[9]) : "-",
          tinFront: row[10] !== undefined && row[10] !== "" ? String(row[10]) : "-",
          tinBack: row[11] !== undefined && row[11] !== "" ? String(row[11]) : "-",
          toutFront: row[12] !== undefined && row[12] !== "" ? String(row[12]) : "-",
          toutBack: row[13] !== undefined && row[13] !== "" ? String(row[13]) : "-",
          techName: row[14] !== undefined && row[14] !== "" ? String(row[14]) : "-",
          techPhone: row[15] !== undefined && row[15] !== "" ? String(row[15]) : "-",
          notes: row[16] !== undefined && row[16] !== "" ? String(row[16]) : "-"
        });
      }
    }
    return historyList;
  } catch (error) {
    return [];
  }
}

// 2. ฟังก์ชันดึงยี่ห้อแอร์จากชีท Error Code Database
function getAirBrands() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Error Code Database");
    
    if (!sheet) {
      return ["Daikin", "Mitsubishi Electric", "Carrier", "Panasonic", "Midea", "Haier", "LG", "Samsung"];
    }
    
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) {
      return ["Daikin", "Mitsubishi Electric", "Carrier", "Panasonic", "Midea", "Haier", "LG", "Samsung"];
    }
    
    var brandSet = new Set();
    for (var i = 1; i < values.length; i++) {
      var brand = String(values[i][0]).trim();
      if (brand !== "") {
        brandSet.add(brand);
      }
    }
    
    var brandsArray = Array.from(brandSet);
    return brandsArray.length > 0 ? brandsArray : ["Daikin", "Mitsubishi Electric", "Carrier", "Panasonic", "Midea", "Haier", "LG", "Samsung"];
  } catch (error) {
    return ["Daikin", "Mitsubishi Electric", "Carrier", "Panasonic", "Midea", "Haier", "LG", "Samsung"];
  }
}

// 3. ฟังก์ชันลบแถวข้อมูลประวัติงานซ่อม
function deleteRepairHistoryRow(pin, rowIndex) {
  try {
    if (pin !== ADMIN_PIN) {
      throw new Error('รหัสผ่าน PIN ไม่ถูกต้อง! ไม่อนุญาตให้ลบข้อมูล');
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("บันทึกหน้างาน");
    if (!sheet) {
      sheet = ss.getSheets()[0];
    }
    
    if (sheet) {
       sheet.deleteRow(Number(rowIndex));
    }
    return 'ลบข้อมูลประวัติสำเร็จ!';
  } catch (error) {
    throw new Error('ลบไม่สำเร็จ: ' + error.message);
  }
}

// 4. ฟังก์ชันบันทึกข้อมูลหน้างานขึ้น Google Sheets
function saveDataToSheet(record) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("บันทึกหน้างาน");
    
    if (!sheet) {
      sheet = ss.insertSheet("บันทึกหน้างาน");
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "วันที่/เวลา", "ชื่อลูกค้า/สถานที่", "ยี่ห้อ", "ขนาด BTU", "อาการเสีย/โค้ด", 
        "รุ่น (Model)", "ชนิดน้ำยา", "แรงดันน้ำยา (PSI)", "กระแสไฟ (Amp)", "แรงดันไฟฟ้า (Volt)", 
        "อุณหภูมิหน้าคอยล์เย็น", "อุณหภูมิหลังคอยล์เย็น", "อุณหภูมิหน้าคอยล์ร้อน", "อุณหภูมิหลังคอยล์ร้อน", 
        "ชื่อผู้บันทึก", "เบอร์โทร", "บันทึกเพิ่มเติม"
      ]);
    }

    var rowData = [
      record.date, record.customer, record.brand, record.btu, record.symptom,
      record.model, record.ref, record.psi, record.amp, record.volt,
      record.tinFront, record.tinBack, record.toutFront, record.toutBack,
      record.techName, record.techPhone, record.notes
    ];
    
    sheet.appendRow(rowData);
    return 'บันทึกข้อมูลขึ้นระบบ Cloud สำเร็จ!';
  } catch (error) {
    throw new Error('บันทึกไม่สำเร็จ: ' + error.message);
  }
}

// 5. ฟังก์ชันค้นหา Error Code จากชีท
function searchErrorCodeFromSheet(brandName, errorCode) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Error Code Database");
    
    if (!sheet) return null;
    
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    brandName = String(brandName).trim().toLowerCase();
    errorCode = String(errorCode).trim().toUpperCase();
    
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var sheetBrand = String(row[0]).trim().toLowerCase();
      var sheetCode = String(row[1]).trim().toUpperCase();
      
      if (sheetBrand === brandName && sheetCode === errorCode) {
        return {
          title: row[2] || "ไม่มีข้อมูลคำอธิบาย",
          points: row[3] ? row[3].split(",") : ["ตรวจสอบตามคู่มือศูนย์บริการ"],
          steps: row[4] ? row[4].split(",") : ["1. ตรวจสอบระบบเบื้องต้น"],
          doDont: "⚠️ โปรดใช้ความระมัดระวังในการตรวจสอบระบบไฟฟ้าและอุปกรณ์"
        };
      }
    }
    return null;
  } catch (error) {
    return null;
  }
}

// 6. ฟังก์ชันอัปโหลด Error Code
function uploadErrorCodeFile(pin, rowsData) {
  try {
    if (pin !== ADMIN_PIN) {
      throw new Error('รหัสผ่าน PIN ไม่ถูกต้อง!');
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Error Code Database");
    
    if (!sheet) {
      sheet = ss.insertSheet("Error Code Database");
      sheet.appendRow(["ยี่ห้อแอร์", "รหัสข้อผิดพลาด", "ความหมาย / อาการ", "จุดที่ต้องตรวจเช็ก", "วิธีแก้ไขเบื้องต้น"]);
    }

    for (var i = 0; i < rowsData.length; i++) {
      var row = rowsData[i];
      if (row.length >= 5 && row[0] !== "") {
        sheet.appendRow([row[0], row[1], row[2], row[3], row[4]]);
      }
    }
    removeDuplicateErrors();
    return 'อัปโหลดสำเร็จ';
  } catch (error) {
    throw new Error('อัปโหลดไม่สำเร็จ: ' + error.message);
  }
}
// ฟังก์ชันดึงรายการอาการเสียจากชีท "วิเคราะห์อาการเสีย" มาใส่ Dropdown (พร้อมกรองข้อความซ้ำอัตโนมัติ)
function getDiagnosisSymptoms() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("วิเคราะห์อาการเสีย");
    
    if (!sheet) {
      return [
        { id: "1", name: "แอร์ไม่เย็น" },
        { id: "2", name: "แอร์ตัดบ่อย" }
      ];
    }
    
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length <= 1) {
      return [];
    }
    
    var symptomList = [];
    var seenCategories = new Set(); // ตัวกรองข้อความซ้ำ
    
    for (var i = 1; i < values.length; i++) {
      var category = String(values[i][0]).trim();
      
      // ตรวจสอบว่ามีค่าและยังไม่เคยถูกเพิ่มเข้ามาในรายการ
      if (category !== "" && !seenCategories.has(category)) {
        seenCategories.add(category);
        symptomList.push({
          id: String(symptomList.length + 1),
          name: category
        });
      }
    }
    
    return symptomList;
  } catch (error) {
    return [];
  }
}
// 7. ฟังก์ชันอัปโหลดเอกสารวิเคราะห์อาการเสีย
function uploadDiagnosisDocFile(pin, rowsData) {
  try {
    if (pin !== ADMIN_PIN) {
      throw new Error('รหัสผ่าน PIN ไม่ถูกต้อง!');
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("วิเคราะห์อาการเสีย");
    
    if (!sheet) {
      sheet = ss.insertSheet("วิเคราะห์อาการเสีย");
      sheet.appendRow(["กลุ่มอาการ (Category)", "หัวข้อหลัก (Topic/Step)", "จุดที่ต้องตรวจสอบ / วัดค่า (Checkpoints)", "คำแนะนำ / วิธีแก้ไข (Guidelines)", "ข้อควรระวัง (Cautions)"]);
    }

    for (var i = 0; i < rowsData.length; i++) {
      var row = rowsData[i];
      if (row.length >= 5 && row[0] !== "") {
        sheet.appendRow([row[0], row[1], row[2], row[3], row[4]]);
      }
    }
    return 'อัปโหลดเอกสารวิเคราะห์อาการเสียสำเร็จ';
  } catch (error) {
    throw new Error('อัปโหลดไม่สำเร็จ: ' + error.message);
  }
}

// 8. ลบข้อมูลซ้ำ Error Code
function removeDuplicateErrors() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Error Code Database");
  
  if (!sheet) return;
  
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  if (values.length <= 1) return;
  
  var uniqueRows = [];
  var seen = new Set();
  var header = values[0]; 
  uniqueRows.push(header);
  
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var brand = String(row[0]).trim().toLowerCase();
    var code = String(row[1]).trim().toUpperCase();
    var key = brand + "_" + code;
    
    if (key !== "_" && !seen.has(key)) {
      seen.add(key);
      uniqueRows.push(row);
    }
  }
  
  sheet.clearContents();
  sheet.getRange(1, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
}
