// --- Configuration ---
// ระบุ ID ของ Google Sheet ที่ต้องการให้สคริปต์นี้ทำงานด้วย
const SPREADSHEET_ID = '1Wka6GisqNZhwh97mRo8-dG1qLM2mCvFhpJNGB7zcL3c';

// ชื่อชีตสำหรับเก็บข้อมูล
const PAYEES_SHEET_NAME = 'Payees'; // ชื่อชีตสำหรับเก็บข้อมูลผู้ถูกหักภาษี
const TRANSACTIONS_SHEET_NAME = 'Transactions'; // ชื่อชีตสำหรับเก็บรายการหักภาษี

/**
 * @description ให้บริการไฟล์ HTML สำหรับเว็บแอป
 * @param {object} e The event parameter.
 * @returns {HtmlOutput} The HTML output for the web app.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบบันทึกภาษีหัก ณ ที่จ่าย')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * @description ดึงชีตด้วยชื่อจาก Spreadsheet ที่ระบุ หรือสร้างขึ้นใหม่พร้อมหัวตารางหากยังไม่มี
 * @param {string} sheetName The name of the sheet.
 * @param {Array<string>} headers An array of header strings.
 * @returns {Sheet} The Google Sheet object.
 */
function getSheet(sheetName, headers) {
  // ใช้ .openById() เพื่อเปิดไฟล์ Spreadsheet ตาม ID ที่ระบุไว้ด้านบน
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  
  // หากไม่พบชีต ให้สร้างชีตใหม่พร้อมกำหนดหัวตาราง
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * @description แปลงข้อมูลจากชีต (Array 2 มิติ) ให้อยู่ในรูปแบบ Array ของ Object
 * @param {Array<Array<any>>} data The 2D array from getValues().
 * @returns {Array<object>} An array of objects.
 */
function sheetDataToObjects(data) {
  if (data.length < 2) return [];
  const headers = data[0].map(h => h.toString().trim());
  return data.slice(1).map((row, index) => {
    let obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    obj.rowIndex = index + 2; // เก็บเลขแถวเดิมไว้สำหรับแก้ไข/ลบข้อมูล
    return obj;
  });
}

// --- Payee Functions ---

/**
 * @description ดึงข้อมูลผู้ถูกหักภาษีทั้งหมด
 * @returns {Array<object>} An array of payee objects.
 */
function getPayees() {
  try {
    const headers = ['id', 'name', 'type', 'taxId', 'address', 'createdAt'];
    const sheet = getSheet(PAYEES_SHEET_NAME, headers);
    const data = sheet.getDataRange().getValues();
    return sheetDataToObjects(data);
  } catch (error) {
    console.error("Error in getPayees:", error);
    return [];
  }
}

/**
 * @description เพิ่มข้อมูลผู้ถูกหักภาษีใหม่
 * @param {object} payee The payee object to add.
 * @returns {object} A success message.
 */
function addPayee(payee) {
  try {
    const headers = ['id', 'name', 'type', 'taxId', 'address', 'createdAt'];
    const sheet = getSheet(PAYEES_SHEET_NAME, headers);
    const newId = new Date().getTime().toString();
    const newRow = [newId, payee.name, payee.type, payee.taxId, payee.address, new Date()];
    sheet.appendRow(newRow);
    return { status: 'success', message: 'เพิ่มข้อมูลผู้ถูกหักภาษีสำเร็จ' };
  } catch (error) {
    console.error("Error in addPayee:", error);
    return { status: 'error', message: error.toString() };
  }
}

/**
 * @description ลบข้อมูลผู้ถูกหักภาษี
 * @param {number} rowIndex The row number to delete.
 * @returns {object} A success message.
 */
function deletePayee(rowIndex) {
   try {
    const sheet = getSheet(PAYEES_SHEET_NAME, []);
    sheet.deleteRow(rowIndex);
    return { status: 'success', message: 'ลบข้อมูลสำเร็จ' };
  } catch (error)
 {
    console.error("Error in deletePayee:", error);
    return { status: 'error', message: error.toString() };
  }
}

/**
 * @description อัปเดตข้อมูลผู้ถูกหักภาษี
 * @param {object} payee The payee object with updated information, including rowIndex.
 * @returns {object} A success status object.
 */
function updatePayee(payee) {
  try {
    const sheet = getSheet(PAYEES_SHEET_NAME, []);
    // Headers: ['id', 'name', 'type', 'taxId', 'address', 'createdAt']
    sheet.getRange(payee.rowIndex, 2, 1, 4).setValues([[payee.name, payee.type, payee.taxId, payee.address]]);
    return { status: 'success', message: 'แก้ไขข้อมูลสำเร็จ' };
  } catch (error) {
    console.error("Error in updatePayee:", error);
    return { status: 'error', message: error.toString() };
  }
}


// --- Transaction Functions ---

/**
 * @description ดึงข้อมูลรายการหักภาษีทั้งหมด
 * @returns {Array<object>} An array of transaction objects.
 */
function getTransactions() {
  try {
    const headers = [
      'id', 'payeeName', 'payeeTaxId', 'payeeAddress', 'payeeType', 
      'description', 'paymentDate', 'totalAmount', 
      'whtAmount', 'netAmount', 'createdAt'
    ];
    const sheet = getSheet(TRANSACTIONS_SHEET_NAME, headers);
    const data = sheet.getDataRange().getValues();
    const transactions = sheetDataToObjects(data);
    
    return transactions.map(t => {
      t.paymentDate = t.paymentDate instanceof Date ? t.paymentDate.toISOString().split('T')[0] : t.paymentDate;
      t.createdAt = t.createdAt instanceof Date ? t.createdAt.toISOString() : t.createdAt;
      return t;
    });
  } catch (error) {
    console.error("Error in getTransactions:", error);
    return [];
  }
}

/**
 * @description เพิ่มรายการหักภาษีใหม่
 * @param {object} tx The transaction object to add.
 * @returns {object} A success message.
 */
function addTransaction(tx) {
  try {
    const headers = [
      'id', 'payeeName', 'payeeTaxId', 'payeeAddress', 'payeeType', 
      'description', 'paymentDate', 'totalAmount', 
      'whtAmount', 'netAmount', 'createdAt'
    ];
    const sheet = getSheet(TRANSACTIONS_SHEET_NAME, headers);
    const newId = new Date().getTime().toString();
    const newRow = [
      newId, tx.payeeName, tx.payeeTaxId, tx.payeeAddress, tx.payeeType,
      tx.description, tx.paymentDate, tx.totalAmount,
      tx.whtAmount, tx.netAmount, new Date()
    ];
    sheet.appendRow(newRow);
    return { status: 'success', message: 'เพิ่มรายการหักภาษีสำเร็จ' };
  } catch (error) {
    console.error("Error in addTransaction:", error);
    return { status: 'error', message: error.toString() };
  }
}

/**
 * @description ลบรายการหักภาษี
 * @param {number} rowIndex The row number to delete.
 * @returns {object} A success message.
 */
function deleteTransaction(rowIndex) {
   try {
    const sheet = getSheet(TRANSACTIONS_SHEET_NAME, []);
    sheet.deleteRow(rowIndex);
    return { status: 'success', message: 'ลบรายการสำเร็จ' };
  } catch (error) {
    console.error("Error in deleteTransaction:", error);
    return { status: 'error', message: error.toString() };
  }
}
