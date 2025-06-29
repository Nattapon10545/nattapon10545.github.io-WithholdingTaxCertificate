/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of file
 * access for this script to only the current document.
 */
const PAYEES_SHEET_NAME = 'Payees'; // ชื่อชีตสำหรับเก็บข้อมูลผู้ถูกหักภาษี
const TRANSACTIONS_SHEET_NAME = 'Transactions'; // ชื่อชีตสำหรับเก็บรายการหักภาษี

/**
 * @description Serves the HTML file for the web app.
 * @param {object} e The event parameter.
 * @returns {HtmlOutput} The HTML output for the web app.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบบันทึกภาษีหัก ณ ที่จ่าย')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * @description Gets a sheet by name, creating it with headers if it doesn't exist.
 * @param {string} sheetName The name of the sheet.
 * @param {Array<string>} headers An array of header strings.
 * @returns {Sheet} The Google Sheet object.
 */
function getSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * @description Converts sheet data (2D array) to an array of objects.
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
    obj.rowIndex = index + 2; // Store original row index for easy updates/deletes
    return obj;
  });
}

// --- Payee Functions ---

/**
 * @description Retrieves all payees from the spreadsheet.
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
 * @description Adds a new payee to the sheet.
 * @param {object} payee The payee object to add.
 * @returns {object} A success message.
 */
function addPayee(payee) {
  try {
    const headers = ['id', 'name', 'type', 'taxId', 'address', 'createdAt'];
    const sheet = getSheet(PAYEES_SHEET_NAME, headers);
    const newId = new Date().getTime().toString(); // Simple unique ID
    const newRow = [newId, payee.name, payee.type, payee.taxId, payee.address, new Date()];
    sheet.appendRow(newRow);
    return { status: 'success', message: 'เพิ่มข้อมูลผู้ถูกหักภาษีสำเร็จ' };
  } catch (error) {
    console.error("Error in addPayee:", error);
    return { status: 'error', message: error.toString() };
  }
}

/**
 * @description Deletes a payee from the sheet.
 * @param {number} rowIndex The row number to delete.
 * @returns {object} A success message.
 */
function deletePayee(rowIndex) {
   try {
    const sheet = getSheet(PAYEES_SHEET_NAME, []);
    sheet.deleteRow(rowIndex);
    return { status: 'success', message: 'ลบข้อมูลสำเร็จ' };
  } catch (error) {
    console.error("Error in deletePayee:", error);
    return { status: 'error', message: error.toString() };
  }
}

/**
 * @description Updates an existing payee's data.
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
 * @description Retrieves all withholding tax transactions.
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
    // Convert date objects back to strings for proper JSON serialization
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
 * @description Adds a new transaction to the sheet.
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
 * @description Deletes a transaction from the sheet.
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
