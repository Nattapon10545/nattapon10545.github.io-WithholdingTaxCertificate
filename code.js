// Google Apps Script to link with the HTML form above
// Auto-create sheets and headers if they don't exist

function doGet(e) {
  return HtmlService.createHtmlOutput("<h1>Google Apps Script Connected</h1>");
}

function getSpreadsheet() {
  const ssId = '1cwzGbVUXFEKNXFJh57haEQzE_DiUvO_0daoAvflNVFA';
  return SpreadsheetApp.openById(ssId);
}

function ensureSheet(name, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function getTaxpayers() {
  const sheet = ensureSheet("Taxpayers", ["id", "name", "type", "taxId", "address"]);
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({
    id: row[0], name: row[1], type: row[2], taxId: row[3], address: row[4]
  }));
}

function getTaxRecords(month) {
  const sheet = ensureSheet("TaxRecords", ["id", "taxpayerId", "name", "type", "taxId", "address", "paymentDetail", "amount", "paymentDate", "taxAmount", "netAmount"]);
  const data = sheet.getDataRange().getValues();
  let records = data.slice(1).map(row => ({
    id: row[0], taxpayerId: row[1], name: row[2], type: row[3], taxId: row[4], address: row[5],
    paymentDetail: row[6], amount: row[7], paymentDate: row[8], taxAmount: row[9], netAmount: row[10]
  }));
  if (month) {
    records = records.filter(r => r.paymentDate && r.paymentDate.startsWith(month));
  }
  return records;
}

function addTaxpayer(data) {
  const sheet = ensureSheet("Taxpayers", ["id", "name", "type", "taxId", "address"]);
  data.id = "tp" + new Date().getTime();
  sheet.appendRow([data.id, data.name, data.type, data.taxId, data.address]);
  return { success: true };
}

function updateTaxpayer(data) {
  const sheet = ensureSheet("Taxpayers", ["id", "name", "type", "taxId", "address"]);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.getRange(i + 1, 2, 1, 4).setValues([[data.name, data.type, data.taxId, data.address]]);
      return { success: true };
    }
  }
  return { success: false, error: 'Not found' };
}

function deleteTaxpayer(data) {
  const sheet = ensureSheet("Taxpayers", ["id", "name", "type", "taxId", "address"]);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Not found' };
}

function addTaxRecord(data) {
  const sheet = ensureSheet("TaxRecords", ["id", "taxpayerId", "name", "type", "taxId", "address", "paymentDetail", "amount", "paymentDate", "taxAmount", "netAmount"]);
  data.id = "tr" + new Date().getTime();
  sheet.appendRow([data.id, data.taxpayerId, data.name, data.type, data.taxId, data.address, data.paymentDetail, data.amount, data.paymentDate, data.taxAmount, data.netAmount]);
  return { success: true };
}

function updateTaxRecord(data) {
  const sheet = ensureSheet("TaxRecords", ["id", "taxpayerId", "name", "type", "taxId", "address", "paymentDetail", "amount", "paymentDate", "taxAmount", "netAmount"]);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.getRange(i + 1, 2, 1, 10).setValues([[data.taxpayerId, data.name, data.type, data.taxId, data.address, data.paymentDetail, data.amount, data.paymentDate, data.taxAmount, data.netAmount]]);
      return { success: true };
    }
  }
  return { success: false, error: 'Not found' };
}

function deleteTaxRecord(data) {
  const sheet = ensureSheet("TaxRecords", ["id", "taxpayerId", "name", "type", "taxId", "address", "paymentDetail", "amount", "paymentDate", "taxAmount", "netAmount"]);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Not found' };
}

function doPost(e) {
  const action = e.parameter.action;
  const data = JSON.parse(e.postData.contents);

  switch (action) {
    case 'getTaxpayers': return ContentService.createTextOutput(JSON.stringify(getTaxpayers())).setMimeType(ContentService.MimeType.JSON);
    case 'getTaxRecords': return ContentService.createTextOutput(JSON.stringify(getTaxRecords(data.month))).setMimeType(ContentService.MimeType.JSON);
    case 'addTaxpayer': return ContentService.createTextOutput(JSON.stringify(addTaxpayer(data))).setMimeType(ContentService.MimeType.JSON);
    case 'updateTaxpayer': return ContentService.createTextOutput(JSON.stringify(updateTaxpayer(data))).setMimeType(ContentService.MimeType.JSON);
    case 'deleteTaxpayer': return ContentService.createTextOutput(JSON.stringify(deleteTaxpayer(data))).setMimeType(ContentService.MimeType.JSON);
    case 'addTaxRecord': return ContentService.createTextOutput(JSON.stringify(addTaxRecord(data))).setMimeType(ContentService.MimeType.JSON);
    case 'updateTaxRecord': return ContentService.createTextOutput(JSON.stringify(updateTaxRecord(data))).setMimeType(ContentService.MimeType.JSON);
    case 'deleteTaxRecord': return ContentService.createTextOutput(JSON.stringify(deleteTaxRecord(data))).setMimeType(ContentService.MimeType.JSON);
    default:
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Invalid action' })).setMimeType(ContentService.MimeType.JSON);
  }
}
