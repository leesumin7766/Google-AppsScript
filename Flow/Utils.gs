function getColumnMap(sheet, headerRowIndex) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const headers = sheet.getRange(headerRowIndex, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    if (header) map[header.toString().trim()] = index + 1;
  });
  return map;
}

function getHeaders(sheet, headerRowIndex) {
  const lastCol = sheet.getLastColumn();
  return sheet.getRange(headerRowIndex, 1, 1, lastCol).getValues()[0];
}

function findRowByTestNumber(sheet, testNumColIndex, testNumber, startRowIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRowIndex) return null;
  const values = sheet.getRange(startRowIndex, testNumColIndex, lastRow - startRowIndex + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim().toUpperCase() === testNumber.trim().toUpperCase()) {
      return startRowIndex + i; 
    }
  }
  return null;
}

function getProcessedMessageIds(logSheet) {
  const ids = new Set();
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return ids; 
  const values = logSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  values.forEach(row => { if (row[0]) ids.add(row[0].toString()); });
  return ids;
}

function logMessage(logSheet, msgId, timestamp, result) {
  logSheet.appendRow([msgId, timestamp, result]);
}
