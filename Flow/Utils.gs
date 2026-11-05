function getColumnMap(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const colMap = {};
  
  headers.forEach((header, index) => {
    if (header) {
      // 2번 패치: 헤더에 괄호가 있을 경우 괄호 앞부분만 잘라내어 공백 제거 후 키로 사용
      const cleanHeader = header.split('(')[0].trim();
      colMap[cleanHeader] = index + 1;
    }
  });
  
  return colMap;
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

// 6번 패치: 수식이 있는 열 대신 실제 데이터 기준으로 마지막 행을 찾는 함수 추가
function getLastDataRow(sheet, columnNumber) {
  const data = sheet.getRange(1, columnNumber, sheet.getMaxRows(), 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== "" && data[i][0] != null) {
      return i + 1; 
    }
  }
  return 1; 
}
