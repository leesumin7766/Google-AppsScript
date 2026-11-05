// 2번 패치: 헤더에 괄호가 있을 경우 괄호 앞부분만 잘라내어 공백 제거 후 키로 사용
// (순수 함수 — tests/ 에서 시트 없이 확인)
function normalizeHeader(header) {
  return String(header).split('(')[0].trim();
}

function getColumnMap(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const colMap = {};
  
  headers.forEach((header, index) => {
    if (header) {
      colMap[normalizeHeader(header)] = index + 1;
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

// 6번 패치 (1단계 C1 에서 다시 작성): 새 행 위치
// 수식 열을 뺀 **모든 열**에서 마지막으로 값이 있는 행 다음 행. 데이터 시작 행(3)보다 위로는 가지 않습니다.
// 시험번호가 빈 착수 행도 상담자·시험안내일·상태가 있으므로 덮어쓰지 않습니다.
// values: 1행부터의 2차원 배열, skipCols: 무시할 열 번호(1부터). 순수 함수.
function findNewRowIndex(values, skipCols, startRow) {
  let last = 0;
  for (let r = values.length - 1; r >= 0; r--) {
    if (rowHasData(values[r], skipCols)) { last = r + 1; break; }
  }
  return Math.max(last + 1, startRow);
}

function rowHasData(rowValues, skipCols) {
  return rowValues.some((v, c) => !skipCols.includes(c + 1) && v !== '' && v !== null && v !== undefined);
}

// 수식 열(CONFIG.FORMULA_COLUMNS)의 열 번호 목록
function getFormulaColumnIndexes(headers) {
  const result = [];
  headers.forEach((h, i) => { if (h && CONFIG.FORMULA_COLUMNS.includes(normalizeHeader(h))) result.push(i + 1); });
  return result;
}
