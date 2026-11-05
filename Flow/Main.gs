// [트리거 1] 매일 오전 7시: 신규 착수 메일 감지 및 새 행 생성
function runInitiationCheck() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

  if (!masterSheet || !logSheet) {
    Logger.log('오류: 마스터 또는 처리로그 시트를 찾을 수 없습니다.');
    return;
  }

  const headers = getHeaders(masterSheet, CONFIG.HEADER_ROW);
  const processedIds = getProcessedMessageIds(logSheet);

  Logger.log('🚀 [착수 메일 감지] 시작');
  processInitiationEmails(masterSheet, logSheet, headers, processedIds);
}

// [트리거 2] 매 1시간마다: GS 라벨 진행 메일 처리 (파싱 + LLM + 옵시디언)
function runProgressCheck() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  const queueSheet = ss.getSheetByName(CONFIG.QUEUE_SHEET_NAME);

  if (!masterSheet || !logSheet || !queueSheet) {
    Logger.log('오류: 마스터, 처리로그, 확인 큐 시트 중 하나를 찾을 수 없습니다.');
    return;
  }

  const colMap = getColumnMap(masterSheet, CONFIG.HEADER_ROW);
  const headers = getHeaders(masterSheet, CONFIG.HEADER_ROW);
  const processedIds = getProcessedMessageIds(logSheet);

  Logger.log('🚀 [진행 메일(GS 라벨) 처리] 시작');
  processGSLabelEmails(masterSheet, logSheet, queueSheet, colMap, headers, processedIds);
}
