// [트리거 1] 매일 오전 7시: 신규 착수 메일 감지 및 새 행 생성
function runInitiationCheck() {
  withScriptLock('착수 메일 감지', () => {
    const deadline = createDeadline();
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const masterSheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

    if (!masterSheet || !logSheet) {
      Logger.log('오류: 마스터 또는 처리로그 시트를 찾을 수 없습니다.');
      return;
    }

    const writer = createWriter(ss);
    if (!writer) return;
    if (!ensureLogHeaders(logSheet, writer.dryRun)) return;

    const headers = getHeaders(masterSheet, CONFIG.HEADER_ROW);
    const processedIds = getProcessedMessageIds(logSheet);

    Logger.log('🚀 [착수 메일 감지] 시작');
    processInitiationEmails(masterSheet, logSheet, headers, processedIds, writer, deadline);
  });
}

// [트리거 2] 매 1시간마다: 확인 큐 결정 반영 → GS 라벨 진행 메일 처리 (파싱 + LLM)
function runProgressCheck() {
  withScriptLock('진행 메일 처리', () => {
    const deadline = createDeadline();
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const masterSheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

    if (!masterSheet || !logSheet) {
      Logger.log('오류: 마스터 또는 처리로그 시트를 찾을 수 없습니다.');
      return;
    }

    const writer = createWriter(ss);
    if (!writer) return;
    if (!ensureLogHeaders(logSheet, writer.dryRun)) return;
    const queueSheet = ensureQueueSheet(ss, writer.dryRun);
    const todoSheet = ensureTodoSheet(ss, writer.dryRun);
    if (!queueSheet || !todoSheet) return;

    const colMap = getColumnMap(masterSheet, CONFIG.HEADER_ROW);
    const headers = getHeaders(masterSheet, CONFIG.HEADER_ROW);

    // B2: 사람·ATDT 가 확인 큐에서 정한 것을 먼저 반영 (onEdit 를 놓친 것 포함)
    applyQueueDecisions(queueSheet, masterSheet, colMap, writer);

    const logState = getLogState(logSheet);
    Logger.log('🚀 [진행 메일(GS 라벨) 처리] 시작');
    processGSLabelEmails(masterSheet, logSheet, queueSheet, todoSheet, colMap, headers, logState, writer, deadline);
  });
}

// B3: 두 트리거가 동시에 돌지 않게 합니다. 다른 실행이 끝나기를 30초까지 기다립니다.
function withScriptLock(name, fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log(`⏳ ${name}: 다른 실행이 진행 중이라 이번에는 건너뜁니다.`);
    return;
  }
  try {
    fn();
  } finally {
    lock.releaseLock();
  }
}

// B3: 실행 시간 제한 (Apps Script 한도 6분 — 5분이 지나면 남은 메일은 다음 주기에)
function createDeadline(budgetMs) {
  const end = Date.now() + (budgetMs || CONFIG.TIME_BUDGET_MS);
  return { expired: () => Date.now() > end };
}
