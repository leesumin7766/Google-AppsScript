// GS 라벨 진행 메일 처리 로직 (시험번호 매칭, LLM 일정 추출, 시트 갱신, 옵시디언 동기화)
// logState: getLogState() 결과. '재시도대기' 인 메일만 다시 처리합니다.
function processGSLabelEmails(masterSheet, logSheet, queueSheet, colMap, headers, logState, writer) {
  const label = GmailApp.getUserLabelByName(CONFIG.LABEL_NAME);
  if (!label) {
    Logger.log(`⚠️ 알림: '${CONFIG.LABEL_NAME}' 라벨이 존재하지 않습니다.`);
    return;
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    Logger.log('오류: 스크립트 속성에 GEMINI_API_KEY가 없습니다.');
    return;
  }

  const threads = label.getThreads(0, Number(getSetting('GS_THREAD_LIMIT', CONFIG.GS_THREAD_LIMIT)));
  if (threads.length === 0) {
    Logger.log(`ℹ️ [알림]: '${CONFIG.LABEL_NAME}' 라벨이 지정된 메일이 0건입니다.`);
    return;
  }

  let newProcessedCount = 0; // 신규 처리 메일 카운터

  threadLoop:
  for (const thread of threads) {
    const messages = thread.getMessages();
    const threadId = thread.getId();
    const mailLink = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;
    const threadSubjects = messages.map(m => m.getSubject());

    for (const msg of messages) {
      const msgId = msg.getId();
      // 1. 이미 처리된 메일이면 스킵 (재시도대기 는 다시 처리)
      const prev = logState[msgId];
      if (prev && prev.status !== LOG_STATUS.RETRY) continue;
      const prevAttempts = prev ? prev.attempts : 0;
      const logRow = prev ? prev.row : undefined;

      const subject = msg.getSubject();
      const sender = msg.getFrom();
      const date = msg.getDate();
      const body = msg.getPlainBody();
      const isMyEmail = extractEmails(sender).includes(CONFIG.MY_EMAIL.toLowerCase());

      // 2. 시험번호 찾기 (A3): 제목 → 같은 스레드 제목 → 메일 도메인
      const resolved = resolveTestNumber({
        subject: subject,
        threadSubjects: threadSubjects,
        from: sender,
        to: [msg.getTo ? msg.getTo() : '', msg.getCc ? msg.getCc() : ''].join(', '),
        isMyEmail: isMyEmail,
        masterValues: () => writer.getAllValues(masterSheet, headers.length),
        colMap: colMap,
      });

      if (!resolved.testNumber) {
        const candidates = (resolved.candidates || []).length ? ` (후보: ${resolved.candidates.join(', ')})` : '';
        const reason = `시험번호 없음: ${resolved.reason}${candidates}`;
        writer.appendRow(queueSheet, ['', mailLink, reason, JSON.stringify({ 제목: subject, 발신: sender, 후보: resolved.candidates || [] }, null, 2), '대기'],
                         { msgId: msgId, testNumber: '', reason: '확인 큐: ' + reason });
        writer.log(logSheet, msgId, date, `확인 큐로 전송: ${reason}`, '', { status: LOG_STATUS.QUEUED, row: logRow });
        continue;
      }

      const testNumber = resolved.testNumber;
      if (resolved.source !== '제목') {
        Logger.log(`🔎 시험번호 ${testNumber} 를 ${resolved.source} 에서 찾았습니다. (메일 ID: ${msgId})`);
        writer.note({ msgId: msgId, testNumber: testNumber, reason: '시험번호 찾기 (A3)' }, CONFIG.MASTER_SHEET_NAME,
                    `${resolved.source} 에서 찾음`);
      }

      const rowIndex = findRowByTestNumber(masterSheet, colMap['시험번호'], testNumber, CONFIG.DATA_START_ROW);

      // 3. 시트에 등록된 시험번호가 없으면 스킵 (LLM 미호출)
      if (!rowIndex) {
        Logger.log(`⏳ [대기]: 시트에 '${testNumber}' 행이 아직 기입되지 않아 다음 주기에 다시 확인합니다. (메일 ID: ${msgId})`);
        writer.note({ msgId: msgId, testNumber: testNumber, reason: '처리로그 없음 → 다음 주기에 다시 확인' }, CONFIG.MASTER_SHEET_NAME, '시트에 시험번호 행 없음');
        continue; // logMessage를 호출하지 않아야 다음 실행 때 다시 읽어옵니다.
      }

      newProcessedCount++; // 유효한 신규 메일 감지

      // C4: 이전 메일 인용문을 자른 본문만 씁니다 (인용된 남의 서명·날짜를 잘못 읽지 않게)
      const cleanBody = cleanEmailBody(body);

      // ----------------------------------------------------
      // [STEP A] 정형 데이터 파싱 기입 (년도, 기업명, 담당자)
      // A4: 내가 보낸 메일이면 건너뜁니다 (내 이름·서명을 담당자로 넣지 않게)
      // ----------------------------------------------------
      if (isMyEmail) {
        writer.note({ msgId: msgId, testNumber: testNumber, reason: 'A4' }, CONFIG.MASTER_SHEET_NAME, '내 메일 — STEP A 건너뜀');
      } else {
        const parsedInfo = parseBasicInfo(subject, sender, cleanBody, testNumber);
        ['년도', '기업명', '담당자'].forEach(headerName => {
          const colIndex = colMap[headerName];
          const value = parsedInfo[headerName];
          if (!colIndex || value === '' || value === null || value === undefined) return;
          const cellValue = writer.getCellValue(masterSheet, rowIndex, colIndex);
          if (!cellValue || cellValue.toString().trim() === "") {
            writer.setCell(masterSheet, rowIndex, colIndex, value,
                           { msgId: msgId, testNumber: testNumber, colName: headerName, reason: 'STEP A: 빈칸 채움' });
          }
        });
      }

      // ----------------------------------------------------
      // [STEP B] LLM을 통한 비정형 데이터(일정, 요약) 기입
      // ----------------------------------------------------
      // C4: 자르고 나서 본문이 거의 비면 LLM 을 부르지 않습니다
      if (isBodyNearlyEmpty(cleanBody)) {
        writer.log(logSheet, msgId, date, `LLM 생략: 인용문을 자르니 본문이 비어 있음 (${testNumber})`, testNumber, { row: logRow });
        continue;
      }

      const maskedBody = maskSensitiveData(cleanBody);
      const currentRowData = getRowDataAsObject(masterSheet, rowIndex, headers, writer);

      const res = callGeminiAPI(apiKey, maskedBody, isMyEmail, date, subject, testNumber, currentRowData);
      const attempts = prevAttempts + 1;

      if (!res.ok) {
        // A2: 키·모델·스키마 문제는 모든 메일에 같으므로 처리로그를 남기지 않고 실행을 멈춥니다
        if (res.kind === 'fatal') {
          Logger.log(`오류: Gemini 호출 불가 (${res.message}). 실행을 멈춥니다. 처리로그는 남기지 않습니다.`);
          writer.note({ msgId: msgId, testNumber: testNumber, reason: 'LLM 호출 불가 — 실행 중단' }, 'LLM', res.message);
          break threadLoop;
        }
        const giveUp = res.kind === 'blocked' || attempts >= CONFIG.LLM_MAX_ATTEMPTS;
        writer.log(logSheet, msgId, date, `LLM 실패(${res.kind}): ${res.message} — ${attempts}/${CONFIG.LLM_MAX_ATTEMPTS} (${testNumber})`, testNumber,
                   { status: giveUp ? LOG_STATUS.GIVE_UP : LOG_STATUS.RETRY, attempts: attempts, row: logRow });
        continue;
      }
      const llmResult = res.data;

      writer.note({ msgId: msgId, testNumber: testNumber, reason: 'LLM 응답 원문 (검토용)' }, 'LLM', JSON.stringify(llmResult));

      // C2·C3·A1·A8: 허용 열·상태 허용값·날짜 형식·같은 날짜·덮어쓰기 충돌·날짜 순서를 코드로 검사
      const check = checkLlmUpdates(llmResult, currentRowData, colMap);
      check.ignored.forEach(i => {
        Logger.log(`LLM 제안 무시 (${testNumber}) [${i.colName}] ${i.value}: ${i.reason}`);
        writer.note({ msgId: msgId, testNumber: testNumber, reason: `LLM 제안 무시: ${i.reason}` }, CONFIG.MASTER_SHEET_NAME,
                    `[${i.colName}] ${i.value}`);
      });

      const sendQueue = (reasons) => writer.appendRow(queueSheet,
        [testNumber, mailLink, reasons.join(" / "), JSON.stringify(llmResult, null, 2), "대기"],
        { msgId: msgId, testNumber: testNumber, reason: '확인 큐: ' + reasons.join(" / ") });

      // 확신도가 낮으면 아무것도 반영하지 않고 전부 확인 큐로 보냅니다 (원칙 1)
      if (llmResult.확신도 === "낮음") {
        const reasons = [`확신도 낮음: ${llmResult.불확실한이유}`].concat(
          check.updates.map(u => `[${u.colName}] 제안: ${u.value}`), check.queueReasons);
        sendQueue(reasons);
        writer.log(logSheet, msgId, date, `LLM: 확신도 낮음 → 확인 큐 (${testNumber})`, testNumber,
                   { status: LOG_STATUS.QUEUED, attempts: attempts, row: logRow });
        continue;
      }

      // A1: 충돌이 없는 열은 반영하고, 충돌한 열만 확인 큐로 보냅니다
      for (const u of check.updates) {
        writer.setCell(masterSheet, rowIndex, colMap[u.colName], u.value,
                       { msgId: msgId, testNumber: testNumber, colName: u.colName, reason: u.reason });
      }

      // 진행상황 추가 기입
      if (llmResult.요약 && colMap['진행상황']) {
        const oldProgress = writer.getCellValue(masterSheet, rowIndex, colMap['진행상황']);
        const mmdd = Utilities.formatDate(date, "Asia/Seoul", "MM/dd");
        const newLine = `[${mmdd}] ${llmResult.요약}`;
        writer.setCell(masterSheet, rowIndex, colMap['진행상황'], oldProgress ? oldProgress + "\n" + newLine : newLine,
                       { msgId: msgId, testNumber: testNumber, colName: '진행상황', reason: '요약 덧붙임' });
      }

      if (check.queued.length > 0) {
        sendQueue(check.queueReasons);
        writer.log(logSheet, msgId, date, `LLM: ${check.updates.length}개 열 반영, ${check.queued.length}개 열 확인 큐 (${testNumber})`, testNumber,
                   { status: LOG_STATUS.QUEUED, attempts: attempts, row: logRow });
      } else {
        writer.log(logSheet, msgId, date, `LLM: 시트 반영 완료 (${testNumber})`, testNumber,
                   { status: LOG_STATUS.DONE, attempts: attempts, row: logRow });
      }

      // ----------------------------------------------------
      // [STEP C] 옵시디언(GitHub) Flow.md 업데이트
      // ----------------------------------------------------
      // DRY_RUN 에서는 GitHub 에 쓰지 않습니다. (호출 자체는 3단계 B1 에서 제거)
      if (!writer.dryRun) {
        const updatedRowData = getRowDataAsObject(masterSheet, rowIndex, headers);
        updateObsidianFlow(testNumber, updatedRowData);
      }
    }
  }

  // 반복 종료 후 결과 요약 로그 출력
  if (newProcessedCount === 0) {
    Logger.log('ℹ️ [진행 메일 탐색 완료]: 처리할 새로운 메일이 없습니다. (모두 이미 처리됨)');
  } else {
    Logger.log(`✅ [진행 메일 처리 완료]: 총 ${newProcessedCount}건의 신규 메일을 처리했습니다.`);
  }
}
