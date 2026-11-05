// GS 라벨 진행 메일 처리 로직 (시험번호 매칭, LLM 일정 추출, 시트 갱신, 옵시디언 동기화)
function processGSLabelEmails(masterSheet, logSheet, queueSheet, colMap, headers, processedIds, writer) {
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

  for (const thread of threads) {
    const messages = thread.getMessages();
    const threadId = thread.getId();

    for (const msg of messages) {
      const msgId = msg.getId();
      // 1. 이미 처리된 메일이면 스킵 (LLM 미호출)
      if (processedIds.has(msgId)) continue;
      
      const subject = msg.getSubject();
      const sender = msg.getFrom();
      const date = msg.getDate();
      const body = msg.getPlainBody();

      // 2. 제목에 시험번호가 없으면 스킵 (LLM 미호출)
      const testNumMatch = subject.match(CONFIG.TEST_NUM_REGEX);
      if (!testNumMatch) {
        writer.log(logSheet, msgId, date, '기존업무: 제목에 시험번호 없음', '');
        continue;
      }
      
      const testNumber = testNumMatch[0]; 
      const rowIndex = findRowByTestNumber(masterSheet, colMap['시험번호'], testNumber, CONFIG.DATA_START_ROW);
      
      // 3. 시트에 등록된 시험번호가 없으면 스킵 (LLM 미호출)
      if (!rowIndex) {
        Logger.log(`⏳ [대기]: 시트에 '${testNumber}' 행이 아직 기입되지 않아 다음 주기에 다시 확인합니다. (메일 ID: ${msgId})`);
        writer.note({ msgId: msgId, testNumber: testNumber, reason: '처리로그 없음 → 다음 주기에 다시 확인' }, CONFIG.MASTER_SHEET_NAME, '시트에 시험번호 행 없음');
        continue; // logMessage를 호출하지 않아야 다음 실행 때 다시 읽어옵니다.
      }

      newProcessedCount++; // 유효한 신규 메일 감지

      // ----------------------------------------------------
      // [STEP A] 정형 데이터 파싱 기입 (년도, 기업명, 담당자)
      // ----------------------------------------------------
      const parsedInfo = parseBasicInfo(subject, sender, body);
      const targetHeaders = ['년도', '기업명', '담당자']; 
      
      targetHeaders.forEach(headerName => {
        const colIndex = colMap[headerName];
        if (colIndex) {
          const cellValue = writer.getCellValue(masterSheet, rowIndex, colIndex);
          if (!cellValue || cellValue.toString().trim() === "") {
            writer.setCell(masterSheet, rowIndex, colIndex, parsedInfo[headerName],
                           { msgId: msgId, testNumber: testNumber, colName: headerName, reason: 'STEP A: 빈칸 채움' });
          }
        }
      });

      // ----------------------------------------------------
      // [STEP B] LLM을 통한 비정형 데이터(일정, 요약) 기입
      // ----------------------------------------------------
      const isMyEmail = sender.includes(CONFIG.MY_EMAIL);
      const maskedBody = maskSensitiveData(body);
      const currentRowData = getRowDataAsObject(masterSheet, rowIndex, headers, writer);

      // 신규 메일 검증 완료 시에만 Gemini 호출
      const llmResult = callGeminiAPI(apiKey, maskedBody, isMyEmail, date, subject, testNumber, currentRowData);
      
      if (!llmResult) {
        writer.log(logSheet, msgId, date, `LLM 호출 실패 (${testNumber})`, testNumber);
        continue;
      }

      writer.note({ msgId: msgId, testNumber: testNumber, reason: 'LLM 응답 원문 (검토용)' }, 'LLM', JSON.stringify(llmResult));

      let sendToQueue = false;
      let queueReason = [];
      const updates = [];

      // 확신도 검사
      if (llmResult.확신도 === "낮음") {
        sendToQueue = true;
        queueReason.push(`확신도 낮음: ${llmResult.불확실한이유}`);
      }

      // 열 갱신 충돌 검사
      if (llmResult.열갱신 && Array.isArray(llmResult.열갱신)) {
        for (const update of llmResult.열갱신) {
          const colName = update.열;
          if (!colMap[colName]) continue;
          
          const currentVal = currentRowData[colName];
          if (currentVal && currentVal.toString().trim() !== "") {
            sendToQueue = true;
            queueReason.push(`[${colName}] 덮어쓰기 충돌 (기존값: ${currentVal})`);
          } else {
            updates.push({ colIndex: colMap[colName], colName: colName, value: update.값, reason: 'LLM: 열갱신' });
          }
        }
      }

      // 상태 변경
      if (llmResult.상태변경) {
        updates.push({ colIndex: colMap['상태'], colName: '상태', value: llmResult.상태변경, reason: 'LLM: 상태변경' });
      }

      // 최종 반영 분기
      if (sendToQueue) {
        const mailLink = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;
        writer.appendRow(queueSheet, [testNumber, mailLink, queueReason.join(" / "), JSON.stringify(llmResult, null, 2), "대기"],
                         { msgId: msgId, testNumber: testNumber, reason: '확인 큐: ' + queueReason.join(" / ") });
        writer.log(logSheet, msgId, date, `LLM: 확인 큐로 전송 (${testNumber})`, testNumber);
      } else {
        // 시트 값 업데이트
        for (const u of updates) {
          writer.setCell(masterSheet, rowIndex, u.colIndex, u.value,
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
        
        writer.log(logSheet, msgId, date, `LLM: 시트 반영 완료 (${testNumber})`, testNumber);

        // ----------------------------------------------------
        // [STEP C] 3단계: 옵시디언(GitHub) Flow.md 업데이트
        // ----------------------------------------------------
        // DRY_RUN 에서는 GitHub 에 쓰지 않습니다. (호출 자체는 3단계 B1 에서 제거)
        if (!writer.dryRun) {
          const updatedRowData = getRowDataAsObject(masterSheet, rowIndex, headers);
          updateObsidianFlow(testNumber, updatedRowData);
        }
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
