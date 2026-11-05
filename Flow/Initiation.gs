// 착수 메일 전용 처리 로직 (시험번호 없음, 새 행 생성)
function processInitiationEmails(masterSheet, logSheet, headers, processedIds, writer, deadline) {
  const threads = GmailApp.search(getSetting('INIT_SEARCH_QUERY', CONFIG.INIT_SEARCH_QUERY));
  const keys = headers.map(h => (h ? normalizeHeader(h) : ''));   // C2: 헤더 정규화
  const skipCols = getFormulaColumnIndexes(headers);

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const msg of messages) {
      const msgId = msg.getId();
      if (processedIds.has(msgId)) continue;
      // B3: 시간이 다 되면 처리로그 없이 끝냅니다 (다음 실행에서 이어서 처리)
      if (deadline && deadline.expired()) {
        Logger.log('⏱️ 실행 시간 5분이 지나 남은 착수 메일은 다음 실행에서 처리합니다.');
        return;
      }

      const subject = msg.getSubject();
      if (subject.match(/^(Re|Fwd|RE|FW):/i)) continue;

      const sender = msg.getFrom();
      const date = msg.getDate();
      const body = msg.getPlainBody();

      const initData = parseInitiationEmail(sender, body, date);

      // C1: 수식 열을 뺀 모든 열 기준으로 새 행 위치를 정합니다 (appendRow 는 F열 수식 아래로 감)
      const rowIndex = findNewRowIndex(writer.getAllValues(masterSheet, headers.length), skipCols, CONFIG.DATA_START_ROW);

      // 쓰기 직전에 대상 행이 비었는지 다시 확인. 비어 있지 않으면 처리로그 없이 넘어가 다음 주기에 다시 봅니다.
      if (rowHasData(writer.getRowValues(masterSheet, rowIndex, headers.length), skipCols)) {
        Logger.log(`⚠️ 착수메일 보류: ${rowIndex}행이 비어 있지 않습니다. 다음 주기에 다시 확인합니다. (메일 ID: ${msgId})`);
        continue;
      }

      // 값이 있는 칸만 씁니다 (수식 열·빈칸은 건드리지 않음)
      keys.forEach((k, i) => {
        const v = initData[k];
        if (!k || v === undefined || v === '' || skipCols.includes(i + 1)) return;
        writer.setCell(masterSheet, rowIndex, i + 1, v, { msgId: msgId, testNumber: '', colName: k, reason: '착수메일: 새 행 생성' });
      });
      writer.log(logSheet, msgId, date, `착수메일 인식: ${rowIndex}행 생성 성공 (시험시작일: ${initData['시험시작일']})`, '');
    }
  }
}
