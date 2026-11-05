// 착수 메일 전용 처리 로직 (시험번호 없음, 새 행 생성)
function processInitiationEmails(masterSheet, logSheet, headers, processedIds, writer) {
  const threads = GmailApp.search(getSetting('INIT_SEARCH_QUERY', CONFIG.INIT_SEARCH_QUERY));

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const msg of messages) {
      const msgId = msg.getId();
      if (processedIds.has(msgId)) continue;

      const subject = msg.getSubject();
      if (subject.match(/^(Re|Fwd|RE|FW):/i)) continue;

      const sender = msg.getFrom();
      const date = msg.getDate();
      const body = msg.getPlainBody();

      const initData = parseInitiationEmail(sender, body, date);

      const newRow = headers.map(header => {
        const h = header.toString().trim();
        return initData[h] !== undefined ? initData[h] : "";
      });

      writer.appendRow(masterSheet, newRow, { msgId: msgId, testNumber: '', headers: headers, reason: '착수메일: 새 행 생성' });
      writer.log(logSheet, msgId, date, `착수메일 인식: 새 행 생성 성공 (시험시작일: ${initData['시험시작일']})`, '');
    }
  }
}
