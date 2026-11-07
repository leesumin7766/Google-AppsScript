// 3단계 B1·B2: `확인 큐`·`할 일` 탭 준비, 확인 큐 결정(적용/무시) 반영
const QUEUE_HEADERS = ['ID', '생성시각', '시험번호', '메시지ID', '메일 링크', '열', '기존값', '제안값', '사유', 'LLM 원문', '처리', '처리시각'];
const QUEUE_DECISIONS = ['대기', '적용', '무시'];
const TODO_HEADERS = ['ID', '시험번호', '내용', '기한', '출처 메시지ID', '생성시각', '상태'];
const TODO_STATUSES = ['신규', '반영', '무시'];
const CELL_TEXT_LIMIT = 45000; // 셀 한 칸 한도(50000자)보다 조금 작게

// DRY_RUN 에서 탭이 아직 없을 때 쓰는 빈 탭 (읽기만 되고 아무것도 없음)
function virtualSheet(name) {
  return {
    isVirtual: true,
    getName: () => name,
    getLastRow: () => 1,
    getLastColumn: () => 0,
  };
}

// 탭이 없으면 만들고, 비어 있으면 머리글·드롭다운을 넣습니다. DRY_RUN 이면 바꾸지 않습니다.
// 머리글이 다르면: 확인 큐가 예전 형식이면 '확인 큐 (v1)' 로 이름을 바꿔 보관하고 새로 만듭니다.
// 반환: 탭 (DRY_RUN 이고 없으면 virtualSheet), 멈춰야 하면 null
function ensureTabWithHeaders(ss, name, headers, dropdown, dryRun, archiveOldFormat) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    if (dryRun) return virtualSheet(name);
    sheet = ss.insertSheet(name);
    writeHeaders(sheet, headers, dropdown);
    Logger.log(`'${name}' 탭을 만들었습니다.`);
    return sheet;
  }
  if (sheet.getLastRow() === 0) {
    if (!dryRun) writeHeaders(sheet, headers, dropdown);
    return sheet;
  }
  const have = sheet.getRange(1, 1, 1, headers.length).getValues()[0].map(v => String(v).trim());
  if (have.join('|') === headers.join('|')) return sheet;

  if (archiveOldFormat && have[0] !== headers[0]) {
    if (dryRun) {
      Logger.log(`DRY_RUN: '${name}' 이 예전 형식입니다. 실제 실행 때는 '${name} (v1)' 로 보관하고 새로 만듭니다.`);
      return virtualSheet(name);
    }
    let archiveName = `${name} (v1)`;
    if (ss.getSheetByName(archiveName)) archiveName = `${name} (v1 ${Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd-HHmmss')})`;
    sheet.setName(archiveName);
    const fresh = ss.insertSheet(name);
    writeHeaders(fresh, headers, dropdown);
    Logger.log(`예전 형식의 '${name}' 을 '${archiveName}' 로 보관하고 새 '${name}' 을 만들었습니다.`);
    return fresh;
  }
  Logger.log(`오류: '${name}' 탭의 머리글이 다릅니다. 기대: ${headers.join(', ')} / 현재: ${have.join(', ')}. 실행을 멈춥니다.`);
  return null;
}

// dropdown: { header, values } — 그 열 2행부터 드롭다운
function writeHeaders(sheet, headers, dropdown) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  if (dropdown) {
    const col = headers.indexOf(dropdown.header) + 1;
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(dropdown.values, true).setAllowInvalid(false).build();
    sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  }
}

function ensureQueueSheet(ss, dryRun) {
  return ensureTabWithHeaders(ss, CONFIG.QUEUE_SHEET_NAME, QUEUE_HEADERS, { header: '처리', values: QUEUE_DECISIONS }, dryRun, true);
}

function ensureTodoSheet(ss, dryRun) {
  return ensureTabWithHeaders(ss, CONFIG.TODO_SHEET_NAME, TODO_HEADERS, { header: '상태', values: TODO_STATUSES }, dryRun, false);
}

function clipText(s) {
  s = String(s === null || s === undefined ? '' : s);
  return s.length > CELL_TEXT_LIMIT ? s.slice(0, CELL_TEXT_LIMIT) + ' …(잘림)' : s;
}

// ---------------------------------------------------------------------------
// 메일 링크: 한 브라우저에 구글 계정이 여러 개면 /u/0 은 다른 계정일 수 있어 계정 이메일을 넣습니다 (2026-11-07 실환경).
// account 가 이메일 모양이 아니면 예전처럼 /u/0.
// ---------------------------------------------------------------------------
function mailThreadLink(threadId, account) {
  const acct = /^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(String(account || '')) ? String(account).toLowerCase() : '0';
  return `https://mail.google.com/mail/u/${acct}/#all/${threadId}`;
}

// 메일 링크에 넣을 계정: 스크립트 속성 MAIL_ACCOUNT, 없으면 CONFIG.MY_EMAIL
function mailAccount() {
  return getSetting('MAIL_ACCOUNT', CONFIG.MY_EMAIL);
}

// ---------------------------------------------------------------------------
// 확인 큐 행 (B2): 변경 1건당 1행. 순수 함수.
// items: [{ colName, current, value, reason }]
// ---------------------------------------------------------------------------
function buildQueueRows(opts) {
  const now = opts.now;
  const link = mailThreadLink(opts.threadId, opts.account);   // 보관함 메일도 열림
  const raw = clipText(typeof opts.llmRaw === 'string' ? opts.llmRaw : JSON.stringify(opts.llmRaw, null, 2));
  const items = opts.items.length ? opts.items : [{ colName: '', current: '', value: '', reason: opts.fallbackReason || '' }];
  return items.map((it, i) => [
    `Q-${opts.msgId}-${i + 1}`, now, opts.testNumber || '', opts.msgId, link,
    it.colName || '', toQueueCell(it.current), toQueueCell(it.value), it.reason || '', raw, '대기', '',
  ]);
}

function toQueueCell(v) {
  if (v === null || v === undefined) return '';
  const ymd = Object.prototype.toString.call(v) === '[object Date]' ? normalizeYmd(v) : null;
  return ymd || v;
}

// ---------------------------------------------------------------------------
// 할 일 (B1): LLM 의 `할일` → `할 일` 탭. 같은 시험번호·내용·기한이면 넣지 않습니다. 순수 함수.
// existingKeys: Set('시험번호|내용|기한')
// ---------------------------------------------------------------------------
function buildTodoRows(todos, testNumber, msgId, now, existingKeys) {
  const rows = [];
  (Array.isArray(todos) ? todos : []).forEach(t => {
    const content = t && t.내용 ? String(t.내용).trim() : '';
    if (!content) return;
    const due = normalizeYmd(t.기한) || '';
    const key = todoKey(testNumber, content, due);
    if (existingKeys.has(key)) return;
    existingKeys.add(key);
    rows.push([`T-${msgId}-${rows.length + 1}`, testNumber, content, due, msgId, now, '신규']);
  });
  return rows;
}

function todoKey(testNumber, content, due) {
  return [String(testNumber || '').trim().toUpperCase(), String(content || '').trim(), normalizeYmd(due) || ''].join('|');
}

function readTodoKeys(todoSheet) {
  const keys = new Set();
  if (todoSheet.isVirtual || todoSheet.getLastRow() < 2) return keys;
  todoSheet.getRange(2, 2, todoSheet.getLastRow() - 1, 3).getValues()
    .forEach(r => { if (r[1]) keys.add(todoKey(r[0], r[1], r[2])); });
  return keys;
}

// ---------------------------------------------------------------------------
// 확인 큐 결정 반영 (B2)
// `처리` 가 적용/무시 이고 `처리시각` 이 빈 행을 처리합니다. 사람이 시트에서 바꾸면 onEdit 트리거가,
// 다른 프로그램(ATDT)이 바꾸면 매시간 실행(runProgressCheck) 시작 때 처리합니다.
// 적용: 마스터의 현재 값이 `기존값` 과 같을 때만 `제안값` 을 씁니다 (그사이 사람이 고쳤으면 덮어쓰지 않음).
// ---------------------------------------------------------------------------
function applyQueueDecisions(queueSheet, masterSheet, colMap, writer) {
  if (queueSheet.isVirtual || queueSheet.getLastRow() < 2) return { applied: 0, ignored: 0, refused: 0 };
  const C = {}; QUEUE_HEADERS.forEach((h, i) => { C[h] = i; });
  const values = queueSheet.getRange(2, 1, queueSheet.getLastRow() - 1, QUEUE_HEADERS.length).getValues();
  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  const result = { applied: 0, ignored: 0, refused: 0 };

  values.forEach((r, i) => {
    const row = i + 2;
    const decision = String(r[C['처리']]).trim();
    if (r[C['처리시각']] !== '' || (decision !== '적용' && decision !== '무시')) return;
    const ctx = { msgId: r[C['메시지ID']], testNumber: r[C['시험번호']], colName: '처리시각', reason: `확인 큐 ${decision}` };

    if (decision === '무시') {
      writer.setCell(queueSheet, row, C['처리시각'] + 1, now, ctx);
      result.ignored++;
      return;
    }
    const refuse = (why) => {
      writer.setCell(queueSheet, row, C['처리'] + 1, '대기', Object.assign({}, ctx, { colName: '처리' }));
      writer.setCell(queueSheet, row, C['사유'] + 1, `${r[C['사유']]} / 적용 안 함(${now}): ${why}`, Object.assign({}, ctx, { colName: '사유' }));
      result.refused++;
    };

    const check = checkQueueApply(r[C['열']], r[C['기존값']], r[C['제안값']]);
    if (!check.ok) return refuse(check.reason);
    const testNumber = String(r[C['시험번호']]).trim();
    const masterRow = testNumber ? findRowByTestNumber(masterSheet, colMap['시험번호'], testNumber, CONFIG.DATA_START_ROW) : null;
    if (!masterRow) return refuse(`마스터에 ${testNumber || '시험번호'} 행이 없음`);
    if (!colMap[check.colName]) return refuse(`마스터에 '${check.colName}' 열이 없음`);
    const nowValue = writer.getCellValue(masterSheet, masterRow, colMap[check.colName]);
    if (!sameCellValue(nowValue, r[C['기존값']])) return refuse(`마스터 값이 바뀜 (현재: ${toQueueCell(nowValue) || '빈칸'})`);

    writer.setCell(masterSheet, masterRow, colMap[check.colName], check.value,
                   Object.assign({}, ctx, { colName: check.colName, reason: `확인 큐 적용 (${r[C['ID']]})` }));
    writer.setCell(queueSheet, row, C['처리시각'] + 1, now, ctx);
    result.applied++;
  });
  if (result.applied + result.ignored + result.refused > 0) {
    Logger.log(`확인 큐 결정 반영: 적용 ${result.applied} · 무시 ${result.ignored} · 적용 안 함 ${result.refused}`);
  }
  return result;
}

// 순수 함수: 확인 큐 한 행을 마스터에 적용해도 되는지 (사람이 제안값을 고쳤을 수도 있으니 다시 검사)
function checkQueueApply(colName, current, value) {
  const col = normalizeHeader(colName || '');
  if (!ALLOWED_COLUMNS.includes(col)) return { ok: false, reason: `'${col || '(열 없음)'}' 은 자동 적용할 수 없는 항목 — 직접 처리하세요` };
  const v = String(value === null || value === undefined ? '' : value).trim();
  if (col === '상태') {
    if (!VALID_STATUSES.includes(v)) return { ok: false, reason: `상태 허용값이 아님: ${v}` };
    return { ok: true, colName: col, value: v };
  }
  const ymd = normalizeYmd(value);
  if (!ymd) return { ok: false, reason: `날짜 형식(yyyy-MM-dd) 아님: ${v}` };
  return { ok: true, colName: col, value: ymd };
}

function sameCellValue(a, b) {
  const blank = x => x === null || x === undefined || String(x).trim() === '';
  if (blank(a) && blank(b)) return true;
  const da = normalizeYmd(a), db = normalizeYmd(b);
  if (da || db) return da === db;
  return String(a).trim() === String(b).trim();
}

// [트리거 3] 확인 큐 `처리` 를 바꾸면 바로 반영 (설치형 onEdit — installQueueEditTrigger 로 한 번 설치)
function onQueueEdit(e) {
  if (e && e.range && e.range.getSheet().getName() !== CONFIG.QUEUE_SHEET_NAME) return;
  withScriptLock('확인 큐 반영', () => {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const writer = createWriter(ss);
    if (!writer) return;
    if (writer.dryRun) { Logger.log('DRY_RUN: 확인 큐 결정은 반영하지 않습니다.'); return; }
    const queueSheet = ss.getSheetByName(CONFIG.QUEUE_SHEET_NAME);
    const masterSheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    if (!queueSheet || !masterSheet) return;
    applyQueueDecisions(queueSheet, masterSheet, getColumnMap(masterSheet, CONFIG.HEADER_ROW), writer);
  });
}

// 사람이 편집기에서 한 번 실행: 대상 시트에 onQueueEdit 트리거를 설치합니다 (이미 있으면 그대로).
function installQueueEditTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onQueueEdit');
  if (exists) { Logger.log('onQueueEdit 트리거가 이미 있습니다.'); return; }
  ScriptApp.newTrigger('onQueueEdit').forSpreadsheet(getSpreadsheetId()).onEdit().create();
  Logger.log(`onQueueEdit 트리거를 설치했습니다 (시트: ${getSpreadsheetId()}).`);
}
