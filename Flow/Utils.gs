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

// D1: Flow 가 쓰는 마스터 열. 없으면 그 열의 값은 조용히 버려지므로 실행할 때마다 로그를 남깁니다.
function requiredMasterColumns() {
  return ['시험번호', '년도', '기업명', '담당자', '상태', '진행상황',
          '상담자', '시험안내일', 'WD', '시험원'].concat(LLM_DATE_COLUMNS);
}

// 반환: 없는 열 이름 목록 (순수 함수)
function findMissingColumns(colMap, required) {
  return required.filter((c, i) => required.indexOf(c) === i && !colMap[c]);
}

// 없는 열을 로그로 남깁니다. 시험번호 열이 없으면 false (진행 메일은 행을 찾을 수 없으니 멈춤).
function checkMasterColumns(colMap) {
  const missing = findMissingColumns(colMap, requiredMasterColumns());
  if (missing.length) Logger.log(`⚠️ 마스터 1행에 없는 열: ${missing.join(', ')} — 이 열에 들어갈 값은 쓰지 않습니다.`);
  return !missing.includes('시험번호');
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

// ---------------------------------------------------------------------------
// 처리로그 (2단계 A2): A 메시지ID · B 처리 시각 · C 결과 · D 상태 · E 시도횟수
//   상태: 완료 / 재시도대기 / 포기 / 확인대기.  D열이 빈 예전 행은 '완료' 로 봅니다.
// ---------------------------------------------------------------------------
const LOG_HEADERS = ['메시지ID', '처리 시각', '결과', '상태', '시도횟수'];
const LOG_STATUS = { DONE: '완료', RETRY: '재시도대기', GIVE_UP: '포기', QUEUED: '확인대기' };

// D1·E1 이 비어 있으면 헤더를 적습니다. 다른 값이 있으면 false (호출 쪽은 실행을 멈춤).
// DRY_RUN 이면 처리로그를 바꾸지 않고 검사만 합니다.
function ensureLogHeaders(logSheet, dryRun) {
  const current = logSheet.getRange(1, 4, 1, 2).getValues()[0];
  for (let i = 0; i < 2; i++) {
    const want = LOG_HEADERS[3 + i];
    const have = String(current[i] === null || current[i] === undefined ? '' : current[i]).trim();
    if (have === want) continue;
    if (have !== '') {
      Logger.log(`오류: 처리로그 ${i === 0 ? 'D1' : 'E1'} 에 '${want}' 대신 '${have}' 가 있습니다. 실행을 멈춥니다.`);
      return false;
    }
    if (!dryRun) logSheet.getRange(1, 4 + i).setValue(want);
  }
  return true;
}

// 메시지ID → { row, status, attempts }. 같은 ID 가 여러 번 있으면 아래쪽 행이 우선합니다.
function getLogState(logSheet) {
  const state = {};
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return state;
  const values = logSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  values.forEach((r, i) => {
    if (!r[0]) return;
    state[r[0].toString()] = {
      row: i + 2,
      status: String(r[3] || '').trim() || LOG_STATUS.DONE,
      attempts: Number(r[4]) || 0,
    };
  });
  return state;
}

// 다시 처리하지 않을 메시지ID (재시도대기 는 빠짐)
function getProcessedMessageIds(logSheet) {
  const state = getLogState(logSheet);
  return new Set(Object.keys(state).filter(id => state[id].status !== LOG_STATUS.RETRY));
}

// row 가 있으면 그 행을 고치고(재시도), 없으면 새 행을 붙입니다.
function logMessage(logSheet, msgId, timestamp, result, status, attempts, row) {
  const values = [msgId, timestamp, result, status || LOG_STATUS.DONE, attempts === undefined ? '' : attempts];
  if (row) logSheet.getRange(row, 1, 1, 5).setValues([values]);
  else logSheet.appendRow(values);
}

// ---------------------------------------------------------------------------
// 날짜 (2단계 A1·A8): 'yyyy-MM-dd' 하나로 맞춥니다. 순수 함수.
// ---------------------------------------------------------------------------
// Date, 'yyyy-MM-dd', 'yyyy.M.d', 'yyyy/M/d', 'yyyy. M. d.' → 'yyyy-MM-dd'. 아니면 null.
function normalizeYmd(v) {
  if (v === null || v === undefined || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? null : Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  const m = String(v).trim().match(/^(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})\.?$/);
  if (!m) return null;
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  const check = new Date(Date.UTC(y, mo - 1, d));
  if (check.getUTCFullYear() !== y || check.getUTCMonth() !== mo - 1 || check.getUTCDate() !== d) return null;
  return `${m[1]}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

// ---------------------------------------------------------------------------
// 시험번호 찾기 (2단계 A3). 순수 함수.
// ---------------------------------------------------------------------------
function extractTestNumbers(texts) {
  const found = [];
  const re = new RegExp(CONFIG.TEST_NUM_REGEX.source, 'g');
  texts.forEach(t => {
    (String(t || '').match(re) || []).forEach(n => {
      const u = n.toUpperCase();
      if (!found.includes(u)) found.push(u);
    });
  });
  return found;
}

// "이름 <a@b.com>, c@d.com" → ['a@b.com', 'c@d.com']
function extractEmails(text) {
  return (String(text || '').match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || []).map(e => e.toLowerCase());
}

function emailDomain(email) {
  const i = String(email).lastIndexOf('@');
  return i < 0 ? '' : String(email).slice(i + 1).toLowerCase();
}

// 마스터의 담당자 메일 도메인이 domains 중 하나인 행의 시험번호 (종료·중단 행 제외)
// masterValues: 1행부터의 2차원 배열
function findTestNumbersByDomain(domains, masterValues, colMap, startRow) {
  const want = domains.filter(d => d && !CONFIG.SHARED_MAIL_DOMAINS.includes(d));
  const tnCol = colMap['시험번호'], mgrCol = colMap['담당자'], stCol = colMap['상태'];
  if (want.length === 0 || !tnCol || !mgrCol) return [];
  const found = [];
  for (let r = startRow - 1; r < masterValues.length; r++) {
    const row = masterValues[r];
    const tn = String(row[tnCol - 1] || '').trim().toUpperCase();
    if (!tn) continue;
    if (stCol && ['종료', '중단'].includes(String(row[stCol - 1] || '').trim())) continue;
    const rowDomains = extractEmails(row[mgrCol - 1]).map(emailDomain);
    if (rowDomains.some(d => want.includes(d)) && !found.includes(tn)) found.push(tn);
  }
  return found;
}

// 제목 → 같은 스레드의 다른 제목 → 발신(내가 보냈으면 받는 사람) 도메인 순서로 찾습니다.
// 후보가 하나일 때만 { testNumber, source }, 아니면 { testNumber: null, candidates, reason }.
function resolveTestNumber(opts) {
  const fromSubject = extractTestNumbers([opts.subject]);
  if (fromSubject.length === 1) return { testNumber: fromSubject[0], source: '제목' };
  if (fromSubject.length > 1) return { testNumber: null, candidates: fromSubject, reason: '제목에 시험번호가 여러 개' };

  const fromThread = extractTestNumbers(opts.threadSubjects || []);
  if (fromThread.length === 1) return { testNumber: fromThread[0], source: '같은 스레드 제목' };
  if (fromThread.length > 1) return { testNumber: null, candidates: fromThread, reason: '같은 스레드에 시험번호가 여러 개' };

  const addresses = opts.isMyEmail ? extractEmails(opts.to) : extractEmails(opts.from);
  const domains = addresses.map(emailDomain);
  // masterValues 는 배열 또는 배열을 돌려주는 함수 (필요할 때만 시트를 읽도록)
  const masterValues = typeof opts.masterValues === 'function' ? opts.masterValues() : (opts.masterValues || []);
  const fromDomain = findTestNumbersByDomain(domains, masterValues, opts.colMap || {}, opts.startRow || CONFIG.DATA_START_ROW);
  if (fromDomain.length === 1) return { testNumber: fromDomain[0], source: `메일 도메인(${domains.join(', ')})` };
  return {
    testNumber: null, candidates: fromDomain,
    reason: fromDomain.length === 0 ? '시험번호 후보 없음' : `도메인이 같은 진행 건이 여러 개`,
  };
}

function escapeRegExp(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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
