// 실행: node Flow/tests/run.test.js
// PASS  : 지금 코드가 기대대로 동작
// FAIL  : 회귀 — 고쳐야 함 (종료 코드 1)
// TODO  : 다음 단계에서 고칠 문제가 지금 재현됨 (예상된 실패)
// FIXED : TODO 였는데 통과함 → todo 표시를 지우세요
const assert = require('assert');
const { FakeSheet, fakeThread, loadFlow, lastLlmBody } = require('./harness');
const F = require('./fixtures');

const tests = [];
const test = (name, fn) => tests.push({ name, fn });
const todo = (stage, name, fn) => tests.push({ name, fn, todo: stage });

// 진행 메일용 기본 LLM 응답
const llm = (over) => Object.assign({
  시험번호: 'TTA-26-00001', 메일유형: '기타', 요약: '패치 일정 안내', 열갱신: [], 상태변경: null,
  할일: [], 확신도: '높음', 불확실한이유: null,
}, over || {});

// 3행 = TTA-26-00001 (진행 중)
const progressRow = (over) => Object.assign({ 시험번호: 'TTA-26-00001', 상태: '진행' }, over || {});

function setup({ dry, live, master, initThreads, gsThreads, llmResponse, props }) {
  const ss = F.makeSpreadsheet(live ? F.LIVE_ID : F.COPY_ID, master);
  const p = Object.assign({ GEMINI_API_KEY: 'test-key' }, props || {});
  if (!live) p.SPREADSHEET_ID = F.COPY_ID;
  if (dry) p.DRY_RUN = 'true';
  const flow = loadFlow({ props: p, spreadsheets: { [F.LIVE_ID]: ss, [F.COPY_ID]: ss }, initThreads, gsThreads, llmResponse });
  return { flow, ss, master: ss.getSheetByName('마스터'), log: ss.getSheetByName('처리로그'), queue: ss.getSheetByName('확인 큐') };
}
const simRows = ss => (ss.getSheetByName('시뮬레이션') ? ss.getSheetByName('시뮬레이션').data.slice(1) : []);
// 시뮬레이션 행 → 객체
const simObjs = ss => simRows(ss).map(r => ({ 메시지ID: r[1], 시험번호: r[2], 탭: r[3], 행: r[4], 열: r[5], 기존값: r[6], 새값: r[7], 사유: r[8] }));

// =========================================================================
// 0단계: 순수 함수 (지금 동작 고정)
// =========================================================================
test('parseInitiationEmail: 착수 메일에서 상담자·날짜·WD·시험원 추출', () => {
  const flow = loadFlow();
  const m = F.MAIL.init1;
  const r = flow.parseInitiationEmail(m.from, m.body, new Date(m.date));
  assert.deepStrictEqual({ ...r }, {
    상담자: '김상담', 시험안내일: '2026-11-05', 시험시작일: '2026-11-10', WD: '20', 시험원: '이시험\n박시험', 상태: '접수',
  });
});

test('parseBasicInfo: 제목에서 기업명, 발신자·서명에서 담당자 4줄', () => {
  const flow = loadFlow();
  const m = F.MAIL.progress1;
  const r = flow.parseBasicInfo(m.subject, m.from, m.body);
  assert.strictEqual(r['기업명'], '가나다소프트');
  assert.strictEqual(r['담당자'], '홍길동\nhong@ganada.example\n***_****_****\n선임연구원');
});

test('maskSensitiveData: 하이픈 휴대폰 번호와 IP 를 가림', () => {
  const flow = loadFlow();
  assert.strictEqual(flow.maskSensitiveData('전화 010-1234-5678, 서버 10.0.12.34'), '전화 ***-****-****, 서버 ***.***.***.***');
});

test('normalizeHeader / getColumnMap: 괄호 앞부분만 키로 사용', () => {
  const flow = loadFlow();
  assert.strictEqual(flow.normalizeHeader('담당자 (고객사 연락처/직함)'), '담당자');
  assert.strictEqual(flow.normalizeHeader('  시험번호 '), '시험번호');
  const map = flow.getColumnMap(F.makeMaster(), 1);
  assert.strictEqual(map['담당자'], F.col('담당자'));
  assert.strictEqual(map['담당자 (고객사 연락처/직함)'], undefined);
});

// =========================================================================
// 0단계: SPREADSHEET_ID · DRY_RUN
// =========================================================================
test('SPREADSHEET_ID 속성이 있으면 그 시트를 엽니다', () => {
  const live = F.makeSpreadsheet(F.LIVE_ID);
  const copy = F.makeSpreadsheet(F.COPY_ID);
  const flow = loadFlow({ props: { SPREADSHEET_ID: F.COPY_ID }, spreadsheets: { [F.LIVE_ID]: live, [F.COPY_ID]: copy },
                          initThreads: [fakeThread('t1', [F.MAIL.init1])] });
  const liveBefore = live.getSheetByName('마스터').snapshot();
  flow.runInitiationCheck();
  assert.strictEqual(live.getSheetByName('마스터').snapshot(), liveBefore, '실제 시트가 바뀌면 안 됨');
  assert.notStrictEqual(copy.getSheetByName('마스터').getLastRow(), 2, '사본에 새 행이 생겨야 함');
});

test('속성이 없으면 Config 의 시트 ID 를 씁니다', () => {
  const flow = loadFlow();
  assert.strictEqual(flow.getSpreadsheetId(), F.LIVE_ID);
  assert.strictEqual(flow.isDryRun(), false);
});

test('INIT_SEARCH_QUERY 속성으로 착수 메일 검색어를 바꿀 수 있습니다 (지난 메일 재처리용)', () => {
  const q = 'subject:"GS 시험 인증 시작 안내" newer_than:60d';
  const { flow } = setup({ dry: true, props: { INIT_SEARCH_QUERY: q } });
  flow.runInitiationCheck();
  assert.deepStrictEqual(flow.searches, [q]);
});

test('DRY_RUN 착수: 마스터·처리로그는 그대로, 시뮬레이션 탭을 만들고 쓰려던 칸을 기록', () => {
  const { flow, ss, master, log } = setup({ dry: true, initThreads: [fakeThread('t1', [F.MAIL.init1])] });
  const mBefore = master.snapshot(), lBefore = log.snapshot();
  flow.runInitiationCheck();
  assert.strictEqual(master.snapshot(), mBefore);
  assert.strictEqual(log.snapshot(), lBefore);
  assert.deepStrictEqual(ss.inserted, ['시뮬레이션']);
  assert.deepStrictEqual(ss.getSheetByName('시뮬레이션').data[0], [...flow.SIM_HEADERS]);
  const recs = simObjs(ss);
  const byCol = Object.fromEntries(recs.filter(r => r.탭 === '마스터').map(r => [r.열, r]));
  assert.strictEqual(byCol['상담자'].새값, '김상담');
  assert.strictEqual(byCol['시험시작일'].새값, '2026-11-10');
  assert.strictEqual(byCol['상태'].행, '3');
  assert.ok(recs.some(r => r.탭 === '처리로그' && r.사유.includes('기록하지 않음')));
});

test('DRY_RUN 인데 실제 시트이고 시뮬레이션 탭이 없으면 아무것도 하지 않음', () => {
  const { flow, ss, master } = setup({ dry: true, live: true, initThreads: [fakeThread('t1', [F.MAIL.init1])] });
  const before = master.snapshot();
  flow.runInitiationCheck();
  assert.strictEqual(master.snapshot(), before);
  assert.deepStrictEqual(ss.inserted, []);
  assert.ok(flow.logs.some(l => l.startsWith('오류: DRY_RUN')));
});

test('DRY_RUN 진행: 마스터·확인 큐·처리로그 그대로, GitHub 호출 없음, LLM 응답 기록', () => {
  const { flow, ss, master, log, queue } = setup({
    dry: true, props: { GITHUB_TOKEN: 'x' },
    master: F.makeMaster({ rows: [progressRow()] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 열갱신: [{ 열: '패치', 값: '2026-11-20' }] }),
  });
  const snap = [master.snapshot(), log.snapshot(), queue.snapshot()];
  flow.runProgressCheck();
  assert.deepStrictEqual([master.snapshot(), log.snapshot(), queue.snapshot()], snap);
  assert.ok(!flow.fetches.some(f => f.url.includes('api.github.com')), 'DRY_RUN 에서 GitHub 호출 금지');
  const recs = simObjs(ss);
  const patch = recs.find(r => r.열 === '패치');
  assert.ok(patch && patch.새값 === '2026-11-20' && patch.행 === '3' && patch.사유 === 'LLM: 열갱신');
  assert.ok(recs.some(r => r.열 === '진행상황' && r.새값 === '[11/06] 패치 일정 안내'));
  assert.ok(recs.some(r => r.열 === '기업명' && r.새값 === '가나다소프트' && r.사유 === 'STEP A: 빈칸 채움'));
  assert.ok(recs.some(r => r.탭 === 'LLM' && r.새값.includes('"열갱신"')));
});

test('DRY_RUN 진행: 시트에 행이 없는 시험번호는 메모만 남기고 LLM 을 부르지 않음', () => {
  const { flow, ss } = setup({ dry: true, gsThreads: [fakeThread('th9', [F.MAIL.noRow])], llmResponse: llm() });
  flow.runProgressCheck();
  assert.strictEqual(flow.fetches.length, 0);
  assert.ok(simObjs(ss).some(r => r.시험번호 === 'TTA-26-09999' && r.새값 === '시트에 시험번호 행 없음'));
});

test('DRY_RUN 은 같은 실행 안에서 앞 메일이 "쓴 셈 친" 값을 다음 메일이 읽음 (실제 실행과 같은 판단)', () => {
  const run = dry => {
    const s = setup({
      dry, master: F.makeMaster({ rows: [progressRow()] }),
      gsThreads: [fakeThread('th1', [F.MAIL.progress1, F.MAIL.progress2])],
      llmResponse: p => (p.contents[0].parts[0].text.includes('패치는')
        ? llm({ 열갱신: [{ 열: '패치', 값: '2026-11-20' }] })
        : llm({ 요약: '일정 문의', 열갱신: [{ 열: '패치', 값: '2026-11-21' }] })),
    });
    s.flow.runProgressCheck();
    return s;
  };
  const real = run(false);
  assert.strictEqual(real.queue.getLastRow(), 2, '실제 실행: 둘째 메일은 패치 충돌로 확인 큐 (머리글 + 1행)');
  const dry = run(true);
  const recs = simObjs(dry.ss);
  assert.strictEqual(recs.filter(r => r.열 === '기업명').length, 1, 'STEP A 는 한 번만');
  assert.ok(recs.some(r => r.탭 === '확인 큐' && r.사유.includes('[패치] 덮어쓰기 충돌')), 'DRY 에서도 둘째 메일은 확인 큐');
});

test('실제 모드(DRY_RUN 끔)는 예전처럼 마스터·처리로그에 씀', () => {
  const { flow, ss, master, log } = setup({
    master: F.makeMaster({ rows: [progressRow()] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 열갱신: [{ 열: '패치', 값: '2026-11-20' }] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('패치')), '2026-11-20');
  assert.strictEqual(master.get(3, F.col('진행상황')), '[11/06] 패치 일정 안내');
  assert.strictEqual(log.get(2, 3), 'LLM: 시트 반영 완료 (TTA-26-00001)');
  assert.strictEqual(ss.getSheetByName('시뮬레이션'), null);
});

// =========================================================================
// 1단계 (C1~C4) — v1.3 에서 고침
// =========================================================================
test('[C1] 착수 메일 2건 → F열 수식이 978행까지 있어도 3행·4행에 들어감', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ formulaTo: 978 }),
    initThreads: [fakeThread('t1', [F.MAIL.init1]), fakeThread('t2', [F.MAIL.init2])],
  });
  flow.runInitiationCheck();
  assert.strictEqual(master.get(3, F.col('상담자')), '김상담', `3행이 비어 있음 (실제로는 ${master.getLastRow() - 1}행 근처에 씀)`);
  assert.strictEqual(master.get(4, F.col('상담자')), '최상담');
});

test('[C1] 시험번호가 빈 착수 행(3행) 다음 착수 메일은 4행에 들어가고 3행을 덮지 않음', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ formulaTo: 978, rows: [{ 상담자: '김상담', 시험안내일: '2026-11-04', 상태: '접수' }] }),
    initThreads: [fakeThread('t2', [F.MAIL.init2])],
  });
  flow.runInitiationCheck();
  assert.strictEqual(master.get(3, F.col('상담자')), '김상담');
  assert.strictEqual(master.get(4, F.col('상담자')), '최상담');
});

test('[C2] getRowDataAsObject 키도 normalizeHeader 를 거침 (LLM 프롬프트의 키 = colMap 키)', () => {
  const flow = loadFlow();
  const master = F.makeMaster({ rows: [progressRow({ 담당자: '홍길동' })] });
  const obj = flow.getRowDataAsObject(master, 3, master.data[0]);
  assert.strictEqual(obj['담당자'], '홍길동');
});

test('[C2·C3] 담당자가 채워진 행에 LLM 이 담당자를 제안해도 덮어쓰지 않음', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ rows: [progressRow({ 담당자: '홍길동\nhong@ganada.example' })] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 열갱신: [{ 열: '담당자', 값: '***-****-****' }] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('담당자')), '홍길동\nhong@ganada.example');
});

test('[C3] ALLOWED_COLUMNS 에 담당자·진행상황이 없음', () => {
  const flow = loadFlow();
  assert.ok(!flow.ALLOWED_COLUMNS.includes('담당자'));
  assert.ok(!flow.ALLOWED_COLUMNS.includes('진행상황'));
});

test('[C3] LLM 이 허용 목록 밖의 열(잔여일자 수식 열·진행상황)을 제안하면 쓰지 않음', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ rows: [progressRow()] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 열갱신: [{ 열: '잔여일자', 값: '0' }, { 열: '진행상황', 값: '덮어쓰기' }] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('잔여일자')), '');
  assert.strictEqual(master.get(3, F.col('진행상황')), '[11/06] 패치 일정 안내');
});

test('[C3] 허용값이 아닌 상태변경(예: "완료")은 쓰지 않음', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ rows: [progressRow()] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 상태변경: '완료' }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('상태')), '진행');
});

F.QUOTE_CASES.forEach(c => {
  test(`[C4] cleanEmailBody: ${c.name}`, () => {
    const flow = loadFlow();
    assert.strictEqual(flow.cleanEmailBody(c.body), c.expected);
  });
});

test('[C4] LLM 에 보내는 본문에서 인용문이 잘려 있음', () => {
  const quoted = Object.assign({}, F.MAIL.progress1, {
    body: '패치 일정 공유드립니다.\n\n-----Original Message-----\nFrom: 시험원\n이전 메일 1차결함 2026-10-01\n',
  });
  const { flow } = setup({ master: F.makeMaster({ rows: [progressRow()] }), gsThreads: [fakeThread('th1', [quoted])], llmResponse: llm() });
  flow.runProgressCheck();
  assert.ok(!lastLlmBody(flow).includes('이전 메일'));
});

test('[C4] 자르고 나서 본문이 거의 비면 LLM 을 부르지 않고 사유를 남김', () => {
  const empty = Object.assign({}, F.MAIL.progress1, { body: '\n\n-----Original Message-----\nFrom: 시험원\n이전 메일\n' });
  const { flow, log } = setup({ master: F.makeMaster({ rows: [progressRow()] }), gsThreads: [fakeThread('th1', [empty])], llmResponse: llm() });
  flow.runProgressCheck();
  assert.strictEqual(flow.fetches.length, 0);
  assert.ok(String(log.get(2, 3)).includes('본문'));
});

// v1.1 에 적혀 있었지만 호출되지 않던 것 — v1.3 에서 연결
test('[마스킹] LLM 에 보내는 본문의 점으로 쓴 전화번호도 가림', () => {
  const { flow } = setup({ master: F.makeMaster({ rows: [progressRow()] }), gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm() });
  flow.runProgressCheck();
  assert.ok(!lastLlmBody(flow).includes('010.1234.5678'));
});

test('[날짜] getRowDataAsObject 가 Date 를 yyyy-MM-dd(Asia/Seoul) 로 바꿈', () => {
  const flow = loadFlow();
  const master = F.makeMaster({ rows: [progressRow({ 시험시작일: F.kstDate('2026-11-10') })] });
  const obj = flow.getRowDataAsObject(master, 3, master.data[0]);
  assert.strictEqual(obj['시험시작일'], '2026-11-10');
});

// =========================================================================
// 1단계에서 추가한 확인
// =========================================================================
test('[C1] F열 수식이 빈 행에도 값(숫자·오류)을 보여도 3행에 들어가고, F열 수식 칸은 건드리지 않음', () => {
  const master = F.makeMaster({ formulaTo: 978 });
  for (let r = 2; r <= 978; r++) master.set(r, F.col('잔여일자'), r === 3 ? 10 : '#VALUE!');
  const { flow } = setup({ master, initThreads: [fakeThread('t1', [F.MAIL.init1])] });
  flow.runInitiationCheck();
  assert.strictEqual(master.get(3, F.col('상담자')), '김상담');
  assert.strictEqual(master.get(3, F.col('잔여일자')), 10, 'F열에 "" 를 덮어쓰면 안 됨');
});

test('[C1] findNewRowIndex: 수식 열 무시, 데이터 시작 행(3) 아래로만', () => {
  const flow = loadFlow();
  const v = [['h1', 'h2'], ['예시', ''], ['', 5], ['', 7]];
  assert.strictEqual(flow.findNewRowIndex(v, [2], 3), 3);
  assert.strictEqual(flow.findNewRowIndex(v, [], 3), 5);
  assert.strictEqual(flow.findNewRowIndex([['h']], [], 3), 3);
});

test('[C1] DRY_RUN 에서도 착수 메일 2건이 3행·4행으로 기록됨', () => {
  const { flow, ss, master } = setup({
    dry: true, master: F.makeMaster({ formulaTo: 978 }),
    initThreads: [fakeThread('t1', [F.MAIL.init1]), fakeThread('t2', [F.MAIL.init2])],
  });
  const before = master.snapshot();
  flow.runInitiationCheck();
  assert.strictEqual(master.snapshot(), before);
  const rows = simObjs(ss).filter(r => r.열 === '상담자').map(r => [r.행, r.새값]);
  assert.deepStrictEqual(rows, [['3', '김상담'], ['4', '최상담']]);
});

test('[C2] 착수 메일: 괄호가 붙은 헤더(예: "시험시작일 (예정)")에도 값이 들어감', () => {
  const master = F.makeMaster();
  master.set(1, F.col('시험시작일'), '시험시작일 (예정)');
  const { flow } = setup({ master, initThreads: [fakeThread('t1', [F.MAIL.init1])] });
  flow.runInitiationCheck();
  assert.strictEqual(master.get(3, F.col('시험시작일')), '2026-11-10');
});

test('[C2·C3] checkLlmUpdates: 괄호 붙은 열 이름 정규화, 형식·허용값·충돌 검사', () => {
  const flow = loadFlow();
  const colMap = flow.getColumnMap(F.makeMaster(), 1);
  const row = { 상태: '진행', 시험종료일: '2026-12-01', 패치: null, '1차결함': null };
  const r = JSON.parse(JSON.stringify(flow.checkLlmUpdates({
    열갱신: [
      { 열: '패치 (예정일)', 값: '2026-11-20' },   // 정규화 → 반영
      { 열: '1차결함', 값: '11월 3일' },           // 날짜 형식 아님 → 확인 큐
      { 열: '시험종료일', 값: '2026-12-05' },      // 기존값과 다름 → 확인 큐
      { 열: '담당자', 값: '***-****-****' },       // 허용되지 않은 열 → 무시
      { 열: '시험번호', 값: 'TTA-26-99999' },      // 허용되지 않은 열 → 무시
    ],
    상태변경: '진행',                              // 이미 같은 상태 → 건너뜀
  }, row, colMap)));
  assert.deepStrictEqual(r.updates.map(u => [u.colName, u.value]), [['패치', '2026-11-20']]);
  assert.strictEqual(r.queueReasons.length, 2);
  assert.ok(r.queueReasons[0].includes('1차결함') && r.queueReasons[1].includes('시험종료일'));
  assert.deepStrictEqual(r.ignored.map(i => i.colName), ['담당자', '시험번호']);
});

test('[C3] 허용 상태값으로의 변경은 반영됨 (진행 → 종료)', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ rows: [progressRow()] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 상태변경: '종료' }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('상태')), '종료');
});

[
  { name: 'Gmail 한국어가 두 줄로 꺾인 경우', body: '네 알겠습니다.\n\n2026년 11월 5일 (목) 오후 3:00, 홍길동 <a@b.example>님이\n작성:\n> 이전', expected: '네 알겠습니다.' },
  { name: 'Outlook 한국어 머리말 묶음 (보낸 사람/보낸 날짜)', body: '수정본 첨부합니다.\n\n보낸 사람: 이시험 <t@tta.example>\n보낸 날짜: 2026년 11월 4일 화요일 오전 10:00\n받는 사람: 홍길동\n제목: 1차 결함\n\n이전 본문', expected: '수정본 첨부합니다.' },
  { name: '줄 맨 앞 ">" 인용', body: '좋습니다.\n> 11/20 패치 가능할까요?', expected: '좋습니다.' },
  { name: '본문 중 "보낸 사람:" 한 줄만 있으면 자르지 않음', body: '보낸 사람: 홍길동 (담당자)\n패치는 11/20 입니다.', expected: '보낸 사람: 홍길동 (담당자)\n패치는 11/20 입니다.' },
].forEach(c => test(`[C4] cleanEmailBody: ${c.name}`, () => {
  assert.strictEqual(loadFlow().cleanEmailBody(c.body), c.expected);
}));

test('[C4] STEP A 도 인용문을 자른 본문으로 파싱 (인용된 시험원 서명을 담당자 직급으로 읽지 않음)', () => {
  const reply = Object.assign({}, F.MAIL.progress1, {
    from: 'hong@ganada.example',
    body: '패치 일정 공유드립니다.\n\n-----Original Message-----\nFrom: 이시험\n이시험 선임\n',
  });
  const { flow, master } = setup({ master: F.makeMaster({ rows: [progressRow()] }), gsThreads: [fakeThread('th1', [reply])], llmResponse: llm() });
  flow.runProgressCheck();
  assert.ok(!String(master.get(3, F.col('담당자'))).includes('선임'));
});

test('[마스킹] 여러 전화번호 형식과 IP 를 가림', () => {
  const flow = loadFlow();
  const s = flow.maskSensitiveData('a 010-1234-5678 b 010.1234.5678 c 01012345678 d 02-123-4567 e +82 10-1234-5678 f 192.168.0.10 g 2026-11-20');
  assert.ok(!/\d{4}[-.\s]?\d{4}\b/.test(s.replace('2026-11-20', '')), s);
  assert.ok(s.includes('2026-11-20'), '날짜는 가리면 안 됨: ' + s);
});

// =========================================================================
// 2단계 (A1~A8)
// =========================================================================
// 같은 사본(ss)으로 진행 메일 처리를 한 번 더 돌림 (재시도·재실행 확인용)
function progressRun(ss, { gsThreads, llmResponse, props, dry }) {
  const p = Object.assign({ GEMINI_API_KEY: 'k', SPREADSHEET_ID: F.COPY_ID }, props || {});
  if (dry) p.DRY_RUN = 'true';
  const flow = loadFlow({ props: p, spreadsheets: { [F.COPY_ID]: ss }, gsThreads, llmResponse });
  flow.runProgressCheck();
  return flow;
}
const logRows = log => log.data.slice(1).filter(r => r[0]).map(r => ({ id: r[0], 결과: r[2], 상태: r[3], 시도: r[4] }));
// 확인 큐 계약 (B2) — 머리글을 바꾸면 이 테스트도 바뀌어야 합니다
const QUEUE_COLS = ['ID', '생성시각', '시험번호', '메시지ID', '메일 링크', '열', '기존값', '제안값', '사유', 'LLM 원문', '처리', '처리시각'];
const queueRows = q => q.data.slice(1).filter(r => r.some(v => v !== ''))
  .map(r => Object.fromEntries(QUEUE_COLS.map((h, i) => [h, r[i] === undefined ? '' : r[i]])));

test('[A1] 같은 날짜는 충돌이 아님 (시트의 Date 와 LLM 의 문자열 비교)', () => {
  const { flow, master, queue, log } = setup({
    master: F.makeMaster({ rows: [progressRow({ 패치: F.kstDate('2026-11-20') })] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 열갱신: [{ 열: '패치', 값: '2026-11-20' }, { 열: '시험종료일', 값: '2026-12-10' }] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(queueRows(queue).length, 0);
  assert.strictEqual(master.get(3, F.col('시험종료일')), '2026-12-10');
  assert.strictEqual(logRows(log)[0].상태, '완료');
});

test('[A1] 충돌한 열만 확인 큐로, 나머지 열과 요약은 반영', () => {
  const { flow, master, queue, log } = setup({
    master: F.makeMaster({ rows: [progressRow({ 패치: '2026-11-18' })] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 열갱신: [{ 열: '패치', 값: '2026-11-20' }, { 열: '시험종료일', 값: '2026-12-10' }] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('패치')), '2026-11-18', '충돌 열은 그대로');
  assert.strictEqual(master.get(3, F.col('시험종료일')), '2026-12-10', '충돌 없는 열은 반영');
  assert.strictEqual(master.get(3, F.col('진행상황')), '[11/06] 패치 일정 안내');
  const q = queueRows(queue);
  assert.strictEqual(q.length, 1);
  assert.strictEqual(q[0].열, '패치');
  assert.strictEqual(q[0].기존값, '2026-11-18');
  assert.strictEqual(q[0].제안값, '2026-11-20');
  assert.ok(q[0].사유.includes('덮어쓰기 충돌'));
  assert.strictEqual(logRows(log)[0].상태, '확인대기');
});

test('[A1] 확신도 낮음이면 아무것도 반영하지 않고 확인 큐', () => {
  const { flow, master, queue } = setup({
    master: F.makeMaster({ rows: [progressRow()] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 확신도: '낮음', 불확실한이유: '날짜가 두 개', 열갱신: [{ 열: '패치', 값: '2026-11-20' }] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('패치')), '');
  assert.strictEqual(master.get(3, F.col('진행상황')), '');
  const q = queueRows(queue);
  assert.deepStrictEqual(q.map(r => [r.열, r.제안값]), [['패치', '2026-11-20']]);
  assert.ok(q[0].사유.includes('확신도 낮음: 날짜가 두 개'));
});

test('[A2] classifyGeminiResponse: 실패 종류 구분', () => {
  const flow = loadFlow();
  const ok = JSON.stringify({ candidates: [{ content: { parts: [{ text: '{"요약":"a"}' }] }, finishReason: 'STOP' }] });
  const kind = (c, t) => { const r = flow.classifyGeminiResponse(c, t); return r.ok ? 'ok' : r.kind; };
  assert.strictEqual(kind(200, ok), 'ok');
  assert.strictEqual(kind(429, '{}'), 'transient');
  assert.strictEqual(kind(503, '<html>'), 'transient');
  assert.strictEqual(kind(400, '{"error":{"message":"bad schema"}}'), 'fatal');
  assert.strictEqual(kind(403, '{"error":{"message":"key"}}'), 'fatal');
  assert.strictEqual(kind(200, '<html>'), 'parse');
  assert.strictEqual(kind(200, JSON.stringify({ candidates: [{ content: { parts: [{ text: '{깨짐' }] } }] })), 'parse');
  assert.strictEqual(kind(200, JSON.stringify({ candidates: [{ finishReason: 'SAFETY' }] })), 'blocked');
  assert.strictEqual(kind(200, JSON.stringify({ promptFeedback: { blockReason: 'OTHER' } })), 'blocked');
});

test('[A2] 일시 오류는 같은 처리로그 행에서 3회까지 다시 시도한 뒤 포기', () => {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  const log = ss.getSheetByName('처리로그');
  const run = () => progressRun(ss, { gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: { __status: 503 } });
  run();
  assert.deepStrictEqual(logRows(log).map(r => [r.상태, r.시도]), [['재시도대기', 1]]);
  run();
  assert.deepStrictEqual(logRows(log).map(r => [r.상태, r.시도]), [['재시도대기', 2]]);
  run();
  assert.deepStrictEqual(logRows(log).map(r => [r.상태, r.시도]), [['포기', 3]]);
  const f = run();
  assert.strictEqual(f.fetches.length, 0, '포기한 메일은 다시 부르지 않음');
  assert.deepStrictEqual([log.get(1, 4), log.get(1, 5)], ['상태', '시도횟수'], 'D1·E1 헤더');
});

test('[A2] 재시도 중 성공하면 같은 행이 완료로 바뀜', () => {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  const log = ss.getSheetByName('처리로그');
  progressRun(ss, { gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: { __raw: '{깨진 JSON' } });
  assert.strictEqual(logRows(log)[0].상태, '재시도대기');
  progressRun(ss, { gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm() });
  assert.deepStrictEqual(logRows(log).map(r => [r.상태, r.시도]), [['완료', 2]]);
});

test('[A2] 안전 차단은 바로 포기', () => {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  progressRun(ss, { gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: { __finish: 'SAFETY' } });
  assert.deepStrictEqual(logRows(ss.getSheetByName('처리로그')).map(r => [r.상태, r.시도]), [['포기', 1]]);
});

test('[A2] 키·모델 오류(4xx)는 처리로그 없이 실행을 멈춤 (다른 메일도 버리지 않음)', () => {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  const f = progressRun(ss, { gsThreads: [fakeThread('th1', [F.MAIL.progress1]), fakeThread('th2', [F.MAIL.progress2])], llmResponse: { __status: 403 } });
  assert.strictEqual(f.fetches.length, 1);
  assert.strictEqual(logRows(ss.getSheetByName('처리로그')).length, 0);
  assert.ok(f.logs.some(l => l.includes('실행을 멈춥니다')));
});

test('[A2] 상태 열이 빈 예전 처리로그 행은 완료로 봄', () => {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  ss.getSheetByName('처리로그').appendRow(['prog-1', '2026-11-01', 'LLM: 시트 반영 완료']);
  const f = progressRun(ss, { gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm() });
  assert.strictEqual(f.fetches.length, 0);
});

test('[A2] 처리로그 D1 에 다른 값이 있으면 멈춤 / DRY_RUN 은 헤더를 쓰지 않음', () => {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  ss.getSheetByName('처리로그').set(1, 4, '비고');
  const f = progressRun(ss, { gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm() });
  assert.strictEqual(f.fetches.length, 0);
  assert.ok(f.logs.some(l => l.includes("D1 에 '상태' 대신 '비고'")));
  const ss2 = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  progressRun(ss2, { dry: true, gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm() });
  assert.strictEqual(ss2.getSheetByName('처리로그').get(1, 4), '');
});

const noNumber = (over) => Object.assign({}, F.MAIL.progress1, { id: 'prog-x', subject: 'Re: 패치 일정 문의' }, over || {});

test('[A3] 제목에 없으면 같은 스레드의 다른 메일 제목에서 찾음', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ rows: [progressRow()] }),
    gsThreads: [fakeThread('th1', [Object.assign({}, F.MAIL.progress2, { id: 'first' }), noNumber()])],
    llmResponse: llm({ 열갱신: [{ 열: '패치', 값: '2026-11-20' }] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('패치')), '2026-11-20');
});

test('[A3] 스레드에도 없으면 발신 도메인을 담당자 메일 도메인과 맞춤', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ rows: [progressRow({ 담당자: '홍길동\nhong@ganada.example' }), { 시험번호: 'TTA-26-00002', 상태: '진행', 담당자: 'x@other.example' }] }),
    gsThreads: [fakeThread('th1', [noNumber({ from: '김부장 <kim@ganada.example>' })])],
    llmResponse: llm({ 열갱신: [{ 열: '패치', 값: '2026-11-20' }] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('패치')), '2026-11-20');
});

test('[A3] 내가 보낸 메일은 받는 사람 도메인으로 맞춤', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ rows: [progressRow({ 담당자: 'hong@ganada.example' })] }),
    gsThreads: [fakeThread('th1', [noNumber({ from: '나 <hojun7766@tta.or.kr>', to: '홍길동 <hong@ganada.example>' })])],
    llmResponse: llm({ 열갱신: [{ 열: '패치', 값: '2026-11-20' }] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('패치')), '2026-11-20');
});

test('[A3] 후보가 없거나(공용 메일) 여러 개면 확인 큐 + 확인대기, 다음 주기에 다시 처리하지 않음', () => {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [
    progressRow({ 담당자: 'a@ganada.example' }), { 시험번호: 'TTA-26-00002', 상태: '진행', 담당자: 'b@ganada.example' },
    { 시험번호: 'TTA-25-00003', 상태: '종료', 담당자: 'c@ganada.example' },
    { 시험번호: 'TTA-26-00004', 상태: '진행', 담당자: '개인메일 업체 x@gmail.com' }] }));  // 공용 도메인은 맞추지 않음
  const threads = [
    fakeThread('th1', [noNumber({ id: 'm-gmail', from: '개인 <someone@gmail.com>' })]),
    fakeThread('th2', [noNumber({ id: 'm-multi', from: 'd@ganada.example' })]),
  ];
  progressRun(ss, { gsThreads: threads, llmResponse: llm() });
  const q = queueRows(ss.getSheetByName('확인 큐'));
  assert.strictEqual(q.length, 2);
  assert.ok(q[0].열 === '시험번호' && q[0].사유.includes('후보 없음') && q[0].시험번호 === '');
  assert.ok(q[1].사유.includes('여러 개') && q[1].제안값 === 'TTA-26-00001, TTA-26-00002');
  assert.deepStrictEqual(logRows(ss.getSheetByName('처리로그')).map(r => r.상태), ['확인대기', '확인대기']);
  const f = progressRun(ss, { gsThreads: threads, llmResponse: llm() });
  assert.strictEqual(queueRows(ss.getSheetByName('확인 큐')).length, 2, '두 번째 실행에서 다시 올리지 않음');
  assert.strictEqual(f.fetches.length, 0);
});

test('[A4] 내가 보낸 메일이면 STEP A(년도·기업명·담당자)를 건너뜀', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ rows: [progressRow()] }),
    gsThreads: [fakeThread('th1', [Object.assign({}, F.MAIL.progress1, { from: '이시험 <HOJUN7766@tta.or.kr>' })])],
    llmResponse: llm(),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('담당자')), '');
  assert.strictEqual(master.get(3, F.col('기업명')), '');
});

test('[A5] 이름에 정규식 특수문자가 있어도 멈추지 않음', () => {
  const flow = loadFlow();
  const r = flow.parseBasicInfo('[GS] TTA-26-00001 가나다', '"(주)홍길동 [영업*]" <h@ganada.example>', '(주)홍길동 [영업*] 팀장\n');
  assert.strictEqual(r['담당자'].split('\n').pop(), '팀장');
});

test('[A6] 년도는 시험번호에서 (TTA-25-… → 2025), 없으면 빈칸', () => {
  const flow = loadFlow();
  assert.strictEqual(flow.parseBasicInfo('[GS] TTA-25-00010 가나다', 'a <a@b.example>', '').년도, 2025);
  assert.strictEqual(flow.parseBasicInfo('문의', 'a <a@b.example>', '', 'TTA-24-00001').년도, 2024);
  assert.strictEqual(flow.parseBasicInfo('문의', 'a <a@b.example>', '').년도, '');
});

test('[A7] 프롬프트: 메일유형 후보, 허용 열과 그 현재 값만, responseSchema', () => {
  const { flow } = setup({
    master: F.makeMaster({ rows: [progressRow({ 담당자: '홍길동\nhong@ganada.example', 시험시작일: F.kstDate('2026-11-10') })] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm(),
  });
  flow.runProgressCheck();
  const payload = JSON.parse(flow.fetches[0].options.payload);
  const sys = payload.systemInstruction.parts[0].text;
  assert.ok(sys.includes('착수, 1차결함, 2차결함'));
  assert.ok(sys.includes('"시험시작일":"2026-11-10"'));
  assert.ok(!sys.includes('hong@ganada.example') && !sys.includes('담당자'), '허용 열이 아닌 값은 보내지 않음');
  const schema = payload.generationConfig.responseSchema;
  assert.strictEqual(schema.type, 'OBJECT');
  assert.ok(schema.properties['열갱신'].items.properties['열'].enum.includes('패치'));
  assert.ok(!schema.properties['열갱신'].items.properties['열'].enum.includes('상태'));
});

test('[A8] 날짜 순서가 어긋난 열만 확인 큐로', () => {
  const { flow, master, queue } = setup({
    master: F.makeMaster({ rows: [progressRow({ 시험시작일: F.kstDate('2026-11-10'), 시험종료일: '2026-12-10' })] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 열갱신: [
      { 열: '1차결함', 값: '2026-11-05' },   // 시작일보다 앞 → 확인 큐
      { 열: '패치', 값: '2026-12-20' },      // 종료일보다 뒤 → 확인 큐
      { 열: '2차결함', 값: '2026-11-25' },   // 문제 없음 → 반영
    ] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('1차결함')), '');
  assert.strictEqual(master.get(3, F.col('패치')), '');
  assert.strictEqual(master.get(3, F.col('2차결함')), '2026-11-25');
  const q = queueRows(queue);
  assert.deepStrictEqual(q.map(r => r.열), ['1차결함', '패치']);
  assert.ok(q.every(r => r.사유.includes('날짜 순서 어긋남')));
});

test('normalizeYmd: 여러 날짜 표기 → yyyy-MM-dd, 없는 날짜는 null', () => {
  const flow = loadFlow();
  assert.strictEqual(flow.normalizeYmd('2026.11.5'), '2026-11-05');
  assert.strictEqual(flow.normalizeYmd('2026/11/05'), '2026-11-05');
  assert.strictEqual(flow.normalizeYmd('2026. 11. 5.'), '2026-11-05');
  assert.strictEqual(flow.normalizeYmd(F.kstDate('2026-11-05')), '2026-11-05');
  assert.strictEqual(flow.normalizeYmd('2026-02-30'), null);
  assert.strictEqual(flow.normalizeYmd('11월 5일'), null);
});

// =========================================================================
// 3단계 (B1~B4)
// =========================================================================
const TODO_COLS = ['ID', '시험번호', '내용', '기한', '출처 메시지ID', '생성시각', '상태'];
const todoRows = t => (t ? t.data.slice(1).filter(r => r.some(v => v !== ''))
  .map(r => Object.fromEntries(TODO_COLS.map((h, i) => [h, r[i] === undefined ? '' : r[i]]))) : []);
const todos = [{ 내용: '패치본 수령 후 재시험', 기한: '2026.11.21' }, { 내용: '결함리포트 회신 확인', 기한: null }];

test('[B1] LLM 할일 → 할 일 탭 (없으면 만들고 머리글·상태 드롭다운), GitHub 는 부르지 않음', () => {
  const { flow, ss } = setup({
    props: { GITHUB_TOKEN: 'x' },
    master: F.makeMaster({ rows: [progressRow()] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 할일: todos }),
  });
  flow.runProgressCheck();
  const t = ss.getSheetByName('할 일');
  assert.deepStrictEqual(t.data[0], TODO_COLS);
  assert.deepStrictEqual([...t.validations[0].rule.list], ['신규', '반영', '무시']);
  assert.deepStrictEqual(todoRows(t).map(r => [r.ID, r.시험번호, r.내용, r.기한, r['출처 메시지ID'], r.상태]), [
    ['T-prog-1-1', 'TTA-26-00001', '패치본 수령 후 재시험', '2026-11-21', 'prog-1', '신규'],
    ['T-prog-1-2', 'TTA-26-00001', '결함리포트 회신 확인', '', 'prog-1', '신규'],
  ]);
  assert.ok(!flow.fetches.some(f => f.url.includes('api.github.com')));
});

test('[B1] 같은 시험번호·내용·기한의 할 일은 다시 넣지 않음 (다음 실행·다른 메일이어도)', () => {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  progressRun(ss, { gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm({ 할일: todos }) });
  progressRun(ss, { gsThreads: [fakeThread('th1', [F.MAIL.progress1, F.MAIL.progress2])],
                    llmResponse: llm({ 할일: [{ 내용: '패치본 수령 후 재시험', 기한: '2026-11-21' }, { 내용: '새 할 일', 기한: '2026-11-30' }] }) });
  assert.deepStrictEqual(todoRows(ss.getSheetByName('할 일')).map(r => r.내용), ['패치본 수령 후 재시험', '결함리포트 회신 확인', '새 할 일']);
});

test('[B1] 확신도 낮음이면 할 일도 넣지 않음 / DRY_RUN 은 할 일 탭을 만들지 않고 기록만', () => {
  const low = setup({ master: F.makeMaster({ rows: [progressRow()] }), gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
                      llmResponse: llm({ 확신도: '낮음', 할일: todos }) });
  low.flow.runProgressCheck();
  assert.strictEqual(todoRows(low.ss.getSheetByName('할 일')).length, 0);
  const dry = setup({ dry: true, master: F.makeMaster({ rows: [progressRow()] }), gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
                      llmResponse: llm({ 할일: todos }) });
  dry.flow.runProgressCheck();
  assert.strictEqual(dry.ss.getSheetByName('할 일'), null);
  assert.strictEqual(simObjs(dry.ss).filter(r => r.탭 === '할 일').length, 2);
});

test('[B2] 확인 큐: 머리글·처리 드롭다운, 변경 1건당 1행, 보관함도 열리는 #all 링크', () => {
  const { flow, queue } = setup({
    master: F.makeMaster({ rows: [progressRow({ 패치: '2026-11-18', 시험종료일: '2026-12-01' })] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 열갱신: [{ 열: '패치', 값: '2026-11-20' }, { 열: '시험종료일', 값: '2026-12-05' }] }),
  });
  flow.runProgressCheck();
  assert.deepStrictEqual(queue.data[0], QUEUE_COLS);
  assert.deepStrictEqual([...queue.validations[0].rule.list], ['대기', '적용', '무시']);
  const q = queueRows(queue);
  assert.deepStrictEqual(q.map(r => [r.ID, r.열, r.처리]), [['Q-prog-1-1', '패치', '대기'], ['Q-prog-1-2', '시험종료일', '대기']]);
  assert.strictEqual(q[0]['메일 링크'], 'https://mail.google.com/mail/u/0/#all/th1');
  assert.ok(q[0]['LLM 원문'].includes('"열갱신"'));
});

test('[B2] 예전 형식 확인 큐는 "확인 큐 (v1)" 로 보관하고 새로 만듦', () => {
  const old = new FakeSheet('확인 큐', [['TTA-26-00001', 'https://mail…', '[패치] 덮어쓰기 충돌', '{}', '대기']]);
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  ss.sheets['확인 큐'] = old; old.ss = ss;
  progressRun(ss, { gsThreads: [], llmResponse: llm() });
  assert.strictEqual(ss.getSheetByName('확인 큐 (v1)').get(1, 3), '[패치] 덮어쓰기 충돌', '예전 행 보존');
  assert.deepStrictEqual(ss.getSheetByName('확인 큐').data[0], QUEUE_COLS);
});

test('[B2] 확인 큐 머리글이 일부만 다르면 실행을 멈춤', () => {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  ss.getSheetByName('확인 큐').getRange(1, 1, 1, 12).setValues([QUEUE_COLS.map((h, i) => (i === 5 ? '컬럼' : h))]);
  const f = progressRun(ss, { gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm() });
  assert.strictEqual(f.fetches.length, 0);
  assert.ok(f.logs.some(l => l.includes("'확인 큐' 탭의 머리글이 다릅니다")));
});

// 충돌 1건이 든 확인 큐를 만든 사본
function queueWithConflict() {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow({ 패치: '2026-11-18' })] }));
  progressRun(ss, { gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm({ 열갱신: [{ 열: '패치', 값: '2026-11-20' }] }) });
  return ss;
}
const qcol = name => QUEUE_COLS.indexOf(name) + 1;

test('[B2] 처리=적용 → 다음 실행 시작 때 마스터에 반영하고 처리시각 기록', () => {
  const ss = queueWithConflict();
  const queue = ss.getSheetByName('확인 큐'), master = ss.getSheetByName('마스터');
  queue.set(2, qcol('처리'), '적용');
  progressRun(ss, { gsThreads: [], llmResponse: llm() });
  assert.strictEqual(master.get(3, F.col('패치')), '2026-11-20');
  assert.ok(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(queue.get(2, qcol('처리시각'))));
});

test('[B2] 사람이 제안값을 고쳐서 적용해도 검사 후 반영 (2026.11.21 → 2026-11-21)', () => {
  const ss = queueWithConflict();
  const queue = ss.getSheetByName('확인 큐');
  queue.set(2, qcol('제안값'), '2026.11.21');
  queue.set(2, qcol('처리'), '적용');
  progressRun(ss, { gsThreads: [], llmResponse: llm() });
  assert.strictEqual(ss.getSheetByName('마스터').get(3, F.col('패치')), '2026-11-21');
});

test('[B2] 그사이 마스터 값이 바뀌었으면 적용하지 않고 대기로 되돌림', () => {
  const ss = queueWithConflict();
  const queue = ss.getSheetByName('확인 큐'), master = ss.getSheetByName('마스터');
  master.set(3, F.col('패치'), '2026-11-25');
  queue.set(2, qcol('처리'), '적용');
  progressRun(ss, { gsThreads: [], llmResponse: llm() });
  assert.strictEqual(master.get(3, F.col('패치')), '2026-11-25');
  assert.strictEqual(queue.get(2, qcol('처리')), '대기');
  assert.strictEqual(queue.get(2, qcol('처리시각')), '');
  assert.ok(queue.get(2, qcol('사유')).includes('적용 안 함') && queue.get(2, qcol('사유')).includes('마스터 값이 바뀜'));
});

test('[B2] 처리=무시 → 처리시각만 기록 / 시험번호 행은 자동 적용하지 않음', () => {
  const ss = queueWithConflict();
  const queue = ss.getSheetByName('확인 큐');
  queue.set(2, qcol('처리'), '무시');
  queue.appendRow(['Q-x-1', '', '', 'x', '', '시험번호', '', 'TTA-26-00001', '시험번호 없음', '{}', '적용', '']);
  progressRun(ss, { gsThreads: [], llmResponse: llm() });
  assert.strictEqual(ss.getSheetByName('마스터').get(3, F.col('패치')), '2026-11-18');
  assert.notStrictEqual(queue.get(2, qcol('처리시각')), '');
  assert.strictEqual(queue.get(3, qcol('처리')), '대기');
  assert.ok(queue.get(3, qcol('사유')).includes('직접 처리하세요'));
});

test('[B2] onQueueEdit: 확인 큐를 고치면 바로 반영, 다른 탭 편집은 무시 / 트리거 설치는 한 번만', () => {
  const ss = queueWithConflict();
  ss.getSheetByName('확인 큐').set(2, qcol('처리'), '적용');
  const flow = loadFlow({ props: { SPREADSHEET_ID: F.COPY_ID }, spreadsheets: { [F.COPY_ID]: ss } });
  flow.onQueueEdit({ range: { getSheet: () => ss.getSheetByName('마스터') } });
  assert.strictEqual(ss.getSheetByName('마스터').get(3, F.col('패치')), '2026-11-18');
  flow.onQueueEdit({ range: { getSheet: () => ss.getSheetByName('확인 큐') } });
  assert.strictEqual(ss.getSheetByName('마스터').get(3, F.col('패치')), '2026-11-20');
  flow.installQueueEditTrigger();
  flow.installQueueEditTrigger();
  assert.strictEqual(flow.triggers.length, 1);
  assert.strictEqual(flow.triggers[0].id, F.COPY_ID);
});

test('[B2] DRY_RUN 은 확인 큐 결정을 반영하지 않고 기록만', () => {
  const ss = queueWithConflict();
  const queue = ss.getSheetByName('확인 큐');
  queue.set(2, qcol('처리'), '적용');
  const before = [ss.getSheetByName('마스터').snapshot(), queue.snapshot()];
  progressRun(ss, { dry: true, gsThreads: [], llmResponse: llm() });
  assert.deepStrictEqual([ss.getSheetByName('마스터').snapshot(), queue.snapshot()], before);
  assert.ok(simObjs(ss).some(r => r.탭 === '마스터' && r.열 === '패치' && r.새값 === '2026-11-20' && r.사유.includes('확인 큐 적용')));
});

test('[B3] 다른 실행이 잠금을 잡고 있으면 아무것도 하지 않음', () => {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  const flow = loadFlow({ props: { GEMINI_API_KEY: 'k', SPREADSHEET_ID: F.COPY_ID }, spreadsheets: { [F.COPY_ID]: ss },
                          gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm(), lockBusy: true });
  flow.runProgressCheck();
  flow.runInitiationCheck();
  assert.strictEqual(flow.fetches.length, 0);
  assert.strictEqual(logRows(ss.getSheetByName('처리로그')).length, 0);
  assert.ok(flow.logs.some(l => l.includes('다른 실행이 진행 중')));
});

test('[B3] 시간이 다 되면 남은 메일은 처리로그 없이 다음 실행으로', () => {
  const ss = F.makeSpreadsheet(F.COPY_ID, F.makeMaster({ rows: [progressRow()] }));
  const flow = loadFlow({ props: { GEMINI_API_KEY: 'k', SPREADSHEET_ID: F.COPY_ID }, spreadsheets: { [F.COPY_ID]: ss },
                          gsThreads: [fakeThread('th1', [F.MAIL.progress1])], initThreads: [fakeThread('t1', [F.MAIL.init1])], llmResponse: llm() });
  flow.CONFIG.TIME_BUDGET_MS = -1;
  flow.runProgressCheck();
  flow.runInitiationCheck();
  assert.strictEqual(flow.fetches.length, 0);
  assert.strictEqual(logRows(ss.getSheetByName('처리로그')).length, 0);
  assert.strictEqual(ss.getSheetByName('마스터').get(4, F.col('상담자')), '');
});

test('[B4] API 키는 URL 이 아니라 x-goog-api-key 헤더로', () => {
  const { flow } = setup({ master: F.makeMaster({ rows: [progressRow()] }), gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm() });
  flow.runProgressCheck();
  const f = flow.fetches[0];
  assert.ok(!f.url.includes('key='), f.url);
  assert.strictEqual(f.options.headers['x-goog-api-key'], 'test-key');
});

// =========================================================================
// 4단계: SHEET_CONTRACT.md 의 탭 표가 코드의 머리글·허용값과 맞는지
// =========================================================================
const CONTRACT = require('fs').readFileSync(require('path').join(__dirname, '..', 'SHEET_CONTRACT.md'), 'utf8');
// '## N. `탭`' 절 안의 '| A | `이름` |' 행 → 이름 목록
function contractColumns(tab) {
  const sections = CONTRACT.split(/\n(?=## )/);
  const sec = sections.find(s => s.startsWith(`## `) && s.split('\n')[0].includes('`' + tab + '`'));
  assert.ok(sec, `SHEET_CONTRACT.md 에 '${tab}' 절이 없음`);
  return sec.split('\n').map(l => l.match(/^\| ([A-Z]{1,2}) \| `([^`]+)` \|/)).filter(Boolean).map(m => m[2]);
}

test('[계약] 확인 큐·할 일·처리로그·시뮬레이션 열 목록이 코드와 같음', () => {
  const { flow } = setup({});
  assert.strictEqual(JSON.stringify(contractColumns('확인 큐')), JSON.stringify(flow.QUEUE_HEADERS));
  assert.strictEqual(JSON.stringify(contractColumns('할 일')), JSON.stringify(flow.TODO_HEADERS));
  assert.strictEqual(JSON.stringify(contractColumns('처리로그')), JSON.stringify(flow.LOG_HEADERS));
  assert.strictEqual(JSON.stringify(contractColumns('시뮬레이션')), JSON.stringify(flow.SIM_HEADERS));
});

test('[계약] 허용값·LLM 허용 열이 문서에 그대로 적혀 있음', () => {
  const { flow } = setup({});
  const inline = list => list.map(v => '`' + v + '`').join('·');
  assert.ok(CONTRACT.includes(inline(flow.QUEUE_DECISIONS)), '확인 큐 처리 값');
  assert.ok(CONTRACT.includes(inline(flow.TODO_STATUSES)), '할 일 상태 값');
  assert.ok(CONTRACT.includes(inline(flow.VALID_STATUSES)), '마스터 상태 값');
  assert.ok(CONTRACT.includes(inline(Object.values(flow.LOG_STATUS))), '처리로그 상태 값');
  assert.ok(CONTRACT.includes(inline(flow.ALLOWED_COLUMNS)), 'LLM 허용 열');
  [flow.CONFIG.MASTER_SHEET_NAME, flow.CONFIG.LOG_SHEET_NAME, flow.CONFIG.QUEUE_SHEET_NAME,
   flow.CONFIG.TODO_SHEET_NAME, flow.CONFIG.SIM_SHEET_NAME].forEach(n => assert.ok(CONTRACT.includes('| `' + n + '` |'), `탭 목록에 ${n}`));
});

let pass = 0, fail = 0, todoFail = 0, fixed = 0;
for (const t of tests) {
  let err = null;
  try { t.fn(); } catch (e) { err = e; }
  const label = t.todo ? `[${t.todo}] ${t.name}` : t.name;
  if (!t.todo && !err) { pass++; console.log(`PASS  ${label}`); }
  else if (!t.todo && err) { fail++; console.log(`FAIL  ${label}\n      ${String(err.message).split('\n').join('\n      ')}`); }
  else if (t.todo && err) { todoFail++; console.log(`TODO  ${label}\n      ↳ ${String(err.message).split('\n')[0]}`); }
  else { fixed++; console.log(`FIXED ${label}  ← todo 표시를 지우세요`); }
}
console.log(`\n합계: PASS ${pass} · FAIL ${fail} · TODO(예상된 실패) ${todoFail} · FIXED ${fixed}`);
process.exit(fail > 0 ? 1 : 0);
