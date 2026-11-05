// 실행: node Flow/tests/run.test.js
// PASS  : 지금 코드가 기대대로 동작
// FAIL  : 회귀 — 고쳐야 함 (종료 코드 1)
// TODO  : 다음 단계에서 고칠 문제가 지금 재현됨 (예상된 실패)
// FIXED : TODO 였는데 통과함 → todo 표시를 지우세요
const assert = require('assert');
const { fakeThread, loadFlow, lastLlmBody } = require('./harness');
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
  assert.strictEqual(real.queue.getLastRow(), 1, '실제 실행: 둘째 메일은 패치 충돌로 확인 큐');
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
// 1단계에서 고칠 것 (C1~C4) — 지금은 실패가 정상
// =========================================================================
todo('1단계 C1', '착수 메일 2건 → F열 수식이 978행까지 있어도 3행·4행에 들어감', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ formulaTo: 978 }),
    initThreads: [fakeThread('t1', [F.MAIL.init1]), fakeThread('t2', [F.MAIL.init2])],
  });
  flow.runInitiationCheck();
  assert.strictEqual(master.get(3, F.col('상담자')), '김상담', `3행이 비어 있음 (실제로는 ${master.getLastRow() - 1}행 근처에 씀)`);
  assert.strictEqual(master.get(4, F.col('상담자')), '최상담');
});

todo('1단계 C1', '시험번호가 빈 착수 행(3행) 다음 착수 메일은 4행에 들어가고 3행을 덮지 않음', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ formulaTo: 978, rows: [{ 상담자: '김상담', 시험안내일: '2026-11-04', 상태: '접수' }] }),
    initThreads: [fakeThread('t2', [F.MAIL.init2])],
  });
  flow.runInitiationCheck();
  assert.strictEqual(master.get(3, F.col('상담자')), '김상담');
  assert.strictEqual(master.get(4, F.col('상담자')), '최상담');
});

todo('1단계 C2', 'getRowDataAsObject 키도 normalizeHeader 를 거침 (LLM 프롬프트의 키 = colMap 키)', () => {
  const flow = loadFlow();
  const master = F.makeMaster({ rows: [progressRow({ 담당자: '홍길동' })] });
  const obj = flow.getRowDataAsObject(master, 3, master.data[0]);
  assert.strictEqual(obj['담당자'], '홍길동');
});

todo('1단계 C2·C3', '담당자가 채워진 행에 LLM 이 담당자를 제안해도 덮어쓰지 않음', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ rows: [progressRow({ 담당자: '홍길동\nhong@ganada.example' })] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 열갱신: [{ 열: '담당자', 값: '***-****-****' }] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('담당자')), '홍길동\nhong@ganada.example');
});

todo('1단계 C3', 'ALLOWED_COLUMNS 에 담당자·진행상황이 없음', () => {
  const flow = loadFlow();
  assert.ok(!flow.ALLOWED_COLUMNS.includes('담당자'));
  assert.ok(!flow.ALLOWED_COLUMNS.includes('진행상황'));
});

todo('1단계 C3', 'LLM 이 허용 목록 밖의 열(잔여일자 수식 열·진행상황)을 제안하면 쓰지 않음', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ rows: [progressRow()] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 열갱신: [{ 열: '잔여일자', 값: '0' }, { 열: '진행상황', 값: '덮어쓰기' }] }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('잔여일자')), '');
  assert.strictEqual(master.get(3, F.col('진행상황')), '[11/06] 패치 일정 안내');
});

todo('1단계 C3', '허용값이 아닌 상태변경(예: "완료")은 쓰지 않음', () => {
  const { flow, master } = setup({
    master: F.makeMaster({ rows: [progressRow()] }),
    gsThreads: [fakeThread('th1', [F.MAIL.progress1])],
    llmResponse: llm({ 상태변경: '완료' }),
  });
  flow.runProgressCheck();
  assert.strictEqual(master.get(3, F.col('상태')), '진행');
});

F.QUOTE_CASES.forEach(c => {
  todo('1단계 C4', `cleanEmailBody: ${c.name}`, () => {
    const flow = loadFlow();
    assert.strictEqual(flow.cleanEmailBody(c.body), c.expected);
  });
});

todo('1단계 C4', 'LLM 에 보내는 본문에서 인용문이 잘려 있음', () => {
  const quoted = Object.assign({}, F.MAIL.progress1, {
    body: '패치 일정 공유드립니다.\n\n-----Original Message-----\nFrom: 시험원\n이전 메일 1차결함 2026-10-01\n',
  });
  const { flow } = setup({ master: F.makeMaster({ rows: [progressRow()] }), gsThreads: [fakeThread('th1', [quoted])], llmResponse: llm() });
  flow.runProgressCheck();
  assert.ok(!lastLlmBody(flow).includes('이전 메일'));
});

todo('1단계 C4', '자르고 나서 본문이 거의 비면 LLM 을 부르지 않고 사유를 남김', () => {
  const empty = Object.assign({}, F.MAIL.progress1, { body: '\n\n-----Original Message-----\nFrom: 시험원\n이전 메일\n' });
  const { flow, log } = setup({ master: F.makeMaster({ rows: [progressRow()] }), gsThreads: [fakeThread('th1', [empty])], llmResponse: llm() });
  flow.runProgressCheck();
  assert.strictEqual(flow.fetches.length, 0);
  assert.ok(String(log.get(2, 3)).includes('본문'));
});

// README "v1.1 적용 완료" 에 적혀 있지만 실제로는 호출되지 않는 것
todo('1단계 (v1.1 미적용)', 'LLM 에 보내는 본문의 점·붙여쓴 전화번호도 가림 (maskPhoneNumber 미사용)', () => {
  const { flow } = setup({ master: F.makeMaster({ rows: [progressRow()] }), gsThreads: [fakeThread('th1', [F.MAIL.progress1])], llmResponse: llm() });
  flow.runProgressCheck();
  assert.ok(!lastLlmBody(flow).includes('010.1234.5678'));
});

todo('1단계 (v1.1 미적용)', 'getRowDataAsObject 가 Date 를 yyyy-MM-dd(Asia/Seoul) 로 바꿈', () => {
  const flow = loadFlow();
  const master = F.makeMaster({ rows: [progressRow({ 시험시작일: F.kstDate('2026-11-10') })] });
  const obj = flow.getRowDataAsObject(master, 3, master.data[0]);
  assert.strictEqual(obj['시험시작일'], '2026-11-10');
});

// =========================================================================
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
