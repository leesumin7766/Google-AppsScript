// 테스트용 고정 데이터.
// ⚠️ 메일 본문은 코드의 정규식에 맞춰 만든 **가상 예시**입니다. 실제 메일에서 뽑은 익명화 본문으로
//    바꾸거나 추가해 주세요 (이름·회사·주소·번호는 가짜로 바꿔서).
// ⚠️ 마스터 헤더도 가정입니다 (E=상태, F=잔여일자 만 README 와 맞춤). 실제 1행 헤더로 바꾸면 더 정확합니다.
const { FakeSheet, FakeSpreadsheet } = require('./harness');

const LIVE_ID = '1LSISjIp3SVuU_bt7GGJrIylnVDn9J-TVhM5Fi4P38Ng'; // Config.gs 의 실제 시트 ID
const COPY_ID = 'COPY-SHEET-ID';

const MASTER_HEADERS = [
  '년도', '시험번호', '기업명', '상담자', '상태', '잔여일자', '시험안내일', '시험시작일', 'WD', '시험원',
  '담당자 (고객사 연락처/직함)', '1차결함', '2차결함', '3차결함', '패치', '시험종료일', '진행상황',
];
const col = name => MASTER_HEADERS.findIndex(h => h.split('(')[0].trim() === name) + 1;

// KST 자정 Date (시트에서 날짜 칸을 읽으면 이런 Date 가 나옵니다)
const kstDate = ymd => new Date(`${ymd}T00:00:00+09:00`);

// opts.formulaTo: F열(잔여일자)에 수식 결과가 이 행까지 차 있다고 가정 (실제 시트는 978행)
// opts.rows: 3행부터 넣을 행 객체 목록 { 헤더키: 값 }
function makeMaster(opts) {
  opts = opts || {};
  const rows = [MASTER_HEADERS.slice()];
  const example = MASTER_HEADERS.map(() => '');
  example[col('시험번호') - 1] = 'TTA-00-00000';
  example[col('기업명') - 1] = '(예시)';
  rows.push(example); // 2행 = 예시 데이터
  (opts.rows || []).forEach(obj => {
    const r = MASTER_HEADERS.map(() => '');
    Object.keys(obj).forEach(k => { r[col(k) - 1] = obj[k]; });
    rows.push(r);
  });
  const sheet = new FakeSheet('마스터', rows);
  // 수식 칸: 값은 "" 로 보이지만 getLastRow() 에는 잡힙니다 (실제 시트와 같음)
  if (opts.formulaTo) for (let r = 2; r <= opts.formulaTo; r++) sheet.markFormula(r, col('잔여일자'));
  return sheet;
}

function makeSpreadsheet(id, master, extra) {
  const sheets = [
    master || makeMaster(),
    new FakeSheet('처리로그', [['메시지ID', '처리 시각', '결과']]),
    new FakeSheet('확인 큐', []),
  ].concat(extra || []);
  return new FakeSpreadsheet(id, sheets);
}

const MAIL = {
  init1: {
    id: 'init-1', subject: 'GS 시험 인증 시작 안내', from: '김상담 책임 <counsel@tta.or.kr>',
    date: '2026-11-05T09:00:00+09:00',
    body: '안녕하세요. GS 인증 시험 시작을 안내드립니다.\n\n시작일: 2026년 11월 10일 (월), 기간 20 WD\n\n시험원: 이시험 선임, 박시험 책임\n',
  },
  init2: {
    id: 'init-2', subject: 'GS 시험 인증 시작 안내', from: '최상담 수석 <counsel2@tta.or.kr>',
    date: '2026-11-05T10:00:00+09:00',
    body: '시작일: 2026년 11월 17일 (월), 기간 15 WD\n시험원: 정시험 전임\n',
  },
  progress1: {
    id: 'prog-1', subject: '[GS] TTA-26-00001 가나다소프트, 1차 결함리포트 회신', from: '홍길동 <hong@ganada.example>',
    date: '2026-11-06T14:00:00+09:00',
    body: '안녕하세요. 가나다소프트 홍길동입니다.\n1차 결함 조치 후 패치는 2026년 11월 20일에 전달드리겠습니다.\n\n홍길동 선임연구원\n연락처: 010.1234.5678\n',
  },
  progress2: {
    id: 'prog-2', subject: 'Re: [GS] TTA-26-00001 가나다소프트, 일정 문의', from: '홍길동 <hong@ganada.example>',
    date: '2026-11-06T15:00:00+09:00',
    body: '종료 일정은 추후 공유드리겠습니다.\n',
  },
  noRow: {
    id: 'prog-9', subject: '[GS] TTA-26-09999 없는회사, 문의', from: '누군가 <a@b.example>',
    date: '2026-11-06T16:00:00+09:00', body: '문의드립니다.\n',
  },
};

// 인용문 절삭 케이스 (C4). expected = 1단계 이후 기대값
const QUOTE_CASES = [
  { name: 'Gmail 한국어 "…님이 작성:"', body: '확인했습니다.\n\n2026년 11월 5일 (목) 오후 3:00, 홍길동 <a@b.example>님이 작성:\n> 이전 내용 1차결함 2026-11-01', expected: '확인했습니다.' },
  { name: 'Outlook "-----원본 메시지-----"', body: '패치 보냈습니다.\n\n-----원본 메시지-----\n보내는 사람: x\n시험종료일 2026-12-01', expected: '패치 보냈습니다.' },
  { name: '영문 "From:" 헤더', body: 'Patch sent.\n\nFrom: Kim <k@x.example>\nSent: Monday\nold text', expected: 'Patch sent.' },
  { name: '"Forwarded message" 아래는 자르지 않음', body: '참고 바랍니다.\n\n---------- Forwarded message ---------\nFrom: 업체 <c@x.example>\n시작일: 2026년 11월 10일', expected: '참고 바랍니다.\n\n---------- Forwarded message ---------\nFrom: 업체 <c@x.example>\n시작일: 2026년 11월 10일' },
  { name: '본문 중간의 ">" 는 인용 아님', body: '응답시간 기준은 > 3초 입니다. 패치는 11/20 예정입니다.', expected: '응답시간 기준은 > 3초 입니다. 패치는 11/20 예정입니다.' },
  { name: '본문 첫 줄 "On 11/10 …" 은 인용 아님', body: 'On 11/10 we will deliver the patch.\nThanks\n\nOn Tue, Nov 4, 2026 Kim wrote:\n> old', expected: 'On 11/10 we will deliver the patch.\nThanks' },
];

module.exports = { LIVE_ID, COPY_ID, MASTER_HEADERS, col, kstDate, makeMaster, makeSpreadsheet, MAIL, QUOTE_CASES };
