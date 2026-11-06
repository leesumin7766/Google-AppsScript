// 테스트용 고정 데이터.
// ⚠️ 메일 본문은 코드의 정규식에 맞춰 만든 **가상 예시**입니다. 실제 메일에서 뽑은 익명화 본문으로
//    바꾸거나 추가해 주세요 (이름·회사·주소·번호는 가짜로 바꿔서).
// 마스터 헤더: 실제 시트 1행 (A~AO, 2026-11-06 받음).
const { FakeSheet, FakeSpreadsheet } = require('./harness');

const LIVE_ID = '1LSISjIp3SVuU_bt7GGJrIylnVDn9J-TVhM5Fi4P38Ng'; // Config.gs 의 실제 시트 ID
const COPY_ID = 'COPY-SHEET-ID';

const MASTER_HEADERS = [
  '구분', '년도', '번호', '시험번호', '상태', '잔여일자', '기업명', '제품명', '배정', '담당자 (고객사 연락처/직함)',
  '상담자', '시험원', 'WD', '신청일', '계약일', '시험안내일', '환경구성일', '시험시작일', '시험종료일', '시험중단일',
  '시험재개일', '기능 리스트 검토', '기능 리스트 회신', '설치일', '1차결함', '2차결함', '3차결함', '패치', '진행증 요청', '종료일 변경제출',
  '진행증 도착', '장비반출 안내', '장비반출', '결과 온/오프', '결과확인 발송일', '합격서 발송일 (의뢰)', '명세 발송일 (입금확인)', '저작권 확인서 제출일', 'ICM 등록일', '현장시험',
  '진행상황 (비고 및 메모)',
];
const col = name => MASTER_HEADERS.findIndex(h => h.split('(')[0].trim() === name) + 1;

// KST 자정 Date (시트에서 날짜 칸을 읽으면 이런 Date 가 나옵니다)
const kstDate = ymd => new Date(`${ymd}T00:00:00+09:00`);

// opts.formulaTo: F열(잔여일자)에 수식 결과가 이 행까지 차 있다고 가정 (실제 시트는 978행)
// opts.rows: 3행부터 넣을 행 객체 목록 { 헤더키: 값 }
function makeMaster(opts) {
  opts = opts || {};
  const rows = [MASTER_HEADERS.slice()];
  rows.push(exampleRow()); // 2행 = 예시 데이터 (실제 시트 2행과 같은 모양)
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

// 실제 시트 2행 (예시). 시험번호가 TTA-26-00001 이므로 테스트의 진행 건과 겹치지 않게 예시 번호만 바꿉니다.
function exampleRow() {
  const r = MASTER_HEADERS.map(() => '');
  const set = (k, v) => { r[col(k) - 1] = v; };
  set('구분', 'GS'); set('년도', 2026); set('번호', '예시'); set('시험번호', 'TTA-00-00000'); set('잔여일자', 55);
  set('기업명', '주식회사 홍길동'); set('제품명', '홍길동 솔루션 v1.0');
  set('담당자', '김고객 팀장(개발본부)\nkim@honggildong.com\n***-****-****');
  set('상담자', '이수민'); set('시험원', '홍길동'); set('WD', 50);
  set('시험안내일', kstDate('2026-10-30')); set('시험시작일', kstDate('2026-11-08')); set('시험종료일', kstDate('2026-12-31'));
  set('설치일', kstDate('2026-11-09')); set('1차결함', kstDate('2026-11-26')); set('2차결함', kstDate('2026-12-10')); set('패치', kstDate('2026-12-22'));
  return r;
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

// sampleMail/ 의 실제 메일 (D2). 본문은 getPlainBody() 처럼 CRLF 그대로입니다.
const REAL_MAIL = {
  schedule: {
    id: 'real-1', file: "sampleMail\\[TTA] TTA-26-00002 주식회사 치이카와, GS시험 일정 회신 및 결함 대응 계획.eml",
    subject: "RE: [TTA] TTA-26-00002 주식회사 치이카와, GS시험 일정 회신 및 결함 대응 계획",
    from: "­이수민 전임 상암AX센터 <hojun7766@tta.or.kr>",
    to: "­이수민 전임 상암AX센터 <hojun7766@tta.or.kr>",
    date: "2026-10-30T16:56:47+09:00",
    body: "안녕하세요. 주식회사 치이카와 개발팀 우사기 수석입니다.\r\n\r\n당사 '치이카와 보안 관제 솔루션 v2.0'의 GS인증 시험과 관련하여, 유선 협의드린 결함 리포트 발행 및 패치 일정을 다음과 같이\r\n확정하여 회신드립니다.\r\n\r\n1. 1차 결함리포트 발송일: 2026-11-27\r\n2. 2차 결함리포트 발송일: 2026-12-11\r\n3. 패치 버전 제출일: 2026-12-23\r\n4. 3차 결함리포트 발송일: 2026-12-30\r\n\r\n해당 일정에 맞추어 결함 분석 및 수정 패치 제출을 차질 없이 진행하겠습니다.\r\n감사합니다.\r\n\r\n---\r\n우사기 수석연구원 / 주식회사 치이카와\r\n연락처: 010-7777-3333\r\n이메일: usagi@chiikawa.com\r\n\r\n-------------------------------------------------------------------------\r\n-----------------------\r\n*한국정보통신기술협회(TTA)*\r\nAX·SW연구소/AX확산단/상암AX센터\r\n\r\n이수민 전임\r\n\r\n(03925) 서울시 마포구 월드컵북로 396 누리꿈스퀘어 비즈니스타워 404호\r\n\r\nMobile : 010-9540-0035                   Email : hojun7766@tta.or.kr\r\n-------------------------------------------------------------------------\r\n-----------------------\r\n",
  },
};

module.exports = { LIVE_ID, COPY_ID, MASTER_HEADERS, col, kstDate, makeMaster, makeSpreadsheet, MAIL, QUOTE_CASES, REAL_MAIL };
