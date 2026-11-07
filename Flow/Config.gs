const CONFIG = {
  SPREADSHEET_ID: '1LSISjIp3SVuU_bt7GGJrIylnVDn9J-TVhM5Fi4P38Ng',
  MASTER_SHEET_NAME: '마스터',
  LOG_SHEET_NAME: '처리로그',
  QUEUE_SHEET_NAME: '확인 큐',
  TODO_SHEET_NAME: '할 일',
  SIM_SHEET_NAME: '시뮬레이션', // DRY_RUN 일 때만 씁니다 (사본에서만 자동 생성)
  TIME_BUDGET_MS: 5 * 60 * 1000, // B3: 5분이 지나면 남은 메일은 다음 주기에
  HEADER_ROW: 1,
  DATA_START_ROW: 3,
  // 수식 열: 스크립트가 쓰지 않고, 새 행 위치를 정할 때도 보지 않습니다 (헤더 이름, 괄호 앞부분)
  FORMULA_COLUMNS: ['잔여일자'],
  LABEL_NAME: 'GS',
  TEST_NUM_REGEX: /[a-zA-Z]+-\d{2}-\d+/,

  // 새로 추가: 착수 안내 메일을 찾기 위한 검색어 (최근 3일치만 검색)
  INIT_SEARCH_QUERY: 'subject:"GS 시험 인증 시작 안내" newer_than:3d',
  GS_THREAD_LIMIT: 20,

  // A3: 발신 도메인으로 시험번호를 찾을 때 쓰지 않는 도메인 (여러 업체가 함께 쓰는 도메인)
  SHARED_MAIL_DOMAINS: ['tta.or.kr', 'gmail.com', 'naver.com', 'daum.net', 'hanmail.net', 'kakao.com',
                        'nate.com', 'outlook.com', 'hotmail.com'],

  // A2: LLM 실패 시 최대 시도 횟수
  LLM_MAX_ATTEMPTS: 3,
  // 2026-11-07 실환경: Gemini 가 503 을 자주 주고 한 번에 3분 넘게 걸리기도 함
  LLM_INNER_RETRY_WAIT_MS: 3000,   // 429·5xx 면 같은 실행 안에서 이만큼 기다렸다 한 번 더 부름
  LLM_MIN_REMAINING_MS: 120000,    // 실행 시간이 이보다 적게 남으면 LLM 을 부르지 않고 다음 실행으로 넘김

  // A7: LLM 이 고를 메일유형 후보
  MAIL_TYPES: ['착수', '1차결함', '2차결함', '3차결함', '패치', '시험종료', '일정변경', '문의', '기타'],

  // 2026-11-07 실환경 점검(checkGemini): 3.8/3.7-flash 는 503 과부하, 3.6-flash 31초, 3.5-flash-lite 0.7초
  GEMINI_MODEL: 'gemini-3.5-flash-lite',
  GEMINI_FALLBACK_MODEL: 'gemini-flash-lite-latest',   // 기본 모델이 429·5xx 면 같은 실행에서 이 모델로 한 번 더
  MY_EMAIL: 'hojun7766@tta.or.kr'
};

// 0단계: 스크립트 속성으로 덮어쓸 수 있는 설정
//   SPREADSHEET_ID    : 대상 시트 (테스트 때는 사본 ID)
//   DRY_RUN           : true 면 마스터·확인 큐·할 일·처리로그에 쓰지 않고 시뮬레이션 탭에만 기록
//   INIT_SEARCH_QUERY : 착수 메일 검색어 (지난 메일을 다시 돌릴 때 newer_than 을 늘림)
//   GS_THREAD_LIMIT   : GS 라벨에서 읽을 스레드 수
function getSetting(name, defaultValue) {
  const v = PropertiesService.getScriptProperties().getProperty(name);
  if (v === null || v === undefined || String(v).trim() === '') return defaultValue;
  return String(v).trim();
}

function getSpreadsheetId() {
  return getSetting('SPREADSHEET_ID', CONFIG.SPREADSHEET_ID);
}

function isDryRun() {
  return /^(true|1|on|yes)$/i.test(getSetting('DRY_RUN', ''));
}
