const CONFIG = {
  SPREADSHEET_ID: '1LSISjIp3SVuU_bt7GGJrIylnVDn9J-TVhM5Fi4P38Ng', 
  MASTER_SHEET_NAME: '마스터',
  LOG_SHEET_NAME: '처리로그',
  QUEUE_SHEET_NAME: '확인 큐',
  HEADER_ROW: 1,      
  DATA_START_ROW: 3,  
  LABEL_NAME: 'GS',
  TEST_NUM_REGEX: /[a-zA-Z]+-\d{2}-\d+/,
  
  // 새로 추가: 착수 안내 메일을 찾기 위한 검색어 (최근 3일치만 검색)
  INIT_SEARCH_QUERY: 'subject:"GS 시험 인증 시작 안내" newer_than:3d',

  GEMINI_MODEL: 'gemini-3.8-flash',
  MY_EMAIL: 'hojun7766@tta.or.kr' 
};
