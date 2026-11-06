// 계약 데이터 내보내기: 실제 .gs 코드의 상수를 읽어 Flow/tests/contract.json 을 만듭니다.
// 실행: node Flow/tests/export-contract.js
// ATDT 테스트가 이 파일을 읽어 같은 머리글·허용값을 씁니다. 코드를 고쳤으면 이 명령을 다시 실행하세요.
// 테스트 `[계약] contract.json 이 코드와 같음` 이 같은 함수로 만든 값과 파일을 비교합니다.
const fs = require('fs');
const path = require('path');
const { loadFlow } = require('./harness');
const { MASTER_HEADERS } = require('./fixtures');

const OUT_PATH = path.join(__dirname, 'contract.json');

// 코드 상수에서 꺼낸 값을 키 순서를 고정한 객체로 만듭니다. (순수 함수: 같은 코드면 같은 결과)
// 코드에 상수가 없는 값(스키마 버전, ID 규칙, 날짜 형식 문자열 등)은 여기에 직접 적습니다.
function buildContract() {
  const flow = loadFlow();
  const toArr = a => Array.from(a);
  const logStatuses = Object.values(flow.LOG_STATUS).map(String);

  return {
    schema: 1, // SHEET_CONTRACT.md 의 스키마 버전 (코드에 없음, 직접 적음)
    tabs: {
      master: flow.CONFIG.MASTER_SHEET_NAME,
      log: flow.CONFIG.LOG_SHEET_NAME,
      queue: flow.CONFIG.QUEUE_SHEET_NAME,
      todo: flow.CONFIG.TODO_SHEET_NAME,
      sim: flow.CONFIG.SIM_SHEET_NAME,
      settings: '설명', // 코드에 없음 (직접 적음): 스키마 버전 셀이 있는 탭
      queueArchive: '확인 큐 (v1)', // 코드에 없음 (직접 적음): 예전 형식 보관 이름 `${name} (v1)` 의 기본형
    },
    headers: {
      queue: toArr(flow.QUEUE_HEADERS),
      todo: toArr(flow.TODO_HEADERS),
      log: toArr(flow.LOG_HEADERS),
      sim: toArr(flow.SIM_HEADERS),
    },
    allowedValues: {
      queueDecision: toArr(flow.QUEUE_DECISIONS),   // 확인 큐 `처리`
      todoStatus: toArr(flow.TODO_STATUSES),        // 할 일 `상태`
      logStatus: logStatuses,                       // 처리로그 `상태`
      masterStatus: toArr(flow.VALID_STATUSES),     // 마스터 `상태`
    },
    llmAllowedColumns: toArr(flow.ALLOWED_COLUMNS), // LLM 이 바꿀 수 있는 마스터 열 (코드 그대로)
    ids: {
      queue: 'Q-<메시지ID>-<n>', // 코드에 없는 문자열 템플릿 규칙 (Tabs.gs buildQueueRows): n 은 1부터
      todo: 'T-<메시지ID>-<n>',  // 코드에 없는 문자열 템플릿 규칙 (Tabs.gs buildTodoRows): n 은 1부터
    },
    formats: {
      timeZone: 'Asia/Seoul',                 // 코드에서는 formatDate 호출 인자로만 있음 (직접 적음)
      date: 'yyyy-MM-dd',
      dateTime: 'yyyy-MM-dd HH:mm:ss',
    },
    masterHeaders: {
      // 실제 시트 1행에서 받은 값 (2026-11-06). tests/fixtures.js 의 MASTER_HEADERS 와 같음
      source: '실제 시트 마스터 1행 (2026-11-06 받음). tests/fixtures.js MASTER_HEADERS',
      values: MASTER_HEADERS.slice(),
    },
  };
}

// 파일에 쓰는 정확한 문자열 (테스트도 이 함수로 비교)
function serializeContract(obj) {
  return JSON.stringify(obj, null, 2) + '\n';
}

module.exports = { buildContract, serializeContract, OUT_PATH };

if (require.main === module) {
  const text = serializeContract(buildContract());
  fs.writeFileSync(OUT_PATH, text, 'utf8');
  console.log(`계약 데이터를 썼습니다: ${OUT_PATH} (${text.length} 글자)`);
}
