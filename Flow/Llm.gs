function callGeminiAPI(apiKey, body, isMyEmail, date, subject, testNumber, currentRowData) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;
  const KSTDate = Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

  const systemPrompt = `너는 GS 인증 시험원의 업무를 돕는 데이터 추출 AI다.
아래 메일 정보를 분석하여 반드시 주어진 JSON 형식 한 개만 출력해라.
- 기준일시(메일 발신일): ${KSTDate}
- 발신자 구분: ${isMyEmail ? "시험원(나)" : "업체"}
- 마스터 시트의 현재 행 상태(비어있는 열 파악용): ${JSON.stringify(currentRowData)}

상대 날짜(예: 내일)는 기준일시를 바탕으로 절대 날짜(YYYY-MM-DD)로 변환해라.
메일에 없는 내용을 추측하지 마라.
추출할 대상이 없거나 모호하면 확신도를 "낮음"으로 설정해라.

출력 JSON 형식:
{
  "시험번호": "${testNumber}",
  "메일유형": "이벤트 중 하나 또는 기타",
  "요약": "메일 내용 한 줄 요약",
  "열갱신": [{"열": "헤더 이름", "값": "들어갈 값"}],
  "상태변경": "상태값 또는 null",
  "할일": [{"내용": "...", "기한": "YYYY-MM-DD 또는 null"}],
  "확신도": "높음|보통|낮음",
  "불확실한이유": "이유 또는 null"
}`;

  const payload = {
    "systemInstruction": { "parts": [{ "text": systemPrompt }] },
    "contents": [{ "role": "user", "parts": [{ "text": `제목: ${subject}\n\n본문: ${body}` }] }],
    "generationConfig": {
      "responseMimeType": "application/json", // JSON만 출력하도록 강제
      "temperature": 0.1
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (json.error) {
      Logger.log("API 오류: " + json.error.message);
      return null;
    }
    const resultText = json.candidates[0].content.parts[0].text;
    return JSON.parse(resultText);
  } catch (e) {
    Logger.log("LLM 파싱 실패: " + e.message);
    return null;
  }
}

// LLM 에 보내기 전 가림: IP → 전화번호(+82 / 하이픈·점·공백·붙여쓰기 / 지역번호)
function maskSensitiveData(text) {
  if (!text) return text;
  let masked = text.replace(/\b(?:\d{1,3}\.){3}\d{1,3}\b/g, '***.***.***.***');
  masked = masked.replace(/\+82[-.\s]?\(?0?\d{1,2}\)?[-.\s]?\d{3,4}[-.\s]?\d{4}\b/g, '***-****-****');
  masked = maskPhoneNumber(masked);
  masked = masked.replace(/\b\d{3}-\d{3,4}-\d{4}\b/g, '***-****-****');
  return masked;
}

// 행 → { 정규화한 헤더: 값 }. 날짜는 yyyy-MM-dd(Asia/Seoul) 문자열로 바꿉니다.
// writer 를 주면 DRY_RUN 중 '쓴 셈 친' 값까지 반영해서 읽습니다.
function getRowDataAsObject(sheet, rowIndex, headers, writer) {
  const lastCol = headers.length;
  const values = writer ? writer.getRowValues(sheet, rowIndex, lastCol)
                        : sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  return rowValuesToObject(headers, values);
}

// 순수 함수 (tests/ 에서 확인)
function rowValuesToObject(headers, values) {
  const rowObj = {};
  headers.forEach((header, i) => {
    if (!header) return;
    const v = values[i];
    rowObj[normalizeHeader(header)] = (v === "" || v === undefined) ? null : toYmdIfDate(v);
  });
  return rowObj;
}

function toYmdIfDate(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  return v;
}

// 원칙 2: LLM 이 쓸 수 있는 열은 날짜 열과 상태뿐 (담당자·진행상황·시험번호·기업명·수식 열 제외)
const LLM_DATE_COLUMNS = ['시험시작일', '1차결함', '2차결함', '3차결함', '패치', '시험종료일'];
const ALLOWED_COLUMNS = LLM_DATE_COLUMNS.concat(['상태']);
const VALID_STATUSES = ['접수', '진행', '중단', '종료'];

// LLM 응답의 열갱신·상태변경을 코드 규칙으로 거릅니다. 시트를 건드리지 않는 순수 함수.
//   updates      : 바로 반영할 것 [{ colName, value, reason }]
//   queueReasons : 확인 큐로 보낼 사유 (하나라도 있으면 이 메일은 확인 큐 — 열별 반영은 2단계 A1)
//   ignored      : 허용되지 않은 열 등 버린 제안 [{ colName, value, reason }]
function checkLlmUpdates(llmResult, currentRowData, colMap) {
  const out = { updates: [], queueReasons: [], ignored: [] };
  const proposals = [];
  (Array.isArray(llmResult.열갱신) ? llmResult.열갱신 : []).forEach(u => {
    if (u && u.열) proposals.push({ colName: normalizeHeader(u.열), value: u.값, source: 'LLM: 열갱신' });
  });
  if (llmResult.상태변경) proposals.push({ colName: '상태', value: llmResult.상태변경, source: 'LLM: 상태변경' });

  for (const p of proposals) {
    const value = (p.value === null || p.value === undefined) ? '' : String(p.value).trim();
    if (!ALLOWED_COLUMNS.includes(p.colName)) { out.ignored.push({ colName: p.colName, value, reason: '허용되지 않은 열' }); continue; }
    if (!colMap[p.colName]) { out.ignored.push({ colName: p.colName, value, reason: '시트에 없는 열' }); continue; }
    if (value === '' || value === 'null') { out.ignored.push({ colName: p.colName, value, reason: '빈 값' }); continue; }

    const current = currentRowData[p.colName];
    const hasCurrent = current !== null && current !== undefined && String(current).trim() !== '';

    if (p.colName === '상태') {
      if (!VALID_STATUSES.includes(value)) { out.queueReasons.push(`[상태] 허용되지 않은 값: ${value}`); continue; }
      if (hasCurrent && String(current).trim() === value) continue; // 이미 같은 상태
      out.updates.push({ colName: '상태', value, reason: p.source });
      continue;
    }

    // 날짜 열
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) { out.queueReasons.push(`[${p.colName}] 날짜 형식(yyyy-MM-dd) 아님: ${value}`); continue; }
    if (hasCurrent) { out.queueReasons.push(`[${p.colName}] 덮어쓰기 충돌 (기존값: ${current})`); continue; }
    out.updates.push({ colName: p.colName, value, reason: p.source });
  }
  return out;
}
