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

function maskSensitiveData(text) {
  let masked = text.replace(/\b\d{3}-\d{3,4}-\d{4}\b/g, '***-****-****');
  masked = masked.replace(/\b(?:\d{1,3}\.){3}\d{1,3}\b/g, '***.***.***.***');
  return masked;
}

// writer 를 주면 DRY_RUN 중 '쓴 셈 친' 값까지 반영해서 읽습니다.
function getRowDataAsObject(sheet, rowIndex, headers, writer) {
  const lastCol = headers.length;
  const values = writer ? writer.getRowValues(sheet, rowIndex, lastCol)
                        : sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  const rowObj = {};
  headers.forEach((header, i) => {
    if (header) rowObj[header.toString().trim()] = values[i] === "" ? null : values[i];
  });
  return rowObj;
}

// 5번 패치: LLM 허용 열 제한 및 상태값 강제 (수식 셀 덮어쓰기 방지)
const ALLOWED_COLUMNS = ['담당자', '진행상황', '시험시작일', '시험종료일', '상태'];
const VALID_STATUSES = ['접수', '진행', '중단', '종료']; 

function filterLlmOutput(parsedData) {
  const filtered = {};
  for (let key in parsedData) {
    // 1. 화이트리스트에 없는 열은 무시
    if (!ALLOWED_COLUMNS.includes(key)) continue;
    
    // 2. 상태 열일 경우 정해진 값이 아니면 무시
    if (key === '상태' && !VALID_STATUSES.includes(parsedData[key])) {
      continue;
    }
    
    filtered[key] = parsedData[key];
  }
  return filtered;
}
