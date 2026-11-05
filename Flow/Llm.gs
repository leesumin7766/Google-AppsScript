// Gemini 호출 (2단계 A2·A7)
// 반환: { ok: true, data } 또는 { ok: false, kind, message }
//   kind: 'transient' (429·5xx·네트워크 — 다시 시도) / 'parse' (응답 해석 실패 — 다시 시도)
//         'blocked' (안전 차단 — 포기) / 'fatal' (400·401·403·404 — 키·모델·스키마 문제. 실행을 멈춤)
// currentRowData 는 허용 열과 그 현재 값만 보냅니다 (담당자 등은 보내지 않음).
function callGeminiAPI(apiKey, body, isMyEmail, date, subject, testNumber, currentRowData) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;
  const payload = buildGeminiPayload(body, isMyEmail, date, subject, testNumber, currentRowData);
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    return { ok: false, kind: 'transient', message: '네트워크 오류: ' + e.message };
  }
  return classifyGeminiResponse(response.getResponseCode(), response.getContentText());
}

// 순수 함수: 프롬프트와 요청 본문 (A7)
function buildGeminiPayload(body, isMyEmail, date, subject, testNumber, currentRowData) {
  const KSTDate = Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  const allowedNow = {};
  ALLOWED_COLUMNS.forEach(c => {
    const v = currentRowData ? currentRowData[c] : undefined;
    allowedNow[c] = (v === undefined || v === '') ? null : v;
  });

  const systemPrompt = `너는 GS 인증 시험원의 업무를 돕는 데이터 추출 AI다.
아래 메일을 분석하여 주어진 JSON 형식 한 개만 출력해라.
- 기준일시(메일 발신일): ${KSTDate}
- 발신자 구분: ${isMyEmail ? "시험원(나)" : "업체"}
- 시험번호: ${testNumber}
- 갱신할 수 있는 열과 현재 값(null 은 빈칸): ${JSON.stringify(allowedNow)}

규칙:
- 메일유형은 다음 중 하나: ${CONFIG.MAIL_TYPES.join(', ')}
- 열갱신의 "열"은 다음 중 하나: ${ALLOWED_COLUMNS.filter(c => c !== '상태').join(', ')}. "값"은 YYYY-MM-DD.
- 상태변경은 다음 중 하나 또는 null: ${VALID_STATUSES.join(', ')}
- 상대 날짜(예: 내일)는 기준일시를 바탕으로 절대 날짜(YYYY-MM-DD)로 바꿔라.
- 메일에 없는 내용을 추측하지 마라. 메일에 날짜가 나온 열만 열갱신에 넣어라.
- 추출할 대상이 없거나 모호하면 확신도를 "낮음"으로 하고 불확실한이유를 적어라.`;

  return {
    "systemInstruction": { "parts": [{ "text": systemPrompt }] },
    "contents": [{ "role": "user", "parts": [{ "text": `제목: ${subject}\n\n본문: ${body}` }] }],
    "generationConfig": {
      "responseMimeType": "application/json",
      "responseSchema": geminiResponseSchema(),
      "temperature": 0.1
    }
  };
}

// A7: generationConfig.responseSchema 로 JSON 구조를 강제합니다
function geminiResponseSchema() {
  const str = (extra) => Object.assign({ type: 'STRING' }, extra || {});
  return {
    type: 'OBJECT',
    properties: {
      '시험번호': str(),
      '메일유형': str({ enum: CONFIG.MAIL_TYPES }),
      '요약': str(),
      '열갱신': {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: { '열': str({ enum: ALLOWED_COLUMNS.filter(c => c !== '상태') }), '값': str() },
          required: ['열', '값']
        }
      },
      '상태변경': str({ enum: VALID_STATUSES, nullable: true }),
      '할일': {
        type: 'ARRAY',
        items: { type: 'OBJECT', properties: { '내용': str(), '기한': str({ nullable: true }) }, required: ['내용'] }
      },
      '확신도': str({ enum: ['높음', '보통', '낮음'] }),
      '불확실한이유': str({ nullable: true })
    },
    required: ['메일유형', '요약', '열갱신', '확신도']
  };
}

// 순수 함수 (A2): HTTP 응답 → { ok, data } / { ok: false, kind, message }
function classifyGeminiResponse(code, text) {
  if (code === 429 || code >= 500) return { ok: false, kind: 'transient', message: `HTTP ${code}` };
  let json;
  try { json = JSON.parse(text); } catch (e) {
    return code >= 400 ? { ok: false, kind: 'fatal', message: `HTTP ${code}` }
                       : { ok: false, kind: 'parse', message: '응답이 JSON 이 아님' };
  }
  if (code >= 400 || json.error) {
    const msg = json.error && json.error.message ? json.error.message : '';
    return { ok: false, kind: 'fatal', message: `HTTP ${code} ${msg}`.trim() };
  }
  if (json.promptFeedback && json.promptFeedback.blockReason) {
    return { ok: false, kind: 'blocked', message: '프롬프트 차단: ' + json.promptFeedback.blockReason };
  }
  const cand = json.candidates && json.candidates[0];
  if (!cand) return { ok: false, kind: 'parse', message: '후보 응답 없음' };
  if (['SAFETY', 'BLOCKLIST', 'PROHIBITED_CONTENT', 'SPII', 'RECITATION'].includes(cand.finishReason)) {
    return { ok: false, kind: 'blocked', message: '응답 차단: ' + cand.finishReason };
  }
  try {
    const data = JSON.parse(cand.content.parts[0].text);
    if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('객체가 아님');
    return { ok: true, data: data };
  } catch (e) {
    return { ok: false, kind: 'parse', message: '응답 JSON 해석 실패: ' + e.message };
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
// 원칙 1 (A8): 날짜 순서. [앞, 뒤] 는 앞 ≤ 뒤 여야 합니다. (시작 ≤ 1차 ≤ 2차 ≤ 3차, 패치 ≤ 종료)
const DATE_ORDER_RULES = [
  ['시험시작일', '1차결함'], ['1차결함', '2차결함'], ['2차결함', '3차결함'],
  ['시험시작일', '2차결함'], ['시험시작일', '3차결함'], ['1차결함', '3차결함'],
  ['패치', '시험종료일'],
];

// LLM 응답의 열갱신·상태변경을 코드 규칙으로 거릅니다. 시트를 건드리지 않는 순수 함수.
//   updates      : 바로 반영할 것 [{ colName, value, reason }]
//   queued       : 확인 큐로 보낼 열 [{ colName, current, value, reason }] (A1: 그 열만 보냄)
//   queueReasons : queued 의 사유 문자열
//   ignored      : 허용되지 않은 열 등 버린 제안 [{ colName, value, reason }]
function checkLlmUpdates(llmResult, currentRowData, colMap) {
  const out = { updates: [], queued: [], queueReasons: [], ignored: [] };
  const queue = (colName, current, value, reason) => {
    out.queued.push({ colName, current: current === undefined ? null : current, value, reason });
    out.queueReasons.push(`[${colName}] ${reason}`);
  };
  const proposals = [];
  (Array.isArray(llmResult.열갱신) ? llmResult.열갱신 : []).forEach(u => {
    if (u && u.열) proposals.push({ colName: normalizeHeader(u.열), value: u.값, source: 'LLM: 열갱신' });
  });
  if (llmResult.상태변경) proposals.push({ colName: '상태', value: llmResult.상태변경, source: 'LLM: 상태변경' });

  const dateUpdates = [];
  for (const p of proposals) {
    const value = (p.value === null || p.value === undefined) ? '' : String(p.value).trim();
    if (!ALLOWED_COLUMNS.includes(p.colName)) { out.ignored.push({ colName: p.colName, value, reason: '허용되지 않은 열' }); continue; }
    if (!colMap[p.colName]) { out.ignored.push({ colName: p.colName, value, reason: '시트에 없는 열' }); continue; }
    if (value === '' || value === 'null') { out.ignored.push({ colName: p.colName, value, reason: '빈 값' }); continue; }

    const current = currentRowData[p.colName];
    const hasCurrent = current !== null && current !== undefined && String(current).trim() !== '';

    if (p.colName === '상태') {
      if (!VALID_STATUSES.includes(value)) { queue('상태', current, value, `허용되지 않은 값: ${value}`); continue; }
      if (hasCurrent && String(current).trim() === value) continue; // 이미 같은 상태
      out.updates.push({ colName: '상태', value, reason: p.source });
      continue;
    }

    // 날짜 열
    const ymd = normalizeYmd(value);
    if (!ymd) { queue(p.colName, current, value, `날짜 형식(yyyy-MM-dd) 아님: ${value}`); continue; }
    if (hasCurrent) {
      if (normalizeYmd(current) === ymd) continue;                 // A1: 같은 날짜는 충돌이 아님
      queue(p.colName, current, ymd, `덮어쓰기 충돌 (기존값: ${current})`);
      continue;
    }
    if (dateUpdates.some(d => d.colName === p.colName)) {          // 같은 열을 두 번 제안
      queue(p.colName, current, ymd, '같은 열에 날짜를 여러 개 제안');
      continue;
    }
    dateUpdates.push({ colName: p.colName, value: ymd, reason: p.source });
  }

  // A8: 반영 후 모습(현재 값 + 제안)으로 날짜 순서 검사. 어긋난 제안 열은 확인 큐로.
  const merged = {};
  LLM_DATE_COLUMNS.forEach(c => { merged[c] = normalizeYmd(currentRowData[c]); });
  dateUpdates.forEach(d => { merged[d.colName] = d.value; });
  for (const d of dateUpdates) {
    const broken = DATE_ORDER_RULES.find(([a, b]) => (a === d.colName || b === d.colName) &&
      merged[a] && merged[b] && merged[a] > merged[b]);
    if (broken) {
      const [a, b] = broken;
      queue(d.colName, currentRowData[d.colName], d.value, `날짜 순서 어긋남 (${a} ${merged[a]} > ${b} ${merged[b]})`);
    } else {
      out.updates.push(d);
    }
  }
  return out;
}
