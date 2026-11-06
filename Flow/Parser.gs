const MASKED_PHONE = '***-****-****';
// 직급 후보. 긴 것을 앞에 둡니다 (같은 위치에서는 앞쪽 후보부터 맞춰 봄)
const TITLE_WORDS = '수석연구원|책임연구원|전임연구원|선임연구원|팀장|과장|대리|사원|책임|선임|수석|연구원|매니저|이사|대표|부장|차장';

// 1. 기존 진행 건의 기본 정보(년도, 기업명, 담당자) 파싱
// testNumber 를 주면 그 번호로, 없으면 제목의 번호로 년도를 정합니다 (A6: TTA-26-00001 → 2026)
function parseBasicInfo(subject, sender, body, testNumber) {
  const year = yearFromTestNumber(testNumber || (String(subject || '').match(CONFIG.TEST_NUM_REGEX) || [''])[0]);
  let company = "";
  const companyMatch = subject.match(/[a-zA-Z]+-\d{2}-\d+\s+([^,\[\]]+)/);
  if (companyMatch) company = companyMatch[1].trim(); 

  // 발신자 이름 및 이메일 추출
  let name = sender;
  let email = "";
  const senderMatch = sender.match(/(.*?)\s*<(.+?)>/);
  if (senderMatch) {
    name = senderMatch[1].replace(/["']/g, "").trim();
    email = senderMatch[2].trim();
  }

  // 연락처: 번호는 시트에 남기지 않고 자리만 표시합니다 (D2: 마스터 2행과 같은 ***-****-****)
  let phone = "";
  const phoneMatch = body.match(/(010|02|0[3-9]{1,2})[-.\s]?\d{3,4}[-.\s]?\d{4}|\*{3}[-_]\*{4}[-_]\*{4}/);
  if (phoneMatch) {
    phone = MASKED_PHONE;
  }

  // 직급 추출 (수석연구원, 책임연구원 등 복합 직급 우선 탐색)
  let title = "";
  // A5: 이름에 특수문자가 있어도 정규식이 깨지지 않게 이스케이프
  // D2: 본문에 여러 번 나오면 가장 긴 직급 (예: 인사말 "우사기 수석입니다" + 서명 "우사기 수석연구원" → 수석연구원)
  const titleRegex = new RegExp(escapeRegExp(name) + "\\s*(" + TITLE_WORDS + ")", "gi");
  const titles = [];
  let tm;
  while ((tm = titleRegex.exec(body)) !== null) titles.push(tm[1]);
  if (titles.length) {
    title = titles.sort((a, b) => b.length - a.length)[0];
  } else {
    // 발신자 이름과 붙어있지 않고 서명 등에 단독 기재된 경우 대비
    const altTitleMatch = body.match(new RegExp("([가-힣]{2,4})\\s+(" + TITLE_WORDS + ")"));
    if (altTitleMatch) {
      if (!name || name.includes("@")) name = altTitleMatch[1];
      title = altTitleMatch[2];
    }
  }

  // 본문에 별도 기재된 이메일이 있다면 우선 보강
  const bodyEmailMatch = body.match(/이메일\s*:\s*([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/i);
  if (bodyEmailMatch && (email.includes("tta.or.kr") || !email)) {
    email = bodyEmailMatch[1];
  }

  // D2: 마스터 2행과 같은 3줄 — "이름 직급" / 메일 / 전화번호(가림)
  const managerLines = [];
  const nameLine = [name, title].filter(Boolean).join(' ');
  if (nameLine) managerLines.push(nameLine);
  if (email) managerLines.push(email);
  if (phone) managerLines.push(phone);

  const managerInfo = managerLines.join("\n");

  return { '년도': year, '기업명': company, '담당자': managerInfo };
}

// A6: 시험번호의 두 자리 년도 → 4자리 (시험번호가 없으면 빈칸)
function yearFromTestNumber(testNumber) {
  const m = String(testNumber || '').match(/-(\d{2})-/);
  return m ? 2000 + Number(m[1]) : '';
}

// 2. 신규 착수 메일 내용 파싱 (기존 로직 유지)
function parseInitiationEmail(sender, body, date) {
  // 상담자: 한글과 공백만 남긴 후 첫 번째 단어(이름)만 추출 (직급, 센터 등 제거)
  let rawCounselor = sender.split('<')[0].replace(/[^가-힣\s]/g, "").trim();
  let counselor = rawCounselor.split(/\s+/)[0]; 
  
  // 시험안내일 (메일 수신일)
  let infoDate = Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd");

  // 시험시작일 및 WD 추출
  let startDate = "";
  let wd = "";
  const startMatch = body.match(/시\s*작\s*일\s*:\s*(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일.*?(\d+)\s*WD/);
  
  if (startMatch) {
    startDate = `${startMatch[1]}-${startMatch[2].padStart(2, '0')}-${startMatch[3].padStart(2, '0')}`;
    wd = startMatch[4];
  } else {
    const altMatch = body.match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/);
    if (altMatch) startDate = `${altMatch[1]}-${altMatch[2].padStart(2, '0')}-${altMatch[3].padStart(2, '0')}`;
  }

  // 시험원 추출
  let testers = [];
  const testerRegex = /([가-힣]{2,4})\s*(책임|전임|선임|수석)/g;
  let match;
  while ((match = testerRegex.exec(body)) !== null) {
    if (!counselor.includes(match[1])) testers.push(match[1]); 
  }

  return {
    '상담자': counselor,
    '시험안내일': infoDate,
    '시험시작일': startDate,
    'WD': wd,
    '시험원': testers.join('\n'), 
    '상태': '접수'
  };
}

// 4번 패치: 전화번호 마스킹 (PII 보호)
function maskPhoneNumber(text) {
  if (!text) return text;
  // 010.1234.5678, 021234567, 010-1234-5678 등 다양한 형태 마스킹 처리
  return text.replace(/\b0\d{1,2}[-.\s)]?\d{3,4}[-.\s]?\d{4}\b/g, '***-****-****');
}

// 1번 패치 (1단계 C4 에서 다시 작성): 회신 메일의 이전 대화(인용문) 절삭
// 줄 단위로 보다가 인용 머리말이 나온 줄부터 끝까지 자릅니다.
//   - "-----Original Message-----" / "-----원본 메시지-----"
//   - Gmail: "… <a@b>님이 작성:" (두 줄로 꺾여도 인식), "On … wrote:"
//   - Outlook: "From:/보낸 사람:" 다음 4줄 안에 "Sent:/Date:/To:/보낸 날짜:/받는 사람:" 이 있는 머리말 묶음
//   - 줄 맨 앞의 ">" 인용
// "Forwarded message" / "전달된 메일" 표시 아래는 전달된 본문이므로 자르지 않습니다.
function cleanEmailBody(body) {
  if (!body) return "";
  const lines = String(body).replace(/\r\n?/g, '\n').split('\n');
  const FORWARD = /^-{3,}\s*(Forwarded message|전달된 메일|전달된 메시지)\s*-{3,}\s*$/i;
  const ORIGINAL = /^-{3,}\s*(Original Message|원본 메시지)\s*-{3,}\s*$/i;
  const HEADER_FROM = /^(From|보낸 사람|보낸사람|보내는 사람)\s*:/i;
  const HEADER_NEXT = /^(Sent|Date|To|Subject|보낸 날짜|받는 사람|날짜|제목)\s*:/i;

  let cut = lines.length;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    const prev = i > 0 ? lines[i - 1].trim() : '';
    if (FORWARD.test(line)) break;
    if (ORIGINAL.test(line)) { cut = i; break; }
    if (/님이 작성:$/.test(line)) { cut = i; break; }
    if (/^작성:$/.test(line) && /님이$/.test(prev)) { cut = i - 1; break; }
    if (/^On .+ wrote:$/.test(line)) { cut = i; break; }
    if (/wrote:$/.test(line) && /^On /.test(prev)) { cut = i - 1; break; }
    if (HEADER_FROM.test(line) && lines.slice(i + 1, i + 5).some(l => HEADER_NEXT.test(l.trim()))) { cut = i; break; }
    if (/^>/.test(line)) { cut = i; break; }
  }
  return lines.slice(0, cut).join('\n').trim();
}

// 인용문을 자른 본문이 거의 비었는지 (공백을 뺀 글자가 5자 미만)
function isBodyNearlyEmpty(text) {
  return String(text || '').replace(/\s/g, '').length < 5;
}
