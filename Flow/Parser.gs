// 1. 기존 진행 건의 기본 정보(년도, 기업명, 담당자) 파싱
function parseBasicInfo(subject, sender, body) {
  const year = new Date().getFullYear();
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

  // 연락처 추출 (마스킹 처리: ***_****_****)
  let phone = "";
  const phoneMatch = body.match(/(010|02|0[3-9]{1,2})[-.\s]?\d{3,4}[-.\s]?\d{4}|\*{3}_\*{4}_\*{4}/);
  if (phoneMatch) {
    phone = "***_****_****";
  }

  // 직급 추출 (수석연구원, 책임연구원 등 복합 직급 우선 탐색)
  let title = "";
  const titleRegex = new RegExp(name + "\\s*(수석연구원|책임연구원|전임연구원|선임연구원|팀장|과장|대리|사원|책임|선임|수석|연구원|매니저|이사|대표|부장|차장)", "i");
  const titleMatch = body.match(titleRegex);
  if (titleMatch) {
    title = titleMatch[1];
  } else {
    // 발신자 이름과 붙어있지 않고 서명 등에 단독 기재된 경우 대비
    const altTitleMatch = body.match(/([가-힣]{2,4})\s+(수석연구원|책임연구원|전임연구원|선임연구원|팀장|과장|대리|사원|책임|선임|수석|연구원|매니저|이사|대표|부장|차장)/);
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

  // [수정] 4줄 줄바꿈 조합: 이름 -> 이메일 -> 전화번호 -> 직급
  const managerLines = [];
  if (name) managerLines.push(name);
  if (email) managerLines.push(email);
  if (phone) managerLines.push(phone);
  if (title) managerLines.push(title);

  const managerInfo = managerLines.join("\n");

  return { '년도': year, '기업명': company, '담당자': managerInfo };
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

// 1번 패치: 회신 메일의 이전 대화(꼬리표) 원문 절삭
function cleanEmailBody(body) {
  if (!body) return "";
  const quoteRegex = /(\r?\n)?(-----Original Message-----|보낸 사람:|On .* wrote:|>\s).*/s;
  return body.split(quoteRegex)[0].trim();
}
