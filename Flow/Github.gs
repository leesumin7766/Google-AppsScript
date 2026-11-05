const GITHUB = {
  OWNER: 'leesumin7766',
  REPO: 'Obsidian-Vault',
  BRANCH: 'main'
};

// 한글(UTF-8) 처리를 위한 Base64 변환 유틸리티
function utf8ToBase64(str) {
  return Utilities.base64Encode(Utilities.newBlob(str).getBytes());
}
function base64ToUtf8(base64Str) {
  return Utilities.newBlob(Utilities.base64Decode(base64Str)).getDataAsString();
}

function formatToMMDD(dateVal) {
  if (!dateVal || dateVal.toString().trim() === "") return "";
  
  try {
    const d = new Date(dateVal);
    if (!isNaN(d.getTime())) {
      // MM-dd 에서 MM/dd 로 변경!
      return Utilities.formatDate(d, "Asia/Seoul", "MM/dd");
    }
    return dateVal.toString(); 
  } catch (e) {
    return dateVal.toString();
  }
}

// 옵시디언 Flow.md 파일을 읽고 괄호 안의 날짜를 갱신하는 함수
function updateObsidianFlow(testNumber, rowData) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    Logger.log('오류: GITHUB_TOKEN이 설정되지 않았습니다.');
    return;
  }

  const getOptions = {
    method: 'get',
    headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/vnd.github.v3+json' },
    muteHttpExceptions: true
  };

  // 1. 'GS인증' 폴더 내부의 목록을 먼저 가져와서 시험번호가 포함된 폴더 찾기
  const baseDirPath = encodeURIComponent('GS인증');
  const dirUrl = `https://api.github.com/repos/${GITHUB.OWNER}/${GITHUB.REPO}/contents/${baseDirPath}?ref=${GITHUB.BRANCH}`;
  
  const dirResponse = UrlFetchApp.fetch(dirUrl, getOptions);
  if (dirResponse.getResponseCode() !== 200) {
    Logger.log(`GitHub 디렉토리 읽기 오류: ${dirResponse.getContentText()}`);
    return;
  }

  const dirContents = JSON.parse(dirResponse.getContentText());
  let targetFolderName = null;

  for (const item of dirContents) {
    if (item.type === 'dir' && item.name.includes(testNumber)) {
      targetFolderName = item.name;
      break;
    }
  }

  if (!targetFolderName) {
    Logger.log(`옵시디언 갱신 대기: '${testNumber}'가 포함된 폴더를 찾을 수 없습니다.`);
    return;
  }

  // 2. 찾은 폴더명을 바탕으로 실제 Flow.md 경로 생성
  const rawPath = `GS인증/${targetFolderName}/Flow.md`;
  const encodedPath = rawPath.split('/').map(encodeURIComponent).join('/');
  const fileUrl = `https://api.github.com/repos/${GITHUB.OWNER}/${GITHUB.REPO}/contents/${encodedPath}?ref=${GITHUB.BRANCH}`;

  // 3. 파일 읽기
  const fileResponse = UrlFetchApp.fetch(fileUrl, getOptions);
  if (fileResponse.getResponseCode() === 404) {
    Logger.log(`옵시디언 갱신 대기: 폴더는 있으나 Flow.md 파일이 없습니다 (${rawPath})`);
    return;
  }
  if (fileResponse.getResponseCode() !== 200) {
    Logger.log(`GitHub 파일 읽기 오류: ${fileResponse.getContentText()}`);
    return;
  }
  
  const fileData = JSON.parse(fileResponse.getContentText());
  const fileSha = fileData.sha;
  let content = base64ToUtf8(fileData.content);

  // 4. [변경] 템플릿의 항목과 시트의 열 데이터(MM-dd 변환) 매핑
  const mapping = {
    'GS 시험 인증 시작 안내 G메일 확인': formatToMMDD(rowData['시험안내일']),
    '1차 결함리포트 발송': formatToMMDD(rowData['1차결함']),
    '2차 결함리포트 발송': formatToMMDD(rowData['2차결함']),
    '3차 결함리포트 발송': formatToMMDD(rowData['3차결함']),
    '패치': formatToMMDD(rowData['패치'])
  };

  let isModified = false;

  // 5. 괄호() 안의 내용 찾아 바꾸기
  for (const [taskName, date] of Object.entries(mapping)) {
    if (date && date.toString().trim() !== "") {
      const regex = new RegExp(`(- \\[[ x]\\] ${taskName})\\([^)]*\\)`, 'g');
      
      const newContent = content.replace(regex, `$1(${date})`);
      if (content !== newContent) {
        content = newContent;
        isModified = true;
      }
    }
  }

  if (!isModified) {
    Logger.log(`옵시디언 갱신 패스: 추가할 새 날짜가 없습니다 (${rawPath})`);
    return;
  }

  // 6. GitHub에 커밋 (수정본 덮어쓰기)
  const putPayload = {
    message: `🤖 자동 업데이트: ${testNumber} Flow 날짜 갱신 (MM-dd)`,
    content: utf8ToBase64(content),
    sha: fileSha,
    branch: GITHUB.BRANCH
  };

  const putOptions = {
    method: 'put',
    headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/vnd.github.v3+json' },
    contentType: 'application/json',
    payload: JSON.stringify(putPayload),
    muteHttpExceptions: true
  };

  const putResponse = UrlFetchApp.fetch(fileUrl, putOptions);
  if (putResponse.getResponseCode() === 200 || putResponse.getResponseCode() === 201) {
    Logger.log(`✅ 옵시디언 갱신 완료: ${rawPath}`);
  } else {
    Logger.log(`옵시디언 갱신 실패: ${putResponse.getContentText()}`);
  }
}
