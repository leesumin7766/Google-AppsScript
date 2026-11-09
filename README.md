# Google-AppsScript

GS 인증 업무용 Google Apps Script 모음. 현재 쓰는 프로젝트는 **[Flow](Flow/README.md)** 하나다.

- **Flow**: GS 인증 메일(Gmail)을 읽어 시험일정 구글 시트(마스터·확인 큐·할 일·처리로그)를 갱신한다. 시트 형식 계약은 [Flow/SHEET_CONTRACT.md](Flow/SHEET_CONTRACT.md).
- 같은 시트를 업무 허브 **ATDT**(데스크톱 앱)가 읽고 확인 큐 결정을 써 넣는다. 시험 수행 웹 **autoTest** 와는 직접 연결하지 않는다.

## 회사에서 사용하기 (2026-11-09)

Flow 는 **회사 구글 계정의 Apps Script 프로젝트 안에서** 시간 트리거로 돈다(착수 메일 매일 07시, 진행 메일 매시간). PC 에 설치하는 것이 없으므로 집·회사 어디서든 따로 할 일이 없다.

남은 일은 **실제 시트로 전환** 하나다(지금은 사본 시트로 검증 중). Apps Script 편집기 → 프로젝트 설정 → 스크립트 속성에서:

1. `SPREADSHEET_ID` 삭제(없으면 원본 시험일정요약 사용), `GITHUB_TOKEN` 삭제(더 쓰지 않음), `INIT_SEARCH_QUERY` 는 지워도 됨
2. 트리거 화면에서 사본용 `onQueueEdit` 를 지우고, 편집기에서 `installQueueEditTrigger` 를 한 번 실행
3. 원본 시트 `설명` 탭 빈 셀에 `schema: 1`
4. ATDT 설정 → 시트 목록에서 사본을 빼고 원본을 고름
5. **전환 전에** 회사 메일함의 시험용 예시 메일(라벨 `GS`)을 지우거나 라벨을 뗀다(원본 처리로그에 없어 다시 처리됨)

`GEMINI_API_KEY` 는 그대로 둔다. 모델은 기본 `gemini-3.5-flash-lite`(실패 시 `gemini-flash-lite-latest`), `GEMINI_MODEL` 속성으로 바꿀 수 있고 편집기에서 `checkGemini` 로 진단한다.

## 회사 PC 에서 코드를 고칠 때

1. 저장소를 clone 하고 `dev-integration` 브랜치에서 작업한다.
2. `Flow/.clasp.json`(스크립트 ID)은 git 에 없다. 로컬에 만든다 — 형식은 [Flow/README.md](Flow/README.md) 'clasp 로 올리기'.
3. `npx @google/clasp@3.4.1 login` 을 **Apps Script 프로젝트 주인 계정(회사 계정)**으로 한 번 하고, `show-file-status` 로 올릴 파일을 확인한 뒤 `push` 한다.
4. 테스트: `node Flow/tests/run.test.js`
