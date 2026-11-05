// 0단계: 시트에 쓰는 일을 한곳으로 모읍니다.
// DRY_RUN 이 켜져 있으면 마스터·확인 큐·할 일·처리로그에 쓰지 않고,
// 쓰려던 내용을 `시뮬레이션` 탭에 한 줄씩 기록합니다.
const SIM_HEADERS = ['시각', '메시지ID', '시험번호', '탭', '행', '열', '기존값', '새 값', '사유'];

// DRY_RUN 인데 시뮬레이션 탭을 준비할 수 없으면 null 을 돌려줍니다 (호출 쪽은 실행을 멈춤).
function createWriter(ss) {
  const dryRun = isDryRun();
  let simSheet = null;

  if (dryRun) {
    simSheet = ss.getSheetByName(CONFIG.SIM_SHEET_NAME);
    if (!simSheet) {
      if (ss.getId() === CONFIG.SPREADSHEET_ID) {
        Logger.log(`오류: DRY_RUN 인데 대상이 실제 시트이고 '${CONFIG.SIM_SHEET_NAME}' 탭이 없습니다. ` +
                   `실제 시트에는 탭을 만들지 않습니다. 스크립트 속성 SPREADSHEET_ID 에 사본 ID 를 넣으세요.`);
        return null;
      }
      simSheet = ss.insertSheet(CONFIG.SIM_SHEET_NAME);
      simSheet.appendRow(SIM_HEADERS);
    }
    Logger.log(`🧪 DRY_RUN: '${ss.getName()}' 에 쓰지 않고 '${CONFIG.SIM_SHEET_NAME}' 탭에만 기록합니다.`);
  }

  // DRY_RUN 중 '쓴 셈 친' 값. 같은 실행 안의 다음 메일이 실제 실행과 같은 값을 읽도록 합니다.
  const overlay = {};
  // DRY_RUN 중 appendRow 가 들어갈 행 번호 (실제 appendRow 와 같은 규칙: getLastRow()+1)
  const nextRow = {};
  const key = (sheet, row, col) => `${sheet.getName()}|${row}|${col}`;

  function record(ctx, tab, row, col, oldValue, newValue, reason) {
    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    simSheet.appendRow([now, ctx.msgId, ctx.testNumber, tab, row, col, oldValue, newValue, reason].map(toSimCell));
  }

  function getCellValue(sheet, row, col) {
    const k = key(sheet, row, col);
    if (dryRun && Object.prototype.hasOwnProperty.call(overlay, k)) return overlay[k];
    return sheet.getRange(row, col).getValue();
  }

  function getRowValues(sheet, row, numCols) {
    const values = sheet.getRange(row, 1, 1, numCols).getValues()[0];
    if (dryRun) {
      for (let c = 1; c <= numCols; c++) {
        const k = key(sheet, row, c);
        if (Object.prototype.hasOwnProperty.call(overlay, k)) values[c - 1] = overlay[k];
      }
    }
    return values;
  }

  // ctx: { msgId, testNumber, colName, reason }
  function setCell(sheet, row, col, value, ctx) {
    if (!dryRun) {
      sheet.getRange(row, col).setValue(value);
      return;
    }
    const oldValue = getCellValue(sheet, row, col);
    overlay[key(sheet, row, col)] = value;
    record(ctx, sheet.getName(), row, ctx.colName || col, oldValue, value, ctx.reason);
  }

  // ctx.headers 가 있으면 빈칸이 아닌 열마다 한 줄씩, 없으면 행 전체를 한 줄로 기록합니다.
  function appendRow(sheet, rowValues, ctx) {
    if (!dryRun) {
      sheet.appendRow(rowValues);
      return;
    }
    const name = sheet.getName();
    const row = nextRow[name] = (nextRow[name] || sheet.getLastRow()) + 1;
    if (ctx.headers) {
      rowValues.forEach((v, i) => {
        if (v !== '' && v !== null && v !== undefined) record(ctx, name, row, ctx.headers[i], '', v, ctx.reason);
      });
    } else {
      record(ctx, name, row, '(행 전체)', '', JSON.stringify(rowValues), ctx.reason);
    }
  }

  // 처리로그. DRY_RUN 이면 처리로그에 쓰지 않고, 남기려던 결과를 시뮬레이션에 적어 둡니다.
  function log(logSheet, msgId, date, result, testNumber) {
    if (!dryRun) {
      logMessage(logSheet, msgId, date, result);
      return;
    }
    record({ msgId: msgId, testNumber: testNumber || '' }, logSheet.getName(), '', '', '', result,
           'DRY_RUN: 처리로그에 기록하지 않음');
  }

  // 검토용 메모 (DRY_RUN 일 때만 기록). 실제 실행에는 영향이 없습니다.
  function note(ctx, tab, text) {
    if (!dryRun) return;
    record(ctx, tab, '', '', '', text, ctx.reason || '');
  }

  return { dryRun, getCellValue, getRowValues, setCell, appendRow, log, note };
}

// 시뮬레이션 칸에 쓸 값: 수식으로 해석되지 않게 문자열로 바꾸고, Date 는 형식을 드러냅니다.
function toSimCell(v) {
  if (v === null || v === undefined) return '';
  let s;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    s = '[Date] ' + Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  } else {
    s = String(v);
  }
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
