// Flow 의 .gs 파일을 Node 에서 불러오는 테스트 도구입니다.
// Apps Script 서비스(SpreadsheetApp, GmailApp, UrlFetchApp, PropertiesService, Utilities, Logger)를
// 메모리 안의 가짜 객체로 바꿔 끼웁니다. 실제 시트·메일·API 에는 접속하지 않습니다.
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const FLOW_DIR = path.join(__dirname, '..');
const GS_FILES = ['Config.gs', 'Utils.gs', 'Parser.gs', 'Llm.gs', 'Writer.gs',
                  'Initiation.gs', 'Progress.gs', 'Github.gs', 'Main.gs'];

// ---- 가짜 시트 -------------------------------------------------------------
class FakeSheet {
  constructor(name, rows) {
    this.name = name;
    this.data = (rows || []).map(r => r.slice());
    this.formulas = new Set(); // 'r,c' — 수식이 든 칸 (값은 "" 여도 getLastRow 에 잡힘)
  }
  markFormula(r, c) { this.formulas.add(`${r},${c}`); }
  getName() { return this.name; }
  _has(v) { return v !== '' && v !== null && v !== undefined; }
  getLastRow() {
    let last = 0;
    for (let r = this.data.length; r >= 1; r--) if ((this.data[r - 1] || []).some(v => this._has(v))) { last = r; break; }
    for (const k of this.formulas) last = Math.max(last, Number(k.split(',')[0]));
    return last;
  }
  getLastColumn() {
    let m = 0;
    for (const row of this.data) for (let c = row.length; c >= 1; c--) if (this._has(row[c - 1])) { m = Math.max(m, c); break; }
    return m;
  }
  getMaxRows() { return Math.max(1000, this.data.length, this.getLastRow()); }
  get(r, c) { const v = (this.data[r - 1] || [])[c - 1]; return v === undefined ? '' : v; }
  set(r, c, v) {
    while (this.data.length < r) this.data.push([]);
    const row = this.data[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  getRange(r, c, nr, nc) {
    if (!Number.isInteger(r) || !Number.isInteger(c)) throw new Error(`잘못된 범위 (${r}, ${c})`);
    nr = nr || 1; nc = nc || 1;
    const s = this;
    const range = {
      getValue: () => s.get(r, c),
      setValue: v => { s.set(r, c, v); return range; },
      getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => s.get(r + i, c + j))),
      setValues: vals => { vals.forEach((row, i) => row.forEach((v, j) => s.set(r + i, c + j, v))); return range; },
    };
    return range;
  }
  appendRow(values) {
    const r = this.getLastRow() + 1;
    values.forEach((v, i) => this.set(r, i + 1, v));
  }
  snapshot() { return JSON.stringify(this.data); }
}

class FakeSpreadsheet {
  constructor(id, sheets) {
    this.id = id;
    this.sheets = {};
    (sheets || []).forEach(s => { this.sheets[s.getName()] = s; });
    this.inserted = [];
  }
  getId() { return this.id; }
  getName() { return `사본(${this.id})`; }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { const s = new FakeSheet(n); this.sheets[n] = s; this.inserted.push(n); return s; }
}

// ---- 가짜 메일 -------------------------------------------------------------
function fakeMessage({ id, subject, from, to, cc, date, body }) {
  return {
    getId: () => id, getSubject: () => subject, getFrom: () => from,
    getTo: () => to || '', getCc: () => cc || '',
    getDate: () => (date instanceof Date ? date : new Date(date)), getPlainBody: () => body,
  };
}
function fakeThread(id, messages) {
  return { getId: () => id, getMessages: () => messages.map(fakeMessage) };
}

// ---- Utilities.formatDate (yyyy MM dd HH mm ss 만 지원) ---------------------
function formatDate(date, tz, pattern) {
  const parts = {};
  new Intl.DateTimeFormat('en-CA', {
    timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
  }).formatToParts(date).forEach(p => { parts[p.type] = p.value; });
  const map = { yyyy: parts.year, MM: parts.month, dd: parts.day, HH: parts.hour, mm: parts.minute, ss: parts.second };
  return pattern.replace(/yyyy|MM|dd|HH|mm|ss/g, t => map[t]);
}

// ---- Flow 불러오기 ---------------------------------------------------------
// opts: { props, spreadsheets: {id: FakeSpreadsheet}, initThreads, gsThreads, llmResponse(fn|obj) }
function loadFlow(opts) {
  opts = opts || {};
  const props = Object.assign({}, opts.props || {});
  const logs = [];
  const fetches = [];
  const searches = [];

  const sandbox = {
    console,
    Logger: { log: m => logs.push(String(m)) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null) }) },
    Utilities: {
      formatDate,
      base64Encode: b => Buffer.from(b).toString('base64'),
      base64Decode: s => Array.from(Buffer.from(s, 'base64')),
      newBlob: d => ({ getBytes: () => Array.from(Buffer.from(String(d))), getDataAsString: () => Buffer.from(d).toString('utf8') }),
    },
    SpreadsheetApp: {
      openById: id => {
        const ss = (opts.spreadsheets || {})[id];
        if (!ss) throw new Error(`시트 없음: ${id}`);
        return ss;
      },
    },
    GmailApp: {
      search: q => { searches.push(q); return opts.initThreads || []; },
      getUserLabelByName: name => ({ getThreads: (start, n) => (opts.gsThreads || []).slice(start, start + n) }),
    },
    UrlFetchApp: {
      fetch: (url, options) => {
        fetches.push({ url, options });
        if (url.includes('generativelanguage.googleapis.com')) {
          const payload = JSON.parse(options.payload);
          const r = typeof opts.llmResponse === 'function' ? opts.llmResponse(payload) : opts.llmResponse;
          // 실패 흉내: { __status: 503 } / { __raw: '깨진 JSON' } / { __finish: 'SAFETY' } / { __throw: '네트워크' }
          if (r && r.__throw) throw new Error(r.__throw);
          if (r && r.__status) {
            const body = r.__body !== undefined ? r.__body : JSON.stringify({ error: { code: r.__status, message: '흉내 오류' } });
            return { getContentText: () => body, getResponseCode: () => r.__status };
          }
          const cand = r && r.__finish
            ? { finishReason: r.__finish }
            : { content: { parts: [{ text: r && r.__raw !== undefined ? r.__raw : JSON.stringify(r) }] }, finishReason: 'STOP' };
          const text = JSON.stringify({ candidates: [cand] });
          return { getContentText: () => text, getResponseCode: () => 200 };
        }
        return { getContentText: () => '{"message":"not found"}', getResponseCode: () => 404 };
      },
    },
  };
  vm.createContext(sandbox);

  const code = GS_FILES.map(f => `// ---- ${f}\n` + fs.readFileSync(path.join(FLOW_DIR, f), 'utf8')).join('\n') +
    '\n;globalThis.__consts = { CONFIG, ALLOWED_COLUMNS, VALID_STATUSES, SIM_HEADERS };';
  vm.runInContext(code, sandbox, { filename: 'flow-bundle.gs' });

  return Object.assign(sandbox, sandbox.__consts, { logs, fetches, searches, props });
}

// LLM 에 실제로 보낸 본문 (가장 최근 호출)
function lastLlmBody(flow) {
  const f = flow.fetches.filter(x => x.url.includes('generativelanguage')).pop();
  if (!f) return null;
  return JSON.parse(f.options.payload).contents[0].parts[0].text;
}

module.exports = { FakeSheet, FakeSpreadsheet, fakeThread, loadFlow, lastLlmBody, formatDate };
