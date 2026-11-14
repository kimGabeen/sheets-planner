/**
 * test/mock.js
 * Apps Script 서비스를 흉내 내는 메모리 목업 (셀 값만 저장, 서식·수식 계산은 하지 않음)
 */
const vm = require('vm');
const fs = require('fs');
const path = require('path');

const pad = (n) => String(n).padStart(2, '0');
// 성능 측정: Apps Script 서비스 호출 수 (종류별)
const calls = { read: 0, write: 0, struct: 0, formula: 0, calRead: 0, calWrite: 0, cache: 0, props: 0, flushRead: 0 };
let dirty = false; // 쓰기 뒤 첫 읽기는 시트 재계산을 기다림 (flush)
const countRead = () => { calls.read++; if (dirty) { calls.flushRead++; dirty = false; } };
const countWrite = (kind) => { calls[kind || 'write']++; dirty = true; };
function resetCalls() { Object.keys(calls).forEach((k) => { calls[k] = 0; }); dirty = false; }
function formatDate(d, tz, f) {
  return f.replace('yyyy', d.getFullYear()).replace('MM', pad(d.getMonth() + 1)).replace('dd', pad(d.getDate()))
    .replace('HH', pad(d.getHours())).replace('mm', pad(d.getMinutes()))
    .replace(/\bM\b/, d.getMonth() + 1).replace(/\bd\b/, d.getDate());
}

/** 아무 메서드나 받아 자기 자신을 돌려주는 객체 (서식 호출용) */
function chain(target) {
  return new Proxy(target, {
    get(t, k) {
      if (k in t) return t[k];
      if (typeof k === 'symbol') return undefined;
      return () => chain(t);
    },
  });
}

class Sheet {
  constructor(name) { this.name = name; this.cells = new Map(); this.maxRows = 100; this.maxCols = 26; this.dv = new Map(); }
  key(r, c) { return r + ',' + c; }
  get(r, c) { const v = this.cells.get(this.key(r, c)); return v === undefined ? '' : v.value; }
  set(r, c, value, formula) {
    if ((value === '' || value === null) && !formula) this.cells.delete(this.key(r, c));
    else this.cells.set(this.key(r, c), { value: value === null ? '' : value, formula });
  }
  getName() { return this.name; }
  getRange(r, c, nr, nc) {
    if (typeof r === 'string') return chain({});
    return new Range(this, r, c, nr || 1, nc || 1);
  }
  getDataRange() {
    let r = 1, c = 1;
    for (const k of this.cells.keys()) { const [a, b] = k.split(',').map(Number); r = Math.max(r, a); c = Math.max(c, b); }
    return new Range(this, 1, 1, r, c);
  }
  getLastRow() { countRead(); let m = 0; for (const k of this.cells.keys()) m = Math.max(m, +k.split(',')[0]); return m; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  insertRowsAfter(row, n) { countWrite('struct'); this.shiftRows(row + 1, n); }
  insertRowsBefore(row, n) { countWrite('struct'); this.shiftRows(row, n); }
  /** from 행부터 아래를 n줄 밀기 (값·검증) */
  shiftRows(from, n) {
    const move = (map) => {
      const next = new Map();
      for (const [k, v] of map) {
        const [r, c] = k.split(',').map(Number);
        next.set((r >= from ? r + n : r) + ',' + c, v);
      }
      return next;
    };
    this.cells = move(this.cells);
    this.dv = move(this.dv);
    this.maxRows += n;
  }
  insertColumnsAfter(_, n) { this.maxCols += n; }
  deleteRows(start, n) {
    countWrite('struct');
    const next = new Map();
    for (const [k, v] of this.cells) {
      const [r, c] = k.split(',').map(Number);
      if (r < start) next.set(k, v); else if (r >= start + n) next.set(this.key(r - n, c), v);
    }
    this.cells = next;
  }
  clear() { this.cells.clear(); this.dv.clear(); }
  getCharts() { return []; }
  isSheetHidden() { return false; }
}
['setHiddenGridlines', 'setColumnWidth', 'setColumnWidths', 'setRowHeight', 'setRowHeights', 'setFrozenRows',
  'hideColumns', 'showColumns', 'hideSheet', 'showSheet', 'clearConditionalFormatRules', 'setConditionalFormatRules',
  'insertChart', 'removeChart'].forEach((m) => { Sheet.prototype[m] = function () { return this; }; });
Sheet.prototype.newChart = function () { return chain({ build: () => ({}) }); };

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); return chain(this); }
  getSheet() { return this.sh; }
  getRow() { return this.r; }
  getColumn() { return this.c; }
  getLastRow() { return this.r + this.nr - 1; }
  getLastColumn() { return this.c + this.nc - 1; }
  each(fn) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) fn(this.r + i, this.c + j, i, j); }
  getValues() { countRead(); const out = []; for (let i = 0; i < this.nr; i++) { out.push([]); for (let j = 0; j < this.nc; j++) out[i].push(this.sh.get(this.r + i, this.c + j)); } return out; }
  getDisplayValues() {
    return this.getValues().map((row) => row.map((v) => (v instanceof Date ? formatDate(v, '', 'yyyy-MM-dd') : v === true ? 'TRUE' : v === false ? 'FALSE' : String(v))));
  }
  getValue() { countRead(); return this.sh.get(this.r, this.c); }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  setValue(v) { countWrite(); this.each((r, c) => this.sh.set(r, c, v)); return this; }
  setValues(vals) {
    countWrite();
    if (vals.length !== this.nr || vals[0].length !== this.nc) throw new Error(`setValues 크기 불일치 ${vals.length}x${vals[0].length} vs ${this.nr}x${this.nc}`);
    this.each((r, c, i, j) => this.sh.set(r, c, vals[i][j])); return this;
  }
  setFormula(f) { countWrite('formula'); this.sh.set(this.r, this.c, '', f); return this; }
  setFormulas(fs) { this.each((r, c, i, j) => this.sh.set(r, c, '', fs[i][j])); return this; }
  clearContent() { countWrite(); this.each((r, c) => this.sh.set(r, c, '')); return this; }
  setDataValidation(v) { this.each((r, c) => this.sh.dv.set(r + ',' + c, v)); return this; }
  getDataValidation() { return this.sh.dv.get(this.r + ',' + this.c) || null; }
  clearDataValidations() { this.sh.dv.clear(); return this; }
  sort(specs) {
    countWrite('struct');
    const rows = this.getValues(); calls.read--;
    rows.sort((a, b) => { for (const s of specs) { const i = s.column - this.c; if (a[i] < b[i]) return -1; if (a[i] > b[i]) return 1; } return 0; });
    this.setValues(rows); return this;
  }
}

/** 목업 스프레드시트 + 서비스로 src/*.js를 불러온 실행 환경 만들기 */
function createEnv() {
  const sheets = [];
  const copies = []; // 백업 사본 이름
  const ss = chain({
    getSheetByName: (n) => sheets.find((s) => s.name === n) || null,
    insertSheet: (n) => { const s = new Sheet(n); sheets.push(s); return s; },
    getSheets: () => sheets.slice(),
    getId: () => 'test-sheet-id',
    copy: (name) => { copies.push(name); return { getName: () => name }; },
    deleteSheet: (s) => sheets.splice(sheets.indexOf(s), 1),
  });
  const props = new Map(); // 스크립트 속성
  const docProps = new Map(); // 문서 속성 (예전 버전 값 옮기기 테스트용)
  const store = (m) => ({ getProperty: (k) => (calls.props++, m.has(k) ? m.get(k) : null), setProperty: (k, v) => (calls.props++, m.set(k, String(v))) });
  const triggers = [];
  const events = [];
  let nextId = 1;
  const cache = new Map();
  const mkEvent = (title, start, end, allDay, opts) => ({
    title, start, end, allDay, description: (opts && opts.description) || '',
    getDescription() { return this.description; },
    setDescription(d) { this.description = d; return this; },
    getTitle() { return this.title; }, isAllDayEvent() { return this.allDay; },
    getStartTime() { return this.start; }, getEndTime() { return this.end; },
    getAllDayStartDate() { return this.start; }, getAllDayEndDate() { return this.end; },
    id: 'id' + (nextId++),
    getId() { return this.id; },
    setTitle(t) { calls.calWrite++; this.title = t; return this; },
    setTime(st, en) { calls.calWrite++; this.start = st; this.end = en; this.allDay = false; return this; },
    setAllDayDate(d) { calls.calWrite++; this.start = d; this.end = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1); this.allDay = true; return this; },
    deleteEvent() { calls.calWrite++; events.splice(events.indexOf(this), 1); },
  });
  const sameDay = (a, b) => a.toDateString() === b.toDateString();
  const calendar = {
    getEvents: (s, e) => (calls.calRead++, events).filter((ev) => ev.start < e && ev.end > s),
    getEventsForDay: (d) => (calls.calRead++, events).filter((ev) => sameDay(ev.start, d)),
    getEventById: (id) => (calls.calRead++, events).find((ev) => ev.id === id) || null,
    createEvent: (t, s, e, o) => { calls.calWrite++; const ev = mkEvent(t, s, e, false, o); events.push(ev); return ev; },
    createAllDayEvent: (t, d, o) => { calls.calWrite++; const ev = mkEvent(t, d, new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1), true, o); events.push(ev); return ev; },
  };
  // "대한민국의 휴일" 공개 캘린더 (subscribed=false면 구독 전 상태)
  const holidays = { subscribed: false, list: [
    ['2026-11-11', '테스트 기념일'], ['2026-12-25', '기독탄신일'], ['2027-01-01', '신정'],
  ] };
  const holidayCal = {
    getEvents: () => (calls.calRead++, holidays.list).map(([k, t]) => {
      const [y, m, d] = k.split('-').map(Number);
      return { getTitle: () => t, getAllDayStartDate: () => new Date(y, m - 1, d), getAllDayEndDate: () => new Date(y, m - 1, d + 1) };
    }),
  };
  const triggerBuilder = (name) => chain({ create: () => { triggers.push({ getHandlerFunction: () => name }); } });
  const ctx = {
    console, Date, Math, JSON, String, Number, Object, Array, parseInt, isNaN,
    SpreadsheetApp: chain({
      getActiveSpreadsheet: () => ss, flush: () => {},
      newDataValidation: () => chain({ build: () => ({ dv: true }) }),
      newConditionalFormatRule: () => chain({ build: () => ({}) }),
      BorderStyle: {}, getUi: () => chain({}),
    }),
    ScriptApp: {
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: (t) => triggers.splice(triggers.indexOf(t), 1),
      newTrigger: triggerBuilder,
    },
    CalendarApp: {
      getDefaultCalendar: () => calendar,
      getCalendarById: () => (holidays.subscribed ? holidayCal : null),
      subscribeToCalendar: () => { holidays.subscribed = true; return holidayCal; },
    },
    PropertiesService: {
      getScriptProperties: () => store(props),
      getDocumentProperties: () => store(docProps),
    },
    HtmlService: chain({}),
    CacheService: { getScriptCache: () => ({ get: (k) => (calls.cache++, cache.get(k) || null), put: (k, v) => (calls.cache++, cache.set(k, v)) }) },
    LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }), getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
    Charts: { ChartType: {}, ChartHiddenDimensionStrategy: {} },
    Utilities: { formatDate },
  };
  vm.createContext(ctx);
  const dir = path.join(__dirname, '..', 'src');
  for (const f of fs.readdirSync(dir).filter((f) => f.endsWith('.js')).sort()) {
    vm.runInContext(fs.readFileSync(path.join(dir, f), 'utf8'), ctx, { filename: f });
  }
  return { ctx, ss, props, docProps, holidays, copies, triggers, events, cache, run: (code) => vm.runInContext(code, ctx) };
}

module.exports = { createEnv, Sheet, calls, resetCalls };
