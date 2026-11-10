/**
 * test/mock.js
 * Apps Script 서비스를 흉내 내는 메모리 목업 (셀 값만 저장, 서식·수식 계산은 하지 않음)
 */
const vm = require('vm');
const fs = require('fs');
const path = require('path');

const pad = (n) => String(n).padStart(2, '0');
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
  getLastRow() { let m = 0; for (const k of this.cells.keys()) m = Math.max(m, +k.split(',')[0]); return m; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  insertRowsAfter(_, n) { this.maxRows += n; }
  insertColumnsAfter(_, n) { this.maxCols += n; }
  deleteRows(start, n) {
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
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { out.push([]); for (let j = 0; j < this.nc; j++) out[i].push(this.sh.get(this.r + i, this.c + j)); } return out; }
  getDisplayValues() {
    return this.getValues().map((row) => row.map((v) => (v instanceof Date ? formatDate(v, '', 'yyyy-MM-dd') : v === true ? 'TRUE' : v === false ? 'FALSE' : String(v))));
  }
  getValue() { return this.sh.get(this.r, this.c); }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  setValue(v) { this.each((r, c) => this.sh.set(r, c, v)); return this; }
  setValues(vals) {
    if (vals.length !== this.nr || vals[0].length !== this.nc) throw new Error(`setValues 크기 불일치 ${vals.length}x${vals[0].length} vs ${this.nr}x${this.nc}`);
    this.each((r, c, i, j) => this.sh.set(r, c, vals[i][j])); return this;
  }
  setFormula(f) { this.sh.set(this.r, this.c, '', f); return this; }
  setFormulas(fs) { this.each((r, c, i, j) => this.sh.set(r, c, '', fs[i][j])); return this; }
  clearContent() { this.each((r, c) => this.sh.set(r, c, '')); return this; }
  setDataValidation(v) { this.each((r, c) => this.sh.dv.set(r + ',' + c, v)); return this; }
  getDataValidation() { return this.sh.dv.get(this.r + ',' + this.c) || null; }
  clearDataValidations() { this.sh.dv.clear(); return this; }
  sort(specs) {
    const rows = this.getValues();
    rows.sort((a, b) => { for (const s of specs) { const i = s.column - this.c; if (a[i] < b[i]) return -1; if (a[i] > b[i]) return 1; } return 0; });
    this.setValues(rows); return this;
  }
}

/** 목업 스프레드시트 + 서비스로 src/*.js를 불러온 실행 환경 만들기 */
function createEnv() {
  const sheets = [];
  const ss = chain({
    getSheetByName: (n) => sheets.find((s) => s.name === n) || null,
    insertSheet: (n) => { const s = new Sheet(n); sheets.push(s); return s; },
    getSheets: () => sheets.slice(),
    deleteSheet: (s) => sheets.splice(sheets.indexOf(s), 1),
  });
  const props = new Map();
  const triggers = [];
  const events = [];
  const cache = new Map();
  const mkEvent = (title, start, end, allDay) => ({
    title, start, end, allDay,
    getTitle() { return this.title; }, isAllDayEvent() { return this.allDay; },
    getStartTime() { return this.start; }, getEndTime() { return this.end; },
    getAllDayStartDate() { return this.start; }, getAllDayEndDate() { return this.end; },
    getId() { return 'id' + events.indexOf(this); },
  });
  const sameDay = (a, b) => a.toDateString() === b.toDateString();
  const calendar = {
    getEvents: (s, e) => events.filter((ev) => ev.start < e && ev.end > s),
    getEventsForDay: (d) => events.filter((ev) => sameDay(ev.start, d)),
    createEvent: (t, s, e) => { const ev = mkEvent(t, s, e, false); events.push(ev); return ev; },
    createAllDayEvent: (t, d) => { const ev = mkEvent(t, d, new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1), true); events.push(ev); return ev; },
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
    CalendarApp: { getDefaultCalendar: () => calendar },
    PropertiesService: { getDocumentProperties: () => ({
      getProperty: (k) => (props.has(k) ? props.get(k) : null), setProperty: (k, v) => props.set(k, String(v)),
    }) },
    CacheService: { getScriptCache: () => ({ get: (k) => cache.get(k) || null, put: (k, v) => cache.set(k, v) }) },
    LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }), getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Charts: { ChartType: {}, ChartHiddenDimensionStrategy: {} },
    Utilities: { formatDate },
  };
  vm.createContext(ctx);
  const dir = path.join(__dirname, '..', 'src');
  for (const f of fs.readdirSync(dir).filter((f) => f.endsWith('.js')).sort()) {
    vm.runInContext(fs.readFileSync(path.join(dir, f), 'utf8'), ctx, { filename: f });
  }
  return { ctx, ss, props, triggers, events, cache, run: (code) => vm.runInContext(code, ctx) };
}

module.exports = { createEnv, Sheet };
