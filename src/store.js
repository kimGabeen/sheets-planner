/**
 * store.js
 * 날짜별 저장소: 할 일·습관·기분·일정 체크·주간 정보를 날짜(주) 기준으로 모두 보관
 * - 위클리·오늘 시트와 웹앱은 이 저장소를 읽고 쓴다 → 어느 날짜·주로든 이동 가능, 개수 제한 없음
 * - 기록 시트(분석용)는 이 저장소에서 만들어진다 (형태 그대로)
 */

const TODO_STORE_COLS_ = 8; // 날짜 | 순서 | 완료 | 시작 | 끝 | 내용 | 메모 | 이월표시
const DAY_STORE_COLS_ = 4; // 날짜 | 구분 | 항목 | 값
const WEEK_STORE_COLS_ = 5; // 주차(월요일) | 구분 | 순서 | 완료 | 내용
const CARRY_FLAG_ = { MISSED: '미완', CARRIED: '넘어옴' }; // 이월표시
const WEEK_INFO_ = ['목표', '보상', '확언'];
const WEEK_MUST_ = '꼭끝낼일';

/** 한 번의 실행 안에서 같은 저장소를 여러 번 읽지 않게 (쓰면 비움) */
let STORE_MEMO_ = {};

/* ---------- 공통 ---------- */

/** 저장소 시트 만들기 (데이터는 유지) */
function buildStoreSheets_(theme) {
  const defs = [
    [CONFIG.SHEETS.TODO_STORE, ['날짜', '순서', '완료', '시작', '끝', '내용', '메모', '이월표시'], [100, 50, 50, 60, 60, 260, 220, 70]],
    [CONFIG.SHEETS.DAY_STORE, ['날짜', '구분', '항목', '값'], [100, 80, 240, 80]],
    [CONFIG.SHEETS.WEEK_STORE, ['주차', '구분', '순서', '완료', '내용'], [100, 80, 50, 50, 300]],
  ];
  defs.forEach(function (d) {
    const sh = ensureSheet_(d[0]);
    ensureSize_(sh, 50, d[1].length);
    if (!sh.getRange(1, 1).getValue()) sh.getRange(1, 1, 1, d[1].length).setValues([d[1]]);
    if (theme) styleHeader_(sh.getRange(1, 1, 1, d[1].length), theme);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
    d[2].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  });
  // 시간·글자는 글자 그대로 (시트가 날짜·숫자로 바꾸지 않게)
  sheet_(CONFIG.SHEETS.TODO_STORE).getRange('D:H').setNumberFormat('@');
  sheet_(CONFIG.SHEETS.DAY_STORE).getRange('B:C').setNumberFormat('@');
  sheet_(CONFIG.SHEETS.WEEK_STORE).getRange('B:B').setNumberFormat('@');
  sheet_(CONFIG.SHEETS.WEEK_STORE).getRange('E:E').setNumberFormat('@');
}

/** 저장소 전체 읽기 → [{row, key(날짜), v(값 배열)}] */
function readStore_(name, cols) {
  if (STORE_MEMO_[name]) return STORE_MEMO_[name];
  const out = readAll_(sheet_(name), cols).map(function (v, i) {
    const d = dateOnly_(v[0]);
    return { row: i + 2, key: d ? dateKey_(d) : '', v: v };
  }).filter(function (r) { return r.key; });
  STORE_MEMO_[name] = out;
  return out;
}

/**
 * 조건에 맞는 줄을 새 줄로 바꾸기 (다른 데이터는 그대로)
 * - 맞는 줄이 한 덩어리로 붙어 있으면 그 자리에 덮어쓰고, 모자라면 아래에 추가·남으면 삭제 → 보통 시트 호출 1~3번
 * - 시트 전체 정렬은 하지 않음 (읽을 때 날짜·순서로 정리하므로)
 * - 메모리 사본도 같이 고쳐서 다시 읽지 않음
 */
function replaceStore_(name, cols, match, rows) {
  const sh = sheet_(name);
  const all = readStore_(name, cols);
  const hits = all.filter(match);
  const pos = hits.map(function (r) { return r.row; });
  const contiguous = pos.length > 0 && pos[pos.length - 1] - pos[0] === pos.length - 1;
  if (contiguous) {
    const start = pos[0], n = Math.min(pos.length, rows.length);
    if (n) sh.getRange(start, 1, n, cols).setValues(rows.slice(0, n));
    if (pos.length > rows.length) sh.deleteRows(start + rows.length, pos.length - rows.length);
    else if (rows.length > pos.length) {
      sh.insertRowsAfter(start + pos.length - 1, rows.length - pos.length);
      sh.getRange(start + pos.length, 1, rows.length - pos.length, cols).setValues(rows.slice(pos.length));
    }
    // 메모리 사본: 그 덩어리만 바꾸고 아래 줄 번호를 밀거나 당김
    const shift = rows.length - pos.length;
    const before = all.filter(function (r) { return r.row < start; });
    const after = all.filter(function (r) { return r.row > pos[pos.length - 1]; }).map(function (r) { return { row: r.row + shift, key: r.key, v: r.v }; });
    const mid = rows.map(function (v, k) { return { row: start + k, key: dateKey_(dateOnly_(v[0])), v: v }; });
    STORE_MEMO_[name] = before.concat(mid, after);
    return;
  }
  // 그 날짜가 처음이면: 맨 아래에 추가 (기존 줄 번호는 그대로라 메모리 사본에 붙이기만)
  if (!pos.length) {
    if (!rows.length) return;
    const at = Math.max(sh.getLastRow(), 1) + 1;
    ensureSize_(sh, at + rows.length + 5, cols);
    sh.getRange(at, 1, rows.length, cols).setValues(rows);
    STORE_MEMO_[name] = all.concat(rows.map(function (v, k) { return { row: at + k, key: dateKey_(dateOnly_(v[0])), v: v }; }));
    return;
  }
  // 흩어져 있으면: 지우고 맨 아래에 추가 (드문 경우라 메모리 사본은 다시 읽게 비움)
  deleteRowList_(sh, pos);
  if (rows.length) {
    const at = Math.max(sh.getLastRow(), 1) + 1;
    ensureSize_(sh, at + rows.length + 5, cols);
    sh.getRange(at, 1, rows.length, cols).setValues(rows);
  }
  delete STORE_MEMO_[name];
}

/* ---------- 할 일 ---------- */
// 할 일 = {done, start, end, text, memo, flag}  (flag: '' / '미완'(전날에 남은 원래 일) / '넘어옴'(다음 날로 복사된 일))

/** 할 일 → 시간 글자 ('11:00-13:00' / '15:00' / '') */
function todoTime_(t) {
  return t.start ? t.start + (t.end ? '-' + t.end : '') : '';
}

/** 시간 글자 → {start, end} */
function splitRange_(text) {
  const r = normRange_(text);
  return { start: r.slice(0, 5), end: r.slice(6, 11) };
}

/** 화면 글자 (이월표시 붙임) */
function todoDisplay_(t) {
  return (t.flag === CARRY_FLAG_.MISSED ? CONFIG.CARRY.MISSED : t.flag === CARRY_FLAG_.CARRIED ? CONFIG.CARRY.CARRIED : '') + t.text;
}

/** 정렬: 미완료(시작 시간순) → 시간 없음(입력 순) → 체크·"미완" 일 */
function sortTodoItems_(items) {
  const group = function (t) { return t.done || t.flag === CARRY_FLAG_.MISSED ? 2 : t.start ? 0 : 1; };
  return items.map(function (t, i) { return { t: t, i: i }; }).sort(function (a, b) {
    const g = group(a.t) - group(b.t);
    if (g) return g;
    if (group(a.t) === 0 && a.t.start !== b.t.start) return a.t.start < b.t.start ? -1 : 1;
    return a.i - b.i;
  }).map(function (x) { return x.t; });
}

/** 입력 정리: 빈 내용 버림, 시간 없고 내용에 '11-13'이 있으면 시간으로, 이월 기호는 표시로 */
function cleanTodo_(t) {
  let text = String(t.text || '').trim();
  let flag = t.flag || '';
  if (isMissed_(text)) flag = CARRY_FLAG_.MISSED;
  else if (text.indexOf(CONFIG.CARRY.CARRIED.trim()) === 0 && !flag) flag = CARRY_FLAG_.CARRIED;
  text = stripCarry_(text);
  let start = normTime_(t.start), end = normTime_(t.end);
  if (!start) {
    const split = splitTimeFromText_(text);
    if (split) { text = split.text; start = split.time.slice(0, 5); end = split.time.slice(6, 11); }
  }
  if (end && end <= start) end = '';
  return { done: t.done === true, start: start, end: start ? end : '', text: text, memo: String(t.memo || '').replace(/\s+$/, ''), flag: flag };
}

/** 날짜 범위의 할 일 → {날짜키: [할 일, …]} (순서대로) */
function todosByDate_(from, to) {
  const a = from ? dateKey_(from) : '', b = to ? dateKey_(to) : '9999';
  const out = {};
  readStore_(CONFIG.SHEETS.TODO_STORE, TODO_STORE_COLS_).forEach(function (r) {
    if (r.key < a || r.key > b) return;
    (out[r.key] = out[r.key] || []).push({ order: Number(r.v[1]) || 0, done: r.v[2] === true || r.v[2] === 'TRUE',
      start: normTime_(r.v[3]), end: normTime_(r.v[4]), text: String(r.v[5]), memo: String(r.v[6] || ''), flag: String(r.v[7] || '') });
  });
  Object.keys(out).forEach(function (k) {
    out[k].sort(function (x, y) { return x.order - y.order; });
    out[k].forEach(function (t) { delete t.order; });
  });
  return out;
}

/** 그날 할 일 (화면 순서) */
function getTodos_(date) {
  return todosByDate_(date, date)[dateKey_(date)] || [];
}

/** 그날 할 일 전체 저장 (정리 → 정렬 → 순서 다시 매김). 개수 제한 없음 */
function setTodos_(date, items) {
  const key = dateKey_(date);
  const list = sortTodoItems_(items.map(cleanTodo_).filter(function (t) { return t.text; }));
  replaceStore_(CONFIG.SHEETS.TODO_STORE, TODO_STORE_COLS_, function (r) { return r.key === key; },
    list.map(function (t, i) { return [date, i + 1, t.done, t.start, t.end, t.text, t.memo, t.flag]; }));
  return list;
}

/* ---------- 습관·기분·일정 체크 (하루 단위) ---------- */

/** 날짜 범위의 하루 데이터 → {날짜키: {habits:{이름:bool}, mood:{score, memo}, checks:{열쇠:true}}} */
function dayDataByDate_(from, to) {
  const a = dateKey_(from), b = dateKey_(to);
  const out = {};
  readStore_(CONFIG.SHEETS.DAY_STORE, DAY_STORE_COLS_).forEach(function (r) {
    if (r.key < a || r.key > b) return;
    const d = out[r.key] = out[r.key] || { habits: {}, mood: { score: '', memo: '' }, checks: {} };
    const type = String(r.v[1]), item = String(r.v[2]), val = r.v[3];
    if (type === CONFIG.CATEGORY.HABIT) d.habits[item] = val === true || val === 'TRUE';
    else if (type === CONFIG.CATEGORY.MOOD) d.mood = { score: Number(val) || '', memo: item };
    else if (type === '일정체크') d.checks[item] = val === true || val === 'TRUE';
  });
  return out;
}

function emptyDay_() { return { habits: {}, mood: { score: '', memo: '' }, checks: {} }; }

/** 그날 하루 데이터 */
function getDayData_(date) {
  return dayDataByDate_(date, date)[dateKey_(date)] || emptyDay_();
}

/** 습관 체크 저장 */
function setHabitDone_(date, name, done) {
  const key = dateKey_(date);
  replaceStore_(CONFIG.SHEETS.DAY_STORE, DAY_STORE_COLS_, function (r) {
    return r.key === key && r.v[1] === CONFIG.CATEGORY.HABIT && String(r.v[2]) === name;
  }, done ? [[date, CONFIG.CATEGORY.HABIT, name, true]] : []);
}

/** 여러 습관 체크를 한 번에 저장 ({이름: bool}) */
function setHabitsDone_(date, map) {
  const key = dateKey_(date);
  const names = Object.keys(map);
  replaceStore_(CONFIG.SHEETS.DAY_STORE, DAY_STORE_COLS_, function (r) {
    return r.key === key && r.v[1] === CONFIG.CATEGORY.HABIT && names.indexOf(String(r.v[2])) >= 0;
  }, names.filter(function (n) { return map[n]; }).map(function (n) { return [date, CONFIG.CATEGORY.HABIT, n, true]; }));
}

/** 기분 저장 (점수 1~5 또는 '', 한 줄 메모) */
function setMoodData_(date, score, memo) {
  const key = dateKey_(date);
  const s = Number(score) || '', m = String(memo || '').trim();
  replaceStore_(CONFIG.SHEETS.DAY_STORE, DAY_STORE_COLS_, function (r) {
    return r.key === key && r.v[1] === CONFIG.CATEGORY.MOOD;
  }, s || m ? [[date, CONFIG.CATEGORY.MOOD, m, s]] : []);
}

/** 일정 완료 체크 저장 (열쇠 = 일정 ID 또는 날짜+시간+제목, 수업은 수업명) */
function setItemCheck_(date, itemKey, done) {
  const key = dateKey_(date);
  replaceStore_(CONFIG.SHEETS.DAY_STORE, DAY_STORE_COLS_, function (r) {
    return r.key === key && r.v[1] === '일정체크' && String(r.v[2]) === itemKey;
  }, done ? [[date, '일정체크', itemKey, true]] : []);
}

/** 습관 이름 바꾸기: from 날짜부터의 체크를 새 이름으로 (지난 기록은 그대로) */
function renameHabitFrom_(from, oldName, newName) {
  const a = dateKey_(from);
  const rows = readStore_(CONFIG.SHEETS.DAY_STORE, DAY_STORE_COLS_).filter(function (r) {
    return r.key >= a && r.v[1] === CONFIG.CATEGORY.HABIT && String(r.v[2]) === oldName;
  });
  if (!rows.length) return;
  const keys = rows.map(function (r) { return r.row; });
  replaceStore_(CONFIG.SHEETS.DAY_STORE, DAY_STORE_COLS_, function (r) { return keys.indexOf(r.row) >= 0; },
    rows.map(function (r) { return [r.v[0], CONFIG.CATEGORY.HABIT, newName, r.v[3]]; }));
}

/* ---------- 주간 (목표·보상·확언·꼭 끝낼 일) ---------- */

/** 그 주 데이터 → {info: [목표, 보상, 확언], must: [{done, text}]} */
function getWeekData_(monday) {
  const key = dateKey_(mondayOf_(monday));
  const info = ['', '', ''], must = [];
  readStore_(CONFIG.SHEETS.WEEK_STORE, WEEK_STORE_COLS_).forEach(function (r) {
    if (r.key !== key) return;
    const type = String(r.v[1]);
    const i = WEEK_INFO_.indexOf(type);
    if (i >= 0) info[i] = String(r.v[4]);
    else if (type === WEEK_MUST_) must.push({ order: Number(r.v[2]) || 0, done: r.v[3] === true || r.v[3] === 'TRUE', text: String(r.v[4]) });
  });
  must.sort(function (a, b) { return a.order - b.order; });
  return { info: info, must: must.map(function (m) { return { done: m.done, text: m.text }; }) };
}

/** 목표·보상·확언 저장 */
function setWeekInfo_(monday, info) {
  const mon = mondayOf_(monday), key = dateKey_(mon);
  replaceStore_(CONFIG.SHEETS.WEEK_STORE, WEEK_STORE_COLS_, function (r) {
    return r.key === key && WEEK_INFO_.indexOf(String(r.v[1])) >= 0;
  }, WEEK_INFO_.map(function (t, i) { return [mon, t, 0, false, String(info[i] || '').trim()]; })
    .filter(function (r) { return r[4]; }));
}

/** 꼭 끝낼 일 저장 (개수 제한 없음, 빈 줄은 버림) */
function setWeekMust_(monday, list) {
  const mon = mondayOf_(monday), key = dateKey_(mon);
  const clean = list.map(function (m) { return { done: m.done === true, text: String(m.text || '').trim() }; })
    .filter(function (m) { return m.text; });
  replaceStore_(CONFIG.SHEETS.WEEK_STORE, WEEK_STORE_COLS_, function (r) {
    return r.key === key && String(r.v[1]) === WEEK_MUST_;
  }, clean.map(function (m, i) { return [mon, WEEK_MUST_, i + 1, m.done, m.text]; }));
  return clean;
}

/* ---------- 예전 데이터 옮기기 (한 번만) ---------- */

/** 저장소가 아직 없으면 백업 → 옮기기 → 화면 다시 그리기 (트리거·웹앱·setup 어디서든 먼저 실행) */
function ensureStore_(skipBuild) {
  const props = props_();
  if (props.getProperty('STORE_VERSION') === CONFIG.STORE_VERSION) return;
  if (!sheet_(CONFIG.SHEETS.WEEKLY)) return; // setup 전
  migrateToStore_();
  props.setProperty('STORE_VERSION', CONFIG.STORE_VERSION);
  if (!skipBuild) buildAllLayouts_(getTheme_()); // 새 수식·화면으로 (사용자가 setup을 다시 안 돌려도)
  bumpCache_();
}

/**
 * 예전 데이터 → 저장소 (원본은 지우지 않음)
 * - 이번 주(위클리 시트): 할 일(시간·메모·이월표시 포함)·습관·기분·꼭 끝낼 일·목표/보상/확언
 * - 그 밖의 날짜(기록 시트): 할 일·습관·기분·주간목표
 * - 오늘 시트 "넘어온 일"로 빠져 있던 일은 오늘 할 일로
 * 반환·설정 시트에 옮긴 줄 수를 남김
 */
function migrateToStore_() {
  const ss = ss_();
  let backup = '';
  try {
    backup = '백업_' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
    ss.copy(backup); // 옮기기 전에 시트 전체 사본
  } catch (err) {
    backup = '실패(' + err.message + ')';
  }
  buildStoreSheets_(getTheme_());
  const W = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY);
  const mon = dateOnly_(sh.getRange(W.WEEK_START.row, W.WEEK_START.col).getValue());
  const weekMon = mon ? mondayOf_(mon) : null;
  const weekKeys = {};
  for (let d = 0; d < 7 && weekMon; d++) weekKeys[dateKey_(addDays_(weekMon, d))] = true;

  const todoRows = [], dayRows = [], weekRows = [];
  const C = CONFIG.CATEGORY;
  // 1) 기록 시트 (이번 주 날짜는 위클리 시트가 더 자세하므로 건너뜀)
  const order = {};
  readArchive_().forEach(function (r) {
    if (weekKeys[r.key] && r.cat !== C.MONTH_GOAL && r.cat !== C.WEEK_GOAL) return;
    const date = dateOnly_(r.key);
    if (r.cat === C.TODO) {
      order[r.key] = (order[r.key] || 0) + 1;
      todoRows.push([date, order[r.key], r.done === true, '', '', String(r.item), '', r.value === CONFIG.CARRY.VALUE ? CARRY_FLAG_.MISSED : '']);
    } else if (r.cat === C.HABIT) {
      if (r.done === true) dayRows.push([date, C.HABIT, String(r.item), true]);
    } else if (r.cat === C.MOOD) {
      dayRows.push([date, C.MOOD, String(r.item), Number(r.value) || '']);
    } else if (r.cat === C.WEEK_GOAL && !(weekMon && r.key === dateKey_(weekMon))) {
      order['w' + r.key] = (order['w' + r.key] || 0) + 1;
      weekRows.push([date, WEEK_MUST_, order['w' + r.key], r.done === true, String(r.item)]);
    }
  });

  // 2) 위클리 시트의 이번 주
  if (weekMon) {
    const habits = getHabits_();
    for (let d = 0; d < 7; d++) {
      const date = addDays_(weekMon, d);
      const block = sh.getRange(W.TODO_START, W.checkCol(d), CONFIG.TODO_ROWS, 3);
      const vals = block.getValues(), disp = block.getDisplayValues();
      const memos = sh.getMaxColumns() >= W.memoCol(d) ? sh.getRange(W.TODO_START, W.memoCol(d), CONFIG.TODO_ROWS, 1).getValues() : [];
      let n = 0;
      vals.forEach(function (r, i) {
        const t = cleanTodo_({ done: r[0] === true, start: splitRange_(disp[i][1]).start, end: splitRange_(disp[i][1]).end,
          text: String(r[2]), memo: memos[i] ? memos[i][0] : '' });
        if (t.text) todoRows.push([date, ++n, t.done, t.start, t.end, t.text, t.memo, t.flag]);
      });
      habits.forEach(function (name, i) {
        if (name && sh.getRange(W.HABIT_START + i, W.HABIT_DAY_COL + d).getValue() === true) dayRows.push([date, C.HABIT, name, true]);
      });
      const score = moodScore_(sh.getRange(W.MOOD, W.checkCol(d)).getValue());
      const memo = String(sh.getRange(W.MOOD_MEMO, W.checkCol(d)).getValue()).trim();
      if (score || memo) dayRows.push([date, C.MOOD, memo, score]);
    }
    sh.getRange(W.MUST_START, 1, CONFIG.MUST_COUNT, 2).getValues().forEach(function (r, i) {
      if (String(r[1]).trim()) weekRows.push([weekMon, WEEK_MUST_, i + 1, r[0] === true, String(r[1]).trim()]);
    });
    [W.GOAL_ROW, W.REWARD_ROW, W.AFFIRM_ROW].forEach(function (row, i) {
      const v = String(sh.getRange(row, W.INFO_VALUE_COL).getValue()).trim();
      if (v) weekRows.push([weekMon, WEEK_INFO_[i], 0, false, v]);
    });
  }

  // 3) "넘어온 일"로 빠져 있던 일 → 그날 할 일 맨 뒤로
  const saved = JSON.parse(props_().getProperty('CARRY_OVERFLOW') || '{}');
  if (saved.date && (saved.items || []).length) {
    const date = dateOnly_(saved.date);
    let n = todoRows.filter(function (r) { return dateKey_(r[0]) === saved.date; }).length;
    saved.items.forEach(function (it) {
      const t = cleanTodo_({ start: splitRange_(it.time).start, end: splitRange_(it.time).end, text: it.text, memo: it.memo });
      todoRows.push([date, ++n, false, t.start, t.end, t.text, t.memo, CARRY_FLAG_.CARRIED]);
    });
  }

  // 날짜별로 화면 정렬 규칙대로 순서 다시 매김 (내용은 그대로)
  const byDate = {};
  todoRows.forEach(function (r) { (byDate[dateKey_(r[0])] = byDate[dateKey_(r[0])] || []).push(r); });
  todoRows.length = 0;
  Object.keys(byDate).sort().forEach(function (k) {
    const items = byDate[k].map(function (r) { return { done: r[2], start: r[3], end: r[4], text: r[5], memo: r[6], flag: r[7], date: r[0] }; });
    sortTodoItems_(items).forEach(function (t, i) { todoRows.push([t.date, i + 1, t.done, t.start, t.end, t.text, t.memo, t.flag]); });
  });

  // 저장 (저장소가 비어 있을 때만 — 두 번 옮기지 않게)
  const write = function (name, cols, rows) {
    const s = sheet_(name);
    if (s.getLastRow() > 1 || !rows.length) return 0;
    ensureSize_(s, rows.length + 5, cols);
    s.getRange(2, 1, rows.length, cols).setValues(rows);
    s.getRange(2, 1, rows.length, cols).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }, { column: 3, ascending: true }]);
    return rows.length;
  };
  const n1 = write(CONFIG.SHEETS.TODO_STORE, TODO_STORE_COLS_, todoRows);
  const n2 = write(CONFIG.SHEETS.DAY_STORE, DAY_STORE_COLS_, dayRows);
  const n3 = write(CONFIG.SHEETS.WEEK_STORE, WEEK_STORE_COLS_, weekRows);
  STORE_MEMO_ = {};
  props_().setProperty('CARRY_OVERFLOW', '{}');
  const report = { todos: n1, day: n2, week: n3, backup: backup };
  props_().setProperty('STORE_REPORT', JSON.stringify(report));
  setStatus_('store', '✓ 옮김: 할 일 ' + n1 + '줄 · 습관·기분 ' + n2 + '줄 · 주간 ' + n3 + '줄 · 백업 ' + backup);
  return report;
}
