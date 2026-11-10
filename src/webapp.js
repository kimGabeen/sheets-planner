/**
 * webapp.js
 * 웹앱(HtmlService) 진입점과 화면에서 부르는 API
 * - 구글 시트는 그대로 저장소: 웹앱 변경은 위클리 시트에 쓰고 기존 함수(정렬·오늘 반영·기록)를 재사용
 * - google.script.run은 이름이 _로 끝나는 함수를 못 부르므로 API는 web으로 시작하는 이름을 씀
 */

/** 웹앱 화면 */
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle(CONFIG.WEB.TITLE)
    .setFaviconUrl(CONFIG.WEB.ICON_URL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('mobile-web-app-capable', 'yes');
}

/** HTML 템플릿에서 다른 파일 끼워 넣기 */
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ---------- 읽기 ---------- */

/**
 * 처음 열 때 데이터 (날짜가 바뀌었으면 넘기기·새 주 처리 먼저)
 * 처음 화면(오늘·위클리)에 필요한 것만 담고, 먼슬리·타임라인은 그 탭을 열 때 따로 받음
 * 서버 캐시에 있으면 시트를 다시 읽지 않음 (시트가 바뀌면 캐시 무효화)
 */
function webBoot() {
  return withLock_(function () {
    dailyRollover_();
    return cached_('boot|' + dateKey_(today_()), bootData_);
  }, true);
}

/** 캘린더 새로고침 후 전체 데이터 */
function webRefresh() {
  return withLock_(function () {
    syncCalendar_();
    dailyRollover_();
    return bootData_();
  });
}

/** 화면에 필요한 데이터 묶음 */
function bootData_() {
  const today = today_();
  const saved = JSON.parse(props_().getProperty('CARRY_OVERFLOW') || '{}');
  return {
    todayKey: dateKey_(today),
    todayIndex: weeklyIndexOf_(today),
    week: webWeekData_(),
    habits: getHabits_(),
    categories: getCategories_(),
    moods: CONFIG.MOOD_EMOJI,
    overflow: saved.date === dateKey_(today) ? saved.items || [] : [],
    timetable: timetableData_(),
    savedAt: stamp_(),
  };
}

/** 위클리 시트를 한 번에 읽어 7일치 + 목표·꼭 끝낼 일로 정리 */
function webWeekData_() {
  const L = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY);
  const mon = weeklyMonday_();
  const cols = Math.min(sh.getMaxColumns(), L.LAST_COL); // setup 전이라 메모 열이 없어도 동작
  const range = sh.getRange(1, 1, L.MOOD_MEMO, cols);
  const vals = range.getValues(), disp = range.getDisplayValues();
  const v = function (r, c) { return c <= cols ? vals[r - 1][c - 1] : ''; };
  const dv = function (r, c) { return String(disp[r - 1][c - 1]).trim(); };

  const byDate = scheduleByDate_(readEvents_());
  const habits = getHabits_();
  const days = [];
  for (let d = 0; d < 7 && mon; d++) {
    const date = addDays_(mon, d), key = dateKey_(date), c = L.checkCol(d);
    const todos = [];
    for (let i = 0; i < CONFIG.TODO_ROWS; i++) {
      const r = L.TODO_START + i;
      todos.push([v(r, c) === true, dv(r, L.timeCol(d)), String(v(r, L.textCol(d))).trim(), String(v(r, L.memoCol(d)))]);
    }
    days.push({
      key: key,
      todos: todos,
      habits: habits.map(function (n, i) { return v(L.HABIT_START + i, L.HABIT_DAY_COL + d) === true; }),
      mood: { score: moodScore_(v(L.MOOD, c)), memo: String(v(L.MOOD_MEMO, c)).trim() },
      schedule: (byDate[key] || {}).items || [],
      holiday: (byDate[key] || {}).holiday || '',
    });
  }
  const must = [];
  for (let i = 0; i < CONFIG.MUST_COUNT; i++) {
    must.push([v(L.MUST_START + i, 1) === true, String(v(L.MUST_START + i, 2)).trim()]);
  }
  return {
    monday: mon ? dateKey_(mon) : '',
    days: days,
    info: [L.GOAL_ROW, L.REWARD_ROW, L.AFFIRM_ROW].map(function (r) { return String(v(r, L.INFO_VALUE_COL)).trim(); }),
    must: must,
  };
}

/**
 * _일정 행 → 날짜별 {items, holiday}
 * items: 화면에 보일 일정 (캘린더 cal / 내 일정 memo / 시간표 tt). 공휴일은 holiday 이름으로만
 * 캘린더에도 있는 메모(연동)는 캘린더 쪽만 보이고, 수정하면 둘 다 바뀜
 * ref = 수정·삭제 때 그 줄을 다시 찾는 열쇠 (종류 + 그날 같은 종류 중 순서 + 글자)
 */
function scheduleByDate_(rows) {
  const out = {}, count = {};
  rows.forEach(function (r) {
    const key = dateKey_(r[0]);
    const entry = out[key] = out[key] || { items: [], holiday: '' };
    if (r[1] === CONFIG.EVENT_TYPE.HOLIDAY) {
      entry.holiday = entry.holiday ? entry.holiday + ', ' + r[2] : r[2];
      return;
    }
    const kind = eventKind_(r[1]);
    const ck = key + '|' + kind;
    const index = count[ck] = count[ck] === undefined ? 0 : count[ck] + 1;
    if (kind === 'memo' && r[5] === '연동') return;
    const p = parseMemoText_(r[2]);
    entry.items.push({ kind: kind, text: r[2], title: p.title, start: p.start, end: p.end, memo: String(r[6] || ''),
      ref: { kind: kind, index: index, text: r[2] } });
  });
  return out;
}

/**
 * date가 든 달(6주 달력) 기준 _일정 행. 동기화 기간 밖이면 그 기간 캘린더 일정을 바로 가져와 합침 (시트에는 안 씀)
 * 먼슬리 화면과 수정·삭제가 같은 목록을 보게 해서 ref 순서가 어긋나지 않게 함
 */
function eventRowsAround_(date) {
  const rows = readEvents_();
  const range = syncRange_();
  const gs = mondayOf_(new Date(date.getFullYear(), date.getMonth(), 1)), ge = addDays_(gs, 42);
  if (gs >= range.start && ge <= range.end) return rows;
  const extra = calendarRows_({ start: gs, end: ge });
  const inGrid = function (d) { return d >= gs && d < ge; };
  return sortEvents_(rows.filter(function (r) { return r[1] !== CONFIG.EVENT_TYPE.CAL || !inGrid(r[0]); })
    .map(function (r) {
      if (r[1] !== CONFIG.EVENT_TYPE.MEMO || !r[4] || !inGrid(r[0])) return r;
      const copy = r.slice();
      copy[5] = extra.ids[r[4]] ? '연동' : '';
      return copy;
    }).concat(extra.rows));
}

/** _일정 구분 → 화면용 종류 */
function eventKind_(type) {
  const T = CONFIG.EVENT_TYPE;
  return type === T.MEMO ? 'memo' : type === T.TIMETABLE ? 'tt' : 'cal';
}

/** 메모 글자 → {start, end, title} ('10:00-11:30 수업' 형식) */
function parseMemoText_(text) {
  const m = String(text).match(/^(\d{2}:\d{2})(?:-(\d{2}:\d{2}))?\s+(.+)$/);
  return m ? { start: m[1], end: m[2] || '', title: m[3] } : { start: '', end: '', title: String(text) };
}

/** 고정 시간표를 화면용으로 (날짜는 'yyyy-MM-dd') */
function timetableData_() {
  const tt = getTimetable_();
  return { start: tt.start ? dateKey_(tt.start) : '', end: tt.end ? dateKey_(tt.end) : '', rows: tt.rows };
}

/** 먼슬리 화면 데이터 (달력·월간 목표·습관 달성률·기분 그래프) */
function webGetMonth(y, m) {
  return withLock_(function () {
    return cached_('month|' + y + '-' + m + '|' + dateKey_(today_()), function () { return monthData_(y, m); });
  }, true);
}

/**
 * 먼슬리 데이터 (잠금 없이, 다른 API 안에서도 사용)
 * 날짜별 items: 캘린더 일정(cal, 읽기 전용) + 내 일정(memo, 수정·삭제 가능). 시간표는 달력에 넣지 않음
 */
function monthData_(y, m) {
  const first = new Date(y, m - 1, 1), days = new Date(y, m, 0).getDate();
  const firstKey = dateKey_(first), prefix = firstKey.slice(0, 8); // 'yyyy-MM-'
  const today = today_();
  const elapsed = today < first ? 0 : Math.min(days, diffDays_(first, today) + 1);

  const dates = [];
  for (let n = 1; n <= days; n++) dates.push({ day: n, items: [], holiday: '', mood: '', todo: [0, 0] });
  const at = function (key) { return key.indexOf(prefix) === 0 ? dates[+key.slice(8) - 1] : null; };

  const byDate = scheduleByDate_(eventRowsAround_(first));
  dates.forEach(function (c) {
    const e = byDate[prefix + ('0' + c.day).slice(-2)] || {};
    c.items = (e.items || []).filter(function (it) { return it.kind !== 'tt'; }); // 달력에는 시간표 안 넣음
    c.holiday = e.holiday || '';
  });

  const habits = getHabits_().filter(String);
  const habitDone = {};
  const goals = [];
  readArchive_().forEach(function (r) {
    const C = CONFIG.CATEGORY;
    if (r.cat === C.MONTH_GOAL && r.key === firstKey) { goals.push([r.done === true, String(r.item)]); return; }
    const cell = at(r.key);
    if (!cell) return;
    if (r.cat === C.MOOD) cell.mood = Number(r.value) || '';
    else if (r.cat === C.TODO) { cell.todo[1]++; if (r.done === true) cell.todo[0]++; }
    else if (r.cat === C.HABIT && r.done === true) habitDone[r.item] = (habitDone[r.item] || 0) + 1;
  });
  while (goals.length < CONFIG.MONTH_GOALS) goals.push([false, '']);

  return {
    y: y, m: m, days: days,
    offset: (first.getDay() + 6) % 7, // 1일 앞 빈칸 수 (월요일 시작)
    todayKey: dateKey_(today),
    dates: dates,
    goals: goals.slice(0, CONFIG.MONTH_GOALS),
    habits: habits.map(function (n) {
      return { name: n, done: habitDone[n] || 0, rate: elapsed ? (habitDone[n] || 0) / elapsed : 0 };
    }),
    elapsed: elapsed,
  };
}

/** 타임라인 하루치 */
function webGetTimeline(key) {
  return withLock_(function () {
    return { key: key, slots: readTimeline_(dateOnly_(key)) };
  }, true);
}

/* ---------- 쓰기 (화면은 먼저 바뀌고, 여기서 뒤에서 저장) ---------- */

/** 날짜 → 위클리 요일 번호 (이번 주가 아니면 오류) */
function webDayIndex_(key) {
  const d = weeklyIndexOf_(dateOnly_(key));
  if (d < 0) throw new Error('이번 주 날짜만 바꿀 수 있어요');
  return d;
}

/** 위클리의 그 요일이 오늘이면 오늘 시트에도 반영 */
function mirrorToToday_(d) {
  if (weeklyIndexOf_(todaySheetDate_()) !== d) return;
  copyBlocks_(sheet_(CONFIG.SHEETS.WEEKLY), sheet_(CONFIG.SHEETS.TODAY), todayBlocks_(d), 'w');
}

/** 할 일 10줄 저장 [체크, 시간, 내용, 메모] → 정렬된 결과 반환 (메모는 같은 줄 숨김 칸) */
function webSetTodos(key, rows) {
  return withLock_(function () {
    const L = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY), d = webDayIndex_(key);
    const clean = [];
    for (let i = 0; i < CONFIG.TODO_ROWS; i++) {
      const r = rows[i] || [];
      const text = String(r[2] || '').trim();
      clean.push([r[0] === true, String(r[1] || '').trim(), text, text ? String(r[3] || '').trim() : '']);
    }
    sh.getRange(L.TODO_START, L.checkCol(d), CONFIG.TODO_ROWS, 3).setValues(clean.map(function (r) { return r.slice(0, 3); }));
    memoRange_(sh, L.TODO_START, L.memoCol(d)).setValues(clean.map(function (r) { return [r[3]]; }));
    sortTodos_(sh, L.TODO_START, L.checkCol(d), L.memoCol(d));
    mirrorToToday_(d);
    archiveDayLive_(dateOnly_(key));
    const range = sh.getRange(L.TODO_START, L.checkCol(d), CONFIG.TODO_ROWS, 3);
    const vals = range.getValues(), disp = range.getDisplayValues();
    const memos = memoRange_(sh, L.TODO_START, L.memoCol(d)).getValues();
    return vals.map(function (r, i) { return [r[0] === true, String(disp[i][1]).trim(), String(r[2]).trim(), String(memos[i][0])]; });
  });
}

/** 습관 체크 */
function webSetHabit(key, index, done) {
  return withLock_(function () {
    const L = WEEKLY, d = webDayIndex_(key);
    sheet_(CONFIG.SHEETS.WEEKLY).getRange(L.HABIT_START + index, L.HABIT_DAY_COL + d).setValue(done === true);
    mirrorToToday_(d);
    archiveDayLive_(dateOnly_(key));
    return true;
  });
}

/** 기분 점수(1~5 또는 '') + 한 줄 메모 */
function webSetMood(key, score, memo) {
  return withLock_(function () {
    const L = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY), d = webDayIndex_(key);
    sh.getRange(L.MOOD, L.checkCol(d)).setValue(score ? CONFIG.MOODS[score - 1] : '');
    sh.getRange(L.MOOD_MEMO, L.checkCol(d)).setValue(String(memo || '').trim());
    mirrorToToday_(d);
    archiveDayLive_(dateOnly_(key));
    return true;
  });
}

/** 이번 주 꼭 끝낼 일 5개 [[체크, 내용]] */
function webSetMust(rows) {
  return withLock_(function () {
    const L = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY);
    const clean = [];
    for (let i = 0; i < CONFIG.MUST_COUNT; i++) {
      const r = rows[i] || [];
      clean.push([r[0] === true, String(r[1] || '').trim()]);
    }
    sh.getRange(L.MUST_START, 1, CONFIG.MUST_COUNT, 2).setValues(clean);
    const td = weeklyIndexOf_(todaySheetDate_());
    if (td >= 0) mirrorToToday_(td);
    archiveWeekGoals_();
    return true;
  });
}

/** 이번 주 목표·보상·확언 */
function webSetWeekInfo(info) {
  return withLock_(function () {
    const L = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY);
    [L.GOAL_ROW, L.REWARD_ROW, L.AFFIRM_ROW].forEach(function (r, i) {
      sh.getRange(r, L.INFO_VALUE_COL).setValue(String(info[i] || '').trim());
    });
    return true;
  });
}

/** 월간 목표 5개 저장 (먼슬리 시트가 같은 달을 보고 있으면 칸도 갱신) */
function webSetMonthGoals(y, m, rows) {
  return withLock_(function () {
    const first = new Date(y, m - 1, 1), L = MONTHLY;
    const clean = [];
    for (let i = 0; i < CONFIG.MONTH_GOALS; i++) {
      const r = rows[i] || [];
      clean.push([r[0] === true, String(r[1] || '').trim()]);
    }
    const ym = getMonthlyYM_();
    if (ym && ym.y === y && ym.m === m) {
      const sh = sheet_(CONFIG.SHEETS.MONTHLY);
      sh.getRange(L.GOAL_START, L.PANEL_COL, CONFIG.MONTH_GOALS, 1).setValues(clean.map(function (r) { return [r[0]]; }));
      sh.getRange(L.GOAL_START, L.PANEL_COL + 1, CONFIG.MONTH_GOALS, 1).setValues(clean.map(function (r) { return [r[1]]; }));
    }
    const out = clean.filter(function (r) { return r[1]; }).map(function (r) {
      return [first, CONFIG.CATEGORY.MONTH_GOAL, r[1], r[0], '', weekKey_(first)];
    });
    replaceArchive_([{ date: first, cat: CONFIG.CATEGORY.MONTH_GOAL }], out);
    return true;
  });
}

/** 새 일정 추가 → {ok, message, schedule(그날 일정)} */
function webAddEvent(input) {
  return withLock_(function () {
    let res;
    try {
      res = createCalendarEvent_(input.date, input.start, input.end, input.title);
    } catch (err) {
      res = { ok: false, message: '오류: ' + err.message };
    }
    res.week = webWeekData_();
    return res;
  });
}

/** 타임라인 하루치 저장 (144칸 배열) */
function webSetTimeline(key, slots) {
  return withLock_(function () {
    const names = {};
    getCategories_().forEach(function (c) { names[c.name] = true; });
    const clean = [];
    for (let i = 0; i < slotCount_(); i++) clean.push(names[slots[i]] ? slots[i] : '');
    writeTimeline_(dateOnly_(key), clean);
    return true;
  });
}

/* ---------- 먼슬리 일정(메모) · 시간표 · 습관 편집 ---------- */

/** 그날 ref에 해당하는 _일정 줄 찾기 (다른 기기에서 바뀌었으면 오류) */
function findItemRow_(rows, date, ref) {
  const key = dateKey_(date);
  const same = rows.filter(function (r) { return dateKey_(r[0]) === key && eventKind_(r[1]) === ref.kind && r[1] !== CONFIG.EVENT_TYPE.HOLIDAY; });
  const row = same[ref.index];
  if (!row || row[2] !== ref.text) throw new Error('다른 곳에서 바뀌었어요. 새로고침 후 다시 해 주세요');
  return row;
}

/** 입력 → 일정 글자 ('11:00-13:00 제목' / '15:00 제목' / '제목') */
function itemText_(shape) {
  const hhmm = function (t) { return ('0' + t.h).slice(-2) + ':' + ('0' + t.m).slice(-2); };
  return (shape.st ? hhmm(shape.st) + (shape.en ? '-' + hhmm(shape.en) : '') + ' ' : '') + shape.title;
}

/** 결과에 최신 위클리(+ 먼슬리) 붙이기 → 화면이 바로 갱신 */
function withViews_(res, y, m) {
  res.week = webWeekData_();
  res.month = y ? monthData_(y, m) : null;
  return res;
}

/** 캘린더 ID로 연동된 메모 줄들 고치기 (text가 null이면 삭제), memo = 숨김 메모 */
function updateLinkedMemos_(id, text, memo) {
  if (!id) return;
  const rows = readEvents_();
  let changed = false;
  const out = rows.filter(function (r) {
    if (r[1] !== CONFIG.EVENT_TYPE.MEMO || r[4] !== id) return true;
    changed = true;
    if (text === null) return false;
    r[2] = text; r[3] = eventSortKey_(text); r[6] = memo || '';
    return true;
  });
  if (changed) { writeEvents_(out); refreshMonthMemos_(); }
}

/**
 * 일정 저장 (오늘·요일 상세·먼슬리 공용)
 * ref = null → 새 일정: 항상 구글 캘린더에 만들고 먼슬리 메모에도 연동해 저장
 * ref.kind = 'cal' → 구글 캘린더 일정을 ID(없으면 날짜+제목)로 찾아 수정 + 연동 메모도 수정
 * ref.kind = 'memo' → 내 일정 수정 (캘린더에도 있으면 같이 수정)
 * 반환 {ok, message, week, month}
 */
function webSaveItem(key, ref, item, y, m) {
  return withLock_(function () {
    const date = dateOnly_(key);
    let res;
    try {
      const shape = eventShape_(date, item.start, item.end, item.title);
      if (shape.error) return withViews_({ ok: false, message: shape.error }, y, m);
      const text = itemText_(shape);
      const memo = String(item.memo || '').trim(); // 숨김 메모 (목록엔 안 보임)

      if (!ref) {
        const cal = createCalendarEvent_(date, item.start, item.end, item.title, memo); // 안에서 캘린더 다시 불러옴
        if (!cal.ok) return withViews_({ ok: false, message: cal.message }, y, m);
        const rows = readEvents_();
        rows.push([date, CONFIG.EVENT_TYPE.MEMO, text, eventSortKey_(text), cal.id, '연동', memo]);
        writeEvents_(rows);
        refreshMonthMemos_();
        res = { ok: true, message: '저장됨 ✓ · 구글 캘린더에도 추가' };
      } else if (ref.kind === 'cal') {
        const row = findItemRow_(eventRowsAround_(date), date, ref);
        const up = updateCalendarEvent_(row[4], date, item.start, item.end, item.title, calendarTitles_(row[2]), memo);
        if (!up.ok) throw new Error(up.found ? up.message : '구글 캘린더에서 찾지 못했어요 (이미 지워졌을 수 있어요)');
        updateLinkedMemos_(row[4], text, memo);
        res = { ok: true, message: '수정됨 ✓ · 구글 캘린더에도 반영' };
      } else if (ref.kind === 'memo') {
        let id = findItemRow_(readEvents_(), date, ref)[4], msg = '수정됨 ✓';
        if (id) {
          const up = updateCalendarEvent_(id, date, item.start, item.end, item.title, [], memo);
          if (up.ok) msg += ' · 구글 캘린더에도 반영';
          else if (!up.found) { id = ''; msg += ' · 구글 캘린더에서 이미 지워진 일정이라 앱에만 저장했어요'; }
          else throw new Error(up.message);
        }
        const rows = readEvents_();
        const old = findItemRow_(rows, date, ref);
        old[2] = text; old[3] = eventSortKey_(text); old[4] = id; old[5] = id ? '연동' : ''; old[6] = memo;
        writeEvents_(rows);
        refreshMonthMemos_();
        res = { ok: true, message: msg };
      } else {
        throw new Error('시간표·공휴일은 여기서 바꿀 수 없어요');
      }
    } catch (err) {
      res = { ok: false, message: err.message };
    }
    return withViews_(res, y, m);
  });
}

/** 일정 삭제 (캘린더 일정이면 구글 캘린더에서도 삭제) → {ok, message, week, month} */
function webDeleteItem(key, ref, y, m) {
  return withLock_(function () {
    const date = dateOnly_(key);
    let res;
    try {
      const row = findItemRow_(ref.kind === 'cal' ? eventRowsAround_(date) : readEvents_(), date, ref);
      if (ref.kind === 'cal') {
        if (!deleteCalendarEvent_(row[4], date, calendarTitles_(row[2]))) {
          throw new Error('구글 캘린더에서 찾지 못했어요 (이미 지워졌을 수 있어요)');
        }
        updateLinkedMemos_(row[4], null);
        res = { ok: true, message: '삭제됨 · 구글 캘린더에서도 삭제' };
      } else if (ref.kind === 'memo') {
        const gone = row[4] ? deleteCalendarEvent_(row[4], date, []) : false;
        const rows = readEvents_();
        const old = findItemRow_(rows, date, ref);
        writeEvents_(rows.filter(function (r) { return r !== old; }));
        refreshMonthMemos_();
        res = { ok: true, message: '삭제됨' + (gone ? ' · 구글 캘린더에서도 삭제' : '') };
      } else {
        throw new Error('시간표·공휴일은 여기서 지울 수 없어요');
      }
    } catch (err) {
      res = { ok: false, message: err.message };
    }
    return withViews_(res, y, m);
  });
}

/** 고정 시간표 저장 → 일정 칸 다시 만들기 → {timetable, week} */
function webSetTimetable(tt) {
  return withLock_(function () {
    const rows = (tt.rows || []).map(function (r) {
      const st = normTime_(r.start), en = normTime_(r.end), name = String(r.name || '').trim();
      if (!name && !st) return null;
      if (!name) throw new Error('이름을 입력해 주세요');
      if (!st) throw new Error(name + ': 시작시간 확인');
      if (en && en <= st) throw new Error(name + ': 끝시간 확인');
      return { day: Number(r.day), start: st, end: en, name: name };
    }).filter(Boolean);
    const start = dateOnly_(tt.start), end = dateOnly_(tt.end);
    if (start && end && end < start) throw new Error('학기 종료가 시작보다 빨라요');
    setTimetable_({ start: start, end: end, rows: rows });
    refreshTimetableEvents_();
    return { timetable: timetableData_(), week: webWeekData_() };
  });
}

/** 습관 목록 저장 (추가·삭제·이름 변경·순서) → 전체 데이터 */
function webSetHabits(list) {
  return withLock_(function () {
    const clean = (list || []).map(function (h) {
      return { name: String(h.name || '').trim(), from: h.from === null || h.from === undefined ? null : Number(h.from) };
    }).filter(function (h) { return h.name; });
    setHabits_(clean);
    const td = weeklyIndexOf_(todaySheetDate_());
    if (td >= 0) mirrorToToday_(td);
    archiveDayLive_(today_()); // 오늘만 새 이름으로 다시 저장, 지난 기록은 그대로
    return bootData_();
  });
}
