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

/** 처음 열 때 전체 데이터 (날짜가 바뀌었으면 넘기기·새 주 처리 먼저) */
function webBoot() {
  return withLock_(function () {
    dailyRollover_();
    return bootData_();
  });
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
    savedAt: stamp_(),
  };
}

/** 위클리 시트를 한 번에 읽어 7일치 + 목표·꼭 끝낼 일로 정리 */
function webWeekData_() {
  const L = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY);
  const mon = weeklyMonday_();
  const range = sh.getRange(1, 1, L.MOOD_MEMO, L.HABIT_PCT_COL);
  const vals = range.getValues(), disp = range.getDisplayValues();
  const v = function (r, c) { return vals[r - 1][c - 1]; };
  const dv = function (r, c) { return String(disp[r - 1][c - 1]).trim(); };

  const schedule = {};
  readEvents_().forEach(function (r) {
    const k = dateKey_(r[0]);
    (schedule[k] = schedule[k] || []).push({ text: r[2], memo: r[1] === CONFIG.EVENT_TYPE.MEMO });
  });
  const habits = getHabits_();
  const days = [];
  for (let d = 0; d < 7 && mon; d++) {
    const date = addDays_(mon, d), key = dateKey_(date), c = L.checkCol(d);
    const todos = [];
    for (let i = 0; i < CONFIG.TODO_ROWS; i++) {
      const r = L.TODO_START + i;
      todos.push([v(r, c) === true, dv(r, L.timeCol(d)), String(v(r, L.textCol(d))).trim()]);
    }
    days.push({
      key: key,
      todos: todos,
      habits: habits.map(function (n, i) { return v(L.HABIT_START + i, L.HABIT_DAY_COL + d) === true; }),
      mood: { score: moodScore_(v(L.MOOD, c)), memo: String(v(L.MOOD_MEMO, c)).trim() },
      schedule: schedule[key] || [],
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

/** 먼슬리 화면 데이터 (달력·월간 목표·습관 달성률·기분 그래프) */
function webGetMonth(y, m) {
  return withLock_(function () {
    const first = new Date(y, m - 1, 1), days = new Date(y, m, 0).getDate();
    const firstKey = dateKey_(first), prefix = firstKey.slice(0, 8); // 'yyyy-MM-'
    const today = today_();
    const elapsed = today < first ? 0 : Math.min(days, diffDays_(first, today) + 1);

    const dates = [];
    for (let n = 1; n <= days; n++) dates.push({ day: n, events: [], memo: '', mood: '', todo: [0, 0] });
    const at = function (key) { return key.indexOf(prefix) === 0 ? dates[+key.slice(8) - 1] : null; };

    readEvents_().forEach(function (r) {
      const cell = at(dateKey_(r[0]));
      if (!cell) return;
      if (r[1] === CONFIG.EVENT_TYPE.MEMO) cell.memo = r[2];
      else cell.events.push(r[2]);
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
  });
}

/** 타임라인 하루치 */
function webGetTimeline(key) {
  return withLock_(function () {
    return { key: key, slots: readTimeline_(dateOnly_(key)) };
  });
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

/** 할 일 10줄 저장 → 정렬된 결과 반환 */
function webSetTodos(key, rows) {
  return withLock_(function () {
    const L = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY), d = webDayIndex_(key);
    const clean = [];
    for (let i = 0; i < CONFIG.TODO_ROWS; i++) {
      const r = rows[i] || [];
      clean.push([r[0] === true, String(r[1] || '').trim(), String(r[2] || '').trim()]);
    }
    sh.getRange(L.TODO_START, L.checkCol(d), CONFIG.TODO_ROWS, 3).setValues(clean);
    sortTodos_(sh, L.TODO_START, L.checkCol(d));
    mirrorToToday_(d);
    archiveDayLive_(dateOnly_(key));
    const range = sh.getRange(L.TODO_START, L.checkCol(d), CONFIG.TODO_ROWS, 3);
    const vals = range.getValues(), disp = range.getDisplayValues();
    return vals.map(function (r, i) { return [r[0] === true, String(disp[i][1]).trim(), String(r[2]).trim()]; });
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
