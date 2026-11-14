/**
 * archive.js
 * 기록 시트: 분석용 "한 줄에 하나" 형태로 쌓는 저장소
 * 열: 날짜 | 구분 | 항목 | 완료여부 | 값 | 주차
 * - 같은 날짜+구분 묶음은 지우고 다시 써서 중복 없이 덮어씀
 * - 행 삭제·추가만 하고 시트를 통째로 지우는 일은 절대 없음
 */

/** 기록 시트 만들기 (이미 있으면 머리글·서식만 확인, 데이터는 그대로) */
function buildArchive_(theme) {
  const sh = ensureSheet_(CONFIG.SHEETS.ARCHIVE);
  const H = CONFIG.ARCHIVE_HEADERS;
  ensureSize_(sh, 50, H.length);
  if (!sh.getRange(1, 1).getValue()) sh.getRange(1, 1, 1, H.length).setValues([H]);
  styleHeader_(sh.getRange(1, 1, 1, H.length), theme);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('F:F').setNumberFormat('@'); // 주차는 글자로 (CSV에서 그대로)
  sh.setColumnWidth(1, 100).setColumnWidth(2, 80).setColumnWidth(3, 240)
    .setColumnWidth(4, 80).setColumnWidth(5, 60).setColumnWidth(6, 100);
  return sh;
}

/** 기록 전체 읽기 → [{row, key, cat, item, done, value, week}] */
function readArchive_() {
  const sh = sheet_(CONFIG.SHEETS.ARCHIVE);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 6).getValues().map(function (r, i) {
    const d = dateOnly_(r[0]);
    return { row: i + 2, key: d ? dateKey_(d) : String(r[0]), cat: r[1], item: r[2], done: r[3], value: r[4], week: String(r[5]) };
  });
}

/**
 * 기록 덮어쓰기
 * keys: [{date, cat}] 에 해당하는 기존 행을 지우고 rows를 추가
 */
function replaceArchive_(keys, rows) {
  const sh = sheet_(CONFIG.SHEETS.ARCHIVE);
  const want = {};
  keys.forEach(function (k) { want[dateKey_(k.date) + '|' + k.cat] = true; });
  const del = readArchive_().filter(function (r) { return want[r.key + '|' + r.cat]; })
    .map(function (r) { return r.row; });

  deleteRowList_(sh, del);
  if (rows.length) {
    const at = Math.max(sh.getLastRow(), 1) + 1;
    ensureSize_(sh, at + rows.length + 5, 6);
    sh.getRange(at, 1, rows.length, 6).setValues(rows);
  }
  // 날짜·구분 순으로 정리 (데이터는 그대로, 순서만)
  const last = sh.getLastRow();
  if (last > 2) sh.getRange(2, 1, last - 1, 6).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
}

/**
 * 저장소의 하루치(할 일·습관·기분·일정 체크) → 기록 행 (개수 제한 없음)
 * 할 일: "미완"(넘겨진 원래 일)은 FALSE + 값 "이월" / 일정: 구분 "일정", 시간표 수업: 구분 "수업"
 */
function buildDayRows_(date) {
  const C = CONFIG.CATEGORY, wk = weekKey_(date);
  const rows = [];
  getTodos_(date).forEach(function (t) {
    const missed = t.flag === CARRY_FLAG_.MISSED;
    rows.push([date, C.TODO, t.text, missed ? false : t.done, missed ? CONFIG.CARRY.VALUE : '', wk]);
  });
  const day = getDayData_(date);
  getHabits_().forEach(function (name) {
    if (name) rows.push([date, C.HABIT, name, !!day.habits[name], '', wk]);
  });
  if (day.mood.score || day.mood.memo) rows.push([date, C.MOOD, day.mood.memo, '', day.mood.score, wk]); // 기분: 항목=메모, 값=점수
  scheduleOf_(date, day).forEach(function (it) {
    rows.push([date, it.kind === 'tt' ? C.CLASS : C.SCHEDULE, it.text, it.done, '', wk]);
  });
  return rows;
}

/** 하루치 덮어쓰기 키 */
function dayKeys_(date) {
  const C = CONFIG.CATEGORY;
  return [C.TODO, C.HABIT, C.MOOD, C.SCHEDULE, C.CLASS].map(function (cat) { return { date: date, cat: cat }; });
}

/** 편집 즉시 그날 기록 갱신 (미래 날짜는 밤에 저장) */
function archiveDayLive_(date) {
  if (!date || date > today_()) return;
  replaceArchive_(dayKeys_(date), buildDayRows_(date));
}

/** 그 주 꼭 끝낼 일 → 주간목표 (날짜 = 그 주 월요일) */
function weekGoalRows_(monday) {
  const mon = mondayOf_(monday || today_());
  const rows = getWeekData_(mon).must.map(function (m) {
    return [mon, CONFIG.CATEGORY.WEEK_GOAL, m.text, m.done, '', dateKey_(mon)];
  });
  return { keys: [{ date: mon, cat: CONFIG.CATEGORY.WEEK_GOAL }], rows: rows };
}

/** 이번 달 목표 → 월간목표 (날짜 = 그 달 1일) */
function monthGoalRows_() {
  const sh = sheet_(CONFIG.SHEETS.MONTHLY), L = MONTHLY;
  const ym = getMonthlyYM_();
  if (!ym) return { keys: [], rows: [] };
  const first = new Date(ym.y, ym.m - 1, 1);
  const rows = [];
  sh.getRange(L.GOAL_START, L.PANEL_COL, CONFIG.MONTH_GOALS, 2).getValues().forEach(function (r) {
    const text = String(r[1]).trim();
    if (text) rows.push([first, CONFIG.CATEGORY.MONTH_GOAL, text, r[0] === true, '', weekKey_(first)]);
  });
  return { keys: [{ date: first, cat: CONFIG.CATEGORY.MONTH_GOAL }], rows: rows };
}

/** 주간목표만 저장 (monday 없으면 이번 주) */
function archiveWeekGoals_(monday) {
  const g = weekGoalRows_(monday);
  replaceArchive_(g.keys, g.rows);
}

/** 월간목표만 저장 */
function archiveMonthGoals_() {
  const g = monthGoalRows_();
  replaceArchive_(g.keys, g.rows);
}

/** from ~ to 하루치 저장 (오늘 이후는 건너뜀) → 저장한 행 수 */
function saveRange_(from, to) {
  let keys = [], rows = [];
  const today = today_();
  for (let d = from; d <= to && d <= today; d = addDays_(d, 1)) {
    keys = keys.concat(dayKeys_(d));
    rows = rows.concat(buildDayRows_(d));
  }
  replaceArchive_(keys, rows);
  return rows.length;
}

/** 이번 주 월요일 ~ 오늘 + 이번 주 꼭 끝낼 일 + 이번 달 목표 저장 */
function saveAll_() {
  const today = today_();
  let n = saveRange_(mondayOf_(today), today);
  [weekGoalRows_(today), monthGoalRows_()].forEach(function (g) {
    replaceArchive_(g.keys, g.rows);
    n += g.rows.length;
  });
  setStatus_('save', '✓ ' + n + '행');
  return n;
}

/** 특정 주차의 기록 행 수 (새 주 시작 전 확인용) */
function countArchiveWeek_(week) {
  return readArchive_().filter(function (r) { return r.week === week; }).length;
}
