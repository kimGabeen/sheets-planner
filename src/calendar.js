/**
 * calendar.js
 * 구글 캘린더 연동 (기본 캘린더만 사용, ID를 코드에 적지 않음)
 * - 불러오기: 일정을 숨김 시트 _일정에 저장 → 먼슬리·위클리·오늘 수식이 읽어 감
 * - 적기: "새 일정 추가" 입력칸 → 캘린더에 생성 (중복 방지)
 */

/** 내 기본 캘린더 */
function calendar_() {
  return CalendarApp.getDefaultCalendar();
}

/* ---------- _일정 시트 (캘린더 일정 + 달력 메모 + 고정 시간표) ---------- */
// 열: 날짜 | 구분(캘린더/메모/시간표/공휴일) | 내용 | 정렬(분+1, 시간 없음 0) | 이벤트ID | 연동 | 설명(숨김 메모)
// 메모는 한 줄에 일정 하나. 캘린더에도 만든 메모는 이벤트ID를 갖고 "연동" 표시 → 일정 칸에 두 번 안 나옴

const EVENT_COLS_ = 7;

/** _일정 시트 만들기 (숨김) */
function buildEventsSheet_() {
  const sh = ensureSheet_(CONFIG.SHEETS.EVENTS);
  ensureSize_(sh, 50, EVENT_COLS_);
  sh.getRange(1, 1, 1, EVENT_COLS_).setValues([['날짜', '구분', '내용', '정렬', '이벤트ID', '연동', '설명']]);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.hideSheet();
  return sh;
}

/** _일정 전체 읽기 → [[날짜, 구분, 내용, 정렬, ID, 연동, 설명], ...] */
function readEvents_() {
  const sh = sheet_(CONFIG.SHEETS.EVENTS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const cols = Math.min(EVENT_COLS_, sh.getMaxColumns());
  return sh.getRange(2, 1, last - 1, cols).getValues()
    .map(function (r) { return [dateOnly_(r[0]), r[1], r[2], Number(r[3]) || 0, r[4] || '', r[5] || '', r[6] || '']; })
    .filter(function (r) { return r[0]; });
}

/** _일정 전체 다시 쓰기 (날짜 → 시간 → 캘린더·시간표·메모 순 정렬) */
function writeEvents_(rows) {
  const sh = sheet_(CONFIG.SHEETS.EVENTS);
  sortEvents_(rows);
  ensureSize_(sh, rows.length + 2, EVENT_COLS_);
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, EVENT_COLS_).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, EVENT_COLS_).setValues(rows);
}

/** _일정 행 정렬 (날짜 → 시간 → 공휴일·캘린더·시간표·메모 순), 같으면 원래 순서 유지 */
function sortEvents_(rows) {
  const T = CONFIG.EVENT_TYPE;
  const order = function (t) { return t === T.HOLIDAY ? -1 : t === T.CAL ? 0 : t === T.TIMETABLE ? 1 : 2; };
  rows.forEach(function (r, i) { r._i = i; });
  rows.sort(function (a, b) {
    return (a[0] - b[0]) || (a[3] - b[3]) || (order(a[1]) - order(b[1])) || (a._i - b._i);
  });
  rows.forEach(function (r) { delete r._i; });
  return rows;
}

/** 일정 글자 앞 시간으로 정렬 값 ('15:00 스터디' → 901, 시간 없음 → 0) */
function eventSortKey_(text) {
  const m = String(text).match(/^(\d{1,2}):(\d{2})/);
  return m ? +m[1] * 60 + +m[2] + 1 : 0;
}

/**
 * 먼슬리 메모 칸에서 숨길 메모인지 (캘린더 일정 줄에 이미 보이는 것)
 * 캘린더와 연동된 메모 또는 같은 날짜에 같은 글자의 캘린더 일정이 있는 메모
 */
function hiddenMemoTest_(rows) {
  const cal = {};
  rows.forEach(function (r) { if (r[1] === CONFIG.EVENT_TYPE.CAL) cal[dateKey_(r[0]) + '|' + r[2]] = true; });
  return function (r) { return r[1] === CONFIG.EVENT_TYPE.MEMO && (r[5] === '연동' || !!cal[dateKey_(r[0]) + '|' + r[2]]); };
}

/** 그날 메모 줄들 (정렬 순서대로) */
function memosOf_(rows, date) {
  const key = dateKey_(date);
  return rows.filter(function (r) { return r[1] === CONFIG.EVENT_TYPE.MEMO && dateKey_(r[0]) === key; });
}

/**
 * 달력 메모 저장 (먼슬리 메모 칸 기준): 칸의 각 줄 = 일정 하나
 * 같은 글자의 예전 줄이 있으면 캘린더 연동 정보를 이어받음
 */
function saveMemos_(entries) {
  let rows = readEvents_();
  entries.forEach(function (e) {
    const old = memosOf_(rows, e.date);
    const key = dateKey_(e.date);
    // 메모 칸에 안 보이는(캘린더 쪽에 보이는) 메모는 그대로 두고, 보이는 메모만 칸 내용으로 바꿈
    const hidden = hiddenMemoTest_(rows);
    rows = rows.filter(function (r) { return !(r[1] === CONFIG.EVENT_TYPE.MEMO && dateKey_(r[0]) === key && !hidden(r)); });
    String(e.text || '').split(/\r?\n/).map(function (t) { return t.trim(); }).filter(String).forEach(function (t) {
      const prev = old.find(function (r) { return r[2] === t; });
      rows.push([e.date, CONFIG.EVENT_TYPE.MEMO, t, eventSortKey_(t), prev ? prev[4] : '', prev ? prev[5] : '', prev ? prev[6] : '']);
    });
  });
  writeEvents_(rows);
}

/** 고정 시간표 → 기간 안 날짜별 일정 행 (학기 기간에만) */
function timetableRows_(range) {
  const tt = getTimetable_();
  if (!tt.rows.length) return [];
  let from = range.start, to = range.end;
  if (tt.start && tt.start > from) from = tt.start;
  if (tt.end && addDays_(tt.end, 1) < to) to = addDays_(tt.end, 1);
  const out = [];
  for (let d = from; d < to; d = addDays_(d, 1)) {
    const wd = (d.getDay() + 6) % 7;
    tt.rows.forEach(function (c) {
      if (c.day !== wd) return;
      const text = c.start + (c.end ? '-' + c.end : '') + ' ' + c.name;
      out.push([d, CONFIG.EVENT_TYPE.TIMETABLE, text, eventSortKey_(text), '', '', '']);
    });
  }
  return out;
}

/**
 * 위클리·오늘 "일정" k번째 줄 수식 (캘린더 + 메모 + 시간표, 연동 메모·공휴일은 제외)
 * 마지막 줄은 넘치면 "+N개 더"
 */
function scheduleLineFormula_(dateRef, k, emptyText) {
  const ev = CONFIG.SHEETS.EVENTS, n = CONFIG.SCHEDULE_LINES;
  // 공휴일은 날짜 표시용이라 일정 목록에서 뺌
  // UNIQUE: 같은 날짜·시간·제목이면 한 줄만 (캘린더 일정과 같은 메모가 있어도 한 번만 보임)
  const list = 'UNIQUE(FILTER(' + colRef_(ev, 3) + ',' + colRef_(ev, 1) + '=' + dateRef + ',' + colRef_(ev, 6) + '<>"연동",' +
    colRef_(ev, 2) + '<>"' + CONFIG.EVENT_TYPE.HOLIDAY + '"))';
  const empty = '"' + emptyText + '"';
  if (k < n) return '=IFERROR(INDEX(' + list + ',' + k + '),' + empty + ')';
  return '=IFERROR(LET(l,' + list + ',c,ROWS(l),IF(c>' + n + ',"+"&(c-' + (n - 1) + ')&"개 더",INDEX(l,' + n + '))),' + empty + ')';
}

/** 시간표만 다시 만들기 (설정·웹앱에서 시간표를 바꿨을 때) */
function refreshTimetableEvents_() {
  const rows = readEvents_().filter(function (r) { return r[1] !== CONFIG.EVENT_TYPE.TIMETABLE; });
  writeEvents_(rows.concat(timetableRows_(syncRange_())));
}

/* ---------- 불러오기 ---------- */

/** 불러올 기간: 이번 달 ~ 다음 달의 6주 달력 전체 + 먼슬리에서 보는 달 + 위클리 주 */
function syncRange_() {
  const now = today_();
  let start = mondayOf_(new Date(now.getFullYear(), now.getMonth(), 1));
  let end = addDays_(mondayOf_(new Date(now.getFullYear(), now.getMonth() + CONFIG.CALENDAR_MONTHS_AHEAD, 1)), 42);
  const ym = getMonthlyYM_();
  if (ym) {
    const gs = monthGridStart_(ym), ge = addDays_(gs, MONTHLY.WEEKS * 7);
    if (gs < start) start = gs;
    if (ge > end) end = ge;
  }
  const mon = weeklyMonday_();
  if (mon) {
    if (mon < start) start = mon;
    if (addDays_(mon, 7) > end) end = addDays_(mon, 7);
  }
  return { start: start, end: end };
}

/**
 * 기본 캘린더에서 기간 안 일정 → _일정 행들 + 이벤트 ID 목록
 * 시간 일정은 '11:00-13:00 제목'(같은 날 끝나면), 종일 일정은 제목 그대로 (여러 날이면 날마다 한 줄)
 */
function calendarRows_(range) {
  const rows = [], ids = {};
  const inRange = function (d) { return d >= range.start && d < range.end; };
  const hm = function (d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, 'HH:mm'); };
  calendar_().getEvents(range.start, range.end).forEach(function (ev) {
    const title = ev.getTitle() || '(제목 없음)';
    const desc = String(ev.getDescription ? ev.getDescription() || '' : ''); // 캘린더 설명 = 메모
    ids[ev.getId()] = true;
    if (ev.isAllDayEvent()) {
      const last = dateOnly_(ev.getAllDayEndDate()); // 끝 날짜는 포함 안 됨
      for (let d = dateOnly_(ev.getAllDayStartDate()); d < last; d = addDays_(d, 1)) {
        if (inRange(d)) rows.push([d, CONFIG.EVENT_TYPE.CAL, title, eventSortKey_(title), ev.getId(), '', desc]);
      }
    } else {
      const s = ev.getStartTime(), e = ev.getEndTime(), d = dateOnly_(s);
      if (!inRange(d)) return;
      const sameDay = dateKey_(e) === dateKey_(s) && e > s;
      rows.push([d, CONFIG.EVENT_TYPE.CAL, hm(s) + (sameDay ? '-' + hm(e) : '') + ' ' + title,
        s.getHours() * 60 + s.getMinutes() + 1, ev.getId(), '', desc]);
    }
  });
  return { rows: rows, ids: ids };
}

/** "대한민국의 휴일" 캘린더 (구독 안 돼 있으면 구독) — 실패하면 null */
function holidayCalendar_() {
  const id = CONFIG.HOLIDAY_CALENDAR_ID;
  try {
    return CalendarApp.getCalendarById(id) || CalendarApp.subscribeToCalendar(id);
  } catch (err) {
    return null;
  }
}

/** 공휴일 행 (올해 1월 1일 ~ 내년 말), 가져오지 못하면 null */
function holidayRows_() {
  const cal = holidayCalendar_();
  if (!cal) return null;
  const y = today_().getFullYear();
  const start = new Date(y, 0, 1), end = new Date(y + 2, 0, 1);
  const rows = [];
  cal.getEvents(start, end).forEach(function (ev) {
    const last = dateOnly_(ev.getAllDayEndDate ? ev.getAllDayEndDate() : ev.getEndTime());
    for (let d = dateOnly_(ev.getAllDayStartDate ? ev.getAllDayStartDate() : ev.getStartTime()); d < last; d = addDays_(d, 1)) {
      rows.push([d, CONFIG.EVENT_TYPE.HOLIDAY, ev.getTitle(), 0, '', '', '']);
    }
  });
  return rows;
}

/**
 * 캘린더 일정 불러오기 (+ 시간표·공휴일 다시 만들기)
 * 기간 안의 기존 캘린더 행은 지우고 새로 채움 → 캘린더에서 지운 일정은 시트에서도 사라짐
 * 반환: 불러온 일정 수
 */
function syncCalendar_() {
  const range = syncRange_();
  const inRange = function (d) { return d >= range.start && d < range.end; };
  const cal = calendarRows_(range);
  const holidays = holidayRows_(); // 실패하면 예전 공휴일 유지
  const kept = readEvents_().filter(function (r) {
    const T = CONFIG.EVENT_TYPE;
    if (r[1] === T.TIMETABLE) return false; // 아래에서 새로 만듦
    if (r[1] === T.HOLIDAY) return !holidays;
    return r[1] !== T.CAL || !inRange(r[0]);
  });
  // 캘린더에도 만든 메모: 캘린더에 아직 있으면 "연동"(일정 칸에서 숨김), 지워졌으면 다시 보이게
  kept.forEach(function (r) {
    if (r[1] === CONFIG.EVENT_TYPE.MEMO && r[4] && inRange(r[0])) r[5] = cal.ids[r[4]] ? '연동' : '';
  });
  writeEvents_(kept.concat(cal.rows, timetableRows_(range), holidays || []));
  setStatus_('sync', '✓ ' + cal.rows.length + '개' + (holidays ? ' · 공휴일 ' + holidays.length + '일' : ' · 공휴일 못 가져옴') +
    ' (' + dateKey_(range.start) + ' ~ ' + dateKey_(addDays_(range.end, -1)) + ')');
  return cal.rows.length;
}

/* ---------- 적기 ---------- */

/** 시간 글자('14:30', '1430', '14') → {h, m}, 비어 있으면 null */
function parseTime_(text) {
  const s = String(text || '').trim().replace(/시\s*/, ':').replace(/분/, '');
  if (!s) return null;
  const m = s.match(/^(\d{1,2})(?::?(\d{2}))?(?::\d{2})?$/);
  if (!m || +m[1] > 23 || +(m[2] || 0) > 59) throw new Error('시간은 14:30처럼 입력해 주세요');
  return { h: +m[1], m: +(m[2] || 0) };
}

/** 날짜 + 시각 → Date */
function atTime_(date, t) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), t.h, t.m);
}

/** 같은 날 같은 제목·시간 일정이 이미 있는지 */
function hasDuplicate_(cal, date, title, start, end) {
  return cal.getEventsForDay(date).some(function (ev) {
    if (ev.getTitle() !== title) return false;
    if (!start) return ev.isAllDayEvent();
    return !ev.isAllDayEvent() && ev.getStartTime().getTime() === start.getTime() &&
      ev.getEndTime().getTime() === end.getTime();
  });
}

/**
 * 캘린더에 일정 만들기 (시트 입력칸·웹앱 공용)
 * 규칙: 시작+끝 → 그 시간대 / 시작만 → 종일 + 제목 앞에 시간 / 둘 다 없음 → 종일
 * 반환 {ok, message, id}. 입력이 잘못되면 Error를 던짐
 */
/**
 * 입력 → 캘린더 일정 모양 (만들기·수정 공용 규칙)
 * 시작+끝 → 그 시간대 / 시작만 → 종일 + 제목 앞에 시간 / 둘 다 없음 → 종일
 * 반환 {date, st, en, allDay, start, end, finalTitle, error}. 입력이 잘못되면 Error
 */
function eventShape_(dateValue, startText, endText, titleText) {
  const date = dateOnly_(dateValue) || today_();
  const title = String(titleText || '').trim();
  if (!title) throw new Error('제목을 입력해 주세요');
  const st = parseTime_(startText);
  const en = parseTime_(endText); // 끝시간은 선택
  if (en && !st) throw new Error('시작시간 확인 (끝시간만 넣을 수는 없어요)');
  const allDay = !(st && en);
  const start = allDay ? null : atTime_(date, st);
  const end = allDay ? null : atTime_(date, en);
  const hhmm = function (t) { return ('0' + t.h).slice(-2) + ':' + ('0' + t.m).slice(-2); };
  return {
    date: date, title: title, st: st, en: en, allDay: allDay, start: start, end: end,
    finalTitle: st && !en ? hhmm(st) + ' ' + title : title,
    error: !allDay && end <= start ? '끝시간 확인' : '',
  };
}

/**
 * 내 캘린더 일정 찾기: 저장해 둔 ID로 먼저, 없으면(예전 일정) 날짜 + 제목으로 (못 찾으면 null)
 * titles = 찾을 제목 후보들
 */
function findCalendarEvent_(id, date, titles) {
  const cal = calendar_();
  if (id) {
    try {
      const ev = cal.getEventById(id);
      if (ev) return ev;
    } catch (err) { /* ID로 못 찾으면 아래에서 날짜+제목 */ }
  }
  if (!date || !titles || !titles.length) return null;
  const list = cal.getEventsForDay(date).filter(function (ev) { return titles.indexOf(ev.getTitle()) >= 0; });
  return list[0] || null;
}

/** _일정 캘린더 줄 글자 → 원래 캘린더 제목 후보 ('11:00-13:00 회의' → ['회의', …]) */
function calendarTitles_(text) {
  const t = String(text);
  const out = [t];
  const m = t.match(/^\d{2}:\d{2}(?:-\d{2}:\d{2})?\s+(.+)$/);
  if (m) out.unshift(m[1]);
  return out;
}

/**
 * 캘린더에도 만든 일정 수정 (ID로 찾아서 제목·날짜·시간을 같은 규칙으로 바꿈)
 * 반환 {ok, found, message}. 캘린더에서 이미 지워졌으면 found=false
 */
function updateCalendarEvent_(id, dateValue, startText, endText, titleText, oldTitles, memo) {
  const shape = eventShape_(dateValue, startText, endText, titleText);
  if (shape.error) return { ok: false, found: true, message: shape.error };
  const ev = findCalendarEvent_(id, shape.date, oldTitles);
  if (!ev) {
    syncCalendar_(); // 남아 있던 옛 캘린더 줄 정리
    return { ok: false, found: false, message: '구글 캘린더에서 이미 지워진 일정이에요' };
  }
  ev.setTitle(shape.finalTitle);
  if (memo !== undefined) ev.setDescription(String(memo || '')); // 숨김 메모 = 캘린더 설명
  if (shape.allDay) ev.setAllDayDate(shape.date); // 시간 있던 일정도 종일로 바뀜
  else ev.setTime(shape.start, shape.end); // 종일이던 일정도 시간 일정으로 바뀜
  syncCalendar_();
  return { ok: true, found: true, message: '구글 캘린더 일정도 수정' };
}

/** 캘린더 일정 삭제 (ID → 없으면 날짜+제목) → 찾아서 지웠으면 true */
function deleteCalendarEvent_(id, date, oldTitles) {
  const ev = findCalendarEvent_(id, date, oldTitles);
  if (!ev) {
    syncCalendar_(); // 남아 있던 옛 캘린더 줄 정리
    return false;
  }
  ev.deleteEvent();
  syncCalendar_();
  return true;
}

function createCalendarEvent_(dateValue, startText, endText, titleText, memo) {
  const shape = eventShape_(dateValue, startText, endText, titleText);
  if (shape.error) return { ok: false, message: shape.error };
  const date = shape.date, title = shape.title, st = shape.st, en = shape.en;
  const allDay = shape.allDay, start = shape.start, end = shape.end, finalTitle = shape.finalTitle;
  const hhmm = function (t) { return ('0' + t.h).slice(-2) + ':' + ('0' + t.m).slice(-2); };

  // 중복 방지 ① 최근 2분 안에 같은 요청 ② 캘린더에 같은 일정 존재
  const key = 'add|' + dateKey_(date) + '|' + (st ? hhmm(st) : '') + '|' + (en ? hhmm(en) : '') + '|' + title;
  const cache = CacheService.getScriptCache();
  const cal = calendar_();
  if (cache.get(key) || hasDuplicate_(cal, date, finalTitle, start, end)) {
    return { ok: false, message: '이미 있는 일정이라 추가하지 않았어요 (중복 방지)' };
  }
  const opts = memo ? { description: String(memo) } : {}; // 숨김 메모 → 캘린더 설명
  const ev = allDay ? cal.createAllDayEvent(finalTitle, date, opts) : cal.createEvent(finalTitle, start, end, opts);
  cache.put(key, '1', 120);

  syncCalendar_(); // 바로 달력에 다시 불러오기
  const when = Utilities.formatDate(date, CONFIG.TIMEZONE, 'M/d') +
    (allDay ? ' 종일' : ' ' + Utilities.formatDate(start, CONFIG.TIMEZONE, 'HH:mm'));
  return { ok: true, message: '추가됨 ✓ ' + when + ' ' + finalTitle, id: ev.getId() };
}

/**
 * "새 일정 추가" 입력칸 처리 (먼슬리·오늘 시트)
 * F = {date, start, end, title, check, result} 각 [행, 열]
 */
function addEventFromForm_(sh, F) {
  const cell = function (k) { return sh.getRange(F[k][0], F[k][1]); };
  const out = cell('result');
  cell('check').setValue(false); // 먼저 체크 해제 → 두 번 눌러도 한 번만 처리
  out.setValue('추가 중…');
  SpreadsheetApp.flush();
  try {
    // 시간은 표시 글자로 읽음 (1899년 기준 시간값의 시간대 오차 방지)
    const res = createCalendarEvent_(cell('date').getValue(), cell('start').getDisplayValue(),
      cell('end').getDisplayValue(), cell('title').getValue());
    if (res.ok) ['date', 'start', 'end', 'title'].forEach(function (k) { cell(k).clearContent(); });
    out.setValue(res.message);
  } catch (err) {
    out.setValue('오류: ' + err.message);
  }
}
