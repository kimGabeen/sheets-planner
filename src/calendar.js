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

/* ---------- _일정 시트 (캘린더 일정 + 달력 메모) ---------- */

/** _일정 시트 만들기 (숨김) */
function buildEventsSheet_() {
  const sh = ensureSheet_(CONFIG.SHEETS.EVENTS);
  if (!sh.getRange(1, 1).getValue()) {
    sh.getRange(1, 1, 1, 5).setValues([['날짜', '구분', '내용', '정렬', '이벤트ID']]);
  }
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.hideSheet();
  return sh;
}

/** _일정 전체 읽기 → [[날짜, 구분, 내용, 정렬, ID], ...] */
function readEvents_() {
  const sh = sheet_(CONFIG.SHEETS.EVENTS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues()
    .map(function (r) { return [dateOnly_(r[0]), r[1], r[2], r[3], r[4]]; })
    .filter(function (r) { return r[0]; });
}

/** _일정 전체 다시 쓰기 (날짜 → 캘린더 먼저 → 시간 순 정렬) */
function writeEvents_(rows) {
  const sh = sheet_(CONFIG.SHEETS.EVENTS);
  const order = function (t) { return t === CONFIG.EVENT_TYPE.CAL ? 0 : 1; };
  rows.sort(function (a, b) {
    return (a[0] - b[0]) || (order(a[1]) - order(b[1])) || (a[3] - b[3]);
  });
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 5).clearContent();
  if (rows.length) {
    ensureSize_(sh, rows.length + 1, 5);
    sh.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

/** 달력 메모 저장 (같은 날짜 메모는 덮어쓰기, 빈 값이면 삭제) */
function saveMemos_(entries) {
  const keys = {};
  entries.forEach(function (e) { keys[dateKey_(e.date)] = true; });
  const rows = readEvents_().filter(function (r) {
    return !(r[1] === CONFIG.EVENT_TYPE.MEMO && keys[dateKey_(r[0])]);
  });
  entries.forEach(function (e) {
    if (e.text) rows.push([e.date, CONFIG.EVENT_TYPE.MEMO, e.text, 9999, '']);
  });
  writeEvents_(rows);
}

/* ---------- 불러오기 ---------- */

/** 불러올 기간: 이번 달 1일 ~ 다음 달 말 + 먼슬리에서 보는 달 + 위클리 주 */
function syncRange_() {
  const now = today_();
  let start = new Date(now.getFullYear(), now.getMonth(), 1);
  let end = new Date(now.getFullYear(), now.getMonth() + 1 + CONFIG.CALENDAR_MONTHS_AHEAD, 1);
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
 * 캘린더 일정 불러오기
 * 기간 안의 기존 캘린더 행은 지우고 새로 채움 → 캘린더에서 지운 일정은 시트에서도 사라짐
 * 반환: 불러온 일정 수
 */
function syncCalendar_() {
  const range = syncRange_();
  const events = calendar_().getEvents(range.start, range.end);
  const inRange = function (d) { return d >= range.start && d < range.end; };
  const fresh = [];
  events.forEach(function (ev) {
    const title = ev.getTitle() || '(제목 없음)';
    if (ev.isAllDayEvent()) {
      // 여러 날 종일 일정은 날마다 한 줄씩
      const last = dateOnly_(ev.getAllDayEndDate()); // 끝 날짜는 포함 안 됨
      for (let d = dateOnly_(ev.getAllDayStartDate()); d < last; d = addDays_(d, 1)) {
        if (inRange(d)) fresh.push([d, CONFIG.EVENT_TYPE.CAL, title, 0, ev.getId()]);
      }
    } else {
      const s = ev.getStartTime(), d = dateOnly_(s);
      if (inRange(d)) {
        fresh.push([d, CONFIG.EVENT_TYPE.CAL, Utilities.formatDate(s, CONFIG.TIMEZONE, 'HH:mm') + ' ' + title,
          s.getHours() * 60 + s.getMinutes() + 1, ev.getId()]);
      }
    }
  });
  const kept = readEvents_().filter(function (r) {
    return r[1] !== CONFIG.EVENT_TYPE.CAL || !inRange(r[0]);
  });
  writeEvents_(kept.concat(fresh));
  setStatus_('sync', '✓ ' + fresh.length + '개 (' + dateKey_(range.start) + ' ~ ' + dateKey_(addDays_(range.end, -1)) + ')');
  return fresh.length;
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
 * 반환 {ok, message}. 입력이 잘못되면 Error를 던짐
 */
function createCalendarEvent_(dateValue, startText, endText, titleText) {
  const date = dateOnly_(dateValue) || today_();
  const title = String(titleText || '').trim();
  if (!title) throw new Error('제목을 입력해 주세요');
  const st = parseTime_(startText);
  const en = parseTime_(endText); // 끝시간은 선택
  if (en && !st) throw new Error('시작시간 확인 (끝시간만 넣을 수는 없어요)');
  const allDay = !(st && en);
  const start = allDay ? null : atTime_(date, st);
  const end = allDay ? null : atTime_(date, en);
  if (!allDay && end <= start) return { ok: false, message: '끝시간 확인' };
  const hhmm = function (t) { return ('0' + t.h).slice(-2) + ':' + ('0' + t.m).slice(-2); };
  const finalTitle = st && !en ? hhmm(st) + ' ' + title : title;

  // 중복 방지 ① 최근 2분 안에 같은 요청 ② 캘린더에 같은 일정 존재
  const key = 'add|' + dateKey_(date) + '|' + (st ? hhmm(st) : '') + '|' + (en ? hhmm(en) : '') + '|' + title;
  const cache = CacheService.getScriptCache();
  const cal = calendar_();
  if (cache.get(key) || hasDuplicate_(cal, date, finalTitle, start, end)) {
    return { ok: false, message: '이미 있는 일정이라 추가하지 않았어요 (중복 방지)' };
  }
  if (allDay) cal.createAllDayEvent(finalTitle, date);
  else cal.createEvent(finalTitle, start, end);
  cache.put(key, '1', 120);

  syncCalendar_(); // 바로 달력에 다시 불러오기
  const when = Utilities.formatDate(date, CONFIG.TIMEZONE, 'M/d') +
    (allDay ? ' 종일' : ' ' + Utilities.formatDate(start, CONFIG.TIMEZONE, 'HH:mm'));
  return { ok: true, message: '추가됨 ✓ ' + when + ' ' + finalTitle };
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
