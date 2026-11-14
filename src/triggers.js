/**
 * triggers.js
 * 트리거 설치(중복 없이)와 트리거가 부르는 함수들, 맥 브라우저용 메뉴
 */

/** 이 프로젝트가 쓰는 트리거 핸들러 이름 */
const TRIGGER_HANDLERS_ = ['handleEdit', 'syncCalendarHourly', 'nightlyArchive', 'midnightJob'];

/** 트리거 설치: 기존 것을 지우고 새로 만들어 몇 번 실행해도 1개씩만 존재 */
function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (TRIGGER_HANDLERS_.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  const T = CONFIG.TRIGGER;
  // 캘린더 접근이 필요하므로 단순 onEdit 대신 설치형 onEdit
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss_()).onEdit().create();
  ScriptApp.newTrigger('syncCalendarHourly').timeBased().everyHours(T.SYNC_EVERY_HOURS).create();
  ScriptApp.newTrigger('nightlyArchive').timeBased().everyDays(1)
    .atHour(T.NIGHTLY_HOUR).nearMinute(T.NIGHTLY_MINUTE).create();
  ScriptApp.newTrigger('midnightJob').timeBased().everyDays(1)
    .atHour(T.MIDNIGHT_HOUR).nearMinute(T.MIDNIGHT_MINUTE).create();
}

/** 트리거 함수 공통: 잠금 + 오류는 설정 시트에 기록 */
function runSafely_(name, fn) {
  try {
    return withLock_(fn);
  } catch (err) {
    setStatus_('error', name + ': ' + err.message);
    console.error(name, err);
  }
}

/** [설치형 onEdit] 시트별로 나눠 처리 */
function handleEdit(e) {
  if (!e || !e.range) return;
  const name = e.range.getSheet().getName(), S = CONFIG.SHEETS;
  runSafely_('편집 처리', function () {
    if (name === S.WEEKLY) onWeeklyEdit_(e);
    else if (name === S.TODAY) onTodayEdit_(e);
    else if (name === S.MONTHLY) onMonthlyEdit_(e);
    else if (name === S.SETTINGS) onSettingsEdit_(e);
  });
}

/** [1시간마다] 캘린더 불러오기 + 날짜 넘어갔는지 확인 */
function syncCalendarHourly() {
  runSafely_('캘린더 불러오기', function () {
    dailyRollover_();
    syncCalendar_();
    flushArchive_(); // 미뤄 둔 기록 저장
  });
}

/** [매일 23:50] 그날 할 일·습관·기분을 기록 시트에 저장 */
function nightlyArchive() {
  runSafely_('밤 저장', function () { saveAll_(); });
}

/** [매일 0시] 못 한 일 넘기기 → (월요일이면) 새 주 시작 → 오늘 시트를 새 날짜로 */
function midnightJob() {
  runSafely_('자정 작업', function () {
    carryOver_();
    ensureCurrentWeek_();
    refreshToday_();
  });
}

/**
 * 날짜가 바뀐 뒤 할 일 (자정 트리거가 늦거나 빠져도 1시간마다 다시 확인)
 * 넘기기는 하루 한 번만 실행되고, 새 주 시작보다 먼저 처리됨
 */
function dailyRollover_() {
  carryOver_();
  ensureCurrentWeek_();
  if (String(todaySheetDate_()) !== String(today_())) refreshToday_();
}

/** 위클리가 지난 주에 머물러 있으면 새 주 시작 (트리거가 늦거나 빠져도 복구) */
function ensureCurrentWeek_() {
  const mon = weeklyMonday_();
  if (mon && mon >= mondayOf_(today_())) return;
  setStatus_('week', startNewWeek_());
}

/* ---------- 맥 브라우저용 메뉴 (폰·아이패드 앱에서는 체크박스 사용) ---------- */

/** 시트를 열 때 메뉴 추가 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('플래너')
    .addItem('지금 기록 저장', 'menuSaveNow')
    .addItem('캘린더 새로고침', 'menuSync')
    .addItem('오늘 시트 새로고침', 'menuRefreshToday')
    .addItem('새 주 시작 (지난 주일 때만)', 'menuNewWeek')
    .addSeparator()
    .addItem('처음 설정 / 다시 그리기 (setup)', 'setup')
    .addToUi();
}

/** 메뉴 결과는 위클리의 결과 칸에 글자로 표시 */
function menuResult_(cell, fn) {
  const sh = sheet_(CONFIG.SHEETS.WEEKLY);
  runSafely_('메뉴', function () {
    sh.getRange(cell.row, WEEKLY.ACTION_RESULT_COL).setValue(fn() + '  ' + stamp_());
  });
}

function menuSaveNow() {
  menuResult_(WEEKLY.ACTION_SAVE, function () { return '저장됨 ✓ ' + saveAll_() + '행'; });
}

function menuSync() {
  menuResult_(WEEKLY.ACTION_SYNC, function () { return '불러옴 ✓ ' + syncCalendar_() + '개'; });
}

function menuRefreshToday() {
  runSafely_('오늘 새로고침', refreshToday_);
}

function menuNewWeek() {
  menuResult_(WEEKLY.ACTION_SAVE, function () {
    const msg = startNewWeek_();
    refreshToday_();
    return msg;
  });
}
