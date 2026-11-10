/**
 * setup.js
 * setup() 한 번 실행하면 시트·서식·수식·차트·트리거까지 전부 만든다.
 * 다시 실행해도 안전: 입력한 내용과 기록 시트 데이터는 지우지 않고 서식·수식만 다시 그림.
 */

/** 처음 설정 (Apps Script 편집기 또는 메뉴에서 실행) */
function setup() {
  const ss = ss_();
  ss.setSpreadsheetLocale(CONFIG.LOCALE);
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  // 1) 데이터 시트 (있으면 그대로 둠)
  const theme = buildSettings_();
  buildEventsSheet_();
  buildArchive_(theme);

  // 2) 화면 시트
  buildAllLayouts_(theme);

  // 3) 처음 값 채우기 (비어 있을 때만)
  initDefaults_();
  arrangeSheets_();

  // 4) 데이터 불러오기·맞추기
  try {
    syncCalendar_();
  } catch (err) {
    setStatus_('sync', '오류: ' + err.message);
  }
  refreshToday_();
  refreshMonthMemos_();
  loadMonthGoals_();
  initLastWeekSelect_();

  // 첫날에는 어제 미완료를 넘기지 않도록 표시만 해 둠
  const props = PropertiesService.getDocumentProperties();
  if (!props.getProperty('CARRY_DONE')) props.setProperty('CARRY_DONE', dateKey_(addDays_(today_(), -1)));

  // 5) 트리거 (중복 없이)
  installTriggers_();
  setStatus_('setup', '✓ 완료 (트리거 ' + ScriptApp.getProjectTriggers().length + '개)');
  ss.setActiveSheet(sheet_(CONFIG.SHEETS.WEEKLY));
}

/** 화면 시트 전부 다시 그리기 (테마 바꿀 때도 사용) */
function buildAllLayouts_(theme) {
  const migrating = needsMigration_();
  styleSettings_(sheet_(CONFIG.SHEETS.SETTINGS), theme);
  buildArchive_(theme);
  buildWeekly_(theme); // 예전 레이아웃이면 입력값을 새 위치로 옮김
  buildToday_(theme);
  buildMonthly_(theme);
  buildLastWeek_(theme);
  PropertiesService.getDocumentProperties().setProperty('LAYOUT_VERSION', CONFIG.LAYOUT_VERSION);
  if (migrating) refreshToday_(); // 오늘 시트는 위클리에서 다시 채움
}

/** 비어 있는 칸에만 기본값 (주 시작 = 이번 주 월요일, 먼슬리 = 이번 달) */
function initDefaults_() {
  const now = today_();
  const w = sheet_(CONFIG.SHEETS.WEEKLY).getRange(WEEKLY.WEEK_START.row, WEEKLY.WEEK_START.col);
  if (!w.getValue()) w.setValue(mondayOf_(now));
  const m = sheet_(CONFIG.SHEETS.MONTHLY);
  const y = m.getRange(MONTHLY.YEAR.row, MONTHLY.YEAR.col);
  const mo = m.getRange(MONTHLY.MONTH.row, MONTHLY.MONTH.col);
  if (!y.getValue()) y.setValue(String(now.getFullYear()));
  if (!mo.getValue()) mo.setValue(String(now.getMonth() + 1));
}

/** 시트 순서 정리 + 처음 생긴 빈 기본 시트 삭제 */
function arrangeSheets_() {
  const ss = ss_(), S = CONFIG.SHEETS;
  const order = [S.WEEKLY, S.TODAY, S.MONTHLY, S.LASTWEEK, S.ARCHIVE, S.SETTINGS, S.EVENTS];
  ss.getSheets().forEach(function (sh) {
    const n = sh.getName();
    if (order.indexOf(n) < 0 && sh.getLastRow() === 0 && /^(시트|Sheet)\d+$/.test(n)) ss.deleteSheet(sh);
  });
  order.forEach(function (n, i) {
    const sh = sheet_(n);
    if (!sh) return;
    if (sh.isSheetHidden()) sh.showSheet();
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });
  sheet_(S.EVENTS).hideSheet();
}
