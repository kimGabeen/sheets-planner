/**
 * utils.js
 * 여러 파일에서 같이 쓰는 작은 도우미 함수 모음 (날짜, 범위, 서식, 상태 표시)
 */

/** 현재 스프레드시트 */
function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** 이름으로 시트 가져오기 */
function sheet_(name) {
  return ss_().getSheetByName(name);
}

/** 시트가 없으면 만들고 반환 (있으면 그대로) */
function ensureSheet_(name) {
  return sheet_(name) || ss_().insertSheet(name);
}

/** 시트의 행·열 수가 모자라면 늘리기 */
function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}

/* ---------- 날짜 ---------- */

/** 오늘 0시 */
function today_() {
  const n = new Date();
  return new Date(n.getFullYear(), n.getMonth(), n.getDate());
}

/** 값(Date 또는 'yyyy-MM-dd')을 날짜(0시)로 바꾸기, 실패하면 null */
function dateOnly_(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const m = String(v || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}

/** n일 더하기 */
function addDays_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}

/** 두 날짜 사이 일수 (b - a) */
function diffDays_(a, b) {
  return Math.round((b - a) / 86400000);
}

/** 그 주의 월요일 */
function mondayOf_(d) {
  return addDays_(d, -((d.getDay() + 6) % 7));
}

/** 'yyyy-MM-dd' 문자열 */
function dateKey_(d) {
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

/** 주차 키 = 그 주 월요일의 'yyyy-MM-dd' */
function weekKey_(d) {
  return dateKey_(mondayOf_(d));
}

/** 'M/d HH:mm' 시각 도장 */
function stamp_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'M/d HH:mm');
}

/* ---------- 범위·수식 ---------- */

/** 열 번호 → 열 문자 (28 → AB) */
function colLetter_(c) {
  let s = '';
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

/** 행·열 → A1 표기 (abs=true면 $A$1) */
function a1_(row, col, abs) {
  return abs ? '$' + colLetter_(col) + '$' + row : colLetter_(col) + row;
}

/** 다른 시트 열 전체 참조 ('기록'!$A$2:$A) */
function colRef_(sheetName, col) {
  const c = colLetter_(col);
  return "'" + sheetName + "'!$" + c + '$2:$' + c;
}

/** 수정된 범위가 주어진 영역과 겹치는지 */
function hit_(range, row, col, nRows, nCols) {
  const r1 = range.getRow(), r2 = range.getLastRow();
  const c1 = range.getColumn(), c2 = range.getLastColumn();
  return r1 <= row + (nRows || 1) - 1 && r2 >= row && c1 <= col + (nCols || 1) - 1 && c2 >= col;
}

/** 체크박스 셀이 방금 체크되었는지 */
function justChecked_(e, row, col) {
  return hit_(e.range, row, col) && e.range.getSheet().getRange(row, col).getValue() === true;
}

/* ---------- 서식 ---------- */

/** 테마색을 흰색과 섞어 연한 색 만들기 */
function tint_(hex, ratio) {
  const n = parseInt(hex.slice(1), 16);
  const mix = function (c) { return Math.round(c + (255 - c) * ratio); };
  const r = mix((n >> 16) & 255), g = mix((n >> 8) & 255), b = mix(n & 255);
  return '#' + [r, g, b].map(function (x) { return ('0' + x.toString(16)).slice(-2); }).join('');
}

/** 테마색을 어둡게 */
function shade_(hex, ratio) {
  const n = parseInt(hex.slice(1), 16);
  const f = function (c) { return Math.round(c * (1 - ratio)); };
  return '#' + [f((n >> 16) & 255), f((n >> 8) & 255), f(n & 255)]
    .map(function (x) { return ('0' + x.toString(16)).slice(-2); }).join('');
}

/** 머리글 스타일 (테마색 배경 + 흰 글씨) */
function styleHeader_(range, theme) {
  range.setBackground(theme).setFontColor(CONFIG.COLORS.WHITE).setFontWeight('bold')
    .setVerticalAlignment('middle');
}

/** 체크박스 넣기 (기존 체크 값은 유지) */
function checkbox_(range) {
  range.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build())
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontSize(CONFIG.SIZE.CHECK_FONT);
}

/** 입력칸 밑줄 (테두리 최소화) */
function underline_(range) {
  range.setBorder(null, null, true, null, null, null, CONFIG.COLORS.LINE,
    SpreadsheetApp.BorderStyle.SOLID);
}

/** 날짜 입력 검증 (폰에서 날짜 선택기 표시) */
function dateInput_(range, format) {
  range.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build())
    .setNumberFormat(format || 'yyyy-mm-dd');
}

/** 기분 드롭다운 */
function moodInput_(range) {
  range.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.MOODS, true).build()).setHorizontalAlignment('center');
}

/** 기분 값('5 😄' 또는 5) → 점수 숫자, 없으면 '' */
function moodScore_(v) {
  const n = parseInt(String(v).trim().charAt(0), 10);
  return n >= 1 && n <= 5 ? n : '';
}

/** 기분 점수에 따라 배경색 바꾸는 조건부 서식 규칙 3개 만들기 */
function moodRules_(sh, ranges, scoreExpr, theme) {
  const mk = function (cond, color) {
    return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=' + cond)
      .setBackground(color).setRanges(ranges).build();
  };
  return [
    mk(scoreExpr + '>=4', tint_(theme, CONFIG.TINT.LIGHT)),
    mk(scoreExpr + '=3', tint_(theme, CONFIG.TINT.PALE)),
    mk('AND(' + scoreExpr + '>=1,' + scoreExpr + '<=2)', CONFIG.COLORS.MOOD_LOW),
  ];
}

/** 시트를 새로 그리기 전에 병합·차트 정리 (입력 값은 건드리지 않음) */
function resetLayout_(sh) {
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.clearConditionalFormatRules();
}

/** 설정 시트 상태 칸에 결과 남기기 (트리거 실행 확인용) */
function setStatus_(key, msg) {
  const sh = sheet_(CONFIG.SHEETS.SETTINGS);
  const row = SETTINGS.STATUS[key];
  if (sh && row) sh.getRange(row, 2).setValue(stamp_() + '  ' + msg);
}

/** 문서 잠금을 잡고 함수 실행 (동시 실행 방지) */
function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
