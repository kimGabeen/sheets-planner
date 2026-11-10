/**
 * settings.js
 * 설정 시트: 테마색, 습관 목록, 자동 실행 상태 표시
 */

const DEFAULT_HABITS_ = ['물 8잔 마시기', '운동 30분', '독서 20분', '영어 공부', '일기 쓰기'];
const THEME_PRESETS_ = '초록 #2E7D32 · 파랑 #1565C0 · 남색 #283593 · 보라 #6A1B9A · 주황 #EF6C00 · 회색 #455A64';

/** 설정 시트 만들기 (이미 입력한 값은 유지) → 테마색 반환 */
function buildSettings_() {
  const sh = ensureSheet_(CONFIG.SHEETS.SETTINGS);
  const S = SETTINGS;
  ensureSize_(sh, 30, 3);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 130).setColumnWidth(2, 280).setColumnWidth(3, 260);
  sh.setRowHeights(1, 30, CONFIG.SIZE.ROW);

  // 처음 만들 때만 기본값 채우기
  const themeCell = sh.getRange(S.THEME.row, S.THEME.col);
  if (!themeCell.getValue()) themeCell.setValue(CONFIG.DEFAULT_THEME);
  const habitRange = sh.getRange(S.HABIT_START, 2, CONFIG.HABIT_COUNT, 1);
  if (habitRange.getValues().every(function (r) { return !r[0]; })) {
    habitRange.setValues(habitRange.getValues().map(function (r, i) { return [DEFAULT_HABITS_[i] || '']; }));
  }

  const theme = getTheme_();
  styleSettings_(sh, theme);
  return theme;
}

/** 설정 시트 서식 (테마 바뀔 때도 다시 실행) */
function styleSettings_(sh, theme) {
  const S = SETTINGS;
  sh.getRange(1, 1, 1, 3).merge().setValue('설정').setFontSize(18).setFontWeight('bold').setFontColor(theme);
  sh.getRange(S.THEME.row, 1).setValue('테마색').setFontWeight('bold');
  sh.getRange(S.THEME.row, 2).setBackground(theme).setFontColor(CONFIG.COLORS.WHITE)
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(S.THEME.row, 3).setValue('← #2E7D32 같은 색 코드를 넣으면 전체 색이 바뀌어요')
    .setFontColor(CONFIG.COLORS.SUB).setFontSize(9);

  // 못 한 일 자동 넘기기 (처음엔 켜짐)
  const carry = sh.getRange(S.CARRY.row, S.CARRY.col);
  if (!carry.getDataValidation()) carry.setValue(true);
  sh.getRange(S.CARRY.row, 1).setValue('못 한 일 자동 넘기기').setFontWeight('bold');
  checkbox_(carry);
  sh.getRange(S.CARRY.row, 3).setValue('← 켜면 매일 0시에 전날 미완료를 다음 날로 넘겨요')
    .setFontColor(CONFIG.COLORS.SUB).setFontSize(9);

  styleHeader_(sh.getRange(S.HABIT_HEADER, 1, 1, 3).merge()
    .setValue('습관 목록 (최대 ' + CONFIG.HABIT_COUNT + '개) — 위클리·오늘·먼슬리에 자동 반영'), theme);
  for (let i = 0; i < CONFIG.HABIT_COUNT; i++) {
    sh.getRange(S.HABIT_START + i, 1).setValue(i + 1).setHorizontalAlignment('center')
      .setFontColor(CONFIG.COLORS.SUB);
    underline_(sh.getRange(S.HABIT_START + i, 2, 1, 2).merge());
  }

  sh.getRange(S.PRESET_ROW, 1).setValue('테마 예시').setFontColor(CONFIG.COLORS.SUB);
  sh.getRange(S.PRESET_ROW, 2, 1, 2).merge().setValue(THEME_PRESETS_).setFontColor(CONFIG.COLORS.SUB)
    .setFontSize(9);

  styleHeader_(sh.getRange(S.STATUS_HEADER, 1, 1, 3).merge().setValue('자동 실행 상태 (읽기 전용)'), theme);
  const labels = { setup: '처음 설정', save: '기록 저장', sync: '캘린더 불러오기', week: '새 주 시작', carry: '못 한 일 넘기기', error: '최근 오류' };
  Object.keys(S.STATUS).forEach(function (k) {
    sh.getRange(S.STATUS[k], 1).setValue(labels[k]).setFontColor(CONFIG.COLORS.SUB);
    sh.getRange(S.STATUS[k], 2, 1, 2).merge().setFontColor(CONFIG.COLORS.SUB).setFontSize(9);
  });
}

/** 테마색 읽기 (형식이 틀리면 기본색) */
function getTheme_() {
  const sh = sheet_(CONFIG.SHEETS.SETTINGS);
  const v = sh ? String(sh.getRange(SETTINGS.THEME.row, SETTINGS.THEME.col).getValue()).trim() : '';
  return /^#[0-9a-fA-F]{6}$/.test(v) ? v.toUpperCase() : CONFIG.DEFAULT_THEME;
}

/** 습관 이름 목록 (빈 칸은 '') */
function getHabits_() {
  const sh = sheet_(CONFIG.SHEETS.SETTINGS);
  return sh.getRange(SETTINGS.HABIT_START, 2, CONFIG.HABIT_COUNT, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); });
}

/** 다른 시트 수식에서 쓸 습관 이름 셀 참조 */
function habitRef_(i) {
  return "'" + CONFIG.SHEETS.SETTINGS + "'!$B$" + (SETTINGS.HABIT_START + i);
}

/** 설정 시트 수정 시: 테마색이 바뀌면 전체 다시 칠하기 */
function onSettingsEdit_(e) {
  const S = SETTINGS;
  if (!hit_(e.range, S.THEME.row, S.THEME.col)) return;
  const v = String(e.range.getSheet().getRange(S.THEME.row, S.THEME.col).getValue()).trim();
  const note = e.range.getSheet().getRange(S.THEME.row, 3);
  if (!/^#[0-9a-fA-F]{6}$/.test(v)) {
    note.setValue('색 코드 형식이 아니에요 (예: #2E7D32)');
    return;
  }
  note.setValue('적용 중…');
  SpreadsheetApp.flush();
  buildAllLayouts_(getTheme_());
  note.setValue('적용됨 ✓ ' + stamp_());
}
