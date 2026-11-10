/**
 * timeline.js
 * 타임라인 시트: 웹앱에서 10분 칸에 칠한 카테고리를 "한 줄에 한 칸"으로 저장
 * 열: 날짜 | 시작(HH:mm) | 카테고리  → 기록 시트처럼 CSV 분석에 바로 사용
 */

/** 하루 칸 수 (10분 → 144칸) */
function slotCount_() {
  return 24 * 60 / CONFIG.TIMELINE.SLOT_MINUTES;
}

/** 칸 번호 → 'HH:mm' */
function slotTime_(i) {
  const m = i * CONFIG.TIMELINE.SLOT_MINUTES;
  return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2);
}

/** 타임라인 시트 만들기 (데이터는 유지) */
function buildTimelineSheet_(theme) {
  const sh = ensureSheet_(CONFIG.SHEETS.TIMELINE);
  ensureSize_(sh, 50, 3);
  if (!sh.getRange(1, 1).getValue()) sh.getRange(1, 1, 1, 3).setValues([['날짜', '시작', '카테고리']]);
  styleHeader_(sh.getRange(1, 1, 1, 3), theme);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('B:B').setNumberFormat('@'); // 시작 시각은 글자로
  sh.setColumnWidth(1, 100).setColumnWidth(2, 70).setColumnWidth(3, 120);
  return sh;
}

/** 하루치 칸 읽기 → 길이 144 배열 (카테고리 이름 또는 '') */
function readTimeline_(date) {
  const sh = sheet_(CONFIG.SHEETS.TIMELINE) || buildTimelineSheet_(getTheme_());
  const slots = new Array(slotCount_()).fill('');
  const last = sh.getLastRow();
  if (last < 2) return slots;
  const key = dateKey_(date);
  sh.getRange(2, 1, last - 1, 3).getDisplayValues().forEach(function (r) {
    if (r[0] !== key) return;
    const t = parseTime_(r[1]);
    const i = t ? (t.h * 60 + t.m) / CONFIG.TIMELINE.SLOT_MINUTES : -1;
    if (i >= 0 && i < slots.length) slots[i] = r[2];
  });
  return slots;
}

/** 하루치 칸 저장 (그날 행을 지우고 칠한 칸만 다시 씀) */
function writeTimeline_(date, slots) {
  const sh = sheet_(CONFIG.SHEETS.TIMELINE) || buildTimelineSheet_(getTheme_());
  const key = dateKey_(date);
  const last = sh.getLastRow();
  if (last >= 2) {
    const del = [];
    sh.getRange(2, 1, last - 1, 1).getDisplayValues().forEach(function (r, i) {
      if (r[0] === key) del.push(i + 2);
    });
    deleteRowList_(sh, del);
  }
  const rows = [];
  slots.forEach(function (cat, i) {
    if (cat) rows.push([date, slotTime_(i), String(cat)]);
  });
  if (rows.length) {
    const at = Math.max(sh.getLastRow(), 1) + 1;
    ensureSize_(sh, at + rows.length + 5, 3);
    sh.getRange(at, 1, rows.length, 3).setValues(rows);
  }
  const end = sh.getLastRow();
  if (end > 2) sh.getRange(2, 1, end - 1, 3).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
}
