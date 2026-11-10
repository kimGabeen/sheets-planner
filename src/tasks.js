/**
 * tasks.js
 * 할 일 10줄 관리: 시간순 자동 정렬, 못 한 일 다음 날로 자동 넘기기
 * 위클리·오늘 시트 모두 "체크 | 시간 | 내용" 3열 묶음을 같은 함수로 다룬다.
 */

/** 시간 글자 → 'HH:mm' (형식이 아니면 '') */
function normTime_(text) {
  try {
    const t = parseTime_(text);
    return t ? ('0' + t.h).slice(-2) + ':' + ('0' + t.m).slice(-2) : '';
  } catch (err) {
    return '';
  }
}

/** 전날 칸에 남은 "미완 →" 할 일인지 */
function isMissed_(text) {
  return String(text).trim().indexOf(CONFIG.CARRY.MISSED.trim()) === 0;
}

/** 앞에 붙은 "미완 →", "↪" 표시 떼기 */
function stripCarry_(text) {
  let s = String(text).trim();
  const marks = [CONFIG.CARRY.MISSED.trim(), CONFIG.CARRY.CARRIED.trim()];
  let changed = true;
  while (changed) {
    changed = false;
    marks.forEach(function (m) {
      if (s.indexOf(m) === 0) { s = s.slice(m.length).trim(); changed = true; }
    });
  }
  return s;
}

/**
 * 할 일 10줄 정렬 (row, col = 체크박스 칸의 첫 줄)
 * 순서: ① 체크 안 된 일 중 시간 있는 것(시간순) ② 시간 없는 것(입력 순)
 *       ③ 체크된 일·"미완 →" 일 ④ 빈 줄
 * 순서가 그대로면 쓰지 않음. 반환: 바뀌었는지
 */
function sortTodos_(sh, row, col) {
  const n = CONFIG.TODO_ROWS;
  const range = sh.getRange(row, col, n, 3);
  const vals = range.getValues(), disp = range.getDisplayValues();
  const items = vals.map(function (r, i) {
    const raw = String(disp[i][1]).trim();
    return { done: r[0] === true, time: normTime_(raw) || raw, sortTime: normTime_(raw),
      text: String(r[2]).trim(), i: i };
  });
  const group = function (it) {
    if (!it.text) return 3;
    if (it.done || isMissed_(it.text)) return 2;
    return it.sortTime ? 0 : 1;
  };
  const sorted = items.slice().sort(function (a, b) {
    const g = group(a) - group(b);
    if (g) return g;
    if (group(a) === 0 && a.sortTime !== b.sortTime) return a.sortTime < b.sortTime ? -1 : 1;
    return a.i - b.i;
  });
  const out = sorted.map(function (it) { return [it.done, it.time, it.text]; });
  const before = items.map(function (it) { return [it.done, String(disp[it.i][1]).trim(), it.text]; });
  if (JSON.stringify(out) === JSON.stringify(before)) return false;
  range.setValues(out);
  return true;
}

/** 설정 시트의 자동 넘기기 체크 여부 */
function carryEnabled_() {
  return sheet_(CONFIG.SHEETS.SETTINGS).getRange(SETTINGS.CARRY.row, SETTINGS.CARRY.col).getValue() === true;
}

/**
 * 못 한 일 자동 넘기기 (매일 0시, 하루 한 번만)
 * 1) 전날 체크 안 된 할 일 → 전날 칸은 "미완 → 할 일"로 남기고 기록에 FALSE + "이월" 저장
 * 2) 전날이 일요일이면 여기서 새 주 시작 (전날 표시·저장을 먼저 끝낸 뒤)
 * 3) 오늘 칸의 빈 줄에 "↪ 할 일"로 복사 (시간도 같이), 자리가 없으면 오늘 시트 "넘어온 일"로
 */
function carryOver_() {
  const props = PropertiesService.getDocumentProperties();
  const today = today_(), y = addDays_(today, -1), yKey = dateKey_(y);
  if (props.getProperty('CARRY_DONE') === yKey) return ''; // 이미 처리함
  if (!carryEnabled_()) {
    props.setProperty('CARRY_DONE', yKey);
    return '';
  }
  const W = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY), n = CONFIG.TODO_ROWS;
  const C = CONFIG.CARRY;

  // 1) 전날 미완료 모으기 + "미완 →" 표시
  const items = [];
  const yd = weeklyIndexOf_(y);
  if (yd >= 0) {
    const range = sh.getRange(W.TODO_START, W.checkCol(yd), n, 3);
    const vals = range.getValues(), disp = range.getDisplayValues();
    const texts = vals.map(function (r) { return [r[2]]; });
    vals.forEach(function (r, i) {
      const t = String(r[2]).trim();
      if (!t || r[0] === true || isMissed_(t)) return;
      const name = stripCarry_(t);
      items.push({ time: normTime_(disp[i][1]), text: name });
      texts[i][0] = C.MISSED + name;
    });
    if (items.length) {
      sh.getRange(W.TODO_START, W.textCol(yd), n, 1).setValues(texts);
      sortTodos_(sh, W.TODO_START, W.checkCol(yd));
      archiveDayLive_(y); // 전날 날짜로 FALSE + 이월
    }
  }

  // 2) 일요일 → 새 주 (지난주 저장 포함)
  ensureCurrentWeek_();

  // 3) 오늘 칸 빈 줄에 넣기
  let overflow = items;
  const td = weeklyIndexOf_(today);
  if (td >= 0 && items.length) {
    const range = sh.getRange(W.TODO_START, W.checkCol(td), n, 3);
    const vals = range.getValues(), disp = range.getDisplayValues();
    const rows = vals.map(function (r, i) { return [r[0], disp[i][1], r[2]]; });
    overflow = [];
    items.forEach(function (it) {
      const free = rows.findIndex(function (r) { return !String(r[2]).trim() && !String(r[1]).trim(); });
      if (free < 0) { overflow.push(it); return; }
      rows[free] = [false, it.time, C.CARRIED + it.text];
    });
    range.setValues(rows);
    sortTodos_(sh, W.TODO_START, W.checkCol(td));
    archiveDayLive_(today);
  }

  props.setProperty('CARRY_OVERFLOW', JSON.stringify({ date: dateKey_(today), items: overflow }));
  props.setProperty('CARRY_DONE', yKey);
  const msg = items.length ? '✓ ' + items.length + '개 넘김' + (overflow.length ? ' (넘어온 일 ' + overflow.length + '개)' : '')
    : '✓ 넘길 일 없음';
  setStatus_('carry', msg);
  return msg;
}
