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

/** 시간 범위 글자('11-13', '11:00~13:00', '1100-1300', '15:00') → 'HH:mm' 또는 'HH:mm-HH:mm' (아니면 '') */
function normRange_(text) {
  const s = String(text || '').trim();
  if (!s) return '';
  const parts = s.split(/\s*[-~–]\s*/);
  if (parts.length === 1) return normTime_(s);
  if (parts.length !== 2) return '';
  const a = normTime_(parts[0]), b = normTime_(parts[1]);
  if (a && !parts[1]) return a;
  return a && b && b > a ? a + '-' + b : '';
}

/** 시간 범위의 시작 'HH:mm' (정렬용) */
function rangeStart_(text) {
  return normRange_(text).slice(0, 5);
}

/** 내용 앞이나 뒤에 적은 시간 범위를 떼어 냄 ('11-13 회의' → {time: '11:00-13:00', text: '회의'}), 없으면 null */
function splitTimeFromText_(text) {
  const R = '(\\d{1,2}(?::?\\d{2})?\\s*[-~–]\\s*\\d{1,2}(?::?\\d{2})?)';
  const tries = [new RegExp('^' + R + '\\s+(.+)$'), new RegExp('^(.+?)\\s+' + R + '$')];
  for (let k = 0; k < tries.length; k++) {
    const m = String(text).trim().match(tries[k]);
    if (!m) continue;
    const range = k === 0 ? m[1] : m[2], rest = k === 0 ? m[2] : m[1];
    const t = normRange_(range);
    if (t.length === 11) return { time: t, text: rest.trim() }; // 시작·끝 둘 다 있을 때만
  }
  return null;
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


/** 설정 시트의 자동 넘기기 체크 여부 */
function carryEnabled_() {
  return sheet_(CONFIG.SHEETS.SETTINGS).getRange(SETTINGS.CARRY.row, SETTINGS.CARRY.col).getValue() === true;
}

/**
 * 못 한 일 자동 넘기기 (매일 0시, 하루 한 번만)
 * 1) 전날 체크 안 된 할 일 → 전날 것은 "미완"으로 남기고(기록: FALSE + 이월), 오늘 할 일로 복사("넘어옴")
 *    개수 제한 없이 전부 넘기고, 시간·메모도 같이
 * 2) 그다음 (월요일이면) 새 주 시작 — 저장소에는 데이터가 남아 있으므로 순서 문제 없음
 */
function carryOver_() {
  const props = props_();
  const today = today_(), y = addDays_(today, -1), yKey = dateKey_(y);
  if (props.getProperty('CARRY_DONE') === yKey) return ''; // 이미 처리함
  if (!carryEnabled_()) {
    props.setProperty('CARRY_DONE', yKey);
    return '';
  }
  const yItems = getTodos_(y);
  const carried = [];
  yItems.forEach(function (t) {
    if (t.done || t.flag === CARRY_FLAG_.MISSED || !t.text) return;
    carried.push({ done: false, start: t.start, end: t.end, text: t.text, memo: t.memo, flag: CARRY_FLAG_.CARRIED });
    t.flag = CARRY_FLAG_.MISSED;
  });
  if (carried.length) {
    setTodos_(y, yItems);
    archiveDayLive_(y); // 전날 날짜로 FALSE + 이월
    setTodos_(today, getTodos_(today).concat(carried));
    archiveDayLive_(today);
    refreshSheetsFor_(y);
    refreshSheetsFor_(today);
  }
  ensureCurrentWeek_();
  props.setProperty('CARRY_DONE', yKey);
  const msg = carried.length ? '✓ ' + carried.length + '개 넘김' : '✓ 넘길 일 없음';
  setStatus_('carry', msg);
  return msg;
}
