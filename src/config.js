/**
 * config.js
 * 플래너 전체 설정값 모음 (줄 수, 시간, 색, 시트 이름, 레이아웃 위치)
 * 숫자만 바꾸면 되도록 다른 파일에서는 이 값을 참조한다.
 */

const CONFIG = {
  // 지역 설정
  TIMEZONE: 'Asia/Seoul',
  LOCALE: 'ko_KR',

  // 기본 테마색 (설정 시트에서 변경 가능)
  DEFAULT_THEME: '#2E7D32',

  // 시트 이름
  SHEETS: {
    WEEKLY: '위클리',
    TODAY: '오늘',
    MONTHLY: '먼슬리',
    LASTWEEK: '지난 주 보기',
    ARCHIVE: '기록',
    SETTINGS: '설정',
    EVENTS: '_일정', // 숨김: 캘린더 일정 + 달력 메모 저장소
  },

  // 개수
  TODO_ROWS: 10, // 요일별 할 일 줄 수 (고정)
  HABIT_COUNT: 10, // 습관 개수
  MUST_COUNT: 5, // 이번 주 꼭 끝낼 일
  SCHEDULE_LINES: 3, // 위클리·오늘의 일정 표시 줄 수
  MONTH_GOALS: 5, // 이번 달 목표

  // 기분 점수 (드롭다운 값과 이모지)
  MOODS: ['1 😞', '2 🙁', '3 😐', '4 🙂', '5 😄'],
  MOOD_EMOJI: ['😞', '🙁', '😐', '🙂', '😄'],

  // 기록 시트의 "구분" 값
  CATEGORY: {
    TODO: '할일',
    HABIT: '습관',
    WEEK_GOAL: '주간목표',
    MONTH_GOAL: '월간목표',
    MOOD: '기분',
  },
  ARCHIVE_HEADERS: ['날짜', '구분', '항목', '완료여부', '값', '주차'],

  // _일정 시트의 구분 값
  EVENT_TYPE: { CAL: '캘린더', MEMO: '메모' },

  // 트리거 시간
  TRIGGER: {
    NIGHTLY_HOUR: 23, // 매일 밤 기록 저장
    NIGHTLY_MINUTE: 50,
    MIDNIGHT_HOUR: 0, // 매일 0시: 오늘 시트 갱신 + 월요일이면 새 주
    MIDNIGHT_MINUTE: 5,
    SYNC_EVERY_HOURS: 1, // 캘린더 불러오기 주기
  },
  CALENDAR_MONTHS_AHEAD: 1, // 이번 달 + 몇 달 뒤까지 불러올지

  // 색
  COLORS: {
    TEXT: '#212121',
    SUB: '#9E9E9E', // 회색 글씨 (읽기 전용)
    LINE: '#E6E6E6', // 얇은 밑줄
    MOOD_LOW: '#EEEEEE', // 기분 낮을 때 배경
    WHITE: '#FFFFFF',
  },
  TINT: { LIGHT: 0.78, PALE: 0.9 }, // 테마색을 흰색과 섞는 비율

  // 크기 (px)
  SIZE: {
    ROW: 30, // 기본 행 높이 (폰에서 누르기 편하게)
    TODAY_ROW: 36,
    LABEL_COL: 70,
    CHECK_COL: 36,
    TEXT_COL: 115,
    GAP_COL: 16,
    HABIT_NAME_COL: 120,
    HABIT_DAY_COL: 34,
    GRID_COL: 34, // 먼슬리 격자 열
    CHECK_FONT: 13, // 체크박스 크기는 글자 크기를 따라감
  },
};

/** 위클리 시트 레이아웃 */
const WEEKLY = (function () {
  const L = {
    TITLE: { row: 1, col: 1 },
    WEEK_START: { row: 1, col: 4 }, // 주 시작 날짜 입력칸 (D1)
    ACTION_SAVE: { row: 2, col: 1 }, // 지금 저장 체크박스
    ACTION_SYNC: { row: 3, col: 1 }, // 캘린더 새로고침 체크박스
    ACTION_RESULT_COL: 4,
    INFO_LABEL_COL: 7,
    INFO_VALUE_COL: 8,
    INFO_VALUE_END: 15,
    GOAL_ROW: 1,
    REWARD_ROW: 2,
    AFFIRM_ROW: 3,
    MUST_HEADER: 5,
    MUST_START: 6,
    DAY_HEADER: 13,
    HABIT_NAME_COL: 17,
    HABIT_DAY_COL: 18,
    HABIT_BAR_COL: 25,
    HABIT_PCT_COL: 26,
    CHART_DATA_COL: 28, // 숨김 차트 데이터
  };
  L.MUST_PROGRESS = L.MUST_START + CONFIG.MUST_COUNT;
  L.OVERALL = { row: L.DAY_HEADER - 1, col: L.HABIT_NAME_COL };
  L.RATE = L.DAY_HEADER + 1;
  L.SCHED_START = L.RATE + 1;
  L.TODO_START = L.SCHED_START + CONFIG.SCHEDULE_LINES;
  L.COUNT = L.TODO_START + CONFIG.TODO_ROWS;
  L.MOOD = L.COUNT + 1;
  L.MOOD_MEMO = L.MOOD + 1;
  L.HABIT_START = L.DAY_HEADER + 1;
  L.LAST_COL = L.CHART_DATA_COL + 2;
  L.checkCol = function (d) { return 2 + 2 * d; }; // 요일 d(0=월)의 체크박스 열
  L.textCol = function (d) { return 3 + 2 * d; }; // 요일 d의 글자 열
  return L;
})();

/** 먼슬리 시트 레이아웃 (34px 격자) */
const MONTHLY = (function () {
  const L = {
    YEAR: { row: 1, col: 3 }, // C1
    MONTH: { row: 1, col: 7 }, // G1
    TITLE_COL: 9,
    WEEKDAY_ROW: 3,
    CAL_START: 4,
    WEEKS: 6,
    DAY_SPAN: 4, // 달력 하루 = 격자 4칸
    PANEL_COL: 30, // 오른쪽 패널 시작 (AD)
    PANEL_END: 41, // (AO)
    GOAL_HEADER: 3,
    GOAL_START: 4,
    ADD_HEADER: 10,
    ADD_LABEL: 11,
    ADD_ROW: 12,
    ADD_RESULT: 13,
    ADD: { date: 30, start: 33, end: 35, title: 37, check: 41 },
    ADD_SPAN: { date: 3, start: 2, end: 2, title: 4, check: 1 },
    SYNC_ROW: 15,
    HABIT_NAME_SPAN: 3,
    HABIT_DAY_COL: 4,
    HABIT_BAR_COL: 35,
    HABIT_PCT_COL: 37,
    CHART_DATA_COL: 44, // 숨김 차트 데이터 (AR~AT)
  };
  L.CAL_END = L.CAL_START + L.WEEKS * 3 - 1;
  L.HABIT_HEADER = L.CAL_END + 2;
  L.HABIT_START = L.HABIT_HEADER + 1;
  L.CHART_ROW = L.HABIT_START + CONFIG.HABIT_COUNT + 1;
  L.LAST_COL = L.CHART_DATA_COL + 2;
  L.dateRow = function (w) { return L.CAL_START + 3 * w; };
  L.eventRow = function (w) { return L.CAL_START + 3 * w + 1; };
  L.memoRow = function (w) { return L.CAL_START + 3 * w + 2; };
  L.dayCol = function (d) { return 1 + L.DAY_SPAN * d; };
  return L;
})();

/** 오늘 시트 레이아웃 (폰 세로, 3열) */
const TODAY = (function () {
  let r = 1;
  const L = { DATE_ROW: r };
  r += 2;
  L.SCHED_HEADER = r; L.SCHED_START = r + 1; r = L.SCHED_START + CONFIG.SCHEDULE_LINES + 1;
  L.TODO_HEADER = r; L.TODO_START = r + 1; r = L.TODO_START + CONFIG.TODO_ROWS + 1;
  L.HABIT_HEADER = r; L.HABIT_START = r + 1; r = L.HABIT_START + CONFIG.HABIT_COUNT + 1;
  L.MUST_HEADER = r; L.MUST_START = r + 1; r = L.MUST_START + CONFIG.MUST_COUNT + 1;
  L.MOOD_HEADER = r; L.MOOD = r + 1; L.MOOD_MEMO = r + 2; r += 4;
  L.ADD_HEADER = r;
  L.ADD = { date: r + 1, start: r + 2, end: r + 3, title: r + 4, check: r + 5, result: r + 6 };
  r += 8;
  L.TOOLS_HEADER = r; L.SYNC_ROW = r + 1; L.SAVE_ROW = r + 2;
  L.LAST_ROW = r + 3;
  return L;
})();

/** 지난 주 보기 레이아웃 */
const LASTWEEK = (function () {
  const L = {
    SELECT: { row: 1, col: 4 }, // 주차 드롭다운 (D1)
    HELPER_COL: 28, // 숨김: 주차 목록
    MONDAY: { row: 1, col: 29 }, // 숨김: 선택 주 월요일 날짜
    MUST_HEADER: 3,
    MUST_START: 4,
    HABIT_NAME_COL: 17,
    HABIT_DAY_COL: 18,
    HABIT_PCT_COL: 25,
  };
  L.DAY_HEADER = L.MUST_START + CONFIG.MUST_COUNT + 1;
  L.RATE = L.DAY_HEADER + 1;
  L.TODO_START = L.RATE + 1;
  L.COUNT = L.TODO_START + CONFIG.TODO_ROWS;
  L.MOOD = L.COUNT + 1;
  L.MOOD_MEMO = L.MOOD + 1;
  L.HABIT_START = L.DAY_HEADER + 1;
  L.LAST_COL = L.MONDAY.col;
  return L;
})();

/** 설정 시트 레이아웃 */
const SETTINGS = {
  THEME: { row: 3, col: 2 },
  HABIT_HEADER: 5,
  HABIT_START: 6,
  PRESET_ROW: 17,
  STATUS_HEADER: 19,
  STATUS: { setup: 20, save: 21, sync: 22, week: 23, error: 24 },
};
