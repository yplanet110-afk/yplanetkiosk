/**
 * 대관현황 조회 전용 read-only API
 * - 프론트(정적 페이지)가 CalendarApp을 직접 못 부르니까 이 GAS가 대신 캘린더를 읽어서 JSON으로 돌려줌
 * - 쓰기 기능 없음, 캘린더 ID도 프론트에 노출 안 됨 (여기 서버 코드에만 있음)
 * - doGet은 서버 코드라서 CSP 제약(var, addEventListener 등)은 해당 없음 → 최신 문법 그대로 사용
 * - GAS 웹앱은 GET 요청에 한해 구글이 자동으로 CORS 허용해줌 (수학챗봇 GAS 백엔드와 동일한 방식)
 */

// 활동실 ↔ 캘린더 ID 매핑 (노션 3-6 표 그대로). 방 추가·캘린더 변경 시 이 객체만 수정하면 됨.
const ROOM_CALENDARS = {
  '와이랩1': 'YOUR_CALENDAR_ID_1@group.calendar.google.com',
  '와이레코딩': 'YOUR_CALENDAR_ID_2@group.calendar.google.com',
  '와이스타': 'YOUR_CALENDAR_ID_3@group.calendar.google.com',
  '와이스테이지1': 'YOUR_CALENDAR_ID_4@group.calendar.google.com',
  '와이스테이지2': 'YOUR_CALENDAR_ID_5@group.calendar.google.com',
  '와이스튜디오': 'YOUR_CALENDAR_ID_6@group.calendar.google.com'
};

const TIMEZONE = 'Asia/Seoul';

function doGet(e) {
  const now = new Date();
  const todayStr = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');
  // 서울 자정~자정 범위로 오늘 일정만 조회
  const startOfDay = new Date(todayStr + 'T00:00:00+09:00');
  const endOfDay = new Date(todayStr + 'T23:59:59+09:00');

  const rooms = {};

  for (const roomName in ROOM_CALENDARS) {
    try {
      const cal = CalendarApp.getCalendarById(ROOM_CALENDARS[roomName]);
      const events = cal.getEvents(startOfDay, endOfDay);
      rooms[roomName] = events.map(ev => ({
        title: ev.getTitle(),
        start: Utilities.formatDate(ev.getStartTime(), TIMEZONE, 'HH:mm'),
        end: Utilities.formatDate(ev.getEndTime(), TIMEZONE, 'HH:mm')
      })).sort((a, b) => a.start.localeCompare(b.start));
    } catch (err) {
      // 캘린더 하나가 접근 실패해도 전체가 죽지 않게 개별 처리
      rooms[roomName] = { error: '캘린더 조회 실패' };
    }
  }

  const payload = { success: true, date: todayStr, rooms: rooms };

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * ============================================================
 * 일일 백업: Supabase → 구글 시트
 * ============================================================
 * - 무료 플랜은 백업이 없어서 이 트리거가 실질적인 백업 역할 (실수 복구용·통계 확인용)
 * - 매일 API 요청이 발생하므로 "1주 미사용 시 프로젝트 일시정지"도 자연스럽게 방지됨
 * - service_role key는 RLS를 우회하는 강력한 권한이라 반드시 Script Properties에 저장
 *   (코드에 직접 안 적음 — 이 파일이 실수로 공개 저장소에 올라가도 안전하게)
 *
 * ★ 최초 1회 설정 (Apps Script 편집기 좌측 톱니바퀴 → 프로젝트 설정 → 스크립트 속성):
 *   SUPABASE_URL              = https://fbszzjugwpetcpdxhzbn.supabase.co
 *   SUPABASE_SERVICE_ROLE_KEY = (Supabase 대시보드 → Settings → API → service_role key)
 *   BACKUP_SHEET_ID           = (백업용 구글 시트 URL의 /d/ 와 /edit 사이 ID)
 *
 * ★ 트리거 설정 (Apps Script 편집기 좌측 시계 아이콘 → 트리거 추가):
 *   실행할 함수: dailyBackup / 이벤트 소스: 시간 기반 / 매일 타이머: 오전 1시~2시
 *
 * ★ 보관 방식: 테이블마다 '{테이블명}_{날짜}' 탭을 매일 새로 만듦 (예: members_2026-07-31)
 *   RETENTION_DAYS(기본 7일)보다 오래된 탭은 실행할 때마다 자동 삭제됨
 */

// 백업 대상 테이블 전체 (노션 5-1 스키마 10개 그대로)
const BACKUP_TABLES = [
  'members', 'attendance', 'snack_logs', 'reservations', 'bonus_credits',
  'devices', 'operating_hours', 'settings', 'random_messages', 'nth_visitor_rules'
];
const RETENTION_DAYS = 7; // 이보다 오래된 날짜별 백업 탭은 자동 삭제

function dailyBackup() {
  const props = PropertiesService.getScriptProperties();
  const supabaseUrl = props.getProperty('SUPABASE_URL');
  const serviceKey = props.getProperty('SUPABASE_SERVICE_ROLE_KEY');
  const sheetId = props.getProperty('BACKUP_SHEET_ID');

  if (!supabaseUrl || !serviceKey || !sheetId) {
    console.error('스크립트 속성(SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY / BACKUP_SHEET_ID)이 설정 안 됨');
    return;
  }

  const ss = SpreadsheetApp.openById(sheetId);
  const todayStr = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  const results = [];

  BACKUP_TABLES.forEach(table => {
    try {
      const count = backupTable(ss, supabaseUrl, serviceKey, table, todayStr);
      results.push(`${table}: ${count}건`);
    } catch (err) {
      results.push(`${table}: 실패 (${err.message})`);
    }
  });

  purgeOldBackups(ss);
  writeBackupLog(ss, results);
}

// 테이블 하나를 Supabase REST API로 읽어와서 '{테이블명}_{날짜}' 탭에 저장 (날짜별로 따로 남김)
// Supabase REST API는 한 번 요청에 최대 1000행까지만 주므로, Range 헤더로 나눠서 전부 받아옴
function backupTable(ss, supabaseUrl, serviceKey, tableName, dateStr) {
  const PAGE_SIZE = 1000;
  let rows = [];
  let from = 0;

  while (true) {
    const url = `${supabaseUrl}/rest/v1/${tableName}?select=*`;
    const res = UrlFetchApp.fetch(url, {
      headers: {
        apikey: serviceKey,
        Authorization: `Bearer ${serviceKey}`,
        Range: `${from}-${from + PAGE_SIZE - 1}` // 예: 0-999, 1000-1999 ...
      },
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    if (code !== 200 && code !== 206) { // 206 Partial Content = 정상 (범위 요청 응답)
      throw new Error(`HTTP ${code}`);
    }

    const page = JSON.parse(res.getContentText());
    rows = rows.concat(page);

    if (page.length < PAGE_SIZE) break; // 다 받았으면 종료
    from += PAGE_SIZE;
  }

  const sheetName = `${tableName}_${dateStr}`; // 날짜별 탭 (예: members_2026-07-31)
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  else sheet.clearContents(); // 같은 날 여러 번 수동 실행했을 때 중복 방지

  if (rows.length === 0) {
    sheet.getRange(1, 1).setValue('(데이터 없음)');
    return 0;
  }

  // 컬럼 이름은 첫 행 기준 (테이블마다 다르므로 하드코딩하지 않고 동적으로 뽑음)
  const headers = Object.keys(rows[0]);
  const data = rows.map(row => headers.map(h => {
    const v = row[h];
    // jsonb·array 컬럼은 객체로 오므로 시트에 넣기 전에 문자열화
    return (v !== null && typeof v === 'object') ? JSON.stringify(v) : v;
  }));

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  return rows.length;
}

// RETENTION_DAYS보다 오래된 '{테이블명}_{날짜}' 탭을 찾아서 삭제 (백업로그처럼 날짜 패턴 없는 탭은 안 건드림)
function purgeOldBackups(ss) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - RETENTION_DAYS);
  const cutoffStr = Utilities.formatDate(cutoff, TIMEZONE, 'yyyy-MM-dd');

  ss.getSheets().forEach(sheet => {
    const match = sheet.getName().match(/^(.+)_(\d{4}-\d{2}-\d{2})$/);
    if (!match) return; // 날짜 패턴 없는 탭(백업로그 등)은 건너뜀
    const sheetDateStr = match[2]; // 'YYYY-MM-DD'는 문자열 비교로도 날짜 순서 비교 가능
    if (sheetDateStr < cutoffStr) {
      ss.deleteSheet(sheet);
    }
  });
}

// 백업이 매일 잘 돌고 있는지 확인용 로그 (실행시각 + 테이블별 건수/실패여부)
function writeBackupLog(ss, results) {
  let logSheet = ss.getSheetByName('백업로그');
  if (!logSheet) {
    logSheet = ss.insertSheet('백업로그');
    logSheet.getRange(1, 1, 1, 2).setValues([['실행시각', '결과']]);
  }
  const now = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  logSheet.appendRow([now, results.join(' / ')]);
}


// ===== 디버그용: 스크립트 속성이 제대로 저장됐는지 확인 (문제 해결되면 지워도 됨) =====
function checkScriptProperties() {
  const props = PropertiesService.getScriptProperties().getProperties();
  console.log('현재 저장된 스크립트 속성:', JSON.stringify(props, null, 2));
}
