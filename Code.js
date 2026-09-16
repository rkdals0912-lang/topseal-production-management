function doPost(e) {
  try {
    // 1. Streamlit 등에서 보낸 파싱 데이터 받기 (JSON 형태인 경우)
    var data = JSON.parse(e.postData.contents);

    // TODO: 전달받은 데이터 처리 로직 작성 (예: 시트에 기록, 데이터 조회 등)
    var resultValue = "처리 성공 데이터";

    // 2. 결과를 JSON 형식으로 다시 반환
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: resultValue
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // 에러 발생 시 응답
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

const LOG_SHEET = '3_출·퇴근기록등록';
const MASTER_SHEET = '작업자마스터';
const AUTH_SHEET = '권한마스터';
const APPROVAL_SHEET = '결재현황';

const VALID_TEAMS = [
  '포장반',
  '제조반A',
  '제조반B',
  'GMP/위험물',
  '품질검사',
  '설비팀'
];

const ADMIN_PASSWORD = '1384';
const TIMEZONE = 'Asia/Seoul';

/* =========================================================
   회사 위치 설정 - 반드시 실제 회사 좌표로 변경하세요.
   예) const COMPANY_LAT = 36.925160017235;
       const COMPANY_LNG = 127.56655810896;
========================================================= */
const COMPANY_LAT = 36.925160017235;
const COMPANY_LNG = 127.56655810896;
const ALLOWED_RADIUS_METERS = 300;
const MAX_GPS_ACCURACY_METERS = 150;

const REMARKS = ['', '연차', '오전반차', '오후반차', '교육', '공가', '휴직', '휴무'];

/* =========================================================
   화면 모드 설정 (관리자 On/Off)
   - Script Properties에 저장되어 모든 사용자/기기에 동일하게 적용됩니다.
   - 값이 없으면 기본값은 '꺼짐'(작업자 출·퇴근 모드 숨김) 입니다.
========================================================= */
const CONFIG_WORKER_MODE_KEY = 'WORKER_MODE_ENABLED';

function getFeatureConfig_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONFIG_WORKER_MODE_KEY);
  return {
    workerModeEnabled: raw === 'Y'
  };
}

function setWorkerModeEnabled(password, enabled) {
  checkAdmin_(password);
  const props = PropertiesService.getScriptProperties();
  props.setProperty(CONFIG_WORKER_MODE_KEY, enabled ? 'Y' : 'N');
  return getFeatureConfig_();
}

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('생산관리 POP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================================================
   최초 1회 설정
   - 기존 시트 유지
   - 권한마스터가 없으면 자동 생성
========================================================= */
function setupOnce() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TIMEZONE);

  const log = ss.getSheetByName(LOG_SHEET);
  const master = ss.getSheetByName(MASTER_SHEET);

  if (!log || !master) {
    throw new Error('3_출·퇴근기록등록 또는 작업자마스터 시트가 없습니다.');
  }

  log.setFrozenRows(1);
  master.setFrozenRows(1);

  // 작업자마스터 E열 제목 보장
  if (!String(master.getRange(1, 5).getDisplayValue() || '').trim()) {
    master.getRange(1, 5).setValue('휴대폰번호');
  }

  // K열 비고 / L~N 내부 관리열 보장
  if (!String(log.getRange(1, 11).getDisplayValue() || '').trim()) {
    log.getRange(1, 11).setValue('비고');
  }
  log.getRange(1, 12).setValue('_조회키');
  log.getRange(1, 13).setValue('_휴직시작일');
  log.getRange(1, 14).setValue('_휴직종료일');

  // O열 작업자명 제목 보장(이미 있으면 유지)
  if (!String(log.getRange(1, 15).getDisplayValue() || '').trim()) {
    log.getRange(1, 15).setValue('작업자명');
  }

  // 기존 출퇴근 데이터의 조회키를 1회 생성/보정
  rebuildAttendanceLookupKeys_();

  // 내부용 L~N열은 사용자 화면에서 숨김
  try { log.hideColumns(12, 3); } catch (e) {}

  let auth = ss.getSheetByName(AUTH_SHEET);
  if (!auth) {
    auth = ss.insertSheet(AUTH_SHEET);
    auth.getRange(1, 1, 1, 5).setValues([[
      '사용자명', '비밀번호', '역할', '담당반', '사용여부'
    ]]);
    auth.setFrozenRows(1);
    auth.getRange('A1:E1').setFontWeight('bold');
  }

  let approval = ss.getSheetByName(APPROVAL_SHEET);
  if (!approval) {
    approval = ss.insertSheet(APPROVAL_SHEET);
    approval.getRange(1, 1, 1, 9).setValues([[
      '작업일', '작업반', '작성자', '작성완료일시', '검토자', '검토완료일시', '승인자', '승인완료일시', '상태'
    ]]);
    approval.setFrozenRows(1);
    approval.getRange('A1:I1').setFontWeight('bold');
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    spreadsheet: ss.getName(),
    timezone: ss.getSpreadsheetTimeZone(),
    masterRows: Math.max(0, master.getLastRow() - 1),
    authSheet: AUTH_SHEET,
    approvalSheet: APPROVAL_SHEET,
    lookupColumn: 'L',
    suspensionColumns: 'M:N',
    locationConfigured: isCompanyLocationConfigured_()
  };
}

function ping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSeoulTimezone_();
  const master = ss.getSheetByName(MASTER_SHEET);

  return {
    ok: true,
    version: 'v23-leave-cleanup-fix',
    spreadsheet: ss.getName(),
    timezone: ss.getSpreadsheetTimeZone(),
    masterExists: !!master,
    masterRows: master ? Math.max(0, master.getLastRow() - 1) : 0,
    teams: VALID_TEAMS,
    remarks: REMARKS,
    location: {
      configured: isCompanyLocationConfigured_(),
      radiusMeters: ALLOWED_RADIUS_METERS,
      maxAccuracyMeters: MAX_GPS_ACCURACY_METERS
    },
    config: getFeatureConfig_()
  };
}

/* =========================================================
   작업자 조회
========================================================= */
function getWorkers(team) {
  validateTeam_(team);

  const sh = getMasterSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  return sh.getRange(2, 1, lastRow - 1, 4).getDisplayValues()
    .map(r => ({
      team: String(r[0] || '').trim(),
      empNo: String(r[1] || '').trim(),
      name: String(r[2] || '').trim(),
      active: String(r[3] || '').trim().toUpperCase()
    }))
    .filter(r => r.team === team && r.active === 'Y' && r.empNo && r.name)
    .map(r => ({ team: r.team, empNo: r.empNo, name: r.name }));
}

/* =========================================================
   작업자 휴대폰번호 인증
   - 작업자마스터 E열 휴대폰번호 기준
   - 등록된 번호 1건만 허용
========================================================= */
function verifyWorkerPhone(phone) {
  return getWorkerByPhone_(phone);
}

function getWorkerPunchStatusByPhone(phone) {
  const worker = getWorkerByPhone_(phone);
  return getWorkerPunchStatus(worker.team, worker.empNo);
}

function workerPunchByPhone(phone, action, latitude, longitude, accuracy) {
  const worker = getWorkerByPhone_(phone);
  return workerPunch(worker.team, worker.empNo, action, latitude, longitude, accuracy);
}

function getWorkerByPhone_(phone) {
  const normalized = normalizePhone_(phone);

  if (!normalized) {
    throw new Error('휴대폰번호를 입력하세요.');
  }

  if (!/^01\d{8,9}$/.test(normalized)) {
    throw new Error('휴대폰번호 형식을 확인하세요. 예: 01012345678');
  }

  const sh = getMasterSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('등록된 작업자가 없습니다.');

  const matches = sh.getRange(2, 1, lastRow - 1, 5).getDisplayValues()
    .map(r => ({
      team: String(r[0] || '').trim(),
      empNo: String(r[1] || '').trim(),
      name: String(r[2] || '').trim(),
      active: String(r[3] || '').trim().toUpperCase(),
      phone: normalizePhone_(r[4])
    }))
    .filter(r => r.active === 'Y' && r.phone === normalized && r.empNo && r.name && VALID_TEAMS.includes(r.team));

  if (!matches.length) {
    throw new Error('등록되지 않은 휴대폰번호입니다. 관리자에게 등록을 요청하세요.');
  }

  if (matches.length > 1) {
    throw new Error('같은 휴대폰번호가 여러 작업자에게 등록되어 있습니다. 관리자에게 확인하세요.');
  }

  const w = matches[0];
  return { team: w.team, empNo: w.empNo, name: w.name };
}

/* =========================================================
   작업자 POP - 현재 상태 조회
   - 오늘 기록 우선
   - 자정 이후 야간 퇴근을 위해 최근 36시간 내 미퇴근 기록도 허용
========================================================= */
function getWorkerPunchStatus(team, empNo) {
  validateTeam_(team);
  ensureSeoulTimezone_();

  const worker = getWorkers(team).find(w => w.empNo === String(empNo || '').trim());
  if (!worker) throw new Error('작업자 정보를 확인하세요.');

  const sh = getLogSheet_();
  const now = new Date();
  const today = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');
  const dates = recentDateKeys_(now, 3);
  const rowNumbers = [];

  dates.forEach(dateKey => {
    findAttendanceRowsByKey_(sh, buildAttendanceKey_(dateKey, worker.empNo))
      .forEach(r => rowNumbers.push(r));
  });

  const records = readAttendanceRows_(sh, [...new Set(rowNumbers)]);

  // 오늘이 휴직 기간으로 등록되어 있으면 출퇴근 버튼을 막는다.
  const suspensionRecord = records.find(item => {
    const r = item.values;
    const key = String(r[11] || '').trim();
    const remark = String(r[10] || '').trim();
    return key === buildAttendanceKey_(today, worker.empNo) && remark === '휴직';
  });

  if (suspensionRecord) {
    return {
      team: worker.team,
      empNo: worker.empNo,
      name: worker.name,
      inTime: '',
      outTime: '',
      status: 'SUSPENDED',
      remark: '휴직',
      workDate: today
    };
  }

  let todayRecord = null;
  let openRecord = null;
  const nowMs = now.getTime();
  const maxAge = 36 * 60 * 60 * 1000;

  records.forEach(item => {
    const r = item.values;
    const inDt = r[1];
    const outDt = r[6];
    if (!(inDt instanceof Date) || isNaN(inDt.getTime())) return;

    const dateKey = Utilities.formatDate(inDt, TIMEZONE, 'yyyy-MM-dd');
    const rec = { row: item.row, inDt: inDt, outDt: outDt, dateKey: dateKey };

    if (dateKey === today) {
      if (!todayRecord || inDt.getTime() > todayRecord.inDt.getTime()) todayRecord = rec;
    }

    if (!(outDt instanceof Date) && nowMs - inDt.getTime() <= maxAge) {
      if (!openRecord || inDt.getTime() > openRecord.inDt.getTime()) openRecord = rec;
    }
  });

  return buildPunchStatus_(worker, openRecord || todayRecord);
}

/* =========================================================
   작업자 POP - 출근/퇴근 기록
   action: IN / OUT
   서버에서 회사 거리 재검증
========================================================= */
function workerPunch(team, empNo, action, latitude, longitude, accuracy) {
  validateTeam_(team);
  ensureSeoulTimezone_();

  const worker = getWorkers(team).find(w => w.empNo === String(empNo || '').trim());
  if (!worker) throw new Error('작업자 정보를 확인하세요.');

  const act = String(action || '').toUpperCase();
  if (!['IN', 'OUT'].includes(act)) throw new Error('출퇴근 구분이 올바르지 않습니다.');

  const loc = validateCompanyLocation_(latitude, longitude, accuracy);
  const sh = getLogSheet_();
  const now = new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const today = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');
    const dates = recentDateKeys_(now, 3);
    const rowNumbers = [];

    dates.forEach(dateKey => {
      findAttendanceRowsByKey_(sh, buildAttendanceKey_(dateKey, worker.empNo))
        .forEach(r => rowNumbers.push(r));
    });

    const records = readAttendanceRows_(sh, [...new Set(rowNumbers)]);

    const suspensionToday = records.some(item => {
      const r = item.values;
      return String(r[11] || '').trim() === buildAttendanceKey_(today, worker.empNo) &&
        String(r[10] || '').trim() === '휴직';
    });

    if (suspensionToday) {
      throw new Error('현재 휴직 기간으로 등록되어 있어 출·퇴근할 수 없습니다.');
    }

    let todayRows = [];
    let openRecord = null;
    const maxAge = 36 * 60 * 60 * 1000;

    records.forEach(item => {
      const r = item.values;
      const inDt = r[1];
      const outDt = r[6];
      if (!(inDt instanceof Date) || isNaN(inDt.getTime())) return;

      const dateKey = Utilities.formatDate(inDt, TIMEZONE, 'yyyy-MM-dd');
      if (dateKey === today) todayRows.push(item.row);

      if (!(outDt instanceof Date) && now.getTime() - inDt.getTime() <= maxAge) {
        if (!openRecord || inDt.getTime() > openRecord.inDt.getTime()) {
          openRecord = { row: item.row, inDt: inDt };
        }
      }
    });

    if (act === 'IN') {
      if (openRecord || todayRows.length) {
        throw new Error('이미 출근 기록이 있습니다. 상태를 다시 확인하세요.');
      }

      const newRow = sh.getLastRow() + 1;
      sh.getRange(newRow, 1, 1, 12).setValues([[
        worker.empNo,
        now,
        '',
        formatLocationText_(loc),
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        buildAttendanceKey_(today, worker.empNo)
      ]]);
      sh.getRange(newRow, 1).setNumberFormat('@');
      sh.getRange(newRow, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    } else {
      if (!openRecord) throw new Error('퇴근할 수 있는 출근 기록이 없습니다.');

      sh.getRange(openRecord.row, 7).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm:ss');
      sh.getRange(openRecord.row, 9).setValue(formatLocationText_(loc));
    }

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return getWorkerPunchStatus(team, worker.empNo);
}

/* =========================================================
   작성자 모드 로그인
   권한마스터 역할은 WRITER 또는 작성자
========================================================= */
function recordLogin(username, password, team) {
  const user = findAuthUser_(username, password, ['WRITER', 'REVIEWER', 'APPROVER']);

  const labels = {
    WRITER: '작성',
    REVIEWER: '검토',
    APPROVER: '승인'
  };

  // 담당반이 ALL이면 로그인 성공 후 화면에서 작업반을 선택하게 한다.
  if (user.team === 'ALL') {
    return {
      ok: true,
      username: user.username,
      role: user.role,
      roleLabel: labels[user.role] || user.role,
      team: 'ALL',
      needTeamSelect: true,
      teams: VALID_TEAMS.slice()
    };
  }

  // 특정 작업반 권한은 로그인과 동시에 담당반으로 진입한다.
  validateTeam_(user.team);

  return {
    ok: true,
    username: user.username,
    role: user.role,
    roleLabel: labels[user.role] || user.role,
    team: user.team,
    needTeamSelect: false,
    teams: []
  };
}

// 구버전 Index.html 호환용
function writerLogin(username, password, team) {
  const session = recordLogin(username, password, team);
  if (session.role !== 'WRITER') {
    throw new Error('작성자 권한으로 로그인하세요.');
  }
  return session;
}

/* =========================================================
   작업일별 기존 근무기록 조회 - 작성자 화면
========================================================= */
function getAttendance(team, workDate) {
  validateTeam_(team);
  validateWorkDate_(workDate);
  ensureSeoulTimezone_();

  const workers = getWorkers(team);
  const workerMap = {};

  workers.forEach(w => {
    workerMap[w.empNo] = {
      team: w.team,
      empNo: w.empNo,
      name: w.name,
      inTime: '',
      outTime: '',
      inTimeDisplay: '',
      outTimeDisplay: '',
      remark: '',
      leaveStartDate: '',
      leaveEndDate: '',
      saved: false
    };
  });

  if (!workers.length) return [];

  const sh = getLogSheet_();
  const rowNumbers = findAttendanceRowsByDate_(sh, workDate);
  const records = readAttendanceRows_(sh, rowNumbers);

  records.forEach(item => {
    const r = item.values;
    const empNo = String(r[0] || '').trim();
    const inDt = r[1];
    const outDt = r[6];
    const remark = String(r[10] || '').trim();
    const leaveStartDate = normalizeWorkDateValue_(r[12]);
    const leaveEndDate = normalizeWorkDateValue_(r[13]);

    if (!workerMap[empNo]) return;

    // 휴직은 실제 출퇴근 시간이 없어도 날짜별 기록으로 표시한다.
    if (remark === '휴직') {
      workerMap[empNo].inTime = '';
      workerMap[empNo].outTime = '';
      workerMap[empNo].inTimeDisplay = '';
      workerMap[empNo].outTimeDisplay = '';
      workerMap[empNo].remark = '휴직';
      workerMap[empNo].leaveStartDate = leaveStartDate || workDate;
      workerMap[empNo].leaveEndDate = leaveEndDate || workDate;
      workerMap[empNo].saved = true;
      return;
    }

    // 휴무도 휴직과 동일하게 실제 출퇴근 시간 없이 하루 단위로 표시한다.
    if (remark === '휴무') {
      workerMap[empNo].inTime = '';
      workerMap[empNo].outTime = '';
      workerMap[empNo].inTimeDisplay = '';
      workerMap[empNo].outTimeDisplay = '';
      workerMap[empNo].remark = '휴무';
      workerMap[empNo].saved = true;
      return;
    }

    if (!(inDt instanceof Date) || isNaN(inDt.getTime())) return;

    // 같은 날짜/작업자가 과거에 중복되어 있어도 화면에는 가장 마지막 기록 1건만 표시
    workerMap[empNo].inTime = Utilities.formatDate(inDt, TIMEZONE, 'HH:mm');
    workerMap[empNo].inTimeDisplay = Utilities.formatDate(inDt, TIMEZONE, 'HH:mm:ss');

    if (outDt instanceof Date && !isNaN(outDt.getTime())) {
      workerMap[empNo].outTime = Utilities.formatDate(outDt, TIMEZONE, 'HH:mm');
      workerMap[empNo].outTimeDisplay = Utilities.formatDate(outDt, TIMEZONE, 'HH:mm:ss');
    } else {
      workerMap[empNo].outTime = '';
      workerMap[empNo].outTimeDisplay = '';
    }

    workerMap[empNo].remark = remark;
    workerMap[empNo].saved = true;
  });

  return workers.map(w => workerMap[w.empNo]);
}

/* =========================================================
   작성자 저장 / 기존값 수정
   - K열 비고
   - 기존 POP 출퇴근 기록 수정 가능

   성능 개선(v22):
   - 기존에는 저장 대상 항목(및 휴직 기간의 날짜 하나하나)마다
     createTextFinder로 L열 전체를 매번 스캔하여, 작업자 수 ×
     휴직일수가 커지면 서버 실행이 20초를 훌쩍 넘겨
     클라이언트 타임아웃("서버 응답이 없습니다")을 유발했습니다.
   - 이번 버전은 저장 시작 시 L열(조회키) 전체를 딱 1번만 읽어
     메모리 상의 Map(key -> [row,...])으로 인덱싱한 뒤,
     그 메모리 인덱스에서만 조회/갱신합니다.
   - 삭제 예정 행 번호가 바뀌지 않도록 실제 시트 삭제는
     맨 마지막에 한 번에 처리합니다(기존과 동일).

   버그 수정(v23):
   - 기존에는 휴직 기간을 줄였을 때, 화면(클라이언트)에서 넘어온
     originalLeaveStartDate/originalLeaveEndDate 문자열이 시트의
     실제 값과 정확히 일치할 때만 "기간 밖으로 벗어난 날짜의
     자동 휴직 행"을 삭제했습니다.
     클라이언트 값이 비어있거나(캐시/재조회 누락 등) 시트 값과
     형식이 다르면 이 삭제가 전혀 일어나지 않아, 휴직을 단축해도
     이전에 등록된 이후 날짜의 휴직 행이 시트에 그대로 남아있고,
     그 날짜에 출근을 시도하면 계속 "휴직 기간"으로 판정되어
     출근이 막히는 문제가 있었습니다.
   - 이제는 클라이언트가 보낸 원래 기간 값에 의존하지 않고,
     저장 시작 시 시트 전체에서 "해당 사원번호 + 비고=휴직"인
     행을 전부 한 번에 읽어 서버 자체 인덱스(empNo -> 행 목록)를
     만들고, 이번에 새로 지정한 휴직 기간(newDateSet)에 포함되지
     않는 기존 휴직 행은 무조건 삭제 대상으로 처리합니다.
========================================================= */
function saveAttendance(team, workDate, rows) {
  validateTeam_(team);
  validateWorkDate_(workDate);
  ensureSeoulTimezone_();

  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error('저장할 작업자를 선택하세요.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sh = getLogSheet_();
    const workerList = getWorkers(team);
    const validEmpNos = new Set(workerList.map(w => w.empNo));
    const nameMap = {};
    workerList.forEach(w => { nameMap[w.empNo] = w.name; });
    const rowsToDelete = [];
    let inserted = 0;
    let updated = 0;
    let suspensionDays = 0;
    let skipped = 0;
    const skippedNames = [];

    // ---- 성능 핵심: L열 조회키 전체를 1회만 읽어 메모리 인덱스 구성 ----
    const keyIndex = buildKeyIndex_(sh);

    // ---- 휴직 정리용: 사원번호별 기존 휴직 행 전체를 1회만 읽어 인덱스 구성 ----
    // 클라이언트가 보내는 원래 휴직 기간 값에 의존하지 않고,
    // 서버가 직접 "이 사원번호의 모든 기존 휴직 행"을 알고 있도록 한다.
    const leaveIndexByEmp = buildLeaveRowIndexByEmp_(sh);

    // 이번 저장 트랜잭션 중 새로 추가되는 행들도 keyIndex에 즉시 반영해서
    // 같은 호출 내에서 중복 삽입되지 않도록 한다.
    function indexAdd_(key, rowNo) {
      if (!keyIndex[key]) keyIndex[key] = [];
      keyIndex[key].push(rowNo);
    }
    function indexLookup_(key) {
      return keyIndex[key] ? keyIndex[key].slice() : [];
    }

    // 여러 행을 한 번에 읽기 위한 캐시(휴직 갱신 시 후보 행 판정용)
    const rowCache = {};
    function readCachedRow_(rowNo) {
      if (rowCache[rowNo]) return rowCache[rowNo];
      const values = sh.getRange(rowNo, 1, 1, 14).getValues()[0];
      rowCache[rowNo] = values;
      return values;
    }
    function invalidateCachedRow_(rowNo) {
      delete rowCache[rowNo];
    }

    function chooseKeeperRowNo_(rowNumbers) {
      if (!rowNumbers.length) throw new Error('기존 출퇴근 기록을 찾을 수 없습니다.');
      const withValues = rowNumbers.map(rowNo => ({ row: rowNo, values: readCachedRow_(rowNo) }));
      return chooseKeeperAttendanceRow_(withValues).row;
    }

    rows.forEach((r, idx) => {
      const empNo = String(r.empNo || '').trim();
      if (!empNo || !validEmpNos.has(empNo)) {
        throw new Error((idx + 1) + '번째 작업자의 사원번호/작업반을 확인하세요.');
      }

      const remark = validateRemark_(r.remark);

      // -------------------------------------------------
      // 휴직: 시작일~종료일까지 날짜별로 자동 생성/수정
      // -------------------------------------------------
      if (remark === '휴직') {
        const leaveStartDate = String(r.leaveStartDate || workDate || '').trim();
        const leaveEndDate = String(r.leaveEndDate || '').trim();

        validateWorkDate_(leaveStartDate);
        validateWorkDate_(leaveEndDate);

        if (leaveEndDate < leaveStartDate) {
          throw new Error((r.name || empNo) + '의 휴직 종료일은 시작일보다 빠를 수 없습니다.');
        }

        const leaveDates = enumerateDateRange_(leaveStartDate, leaveEndDate, 1826);
        const newDateSet = new Set(leaveDates);
        suspensionDays += leaveDates.length;

        leaveDates.forEach(dateKey => {
          const key = buildAttendanceKey_(dateKey, empNo);
          const matchingRows = indexLookup_(key);

          if (matchingRows.length) {
            const targetRow = chooseKeeperRowNo_(matchingRows);

            sh.getRange(targetRow, 1).setValue(empNo).setNumberFormat('@');
            sh.getRange(targetRow, 2).clearContent();
            sh.getRange(targetRow, 4).clearContent();
            sh.getRange(targetRow, 7).clearContent();
            sh.getRange(targetRow, 9).clearContent();
            sh.getRange(targetRow, 11).setValue('휴직');
            sh.getRange(targetRow, 12).setValue(key);
            sh.getRange(targetRow, 13).setValue(leaveStartDate).setNumberFormat('@');
            sh.getRange(targetRow, 14).setValue(leaveEndDate).setNumberFormat('@');
            sh.getRange(targetRow, 15).setValue(nameMap[empNo] || r.name || '');
            invalidateCachedRow_(targetRow);

            matchingRows.forEach(rowNo => {
              if (rowNo !== targetRow) rowsToDelete.push(rowNo);
            });
            updated++;
          } else {
            const newRow = sh.getLastRow() + 1;
            sh.getRange(newRow, 1, 1, 15).setValues([[
              empNo, '', '', '', '', '', '', '', '', '', '휴직', key,
              leaveStartDate, leaveEndDate, nameMap[empNo] || r.name || ''
            ]]);
            sh.getRange(newRow, 1).setNumberFormat('@');
            sh.getRange(newRow, 12, 1, 3).setNumberFormat('@');
            indexAdd_(key, newRow);
            inserted++;
          }
        });

        // ---- 휴직 기간을 줄이거나 옮겼을 경우, 새 기간 밖에 남아있는
        //      이 사원번호의 기존 휴직 행을 전부 정리한다. ----
        // 클라이언트가 보낸 원래 기간 값에 의존하지 않고,
        // 저장 시작 시 서버가 직접 읽어둔 leaveIndexByEmp를 사용한다.
        const existingLeaveForEmp = leaveIndexByEmp[empNo] || [];
        existingLeaveForEmp.forEach(item => {
          if (!newDateSet.has(item.dateKey)) {
            rowsToDelete.push(item.row);
          }
        });

        return;
      }

      // -------------------------------------------------
      // 휴무: 휴직처럼 시간 없이 해당 작업일 하루만 공란으로 저장
      // -------------------------------------------------
      if (remark === '휴무') {
        const key = buildAttendanceKey_(workDate, empNo);
        const matchingRows = indexLookup_(key);

        if (matchingRows.length) {
          const targetRow = chooseKeeperRowNo_(matchingRows);

          sh.getRange(targetRow, 1).setValue(empNo).setNumberFormat('@');
          sh.getRange(targetRow, 2).clearContent();
          sh.getRange(targetRow, 4).clearContent();
          sh.getRange(targetRow, 7).clearContent();
          sh.getRange(targetRow, 9).clearContent();
          sh.getRange(targetRow, 11).setValue('휴무');
          sh.getRange(targetRow, 12).setValue(key);
          sh.getRange(targetRow, 13, 1, 2).clearContent();
          sh.getRange(targetRow, 15).setValue(nameMap[empNo] || r.name || '');
          invalidateCachedRow_(targetRow);

          matchingRows.forEach(rowNo => {
            if (rowNo !== targetRow) rowsToDelete.push(rowNo);
          });
          updated++;
        } else {
          const newRow = sh.getLastRow() + 1;
          sh.getRange(newRow, 1, 1, 15).setValues([[
            empNo, '', '', '', '', '', '', '', '', '', '휴무', key, '', '', nameMap[empNo] || r.name || ''
          ]]);
          sh.getRange(newRow, 1).setNumberFormat('@');
          sh.getRange(newRow, 12).setNumberFormat('@');
          indexAdd_(key, newRow);
          inserted++;
        }

        // 휴무로 바뀐 작업자에 대해서도, 이 작업일에 남아있는 기존 휴직 행이
        // 있다면 함께 정리한다(예: 휴직 → 휴무로 비고를 바꾼 경우).
        const existingLeaveForEmp = leaveIndexByEmp[empNo] || [];
        existingLeaveForEmp.forEach(item => {
          if (item.dateKey === workDate) {
            rowsToDelete.push(item.row);
          }
        });

        return;
      }

      // -------------------------------------------------
      // 일반 / 연차 / 반차 / 교육 / 공가
      // -------------------------------------------------
      const key = buildAttendanceKey_(workDate, empNo);
      const isAnnualLeave = remark === '연차';
      const inTimeValue = isAnnualLeave ? '08:00' : r.inTime;
      const outTimeValue = isAnnualLeave ? '17:00' : r.outTime;

      // 비고를 휴직/휴무에서 다른 상태로 바꾼 경우를 대비해,
      // 이 작업일에 남아있는 기존 휴직 행이 있으면 함께 정리한다.
      const existingLeaveForEmp = leaveIndexByEmp[empNo] || [];
      existingLeaveForEmp.forEach(item => {
        if (item.dateKey === workDate) {
          rowsToDelete.push(item.row);
        }
      });

      if (!inTimeValue || !outTimeValue) {
        // 출퇴근시간 미입력자는 저장을 막지 않고, 대신 기록을 생성/유지하지 않는다.
        // 기존에 저장된 기록이 있다면 함께 삭제해 미입력 상태와 일치시킨다.
        const existingRows = indexLookup_(key);
        existingRows.forEach(rowNo => rowsToDelete.push(rowNo));
        skipped++;
        skippedNames.push(r.name || empNo);
        return;
      }

      const inDt = parseWorkDateTime_(workDate, inTimeValue);
      let outDt = parseWorkDateTime_(workDate, outTimeValue);

      if (!inDt || !outDt) {
        throw new Error((r.name || empNo) + '의 시간을 확인하세요.');
      }

      if (outDt.getTime() <= inDt.getTime()) {
        outDt = new Date(outDt.getTime() + 24 * 60 * 60 * 1000);
      }

      const matchingRows = indexLookup_(key);

      if (matchingRows.length) {
        const targetRow = chooseKeeperRowNo_(matchingRows);

        sh.getRange(targetRow, 1).setValue(empNo).setNumberFormat('@');
        sh.getRange(targetRow, 2).setValue(inDt).setNumberFormat('yyyy-mm-dd hh:mm');
        sh.getRange(targetRow, 7).setValue(outDt).setNumberFormat('yyyy-mm-dd hh:mm');

        if (isAnnualLeave) {
          sh.getRange(targetRow, 4).clearContent();
          sh.getRange(targetRow, 9).clearContent();
        }

        sh.getRange(targetRow, 11).setValue(remark);
        sh.getRange(targetRow, 12).setValue(key);
        sh.getRange(targetRow, 13, 1, 2).clearContent();
        sh.getRange(targetRow, 15).setValue(nameMap[empNo] || r.name || '');
        invalidateCachedRow_(targetRow);

        matchingRows.forEach(rowNo => {
          if (rowNo !== targetRow) rowsToDelete.push(rowNo);
        });
        updated++;
      } else {
        const newRow = sh.getLastRow() + 1;
        sh.getRange(newRow, 1, 1, 15).setValues([[
          empNo, inDt, '', '', '', '', outDt, '', '', '', remark, key, '', '', nameMap[empNo] || r.name || ''
        ]]);
        sh.getRange(newRow, 1).setNumberFormat('@');
        sh.getRange(newRow, 2).setNumberFormat('yyyy-mm-dd hh:mm');
        sh.getRange(newRow, 7).setNumberFormat('yyyy-mm-dd hh:mm');
        indexAdd_(key, newRow);
        inserted++;
      }
    });

    const uniqueRows = [...new Set(rowsToDelete)].sort((a, b) => b - a);
    uniqueRows.forEach(rowNo => sh.deleteRow(rowNo));

    SpreadsheetApp.flush();
    return {
      count: inserted + updated,
      inserted: inserted,
      updated: updated,
      suspensionDays: suspensionDays,
      duplicatesRemoved: uniqueRows.length,
      skipped: skipped,
      skippedNames: skippedNames
    };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
   출퇴근 기록 화면 - 작성/검토/승인 상태 조회
========================================================= */
function getApprovalState(team, workDate) {
  validateTeam_(team);
  validateWorkDate_(workDate);
  ensureSeoulTimezone_();

  const sh = getApprovalSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return emptyApprovalState_(team, workDate);

  const values = sh.getRange(2, 1, lastRow - 1, 9).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i];
    const d = normalizeApprovalDate_(r[0]);
    const t = String(r[1] || '').trim();
    if (d === workDate && t === team) {
      return {
        workDate: workDate,
        team: team,
        writer: String(r[2] || ''),
        writerAt: formatApprovalDateTime_(r[3]),
        reviewer: String(r[4] || ''),
        reviewerAt: formatApprovalDateTime_(r[5]),
        approver: String(r[6] || ''),
        approverAt: formatApprovalDateTime_(r[7]),
        status: String(r[8] || '작성중')
      };
    }
  }
  return emptyApprovalState_(team, workDate);
}

/* =========================================================
   [신규 추가] 공휴일 판정
   - 매년 날짜가 고정인 공휴일만 자동 반영합니다.
   - 설날/추석 등 음력 공휴일은 연도마다 달라지므로
     LUNAR_HOLIDAYS 객체에 'yyyy-MM-dd' 형태로 직접 추가해서 사용하세요.
   예) '2026-09-24': '추석 연휴', '2026-09-25': '추석', '2026-09-26': '추석 연휴'
========================================================= */
const FIXED_HOLIDAYS_MMDD = ['01-01', '03-01', '05-05', '06-06', '08-15', '10-03', '10-09', '12-25'];
const LUNAR_HOLIDAYS = {
  // 'yyyy-MM-dd': '공휴일명'
};

function isHoliday_(dateKey) {
  const mmdd = String(dateKey || '').slice(5);
  if (FIXED_HOLIDAYS_MMDD.includes(mmdd)) return true;
  if (LUNAR_HOLIDAYS[dateKey]) return true;
  return false;
}

/* =========================================================
   [신규 추가] 월별 근태 등록 캘린더 조회 - 작성/검토/승인 권한자 전용
   - 작업반의 작업자별로 해당 월 1일~말일까지 출퇴근 등록 여부를 조회한다.
   - 휴직/휴무는 실제 시간이 없어도 등록완료로 처리한다.
   - 미래 날짜는 아직 근무 전이므로 미등록 판정에서 제외한다.
========================================================= */
function getMonthlyAttendanceCalendar(username, password, team, yearMonth) {
  validateTeam_(team);
  ensureSeoulTimezone_();

  const user = findAuthUser_(username, password, ['WRITER', 'REVIEWER', 'APPROVER']);
  if (user.team !== 'ALL' && user.team !== team) {
    throw new Error('해당 작업반 권한이 없습니다.');
  }

  if (!/^\d{4}-\d{2}$/.test(String(yearMonth || ''))) {
    throw new Error('조회 월을 선택하세요.');
  }

  const parts = yearMonth.split('-');
  const yearStr = parts[0];
  const monthStr = parts[1];
  const year = Number(yearStr);
  const month = Number(monthStr);
  const daysInMonth = new Date(year, month, 0).getDate();

  const workers = getWorkers(team);
  const sh = getLogSheet_();

  // 사원번호별 '일' -> 등록여부 인덱스
  const dayMap = {};
  workers.forEach(w => { dayMap[w.empNo] = {}; });

  for (let day = 1; day <= daysInMonth; day++) {
    const dateKey = yearStr + '-' + monthStr + '-' + String(day).padStart(2, '0');
    const rowNumbers = findAttendanceRowsByDate_(sh, dateKey);
    if (!rowNumbers.length) continue;

    const records = readAttendanceRows_(sh, rowNumbers);
    records.forEach(item => {
      const r = item.values;
      const empNo = String(r[0] || '').trim();
      if (!dayMap[empNo]) return;

      const remark = String(r[10] || '').trim();
      const inDt = r[1];
      const outDt = r[6];
      let written = false;

      if (remark === '휴직' || remark === '휴무') {
        written = true;
      } else if (inDt instanceof Date && !isNaN(inDt.getTime()) &&
                 outDt instanceof Date && !isNaN(outDt.getTime())) {
        written = true;
      }

      dayMap[empNo][day] = written;
    });
  }

  const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');

  const resultWorkers = workers.map(w => {
    const days = [];
    for (let day = 1; day <= daysInMonth; day++) {
      const dateKey = yearStr + '-' + monthStr + '-' + String(day).padStart(2, '0');
      days.push({
        day: day,
        workDate: dateKey,
        written: !!dayMap[w.empNo][day],
        isFuture: dateKey > today
      });
    }
    return { empNo: w.empNo, name: w.name, days: days };
  });

  const dayInfo = [];
  for (let day = 1; day <= daysInMonth; day++) {
    const dateKey = yearStr + '-' + monthStr + '-' + String(day).padStart(2, '0');
    const dow = new Date(dateKey + 'T00:00:00+09:00').getDay();
    dayInfo.push({ day: day, dow: dow, holiday: isHoliday_(dateKey) });
  }

  return { year: year, month: month, daysInMonth: daysInMonth, days: dayInfo, workers: resultWorkers };
}


/* =========================================================
   미처리 검토/승인 알림
   - REVIEWER: 작성완료 건
   - APPROVER: 검토완료 건
   - 선택한 담당반 기준 최근 60일
========================================================= */
function getPendingApprovals(username, password, team) {
  validateTeam_(team);
  ensureSeoulTimezone_();

  const user = findAuthUser_(username, password, ['WRITER', 'REVIEWER', 'APPROVER']);

  if (user.team !== 'ALL' && user.team !== team) {
    throw new Error('해당 작업반 권한이 없습니다.');
  }

  if (user.role === 'WRITER') return [];

  const wantedStatus = user.role === 'REVIEWER' ? '작성완료' : '검토완료';
  const sh = getApprovalSheet_();
  const lastRow = sh.getLastRow();

  if (lastRow < 2) return [];

  const today = new Date();
  const cutoff = new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000);
  const values = sh.getRange(2, 1, lastRow - 1, 9).getValues();
  const items = [];

  values.forEach(r => {
    const workDate = normalizeApprovalDate_(r[0]);
    const rowTeam = String(r[1] || '').trim();
    const status = String(r[8] || '작성중').replace(/\s+/g, '').trim();

    if (!workDate || rowTeam !== team || status !== wantedStatus) return;

    const dateObj = new Date(workDate + 'T00:00:00+09:00');
    if (!isNaN(dateObj.getTime()) && dateObj < cutoff) return;

    items.push({
      workDate: workDate,
      team: rowTeam,
      status: status,
      message: user.role === 'REVIEWER' ? '검토 필요' : '승인 필요'
    });
  });

  items.sort((a, b) => a.workDate.localeCompare(b.workDate));

  return items.slice(0, 20);
}

/* =========================================================
   출퇴근 기록 화면 - 작성완료 / 검토완료 / 최종승인
========================================================= */
function updateApprovalState(username, password, team, workDate, action) {
  validateTeam_(team);
  validateWorkDate_(workDate);
  ensureSeoulTimezone_();

  const user = findAuthUser_(username, password, ['WRITER', 'REVIEWER', 'APPROVER']);
  if (user.team !== 'ALL' && user.team !== team) {
    throw new Error('해당 작업반 권한이 없습니다.');
  }

  const allowed = {
    WRITE_COMPLETE: 'WRITER',
    REVIEW_COMPLETE: 'REVIEWER',
    APPROVE_COMPLETE: 'APPROVER'
  };
  const requiredRole = allowed[String(action || '')];
  if (!requiredRole || user.role !== requiredRole) {
    throw new Error('현재 계정으로 처리할 수 없는 단계입니다.');
  }

  const sh = getApprovalSheet_();
  const lastRow = sh.getLastRow();
  const values = lastRow >= 2 ? sh.getRange(2, 1, lastRow - 1, 9).getValues() : [];
  let targetRow = 0;
  let row = [workDate, team, '', '', '', '', '', '', '작성중'];

  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i];
    if (normalizeApprovalDate_(r[0]) === workDate && String(r[1] || '').trim() === team) {
      targetRow = i + 2;
      row = r.slice(0, 9);
      break;
    }
  }

  const currentStatus = String(row[8] || '작성중');
  const now = new Date();

  if (action === 'WRITE_COMPLETE') {
    if (currentStatus === '승인완료') throw new Error('이미 최종 승인된 기록입니다.');
    row[0] = workDate;
    row[1] = team;
    row[2] = user.username;
    row[3] = now;
    row[4] = '';
    row[5] = '';
    row[6] = '';
    row[7] = '';
    row[8] = '작성완료';
  } else if (action === 'REVIEW_COMPLETE') {
    if (currentStatus !== '작성완료') throw new Error('작성완료 상태에서만 검토완료할 수 있습니다.');
    row[4] = user.username;
    row[5] = now;
    row[6] = '';
    row[7] = '';
    row[8] = '검토완료';
  } else if (action === 'APPROVE_COMPLETE') {
    if (currentStatus !== '검토완료') throw new Error('검토완료 상태에서만 최종승인할 수 있습니다.');
    row[6] = user.username;
    row[7] = now;
    row[8] = '승인완료';
  }

  if (targetRow) {
    sh.getRange(targetRow, 1, 1, 9).setValues([row]);
  } else {
    sh.appendRow(row);
    targetRow = sh.getLastRow();
  }

  sh.getRange(targetRow, 4).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange(targetRow, 6).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange(targetRow, 8).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  SpreadsheetApp.flush();

  return getApprovalState(team, workDate);
}

function getApprovalSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(APPROVAL_SHEET);
  if (!sh) {
    sh = ss.insertSheet(APPROVAL_SHEET);
    sh.getRange(1, 1, 1, 9).setValues([[
      '작업일', '작업반', '작성자', '작성완료일시', '검토자', '검토완료일시', '승인자', '승인완료일시', '상태'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange('A1:I1').setFontWeight('bold');
  }
  return sh;
}

function emptyApprovalState_(team, workDate) {
  return {
    workDate: workDate,
    team: team,
    writer: '', writerAt: '',
    reviewer: '', reviewerAt: '',
    approver: '', approverAt: '',
    status: '작성중'
  };
}

function normalizeApprovalDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TIMEZONE, 'yyyy-MM-dd');
  }
  return String(value || '').trim().slice(0, 10);
}

function formatApprovalDateTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value || '').trim();
}

/* =========================================================
   관리자
========================================================= */
function adminLogin(password) {
  checkAdmin_(password);
  return { ok: true };
}

function getMasterForAdmin(password) {
  checkAdmin_(password);
  const sh = getMasterSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  return sh.getRange(2, 1, lastRow - 1, 5).getDisplayValues().map((r, i) => ({
    row: i + 2,
    team: String(r[0] || '').trim(),
    empNo: String(r[1] || '').trim(),
    name: String(r[2] || '').trim(),
    active: String(r[3] || 'Y').trim().toUpperCase() === 'N' ? 'N' : 'Y',
    phone: normalizePhone_(r[4])
  }));
}

function saveMasterForAdmin(password, rows) {
  checkAdmin_(password);
  if (!Array.isArray(rows)) throw new Error('작업자 목록 형식이 올바르지 않습니다.');

  const cleaned = [];
  const seen = new Set();
  const seenPhones = new Set();

  rows.forEach((r, idx) => {
    const team = String(r.team || '').trim();
    const empNo = String(r.empNo || '').trim();
    const name = String(r.name || '').trim();
    const active = String(r.active || 'Y').trim().toUpperCase() === 'N' ? 'N' : 'Y';
    const phone = normalizePhone_(r.phone);

    if (!team && !empNo && !name && !phone) return;
    if (!VALID_TEAMS.includes(team)) throw new Error((idx + 1) + '번째 행의 작업반을 확인하세요.');
    if (!empNo) throw new Error((idx + 1) + '번째 행의 사원번호를 입력하세요.');
    if (!name) throw new Error((idx + 1) + '번째 행의 작업자명을 입력하세요.');
    if (seen.has(empNo)) throw new Error('사원번호가 중복되었습니다: ' + empNo);
    if (phone && !/^01\d{8,9}$/.test(phone)) throw new Error(name + '의 휴대폰번호를 확인하세요.');
    if (phone && seenPhones.has(phone)) throw new Error('휴대폰번호가 중복되었습니다: ' + phone);

    seen.add(empNo);
    if (phone) seenPhones.add(phone);
    cleaned.push([team, empNo, name, active, phone]);
  });

  const sh = getMasterSheet_();
  const existingRows = Math.max(0, sh.getLastRow() - 1);
  if (existingRows) sh.getRange(2, 1, existingRows, 5).clearContent();

  if (cleaned.length) {
    sh.getRange(2, 1, cleaned.length, 5).setValues(cleaned);
    sh.getRange(2, 2, cleaned.length, 1).setNumberFormat('@');
    sh.getRange(2, 5, cleaned.length, 1).setNumberFormat('@');
  }

  SpreadsheetApp.flush();
  return { ok: true, count: cleaned.length };
}

/* =========================================================
   권한마스터 관리자 조회/저장
   역할: WRITER / REVIEWER / APPROVER
========================================================= */
function getAuthForAdmin(password) {
  checkAdmin_(password);
  const sh = getAuthSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  return sh.getRange(2, 1, lastRow - 1, 5).getDisplayValues().map((r, i) => ({
    row: i + 2,
    username: String(r[0] || '').trim(),
    password: String(r[1] || '').trim(),
    role: normalizeRole_(r[2]),
    team: String(r[3] || 'ALL').trim() || 'ALL',
    active: String(r[4] || 'Y').trim().toUpperCase() === 'N' ? 'N' : 'Y'
  }));
}

function saveAuthForAdmin(password, rows) {
  checkAdmin_(password);
  if (!Array.isArray(rows)) throw new Error('권한 목록 형식이 올바르지 않습니다.');

  const cleaned = [];
  rows.forEach((r, idx) => {
    const username = String(r.username || '').trim();
    const pw = String(r.password || '').trim();
    const role = normalizeRole_(r.role);
    const team = String(r.team || 'ALL').trim() || 'ALL';
    const active = String(r.active || 'Y').trim().toUpperCase() === 'N' ? 'N' : 'Y';

    if (!username && !pw) return;
    if (!username || !pw) throw new Error((idx + 1) + '번째 권한 사용자의 이름/비밀번호를 입력하세요.');
    if (!['WRITER', 'REVIEWER', 'APPROVER'].includes(role)) throw new Error((idx + 1) + '번째 역할을 확인하세요.');
    if (team !== 'ALL' && !VALID_TEAMS.includes(team)) throw new Error((idx + 1) + '번째 담당반을 확인하세요.');

    cleaned.push([username, pw, role, team, active]);
  });

  const sh = getAuthSheet_();
  const existingRows = Math.max(0, sh.getLastRow() - 1);
  if (existingRows) sh.getRange(2, 1, existingRows, 5).clearContent();
  if (cleaned.length) sh.getRange(2, 1, cleaned.length, 5).setValues(cleaned);

  SpreadsheetApp.flush();
  return { ok: true, count: cleaned.length };
}

/* =========================================================
   엑셀 파일 생성
========================================================= */
function exportExcel(workDate) {
  validateWorkDate_(workDate);
  ensureSeoulTimezone_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();

  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('엑셀 파일 생성에 실패했습니다. HTTP ' + response.getResponseCode());
  }

  const safeName = ss.getName().replace(/[\\/:*?"<>|]/g, '_');
  return {
    filename: safeName + '_' + workDate + '.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    base64: Utilities.base64Encode(response.getBlob().getBytes())
  };
}


/* =========================================================
   출퇴근 고속 조회용 내부 인덱스
   - L열 _조회키 = yyyy-MM-dd|사원번호
   - 사용자에게는 숨겨진 내부용 열
========================================================= */
function buildAttendanceKey_(workDate, empNo) {
  return String(workDate || '').trim() + '|' + String(empNo || '').trim();
}

// L열(조회키) 전체를 1회만 읽어 key -> [row번호,...] 형태의 메모리 인덱스로 구성.
// saveAttendance처럼 같은 실행 내에서 반복 조회가 필요한 경우에 사용해서
// createTextFinder 반복 호출로 인한 실행시간 초과를 방지한다.
function buildKeyIndex_(sh) {
  const lastRow = sh.getLastRow();
  const index = {};
  if (lastRow < 2) return index;

  const keyValues = sh.getRange(2, 12, lastRow - 1, 1).getValues();
  keyValues.forEach((r, i) => {
    const k = String(r[0] || '').trim();
    if (!k) return;
    const rowNo = i + 2;
    if (!index[k]) index[k] = [];
    index[k].push(rowNo);
  });

  return index;
}

// 사원번호별 "비고=휴직" 행 전체를 1회만 읽어 인덱스로 구성.
// { empNo: [ { row, key, dateKey }, ... ], ... }
// dateKey는 L열(조회키)에서 파싱한 "yyyy-MM-dd" 날짜 부분이다.
// saveAttendance가 휴직 기간을 새로 저장할 때, 클라이언트가 보낸
// 원래 기간 값과 무관하게 "새 기간 밖에 남은 기존 휴직 행"을
// 정확히 찾아 정리하기 위해 사용한다.
function buildLeaveRowIndexByEmp_(sh) {
  const lastRow = sh.getLastRow();
  const index = {};
  if (lastRow < 2) return index;

  // A열(사원번호), K열(비고), L열(조회키)만 읽으면 되지만,
  // 열이 떨어져 있으므로 A:L 범위를 한 번에 읽고 필요한 열만 사용한다.
  const values = sh.getRange(2, 1, lastRow - 1, 12).getValues();

  values.forEach((r, i) => {
    const empNo = String(r[0] || '').trim();
    const remark = String(r[10] || '').trim();
    const key = String(r[11] || '').trim();
    if (!empNo || remark !== '휴직' || !key) return;

    const sepIdx = key.lastIndexOf('|');
    if (sepIdx < 0) return;
    const dateKey = key.slice(0, sepIdx);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) return;

    const rowNo = i + 2;
    if (!index[empNo]) index[empNo] = [];
    index[empNo].push({ row: rowNo, key: key, dateKey: dateKey });
  });

  return index;
}

function rebuildAttendanceLookupKeys_() {
  const sh = getLogSheet_();
  sh.getRange(1, 12).setValue('_조회키');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };

  const values = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  const keys = values.map(r => {
    const empNo = String(r[0] || '').trim();
    const inDt = r[1];
    const remark = String(r[10] || '').trim();
    const existingKey = String(r[11] || '').trim();

    if (!empNo) return [''];

    if (inDt instanceof Date && !isNaN(inDt.getTime())) {
      const dateKey = Utilities.formatDate(inDt, TIMEZONE, 'yyyy-MM-dd');
      return [buildAttendanceKey_(dateKey, empNo)];
    }

    // 휴직 및 휴무는 실제 출퇴근시간이 없으므로 기존 L열 조회키를 유지한다.
    if ((remark === '휴직' || remark === '휴무') && existingKey) {
      return [existingKey];
    }

    return [''];
  });

  sh.getRange(2, 12, keys.length, 1).setValues(keys).setNumberFormat('@');
  SpreadsheetApp.flush();
  return { ok: true, count: keys.length };
}

// 필요 시 Apps Script 편집기에서 수동 실행 가능한 복구 함수
function rebuildAttendanceIndex() {
  ensureSeoulTimezone_();
  return rebuildAttendanceLookupKeys_();
}

// [진단용] 특정 사원번호의 남아있는 휴직 행을 모두 로그로 출력한다.
// Apps Script 편집기에서 debugFindLeaveRows('사원번호')를 직접 실행해서
// 로그(보기 > 로그)로 확인할 수 있다. 문제가 되는 행을 찾으면
// 시트에서 직접 삭제하거나, saveAttendance를 다시 실행해 정리한다.
function debugFindLeaveRows(empNo) {
  const sh = getLogSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    Logger.log('데이터가 없습니다.');
    return [];
  }
  const values = sh.getRange(2, 1, lastRow - 1, 15).getValues();
  const found = [];
  values.forEach((r, i) => {
    if (String(r[0] || '').trim() === String(empNo || '').trim() && String(r[10] || '').trim() === '휴직') {
      found.push({
        row: i + 2,
        key: r[11],
        leaveStart: r[12],
        leaveEnd: r[13]
      });
    }
  });
  Logger.log(JSON.stringify(found, null, 2));
  return found;
}

function findAttendanceRowsByKey_(sh, key) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  return sh.getRange(2, 12, lastRow - 1, 1)
    .createTextFinder(String(key || ''))
    .matchEntireCell(true)
    .findAll()
    .map(cell => cell.getRow())
    .sort((a, b) => a - b);
}

function findAttendanceRowsByDate_(sh, workDate) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const escaped = String(workDate || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return sh.getRange(2, 12, lastRow - 1, 1)
    .createTextFinder('^' + escaped + '\\|')
    .useRegularExpression(true)
    .findAll()
    .map(cell => cell.getRow())
    .sort((a, b) => a - b);
}

function readAttendanceRows_(sh, rowNumbers) {
  const rows = [...new Set((rowNumbers || []).map(Number).filter(n => n >= 2))]
    .sort((a, b) => a - b);
  if (!rows.length) return [];

  // 연속된 행을 한 번에 읽어 Spreadsheet 호출 횟수 최소화
  const blocks = [];
  let start = rows[0];
  let prev = rows[0];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row === prev + 1) {
      prev = row;
      continue;
    }
    blocks.push([start, prev]);
    start = prev = row;
  }
  blocks.push([start, prev]);

  const result = [];
  blocks.forEach(block => {
    const first = block[0];
    const last = block[1];
    const values = sh.getRange(first, 1, last - first + 1, 14).getValues();
    values.forEach((r, i) => result.push({ row: first + i, values: r }));
  });

  return result;
}

function chooseKeeperAttendanceRow_(records) {
  if (!records || !records.length) throw new Error('기존 출퇴근 기록을 찾을 수 없습니다.');

  // GPS 위치(D/I)가 있는 행을 우선 보존, 그다음 가장 위의 행
  return records.slice().sort((a, b) => {
    const aGps = String(a.values[3] || '').trim() || String(a.values[8] || '').trim() ? 1 : 0;
    const bGps = String(b.values[3] || '').trim() || String(b.values[8] || '').trim() ? 1 : 0;
    if (aGps !== bGps) return bGps - aGps;
    return a.row - b.row;
  })[0];
}

function recentDateKeys_(baseDate, days) {
  const result = [];
  for (let i = 0; i < days; i++) {
    const d = new Date(baseDate.getTime() - i * 24 * 60 * 60 * 1000);
    result.push(Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd'));
  }
  return result;
}

/* =========================================================
   내부 유틸
========================================================= */
function getMasterSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
  if (!sh) throw new Error('작업자마스터 시트가 없습니다.');
  return sh;
}

function getLogSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sh) throw new Error(LOG_SHEET + ' 시트가 없습니다.');
  return sh;
}

function getAuthSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AUTH_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AUTH_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([['사용자명', '비밀번호', '역할', '담당반', '사용여부']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function checkAdmin_(password) {
  if (String(password || '') !== ADMIN_PASSWORD) {
    throw new Error('관리자 비밀번호가 올바르지 않습니다.');
  }
}

function findAuthUser_(username, password, roles) {
  const sh = getAuthSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('권한마스터에 등록된 사용자가 없습니다.');

  const u = String(username || '').trim();
  const p = String(password || '').trim();
  const allowedRoles = (roles || []).map(normalizeRole_);

  const found = sh.getRange(2, 1, lastRow - 1, 5).getDisplayValues()
    .map(r => ({
      username: String(r[0] || '').trim(),
      password: String(r[1] || '').trim(),
      role: normalizeRole_(r[2]),
      team: String(r[3] || 'ALL').trim() || 'ALL',
      active: String(r[4] || 'Y').trim().toUpperCase()
    }))
    .find(r => r.username === u && r.password === p && r.active === 'Y' && allowedRoles.includes(r.role));

  if (!found) throw new Error('사용자명, 비밀번호 또는 권한을 확인하세요.');
  return found;
}

function normalizeRole_(role) {
  const value = String(role || '').trim().toUpperCase();
  if (value === '작성자') return 'WRITER';
  if (value === '검토자') return 'REVIEWER';
  if (value === '승인자') return 'APPROVER';
  return value;
}

function normalizePhone_(phone) {
  return String(phone || '').replace(/[^0-9]/g, '').trim();
}

function normalizeWorkDateValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TIMEZONE, 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  const m = text.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0');
}

function enumerateDateRange_(startDate, endDate, maxDays) {
  validateWorkDate_(startDate);
  validateWorkDate_(endDate);
  if (endDate < startDate) throw new Error('종료일은 시작일보다 빠를 수 없습니다.');

  const start = parseWorkDateTime_(startDate, '00:00');
  const end = parseWorkDateTime_(endDate, '00:00');
  const limit = Number(maxDays || 1826);
  const dates = [];
  let cursor = new Date(start.getTime());

  while (cursor.getTime() <= end.getTime()) {
    dates.push(Utilities.formatDate(cursor, TIMEZONE, 'yyyy-MM-dd'));
    if (dates.length > limit) {
      throw new Error('휴직 기간이 너무 깁니다. 최대 ' + limit + '일까지 설정할 수 있습니다.');
    }
    cursor = new Date(cursor.getTime() + 24 * 60 * 60 * 1000);
  }

  return dates;
}

function validateRemark_(remark) {
  const value = String(remark || '').trim();
  if (!REMARKS.includes(value)) {
    throw new Error('비고는 연차/오전반차/오후반차/교육/공가/휴직/휴무 중에서 선택하세요.');
  }
  return value;
}

function validateTeam_(team) {
  if (!VALID_TEAMS.includes(team)) throw new Error('올바른 작업반을 선택하세요.');
}

function validateWorkDate_(workDate) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(workDate || ''))) {
    throw new Error('작업일을 선택하세요.');
  }
}

function parseWorkDateTime_(workDate, hhmm) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(workDate || ''))) return null;
  if (!/^\d{2}:\d{2}$/.test(String(hhmm || ''))) return null;

  const dateParts = workDate.split('-').map(Number);
  const timeParts = hhmm.split(':').map(Number);
  const utcMillis = Date.UTC(
    dateParts[0], dateParts[1] - 1, dateParts[2],
    timeParts[0] - 9, timeParts[1], 0, 0
  );
  const result = new Date(utcMillis);
  return isNaN(result.getTime()) ? null : result;
}

function ensureSeoulTimezone_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSpreadsheetTimeZone() !== TIMEZONE) {
    ss.setSpreadsheetTimeZone(TIMEZONE);
    SpreadsheetApp.flush();
  }
}

function isCompanyLocationConfigured_() {
  return typeof COMPANY_LAT === 'number' && typeof COMPANY_LNG === 'number' &&
    isFinite(COMPANY_LAT) && isFinite(COMPANY_LNG);
}

function validateCompanyLocation_(latitude, longitude, accuracy) {
  if (!isCompanyLocationConfigured_()) {
    throw new Error('회사 위치가 아직 설정되지 않았습니다. Code.gs 상단의 COMPANY_LAT / COMPANY_LNG를 실제 좌표로 입력하세요.');
  }

  const lat = Number(latitude);
  const lng = Number(longitude);
  const acc = Number(accuracy);

  if (!isFinite(lat) || !isFinite(lng)) throw new Error('현재 위치를 확인할 수 없습니다.');
  if (isFinite(acc) && acc > MAX_GPS_ACCURACY_METERS) {
    throw new Error('GPS 정확도가 낮습니다(±' + Math.round(acc) + 'm). 창가나 실외에서 위치를 다시 확인하세요.');
  }

  const distance = distanceMeters_(COMPANY_LAT, COMPANY_LNG, lat, lng);
  if (distance > ALLOWED_RADIUS_METERS) {
    throw new Error('회사에서 약 ' + Math.round(distance) + 'm 떨어져 있어 출·퇴근할 수 없습니다. 허용 반경은 ' + ALLOWED_RADIUS_METERS + 'm입니다.');
  }

  return { latitude: lat, longitude: lng, accuracy: acc, distance: distance };
}

function distanceMeters_(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function formatLocationText_(loc) {
  const accuracy = isFinite(loc.accuracy) ? ', GPS±' + Math.round(loc.accuracy) + 'm' : '';
  return '회사 / ' + Math.round(loc.distance) + 'm' + accuracy;
}

function buildPunchStatus_(worker, rec) {
  const result = {
    team: worker.team,
    empNo: worker.empNo,
    name: worker.name,
    inTime: '',
    outTime: '',
    status: 'NOT_IN',
    workDate: Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd')
  };

  if (!rec) return result;

  result.workDate = rec.dateKey;
  result.inTime = Utilities.formatDate(rec.inDt, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  if (rec.outDt instanceof Date && !isNaN(rec.outDt.getTime())) {
    result.outTime = Utilities.formatDate(rec.outDt, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    result.status = 'DONE';
  } else {
    result.status = 'WORKING';
  }

  return result;
}
