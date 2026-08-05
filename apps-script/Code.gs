// ============================================================
// 2급 자격과정 운영 통합 프로그램 - Apps Script 백엔드
// 배포 방법은 apps-script/README.md 참고
// ============================================================

var SPREADSHEET_ID = '16Qy0tc7aJKq0vo4h_-GpnPSRu1TuLlU6Nw72dA-t6Vk';
var CERT_SPREADSHEET_ID = '1Rj9PRw-BHkBVmDdm_x16W-Rv96Rx7lRBDbBOxpeYQgI';
var KAKAO_REST_API_KEY = '3842d19d7f043bd4956dc2acdf5aa557';
var SESSION_TTL_MS = 12 * 60 * 60 * 1000; // 12시간

// ------------------------------------------------------------
// 진입점
// ------------------------------------------------------------

function doGet(e) {
  var p = e.parameter;

  if (p.code && !p.action) {
    return handleKakaoOAuthCallback(p.code);
  }
  if (p.diag === '1') {
    return jsonOutput(sendKakaoToMe('[진단 메시지] 이 메시지가 왔다면 카톡 연동이 정상입니다.'));
  }

  if (p.action === 'me') return jsonOutput(actionMe(p.token));
  if (p.action === 'getCourseByCode') return jsonOutput(actionGetCourseByCode(p.courseCode));
  if (p.action === 'studentLookup') return jsonOutput(actionStudentLookup(p.phone, p.birth));

  var session = requireSession(p.token);
  if (!session) return jsonOutput({ ok: false, error: 'unauthorized' });

  if (p.action === 'myCourses') return jsonOutput(actionMyCourses(session));
  if (p.action === 'myReports') return jsonOutput(actionMyReports(session));
  if (p.action === 'myStudents') return jsonOutput(actionMyStudents(session));

  if (session.role === 'admin') {
    if (p.action === 'listCourses') return jsonOutput(actionAdminListCourses());
    if (p.action === 'listInstructors') return jsonOutput(actionAdminListInstructors());
    if (p.action === 'listAllReports') return jsonOutput(actionAdminListAllReports());
    if (p.action === 'listAllApplications') return jsonOutput(actionAdminListAllApplications());
  }

  return jsonOutput({ ok: false, error: 'unknown action' });
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || data.formType;

    if (action === 'login') return jsonOutput(actionLogin(data.id, data.pw));
    if (action === 'logout') return jsonOutput(actionLogout(data.token));
    if (action === 'studentApplication') return jsonOutput(actionStudentApplication(data));
    if (action === 'registerInstructor') return jsonOutput(actionRegisterInstructor(data));

    // 아래는 모두 로그인 세션이 필요
    var session = requireSession(data.token);
    if (!session) return jsonOutput({ ok: false, error: 'unauthorized' });

    if (action === 'report') return jsonOutput(actionSubmitReport(session, data));
    if (action === 'certificateBatch') return jsonOutput(actionCertificateBatch(session, data));
    if (action === 'updateMyAccount') return jsonOutput(actionUpdateMyAccount(session, data));

    if (session.role === 'admin') {
      if (action === 'createInstructor') return jsonOutput(actionAdminCreateInstructor(data));
      if (action === 'createCourse') return jsonOutput(actionAdminCreateCourse(data));
      if (action === 'updateCourse') return jsonOutput(actionAdminUpdateCourse(data));
      if (action === 'markReportReviewed') return jsonOutput(actionAdminMarkReportReviewed(data));
    }

    return jsonOutput({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// 시트 헬퍼
// ------------------------------------------------------------

function getUsersSheet_() {
  var ss = SpreadsheetApp.openById(CERT_SPREADSHEET_ID);
  var sheet = ss.getSheetByName('사용자');
  if (!sheet) {
    sheet = ss.insertSheet('사용자');
    sheet.appendRow(['아이디', '비밀번호해시', '솔트', '역할', '이름', '연락처', '이메일', '은행', '계좌번호', '예금주', '활성여부', '생성일']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getCoursesSheet_() {
  var ss = SpreadsheetApp.openById(CERT_SPREADSHEET_ID);
  var sheet = ss.getSheetByName('과정');
  if (!sheet) {
    sheet = ss.insertSheet('과정');
    sheet.appendRow(['과정ID', '과정명', '담당강사아이디', '담당강사명', '교육일시', '정원', '교육비', '상태', '신청코드', '생성일']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSessionsSheet_() {
  var ss = SpreadsheetApp.openById(CERT_SPREADSHEET_ID);
  var sheet = ss.getSheetByName('세션');
  if (!sheet) {
    sheet = ss.insertSheet('세션');
    sheet.appendRow(['토큰', '아이디', '역할', '이름', '발급시각', '만료시각']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetRowsAsObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) row[headers[j]] = values[i][j];
    row._rowIndex = i + 1; // 1-based sheet row number
    rows.push(row);
  }
  return rows;
}

// ------------------------------------------------------------
// 인증 / 세션
// ------------------------------------------------------------

function hashPassword_(password, salt) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + ':' + salt, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function findUserById_(id) {
  var rows = sheetRowsAsObjects_(getUsersSheet_());
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['아이디']) === String(id)) return rows[i];
  }
  return null;
}

function actionLogin(id, pw) {
  if (!id || !pw) return { ok: false, error: '아이디/비밀번호를 입력해주세요.' };
  var user = findUserById_(id);
  if (!user || String(user['활성여부']).toLowerCase() === 'false') {
    return { ok: false, error: '아이디 또는 비밀번호가 올바르지 않습니다.' };
  }
  var hash = hashPassword_(pw, user['솔트']);
  if (hash !== user['비밀번호해시']) {
    return { ok: false, error: '아이디 또는 비밀번호가 올바르지 않습니다.' };
  }
  var token = Utilities.getUuid();
  var now = new Date();
  var expires = new Date(now.getTime() + SESSION_TTL_MS);
  getSessionsSheet_().appendRow([token, user['아이디'], user['역할'], user['이름'], now, expires]);
  return { ok: true, token: token, role: user['역할'], name: user['이름'], id: user['아이디'] };
}

function actionLogout(token) {
  var sheet = getSessionsSheet_();
  var rows = sheetRowsAsObjects_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i]['토큰'] === token) {
      sheet.deleteRow(rows[i]._rowIndex);
      break;
    }
  }
  return { ok: true };
}

function requireSession(token) {
  if (!token) return null;
  var rows = sheetRowsAsObjects_(getSessionsSheet_());
  var now = new Date();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i]['토큰'] === token) {
      if (new Date(rows[i]['만료시각']) < now) return null;
      return { token: token, id: rows[i]['아이디'], role: rows[i]['역할'], name: rows[i]['이름'] };
    }
  }
  return null;
}

function actionMe(token) {
  var session = requireSession(token);
  if (!session) return { ok: false, error: 'unauthorized' };
  var user = findUserById_(session.id);
  return {
    ok: true, id: session.id, role: session.role, name: session.name,
    bank: user ? user['은행'] : '', account: user ? user['계좌번호'] : '', holder: user ? user['예금주'] : ''
  };
}

// ------------------------------------------------------------
// 과정 (courses)
// ------------------------------------------------------------

function generateCourseCode_() {
  // 전부 숫자로만 된 코드는 시트에 숫자형으로 저장되어 조회 시 혼동을 줄 수 있으므로 문자가 하나 이상 섞인 코드만 사용한다.
  for (var i = 0; i < 20; i++) {
    var code = Utilities.getUuid().replace(/-/g, '').substring(0, 6).toUpperCase();
    if (/[A-F]/.test(code)) return code;
  }
  return 'A' + Utilities.getUuid().replace(/-/g, '').substring(0, 5).toUpperCase();
}

function findCourseById_(courseId) {
  var rows = sheetRowsAsObjects_(getCoursesSheet_());
  for (var i = 0; i < rows.length; i++) {
    if (rows[i]['과정ID'] === courseId) return rows[i];
  }
  return null;
}

function findCourseByCode_(code) {
  var rows = sheetRowsAsObjects_(getCoursesSheet_());
  // 신청코드가 전부 숫자인 경우 시트가 자동으로 숫자형으로 저장하므로 문자열로 맞춰 비교한다.
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['신청코드']) === String(code)) return rows[i];
  }
  return null;
}

function actionGetCourseByCode(code) {
  if (!code) return { ok: false, error: '과정 코드가 없습니다.' };
  var course = findCourseByCode_(code);
  if (!course) return { ok: false, error: '올바르지 않은 과정 코드입니다. 협회 운영진/강사님께 문의해주세요.' };
  var instructor = findUserById_(course['담당강사아이디']);
  return {
    ok: true,
    course: {
      courseId: course['과정ID'],
      name: course['과정명'],
      teacherName: course['담당강사명'],
      datetime: course['교육일시'],
      fee: course['교육비'],
      status: course['상태'],
      bank: instructor ? instructor['은행'] : '',
      account: instructor ? instructor['계좌번호'] : '',
      holder: instructor ? instructor['예금주'] : ''
    }
  };
}

function actionMyCourses(session) {
  var rows = sheetRowsAsObjects_(getCoursesSheet_());
  var mine = rows.filter(function (c) { return c['담당강사아이디'] === session.id; });
  return { ok: true, courses: mine };
}

function actionAdminListCourses() {
  return { ok: true, courses: sheetRowsAsObjects_(getCoursesSheet_()) };
}

function actionAdminCreateCourse(data) {
  var instructor = findUserById_(data.instructorId);
  if (!instructor) return { ok: false, error: '존재하지 않는 강사 아이디입니다.' };
  var courseId = Utilities.getUuid();
  var code = generateCourseCode_();
  getCoursesSheet_().appendRow([
    courseId, data.name || '', data.instructorId, instructor['이름'],
    data.datetime || '', data.capacity || '', data.fee || '', data.status || '모집중', code, new Date()
  ]);
  return { ok: true, courseId: courseId, code: code };
}

function actionAdminUpdateCourse(data) {
  var course = findCourseById_(data.courseId);
  if (!course) return { ok: false, error: '존재하지 않는 과정입니다.' };
  var sheet = getCoursesSheet_();
  var instructor = data.instructorId ? findUserById_(data.instructorId) : findUserById_(course['담당강사아이디']);
  sheet.getRange(course._rowIndex, 1, 1, 10).setValues([[
    course['과정ID'],
    data.name || course['과정명'],
    data.instructorId || course['담당강사아이디'],
    instructor ? instructor['이름'] : course['담당강사명'],
    data.datetime || course['교육일시'],
    data.capacity || course['정원'],
    data.fee || course['교육비'],
    data.status || course['상태'],
    course['신청코드'],
    course['생성일']
  ]]);
  return { ok: true };
}

// ------------------------------------------------------------
// 강사 관리 (admin)
// ------------------------------------------------------------

function actionAdminListInstructors() {
  var rows = sheetRowsAsObjects_(getUsersSheet_()).filter(function (u) { return u['역할'] === 'instructor'; });
  return { ok: true, instructors: rows.map(function (u) {
    return { id: u['아이디'], name: u['이름'], phone: u['연락처'], email: u['이메일'], bank: u['은행'], account: u['계좌번호'], holder: u['예금주'], active: u['활성여부'] };
  }) };
}

function actionAdminCreateInstructor(data) {
  if (findUserById_(data.id)) return { ok: false, error: '이미 존재하는 아이디입니다.' };
  var salt = Utilities.getUuid();
  getUsersSheet_().appendRow([
    data.id, hashPassword_(data.pw, salt), salt, 'instructor', data.name || '',
    data.phone || '', data.email || '', data.bank || '', data.account || '', data.holder || '',
    true, new Date()
  ]);
  return { ok: true };
}

// 강사 본인이 직접 아이디/비밀번호를 정해 가입 (관리자 승인 없이 즉시 사용 가능, 비밀번호는 본인만 앎)
function actionRegisterInstructor(data) {
  if (!data.id || !data.pw || !data.name) return { ok: false, error: '아이디, 비밀번호, 이름을 입력해주세요.' };
  if (String(data.pw).length < 4) return { ok: false, error: '비밀번호는 4자 이상으로 설정해주세요.' };
  if (findUserById_(data.id)) return { ok: false, error: '이미 사용 중인 아이디입니다. 다른 아이디를 입력해주세요.' };
  var salt = Utilities.getUuid();
  getUsersSheet_().appendRow([
    data.id, hashPassword_(data.pw, salt), salt, 'instructor', data.name || '',
    data.phone || '', data.email || '', data.bank || '', data.account || '', data.holder || '',
    true, new Date()
  ]);
  var msg = '[강사 등록]\n이름: ' + data.name + '\n아이디: ' + data.id + (data.phone ? ('\n연락처: ' + data.phone) : '');
  var kakaoResult = sendKakaoToMe(msg);
  return { ok: true, kakao: kakaoResult };
}

function actionUpdateMyAccount(session, data) {
  var user = findUserById_(session.id);
  if (!user) return { ok: false, error: 'not found' };
  var sheet = getUsersSheet_();
  sheet.getRange(user._rowIndex, 8, 1, 3).setValues([[data.bank || '', data.account || '', data.holder || '']]);
  return { ok: true };
}

// ------------------------------------------------------------
// 과정 운영 신고서 (instructor -> admin)
// ------------------------------------------------------------

function actionSubmitReport(session, data) {
  var course = findCourseById_(data.courseId);
  if (!course || course['담당강사아이디'] !== session.id) {
    return { ok: false, error: '본인에게 배정된 과정만 신고할 수 있습니다.' };
  }
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var headers = ['타임스탬프', '과정ID', '과정명', '강사 성함', '연락처', '교육 일시', '교육 인원', '교육비(1인)', '강의안 파일', '확인여부', '비고'];
  var fileUrl = '';
  if (data.fileBase64 && data.fileName) {
    fileUrl = saveFile(data, '2급자격_강의안_모음');
  }
  var user = findUserById_(session.id);
  var row = [
    new Date(), course['과정ID'], course['과정명'], session.name, user ? user['연락처'] : '',
    data.date || '', data.headcount || '', data.fee || '', fileUrl || '(없음)', false, data.note || ''
  ];
  appendToSheet(ss, '전체', headers, row);
  appendToSheet(ss, sanitizeSheetName(session.name || '무명'), headers, row);

  var msg = '[2급자격 운영 신고 접수]\n과정명: ' + course['과정명'] + '\n강사: ' + session.name +
    '\n일시: ' + (data.date || '') + '\n인원: ' + (data.headcount || '') + '명\n교육비: ' + (data.fee || '') + '원' +
    (fileUrl ? ('\n강의안: ' + fileUrl) : '');
  var kakaoResult = sendKakaoToMe(msg);
  return { ok: true, kakao: kakaoResult };
}

function actionMyReports(session) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('전체');
  if (!sheet) return { ok: true, reports: [] };
  var rows = sheetRowsAsObjects_(sheet).filter(function (r) { return r['강사 성함'] === session.name; });
  return { ok: true, reports: rows };
}

function actionAdminListAllReports() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('전체');
  return { ok: true, reports: sheet ? sheetRowsAsObjects_(sheet) : [] };
}

function actionAdminMarkReportReviewed(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('전체');
  if (!sheet) return { ok: false, error: 'not found' };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = headers.indexOf('확인여부') + 1;
  if (col < 1) return { ok: false, error: '확인여부 컬럼이 없습니다.' };
  sheet.getRange(data.rowIndex, col).setValue(!!data.reviewed);
  return { ok: true };
}

// ------------------------------------------------------------
// 수강생 신청 (student, 로그인 불필요 - 과정코드로 진입)
// ------------------------------------------------------------

function actionStudentApplication(data) {
  var course = findCourseById_(data.courseId);
  if (!course) return { ok: false, error: '올바르지 않은 과정입니다.' };
  var ss = SpreadsheetApp.openById(CERT_SPREADSHEET_ID);
  // 주의: '학습자신청' 시트는 수강생 본인 신청용이며, '수료증신청' 시트(강사 일괄 신청, 자격관리번호 포함)와는
  // 컬럼 구조가 달라 별개로 유지한다. 기존 라이브 데이터가 '수료증신청'에 있으므로 절대 같은 시트를 쓰지 말 것.
  var headers = ['타임스탬프', '과정ID', '구분', '자격/과정명', '신청자 성명', '생년월일', '연락처', '이메일', '우편주소', '담당 강사', '입금자명', '입금일자', '비고'];
  var row = [
    new Date(), course['과정ID'], data.docLabel || '', data.certName || '', data.applicantName || '',
    data.birthdate || '', data.applicantPhone || '', data.email || '', data.address || '',
    course['담당강사명'], data.payerName || '', data.paymentDate || '', data.note || ''
  ];
  appendToSheet(ss, '학습자신청', headers, row);
  appendToSheet(ss, '학습자신청_' + sanitizeSheetName(course['담당강사명'] || '무명'), headers, row);

  var msg = '[수료증/자격증 신청 접수]\n과정명: ' + course['과정명'] + '\n담당 강사: ' + course['담당강사명'] +
    '\n신청자: ' + (data.applicantName || '') + '\n연락처: ' + (data.applicantPhone || '') +
    '\n입금자명: ' + (data.payerName || '') + '\n입금일자: ' + (data.paymentDate || '');
  var kakaoResult = sendKakaoToMe(msg);
  return { ok: true, kakao: kakaoResult };
}

function actionMyStudents(session) {
  var ss = SpreadsheetApp.openById(CERT_SPREADSHEET_ID);
  var sheet = ss.getSheetByName('학습자신청');
  if (!sheet) return { ok: true, students: [] };
  var rows = sheetRowsAsObjects_(sheet).filter(function (r) { return r['담당 강사'] === session.name; });
  return { ok: true, students: rows };
}

function actionAdminListAllApplications() {
  var ss = SpreadsheetApp.openById(CERT_SPREADSHEET_ID);
  var sheet = ss.getSheetByName('학습자신청');
  return { ok: true, students: sheet ? sheetRowsAsObjects_(sheet) : [] };
}

// 수강생용 경량 "로그인": 연락처+생년월일 일치로 본인 신청 내역만 조회 (계정 발급 없이)
function actionStudentLookup(phone, birth) {
  if (!phone || !birth) return { ok: false, error: '연락처와 생년월일을 입력해주세요.' };
  var ss = SpreadsheetApp.openById(CERT_SPREADSHEET_ID);
  var sheet = ss.getSheetByName('학습자신청');
  if (!sheet) return { ok: true, applications: [] };
  var rows = sheetRowsAsObjects_(sheet).filter(function (r) {
    return String(r['연락처']) === String(phone) && String(r['생년월일']) === String(birth);
  });
  return { ok: true, applications: rows };
}

// ------------------------------------------------------------
// 수료증/자격증 일괄 신청 (instructor -> 협회)
// ------------------------------------------------------------

function actionCertificateBatch(session, data) {
  var ss = SpreadsheetApp.openById(CERT_SPREADSHEET_ID);
  var user = findUserById_(session.id);
  // 주의: 기존 라이브 데이터가 이미 이 12개 컬럼(과정ID 없음)으로 쌓여있으므로 헤더 순서를 바꾸지 말 것.
  var headers = ['타임스탬프', '자격관리번호', '구분', '자격/과정명', '학생 성명', '생년월일', '담당 강사', '강사 연락처', '강사 이메일', '우편주소', '입금자명', '비고'];
  var students = data.students || [];
  var count = 0;
  for (var i = 0; i < students.length; i++) {
    var s = students[i];
    if (!s.name) continue;
    var certNumber = getNextCertNumber(data.docType);
    var row = [
      new Date(), certNumber, data.docLabel || '', data.certName || '',
      s.name || '', s.birthdate || '', session.name, user ? user['연락처'] : '', user ? user['이메일'] : '',
      data.mailAddress || '', data.payerName || '', data.note || ''
    ];
    appendToSheet(ss, '수료증신청', headers, row);
    appendToSheet(ss, '수료증신청_' + sanitizeSheetName(session.name || '무명'), headers, row);
    count++;
  }
  var msg = '[' + (data.docLabel || '서류') + ' 신청 접수]\n담당 강사: ' + session.name +
    '\n신청 인원: ' + count + '명\n자격/과정명: ' + (data.certName || '') +
    '\n입금자명: ' + (data.payerName || '');
  var kakaoResult = sendKakaoToMe(msg);
  return { ok: true, count: count, kakao: kakaoResult };
}

// ------------------------------------------------------------
// 카카오 알림 (기존 로직 그대로 유지)
// ------------------------------------------------------------

function handleKakaoOAuthCallback(code) {
  try {
    var resp = UrlFetchApp.fetch('https://kauth.kakao.com/oauth/token', {
      method: 'post',
      payload: {
        grant_type: 'authorization_code',
        client_id: KAKAO_REST_API_KEY,
        redirect_uri: ScriptApp.getService().getUrl(),
        code: code
      },
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    if (data.access_token) {
      var props = PropertiesService.getScriptProperties();
      props.setProperty('KAKAO_ACCESS_TOKEN', data.access_token);
      props.setProperty('KAKAO_REFRESH_TOKEN', data.refresh_token);
      return ContentService.createTextOutput('카카오 연동 완료! 이 창은 닫으셔도 됩니다.');
    }
    return ContentService.createTextOutput('연동 실패: ' + resp.getContentText());
  } catch (err) {
    return ContentService.createTextOutput('오류: ' + String(err));
  }
}

function getKakaoAccessToken() {
  var props = PropertiesService.getScriptProperties();
  var refreshToken = props.getProperty('KAKAO_REFRESH_TOKEN');
  if (!refreshToken) return null;
  var resp = UrlFetchApp.fetch('https://kauth.kakao.com/oauth/token', {
    method: 'post',
    payload: { grant_type: 'refresh_token', client_id: KAKAO_REST_API_KEY, refresh_token: refreshToken },
    muteHttpExceptions: true
  });
  var data = JSON.parse(resp.getContentText());
  if (data.access_token) {
    props.setProperty('KAKAO_ACCESS_TOKEN', data.access_token);
    if (data.refresh_token) props.setProperty('KAKAO_REFRESH_TOKEN', data.refresh_token);
    return data.access_token;
  }
  Logger.log('refresh failed: ' + resp.getContentText());
  return null;
}

function sendKakaoToMe(message) {
  try {
    var token = getKakaoAccessToken();
    if (!token) return { ok: false, stage: 'token', detail: 'no refresh token or refresh failed' };
    var template = {
      object_type: 'text', text: message,
      link: { web_url: 'https://tweetyseoul1.github.io/level2-training-report/', mobile_web_url: 'https://tweetyseoul1.github.io/level2-training-report/' }
    };
    var resp = UrlFetchApp.fetch('https://kapi.kakao.com/v2/api/talk/memo/default/send', {
      method: 'post',
      headers: { Authorization: 'Bearer ' + token },
      payload: { template_object: JSON.stringify(template) },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code >= 200 && code < 300) return { ok: true };
    return { ok: false, stage: 'send', code: code, detail: resp.getContentText() };
  } catch (err) {
    return { ok: false, stage: 'exception', detail: String(err) };
  }
}

// ------------------------------------------------------------
// 공용 유틸 (기존 로직 그대로 유지)
// ------------------------------------------------------------

function getNextCertNumber(docType) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var year = new Date().getFullYear();
    var prefix = docType === 'qualification' ? '자격' : '수료';
    var key = 'CERTNUM_' + docType + '_' + year;
    var current = parseInt(props.getProperty(key) || '0', 10);
    var next = current + 1;
    props.setProperty(key, String(next));
    var padded = ('0000' + next).slice(-4);
    return prefix + '-' + year + '-' + padded;
  } finally {
    lock.releaseLock();
  }
}

function saveFile(data, folderName) {
  var folder = getOrCreateFolder(folderName);
  var bytes = Utilities.base64Decode(data.fileBase64);
  var blob = Utilities.newBlob(bytes, data.fileMime || 'application/octet-stream', data.fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function appendToSheet(ss, name, headers, row) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  } else {
    var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    var same = existing.length === headers.length && existing.every(function (h, i) { return h === headers[i]; });
    if (!same) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  sheet.appendRow(row);
}

function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function sanitizeSheetName(name) {
  var clean = name.replace(/[\[\]\*\/\\\?:]/g, '').trim();
  return (clean || '무명').substring(0, 95);
}

// ------------------------------------------------------------
// 1회성 마이그레이션 (스크립트 편집기에서 딱 한 번 수동 실행)
// ------------------------------------------------------------

function migrateOnce_setupAdminAndInstructors() {
  var usersSheet = getUsersSheet_();
  if (findUserById_('admin')) {
    Logger.log('이미 마이그레이션이 실행된 것 같습니다 (admin 계정 존재). 중복 실행하지 마세요.');
    return;
  }

  // 1) 관리자 계정 생성 (기존 공유 비밀번호 ai2026 유지, 로그인 후 반드시 변경할 것)
  var adminSalt = Utilities.getUuid();
  usersSheet.appendRow(['admin', hashPassword_('ai2026', adminSalt), adminSalt, 'admin', '협회 운영진', '', '', '', '', '', true, new Date()]);

  // 2) 기존 강사계좌 시트 -> 사용자 시트로 이전 (아이디/비밀번호는 강사 전화번호 뒷자리 등으로 협회 운영진이 직접 지정 후 안내)
  var certSs = SpreadsheetApp.openById(CERT_SPREADSHEET_ID);
  var accSheet = certSs.getSheetByName('강사계좌');
  if (accSheet) {
    var rows = sheetRowsAsObjects_(accSheet);
    rows.forEach(function (r) {
      if (!r['강사명']) return;
      var id = r['강사명']; // 임시 아이디 = 강사명. 협회 운영진이 관리자 화면에서 아이디/비번을 재설정해줄 것.
      if (findUserById_(id)) return;
      var salt = Utilities.getUuid();
      var tempPw = '0000';
      usersSheet.appendRow([id, hashPassword_(tempPw, salt), salt, 'instructor', r['강사명'], '', '', r['은행'] || '', r['계좌번호'] || '', r['예금주'] || '', true, new Date()]);
    });
  }
  Logger.log('마이그레이션 완료. 사용자 시트를 확인하고, 각 강사 아이디/임시비밀번호(0000)를 안내한 뒤 반드시 비밀번호를 바꾸게 하세요.');
}
