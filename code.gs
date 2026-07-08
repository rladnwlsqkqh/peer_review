/**
 * 동료평가 시스템 - 서버 로직
 * 데이터는 이 스크립트가 연결된 Google Sheet에 저장됩니다.
 *
 * [v2 변경 사항]
 * 1. LockService 적용: 여러 명이 동시에 저장/제출해도 데이터가 유실되지 않도록 함
 * 2. 권한 검증: 선생님 전용 함수는 teacherId + 세션 소유권 확인 후에만 동작
 * 3. 제출 검증: 마감된 평가는 제출 불가, 점수는 1~5 정수만 허용
 * 4. 문항 잠금: 제출된 응답이 하나라도 있으면 문항 수정 불가
 */

var SHEET_SCHEMAS = {
  Teachers: ['teacherId', 'name', 'subject', 'password', 'createdAt'],
  Sessions: ['sessionId', 'teacherId', 'title', 'accessCode', 'questionsJson', 'status', 'createdAt'],
  Students: ['sessionId', 'studentId', 'name', 'groupName'],
  Responses: ['sessionId', 'evaluatorStudentId', 'targetStudentId', 'answersJson', 'submittedAt']
};

function doGet(e) {
  initializeSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('동료평가 시스템')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 외부(깃허브 페이지 등)에서 오는 요청을 받는 API 창구 ---------- */

// 외부에서 호출을 허용할 함수 목록 (여기 없는 함수는 절대 실행되지 않음)
function getApiFunctions_() {
  return {
    signUpTeacher: signUpTeacher,
    loginTeacher: loginTeacher,
    createSession: createSession,
    getTeacherSessions: getTeacherSessions,
    getSessionDetail: getSessionDetail,
    setSessionStatus: setSessionStatus,
    saveRoster: saveRoster,
    saveGroups: saveGroups,
    saveQuestions: saveQuestions,
    getMyAssignedSessions: getMyAssignedSessions,
    loginStudent: loginStudent,
    getGroupMembers: getGroupMembers,
    submitEvaluation: submitEvaluation,
    getSessionResults: getSessionResults
  };
}

// 깃허브에 올린 index.html이 fetch(POST)로 {fn: 함수이름, args: [인자들]}를
// 보내면, 해당 함수를 실행하고 결과를 JSON으로 돌려줍니다.
function doPost(e) {
  initializeSheets_();
  var res;
  try {
    var req = JSON.parse(e.postData.contents);
    var fn = getApiFunctions_()[req.fn];
    if (!fn) {
      res = { ok: false, message: '알 수 없는 요청입니다: ' + req.fn };
    } else {
      res = fn.apply(null, req.args || []);
    }
  } catch (err) {
    res = { ok: false, message: '서버 오류: ' + err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}

// Apps Script 편집기에서 이 함수를 한 번 실행하면 시트가 만들어지고 권한 승인 창이 뜹니다.
function setup() {
  initializeSheets_();
}

// [주의] 테스트 데이터 전체 삭제. 선생님 계정/평가/명단/응답을 모두 지우고 빈 시트로 되돌립니다.
// Apps Script 편집기 상단에서 함수를 resetAllData 로 선택하고 실행하세요. (되돌릴 수 없음)
function resetAllData() {
  initializeSheets_();
  Object.keys(SHEET_SCHEMAS).forEach(function (name) {
    var sheet = getSheet_(name);
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, SHEET_SCHEMAS[name].length).clearContent();
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('모든 데이터를 삭제했습니다.', '초기화 완료', 5);
  return '초기화 완료: 모든 데이터가 삭제되었습니다.';
}

function initializeSheets_() {
  Object.keys(SHEET_SCHEMAS).forEach(function (name) {
    getSheet_(name);
  });
}

function getSheet_(name) {
  var headers = SHEET_SCHEMAS[name];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.join('') === '') continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = row[j];
    obj._row = i + 1;
    out.push(obj);
  }
  return out;
}

function appendObject_(sheet, obj) {
  var headers = SHEET_SCHEMAS[sheet.getName()];
  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.appendRow(row);
}

function updateObjectRow_(sheet, rowIndex, obj) {
  var headers = SHEET_SCHEMAS[sheet.getName()];
  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

function rewriteSheetRows_(sheet, objects) {
  var headers = SHEET_SCHEMAS[sheet.getName()];
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  if (objects.length === 0) return;
  var data = objects.map(function (obj) {
    return headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  });
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
}

function generateId_(prefix) {
  return prefix + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
}

function generateAccessCode_(sessions) {
  var existing = sessions.map(function (s) { return s.accessCode; });
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 혼동되는 0/O, 1/I 제외
  var code;
  do {
    code = '';
    for (var i = 0; i < 6; i++) code += chars.charAt(Math.floor(Math.random() * chars.length));
  } while (existing.indexOf(code) !== -1);
  return code;
}

/* ---------- 공통: 락 / 권한 헬퍼 ---------- */

// 시트에 쓰는(저장/제출) 작업은 전부 이 함수로 감싸서 한 번에 하나씩만 실행되게 합니다.
// 이렇게 하면 여러 명이 동시에 제출해도 서로의 데이터를 덮어쓰지 않습니다.
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 최대 15초 대기
  } catch (e) {
    return { ok: false, message: '접속이 많아 처리하지 못했습니다. 잠시 후 다시 시도해주세요.' };
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// teacherId가 실제 선생님 계정인지 확인
function findTeacher_(teacherId) {
  if (!teacherId) return null;
  var teachers = sheetToObjects_(getSheet_('Teachers'));
  return teachers.find(function (t) { return t.teacherId === teacherId; }) || null;
}

// 해당 세션이 존재하고, 그 세션의 주인이 teacherId인지 확인.
// 학생이 브라우저 콘솔에서 선생님 함수를 직접 호출해도
// teacherId를 모르면 아무것도 못 하게 막는 핵심 장치입니다.
function findSessionOwned_(teacherId, sessionId) {
  if (!teacherId || !sessionId) return null;
  var sessions = sheetToObjects_(getSheet_('Sessions'));
  var session = sessions.find(function (s) { return s.sessionId === sessionId; });
  if (!session || session.teacherId !== teacherId) return null;
  return session;
}

function authFail_() {
  return { ok: false, message: '권한이 없습니다. 다시 로그인해주세요.' };
}

/* ---------- 선생님: 회원가입 / 로그인 ---------- */

function signUpTeacher(name, subject, password) {
  name = (name || '').trim();
  subject = (subject || '').trim();
  password = (password || '').trim();
  if (!name || !subject || !password) {
    return { ok: false, message: '이름, 과목, 비밀번호를 모두 입력해주세요.' };
  }
  if (password.length < 4) {
    return { ok: false, message: '비밀번호는 4자 이상으로 설정해주세요.' };
  }
  return withLock_(function () {
    var sheet = getSheet_('Teachers');
    var teachers = sheetToObjects_(sheet);
    var dup = teachers.some(function (t) { return String(t.password) === password; });
    if (dup) {
      return { ok: false, message: '이미 사용 중인 비밀번호입니다. 다른 비밀번호를 사용해주세요.' };
    }
    var teacherId = generateId_('T');
    appendObject_(sheet, { teacherId: teacherId, name: name, subject: subject, password: password, createdAt: new Date() });
    return { ok: true, teacherId: teacherId, name: name, subject: subject };
  });
}

function loginTeacher(password) {
  password = (password || '').trim();
  if (!password) return { ok: false, message: '비밀번호를 입력해주세요.' };
  var teachers = sheetToObjects_(getSheet_('Teachers'));
  var found = teachers.find(function (t) { return String(t.password) === password; });
  if (!found) return { ok: false, message: '비밀번호가 일치하는 계정이 없습니다.' };
  return { ok: true, teacherId: found.teacherId, name: found.name, subject: found.subject };
}

/* ---------- 선생님: 세션(평가 활동) 관리 ---------- */

function createSession(teacherId, title) {
  if (!findTeacher_(teacherId)) return authFail_();
  title = (title || '').trim();
  if (!title) return { ok: false, message: '평가 제목을 입력해주세요.' };
  return withLock_(function () {
    var sheet = getSheet_('Sessions');
    var sessions = sheetToObjects_(sheet);
    var sessionId = generateId_('S');
    var accessCode = generateAccessCode_(sessions);
    appendObject_(sheet, {
      sessionId: sessionId, teacherId: teacherId, title: title, accessCode: accessCode,
      questionsJson: JSON.stringify([]), status: 'open', createdAt: new Date()
    });
    return { ok: true, sessionId: sessionId, accessCode: accessCode };
  });
}

function getTeacherSessions(teacherId) {
  if (!findTeacher_(teacherId)) return [];
  var sessions = sheetToObjects_(getSheet_('Sessions'));
  return sessions
    .filter(function (s) { return s.teacherId === teacherId; })
    .map(function (s) { return { sessionId: s.sessionId, title: s.title, accessCode: s.accessCode, status: s.status }; })
    .reverse();
}

// 내부용: 권한 검증 없이 세션 상세 정보를 조립 (getSessionResults 등에서 재사용)
function buildSessionDetail_(session) {
  var students = sheetToObjects_(getSheet_('Students'))
    .filter(function (st) { return st.sessionId === session.sessionId; })
    .map(function (st) { return { studentId: String(st.studentId), name: st.name, groupName: st.groupName || '' }; });
  var questions = [];
  try { questions = JSON.parse(session.questionsJson || '[]'); } catch (e) { questions = []; }
  return {
    ok: true, sessionId: session.sessionId, title: session.title, accessCode: session.accessCode,
    status: session.status, students: students, questions: questions
  };
}

function getSessionDetail(teacherId, sessionId) {
  var session = findSessionOwned_(teacherId, sessionId);
  if (!session) return authFail_();
  var detail = buildSessionDetail_(session);
  // 제출된 응답이 있으면 화면에서 문항 편집을 잠그기 위한 플래그
  detail.hasResponses = sheetToObjects_(getSheet_('Responses'))
    .some(function (r) { return r.sessionId === sessionId; });
  return detail;
}

function setSessionStatus(teacherId, sessionId, status) {
  if (!findSessionOwned_(teacherId, sessionId)) return authFail_();
  return withLock_(function () {
    var sheet = getSheet_('Sessions');
    var sessions = sheetToObjects_(sheet);
    var idx = sessions.findIndex(function (s) { return s.sessionId === sessionId; });
    if (idx === -1) return { ok: false, message: '평가를 찾을 수 없습니다.' };
    var target = sessions[idx];
    var rowIndex = target._row;
    target.status = status === 'closed' ? 'closed' : 'open';
    delete target._row;
    updateObjectRow_(sheet, rowIndex, target);
    return { ok: true };
  });
}

function saveRoster(teacherId, sessionId, rosterText) {
  if (!findSessionOwned_(teacherId, sessionId)) return authFail_();
  var lines = (rosterText || '').split('\n');
  var parsed = [];
  var seen = {};
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    var parts = line.split(',');
    if (parts.length < 2) continue;
    var studentId = parts[0].trim();
    var name = parts.slice(1).join(',').trim();
    if (!studentId || !name || seen[studentId]) continue;
    seen[studentId] = true;
    parsed.push({ studentId: studentId, name: name });
  }
  if (parsed.length === 0) {
    return { ok: false, message: '유효한 학생 정보가 없습니다. "학번,이름" 형식으로 한 줄에 한 명씩 입력해주세요.' };
  }
  return withLock_(function () {
    var sheet = getSheet_('Students');
    var all = sheetToObjects_(sheet);
    var others = all.filter(function (s) { return s.sessionId !== sessionId; });
    var groupMap = {};
    all.filter(function (s) { return s.sessionId === sessionId; })
      .forEach(function (s) { groupMap[String(s.studentId)] = s.groupName; });

    var newRows = others.concat(parsed.map(function (p) {
      return { sessionId: sessionId, studentId: p.studentId, name: p.name, groupName: groupMap[p.studentId] || '' };
    }));
    rewriteSheetRows_(sheet, newRows);
    return { ok: true, count: parsed.length };
  });
}

function saveGroups(teacherId, sessionId, assignments) {
  if (!findSessionOwned_(teacherId, sessionId)) return authFail_();
  return withLock_(function () {
    var sheet = getSheet_('Students');
    var all = sheetToObjects_(sheet);
    var groupByStudent = {};
    (assignments || []).forEach(function (a) { groupByStudent[String(a.studentId)] = a.groupName || ''; });
    var updated = all.map(function (s) {
      if (s.sessionId === sessionId && groupByStudent.hasOwnProperty(String(s.studentId))) {
        return { sessionId: s.sessionId, studentId: s.studentId, name: s.name, groupName: groupByStudent[String(s.studentId)] };
      }
      return { sessionId: s.sessionId, studentId: s.studentId, name: s.name, groupName: s.groupName };
    });
    rewriteSheetRows_(sheet, updated);
    return { ok: true };
  });
}

function saveQuestions(teacherId, sessionId, questions) {
  if (!findSessionOwned_(teacherId, sessionId)) return authFail_();
  questions = (questions || [])
    .filter(function (q) { return q && q.text && q.text.trim(); })
    .map(function (q) { return { text: q.text.trim(), type: q.type === 'text' ? 'text' : 'score' }; });
  if (questions.length === 0) return { ok: false, message: '문항을 1개 이상 입력해주세요.' };
  return withLock_(function () {
    // 이미 제출된 응답이 있으면 문항 수정 불가.
    // (답변이 문항 순서대로 저장되기 때문에, 문항을 바꾸면 기존 답이 엉뚱한 문항에 붙게 됩니다)
    var hasResponses = sheetToObjects_(getSheet_('Responses'))
      .some(function (r) { return r.sessionId === sessionId; });
    if (hasResponses) {
      return { ok: false, message: '이미 제출된 응답이 있어 문항을 수정할 수 없습니다. 문항을 바꾸려면 새 평가를 만들어주세요.' };
    }
    var sheet = getSheet_('Sessions');
    var all = sheetToObjects_(sheet);
    var idx = all.findIndex(function (s) { return s.sessionId === sessionId; });
    if (idx === -1) return { ok: false, message: '평가를 찾을 수 없습니다.' };
    var target = all[idx];
    var rowIndex = target._row;
    target.questionsJson = JSON.stringify(questions);
    delete target._row;
    updateObjectRow_(sheet, rowIndex, target);
    return { ok: true };
  });
}

/* ---------- 학생: 로그인 / 평가 ---------- */

// 학번+이름으로 본인을 확인한 뒤, 그 학생이 명단에 들어 있는 "진행중(open)" 평가만 반환
function getMyAssignedSessions(studentId, name) {
  studentId = (studentId || '').trim();
  name = (name || '').trim();
  if (!studentId || !name) return { ok: false, message: '학번과 이름을 모두 입력해주세요.' };

  var students = sheetToObjects_(getSheet_('Students'));
  var myRows = students.filter(function (s) {
    return String(s.studentId).trim() === studentId && String(s.name).trim() === name;
  });
  if (!myRows.length) return { ok: false, message: '학번 또는 이름이 명단에 없습니다. 선생님께 확인해주세요.' };

  var mySessionIds = {};
  myRows.forEach(function (s) { mySessionIds[s.sessionId] = true; });

  var teachers = sheetToObjects_(getSheet_('Teachers'));
  var teacherMap = {};
  teachers.forEach(function (t) { teacherMap[t.teacherId] = t; });

  var sessions = sheetToObjects_(getSheet_('Sessions'))
    .filter(function (s) { return mySessionIds[s.sessionId] && s.status === 'open'; })
    .map(function (s) {
      var t = teacherMap[s.teacherId];
      return { sessionId: s.sessionId, title: s.title, teacherName: t ? t.name : '', subject: t ? t.subject : '' };
    }).reverse();

  return { ok: true, sessions: sessions };
}

function loginStudent(sessionId, studentId, name) {
  sessionId = (sessionId || '').trim();
  studentId = (studentId || '').trim();
  name = (name || '').trim();
  if (!sessionId || !studentId || !name) {
    return { ok: false, message: '평가를 선택하고 학번, 이름을 모두 입력해주세요.' };
  }
  var sessions = sheetToObjects_(getSheet_('Sessions'));
  var session = sessions.find(function (s) { return s.sessionId === sessionId; });
  if (!session) return { ok: false, message: '선택한 평가를 찾을 수 없습니다.' };
  if (session.status !== 'open') return { ok: false, message: '이 평가는 현재 마감되어 참여할 수 없습니다.' };

  var students = sheetToObjects_(getSheet_('Students')).filter(function (s) { return s.sessionId === session.sessionId; });
  var me = students.find(function (s) { return String(s.studentId) === studentId && String(s.name).trim() === name; });
  if (!me) return { ok: false, message: '학번 또는 이름이 명단과 일치하지 않습니다. 선생님께 확인해주세요.' };

  var questions = [];
  try { questions = JSON.parse(session.questionsJson || '[]'); } catch (e) { questions = []; }
  return {
    ok: true, sessionId: session.sessionId, title: session.title,
    studentId: String(me.studentId), name: me.name, groupName: me.groupName || '',
    questions: questions
  };
}

function getGroupMembers(sessionId, studentId) {
  var students = sheetToObjects_(getSheet_('Students')).filter(function (s) { return s.sessionId === sessionId; });
  var me = students.find(function (s) { return String(s.studentId) === String(studentId); });
  if (!me) return { ok: false, message: '학생 정보를 찾을 수 없습니다.' };
  var groupName = me.groupName || '';
  if (!groupName) return { ok: true, groupName: '', members: [] };

  var members = students
    .filter(function (s) { return (s.groupName || '') === groupName && String(s.studentId) !== String(studentId); })
    .map(function (s) { return { studentId: String(s.studentId), name: s.name }; });

  var responses = sheetToObjects_(getSheet_('Responses'))
    .filter(function (r) { return r.sessionId === sessionId && String(r.evaluatorStudentId) === String(studentId); });
  var submittedSet = {};
  responses.forEach(function (r) { submittedSet[String(r.targetStudentId)] = true; });
  members.forEach(function (m) { m.submitted = !!submittedSet[m.studentId]; });

  return { ok: true, groupName: groupName, members: members };
}

function submitEvaluation(sessionId, evaluatorId, targetId, answers) {
  if (!sessionId || !evaluatorId || !targetId) return { ok: false, message: '잘못된 요청입니다.' };
  if (String(evaluatorId) === String(targetId)) return { ok: false, message: '본인은 평가할 수 없습니다.' };

  return withLock_(function () {
    // 세션이 존재하고 아직 열려 있는지 확인 (마감 후 제출 방지)
    var sessions = sheetToObjects_(getSheet_('Sessions'));
    var session = sessions.find(function (s) { return s.sessionId === sessionId; });
    if (!session) return { ok: false, message: '평가를 찾을 수 없습니다.' };
    if (session.status !== 'open') return { ok: false, message: '이 평가는 마감되어 더 이상 제출할 수 없습니다.' };

    var students = sheetToObjects_(getSheet_('Students')).filter(function (s) { return s.sessionId === sessionId; });
    var evaluator = students.find(function (s) { return String(s.studentId) === String(evaluatorId); });
    var target = students.find(function (s) { return String(s.studentId) === String(targetId); });
    if (!evaluator || !target) return { ok: false, message: '학생 정보를 찾을 수 없습니다.' };
    if (!evaluator.groupName || evaluator.groupName !== target.groupName) {
      return { ok: false, message: '같은 모둠원만 평가할 수 있습니다.' };
    }

    // 답변 검증: 점수형은 1~5 정수만 허용, 서술형은 최대 2000자로 제한
    var questions = [];
    try { questions = JSON.parse(session.questionsJson || '[]'); } catch (e) { questions = []; }
    var raw = answers || [];
    var clean = [];
    for (var i = 0; i < questions.length; i++) {
      var q = questions[i];
      var ans = raw[i];
      if (q.type === 'score') {
        var val = Number(ans);
        if (ans === '' || ans === null || ans === undefined || isNaN(val) || val % 1 !== 0 || val < 1 || val > 5) {
          return { ok: false, message: '점수형 문항에는 1~5점만 제출할 수 있습니다.' };
        }
        clean.push(val);
      } else {
        var text = (ans === null || ans === undefined) ? '' : String(ans).trim();
        if (text.length > 2000) text = text.slice(0, 2000);
        clean.push(text);
      }
    }

    var sheet = getSheet_('Responses');
    var all = sheetToObjects_(sheet);
    var idx = all.findIndex(function (r) {
      return r.sessionId === sessionId && String(r.evaluatorStudentId) === String(evaluatorId) && String(r.targetStudentId) === String(targetId);
    });
    var obj = {
      sessionId: sessionId, evaluatorStudentId: evaluatorId, targetStudentId: targetId,
      answersJson: JSON.stringify(clean), submittedAt: new Date()
    };
    if (idx === -1) {
      appendObject_(sheet, obj);
    } else {
      updateObjectRow_(sheet, all[idx]._row, obj);
    }
    return { ok: true };
  });
}

/* ---------- 선생님: 결과 집계 ---------- */

function getSessionResults(teacherId, sessionId) {
  var owned = findSessionOwned_(teacherId, sessionId);
  if (!owned) return authFail_();
  var session = buildSessionDetail_(owned);

  var hasText = session.questions.some(function (q) { return q.type === 'text'; });

  var byStudent = {};
  session.students.forEach(function (st) {
    byStudent[st.studentId] = {
      studentId: st.studentId, name: st.name, groupName: st.groupName,
      receivedCount: 0,   // 받은 평가 수
      givenCount: 0,      // 적은 평가 수
      scoreAverages: [],
      comments: [],       // 받은 서술형 답변
      receivedChars: 0,   // 받은 글자수
      writtenChars: 0     // 적은 글자수
    };
  });

  var sums = {};
  var counts = {};
  var responses = sheetToObjects_(getSheet_('Responses')).filter(function (r) { return r.sessionId === sessionId; });

  responses.forEach(function (r) {
    var targetId = String(r.targetStudentId);
    var evaluatorId = String(r.evaluatorStudentId);
    var target = byStudent[targetId];
    var evaluator = byStudent[evaluatorId];
    var answers;
    try { answers = JSON.parse(r.answersJson || '[]'); } catch (e) { answers = []; }

    if (target) target.receivedCount++;
    if (evaluator) evaluator.givenCount++;

    session.questions.forEach(function (q, qi) {
      var ans = answers[qi];
      if (q.type === 'score') {
        if (target) {
          var val = Number(ans);
          if (!isNaN(val) && ans !== '' && ans !== null) {
            sums[targetId] = sums[targetId] || {};
            counts[targetId] = counts[targetId] || {};
            sums[targetId][qi] = (sums[targetId][qi] || 0) + val;
            counts[targetId][qi] = (counts[targetId][qi] || 0) + 1;
          }
        }
      } else if (q.type === 'text') {
        var textLen = ans ? String(ans).trim().length : 0;
        if (target && textLen > 0) {
          target.comments.push(String(ans).trim());
          target.receivedChars += textLen;
        }
        if (evaluator) evaluator.writtenChars += textLen;
      }
    });
  });

  var results = Object.keys(byStudent).map(function (sid) {
    var t = byStudent[sid];
    t.scoreAverages = session.questions.map(function (q, qi) {
      if (q.type !== 'score') return null;
      var s = (sums[sid] && sums[sid][qi]) || 0;
      var c = (counts[sid] && counts[sid][qi]) || 0;
      return c > 0 ? Math.round((s / c) * 10) / 10 : null;
    });
    return t;
  });

  return { ok: true, questions: session.questions, hasText: hasText, results: results };
}
