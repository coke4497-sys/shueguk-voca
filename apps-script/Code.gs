/***************************************************************
 * 슈국 어휘 테스트 — 결과 수집/조회 Apps Script
 * ------------------------------------------------------------
 *  - doPost : 학생 test.html 이 보낸 결과를 시트에 저장
 *  - doGet  : 교사 대시보드(shueguk-teacher-dashboard.html)가
 *             결과를 JSONP 로 읽어가는 통로
 *
 *  ▶ 설치
 *    1) 결과가 쌓이는 구글 시트를 열고 [확장 프로그램 → Apps Script]
 *    2) Code.gs 에 이 내용을 붙여넣고 저장(Ctrl+S)
 *    3) [배포 → 배포 관리 → (연필 ✏️) → 버전: 새 버전 → 배포]
 *       (이렇게 해야 기존 /exec 주소가 그대로 유지됩니다)
 *    4) 액세스 권한: "모든 사용자"
 *
 *  ▶ 주의
 *    - ACCESS_KEY 는 대시보드의 ACCESS_KEY 와 반드시 동일해야 함.
 *    - SCRIPT_URL(=/exec) 은 test.html / 대시보드의 SCRIPT_URL 과 같아야 함.
 ***************************************************************/

var ACCESS_KEY = 'shueguk2026';   // 대시보드와 동일하게 유지
var SHEET_NAME = '';              // 결과 시트 이름. 비워두면 첫 번째 시트를 사용
var HEADERS = ['time', 'name', 'school', 'grade', 'round', 'score', 'details'];

/* 결과 시트 가져오기 (없으면 헤더 만들기) */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : null;
  if (!sh) sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);   // 완전히 빈 시트일 때만 헤더 추가
  return sh;
}

/* 학생 제출 (test.html → fetch POST) */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sh = getSheet_();
    var row = HEADERS.map(function (h) { return data[h] != null ? data[h] : ''; });
    sh.appendRow(row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* 교사 대시보드 (JSONP GET) */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback || '';
  var payload = (p.key !== ACCESS_KEY)
    ? { ok: false, error: 'unauthorized' }
    : { ok: true, rows: readRows_() };
  var out = JSON.stringify(payload);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + out + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(out)
    .setMimeType(ContentService.MimeType.JSON);
}

/* 시트 → 행 객체 배열. 헤더가 있으면 헤더 기준, 없으면 HEADERS 순서로 해석 */
function readRows_() {
  var values = getSheet_().getDataRange().getValues();
  if (!values.length) return [];
  var known = { time: 1, name: 1, school: 1, grade: 1, round: 1, score: 1, details: 1 };
  var head = values[0].map(canon_);
  var hasHeader = head.some(function (k) { return known[k]; });
  var start = hasHeader ? 1 : 0;
  var keys = hasHeader ? head : HEADERS.slice();
  var rows = [];
  for (var i = start; i < values.length; i++) {
    var obj = {};
    for (var j = 0; j < keys.length; j++) obj[keys[j] || ('col' + j)] = values[i][j];
    rows.push(obj);
  }
  return rows;
}

/* 헤더 이름 표준화 (한글/영문 모두 허용) */
function canon_(h) {
  h = ('' + h).trim().toLowerCase();
  var map = {
    'time': 'time', '시각': 'time', '제출시각': 'time', 'timestamp': 'time', '타임스탬프': 'time',
    'name': 'name', '이름': 'name',
    'school': 'school', '학교': 'school',
    'grade': 'grade', '학년': 'grade',
    'round': 'round', '주차': 'round', '회차': 'round',
    'score': 'score', '점수': 'score',
    'details': 'details', '상세': 'details'
  };
  return map[h] || h;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
