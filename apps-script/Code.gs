/***************************************************************
 * 슈국 어휘 테스트 — 결과 수집/조회 Apps Script
 * ------------------------------------------------------------
 *  - doPost : 학생 test.html 이 보낸 결과를 시트에 저장
 *             (시트의 "헤더 이름"에 맞춰 저장 → 칼럼 순서가 달라도 안전)
 *  - doGet  : 교사 대시보드가 결과를 JSONP 로 읽어가는 통로
 *
 *  ▶ 설치
 *    1) 결과 시트(「2026 어휘, 시작이 반이다」) → [확장 프로그램 → Apps Script]
 *    2) Code.gs 에 이 내용을 붙여넣고 저장(Ctrl+S)
 *    3) [배포 → 배포 관리 → ✏️ → 버전: 새 버전 → 배포]  (주소 유지)
 *    4) 액세스 권한: "모든 사용자"
 *
 *  ▶ 권장: 기존 시트에 칼럼이 어긋난 옛 데이터가 있으면, 시트를 비우고
 *          시작하세요. 다음 제출 때 아래 HEADERS 헤더가 자동 생성됩니다.
 ***************************************************************/

var ACCESS_KEY = 'shueguk2026';   // 대시보드와 동일하게 유지
var SHEET_NAME = '';              // 결과 시트 이름. 비워두면 첫 번째 시트를 사용
var HEADERS = ['time', 'name', 'school', 'grade', 'phone4', 'round', 'score', 'details'];

/* 결과 시트 가져오기 (완전히 빈 시트면 헤더 생성) */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : null;
  if (!sh) sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

/* 헤더 행을 읽어 { 표준키: 0-based 열번호 } 로 변환 */
function headerMap_(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return {};
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < hdr.length; i++) {
    var k = canon_(hdr[i]);
    if (k && map[k] === undefined) map[k] = i;
  }
  return map;
}

/* 학생 제출 (test.html → fetch POST) — 헤더 이름에 맞춰 저장 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sh = getSheet_();
    var map = headerMap_(sh);
    // 빠진 항목이 있으면 헤더 끝에 칼럼 추가
    HEADERS.forEach(function (h) {
      if (map[h] === undefined) {
        var col = sh.getLastColumn() + 1;
        sh.getRange(1, col).setValue(h);
        map[h] = col - 1;
      }
    });
    var width = sh.getLastColumn();
    var row = [];
    for (var i = 0; i < width; i++) row.push('');
    HEADERS.forEach(function (h) { row[map[h]] = (data[h] != null ? data[h] : ''); });
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

/* 시트 → 행 객체 배열. 헤더 이름 기준으로 매핑 */
function readRows_() {
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var keys = values[0].map(canon_);
  var rows = [];
  for (var i = 1; i < values.length; i++) {
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
    'phone4': 'phone4', '전화': 'phone4', '전화4': 'phone4', '전화뒤4': 'phone4', '전화뒷4': 'phone4', '휴대전화': 'phone4', '식별번호': 'phone4',
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
