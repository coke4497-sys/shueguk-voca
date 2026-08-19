/***************************************************************
 * 슈국 어휘 테스트 — 결과 수집/조회/삭제 Apps Script
 * ------------------------------------------------------------
 *  - doPost : 학생 test.html 이 보낸 결과를 시트에 저장
 *             (시트의 "헤더 이름"에 맞춰 저장 → 칼럼 순서가 달라도 안전)
 *  - doGet  : 교사 대시보드가 결과를 JSONP 로 읽어가는 통로
 *             · action=delete → 선택한 행 삭제 (대시보드 "선택 삭제")
 *
 *  ▶ 설치 / 갱신
 *    1) 결과 시트(「2026 어휘, 시작이 반이다」) → [확장 프로그램 → Apps Script]
 *    2) Code.gs 에 이 내용을 붙여넣고 저장(Ctrl+S)
 *    3) [배포 → 배포 관리 → ✏️ → 버전: 새 버전 → 배포]  (주소 유지)
 *    4) 액세스 권한: "모든 사용자"
 *    ※ 삭제 기능을 쓰려면 이 새 버전으로 "반드시" 재배포해야 합니다.
 *
 *  ▶ 권장: 기존 시트에 칼럼이 어긋난 옛 데이터가 있으면, 시트를 비우고
 *          시작하세요. 다음 제출 때 아래 HEADERS 헤더가 자동 생성됩니다.
 ***************************************************************/

var ACCESS_KEY = 'shueguk2026';   // 대시보드와 동일하게 유지
var SHEET_NAME = '';              // 결과 시트 이름. 비워두면 첫 번째 시트를 사용
var HEADERS = ['time', 'name', 'school', 'grade', 'phone4', 'round', 'score', 'details'];

// 리포트 Apps Script의 열린 주차 조회 주소 (test.html의 STATUS_URL과 동일)
var STATUS_URL = 'https://script.google.com/macros/s/AKfycbzhCncBwn-JlqXARC3wfrWUCuNHzlNK2df0bdhx-w78Xr8mzYUcIYZOJdRi9N4bHtsb/exec?action=vocaStatus';

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

/* 열린 주차 조회 (60초 캐시 — 제출마다 리포트 스크립트를 부르지 않도록) */
function vocaStatus_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('vocaStatus');
  if (hit) return JSON.parse(hit);
  var st = JSON.parse(UrlFetchApp.fetch(STATUS_URL, { muteHttpExceptions: true }).getContentText());
  cache.put('vocaStatus', JSON.stringify(st), 60);
  return st;
}

/* 학생 제출 (test.html → fetch POST) — 헤더 이름에 맞춰 저장 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    // 지난 주차 제출 차단 — 열린 주차와 다르면 기록하지 않는다.
    // (test.html도 제출 시점에 같은 확인을 하지만, 확인이 실패하면 열어주므로 여기가 최종 방어선.
    //  주차 조회 자체가 실패하면 정상 제출을 잃지 않도록 받아준다.)
    try {
      var st = vocaStatus_();
      if (st && st.result === 'success') {
        if (st.open === false) return json_({ ok: false, error: 'closed' });
        if (st.week >= 1 && String(st.week) !== String(data.round || '').trim()) {
          return json_({ ok: false, error: 'wrong_week', week: st.week });
        }
      }
    } catch (ignore) {}
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
    try { CacheService.getScriptCache().remove('takenIdx'); } catch (e) {}
    clearLiteCache_();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* 교사 대시보드 (JSONP GET) — 조회 / 삭제 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback || '';
  var payload;
  if (p.action === 'taken') {
    // 학생 본인 확인용: 이름+전화4+주차가 맞는 결과가 있으면 taken:true (키 불필요, 데이터 비노출)
    payload = checkTaken_(p.name, p.phone4, p.round);
  } else if (p.key !== ACCESS_KEY) {
    payload = { ok: false, error: 'unauthorized' };
  } else if (p.action === 'delete') {
    payload = deleteRows_(p.ids || '');
  } else if (p.action === 'detail') {
    payload = detailRow_(p.row, p.sig);
  } else if (p.lite === '1') {
    // lite=1: 목록에서 상세(details)를 뺀다 — 대시보드 첫 화면용(응답 1MB→수십 KB).
    // 상세는 행별 action=detail, CSV 내보내기는 lite 없이 전체 조회.
    // 접속이 몰릴 때 매번 시트 전체를 읽지 않도록 결과를 gzip으로 60초 캐시(제출·삭제 시 즉시 무효화).
    var out2 = liteListJson_();
    if (cb) {
      return ContentService.createTextOutput(cb + '(' + out2 + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(out2).setMimeType(ContentService.MimeType.JSON);
  } else {
    payload = { ok: true, rows: readRows_(false) };
  }
  var out = JSON.stringify(payload);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + out + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(out)
    .setMimeType(ContentService.MimeType.JSON);
}

/* 시트 → 행 객체 배열. 헤더 이름 기준으로 매핑.
 * 각 행에 시트 행번호(_row, 1-based)와 내용 해시(_sig)를 덧붙여
 * 대시보드가 안전하게 삭제 대상을 지정할 수 있게 한다. */
var LITE_CACHE_KEY = 'liteRowsGz';
function liteListJson_() {
  var cache = CacheService.getScriptCache();
  try {
    var hit = cache.get(LITE_CACHE_KEY);
    if (hit) return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(hit), 'application/x-gzip')).getDataAsString();
  } catch (e) {}
  var json = JSON.stringify({ ok: true, rows: readRows_(true) });
  try {
    var gz = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/octet-stream')).getBytes());
    cache.put(LITE_CACHE_KEY, gz, 60);   // 100KB 초과 등이면 캐시 없이 진행
  } catch (e2) {}
  return json;
}
function clearLiteCache_() {
  try { CacheService.getScriptCache().remove(LITE_CACHE_KEY); } catch (e) {}
}

function readRows_(lite) {
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var keys = values[0].map(canon_);
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var obj = {};
    for (var j = 0; j < keys.length; j++) obj[keys[j] || ('col' + j)] = values[i][j];
    if (lite) obj.details = '';   // 상세는 빼고 보냄 (_sig는 상세 포함 전체 행으로 계산 — 삭제 대조용)
    obj._row = i + 1;             // 시트상의 실제 행번호 (헤더가 1행)
    obj._sig = rowSig_(values[i]); // 행 내용 해시 (삭제 시 대조용)
    rows.push(obj);
  }
  return rows;
}

/* 한 행의 상세(details)만 반환 — 대시보드 '상세' 버튼용. 해시가 다르면(그 사이 변경/삭제) 오류. */
function detailRow_(rowNo, sig) {
  rowNo = parseInt(rowNo, 10);
  var sh = getSheet_();
  if (!(rowNo >= 2) || rowNo > sh.getLastRow()) return { ok: false, error: 'gone' };
  var lastCol = sh.getLastColumn();
  var keys = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(canon_);
  var v = sh.getRange(rowNo, 1, 1, lastCol).getValues()[0];
  if (rowSig_(v) !== ('' + (sig == null ? '' : sig))) return { ok: false, error: 'changed' };
  var iD = keys.indexOf('details');
  return { ok: true, row: rowNo, details: iD >= 0 ? '' + (v[iD] == null ? '' : v[iD]) : '' };
}

/* 응시 여부 조회용 [이름, 전화4, 주차] 목록 — 60초 캐시.
 * 슈퍼스타 개별 페이지가 열릴 때마다 taken 조회가 오는데, 접속이 몰리면
 * 매번 시트 전체를 읽다가 서버가 밀리므로 압축 목록만 캐시해 둔다.
 * (새 제출이 오면 doPost 에서 즉시 캐시를 비워 최신을 유지) */
function takenIndex_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('takenIdx');
  if (hit) return JSON.parse(hit);
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  var list = [];
  if (values.length >= 2) {
    var keys = values[0].map(canon_);
    var iName = keys.indexOf('name'), iPhone = keys.indexOf('phone4'), iRound = keys.indexOf('round');
    for (var i = 1; i < values.length; i++) {
      list.push([
        ('' + (iName  >= 0 && values[i][iName]  != null ? values[i][iName]  : '')).trim(),
        (iPhone >= 0 ? ('' + (values[i][iPhone] == null ? '' : values[i][iPhone])).trim() : null),
        ('' + (iRound >= 0 && values[i][iRound] != null ? values[i][iRound] : '')).trim()
      ]);
    }
  }
  try { cache.put('takenIdx', JSON.stringify(list), 60); } catch (e) {}  // 100KB 초과 등이면 캐시 없이 진행
  return list;
}

/* 학생 본인 응시 여부 확인 (이름+주차, 전화4는 보조). 데이터는 반환하지 않고 boolean 만.
 * ※ 학생이 테스트에 입력하는 전화4(본인 번호)와 개인 페이지가 보내는 전화4(학생ID=부모님
 *   번호 뒤 4자리)가 다를 수 있어, 전화4 불일치를 이유로 응시 기록을 버리지 않는다.
 *   이름+주차가 일치하면 응시로 인정(전화4까지 일치하면 즉시 확정). */
function checkTaken_(name, phone4, round) {
  name = ('' + (name == null ? '' : name)).trim();
  phone4 = ('' + (phone4 == null ? '' : phone4)).trim();
  round = ('' + (round == null ? '' : round)).trim();
  if (!name || !round) return { ok: true, taken: false };
  var list = takenIndex_();
  var nameMatch = false;
  for (var i = 0; i < list.length; i++) {
    if (list[i][2] !== round) continue;
    if (list[i][0] !== name) continue;
    nameMatch = true;   // 이름+주차 일치 → 응시로 인정
    if (phone4 && list[i][1] != null) {
      var ph = list[i][1];
      if (!ph || ph === phone4) return { ok: true, taken: true };   // 전화4까지 일치 → 즉시 확정
    } else {
      return { ok: true, taken: true };
    }
  }
  return { ok: true, taken: nameMatch };
}

/* 선택 행 삭제. ids = "행번호:해시,행번호:해시,..."
 * - 해시가 현재 행 내용과 일치할 때만 삭제 (그 사이 변경되면 건너뜀)
 * - 행번호 내림차순으로 삭제해 인덱스 밀림 방지 */
function deleteRows_(idsStr) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok: false, error: 'busy' }; }
  try {
    var sh = getSheet_();
    var values = sh.getDataRange().getValues();   // 삭제 전 스냅샷
    var want = {};
    ('' + idsStr).split(',').forEach(function (tok) {
      tok = tok.trim(); if (!tok) return;
      var idx = tok.indexOf(':');
      var rn = parseInt(idx >= 0 ? tok.slice(0, idx) : tok, 10);
      if (rn >= 2) want[rn] = idx >= 0 ? tok.slice(idx + 1) : '';
    });
    var nums = Object.keys(want).map(Number).sort(function (a, b) { return b - a; });
    var deleted = 0, skipped = 0;
    nums.forEach(function (rn) {
      if (rn > values.length) { skipped++; return; }
      if (rowSig_(values[rn - 1]) === want[rn]) { sh.deleteRow(rn); deleted++; }
      else skipped++;
    });
    try { CacheService.getScriptCache().remove('takenIdx'); } catch (e2) {}
    clearLiteCache_();
    return { ok: true, deleted: deleted, skipped: skipped };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/* 행 내용 해시 (MD5 앞 10자리) — Date는 epoch ms 로 정규화 */
function rowSig_(vals) {
  var s = vals.map(function (v) {
    if (Object.prototype.toString.call(v) === '[object Date]') return v.getTime();
    return '' + v;
  }).join('');
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < raw.length; i++) { var b = (raw[i] + 256) % 256; hex += ('0' + b.toString(16)).slice(-2); }
  return hex.slice(0, 10);
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
/** 주간 어휘 테스트 — 반점/키워드 채점 소급 보정 (1회 실행용)
 *  「2026 어휘, 시작이 반이다(응답)」 시트의 확장 프로그램 > Apps Script에 붙여넣고
 *  [반점보정_실행] 함수를 실행하세요. 실행 후 이 코드는 지워도 됩니다.
 *  각 행의 이름·기존 점수가 일치할 때만 수정하며, 결과는 실행 로그에 남습니다. */
function 반점보정_실행() {
  var FIXES = [
 {
  "row": 393,
  "name": "이주원a",
  "oldScore": "8 / 17",
  "newScore": "10 / 17",
  "repl": [
   [
    "6. ✗ 정신이 얼떨떨하여 어찌할 바를 모르는 모양 → 정답: 얼떨떨하여, 모르는",
    "6. ✓ 정신이 얼떨떨하여 어찌할 바를 모르는 모양"
   ],
   [
    "9. ✗ 성격따위가 밝고 명랑하여 구김살이 없다 → 정답: 명랑, 구김살",
    "9. ✓ 성격따위가 밝고 명랑하여 구김살이 없다"
   ]
  ]
 },
 {
  "row": 456,
  "name": "이태림",
  "oldScore": "11 / 17",
  "newScore": "12 / 17",
  "repl": [
   [
    "9. ✗ 성격 따위가 밝고 명랑하여 구김살이 없다 → 정답: 명랑, 구김살",
    "9. ✓ 성격 따위가 밝고 명랑하여 구김살이 없다"
   ]
  ]
 },
 {
  "row": 475,
  "name": "최은서b",
  "oldScore": "15 / 17",
  "newScore": "17 / 17",
  "repl": [
   [
    "6. ✗ 얼떨떨하여/모르는 → 정답: 얼떨떨하여, 모르는",
    "6. ✓ 얼떨떨하여/모르는"
   ],
   [
    "9. ✗ 명랑하여/구김살 → 정답: 명랑, 구김살",
    "9. ✓ 명랑하여/구김살"
   ]
  ]
 },
 {
  "row": 806,
  "name": "김지아",
  "oldScore": "13 / 17",
  "newScore": "14 / 17",
  "repl": [
   [
    "4. ✗ 한계 최후의 → 정답: 한계, 최후",
    "4. ✓ 한계 최후의"
   ]
  ]
 },
 {
  "row": 860,
  "name": "유채아",
  "oldScore": "10 / 17",
  "newScore": "11 / 17",
  "repl": [
   [
    "4. ✗ 궁극의  한계 도달할 수 있는 최후의 단계 → 정답: 한계, 최후",
    "4. ✓ 궁극의  한계 도달할 수 있는 최후의 단계"
   ]
  ]
 },
 {
  "row": 881,
  "name": "나세현",
  "oldScore": "16 / 17",
  "newScore": "17 / 17",
  "repl": [
   [
    "4. ✗ 한계, 최후의 → 정답: 한계, 최후",
    "4. ✓ 한계, 최후의"
   ]
  ]
 },
 {
  "row": 944,
  "name": "김건우",
  "oldScore": "14 / 17",
  "newScore": "15 / 17",
  "repl": [
   [
    "4. ✗ 한계,최후의 → 정답: 한계, 최후",
    "4. ✓ 한계,최후의"
   ]
  ]
 },
 {
  "row": 1094,
  "name": "김상효",
  "oldScore": "14 / 17",
  "newScore": "15 / 17",
  "repl": [
   [
    "12. ✗ 알,꿩 → 정답: 꿩, 알",
    "12. ✓ 알,꿩"
   ]
  ]
 },
 {
  "row": 1118,
  "name": "한세준",
  "oldScore": "12 / 17",
  "newScore": "15 / 17",
  "repl": [
   [
    "11. ✗ 누이 좋고 매부 좋다 → 정답: 누이, 매부",
    "11. ✓ 누이 좋고 매부 좋다"
   ],
   [
    "12. ✗ 꿩 먹고 알 먹기 → 정답: 꿩, 알",
    "12. ✓ 꿩 먹고 알 먹기"
   ],
   [
    "13. ✗ 도랑 치고 가재 잡는다 → 정답: 도랑, 가재",
    "13. ✓ 도랑 치고 가재 잡는다"
   ]
  ]
 },
 {
  "row": 1131,
  "name": "송어진",
  "oldScore": "10 / 17",
  "newScore": "13 / 17",
  "repl": [
   [
    "11. ✗ 누이 좋고 매부 좋다 → 정답: 누이, 매부",
    "11. ✓ 누이 좋고 매부 좋다"
   ],
   [
    "12. ✗ 꿩먹고 알먹기 → 정답: 꿩, 알",
    "12. ✓ 꿩먹고 알먹기"
   ],
   [
    "13. ✗ 도랑치고 가재잡는다 → 정답: 도랑, 가재",
    "13. ✓ 도랑치고 가재잡는다"
   ]
  ]
 },
 {
  "row": 1146,
  "name": "김유하",
  "oldScore": "12 / 17",
  "newScore": "15 / 17",
  "repl": [
   [
    "11. ✗ 누이 좋고 매부 좋다. → 정답: 누이, 매부",
    "11. ✓ 누이 좋고 매부 좋다."
   ],
   [
    "12. ✗ 꿩 먹고 알 먹기 → 정답: 꿩, 알",
    "12. ✓ 꿩 먹고 알 먹기"
   ],
   [
    "13. ✗ 도랑 치고 가재 잡는다 → 정답: 도랑, 가재",
    "13. ✓ 도랑 치고 가재 잡는다"
   ]
  ]
 },
 {
  "row": 1148,
  "name": "김도연",
  "oldScore": "13 / 17",
  "newScore": "16 / 17",
  "repl": [
   [
    "11. ✗ 누이 좋고 매부 좋다. → 정답: 누이, 매부",
    "11. ✓ 누이 좋고 매부 좋다."
   ],
   [
    "12. ✗ 꿩 먹고 알 먹기 → 정답: 꿩, 알",
    "12. ✓ 꿩 먹고 알 먹기"
   ],
   [
    "13. ✗ 도랑 치고 가재 잡는다. → 정답: 도랑, 가재",
    "13. ✓ 도랑 치고 가재 잡는다."
   ]
  ]
 },
 {
  "row": 1157,
  "name": "배민서",
  "oldScore": "11 / 17",
  "newScore": "14 / 17",
  "repl": [
   [
    "11. ✗ 누이 좋고 매부 좋다 → 정답: 누이, 매부",
    "11. ✓ 누이 좋고 매부 좋다"
   ],
   [
    "12. ✗ 꿩 먹고 알 먹기 → 정답: 꿩, 알",
    "12. ✓ 꿩 먹고 알 먹기"
   ],
   [
    "13. ✗ 도랑치고 가재잡는다 → 정답: 도랑, 가재",
    "13. ✓ 도랑치고 가재잡는다"
   ]
  ]
 },
 {
  "row": 1172,
  "name": "김준섭",
  "oldScore": "9 / 17",
  "newScore": "12 / 17",
  "repl": [
   [
    "11. ✗ 누이 좋고 매부 좋다 → 정답: 누이, 매부",
    "11. ✓ 누이 좋고 매부 좋다"
   ],
   [
    "12. ✗ 꿩 먹고 알먹기 → 정답: 꿩, 알",
    "12. ✓ 꿩 먹고 알먹기"
   ],
   [
    "13. ✗ 도랑 치고 가재 잡는다 → 정답: 도랑, 가재",
    "13. ✓ 도랑 치고 가재 잡는다"
   ]
  ]
 },
 {
  "row": 1186,
  "name": "김상휘",
  "oldScore": "12 / 17",
  "newScore": "15 / 17",
  "repl": [
   [
    "11. ✗ 누이 좋고 매부 좋다 → 정답: 누이, 매부",
    "11. ✓ 누이 좋고 매부 좋다"
   ],
   [
    "12. ✗ 꿩 먹고 알 먹기 → 정답: 꿩, 알",
    "12. ✓ 꿩 먹고 알 먹기"
   ],
   [
    "13. ✗ 도랑 치고 가재 잡는다 → 정답: 도랑, 가재",
    "13. ✓ 도랑 치고 가재 잡는다"
   ]
  ]
 },
 {
  "row": 1305,
  "name": "백예음",
  "oldScore": "13 / 17",
  "newScore": "16 / 17",
  "repl": [
   [
    "11. ✗ 누이 좋고 매부 좋다 → 정답: 누이, 매부",
    "11. ✓ 누이 좋고 매부 좋다"
   ],
   [
    "12. ✗ 꿩 먹고 알 먹기 → 정답: 꿩, 알",
    "12. ✓ 꿩 먹고 알 먹기"
   ],
   [
    "13. ✗ 도랑 치고 가재 잡는다 → 정답: 도랑, 가재",
    "13. ✓ 도랑 치고 가재 잡는다"
   ]
  ]
 },
 {
  "row": 1330,
  "name": "최현준",
  "oldScore": "12 / 17",
  "newScore": "15 / 17",
  "repl": [
   [
    "11. ✗ 누이 좋고 매부 좋다 → 정답: 누이, 매부",
    "11. ✓ 누이 좋고 매부 좋다"
   ],
   [
    "12. ✗ 꿩 먹고 알 먹기 → 정답: 꿩, 알",
    "12. ✓ 꿩 먹고 알 먹기"
   ],
   [
    "13. ✗ 도랑 치고 가재 잡는다 → 정답: 도랑, 가재",
    "13. ✓ 도랑 치고 가재 잡는다"
   ]
  ]
 },
 {
  "row": 1337,
  "name": "이서린",
  "oldScore": "13 / 17",
  "newScore": "16 / 17",
  "repl": [
   [
    "11. ✗ 누이 좋고 매부 좋다 → 정답: 누이, 매부",
    "11. ✓ 누이 좋고 매부 좋다"
   ],
   [
    "12. ✗ 꿩 먹고 알 먹기 → 정답: 꿩, 알",
    "12. ✓ 꿩 먹고 알 먹기"
   ],
   [
    "13. ✗ 도랑 치고 가재 잡는다 → 정답: 도랑, 가재",
    "13. ✓ 도랑 치고 가재 잡는다"
   ]
  ]
 },
 {
  "row": 1405,
  "name": "양현서",
  "oldScore": "11 / 17",
  "newScore": "14 / 17",
  "repl": [
   [
    "11. ✗ 누이.매부 → 정답: 누이, 매부",
    "11. ✓ 누이.매부"
   ],
   [
    "12. ✗ 꿩.알 → 정답: 꿩, 알",
    "12. ✓ 꿩.알"
   ],
   [
    "13. ✗ 도랑.가재 → 정답: 도랑, 가재",
    "13. ✓ 도랑.가재"
   ]
  ]
 },
 {
  "row": 1406,
  "name": "조수현",
  "oldScore": "11 / 17",
  "newScore": "14 / 17",
  "repl": [
   [
    "11. ✗ 누이 좋고 매부 좋다 → 정답: 누이, 매부",
    "11. ✓ 누이 좋고 매부 좋다"
   ],
   [
    "12. ✗ 꿩 먹고 알 먹기 → 정답: 꿩, 알",
    "12. ✓ 꿩 먹고 알 먹기"
   ],
   [
    "13. ✗ 도랑 치고 가재 잡는다 → 정답: 도랑, 가재",
    "13. ✓ 도랑 치고 가재 잡는다"
   ]
  ]
 }
];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var done = 0, skip = [];
  FIXES.forEach(function (f) {
    var vals = sh.getRange(f.row, 1, 1, 11).getValues()[0];
    var name = String(vals[1]).trim();
    var score = String(vals[5]).replace(/\s+/g, '');
    if (name !== f.name || score !== f.oldScore.replace(/\s+/g, '')) {
      skip.push(f.row + '행 ' + f.name + ' (현재: ' + name + ' / ' + vals[5] + ') — 불일치, 건너뜀');
      return;
    }
    var details = String(vals[10]);
    var ok = true;
    f.repl.forEach(function (p) { if (details.indexOf(p[0]) < 0) ok = false; });
    if (!ok) { skip.push(f.row + '행 ' + f.name + ' — 상세 텍스트 불일치, 건너뜀'); return; }
    f.repl.forEach(function (p) { details = details.replace(p[0], p[1]); });
    sh.getRange(f.row, 6).setValue(f.newScore);
    sh.getRange(f.row, 11).setValue(details);
    done++;
  });
  Logger.log('보정 완료: ' + done + '건 / 대상 ' + FIXES.length + '건');
  skip.forEach(function (s) { Logger.log('건너뜀: ' + s); });
  SpreadsheetApp.getUi ? null : null;
}
