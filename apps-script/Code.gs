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

/* ── 편집기에서 한 번만 실행하는 권한 승인 함수 ──────────────────
 * 이 스크립트는 UrlFetchApp(외부 서비스 연결)을 쓴다:
 *   - vocaStatus_  : 열린 주차 조회 → 지난 주차 제출 차단(#34)
 *   - sbMirrorVoca_: 수파베이스 미러 기록
 * 그런데 소유자가 그 권한을 승인하기 전까지 **배포본에서 조용히 실패한다**
 * (오류도 안 나고 그냥 아무 일도 일어나지 않는다 — 2026-08-25 확인:
 *  지난 주차 제출이 그대로 기록되고 있었고 미러도 안 들어갔다).
 *
 * 고치는 법: 편집기에서 이 함수를 한 번 실행 → 권한 요청이 뜨면 허용.
 * 그 뒤로는 배포본도 정상 동작한다(승인은 버전이 아니라 계정에 붙는다).
 * 아래 로그에 열린 주차와 미러 응답이 찍히면 성공. */
function 권한승인() {
  var out = [];
  try {
    out.push('열린 주차 조회: ' + UrlFetchApp.fetch(STATUS_URL, { muteHttpExceptions: true }).getContentText());
  } catch (e) { out.push('열린 주차 조회 실패: ' + e); }
  try {
    var res = UrlFetchApp.fetch(SB_REST + '/voca_results?select=id&limit=1', {
      headers: { 'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY }, muteHttpExceptions: true
    });
    out.push('미러 연결: ' + res.getResponseCode() + ' ' + res.getContentText());
  } catch (e2) { out.push('미러 연결 실패: ' + e2); }
  var msg = out.join('\n');
  Logger.log(msg);
  return msg;
}

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

/* ── 수파베이스 미러 (voca_results) ────────────────────────────────
 * 시트가 원본이고 이건 DB 이전 준비용 읽기 미러다.
 *
 * 예전에는 학생 브라우저(test.html)에서 넣었는데 두 가지로 새고 있었다
 * (2026-08-24 확인 — 시트 2,433건 vs 미러 2,425건, 빠진 8건 전부 당일 제출):
 *   1. 제출 직후 학생이 창을 닫으면 미러 요청이 취소된다
 *   2. 제출이 no-cors라 페이지는 백엔드 거절(closed·wrong_week)을 모른다 —
 *      시트에 없는 행이 미러에만 들어갔다
 * 이제 시트에 실제로 쓴 뒤 여기서(서버끼리) 넣으므로 둘 다 생기지 않는다.
 *
 * 실패해도 제출은 이미 시트에 저장된 뒤라 무해하다 — 일일 점검(audit_heal --voca)이
 * 시트 기준으로 복구한다. 잠깐 끊긴 경우를 대비해 한 번만 더 시도한다. */
var SB_REST = 'https://bangdbhqpphqqdwcledg.supabase.co/rest/v1';
var SB_KEY  = 'sb_publishable_dE9d1KIbpgYaQkaS2MSrlg_-7SiRJuT';

function sbMirrorVoca_(d) {
  var body = JSON.stringify([{
    ts:      sbStr_(d.time),     // 'yyyy-MM-dd HH:mm' — 페이지가 보낸 문자열 그대로 (시트와 일치)
    name:    sbStr_(d.name),
    school:  sbStr_(d.school),
    grade:   sbStr_(d.grade),
    phone4:  sbNum_(d.phone4),   // 시트가 숫자로 바꾸며 앞의 0을 지우므로 똑같이 맞춘다('0913'→'913')
    round:   sbNum_(d.round),
    score:   sbStr_(d.score),
    details: sbStr_(d.details)
  }]);
  for (var try_ = 0; try_ < 2; try_++) {
    try {
      var res = UrlFetchApp.fetch(SB_REST + '/voca_results', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY, 'Prefer': 'return=minimal' },
        payload: body,
        muteHttpExceptions: true
      });
      var code = res.getResponseCode();
      if (code >= 200 && code < 300) return true;
    } catch (e) {}
  }
  return false;
}
function sbStr_(v) { return '' + (v == null ? '' : v); }

/* 시트에 숫자로 저장되는 칸(전화 뒤 4자리·주차)을 시트와 같은 표기로 맞춘다.
 * 학생이 '0913'을 넣으면 시트는 숫자 913으로 바꿔 앞의 0을 잃는데,
 * 미러가 '0913'을 그대로 담으면 시트와 어긋나 일일 점검이 매번 고치게 된다
 * (2026-08-25 실제 발생 — 한예림 7주차). */
function sbNum_(v) {
  var t = ('' + (v == null ? '' : v)).trim();
  return /^\d+$/.test(t) ? '' + parseInt(t, 10) : t;
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
    try { sbMirrorVoca_(data); } catch (e3) {}   // 미러는 실패해도 제출에 영향 없음
    // 캐시를 비우지 않는다 — 새로 붙은 줄은 조회할 때 꼬리에서 읽어 잇는다.
    // (예전에는 여기서 비우는 바람에 제출이 몰리는 시간대에 캐시가 늘 비어 있었다)
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
  } else if (p.action === 'fresh') {
    // 시트를 손으로 고친 뒤처럼 캐시를 강제로 다시 만들 때 (?key=…&action=fresh)
    clearLiteCache_(); clearTakenCache_();
    payload = { ok: true, cleared: true };
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
var LITE_CACHE_N   = 'liteRowsGzN';    // 나눠 담은 조각 수
var LITE_CACHE_ROW = 'liteRowsGzLast'; // 그 캐시를 만들 때 시트의 마지막 행
var LITE_CHUNK     = 90000;            // CacheService 한 칸 상한(100KB)보다 여유 있게
var LITE_MAX_CHUNK = 12;               // 뒷정리·조회에 쓸 최대 조각 수
var LITE_TTL       = 21600;            // 6시간 (CacheService 최대). 새 제출은 꼬리만 읽어 붙이므로 길게 둔다
var LITE_TAIL_MAX  = 400;              // 캐시 이후 늘어난 줄이 이보다 많으면 통째로 다시 만든다

/* lite 목록 JSON — 시트를 통째로 다시 읽지 않는다.
 *
 * 예전에는 제출이 올 때마다(doPost) 캐시를 통째로 비웠다. 그래서 학생들이 제출하는
 * 저녁 시간대 — 교사가 결과를 확인하는 바로 그 시간 — 에는 캐시가 늘 비어 있어
 * 조회할 때마다 시트 2,400줄을 처음부터 읽었고, 6초에서 80초까지 걸렸다
 * (2026-08-24 측정. 전송량이 아니라 시트 읽기가 원인).
 *
 * 이제 캐시에 '만들 때의 마지막 행'을 함께 적어 두고, 그 뒤에 늘어난 줄만
 * 꼬리에서 읽어 이어 붙인다. 제출이 아무리 들어와도 캐시는 살아 있다.
 * 줄이 지워지면(행 번호가 밀린다) deleteRows_ 가 캐시를 비워 다시 만들게 한다. */
function liteListJson_() {
  var cache = CacheService.getScriptCache();
  var sh = getSheet_();
  var last = sh.getLastRow();
  var base = liteCacheRead_(cache);
  if (base) {
    if (base.lastRow === last) return base.json;                       // 그대로
    if (base.lastRow < last && (last - base.lastRow) <= LITE_TAIL_MAX) {
      var tail = readRowsRange_(sh, base.lastRow + 1, last, true);     // 새 줄만 읽어 붙임
      if (!tail.length) return base.json;
      var merged = appendRowsJson_(base.json, tail);
      if (merged) return merged;
    }
  }
  var json = JSON.stringify({ ok: true, rows: readRows_(true) });      // 통째로 다시
  liteCacheWrite_(cache, json, last);
  return json;
}

/* 이미 만들어 둔 목록 JSON 끝에 행들을 이어 붙인다 (다시 파싱하지 않도록 문자열로).
 * 모양이 예상과 다르면 null 을 돌려 호출부가 통째로 다시 만들게 한다. */
function appendRowsJson_(json, rows) {
  if (!json || json.slice(-2) !== ']}') return null;
  var head = json.slice(0, -2);
  var inner = JSON.stringify(rows);
  inner = inner.slice(1, -1);                                          // 바깥 [ ] 제거
  return head + (head.slice(-1) === '[' ? '' : ',') + inner + ']}';
}

function liteCacheRead_(cache) {
  try {
    var n = parseInt(cache.get(LITE_CACHE_N), 10);
    var lastRow = parseInt(cache.get(LITE_CACHE_ROW), 10);
    if (!(n >= 1) || !(lastRow >= 1)) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(LITE_CACHE_KEY + i);
    var got = cache.getAll(keys), gz = '';
    for (var j = 0; j < n; j++) {
      var part = got[LITE_CACHE_KEY + j];
      if (part == null) return null;                                   // 한 조각이라도 만료면 무시
      gz += part;
    }
    return {
      json: Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(gz), 'application/x-gzip')).getDataAsString(),
      lastRow: lastRow
    };
  } catch (e) { return null; }
}

/* 캐시 한 칸은 100KB까지라 90KB씩 조각내 여러 칸에 담는다.
 * (한 칸을 넘기면 put 이 조용히 실패해 모든 조회가 시트 전체 읽기로 되돌아간다) */
function liteCacheWrite_(cache, json, lastRow) {
  try {
    var gz = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/octet-stream')).getBytes());
    var cnt = Math.ceil(gz.length / LITE_CHUNK);
    if (!(cnt >= 1 && cnt <= LITE_MAX_CHUNK)) return;
    var put = {};
    for (var k = 0; k < cnt; k++) put[LITE_CACHE_KEY + k] = gz.substr(k * LITE_CHUNK, LITE_CHUNK);
    cache.putAll(put, LITE_TTL);
    cache.put(LITE_CACHE_N, '' + cnt, LITE_TTL);
    cache.put(LITE_CACHE_ROW, '' + lastRow, LITE_TTL);
  } catch (e) {}
}

function clearLiteCache_() {
  try {
    var keys = [LITE_CACHE_N, LITE_CACHE_ROW];
    for (var i = 0; i < LITE_MAX_CHUNK; i++) keys.push(LITE_CACHE_KEY + i);
    CacheService.getScriptCache().removeAll(keys);
  } catch (e) {}
}

/* lite 목록에 담을 열 — 대시보드가 실제로 읽는 것만.
 * 시트에 남아 있는 옛 빈 열('회차(숫자만 입력)'·'문항별 결과'·'제출 시각')이
 * 응답의 3분의 1을 차지하고 있었다(603KB → 400KB). */
var LITE_KEYS = { time: 1, name: 1, school: 1, grade: 1, phone4: 1, round: 1, score: 1 };

function readRows_(lite) {
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  return buildRows_(values[0].map(canon_), values.slice(1), 2, lite);
}

/* 시트의 일부 구간만 읽어 같은 모양으로 만든다 (새 제출 꼬리 읽기용). */
function readRowsRange_(sh, fromRow, toRow, lite) {
  if (!(toRow >= fromRow) || fromRow < 2) return [];
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return [];
  var keys = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(canon_);
  var values = sh.getRange(fromRow, 1, toRow - fromRow + 1, lastCol).getValues();
  return buildRows_(keys, values, fromRow, lite);
}

/* 값 배열 → 행 객체 배열. firstRow 는 values[0] 이 시트의 몇 번째 행인지. */
function buildRows_(keys, values, firstRow, lite) {
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var obj = {};
    for (var j = 0; j < keys.length; j++) {
      var k = keys[j] || ('col' + j);
      if (lite && !LITE_KEYS[k]) continue;   // 상세·빈 옛 열은 빼고 보냄
      obj[k] = values[i][j];
    }
    if (lite) obj.details = '';   // (_sig는 상세 포함 전체 행으로 계산 — 삭제 대조용)
    obj._row = firstRow + i;      // 시트상의 실제 행번호 (헤더가 1행)
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

/* 응시 여부 조회용 [이름, 전화4, 주차] 목록.
 * 슈퍼스타 개별 페이지가 열릴 때마다 taken 조회가 오는데, 접속이 몰리면
 * 매번 시트 전체를 읽다가 서버가 밀리므로 목록을 캐시해 둔다.
 * lite 목록과 같은 이유로 제출 때 캐시를 비우지 않고 꼬리만 읽어 잇는다
 * (학생 제출이 몰리는 시간이 곧 개별 페이지가 몰리는 시간이라, 비우면 캐시가 없는 것과 같다). */
var TAKEN_TTL      = 21600;   // 6시간
var TAKEN_TAIL_MAX = 400;
function takenIndex_() {
  var cache = CacheService.getScriptCache();
  var sh = getSheet_();
  var last = sh.getLastRow();
  var list = null, cachedLast = 0;
  try {
    var hit = cache.get('takenIdx'), hr = parseInt(cache.get('takenIdxRow'), 10);
    if (hit && hr >= 1) { list = JSON.parse(hit); cachedLast = hr; }
  } catch (e) { list = null; }
  if (list) {
    if (cachedLast === last) return list;
    if (cachedLast < last && (last - cachedLast) <= TAKEN_TAIL_MAX) {
      return list.concat(takenRange_(sh, cachedLast + 1, last));
    }
  }
  var out = takenRange_(sh, 2, last);
  // 100KB 초과 등이면 캐시 없이 진행
  try { cache.put('takenIdx', JSON.stringify(out), TAKEN_TTL); cache.put('takenIdxRow', '' + last, TAKEN_TTL); } catch (e2) {}
  return out;
}

function takenRange_(sh, fromRow, toRow) {
  if (!(toRow >= fromRow) || fromRow < 2) return [];
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return [];
  var keys = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(canon_);
  var iName = keys.indexOf('name'), iPhone = keys.indexOf('phone4'), iRound = keys.indexOf('round');
  var values = sh.getRange(fromRow, 1, toRow - fromRow + 1, lastCol).getValues();
  var list = [];
  for (var i = 0; i < values.length; i++) {
    list.push([
      ('' + (iName  >= 0 && values[i][iName]  != null ? values[i][iName]  : '')).trim(),
      (iPhone >= 0 ? ('' + (values[i][iPhone] == null ? '' : values[i][iPhone])).trim() : null),
      ('' + (iRound >= 0 && values[i][iRound] != null ? values[i][iRound] : '')).trim()
    ]);
  }
  return list;
}

function clearTakenCache_() {
  try { CacheService.getScriptCache().removeAll(['takenIdx', 'takenIdxRow']); } catch (e) {}
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
    clearTakenCache_();   // 행 번호가 밀리므로 두 캐시 모두 다시 만든다
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
