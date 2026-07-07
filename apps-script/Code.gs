/****************************************************************
 * 가족 가계부 — Google Apps Script 백엔드
 * ------------------------------------------------------------
 * 이 코드를 스프레드시트에 붙여 "웹 앱"으로 배포하면,
 * 웹앱(GitHub Pages)에서 이 시트로 직접 저장/조회할 수 있습니다.
 *
 * ▶ 배포 방법
 *   1) 저장할 구글 스프레드시트를 연다
 *   2) 상단 메뉴  확장 프로그램  →  Apps Script
 *   3) 기본 코드(Code.gs)를 지우고 이 파일 내용을 전부 붙여넣기 → 저장(💾)
 *   4) (선택) 왼쪽 함수 목록에서 setup 실행 → 시트 초기화 (권한 승인 1회)
 *   5) 오른쪽 위  배포  →  새 배포  →  유형: 웹 앱
 *        - 실행 계정: 나
 *        - 액세스 권한: 모든 사용자   ← 반드시 이걸로
 *   6) 배포 → 나오는 "웹 앱 URL" 복사  (https://script.google.com/macros/s/..../exec)
 *   7) 이 URL을 웹앱의 [설정]에 붙여넣고, 위 SECRET 과 똑같은 비밀번호도 입력하면 연결 완료
 *      (URL + 비밀번호를 둘 다 알아야만 접근 가능 → 실질적으로 두 분 전용)
 *
 * ※ 코드를 수정하면 매번 "배포 → 배포 관리 → 편집(연필) → 버전: 새 버전 → 배포"
 *    를 해야 반영됩니다. (URL은 그대로 유지됩니다)
 ****************************************************************/

// 비밀번호. 지금은 비워두면(''), 검사 없이 URL만으로 작동합니다.
// 나중에 여기에 비밀번호를 넣고 앱 [설정]에도 똑같이 넣으면 자동으로 잠깁니다.
var SECRET = '';

var HEADER = ['id', 'date', 'amount', 'category', 'user', 'card', 'memo'];
var META = '_meta';
var LISTS_KEY = '__lists__'; // _meta 시트의 특수 행: 항목/사용자/카드 목록 저장

/* ---------- HTTP 엔드포인트 ---------- */

function doGet(e) {
  var token = e && e.parameter ? e.parameter.token : '';
  return json(handle('all', {}, token));
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return json(handle(body.action, body, body.token));
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 모든 정상 응답: 월별 데이터 + 목록(항목/사용자/카드) 설정을 함께 반환
function respond() {
  return { ok: true, months: readAll(), config: readConfig() };
}

function handle(action, p, token) {
  if (SECRET && String(token) !== String(SECRET)) return { ok: false, error: 'unauthorized' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok: false, error: 'busy' }; }
  try {
    switch (action) {
      case 'all':          return respond();
      case 'addEntry':     addEntry(p.month, p.entry);        return respond();
      case 'deleteEntry':  deleteEntry(p.month, p.id);        return respond();
      case 'setBudget':    setBudget(p.month, p.budget);      return respond();
      case 'setCatBudget': setCatBudget(p.month, p.catBudgets); return respond();
      case 'setLists':     setLists(p.lists);                 return respond();
      default:             return { ok: false, error: 'unknown action: ' + action };
    }
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 시트 헬퍼 ---------- */

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function monthSheet(month, create) {
  var sh = ss().getSheetByName(month);
  if (!sh && create) {
    sh = ss().insertSheet(month);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    // 날짜 컬럼(B)을 텍스트로 고정 → '2026-07-01'이 날짜로 자동변환되지 않게
    sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  return sh;
}

function metaSheet() {
  var sh = ss().getSheetByName(META);
  if (!sh) {
    sh = ss().insertSheet(META);
    sh.appendRow(['month', 'budget', 'catBudgets']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function toDateStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

// _meta의 month 셀이 날짜로 자동변환됐어도 'YYYY-MM' 문자열로 되돌림
function toMonthKey(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ss().getSpreadsheetTimeZone(), 'yyyy-MM');
  return String(v);
}

/* ---------- 읽기 ---------- */

function readAll() {
  var months = {};

  // _meta (예산 정보)
  var mv = metaSheet().getDataRange().getValues();
  for (var i = 1; i < mv.length; i++) {
    var mo = toMonthKey(mv[i][0]); if (!mo) continue;
    if (!/^\d{4}-\d{2}$/.test(mo)) continue; // __lists__ 등 특수 행은 건너뜀
    months[mo] = months[mo] || { budget: 0, catBudgets: {}, entries: [] };
    // 중복 행이 남아있어도 0이 아닌 예산/비어있지 않은 항목예산이 이기도록 병합
    var b = Number(mv[i][1]) || 0;
    if (b) months[mo].budget = b;
    try { var cb = mv[i][2] ? JSON.parse(mv[i][2]) : {}; if (cb && Object.keys(cb).length) months[mo].catBudgets = cb; }
    catch (e) {}
  }

  // 달별 지출 시트 (YYYY-MM 이름만)
  var sheets = ss().getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var name = sheets[s].getName();
    if (name === META) continue;
    if (!/^\d{4}-\d{2}$/.test(name)) continue;
    months[name] = months[name] || { budget: 0, catBudgets: {}, entries: [] };
    var vals = sheets[s].getDataRange().getValues();
    for (var r = 1; r < vals.length; r++) {
      var row = vals[r];
      if (!row[0]) continue;
      months[name].entries.push({
        id: String(row[0]),
        date: toDateStr(row[1]),
        amount: Number(row[2]) || 0,
        category: String(row[3]),
        user: String(row[4]),
        card: String(row[5]),
        memo: String(row[6] || '')
      });
    }
  }
  return months;
}

/* ---------- 쓰기 ---------- */

function addEntry(month, entry) {
  var sh = monthSheet(month, true);
  var id = entry.id || (Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36));
  sh.appendRow([
    id,
    String(entry.date || ''),
    Number(entry.amount) || 0,
    String(entry.category || ''),
    String(entry.user || ''),
    String(entry.card || ''),
    String(entry.memo || '')
  ]);
}

function deleteEntry(month, id) {
  var sh = monthSheet(month, false);
  if (!sh) return;
  var vals = sh.getDataRange().getValues();
  for (var r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][0]) === String(id)) sh.deleteRow(r + 1);
  }
}

/* ---------- 목록(항목/사용자/카드) ---------- */

// _meta 시트의 __lists__ 행(catBudgets 열)에 JSON으로 저장된 목록을 읽음
function readConfig() {
  var sh = metaSheet();
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === LISTS_KEY) {
      try { return JSON.parse(vals[i][2] || 'null'); } catch (e) { return null; }
    }
  }
  return null;
}

function setLists(lists) {
  upsertMeta(LISTS_KEY, function (row) { row[2] = JSON.stringify(lists || {}); });
}

function setBudget(month, budget) {
  upsertMeta(month, function (row) { row[1] = Number(budget) || 0; });
}

function setCatBudget(month, catBudgets) {
  upsertMeta(month, function (row) { row[2] = JSON.stringify(catBudgets || {}); });
}

function upsertMeta(month, mutate) {
  var sh = metaSheet();
  var vals = sh.getDataRange().getValues();
  var firstIdx = -1;
  var merged = null;
  // 이 달에 해당하는 모든 행을 찾아 병합(0이 아닌 예산/비어있지 않은 항목예산이 이김)
  for (var i = 1; i < vals.length; i++) {
    if (toMonthKey(vals[i][0]) === month) {
      if (firstIdx < 0) { firstIdx = i; merged = [month, vals[i][1], vals[i][2]]; }
      else {
        if (Number(vals[i][1]) || 0) merged[1] = vals[i][1];
        if (vals[i][2] && vals[i][2] !== '{}') merged[2] = vals[i][2];
      }
    }
  }
  if (firstIdx >= 0) {
    mutate(merged);
    // 정규화된 month 문자열로 다시 써서 날짜로 변환된 기존 행 복구
    sh.getRange(firstIdx + 1, 1, 1, 3).setValues([[month, merged[1], merged[2]]]);
    // 중복 행 제거(아래에서 위로)
    for (var j = vals.length - 1; j > firstIdx; j--) {
      if (toMonthKey(vals[j][0]) === month) sh.deleteRow(j + 1);
    }
    return;
  }
  var nr = [month, 0, '{}'];
  mutate(nr);
  sh.appendRow(nr);
}

/* ---------- 최초 1회 초기화(선택) ---------- */

function setup() {
  metaSheet();
  monthSheet('2026-07', true);
}
