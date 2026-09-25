/**
 * 유통사업부 도매정산 — Google Sheets 저장/조회 API (Google Apps Script 웹앱)
 *
 * 설치: README.md 의 "구글 시트 연동 설정" 참고
 *  1) 구글 시트 → 확장 프로그램 → Apps Script 에 이 파일 내용을 붙여넣기
 *  2) setup 함수 1회 실행 (권한 승인)
 *  3) 배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자) → URL 을 앱에 입력
 *     코드 수정 후에는 배포 관리 → 수정 → 버전: 새 버전 으로 재배포 (URL 유지)
 *
 * 로그인: 모든 요청은 로그인 후 받은 토큰이 있어야 처리됩니다.
 *  - 계정은 스크립트 속성 USERS 에 {아이디: {salt, hash}} 형태(SHA-256)로 저장됩니다.
 *  - USERS 가 비어 있으면 초기 계정(admin)을 사용합니다. 첫 로그인 후 비밀번호를 꼭 변경하세요.
 */

const SPREADSHEET_ID = '1rFwSqrdn_QTIm_Nz6g386mrRREa-q4jMHnGlQ1uP03o';
const DATA_SHEET = '정산데이터';
const SUMMARY_SHEET = '월별요약';
const DATA_HEADER = ['월', '상품', '건수', '상부정산', '전체정산', '마진', '저장일시', '메모'];
const SUMMARY_HEADER = ['월', '건수', '상부정산', '전체정산', '마진', '건당마진', '마진율(%)', '저장일시'];

// 초기 계정 (비밀번호 원문은 저장하지 않고 salt + SHA-256 해시만 보관)
const DEFAULT_USERS = { admin: { salt: 'd4cbc8afee22be88', hash: 'ee7a69d5f5203cc5fdd7a26a1025969d387eb31004743ee681e18cd6e1a1471e' } };
const TOKEN_TTL_SEC = 6 * 60 * 60;   // 로그인 유지 6시간
const MAX_FAILS = 5;                 // 10분 내 5회 실패 시 잠금
const LOCK_SEC = 10 * 60;

// ── 엔트리 포인트 ─────────────────────────────────────────────
function doGet(e) {
  return handle_(() => {
    const p = (e && e.parameter) || {};
    if (p.action === 'ping') return { ok: true, message: 'pong' };
    const user = auth_(p.token);
    if ((p.action || 'list') === 'list') return { ok: true, user: user, rows: readAll_() };
    throw new Error('알 수 없는 action: ' + p.action);
  });
}

function doPost(e) {
  return handle_(() => {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'ping') return { ok: true, message: 'pong' };
    if (body.action === 'login') return login_(body.id, body.password);
    const user = auth_(body.token);
    if (body.action === 'logout') { CacheService.getScriptCache().remove('tok_' + body.token); return { ok: true }; }
    if (body.action === 'list') return { ok: true, user: user, rows: readAll_() };
    if (body.action === 'changePassword') return changePassword_(user, body.current, body.next);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (body.action === 'save') return saveMonth_(body.month, body.rows || [], body.memo || '');
      if (body.action === 'delete') return deleteMonth_(body.month);
      throw new Error('알 수 없는 action: ' + body.action);
    } finally {
      lock.releaseLock();
    }
  });
}

// ── 인증 ─────────────────────────────────────────────────────
function users_() {
  const raw = PropertiesService.getScriptProperties().getProperty('USERS');
  return raw ? JSON.parse(raw) : DEFAULT_USERS;
}

function hash_(salt, password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function login_(id, password) {
  id = String(id || '').trim();
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + id;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS) throw new Error('로그인 시도가 너무 많습니다. 10분 후 다시 시도하세요.');
  const u = users_()[id];
  if (!u || hash_(u.salt, String(password || '')) !== u.hash) {
    cache.put(failKey, String(fails + 1), LOCK_SEC);
    throw new Error('아이디 또는 비밀번호가 올바르지 않습니다.');
  }
  cache.remove(failKey);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('tok_' + token, id, TOKEN_TTL_SEC);
  return { ok: true, token: token, user: id };
}

function auth_(token) {
  const id = token && CacheService.getScriptCache().get('tok_' + token);
  if (!id) throw new Error('AUTH: 로그인이 필요합니다.');
  return id;
}

function changePassword_(id, current, next) {
  const all = users_();
  const u = all[id];
  if (!u || hash_(u.salt, String(current || '')) !== u.hash) throw new Error('현재 비밀번호가 올바르지 않습니다.');
  next = String(next || '');
  if (next.length < 6) throw new Error('새 비밀번호는 6자 이상이어야 합니다.');
  const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  all[id] = { salt: salt, hash: hash_(salt, next) };
  PropertiesService.getScriptProperties().setProperty('USERS', JSON.stringify(all));
  return { ok: true };
}

// ── 구현 ─────────────────────────────────────────────────────
function handle_(fn) {
  let out;
  try {
    out = fn();
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, header) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#eef2fb');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@'); // '2026-08' 이 날짜로 바뀌지 않도록 텍스트 고정
  }
  return sh;
}

function monthStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  return String(v || '').trim();
}

function validMonth_(m) {
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(m || ''))) throw new Error('월 형식 오류(YYYY-MM): ' + m);
}

function readAll_() {
  const sh = sheet_(DATA_SHEET, DATA_HEADER);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, DATA_HEADER.length).getValues()
    .filter(r => r[0] !== '' && r[1] !== '')
    .map(r => ({
      month: monthStr_(r[0]),
      name: String(r[1]),
      건수: Number(r[2]) || 0,
      상부: Number(r[3]) || 0,
      전체: Number(r[4]) || 0,
      마진: Number(r[5]) || 0,
      savedAt: r[6] instanceof Date ? r[6].toISOString() : String(r[6] || ''),
      memo: String(r[7] || ''),
    }));
}

function removeMonthRows_(sh, month) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  let removed = 0;
  // 아래에서 위로 삭제 (연속 구간은 한 번에)
  for (let i = col.length - 1; i >= 0; i--) {
    if (monthStr_(col[i][0]) !== month) continue;
    let start = i;
    while (start - 1 >= 0 && monthStr_(col[start - 1][0]) === month) start--;
    sh.deleteRows(start + 2, i - start + 1);
    removed += i - start + 1;
    i = start;
  }
  return removed;
}

function saveMonth_(month, rows, memo) {
  validMonth_(month);
  if (!rows.length) throw new Error('저장할 데이터가 없습니다.');
  const sh = sheet_(DATA_SHEET, DATA_HEADER);
  const replaced = removeMonthRows_(sh, month);
  const now = new Date();
  const values = rows.map(r => [
    month, String(r.name), Number(r.건수) || 0, Math.round(Number(r.상부) || 0),
    Math.round(Number(r.전체) || 0), Math.round(Number(r.마진) || 0), now, memo,
  ]);
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, values.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, values.length, DATA_HEADER.length).setValues(values);
  sh.getRange(start, 3, values.length, 4).setNumberFormat('#,##0');
  sortByMonth_(sh, DATA_HEADER.length);
  rebuildSummary_();
  return { ok: true, month: month, saved: values.length, replaced: replaced };
}

function deleteMonth_(month) {
  validMonth_(month);
  const removed = removeMonthRows_(sheet_(DATA_SHEET, DATA_HEADER), month);
  rebuildSummary_();
  return { ok: true, month: month, removed: removed };
}

function sortByMonth_(sh, width) {
  const last = sh.getLastRow();
  if (last > 2) sh.getRange(2, 1, last - 1, width).sort([{ column: 1, ascending: true }]);
}

// 사람이 시트에서 바로 볼 수 있는 월별 요약 (저장/삭제 때마다 재생성)
function rebuildSummary_() {
  const rows = readAll_();
  const by = {};
  rows.forEach(r => {
    const m = by[r.month] || (by[r.month] = { 건수: 0, 상부: 0, 전체: 0, 마진: 0, savedAt: r.savedAt });
    m.건수 += r.건수; m.상부 += r.상부; m.전체 += r.전체; m.마진 += r.마진;
  });
  const sh = sheet_(SUMMARY_SHEET, SUMMARY_HEADER);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, SUMMARY_HEADER.length).clearContent();
  const months = Object.keys(by).sort();
  if (!months.length) return;
  const values = months.map(m => {
    const t = by[m];
    return [m, t.건수, t.상부, t.전체, t.마진,
      t.건수 ? Math.round(t.마진 / t.건수) : 0,
      t.상부 ? Math.round(t.마진 / t.상부 * 1000) / 10 : 0,
      t.savedAt ? new Date(t.savedAt) : ''];
  });
  sh.getRange(2, 1, values.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, values.length, SUMMARY_HEADER.length).setValues(values);
  sh.getRange(2, 2, values.length, 5).setNumberFormat('#,##0');
}

// 계정 추가/비밀번호 초기화: 편집기에서 아이디·비밀번호를 바꿔 실행
function addUser() {
  const id = 'newuser', password = 'change-me-123';
  const all = users_();
  const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  all[id] = { salt: salt, hash: hash_(salt, password) };
  PropertiesService.getScriptProperties().setProperty('USERS', JSON.stringify(all));
  Logger.log('계정 저장: ' + id + ' (전체: ' + Object.keys(all).join(', ') + ')');
}

// 편집기에서 한 번 실행하면 권한 승인 + 시트 생성
function setup() {
  sheet_(DATA_SHEET, DATA_HEADER);
  sheet_(SUMMARY_SHEET, SUMMARY_HEADER);
  Logger.log('시트 준비 완료');
}
