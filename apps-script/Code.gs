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
// 시트 구조 (유통 마진 분석)
const SHEETS = {
  summary: { name: '유통_월별요약', header: ['월', '고객수', '외부고객', '자점고객', '상부정산', '자점마진', '전체정산', '유통마진', '고객당마진', '1만원미만', '역마진', '저장일시', '저장자', '메모'] },
  partner: { name: '유통_협력점별', header: ['월', '협력점', '자점', '고객수', '상부정산', '자점마진', '전체정산', '유통마진', '1만원미만', '역마진'] },
  low:     { name: '유통_저마진고객', header: ['월', '고객키', '고객명', '연락처', '협력점', '자점', '상부점', '통신사', '상품', '상부정산', '자점마진', '전체정산', '유통마진', '사유', '사유작성자', '사유수정일시', '구분'] },
  category:{ name: '유통_구분별', header: ['월', '구분', '건수', '상부정산', '자점마진', '전체정산', '유통마진', '1만원미만', '역마진'] },
};
// 시트 열 이름 ↔ 앱 필드
const FIELD = {
  월: 'month', 고객수: 'customers', 외부고객: 'extCustomers', 자점고객: 'jaCustomers', 상부정산: 'upper', 자점마진: 'jaMargin',
  전체정산: 'total', 유통마진: 'margin', 고객당마진: 'perCustomer', '1만원미만': 'lowCount', 역마진: 'negCount', 저장일시: 'savedAt',
  저장자: 'savedBy', 메모: 'memo', 협력점: 'partner', 자점: 'ja', 고객키: 'key', 고객명: 'name', 연락처: 'phone', 상부점: 'upperShop',
  통신사: 'carrier', 상품: 'products', 사유: 'reason', 사유작성자: 'reasonBy', 사유수정일시: 'reasonAt', 구분: 'category', 건수: 'customers',
};
const MONEY_COLS = ['상부정산', '자점마진', '전체정산', '유통마진', '고객당마진'];

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
    if ((p.action || 'list') === 'list') return listAll_(user);
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
    if (body.action === 'list') return listAll_(user);
    if (body.action === 'changePassword') return changePassword_(user, body.current, body.next);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (body.action === 'save') return saveMonth_(user, body.month, body.summary || {}, body.partners || [], body.lows || [], body.memo || '', body.categories || []);
      if (body.action === 'delete') return deleteMonth_(body.month);
      if (body.action === 'reason') return setReason_(user, body.month, body.key, body.reason);
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

function sheet_(def) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(def.name);
  if (!sh) sh = ss.insertSheet(def.name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(def.header);
    sh.getRange(1, 1, 1, def.header.length).setFontWeight('bold').setBackground('#eef2fb');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@'); // '2026-08' 이 날짜로 바뀌지 않도록 텍스트 고정
  } else if (sh.getLastColumn() < def.header.length) {
    // 열이 추가된 경우 (이전 버전 시트) 헤더만 뒤에 덧붙임
    sh.getRange(1, 1, 1, def.header.length).setValues([def.header]).setFontWeight('bold').setBackground('#eef2fb');
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

function readSheet_(def) {
  const sh = sheet_(def);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, def.header.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => {
      const o = {};
      def.header.forEach((h, i) => {
        let v = r[i];
        if (h === '월') v = monthStr_(v);
        else if (v instanceof Date) v = v.toISOString();
        else if (h === '자점') v = v === 'Y' || v === true;
        o[FIELD[h]] = v;
      });
      return o;
    });
}

function listAll_(user) {
  return { ok: true, user: user, summaries: readSheet_(SHEETS.summary), partners: readSheet_(SHEETS.partner), lows: readSheet_(SHEETS.low), categories: readSheet_(SHEETS.category) };
}

function removeMonthRows_(sh, month) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  let removed = 0;
  for (let i = col.length - 1; i >= 0; i--) {   // 아래에서 위로, 연속 구간은 한 번에 삭제
    if (monthStr_(col[i][0]) !== month) continue;
    let start = i;
    while (start - 1 >= 0 && monthStr_(col[start - 1][0]) === month) start--;
    sh.deleteRows(start + 2, i - start + 1);
    removed += i - start + 1;
    i = start;
  }
  return removed;
}

function writeRows_(def, objs) {
  if (!objs.length) return;
  const sh = sheet_(def);
  const values = objs.map(o => def.header.map(h => {
    const v = o[FIELD[h]];
    if (h === '자점') return v ? 'Y' : '';
    return v === undefined || v === null ? '' : v;
  }));
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, values.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, values.length, def.header.length).setValues(values);
  def.header.forEach((h, i) => {
    if (MONEY_COLS.indexOf(h) >= 0) sh.getRange(start, i + 1, values.length, 1).setNumberFormat('#,##0');
  });
  const last = sh.getLastRow();
  if (last > 2) sh.getRange(2, 1, last - 1, def.header.length).sort([{ column: 1, ascending: true }]);
}

function saveMonth_(user, month, summary, partners, lows, memo, categories) {
  validMonth_(month);
  if (!partners.length) throw new Error('저장할 데이터가 없습니다.');
  // 같은 월을 다시 저장해도 이미 적어둔 사유는 고객키 기준으로 유지
  const kept = {};
  readSheet_(SHEETS.low).filter(r => r.month === month && r.reason).forEach(r => { kept[r.key] = r; });
  const now = new Date();
  [SHEETS.summary, SHEETS.partner, SHEETS.low, SHEETS.category].forEach(def => removeMonthRows_(sheet_(def), month));
  writeRows_(SHEETS.category, categories.map(c => Object.assign({}, c, { month: month })));
  writeRows_(SHEETS.summary, [Object.assign({}, summary, { month: month, savedAt: now, savedBy: user, memo: memo })]);
  writeRows_(SHEETS.partner, partners.map(p => Object.assign({}, p, { month: month })));
  writeRows_(SHEETS.low, lows.map(l => {
    const k = kept[l.key];
    return Object.assign({}, l, { month: month, reason: k ? k.reason : '', reasonBy: k ? k.reasonBy : '', reasonAt: k && k.reasonAt ? new Date(k.reasonAt) : '' });
  }));
  return { ok: true, month: month, partners: partners.length, lows: lows.length, keptReasons: Object.keys(kept).length };
}

function deleteMonth_(month) {
  validMonth_(month);
  [SHEETS.summary, SHEETS.partner, SHEETS.low, SHEETS.category].forEach(def => removeMonthRows_(sheet_(def), month));
  return { ok: true, month: month };
}

function setReason_(user, month, key, reason) {
  validMonth_(month);
  const def = SHEETS.low, sh = sheet_(def);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('해당 고객을 찾을 수 없습니다.');
  const keyCol = def.header.indexOf('고객키'), reasonCol = def.header.indexOf('사유');
  const vals = sh.getRange(2, 1, last - 1, keyCol + 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (monthStr_(vals[i][0]) === month && String(vals[i][keyCol]) === String(key)) {
      const now = new Date();
      sh.getRange(i + 2, reasonCol + 1, 1, 3).setValues([[String(reason || ''), reason ? user : '', reason ? now : '']]);
      return { ok: true, reasonBy: reason ? user : '', reasonAt: reason ? now.toISOString() : '' };
    }
  }
  throw new Error('해당 고객을 찾을 수 없습니다. 월 데이터를 다시 저장했는지 확인하세요.');
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
  Object.keys(SHEETS).forEach(k => sheet_(SHEETS[k]));
  Logger.log('시트 준비 완료');
}
