/**
 * 유통사업부 도매정산 — Google Sheets 저장/조회 API (Google Apps Script 웹앱)
 *
 * 설치: README.md 의 "구글 시트 연동 설정" 참고
 *  1) 구글 시트 → 확장 프로그램 → Apps Script 에 이 파일 내용을 붙여넣기
 *  2) (선택) 프로젝트 설정 → 스크립트 속성에 API_KEY 추가
 *  3) 배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자) → URL 을 앱 설정에 입력
 */

const SPREADSHEET_ID = '1rFwSqrdn_QTIm_Nz6g386mrRREa-q4jMHnGlQ1uP03o';
const DATA_SHEET = '정산데이터';
const SUMMARY_SHEET = '월별요약';
const DATA_HEADER = ['월', '상품', '건수', '상부정산', '전체정산', '마진', '저장일시', '메모'];
const SUMMARY_HEADER = ['월', '건수', '상부정산', '전체정산', '마진', '건당마진', '마진율(%)', '저장일시'];

// ── 엔트리 포인트 ─────────────────────────────────────────────
function doGet(e) {
  return handle_(() => {
    const p = (e && e.parameter) || {};
    checkKey_(p.key);
    const action = p.action || 'list';
    if (action === 'ping') return { ok: true, message: 'pong' };
    if (action === 'list') return { ok: true, rows: readAll_() };
    throw new Error('알 수 없는 action: ' + action);
  });
}

function doPost(e) {
  return handle_(() => {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkKey_(body.key);
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

function checkKey_(key) {
  const required = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (required && key !== required) throw new Error('인증 실패: API 키가 올바르지 않습니다.');
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

// 편집기에서 한 번 실행하면 권한 승인 + 시트 생성
function setup() {
  sheet_(DATA_SHEET, DATA_HEADER);
  sheet_(SUMMARY_SHEET, SUMMARY_HEADER);
  Logger.log('시트 준비 완료');
}
