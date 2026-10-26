/**
 * 자취 가계부 — 구글 시트 저장용 Apps Script
 *
 * 쓰는 법
 * 1) 새 구글 시트를 만들고 [확장 프로그램] → [Apps Script]를 연다.
 * 2) 기본으로 있는 코드를 지우고 이 파일 내용을 통째로 붙여넣는다.
 * 3) [배포] → [새 배포] → 유형: 웹 앱 → 실행: 나 → 액세스 권한: 모든 사용자 → [배포]
 * 4) 권한 승인 후 나오는 웹 앱 URL(…/exec)을 가계부 사이트 첫 화면에 붙여넣는다.
 *
 * 코드를 고친 뒤에는 [배포] → [배포 관리] → 연필 아이콘 → 버전: 새 버전 → [배포]
 * 로 올려야 URL이 바뀌지 않는다.
 */

const SHEETS = {
  expenses:  { name: '지출',   cols: ['id', 'date', 'cat', 'memo', 'amount', 'pay', 'fixed', 'createdAt'] },
  utilities: { name: '공과금', cols: ['id', 'month', 'elec', 'water', 'gas', 'kwh', 'updatedAt'] },
  invest:    { name: '투자',   cols: ['id', 'date', 'kind', 'name', 'amount', 'createdAt'] },
};
const TEXT_COLS = ['id', 'date', 'month', 'cat', 'memo', 'pay', 'kind', 'name'];
const NUM_COLS = ['amount', 'elec', 'water', 'gas', 'kwh', 'createdAt', 'updatedAt'];
const SETTINGS_SHEET = '설정';

// 처음 한 번만 쓰이는 기본 예산 (2026-10-26 대화 기준). 이후로는 사이트의 '예산 설정' 탭에서 고친다.
const DEFAULT_SETTINGS = {
  salary: 3176000, loanPrincipal: 109000000, loanRatePct: 2.6, saving: 500000, tradingDays: 21,
  items: [
    { cat: '주거', name: '전세대출 이자', auto: 'loanInterest', amount: 0, day: '' },
    { cat: '주거', name: '관리비', amount: 30000, day: '' },
    { cat: '공과금', name: '전기·수도·가스 (가정치)', amount: 100000, day: '' },
    { cat: '고정', name: '애플 iCloud', amount: 3300, day: '13' },
    { cat: '고정', name: '공과대학 회비', amount: 20000, day: '11' },
    { cat: '고정', name: '미용실', amount: 25000, day: '' },
    { cat: '고정', name: '교통카드', amount: 55000, day: '' },
    { cat: '고정', name: '삼성 보험', amount: 8000, day: '5' },
    { cat: '고정', name: '핸드폰 요금', amount: 10000, day: '21' },
    { cat: '식비', name: '밀프랩 (하루 2끼)', amount: 200000, day: '' },
    { cat: '생활용품', name: '휴지·생필품', amount: 50000, day: '' },
    { cat: '데이트·문화', name: '여자친구 문화생활', amount: 100000, day: '' },
    { cat: '경조사', name: '경조사비 적립', amount: 100000, day: '' },
    { cat: '의류·비상', name: '의류·비상금 적립', amount: 100000, day: '' },
  ],
};

/* ---------- 시트 도우미 ---------- */
function sheet_(key) {
  const ss = SpreadsheetApp.getActive();
  const def = SHEETS[key];
  if (!def) throw new Error('알 수 없는 시트: ' + key);
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    def.cols.forEach((c, i) => {
      if (TEXT_COLS.indexOf(c) >= 0) sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  }
  return sh;
}

function cell_(col, v) {
  if (v instanceof Date) {
    const tz = Session.getScriptTimeZone();
    return Utilities.formatDate(v, tz, col === 'month' ? 'yyyy-MM' : 'yyyy-MM-dd');
  }
  if (NUM_COLS.indexOf(col) >= 0) return Number(v) || 0;
  if (col === 'fixed') return v === true || v === 'TRUE' || v === 'true';
  return v === null || v === undefined ? '' : String(v);
}

function readAll_(key) {
  const sh = sheet_(key);
  const values = sh.getDataRange().getValues();
  const head = values.shift() || [];
  return values
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => {
      const o = {};
      head.forEach((h, i) => { o[h] = cell_(h, r[i]); });
      return o;
    });
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function rowOf_(key, data) {
  return SHEETS[key].cols.map(c => (data[c] === undefined ? '' : data[c]));
}

function readSettings_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET);
    sh.getRange('A1').setNumberFormat('@').setValue(JSON.stringify(DEFAULT_SETTINGS));
    sh.getRange('A2').setValue('↑ 사이트의 예산 설정 값이 저장되는 칸이에요. 직접 고치지 마세요.');
  }
  const raw = sh.getRange('A1').getValue();
  try { return raw ? JSON.parse(raw) : DEFAULT_SETTINGS; } catch (e) { return DEFAULT_SETTINGS; }
}

function writeSettings_(data) {
  readSettings_();
  SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET).getRange('A1').setValue(JSON.stringify(data));
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 웹 앱 진입점 ---------- */
function doGet(e) {
  try {
    return json_({
      ok: true,
      settings: readSettings_(),
      expenses: readAll_('expenses'),
      utilities: readAll_('utilities'),
      invest: readAll_('invest'),
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'settings') {
      writeSettings_(b.data);
    } else if (b.action === 'add') {
      sheet_(b.col).appendRow(rowOf_(b.col, b.data));
    } else if (b.action === 'upsert') {
      const sh = sheet_(b.col);
      const row = findRow_(sh, b.data.id);
      const vals = [rowOf_(b.col, b.data)];
      if (row > 0) sh.getRange(row, 1, 1, vals[0].length).setValues(vals);
      else sh.appendRow(vals[0]);
    } else if (b.action === 'delete') {
      const sh = sheet_(b.col);
      const row = findRow_(sh, b.id);
      if (row > 0) sh.deleteRow(row);
    } else {
      throw new Error('알 수 없는 요청: ' + b.action);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
