/**
 * 촬영팀 공동 스케줄 — 구글 시트 연결 코드
 * 붙여넣는 곳: 새 구글 시트 > 확장 프로그램 > Apps Script
 * (조성민 대표님 일정 시트와 다른, 새 시트에 붙여넣어야 데이터가 섞이지 않습니다)
 * 탭(설정 · 멤버 · 가능일정)이 없으면 처음 열릴 때 자동으로 만들어집니다.
 */

const TZ = 'Asia/Seoul';
const STATUS = ['가능', '불가', '미정'];

// 시트가 비어 있을 때 자동으로 채울 기본값
const DEFAULTS = {
  config: [['시작 월', '2026-10'], ['개월 수', 3], ['페이지 제목', '촬영팀 공동 스케줄']],
  members: [['구주인', '촬영팀'], ['김남영', '촬영팀'], ['김효년', '촬영팀'], ['신소윤', '촬영팀']]
};

function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

function toDateStr_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone() || TZ, 'yyyy-MM-dd');
  return String(v || '').trim();
}

function toMonthStr_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM');
  return String(v || '').trim().slice(0, 7);
}

// 필요한 탭이 없으면 만들고 기본값을 채움
function setup_() {
  const ss = SpreadsheetApp.getActive();
  const make = (name, header, rows) => {
    let sh = ss.getSheetByName(name);
    if (sh) return;
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    if (rows && rows.length) {
      sh.getRange(2, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
    }
    sh.setFrozenRows(1);
  };
  make('설정', ['항목', '값'], DEFAULTS.config);
  make('멤버', ['이름', '구분'], DEFAULTS.members);
  make('가능일정', ['날짜', '이름', '상태', '수정 시각']);
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('탭을 찾을 수 없습니다: ' + name);
  return sh;
}

function rows_(name) {
  const values = sheet_(name).getDataRange().getValues();
  values.shift(); // 헤더 제외
  return values.filter(r => String(r[0]).trim() !== '');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 홈페이지가 페이지를 열 때 전체 데이터를 가져감
function doGet() {
  setup_();
  const cfg = {};
  rows_('설정').forEach(r => { cfg[String(r[0]).trim()] = r[1]; });

  return json_({
    ok: true,
    config: {
      startMonth: toMonthStr_(cfg['시작 월']) || Utilities.formatDate(new Date(), TZ, 'yyyy-MM'),
      months: Number(cfg['개월 수']) || 3,
      title: String(cfg['페이지 제목'] || '촬영팀 공동 스케줄')
    },
    members: rows_('멤버').map(r => ({ name: String(r[0]).trim(), group: String(r[1]).trim() })),
    availability: rows_('가능일정').map(r => ({
      date: toDateStr_(r[0]), name: String(r[1]).trim(), status: String(r[2]).trim()
    }))
  });
}

// 여러 날짜를 한 번에 저장 (items: [{date, status}], status가 빈 값이면 기록 삭제)
function saveBatch_(name, items) {
  const members = rows_('멤버').map(r => String(r[0]).trim());
  if (members.indexOf(name) === -1) return { ok: false, error: '멤버에 없는 이름' };
  const want = {};
  for (const it of items) {
    const d = String(it.date || '').trim(), st = String(it.status || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return { ok: false, error: '날짜 형식 오류' };
    if (st && STATUS.indexOf(st) === -1) return { ok: false, error: '상태 값 오류' };
    want[d] = st;
  }
  const sh = sheet_('가능일정');
  const last = sh.getLastRow();
  const old = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  const out = [];
  old.forEach(r => {
    const d = toDateStr_(r[0]), n = String(r[1]).trim();
    if (!d) return;
    if (n === name && Object.prototype.hasOwnProperty.call(want, d)) {
      if (want[d]) out.push([d, n, want[d], now]);
      delete want[d];
    } else {
      out.push([d, n, String(r[2]).trim(), r[3]]);
    }
  });
  Object.keys(want).forEach(d => { if (want[d]) out.push([d, name, want[d], now]); });
  if (old.length) sh.getRange(2, 1, old.length, 4).clearContent();
  if (out.length) {
    sh.getRange(2, 1, out.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, out.length, 4).setValues(out);
  }
  return { ok: true, saved: items.length };
}

// 홈페이지에서 날짜 상태를 바꿀 때 호출됨 (빈 상태 = 선택 취소)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    setup_();
    const body = JSON.parse(e.postData.contents);
    if (Array.isArray(body.items)) return json_(saveBatch_(String(body.name || '').trim(), body.items));
    const date = String(body.date || '').trim();
    const name = String(body.name || '').trim();
    const status = String(body.status || '').trim();

    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({ ok: false, error: '날짜 형식 오류' });
    const members = rows_('멤버').map(r => String(r[0]).trim());
    if (members.indexOf(name) === -1) return json_({ ok: false, error: '멤버에 없는 이름' });
    if (status && STATUS.indexOf(status) === -1) return json_({ ok: false, error: '상태 값 오류' });

    const sh = sheet_('가능일정');
    const values = sh.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (toDateStr_(values[i][0]) === date && String(values[i][1]).trim() === name) { rowIndex = i + 1; break; }
    }
    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');

    if (!status) {
      if (rowIndex > 0) sh.deleteRow(rowIndex);
    } else if (rowIndex > 0) {
      sh.getRange(rowIndex, 3, 1, 2).setValues([[status, now]]);
    } else {
      sh.appendRow(["'" + date, name, status, now]);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
