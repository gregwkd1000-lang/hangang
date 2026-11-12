/**
 * 재고 약도 — 저장 서버 (구글 앱스 스크립트)
 * 이 스크립트는 구글 시트에 붙어 있어야 합니다. (시트 > 확장 프로그램 > Apps Script)
 *  - "수정권한" 시트에 등록된 사람만 이름+비밀번호로 로그인해서 볼 수 있음
 *  - 권한(C칸): 조회 = 보기만 · 관리자(또는 수정·빈칸) = 보기+고치기 · 마스터는 안상우 한 명 (고치기+백업+회원등록)
 */
const ADMIN_NAME = '안상우';    // 마스터는 이 한 명뿐 (다른 회원은 조회/관리자만)
const CHUNK = 40000;           // 시트 한 칸에 넣는 글자 수 (칸 한도 5만 자)

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, header) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    if (header) sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  return sh;
}
function docsSheet_() { return sheet_('docs', ['탭', '버전', '저장시각', '저장한 사람', '조각수', '데이터']); }
function logSheet_()  { return sheet_('log', ['시각', '이름', '탭', '내용']); }
function editSheet_() { return sheet_('수정권한', ['이름', '코드', '권한']); }

/** 처음 한 번 실행: 시트를 만들고 관리자 코드를 하나 만들어 둡니다. */
function setup() {
  docsSheet_(); logSheet_();
  const ed = editSheet_();
  if (ed.getLastRow() < 2) {
    const code = String(Math.floor(100000 + Math.random() * 900000));
    ed.getRange(2, 1, 1, 2).setValues([['관리자', code]]);
  }
  const first = ss_().getSheetByName('Sheet1') || ss_().getSheetByName('시트1');
  if (first && first.getLastRow() === 0 && ss_().getSheets().length > 1) ss_().deleteSheet(first);
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
/* 한 번만: "관리자" → "윤명희"로 이름 바꾸고, 관리자 계정(안상우) 추가, 권한 칸 만들기 */
function migrate_() {
  const pr = PropertiesService.getScriptProperties();
  if (pr.getProperty('mig2') === '1') return;
  const sh = editSheet_();
  sh.getRange(1, 1, sh.getMaxRows(), 3).setNumberFormat('@');
  sh.getRange(1, 3).setValue('권한').setFontWeight('bold');
  const v = sh.getDataRange().getDisplayValues();
  let hasAdmin = false;
  for (let i = 1; i < v.length; i++) {
    const n = String(v[i][0]).trim();
    if (n === '관리자') sh.getRange(i + 1, 1).setValue('윤명희');
    if (n === '안상우') { hasAdmin = true; sh.getRange(i + 1, 3).setValue('관리자'); }
  }
  if (!hasAdmin) { const r = sh.getLastRow() + 1; sh.getRange(r, 1, 1, 3).setNumberFormat('@').setValues([['안상우', '031700', '관리자']]); }
  pr.setProperty('mig2', '1'); edDirty_();
}
/* 한 번만: 강미소 계정 추가(031700), 윤명희 비밀번호 123456으로 */
function migrate3_() {
  const pr = PropertiesService.getScriptProperties();
  if (pr.getProperty('mig3') === '1') return;
  const sh = editSheet_(), v = sh.getDataRange().getDisplayValues();
  let hasKang = false;
  for (let i = 1; i < v.length; i++) {
    const n = String(v[i][0]).trim();
    if (n === '윤명희' || n === '관리자') sh.getRange(i + 1, 1, 1, 2).setNumberFormat('@').setValues([['윤명희', '123456']]);
    if (n === '강미소') { hasKang = true; sh.getRange(i + 1, 2).setNumberFormat('@').setValue('031700'); }
  }
  if (!hasKang) { const r = sh.getLastRow() + 1; sh.getRange(r, 1, 1, 3).setNumberFormat('@').setValues([['강미소', '031700', '']]); }
  pr.setProperty('mig3', '1'); edDirty_();
}
/* 한 번만: 권한 이름 바꾸기 — 안상우 = 마스터, 수정 = 관리자 */
function migrate4_() {
  const pr = PropertiesService.getScriptProperties();
  if (pr.getProperty('mig4') === '1') return;
  const sh = editSheet_(), v = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < v.length; i++) {
    const n = String(v[i][0]).trim(), r = String(v[i][2]).trim(); if (!n) continue;
    if (n === ADMIN_NAME) sh.getRange(i + 1, 3).setNumberFormat('@').setValue('마스터');
    else if (r === '' || r === '수정') sh.getRange(i + 1, 3).setNumberFormat('@').setValue('관리자');
  }
  pr.setProperty('mig4', '1'); edDirty_();
}
/* 이름(있으면)과 코드가 같은 줄에 있어야 통과. 돌려주는 값: {name, admin, row} */
function edRows_() {   // 회원 목록 (30초 캐시: 12초마다 오는 조회 요청 때문에)
  const ck = CacheService.getScriptCache(), hit = ck.get('ed');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const v = editSheet_().getDataRange().getDisplayValues();
  try { ck.put('ed', JSON.stringify(v), 30); } catch (e) {}
  return v;
}
function edDirty_() { try { CacheService.getScriptCache().remove('ed'); } catch (e) {} }
function roleOf_(r) { r = String(r || '').trim(); if (r === '마스터' || r === '관리자' || r.toLowerCase() === 'admin') return 'admin'; if (r === '조회' || r.indexOf('조회') === 0) return 'view'; return 'edit'; }
function who_(code, name) {
  code = String(code || '').trim(); if (!code) return null;
  name = name === undefined || name === null ? '' : String(name).trim();
  const v = edRows_();
  for (let i = 1; i < v.length; i++) {
    const n = String(v[i][0]).trim(), c = String(v[i][1]).trim();
    let role = roleOf_(v[i][2]); if (role === 'admin' && n !== ADMIN_NAME) role = 'edit';
    if (n && c && c === code && (!name || name === n)) return { name: n, admin: role === 'admin', role: role, row: i + 1 };
  }
  return null;
}
function ensureCols_(sh, n) { const need = n - sh.getMaxColumns(); if (need > 0) sh.insertColumnsAfter(sh.getMaxColumns(), need); }
function chunks_(json) { const a = []; for (let i = 0; i < json.length; i += CHUNK) a.push(json.slice(i, i + CHUNK)); return a.length ? a : ['']; }
function rowsOf_(sh) { const n = sh.getLastRow(); return n < 2 ? [] : sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues(); }

/* ---------- 문서(탭) ---------- */
function readDocs_(only) {
  const res = {};
  rowsOf_(docsSheet_()).forEach(function (r) {
    const k = String(r[0]); if (!k || (only && only.indexOf(k) < 0)) return;
    const n = +r[4] || 1;
    try { res[k] = { v: +r[1] || 0, at: String(r[2]), by: String(r[3]), data: JSON.parse(r.slice(5, 5 + n).join('')) }; } catch (e) {}
  });
  return res;
}
function versions_() {
  const sh = docsSheet_(), n = sh.getLastRow(), res = {};
  if (n < 2) return res;
  sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (r) { if (r[0]) res[String(r[0])] = +r[1] || 0; });
  return res;
}
function writeDoc_(cid, v, name, data) {
  const sh = docsSheet_(), keys = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  let row = 0; for (let i = 0; i < keys.length; i++) if (String(keys[i][0]) === cid) { row = i + 2; break; }
  if (!row) row = sh.getLastRow() + 1;
  const ch = chunks_(JSON.stringify(data));
  ensureCols_(sh, 5 + ch.length);
  const rg = sh.getRange(row, 1, 1, 5 + ch.length);
  rg.setNumberFormat('@');
  rg.setValues([[cid, String(v), new Date().toISOString(), name, String(ch.length)].concat(ch)]);
}

/* ---------- 지난 기록 ----------
 * 연도별 탭("기록_2026", "기록_2027" …)에 한 회차 = 한 줄로 저장.
 * 앞 6칸(키·탭·날짜·회차·저장시각·조각수)이 이름표라서, 열 때는 이름표만 훑고 필요한 줄만 꺼냄 → 몇 년이 쌓여도 빠름.
 * 10년이 지난 연도 탭은 통째로 자동 삭제.
 */
const HMETA = 6, HIST_YEARS = 10;
function yearOf_(data) { const d = String((data && data.date) || ''); return /^\d{4}-/.test(d) ? d.slice(0, 4) : String(new Date().getFullYear()); }
function histYear_(y) { return sheet_('기록_' + y, ['키', '탭', '날짜', '회차', '저장시각', '조각수', '데이터']); }
function histYears_() {   // 있는 연도 탭, 최근 연도부터
  return ss_().getSheets().map(function (sh) { const m = /^기록_(\d{4})$/.exec(sh.getName()); return m ? { y: +m[1], sh: sh } : null; })
    .filter(Boolean).sort(function (a, b) { return b.y - a.y; });
}
function histMeta_(sh) { const n = sh.getLastRow(); return n < 2 ? [] : sh.getRange(2, 1, n - 1, HMETA).getValues(); }
/* 고른 줄들만 꺼내기: 그 줄들이 있는 구간만 한 번에 읽음 */
function histRows_(sh, rows) {
  if (!rows.length) return [];
  const lo = Math.min.apply(null, rows), hi = Math.max.apply(null, rows);
  const vals = sh.getRange(lo, 1, hi - lo + 1, sh.getLastColumn()).getValues(), out = [];
  rows.forEach(function (r) { const v = vals[r - lo], n = +v[5] || 1; try { out.push(JSON.parse(v.slice(HMETA, HMETA + n).join(''))); } catch (e) {} });
  return out;
}
function readHist_(cid, limit) {   // 이 탭의 최근 기록 limit개
  limit = limit || 60; const picks = [];
  const ys = histYears_();
  for (let k = 0; k < ys.length && picks.length < limit * 3; k++) {
    histMeta_(ys[k].sh).forEach(function (m, i) { if (String(m[1]) === cid) picks.push({ sh: ys[k].sh, row: i + 2, t: +m[4] || 0 }); });
  }
  picks.sort(function (a, b) { return b.t - a.t; });
  const top = picks.slice(0, limit), by = {};
  top.forEach(function (p) { const k = p.sh.getName(); (by[k] = by[k] || { sh: p.sh, rows: [] }).rows.push(p.row); });
  let list = []; Object.keys(by).forEach(function (k) { list = list.concat(histRows_(by[k].sh, by[k].rows)); });
  list.sort(function (a, b) { return (b.savedAt || 0) - (a.savedAt || 0); });
  return list;
}
function readDay_(date) {   // 그 날짜의 모든 탭 기록 {탭: [기록…]}
  const y = String(date).slice(0, 4), sh = ss_().getSheetByName('기록_' + y), res = {};
  if (!sh) return res;
  const rows = [], cids = [];
  histMeta_(sh).forEach(function (m, i) { if (String(m[2]) === date) { rows.push(i + 2); cids.push(String(m[1])); } });
  if (!rows.length) return res;
  const lo = Math.min.apply(null, rows), hi = Math.max.apply(null, rows), vals = sh.getRange(lo, 1, hi - lo + 1, sh.getLastColumn()).getValues();
  rows.forEach(function (r, k) { const v = vals[r - lo], n = +v[5] || 1; try { (res[cids[k]] = res[cids[k]] || []).push(JSON.parse(v.slice(HMETA, HMETA + n).join(''))); } catch (e) {} });
  return res;
}
function writeHist_(cid, id, data) {
  const y = yearOf_(data), sh = histYear_(y), key = cid + '|' + id, n = sh.getLastRow();
  const keys = n < 2 ? [] : sh.getRange(2, 1, n - 1, 1).getValues();
  let row = 0; for (let i = keys.length - 1; i >= 0; i--) if (String(keys[i][0]) === key) { row = i + 2; break; }
  if (!row) row = n + 1;
  const ch = chunks_(JSON.stringify(data));
  ensureCols_(sh, HMETA + ch.length);
  const rg = sh.getRange(row, 1, 1, HMETA + ch.length);
  rg.setNumberFormat('@');
  rg.setValues([[key, cid, String(data.date || ''), String(data.round || ''), String(data.savedAt || Date.now()), String(ch.length)].concat(ch)]);
  pruneYears_();
}
function pruneYears_() {   // 10년 지난 연도 탭 삭제 (하루 한 번만 확인)
  const pr = PropertiesService.getScriptProperties(), today = new Date().toISOString().slice(0, 10);
  if (pr.getProperty('pruneDay') === today) return; pr.setProperty('pruneDay', today);
  const min = new Date().getFullYear() - HIST_YEARS;
  histYears_().forEach(function (o) { if (o.y < min && ss_().getSheets().length > 1) ss_().deleteSheet(o.sh); });
}
/* 한 번만: 예전 "history" 탭 → 연도별 탭으로 옮기기 (예전 탭은 "history_예전"으로 이름만 바꿔 보관) */
function migrateHist_() {
  const pr = PropertiesService.getScriptProperties();
  if (pr.getProperty('mig5') === '1') return;
  const lk = LockService.getScriptLock(); if (!lk.tryLock(30000)) return;
  try {
  if (pr.getProperty('mig5') === '1') return;
  const old = ss_().getSheetByName('history');
  if (old && old.getLastRow() >= 2) {
    const by = {};
    old.getRange(2, 1, old.getLastRow() - 1, old.getLastColumn()).getValues().forEach(function (r) {
      const key = String(r[0]); if (!key) return; const n = +r[3] || 1, json = r.slice(4, 4 + n).join('');
      let d; try { d = JSON.parse(json); } catch (e) { return; }
      const y = yearOf_(d); (by[y] = by[y] || []).push([key, String(r[1]), String(d.date || ''), String(d.round || ''), String(d.savedAt || r[2] || ''), String(n)].concat(r.slice(4, 4 + n).map(String)));
    });
    Object.keys(by).forEach(function (y) {
      const sh = histYear_(y), have = {}; histMeta_(sh).forEach(function (m) { have[String(m[0])] = 1; });
      const rows = by[y].filter(function (r) { return !have[r[0]]; }); if (!rows.length) return;
      const w = Math.max.apply(null, rows.map(function (r) { return r.length; }));
      rows.forEach(function (r) { while (r.length < w) r.push(''); });
      ensureCols_(sh, w);
      const rg = sh.getRange(sh.getLastRow() + 1, 1, rows.length, w); rg.setNumberFormat('@'); rg.setValues(rows);
    });
  }
  if (old) old.setName('history_예전');
  pr.setProperty('mig5', '1');
  } finally { lk.releaseLock(); }
}
function log_(name, cid, text) {
  const ck = CacheService.getScriptCache(), k = 'log|' + name + '|' + cid;
  if (ck.get(k)) return;                 // 같은 사람·같은 탭은 10분에 한 줄만
  ck.put(k, '1', 600);
  logSheet_().appendRow([new Date(), name, cid, text]);
}

/* ---------- 웹 요청 ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action) {                      // 등록된 사람만 조회 가능
      try { migrate_(); migrate3_(); migrate4_(); migrateHist_(); } catch (err) {}
      if (!String(p.name || '').trim() || !who_(p.code, p.name)) return out_({ ok: false, error: 'auth' });
    }
    const me = p.action ? who_(p.code, p.name) : null, role = me ? me.role : '';
    switch (p.action) {
      case 'ver':  return out_({ ok: true, vers: versions_(), role: role });
      case 'all':  return out_({ ok: true, docs: readDocs_(), role: role });
      case 'docs': return out_({ ok: true, docs: readDocs_(String(p.cids || '').split(',')) });
      case 'hist': return out_({ ok: true, items: readHist_(String(p.cid || ''), Math.min(+p.limit || 60, 150)) });
      case 'histDay': return out_({ ok: true, byTab: readDay_(String(p.date || '')) });
      default:     return out_({ ok: true, app: '재고 약도 저장 서버' });
    }
  } catch (err) { return out_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad' }); }
  try { migrate_(); migrate3_(); migrate4_(); migrateHist_(); } catch (err) {}
  if (b.action === 'auth') {
    const nm = String(b.name || '').trim();
    if (!nm) return out_({ ok: false, error: 'name' });   // 이름과 코드를 둘 다 넣어야 로그인
    // 5번 연속 틀리면 5분 동안 로그인 막기 (성공하면 틀린 횟수 0으로)
    const ck = CacheService.getScriptCache(), fk = 'fail_' + nm, lk = 'lock_' + nm, now = Date.now();
    const until = +ck.get(lk) || 0;
    if (until > now) return out_({ ok: false, error: 'locked', until: until });
    const a = who_(b.code, b.name);
    if (a) { ck.remove(fk); return out_({ ok: true, name: a.name, admin: a.admin, role: a.role }); }
    const n = (+ck.get(fk) || 0) + 1;
    if (n >= 5) { const u = now + 5 * 60 * 1000; ck.put(lk, String(u), 300); ck.remove(fk); return out_({ ok: false, error: 'locked', fails: 5, until: u }); }
    ck.put(fk, String(n), 21600);
    return out_({ ok: false, fails: n });
  }
  const me = who_(b.code, b.name);
  if (!me) return out_({ ok: false, error: 'auth' });
  const name = me.name;
  if ((b.action === 'set' || b.action === 'hist' || b.action === 'histBulk') && me.role === 'view') return out_({ ok: false, error: 'view' });
  if (b.action === 'addUser') {         // 회원 등록: 관리자만
    if (!me.admin) return out_({ ok: false, error: 'admin' });
    const nn = String(b.newName || '').trim(), nc = String(b.newCode || '').trim();
    const nr = b.newRole === 'view' ? '조회' : '관리자';   // 관리자 = 수정 가능 (마스터는 안상우 한 명)
    if (!nn || nn.length > 20) return out_({ ok: false, error: 'name' });
    if (nc.length < 4 || nc.length > 20) return out_({ ok: false, error: 'len' });
    const v = editSheet_().getDataRange().getDisplayValues();
    for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === nn) return out_({ ok: false, error: 'exists' });
    const sh = editSheet_(), r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, 3).setNumberFormat('@').setValues([[nn, nc, nr]]);
    edDirty_();
    logSheet_().appendRow([new Date(), name, '-', '회원 등록: ' + nn + ' (' + nr + ')']);
    return out_({ ok: true });
  }
  if (b.action === 'listUsers') {       // 회원 목록: 관리자만 (비밀번호는 보내지 않음)
    if (!me.admin) return out_({ ok: false, error: 'admin' });
    const v = editSheet_().getDataRange().getDisplayValues(), us = [];
    for (let i = 1; i < v.length; i++) { const n = String(v[i][0]).trim(); if (n && String(v[i][1]).trim()) { let r = roleOf_(v[i][2]); if (r === 'admin' && n !== ADMIN_NAME) r = 'edit'; us.push({ name: n, role: r }); } }
    return out_({ ok: true, users: us });
  }
  if (b.action === 'resetPw') {         // 관리자가 다른 사람 비밀번호 새로 정해 주기
    if (!me.admin) return out_({ ok: false, error: 'admin' });
    const tn = String(b.target || '').trim(), nc = String(b.newCode || '').trim();
    if (nc.length < 4 || nc.length > 20) return out_({ ok: false, error: 'len' });
    const sh = editSheet_(), v = sh.getDataRange().getDisplayValues();
    for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === tn) {
      sh.getRange(i + 1, 2).setNumberFormat('@').setValue(nc); edDirty_();
      logSheet_().appendRow([new Date(), name, '-', '비밀번호 재지정: ' + tn]);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'none' });
  }
  if (b.action === 'setRole' || b.action === 'delUser') {   // 권한 바꾸기 / 회원 삭제: 관리자만
    if (!me.admin) return out_({ ok: false, error: 'admin' });
    const tn = String(b.target || '').trim(); if (!tn) return out_({ ok: false, error: 'name' });
    if (tn === name) return out_({ ok: false, error: 'self' });     // 자기 자신은 못 바꿈 (관리자가 0명 되는 것 방지)
    const sh = editSheet_(), v = sh.getDataRange().getDisplayValues();
    for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === tn) {
      if (b.action === 'delUser') { sh.deleteRow(i + 1); edDirty_(); logSheet_().appendRow([new Date(), name, '-', '회원 삭제: ' + tn]); return out_({ ok: true }); }
      const nr = b.newRole === 'view' ? '조회' : '관리자';   // 관리자 = 수정 가능 (마스터는 안상우 한 명)
      sh.getRange(i + 1, 3).setNumberFormat('@').setValue(nr); edDirty_();
      logSheet_().appendRow([new Date(), name, '-', '권한 변경: ' + tn + ' → ' + nr]);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'none' });
  }
  if (b.action === 'passwd') {          // 비밀번호(수정 코드) 바꾸기: 기존 코드가 맞아야 함
    const nc = String(b.newCode || '').trim();
    if (nc.length < 4 || nc.length > 20) return out_({ ok: false, error: 'len' });
    // 다른 사람과 같은 비밀번호여도 됨 (로그인할 때 이름으로 구분)
    editSheet_().getRange(me.row, 2).setNumberFormat('@').setValue(nc); edDirty_();
    logSheet_().appendRow([new Date(), name, '-', '비밀번호 변경']);
    return out_({ ok: true });
  }
  const lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (err) { return out_({ ok: false, error: 'busy' }); }
  try {
    if (b.action === 'set') {
      const cid = String(b.cid || ''); if (!cid || !b.data) return out_({ ok: false, error: 'bad' });
      const cur = readDocs_([cid])[cid], curV = cur ? cur.v : 0;
      if ((+b.baseV || 0) !== curV) return out_({ ok: false, conflict: true, v: curV, data: cur ? cur.data : null });
      writeDoc_(cid, curV + 1, name, b.data);
      const sv = b.data.survey || {};
      log_(name, cid, (sv.date || '') + ' ' + (sv.round || '') + '차 수정');
      return out_({ ok: true, v: curV + 1, name: name });
    }
    if (b.action === 'hist') {
      const cid = String(b.cid || ''), id = String(b.id || '');
      if (!cid || !id || !b.data) return out_({ ok: false, error: 'bad' });
      writeHist_(cid, id, b.data);
      return out_({ ok: true });
    }
    if (b.action === 'histBulk') {
      if (!me.admin) return out_({ ok: false, error: 'admin' });   // 백업 불러오기는 관리자만
      (b.items || []).forEach(function (x) { if (x && x.cid && x.id && x.data) writeHist_(String(x.cid), String(x.id), x.data); });
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'bad' });
  } finally { lock.releaseLock(); }
}
