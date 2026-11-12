/**
 * 재고 약도 — 저장 서버 (구글 앱스 스크립트)
 * 이 스크립트는 구글 시트에 붙어 있어야 합니다. (시트 > 확장 프로그램 > Apps Script)
 *  - 보기: 누구나 (로그인 없음)
 *  - 저장: "수정권한" 시트에 적힌 이름+코드(비밀번호)가 맞는 사람만 · 권한 칸이 "관리자"면 백업 불러오기 가능
 */
const CHUNK = 40000;           // 시트 한 칸에 넣는 글자 수 (칸 한도 5만 자)
const HIST_KEEP = 80;          // 탭마다 남길 지난 기록 수

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
function histSheet_() { return sheet_('history', ['키', '탭', '저장시각', '조각수', '데이터']); }
function logSheet_()  { return sheet_('log', ['시각', '이름', '탭', '내용']); }
function editSheet_() { return sheet_('수정권한', ['이름', '코드', '권한']); }

/** 처음 한 번 실행: 시트를 만들고 관리자 코드를 하나 만들어 둡니다. */
function setup() {
  docsSheet_(); histSheet_(); logSheet_();
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
  pr.setProperty('mig2', '1');
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
  pr.setProperty('mig3', '1');
}
/* 이름(있으면)과 코드가 같은 줄에 있어야 통과. 돌려주는 값: {name, admin, row} */
function who_(code, name) {
  code = String(code || '').trim(); if (!code) return null;
  name = name === undefined || name === null ? '' : String(name).trim();
  const v = editSheet_().getDataRange().getDisplayValues();
  for (let i = 1; i < v.length; i++) {
    const n = String(v[i][0]).trim(), c = String(v[i][1]).trim(), role = String(v[i][2] || '').trim();
    if (n && c && c === code && (!name || name === n)) return { name: n, admin: role === '관리자' || role.toLowerCase() === 'admin', row: i + 1 };
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

/* ---------- 지난 기록 ---------- */
function readHist_(cid, limit) {
  const list = [];
  rowsOf_(histSheet_()).forEach(function (r) {
    if (String(r[1]) !== cid) return;
    const n = +r[3] || 1;
    try { list.push(JSON.parse(r.slice(4, 4 + n).join(''))); } catch (e) {}
  });
  list.sort(function (a, b) { return (b.savedAt || 0) - (a.savedAt || 0); });
  return list.slice(0, limit || 60);
}
function writeHist_(cid, id, data) {
  const sh = histSheet_(), key = cid + '|' + id, n = sh.getLastRow();
  const meta = n < 2 ? [] : sh.getRange(2, 1, n - 1, 3).getValues();
  let row = 0; for (let i = 0; i < meta.length; i++) if (String(meta[i][0]) === key) { row = i + 2; break; }
  if (!row) row = n + 1;
  const ch = chunks_(JSON.stringify(data));
  ensureCols_(sh, 4 + ch.length);
  const rg = sh.getRange(row, 1, 1, 4 + ch.length);
  rg.setNumberFormat('@');
  rg.setValues([[key, cid, String(data.savedAt || Date.now()), String(ch.length)].concat(ch)]);
  // 오래된 기록 정리
  const mine = []; meta.forEach(function (m, i) { if (String(m[1]) === cid && String(m[0]) !== key) mine.push({ row: i + 2, t: +m[2] || 0 }); });
  if (mine.length + 1 > HIST_KEEP) {
    mine.sort(function (a, b) { return a.t - b.t; });
    mine.slice(0, mine.length + 1 - HIST_KEEP).map(function (m) { return m.row; }).sort(function (a, b) { return b - a; })
      .forEach(function (r) { sh.deleteRow(r); });
  }
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
    switch (p.action) {
      case 'ver':  return out_({ ok: true, vers: versions_() });
      case 'all':  return out_({ ok: true, docs: readDocs_() });
      case 'docs': return out_({ ok: true, docs: readDocs_(String(p.cids || '').split(',')) });
      case 'hist': return out_({ ok: true, items: readHist_(String(p.cid || ''), +p.limit || 60) });
      default:     return out_({ ok: true, app: '재고 약도 저장 서버' });
    }
  } catch (err) { return out_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad' }); }
  try { migrate_(); migrate3_(); } catch (err) {}
  if (b.action === 'auth') {
    if (!String(b.name || '').trim()) return out_({ ok: false, error: 'name' });   // 이름과 코드를 둘 다 넣어야 로그인
    const a = who_(b.code, b.name);
    return out_({ ok: !!a, name: a ? a.name : '', admin: a ? a.admin : false });
  }
  const me = who_(b.code, b.name);
  if (!me) return out_({ ok: false, error: 'auth' });
  const name = me.name;
  if (b.action === 'addUser') {         // 회원 등록: 관리자만
    if (!me.admin) return out_({ ok: false, error: 'admin' });
    const nn = String(b.newName || '').trim(), nc = String(b.newCode || '').trim();
    if (!nn || nn.length > 20) return out_({ ok: false, error: 'name' });
    if (nc.length < 4 || nc.length > 20) return out_({ ok: false, error: 'len' });
    const v = editSheet_().getDataRange().getDisplayValues();
    for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === nn) return out_({ ok: false, error: 'exists' });
    const sh = editSheet_(), r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, 3).setNumberFormat('@').setValues([[nn, nc, '']]);
    logSheet_().appendRow([new Date(), name, '-', '회원 등록: ' + nn]);
    return out_({ ok: true });
  }
  if (b.action === 'passwd') {          // 비밀번호(수정 코드) 바꾸기: 기존 코드가 맞아야 함
    const nc = String(b.newCode || '').trim();
    if (nc.length < 4 || nc.length > 20) return out_({ ok: false, error: 'len' });
    // 다른 사람과 같은 비밀번호여도 됨 (로그인할 때 이름으로 구분)
    editSheet_().getRange(me.row, 2).setNumberFormat('@').setValue(nc);
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
