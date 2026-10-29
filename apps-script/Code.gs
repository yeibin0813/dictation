// 딕테이션 서버 (Apps Script). 구글 시트 → 확장 프로그램 → Apps Script 에 이 코드를 붙여 넣고 웹 앱으로 배포하세요.
//
// 문장·정답은 이 시트의 「문제」 탭에만 있고, 사이트(GitHub)에는 없습니다. 채점도 여기서 합니다.
// 음성은 내 드라이브의 「딕테이션 음원」 폴더에 세트마다 파일 하나(<테스트 id>_<세트 id>.json)로 있고,
// 입장 코드가 맞는 학생에게만 이 서버가 꺼내 보냅니다.
// 학생은 사이트에서 기간 → 반 → 이름 → 입장 코드로 들어옵니다.
// 관리자 페이지(admin.html)의 「설정」 탭에서 입장 코드를 정하고, 문제 파일(JSON)을 올립니다.
//
// 시트 탭
//   문제               : 책·테스트·세트·문장·빈칸·음성 파일 (관리자 페이지에서 올림, 직접 고쳐도 됨)
//   기록               : 모든 반의 풀이 기록 (한 번 풀 때마다 한 줄)
//   문장별             : 모든 반의 문장별 답 (문장마다 한 줄)
//   <기간> <반>        : 반별 풀이 기록 (예: "2026년 10월 주니어 화목 U반")
//
// 관리자 비밀번호는 ⚙ 프로젝트 설정 → 스크립트 속성에 ADMIN_KEY 라는 이름으로 넣어 주세요.

const SUMMARY = '기록';
const DETAIL = '문장별';
const PROBLEMS = '문제';
const AUDIO_FOLDER = '딕테이션 음원';
const LEGACY_UNTIL = new Date('2026-10-29T21:30:00+09:00');
const SUMMARY_HEADER = ['일시', '기간', '반', '이름', '시험', '세트', '유형', '방식', '총점(%)', '걸린 시간', '걸린 시간(초)', '문장당 듣기', '평균 들은 횟수'];
const DETAIL_HEADER = ['일시', '기간', '반', '이름', '시험', '세트', '방식', '번호', '정답 문장', '학생 답', '점수(%)', '들은 횟수'];
const PROBLEM_HEADER = ['책 id', '책', '테스트 id', '테스트', '모듈', '세트 id', '세트', '지문 이름', '유형', '번호', '문장', '빈칸 ( | 로 구분)', '음성 파일', '시작(초)', '끝(초)'];

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

const clean_ = v => String(v == null ? '' : v).trim().slice(0, 100);
const text_ = v => "'" + clean_(v); // 시트가 날짜·숫자로 바꾸지 않도록 글자로 저장
const props_ = () => PropertiesService.getScriptProperties();

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 시트 값을 JSON으로 보낼 수 있게 (날짜는 ISO 문자열로)
function rows_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
    .map(r => r.map(v => v instanceof Date ? v.toISOString() : v));
}

// 관리자 비밀번호 확인. 10분에 10번 넘게 틀리면 잠시 막습니다.
function adminOk_(key) {
  const real = props_().getProperty('ADMIN_KEY');
  if (!real) return 'nokey';
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('adminFails') || 0);
  if (fails >= 10) return 'locked';
  if (String(key || '') !== real) { cache.put('adminFails', String(fails + 1), 600); return 'wrong'; }
  return 'ok';
}

const settings_ = () => ({ code: props_().getProperty('ENTRY_CODE') || '', showAnswers: props_().getProperty('SHOW_ANSWERS') !== 'N' });
const codeOk_ = code => { const real = settings_().code; return !!real && String(code || '').trim() === real; };

function doPost(e) {
  try { return handle_(JSON.parse(e.postData.contents)); }
  catch (err) { return json_({ ok: false, error: 'server', message: String(err && err.message || err) }); }
}
function handle_(r) {
  const a = r.action;
  if (a === 'admin' || a === 'settings' || a === 'upload' || a === 'reload' || a === 'saveAudio' || a === 'audioStatus') {
    const ok = adminOk_(r.key);
    if (ok !== 'ok') return json_({ ok: false, error: ok });
    if (a === 'settings') {
      if (r.code != null) props_().setProperty('ENTRY_CODE', clean_(r.code));
      if (r.showAnswers != null) props_().setProperty('SHOW_ANSWERS', r.showAnswers ? 'Y' : 'N');
    }
    if (a === 'upload') { const n = upload_(r.data); bump_(); return json_({ ok: true, rows: n, books: catalog_() }); }
    if (a === 'saveAudio') return json_(Object.assign({ ok: true }, saveAudio_(r.name, r.data)));
    if (a === 'audioStatus') return json_({ ok: true, audio: audioStatus_() });
    if (a === 'reload') bump_();
    if (a === 'admin') return json_({ ok: true, records: rows_(SUMMARY), details: rows_(DETAIL), books: catalog_(), settings: settings_() });
    return json_({ ok: true, books: catalog_(), settings: settings_() });
  }
  // 예전 사이트(입장 코드 없음)의 기록은 새 사이트로 바꾸는 날 밤까지만 받습니다.
  if (a === 'submit' && !r.code && new Date() < LEGACY_UNTIL) return legacySubmit_(r);
  if (!codeOk_(r.code)) return json_({ ok: false, error: settings_().code ? 'code' : 'nocode' });
  if (a === 'catalog') return json_({ ok: true, books: catalog_() });
  if (a === 'set') return json_(setForStudent_(r));
  if (a === 'audio') return audio_(r);
  if (a === 'grade') return json_(gradeOne_(r));
  if (a === 'submit') return submit_(r);
  return json_({ ok: false, error: 'action' });
}

/* ---------- 문제 ---------- */
// 관리자 페이지에서 올린 책 한 권({ book:{id,title}, tests:[...] })을 「문제」 탭에 넣습니다. 같은 책의 예전 줄은 지웁니다.
function upload_(data) {
  const book = data.book, rows = [];
  if (!book || !book.id || !book.title || !Array.isArray(data.tests)) throw new Error('문제 파일 형식이 맞지 않습니다.');
  data.tests.forEach(t => t.sets.forEach(s => s.items.forEach((it, k) => rows.push([
    book.id, book.title, t.id, t.title, t.sub || '', s.id, s.title, s.name || '', s.kind || '', k + 1,
    it.t, (it.b || []).join(' | '), it.src, it.st || '', it.en || '']))));
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(PROBLEMS, PROBLEM_HEADER);
    const keep = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, PROBLEM_HEADER.length).getValues().filter(r => String(r[0]) !== book.id) : [];
    const all = keep.concat(rows);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, PROBLEM_HEADER.length).clearContent();
    if (all.length) {
      const rg = sh.getRange(2, 1, all.length, PROBLEM_HEADER.length);
      rg.setNumberFormat('@'); // 「2-01」 같은 글자가 날짜로 바뀌지 않게
      rg.setValues(all.map(r => r.map(v => v === '' ? '' : String(v))));
    }
    return rows.length;
  } finally {
    lock.releaseLock();
  }
}

// 캐시 버전을 올려 다음 요청부터 「문제」 탭을 새로 읽게 합니다.
function bump_() { props_().setProperty('PROB_VER', String(Date.now())); }

// 「문제」 탭 → 목록(정답 없음)과 세트별 문장. 6시간 캐시합니다. 시트를 직접 고쳤다면 관리자 페이지의 「시트 다시 읽기」.
function build_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROBLEMS);
  const vals = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, PROBLEM_HEADER.length).getDisplayValues() : [];
  const books = [], sets = {};
  vals.forEach(r => {
    const [bid, btitle, tid, ttitle, sub, sid, stitle, name, kind, , t, b, src, st, en] = r.map(v => String(v).trim());
    if (!bid || !tid || !sid || !t) return;
    let bk = books.find(x => x.id === bid); if (!bk) books.push(bk = { id: bid, title: btitle, tests: [] });
    let ts = bk.tests.find(x => x.id === tid); if (!ts) bk.tests.push(ts = { id: tid, title: ttitle, sub, sets: [] });
    let se = ts.sets.find(x => x.id === sid); if (!se) ts.sets.push(se = { id: sid, title: stitle, name, kind, n: 0 });
    se.n++;
    const key = tid + '/' + sid;
    (sets[key] = sets[key] || []).push({ t, b: b ? b.split('|').map(x => x.trim()).filter(Boolean) : [], src, st: +st || 0, en: +en || 0 });
  });
  books.sort((x, y) => x.id.localeCompare(y.id, 'en', { numeric: true })); // 책 순서는 시트 순서가 아니라 책 id 순 (book1, book2, …)
  return { books, sets };
}
function cached_(what) {
  const cache = CacheService.getScriptCache(), ver = props_().getProperty('PROB_VER') || '0';
  const hit = cache.get(ver + ':' + what);
  if (hit) return JSON.parse(hit);
  const { books, sets } = build_(), put = { [ver + ':cat']: JSON.stringify(books) };
  Object.keys(sets).forEach(k => put[ver + ':set:' + k] = JSON.stringify(sets[k]));
  try { cache.putAll(put, 21600); } catch (e) { /* 너무 크면 캐시 없이 */ }
  return what === 'cat' ? books : (sets[what.slice(4)] || null);
}
const catalog_ = () => cached_('cat');
const items_ = (test, set) => cached_('set:' + test + '/' + set);
function meta_(test, set) {
  for (const b of catalog_()) for (const t of b.tests) if (t.id === test) {
    const s = t.sets.find(x => x.id === set);
    if (s) return { book: b, test: t, set: s };
  }
  return null;
}

// 학생에게 보낼 세트: 음성과 (빈칸 모드면) 빈칸이 뚫린 문장. 전체 받아쓰기는 문장을 보내지 않습니다.
function setForStudent_(r) {
  const items = items_(r.test, r.set);
  if (!items) return { ok: false, error: 'noset' };
  return { ok: true, items: items.map(it => {
    const o = { src: it.src, st: it.st, en: it.en };
    if (r.mode === 'blank') o.c = pieces_(it).map(p => p.x != null ? { x: p.x } : { n: p.b.length });
    return o;
  }) };
}

/* ---------- 음성 (드라이브) ---------- */
// 세트 하나의 음성 묶음: { clips: { "<음성 파일>": base64, ... }, type: "audio/mp4" }
const packName_ = (test, set) => String(test) + '_' + String(set) + '.json';
const PACK_RE = /^[A-Za-z0-9-]+_[A-Za-z0-9-]+\.json$/;
// 「딕테이션 음원」 폴더. 내 소유 폴더만 쓰고, 한 번 찾은 폴더는 ID를 기억합니다. create면 없을 때 만듭니다.
function audioFolder_(create) {
  const id = props_().getProperty('audioFolder');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* 다시 찾음 */ } }
  const it = DriveApp.searchFolders("title = '" + AUDIO_FOLDER + "' and 'me' in owners and trashed = false");
  let f = it.hasNext() ? it.next() : null;
  if (!f && create) f = DriveApp.createFolder(AUDIO_FOLDER);
  if (!f) return null;
  props_().setProperty('audioFolder', f.getId());
  return f;
}
// 관리자 페이지에서 올린 묶음 하나를 저장합니다. 같은 이름이 있으면 휴지통으로 보내고 새로 만듭니다.
function saveAudio_(name, data) {
  name = String(name || '');
  if (!PACK_RE.test(name)) throw new Error('음성 파일 이름이 잘못됐습니다: ' + name);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const folder = audioFolder_(true), old = folder.getFilesByName(name);
    while (old.hasNext()) old.next().setTrashed(true);
    const f = folder.createFile(name, String(data || ''), 'application/json');
    CacheService.getScriptCache().put('af:' + name, f.getId(), 21600);
    return { name, size: f.getSize() };
  } finally {
    lock.releaseLock();
  }
}
function audioFile_(name) {
  const cache = CacheService.getScriptCache(), id = cache.get('af:' + name);
  if (id) { try { const f = DriveApp.getFileById(id); if (!f.isTrashed()) return f; } catch (e) { /* 다시 찾음 */ } }
  const folder = audioFolder_(false), it = folder && folder.getFilesByName(name);
  const f = it && it.hasNext() ? it.next() : null;
  if (f) cache.put('af:' + name, f.getId(), 21600);
  return f;
}
// 학생에게 세트 음성 보내기. 묶음은 이미 JSON 글자라 풀지 않고 그대로 감싸 보냅니다.
function audio_(r) {
  if (!items_(r.test, r.set)) return json_({ ok: false, error: 'noset' });
  const f = audioFile_(packName_(r.test, r.set));
  if (!f) return json_({ ok: false, error: 'noaudio' });
  return ContentService.createTextOutput('{"ok":true,"pack":' + f.getBlob().getDataAsString() + '}').setMimeType(ContentService.MimeType.JSON);
}
// 관리자 화면용: 올라온 음성 파일 수와 음성이 빠진 세트
function audioStatus_() {
  const folder = audioFolder_(false), names = new Set();
  if (folder) { const it = folder.getFiles(); while (it.hasNext()) { const f = it.next(); if (!f.isTrashed() && PACK_RE.test(f.getName())) names.add(f.getName()); } }
  const missing = [];
  catalog_().forEach(b => b.tests.forEach(t => t.sets.forEach(s => {
    if (!names.has(packName_(t.id, s.id))) missing.push(b.title + ' · ' + t.title + (t.sub ? ' · ' + t.sub : '') + ' · ' + s.title);
  })));
  return { folder: !!folder, files: names.size, missing };
}

// 편집기에서 실행(관리자 페이지 없이 처음 설정할 때): 「딕테이션 음원」 폴더에 넣어 둔 문제 파일(dictation-problems-*.json)을
// 「문제」 탭으로 불러오고, 입장 코드가 없으면 새로 만듭니다. 결과(입장 코드 포함)는 실행 로그에 나옵니다.
function setupFromDrive() {
  const folder = audioFolder_(false);
  if (!folder) throw new Error('드라이브에 「' + AUDIO_FOLDER + '」 폴더가 없습니다.');
  const it = folder.getFiles(), done = [];
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed() && /^dictation-problems.*\.json$/.test(f.getName())) done.push(f.getName() + ' → 문장 ' + upload_(JSON.parse(f.getBlob().getDataAsString())) + '개');
  }
  bump_();
  if (!settings_().code) props_().setProperty('ENTRY_CODE', 'usher' + String(Math.floor(1000 + Math.random() * 9000)));
  const st = audioStatus_();
  Logger.log('문제: ' + (done.join(', ') || '불러온 파일 없음'));
  Logger.log('책: ' + catalog_().map(b => b.title + ' (테스트 ' + b.tests.length + '개)').join(', '));
  Logger.log('음성: 세트 ' + st.files + '개, 음성 없는 세트 ' + st.missing.length + '개');
  Logger.log('입장 코드: ' + settings_().code);
}

/* ---------- 채점 ---------- */
const NUM = { mister: 'mr' };
'zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen seventeen eighteen nineteen twenty'.split(' ').forEach((w, i) => NUM[w] = String(i));
Object.assign(NUM, { thirty: '30', forty: '40', fifty: '50', sixty: '60', seventy: '70', eighty: '80', ninety: '90' });
const esc_ = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
function norm_(s) {
  s = s.toLowerCase().replace(/[’‘`]/g, "'").replace(/one hundred|a hundred/g, '100').replace(/[—–\-]/g, ' ').replace(/[^a-z0-9'\s]/g, '');
  return s.split(/\s+/).map(w => w.replace(/^'+|'+$/g, '')).filter(w => w && w !== 'dollars' && w !== 'dollar').map(w => NUM[w] || w);
}
function lcs_(a, b) {
  const n = a.length, m = b.length, d = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) d[i][j] = a[i] === b[j] ? d[i + 1][j + 1] + 1 : Math.max(d[i + 1][j], d[i][j + 1]);
  const fa = new Array(n).fill(false); let i = 0, j = 0;
  while (i < n && j < m) { if (a[i] === b[j]) { fa[i] = true; i++; j++; } else if (d[i + 1][j] >= d[i][j + 1]) i++; else j++; }
  return fa;
}
// 정답 ref 와 학생 답 typed 비교 → {hit,total,html} (html: 틀린 단어에 miss 표시한 정답)
function grade_(ref, typed) {
  const words = ref.split(/\s+/).filter(Boolean), tok = [], own = [];
  words.forEach((w, k) => norm_(w).forEach(t => { tok.push(t); own.push(k); }));
  const fa = lcs_(tok, norm_(String(typed || '')));
  const bad = new Set(); fa.forEach((ok, x) => { if (!ok) bad.add(own[x]); });
  return { hit: fa.filter(Boolean).length, total: tok.length, html: words.map((w, k) => `<span class="${bad.has(k) ? 'miss' : 'ok'}">${esc_(w)}</span>`).join(' ') };
}
function pieces_(it) {
  const out = []; let pos = 0;
  it.b.forEach(p => { const k = it.t.indexOf(p, pos); if (k < 0) return; if (k > pos) out.push({ x: it.t.slice(pos, k) }); out.push({ b: p }); pos = k + p.length; });
  if (pos < it.t.length) out.push({ x: it.t.slice(pos) }); return out;
}
const pct_ = (h, t) => t ? Math.round(100 * h / t) : 0;
// 한 문장 채점. a: 전체 받아쓰기면 글자, 빈칸이면 빈칸별 글자 배열 → {s, html, typed}
function gradeItem_(it, mode, a) {
  if (mode !== 'blank') { const g = grade_(it.t, a); return { s: pct_(g.hit, g.total), html: g.html, typed: String(a || '').slice(0, 400) }; }
  const arr = Array.isArray(a) ? a : []; let hit = 0, total = 0, j = 0; const typed = [];
  const html = pieces_(it).map(p => {
    if (p.x != null) return esc_(p.x);
    const v = String(arr[j++] || ''); typed.push(v.trim() || '∅');
    const g = grade_(p.b, v); hit += g.hit; total += g.total; return g.html;
  }).join('');
  return { s: pct_(hit, total), html, typed: typed.join(' / ').slice(0, 400) };
}
function gradeOne_(r) {
  const items = items_(r.test, r.set), it = items && items[r.k];
  if (!it) return { ok: false, error: 'noset' };
  return { ok: true, s: gradeItem_(it, r.mode, r.a).s };
}

/* ---------- 제출 ---------- */
function submit_(r) {
  const period = clean_(r.period), cls = clean_(r.cls), name = clean_(r.stu);
  if (!period || !cls || !name) return json_({ ok: false, error: 'auth' });
  const m = meta_(r.test, r.set), items = items_(r.test, r.set);
  if (!m || !items) return json_({ ok: false, error: 'noset' });
  const mode = r.mode === 'blank' ? 'blank' : 'full', ans = r.items || [];
  const graded = items.map((it, k) => Object.assign(gradeItem_(it, mode, (ans[k] || {}).a), { p: Math.max(0, Math.round(+(ans[k] || {}).p || 0)) }));
  const score = Math.round(graded.reduce((a, g) => a + g.s, 0) / graded.length);
  const secs = Math.max(0, Math.round(+r.secs || 0)), reps = Math.min(5, Math.max(1, Math.round(+r.reps || 1)));
  const test = m.book.title + ' · ' + m.test.title + (m.test.sub ? ' · ' + m.test.sub : '');
  const setName = m.set.title + (m.set.name ? ' ' + m.set.name : '');
  const modeKo = mode === 'blank' ? '빈칸' : '전체';
  const avgPlays = graded.reduce((a, g) => a + g.p, 0) / graded.length;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    const row = [now, text_(period), text_(cls), text_(name), text_(test), text_(setName), m.set.kind, modeKo, score,
      "'" + Math.floor(secs / 60) + ':' + String(secs % 60).padStart(2, '0'), secs, reps, Math.round(avgPlays * 10) / 10];
    sheet_(SUMMARY, SUMMARY_HEADER).appendRow(row);
    // 반별 탭. 시트 탭 이름에 쓸 수 없는 문자( [ ] : * ? / \ )는 빼고 만듭니다.
    const tab = (period + ' ' + cls).replace(/[\[\]:*?\/\\]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 95);
    sheet_(tab, SUMMARY_HEADER).appendRow(row);
    const rows = graded.map((g, k) => [now, text_(period), text_(cls), text_(name), text_(test), text_(setName), modeKo, k + 1, items[k].t, g.typed, g.s, g.p]);
    const sh = sheet_(DETAIL, DETAIL_HEADER);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  const show = settings_().showAnswers;
  return json_({ ok: true, score, showAnswers: show, items: graded.map(g => show ? { s: g.s, html: g.html } : { s: g.s }) });
}

// 예전 사이트(문장이 사이트에 있던 때)가 보내는 기록. LEGACY_UNTIL 까지만 받습니다.
function legacySubmit_(r) {
  const period = clean_(r.period), cls = clean_(r.cls), name = clean_(r.stu);
  if (!period || !cls || !name) return json_({ ok: false, error: 'auth' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date(), items = r.items || [];
    const avgPlays = items.length ? items.reduce((a, x) => a + (x.p || 0), 0) / items.length : 0;
    const row = [now, text_(period), text_(cls), text_(name), text_(r.test), text_(r.set), r.kind, r.mode, r.score,
      "'" + (r.time || ''), r.secs || 0, r.reps || 1, Math.round(avgPlays * 10) / 10];
    sheet_(SUMMARY, SUMMARY_HEADER).appendRow(row);
    const tab = (period + ' ' + cls).replace(/[\[\]:*?\/\\]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 95);
    sheet_(tab, SUMMARY_HEADER).appendRow(row);
    const rows = items.map(x => [now, text_(period), text_(cls), text_(name), text_(r.test), text_(r.set), r.mode, x.n, x.t, x.a, x.s, x.p]);
    if (rows.length) { const sh = sheet_(DETAIL, DETAIL_HEADER); sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows); }
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
