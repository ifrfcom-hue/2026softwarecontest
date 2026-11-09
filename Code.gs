/**
 * 2026 전곽 소프트웨어 경시대회 (프로젝트 부문) — 제출·팀 평가 웹앱 (Google Apps Script)
 *
 * 처음 한 번:
 *   1) 아래 CONFIG.ADMIN_PASSWORD 를 선생님만 아는 값으로 바꾸기
 *   2) 함수 목록에서 setup 실행 → 권한 승인 (드라이브 폴더·시트 자동 생성)
 *   3) 배포 → 새 배포 → 유형: 웹 앱
 *        - 실행 사용자: 나
 *        - 액세스 권한: 모든 사용자
 *   4) 웹 앱 URL 을 학생에게 공유 → 관리자 화면에서 양식 3종(신청서·결과보고서·AI 활용 기록지) 올리기
 */

const CONFIG = {
  ADMIN_PASSWORD: 'admin2316',
  MAX_FILE_MB: 20,
  APP_TITLE: '2026 전곽 소프트웨어 경시대회',
  FOLDER_NAME: '2026 전곽 SW 경시대회 제출함',
  SHEET_NAME: '2026 전곽 SW 경시대회 참가팀'
};

const KINDS = {
  apply:  { label: '참가신청서',      col: 'app' },
  report: { label: '결과보고서',      col: 'rep' },
  ai:     { label: 'AI활용기록지',    col: 'ai'  }
};

const HEADERS = ['id', 'teamName', 'field', 'devType', 'project', 'leader', 'members',
  'sourceUrl', 'videoUrl', 'salt', 'pwHash',
  'appFileId', 'appFileName', 'appAt', 'repFileId', 'repFileName', 'repAt', 'aiFileId', 'aiFileName', 'aiAt',
  'createdAt', 'updatedAt', 'workUrl', 'workAt'];

const VOTE_HEADERS = ['voterId', 'voterName', 'teamId', 'at'];

const DEFAULT_SETTINGS = {
  notice: '학생이 실제 문제를 스스로 발견하고 소프트웨어로 해결한 뒤, 그 효과를 증명하는 전 과정을 평가하는 문제해결 역량 중심 프로젝트 경시대회입니다.',
  schedule: [
    { date: '11. 16.(월)', label: '서류 검토 · 발표 순서 안내' },
    { date: '11. 17.(화) 19:00 ~', label: '발표 심사 · 관람 학생 투표' }
  ],
  /* 제출 기간 (한국 시간). 관리자 화면에서 바꿀 수 있음 */
  applyStart: '2026-11-04T00:00', applyEnd: '2026-11-11T23:59',   // 팀 등록 · 참가 신청서
  workStart: '2026-11-04T00:00',  workEnd: '2026-11-15T23:59',    // 작품 URL · 결과보고서 · AI 활용 기록지
  allowLate: true,   // 마감 뒤에도 받되 '마감 후'로 표시 (감점 확인용)
  deadline: '2026-11-15T23:59',
  contact: '영재교육정보부 정보과',
  registerOpen: true,
  submitOpen: true,
  voteOpen: false,
  voteMode: 'perTeam',      // perTeam: 팀마다 1번씩 / single: 1인 1팀
  showVotes: true,
  showcaseDate: '2026-11-17'  // 다른 팀이 작품 링크를 볼 수 있는 날 (대회 당일)
};

/* ───────── 웹앱 진입 ───────── */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* GitHub Pages 등 외부에 올린 index.html 이 호출하는 창구 */
const API = {
  getPublic: getPublic, registerTeam: registerTeam, unlockTeam: unlockTeam,
  uploadTeamFile: uploadTeamFile, removeTeamFile: removeTeamFile, updateTeam: updateTeam, deleteTeam: deleteTeam,
  setWorkUrl: setWorkUrl, myVotes: myVotes, vote: vote,
  adminLogin: adminLogin, adminUploadForm: adminUploadForm, adminRemoveForm: adminRemoveForm,
  adminSaveSettings: adminSaveSettings, adminResetPassword: adminResetPassword,
  adminDeleteTeam: adminDeleteTeam, adminResetVotes: adminResetVotes, adminSetFolder: adminSetFolder
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = Object.prototype.hasOwnProperty.call(API, req.fn) ? API[req.fn] : null;
    if (!fn) throw new Error('알 수 없는 요청입니다.');
    out = { ok: true, data: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function setup() {
  const s = store_();
  Logger.log('시트: ' + s.ss.getUrl());
  Logger.log('폴더: ' + s.folder.getUrl());
}

/* ───────── 저장소 ───────── */

function props_() { return PropertiesService.getScriptProperties(); }

function sheetWith_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 2000, headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < headers.length) {
    sh.getRange(1, 1, 2000, headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sh;
}

function store_() {
  const p = props_();
  let folder = null, ss = null;
  try { folder = p.getProperty('FOLDER_ID') ? DriveApp.getFolderById(p.getProperty('FOLDER_ID')) : null; } catch (e) { folder = null; }
  if (!folder) { folder = DriveApp.createFolder(CONFIG.FOLDER_NAME); p.setProperty('FOLDER_ID', folder.getId()); }
  try { ss = p.getProperty('SS_ID') ? SpreadsheetApp.openById(p.getProperty('SS_ID')) : null; } catch (e) { ss = null; }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.SHEET_NAME);
    ss.getSheets()[0].setName('_');
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    p.setProperty('SS_ID', ss.getId());
  }
  const schema = 'v' + HEADERS.length + '_' + VOTE_HEADERS.length + '_' + ss.getId();
  let teams, votes;
  if (p.getProperty('SCHEMA') === schema) {
    teams = ss.getSheetByName('teams'); votes = ss.getSheetByName('votes');
  }
  if (!teams || !votes) {
    teams = sheetWith_(ss, 'teams', HEADERS);
    votes = sheetWith_(ss, 'votes', VOTE_HEADERS);
    const blank = ss.getSheetByName('_');
    if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);
    p.setProperty('SCHEMA', schema);
  }
  return { ss: ss, sheet: teams, votes: votes, folder: folder };
}

function sub_(folder, name) {
  const it = folder.getFoldersByName(name);
  return it.hasNext() ? it.next() : folder.createFolder(name);
}

function rows_(sheet, headers) {
  const v = sheet.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < v.length; r++) {
    if (!v[r][0]) continue;
    const o = { _row: r + 1 };
    headers.forEach(function (h, i) { o[h] = String(v[r][i] == null ? '' : v[r][i]); });
    out.push(o);
  }
  return out;
}

function readTeams_(sheet) { return rows_(sheet, HEADERS); }

function writeTeam_(sheet, t) {
  const row = HEADERS.map(function (h) { return t[h] == null ? '' : String(t[h]); });
  if (t._row) sheet.getRange(t._row, 1, 1, HEADERS.length).setValues([row]);
  else sheet.appendRow(row);
}

function findTeam_(sheet, id) {
  const t = readTeams_(sheet).filter(function (x) { return x.id === id; })[0];
  if (!t) throw new Error('팀을 찾을 수 없습니다. 새로고침 후 다시 시도하세요.');
  return t;
}

function settings_() {
  const raw = props_().getProperty('SETTINGS');
  const s = raw ? JSON.parse(raw) : {};
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) { if (s[k] === undefined) s[k] = DEFAULT_SETTINGS[k]; });
  return s;
}

function forms_() { return JSON.parse(props_().getProperty('FORMS') || '{}'); }

function url_(id) { return id ? 'https://drive.google.com/file/d/' + id + '/view' : ''; }

/* 여러 기기가 동시에 저장할 때 순서대로 처리 (기록하는 짧은 순간에만 잠금) */
const BUSY_MSG = '접속자가 많아 잠시 기다리는 중입니다. 잠시 후 다시 눌러 주세요.';
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error(BUSY_MSG);
  try { const out = fn(); bust_(); return out; } finally { lock.releaseLock(); }
}

/* 공개 정보는 10초 동안 캐시 → 많은 기기가 동시에 새로고침해도 시트를 한 번만 읽음 */
const PUB_CACHE = 'pub_v2';
function bust_() { try { CacheService.getScriptCache().remove(PUB_CACHE); } catch (e) { } }

/* ───────── 보안 ───────── */

function hash_(salt, pw) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return d.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function guard_(key, ok) {
  const cache = CacheService.getScriptCache();
  const k = 'fail_' + key;
  const n = Number(cache.get(k) || 0);
  if (n >= 5) throw new Error('비밀번호를 5번 틀렸습니다. 10분 뒤 다시 시도하세요.');
  if (!ok) { cache.put(k, String(n + 1), 600); throw new Error('비밀번호가 맞지 않습니다. (' + (n + 1) + '/5)'); }
  cache.remove(k);
}

function checkTeam_(t, pw) { guard_(t.id, hash_(t.salt, String(pw || '')) === t.pwHash); }
function checkAdmin_(pw) { guard_('admin', String(pw || '') === CONFIG.ADMIN_PASSWORD); }
function clean_(s, max) { return String(s || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max); }
function cleanUrl_(s) { s = clean_(s, 300); return /^https?:\/\//i.test(s) ? s : (s ? 'https://' + s : ''); }
function validatePw_(pw) { if (String(pw || '').length < 4) throw new Error('비밀번호는 4자 이상으로 정해 주세요.'); }
function ids_(text) { return (String(text || '').match(/\d{5}/g) || []); }

/* ───────── 파일 ───────── */

function saveFile_(folder, sub, prefix, file) {
  if (!file || !file.data) throw new Error('파일이 비어 있습니다.');
  const bytes = Utilities.base64Decode(file.data);
  if (bytes.length > CONFIG.MAX_FILE_MB * 1024 * 1024) throw new Error(CONFIG.MAX_FILE_MB + 'MB 이하 파일만 올릴 수 있습니다.');
  const name = clean_(file.name, 120) || 'file';
  const f = sub_(folder, sub).createFile(Utilities.newBlob(bytes, file.mimeType || 'application/octet-stream', prefix + name));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* 도메인 정책으로 막힐 수 있음 */ }
  return { id: f.getId(), name: name };
}

function trash_(id) { if (!id) return; try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { } }

/* ───────── 응답 모양 ───────── */

function voteCounts_(votesSheet) {
  const c = {};
  rows_(votesSheet, VOTE_HEADERS).forEach(function (v) { c[v.teamId] = (c[v.teamId] || 0) + 1; });
  return c;
}

/* 'YYYY-MM-DDTHH:mm' (한국 시간) → 밀리초 */
function kst_(s) { s = String(s || ''); if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(s)) return null; const ms = new Date(s + ':00+09:00').getTime(); return isNaN(ms) ? null : ms; }
function fmtKst_(s) { const ms = kst_(s); if (ms === null) return ''; return Utilities.formatDate(new Date(ms), 'Asia/Seoul', 'M. d. HH:mm'); }
/* 'apply' 또는 'work' 기간이 지금 어떤 상태인지: off(관리자가 멈춤) / before / open / late / closed */
function phase_(st, which) {
  if (which === 'apply' ? !st.registerOpen : !st.submitOpen) return 'off';
  const now = Date.now();
  const a = kst_(which === 'apply' ? st.applyStart : st.workStart), b = kst_(which === 'apply' ? st.applyEnd : st.workEnd);
  if (a !== null && now < a) return 'before';
  if (b !== null && now > b + 59999) return st.allowLate ? 'late' : 'closed';
  return 'open';
}
function needOpen_(which, what) {
  const st = settings_(), p = phase_(st, which);
  const s = which === 'apply' ? st.applyStart : st.workStart, e = which === 'apply' ? st.applyEnd : st.workEnd;
  if (p === 'off') throw new Error(what + ' 접수가 멈춰 있습니다. 담당 선생님께 문의하세요.');
  if (p === 'before') throw new Error(what + ' 제출 기간 전입니다. (' + fmtKst_(s) + '부터)');
  if (p === 'closed') throw new Error(what + ' 제출이 마감되었습니다. (' + fmtKst_(e) + ' 마감)');
}

function today_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'); }
function showcaseToday_(st) { return !!st.showcaseDate && st.showcaseDate === today_(); }

function publicTeam_(t, counts, reveal) {
  return {
    hasWork: !!t.workUrl,
    workUrl: reveal ? (t.workUrl || '') : '',
    id: t.id, teamName: t.teamName, field: t.field, devType: t.devType, project: t.project,
    memberCount: 1 + (t.members ? t.members.split(',').filter(function (x) { return x.trim(); }).length : 0),
    hasApp: !!t.appFileId, hasRep: !!t.repFileId, hasAi: !!t.aiFileId,
    votes: counts ? (counts[t.id] || 0) : null,
    updatedAt: Number(t.updatedAt) || 0
  };
}

function detailTeam_(t) {
  const p = publicTeam_(t, null);
  p.leader = t.leader; p.members = t.members; p.sourceUrl = t.sourceUrl; p.videoUrl = t.videoUrl;
  p.workUrl = t.workUrl || ''; p.workAt = Number(t.workAt) || 0;
  Object.keys(KINDS).forEach(function (k) {
    const c = KINDS[k].col;
    p[k] = t[c + 'FileId'] ? { name: t[c + 'FileName'], url: url_(t[c + 'FileId']), at: Number(t[c + 'At']) } : null;
  });
  return p;
}

function getPublic() {
  const cache = CacheService.getScriptCache();
  try { const hit = cache.get(PUB_CACHE); if (hit) return JSON.parse(hit); } catch (e) { }
  const out = buildPublic_();
  try { cache.put(PUB_CACHE, JSON.stringify(out), 10); } catch (e) { }
  return out;
}

function buildPublic_() {
  const s = store_();
  const st = settings_();
  const f = forms_();
  const forms = {};
  Object.keys(KINDS).forEach(function (k) { if (f[k]) forms[k] = { name: f[k].name, url: url_(f[k].id), at: f[k].at }; });
  const counts = st.showVotes ? voteCounts_(s.votes) : null;
  const reveal = showcaseToday_(st);
  const teams = readTeams_(s.sheet).map(function (t) { return publicTeam_(t, counts, reveal); })
    .sort(function (a, b) { return a.teamName.localeCompare(b.teamName, 'ko'); });
  return { settings: st, forms: forms, teams: teams, maxMb: CONFIG.MAX_FILE_MB, showcaseToday: reveal, today: today_(), serverNow: Date.now() };
}

/* ───────── 팀 기능 ───────── */

function teamInfo_(t, info) {
  t.teamName = clean_(info.teamName, 20);
  t.field = clean_(info.field, 20);
  t.devType = clean_(info.devType, 30);
  t.project = clean_(info.project, 60);
  t.leader = clean_(info.leader, 30);
  t.members = clean_(info.members, 80);
  t.sourceUrl = cleanUrl_(info.sourceUrl);
  t.videoUrl = cleanUrl_(info.videoUrl);
  if (!t.teamName) throw new Error('팀명을 입력하세요.');
  if (!t.leader) throw new Error('대표 학생(학번 이름)을 입력하세요.');
  if (t.members.split(',').filter(function (x) { return x.trim(); }).length > 2) throw new Error('팀은 최대 3명(대표 포함)까지입니다.');
}

function registerTeam(info, file) {
  needOpen_('apply', '팀 등록·참가 신청서');
  const t = {};
  teamInfo_(t, info);
  validatePw_(info.password);
  const s = store_();
  if (readTeams_(s.sheet).some(function (x) { return x.teamName.toLowerCase() === t.teamName.toLowerCase(); })) throw new Error('이미 등록된 팀명입니다. 다른 팀명을 정해 주세요.');
  const saved = file ? saveFile_(s.folder, KINDS.apply.label, t.teamName + '_' + KINDS.apply.label + '_', file) : null;  // 파일 저장은 잠금 밖에서
  try {
    return withLock_(function () {
      if (readTeams_(s.sheet).some(function (x) { return x.teamName.toLowerCase() === t.teamName.toLowerCase(); })) throw new Error('이미 등록된 팀명입니다. 다른 팀명을 정해 주세요.');
      const now = Date.now();
      t.id = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
      t.salt = Utilities.getUuid();
      t.pwHash = hash_(t.salt, String(info.password));
      t.createdAt = now; t.updatedAt = now;
      if (saved) { t.appFileId = saved.id; t.appFileName = saved.name; t.appAt = now; }
      writeTeam_(s.sheet, t);
      return detailTeam_(t);
    });
  } catch (e) { if (saved) trash_(saved.id); throw e; }
}

function unlockTeam(id, pw) {
  const t = findTeam_(store_().sheet, id);
  checkTeam_(t, pw);
  return detailTeam_(t);
}

function uploadTeamFile(id, pw, kind, file) {
  if (!KINDS[kind]) throw new Error('알 수 없는 파일 종류입니다.');
  needOpen_(kind === 'apply' ? 'apply' : 'work', KINDS[kind].label);
  const s = store_();
  const t0 = findTeam_(s.sheet, id);
  checkTeam_(t0, pw);
  const saved = saveFile_(s.folder, KINDS[kind].label, t0.teamName + '_' + KINDS[kind].label + '_', file);  // 파일 저장은 잠금 밖에서
  try {
    return withLock_(function () {
      const t = findTeam_(s.sheet, id);  // 잠금 안에서 최신 행을 다시 읽음
      const c = KINDS[kind].col, now = Date.now();
      trash_(t[c + 'FileId']);
      t[c + 'FileId'] = saved.id; t[c + 'FileName'] = saved.name; t[c + 'At'] = now; t.updatedAt = now;
      writeTeam_(s.sheet, t);
      return detailTeam_(t);
    });
  } catch (e) { trash_(saved.id); throw e; }
}

/* 작품 URL 저장 (빈 값이면 지움) */
function setWorkUrl(id, pw, url) {
  return withLock_(function () {
    const s = store_();
    const t = findTeam_(s.sheet, id);
    checkTeam_(t, pw);
    needOpen_('work', '작품 URL');
    const u = clean_(url, 300);
    if (u && !/^https?:\/\/[^\s.]+\.[^\s]+$/i.test(u)) throw new Error('http:// 또는 https:// 로 시작하는 전체 주소를 넣어 주세요.');
    t.workUrl = u; t.workAt = u ? Date.now() : ''; t.updatedAt = Date.now();
    writeTeam_(s.sheet, t);
    return detailTeam_(t);
  });
}

function removeTeamFile(id, pw, kind) {
  if (!KINDS[kind]) throw new Error('알 수 없는 파일 종류입니다.');
  return withLock_(function () {
    const s = store_();
    const t = findTeam_(s.sheet, id);
    checkTeam_(t, pw);
    const c = KINDS[kind].col;
    trash_(t[c + 'FileId']);
    t[c + 'FileId'] = ''; t[c + 'FileName'] = ''; t[c + 'At'] = ''; t.updatedAt = Date.now();
    writeTeam_(s.sheet, t);
    return detailTeam_(t);
  });
}

function updateTeam(id, pw, info) {
  return withLock_(function () {
    const s = store_();
    const teams = readTeams_(s.sheet);
    const t = teams.filter(function (x) { return x.id === id; })[0];
    if (!t) throw new Error('팀을 찾을 수 없습니다.');
    checkTeam_(t, pw);
    teamInfo_(t, info);
    if (teams.some(function (x) { return x.id !== id && x.teamName.toLowerCase() === t.teamName.toLowerCase(); })) throw new Error('이미 등록된 팀명입니다.');
    if (info.newPassword) { validatePw_(info.newPassword); t.salt = Utilities.getUuid(); t.pwHash = hash_(t.salt, String(info.newPassword)); }
    t.updatedAt = Date.now();
    writeTeam_(s.sheet, t);
    return detailTeam_(t);
  });
}

function deleteTeam(id, pw) {
  return withLock_(function () {
    const s = store_();
    const t = findTeam_(s.sheet, id);
    checkTeam_(t, pw);
    removeTeam_(s, t);
    return true;
  });
}

function removeTeam_(s, t) {
  trash_(t.appFileId); trash_(t.repFileId); trash_(t.aiFileId);
  s.sheet.deleteRow(t._row);
  const vs = rows_(s.votes, VOTE_HEADERS).filter(function (v) { return v.teamId === t.id; }).map(function (v) { return v._row; });
  vs.sort(function (a, b) { return b - a; }).forEach(function (r) { s.votes.deleteRow(r); });
}

/* ───────── 팀 평가(투표) ───────── */

function normVoter_(voterId, voterName) {
  const id = String(voterId || '').replace(/\D/g, '');
  if (!/^\d{5}$/.test(id)) throw new Error('학번 5자리를 입력하세요. (예: 10101)');
  const name = clean_(voterName, 10);
  if (!name) throw new Error('이름을 입력하세요.');
  return { id: id, name: name };
}

function myVotes(voterId, voterName) {
  const v = normVoter_(voterId, voterName);
  const s = store_();
  return rows_(s.votes, VOTE_HEADERS).filter(function (x) { return x.voterId === v.id; }).map(function (x) { return x.teamId; });
}

function vote(voterId, voterName, teamId) {
  return withLock_(function () {
    const st = settings_();
    if (!st.voteOpen) throw new Error('지금은 투표 시간이 아닙니다.');
    const v = normVoter_(voterId, voterName);
    const s = store_();
    const t = findTeam_(s.sheet, teamId);
    if (ids_(t.leader + ',' + t.members).indexOf(v.id) >= 0) throw new Error('우리 팀에는 투표할 수 없습니다.');
    const mine = rows_(s.votes, VOTE_HEADERS).filter(function (x) { return x.voterId === v.id; });
    if (mine.some(function (x) { return x.teamId === teamId; })) throw new Error('이 팀에는 이미 투표했습니다.');
    if (st.voteMode === 'single' && mine.length) throw new Error('이미 한 팀에 투표했습니다. (1인 1팀)');
    s.votes.appendRow([v.id, v.name, teamId, String(Date.now())]);
    bust_();
    const out = buildPublic_();
    out.myVotes = mine.map(function (x) { return x.teamId; }).concat([teamId]);
    return out;
  });
}

/* ───────── 관리자 ───────── */

function adminLogin(pw) {
  checkAdmin_(pw);
  const s = store_();
  const counts = voteCounts_(s.votes);
  const voters = {};
  rows_(s.votes, VOTE_HEADERS).forEach(function (v) { voters[v.voterId] = 1; });
  return {
    teams: readTeams_(s.sheet).map(function (t) { const d = detailTeam_(t); d.votes = counts[t.id] || 0; return d; })
      .sort(function (a, b) { return a.teamName.localeCompare(b.teamName, 'ko'); }),
    voterCount: Object.keys(voters).length,
    sheetUrl: s.ss.getUrl(), folderUrl: s.folder.getUrl(),
    folderName: s.folder.getName(), sheetName: s.ss.getName(), live: true
  };
}

function folderIdFrom_(link) {
  const s = String(link || '').trim();
  const m = s.match(/folders\/([A-Za-z0-9_-]{10,})/) || s.match(/[?&]id=([A-Za-z0-9_-]{10,})/) || s.match(/^([A-Za-z0-9_-]{10,})$/);
  if (!m) throw new Error('드라이브 폴더 주소를 확인해 주세요. (https://drive.google.com/drive/folders/... 형식)');
  return m[1];
}

/* 제출물을 저장할 드라이브 폴더를 바꾸기 (이미 받은 파일은 원래 폴더에 남음) */
function adminSetFolder(pw, link) {
  checkAdmin_(pw);
  const id = folderIdFrom_(link);
  let folder;
  try { folder = DriveApp.getFolderById(id); folder.getName(); }
  catch (e) { throw new Error('그 폴더를 열 수 없습니다. 웹 앱을 배포한 구글 계정이 편집할 수 있는 폴더인지 확인해 주세요.'); }
  props_().setProperty('FOLDER_ID', id);
  try { DriveApp.getFileById(props_().getProperty('SS_ID')).moveTo(folder); } catch (e) { }
  return adminLogin(pw);
}

function adminUploadForm(pw, kind, file) {
  checkAdmin_(pw);
  if (!KINDS[kind]) throw new Error('알 수 없는 양식입니다.');
  const saved = saveFile_(store_().folder, '양식', '', file);
  withLock_(function () {
    const f = forms_();
    if (f[kind]) trash_(f[kind].id);
    f[kind] = { id: saved.id, name: saved.name, at: Date.now() };
    props_().setProperty('FORMS', JSON.stringify(f));
  });
  return buildPublic_();
}

function adminRemoveForm(pw, kind) {
  checkAdmin_(pw);
  const f = forms_();
  if (f[kind]) { trash_(f[kind].id); delete f[kind]; }
  props_().setProperty('FORMS', JSON.stringify(f));
  bust_();
  return buildPublic_();
}

function adminSaveSettings(pw, s) {
  checkAdmin_(pw);
  const dt = function (v, name) { v = clean_(v, 16); if (v && kst_(v) === null) throw new Error(name + ' 날짜 형식이 올바르지 않습니다.'); return v; };
  const as = dt(s.applyStart, '신청 시작'), ae = dt(s.applyEnd, '신청 마감'), ws = dt(s.workStart, '작품 시작'), we = dt(s.workEnd, '작품 마감');
  if (as && ae && kst_(as) > kst_(ae)) throw new Error('신청서 제출 기간: 시작이 마감보다 늦습니다.');
  if (ws && we && kst_(ws) > kst_(we)) throw new Error('작품 제출 기간: 시작이 마감보다 늦습니다.');
  const clean = {
    applyStart: as, applyEnd: ae, workStart: ws, workEnd: we, allowLate: !!s.allowLate,
    notice: clean_(s.notice, 1000),
    schedule: (s.schedule || []).slice(0, 12).map(function (x) { return { date: clean_(x.date, 40), label: clean_(x.label, 40) }; })
      .filter(function (x) { return x.date || x.label; }),
    deadline: we,
    contact: clean_(s.contact, 80),
    registerOpen: !!s.registerOpen, submitOpen: !!s.submitOpen, voteOpen: !!s.voteOpen,
    voteMode: s.voteMode === 'single' ? 'single' : 'perTeam',
    showVotes: !!s.showVotes,
    showcaseDate: /^\d{4}-\d{2}-\d{2}$/.test(String(s.showcaseDate || '')) ? s.showcaseDate : ''
  };
  props_().setProperty('SETTINGS', JSON.stringify(clean));
  bust_();
  return buildPublic_();
}

function adminResetPassword(pw, id, newPw) {
  checkAdmin_(pw);
  validatePw_(newPw);
  return withLock_(function () {
    const s = store_();
    const t = findTeam_(s.sheet, id);
    t.salt = Utilities.getUuid(); t.pwHash = hash_(t.salt, String(newPw)); t.updatedAt = Date.now();
    writeTeam_(s.sheet, t);
    CacheService.getScriptCache().remove('fail_' + id);
    return true;
  });
}

function adminDeleteTeam(pw, id) {
  checkAdmin_(pw);
  return withLock_(function () {
    const s = store_();
    removeTeam_(s, findTeam_(s.sheet, id));
    return true;
  });
}

function adminResetVotes(pw) {
  checkAdmin_(pw);
  return withLock_(function () {
    const s = store_();
    const n = s.votes.getLastRow();
    if (n > 1) s.votes.deleteRows(2, n - 1);
    return true;
  });
}
