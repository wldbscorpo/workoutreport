/**
 * PT 운동기록 - Google Apps Script 서버 코드
 * 이 스크립트는 구글 시트에 붙여서(확장 프로그램 → Apps Script) 사용합니다.
 * 같은 프로젝트에 HTML 파일 'index' (index.html 내용)를 함께 넣어야 합니다.
 */

// 시트 탭별 열 구성: [코드에서 쓰는 이름, 시트 머리글]
const SCHEMA = {
  '회원': [['id','ID'],['name','이름'],['gender','성별'],['birthYear','출생 연도'],['job','직업·생활 환경'],['regCount','등록 횟수'],
          ['monthGoal','이번 달 목표'],['purpose','운동 목적'],['freq','운동 빈도'],['sleep','수면 패턴'],['diet','식습관'],['alcohol','음주 습관'],
          ['medical','병력 및 약물'],['special','특이사항'],['caution','주의사항'],['painArea','제일 불편한 부위'],['painMove','제일 불편한 동작'],
          ['breathing','호흡'],['createdAt','만든 시각'],['updatedAt','수정 시각']],
  '수업': [['id','ID'],['memberId','회원 ID'],['date','날짜'],['no','회차'],['memo','수업 메모'],['updatedAt','수정 시각']],
  '세트': [['sessionId','수업 ID'],['memberId','회원 ID'],['date','날짜'],['exOrder','운동 순서'],['exercise','운동'],['note','추가 메모'],['part','부위'],
          ['setNo','세트'],['weightText','무게 입력'],['kg','환산 kg'],['band','밴드'],['mode','횟수/시간'],['reps','횟수'],['sec','시간(초)']],
  '인바디': [['id','ID'],['memberId','회원 ID'],['date','측정일'],['weight','체중'],['smm','골격근량'],['fatMass','체지방량'],['fatPct','체지방률'],
           ['abdominal','복부비만률'],['visceral','내장지방레벨']],
  '평가': [['id','ID'],['memberId','회원 ID'],['date','평가일'],['itemsJson','항목(JSON)'],['memo','종합 메모']],
  '운동목록': [['name','운동'],['part','부위']],
  '설정': [['key','항목'],['valueJson','값(JSON)']]
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('PT 운동기록')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 시트 도우미 ---------- */
function keys_(name) { return SCHEMA[name].map(function (c) { return c[0]; }); }

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    const labels = SCHEMA[name].map(function (c) { return c[1]; });
    sh.getRange(1, 1, 1, labels.length).setValues([labels]).setFontWeight('bold').setBackground('#e9eef5');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readRows_(name) {
  const sh = sheet_(name);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const ks = keys_(name);
  const vals = sh.getRange(2, 1, n, ks.length).getDisplayValues();
  return vals.map(function (r) {
    const o = {};
    ks.forEach(function (k, i) { o[k] = r[i]; });
    return o;
  });
}

function toRow_(name, obj) {
  return keys_(name).map(function (k) {
    const v = obj[k];
    return v === undefined || v === null ? '' : String(v);
  });
}

// 모든 값을 '텍스트'로 저장해 날짜·숫자가 자동 변환되지 않게 합니다.
function writeRows_(sh, startRow, rows) {
  if (!rows.length) return;
  const rg = sh.getRange(startRow, 1, rows.length, rows[0].length);
  rg.setNumberFormat('@');
  rg.setValues(rows);
}

function findRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

function upsert_(name, obj) {
  const sh = sheet_(name);
  const row = toRow_(name, obj);
  const r = findRow_(sh, String(obj.id));
  writeRows_(sh, r > 0 ? r : sh.getLastRow() + 1, [row]);
}

function deleteWhere_(name, key, value) {
  const sh = sheet_(name);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const col = keys_(name).indexOf(key) + 1;
  const vals = sh.getRange(2, col, n, 1).getDisplayValues();
  const rows = [];
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === value) rows.push(i + 2);
  // 아래쪽부터 연속 구간 단위로 삭제
  for (let i = rows.length - 1; i >= 0;) {
    let end = rows[i], start = end;
    while (i - 1 >= 0 && rows[i - 1] === start - 1) { i--; start--; }
    sh.deleteRows(start, end - start + 1);
    i--;
  }
}

function replaceAll_(name, rows) {
  const sh = sheet_(name);
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, keys_(name).length).clearContent();
  writeRows_(sh, 2, rows);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function parseJSON_(s, f) { try { return JSON.parse(s); } catch (e) { return f; } }

/* ---------- 페이지에서 부르는 함수 ---------- */
function loadAll() {
  const members = readRows_('회원');
  const sessRows = readRows_('수업');
  const setRows = readRows_('세트');
  const bySession = {};
  setRows.forEach(function (r) { (bySession[r.sessionId] = bySession[r.sessionId] || []).push(r); });
  const sessions = sessRows.map(function (s) {
    const rows = (bySession[s.id] || []).sort(function (a, b) { return (+a.exOrder - +b.exOrder) || (+a.setNo - +b.setNo); });
    const exs = [];
    let cur = null, curOrder = null;
    rows.forEach(function (r) {
      if (r.exOrder !== curOrder) {
        cur = { name: r.exercise, note: r.note, part: r.part, sets: [] };
        exs.push(cur); curOrder = r.exOrder;
      }
      if (r.setNo === '0') return; // 세트 없이 저장된 운동
      cur.sets.push({ w: r.weightText, kg: r.kg, band: r.band, mode: r.mode, reps: r.reps, sec: r.sec });
    });
    return { id: s.id, memberId: s.memberId, date: s.date, no: s.no, memo: s.memo, updatedAt: s.updatedAt, exercises: exs };
  });
  const settings = {};
  readRows_('설정').forEach(function (r) { settings[r.key] = parseJSON_(r.valueJson, null); });
  return {
    members: members,
    sessions: sessions,
    inbody: readRows_('인바디'),
    assessments: readRows_('평가').map(function (a) { return { id: a.id, memberId: a.memberId, date: a.date, items: parseJSON_(a.itemsJson, {}), memo: a.memo }; }),
    catalog: readRows_('운동목록').filter(function (c) { return c.name; }),
    settings: settings
  };
}

function saveMember(m) { return withLock_(function () { upsert_('회원', m); return true; }); }

function deleteMember(id) {
  return withLock_(function () {
    deleteWhere_('회원', 'id', id);
    deleteWhere_('수업', 'memberId', id);
    deleteWhere_('세트', 'memberId', id);
    deleteWhere_('인바디', 'memberId', id);
    deleteWhere_('평가', 'memberId', id);
    return true;
  });
}

function saveSession(s) {
  return withLock_(function () {
    upsert_('수업', { id: s.id, memberId: s.memberId, date: s.date, no: s.no, memo: s.memo, updatedAt: s.updatedAt });
    deleteWhere_('세트', 'sessionId', s.id);
    const rows = [];
    (s.exercises || []).forEach(function (ex, i) {
      const base = { sessionId: s.id, memberId: s.memberId, date: s.date, exOrder: i + 1, exercise: ex.name, note: ex.note, part: ex.part };
      if (!ex.sets || !ex.sets.length) { rows.push(toRow_('세트', Object.assign({}, base, { setNo: 0 }))); return; }
      ex.sets.forEach(function (st, j) {
        rows.push(toRow_('세트', Object.assign({}, base, {
          setNo: j + 1, weightText: st.w, kg: st.kg, band: st.band ? '1' : '', mode: st.mode, reps: st.reps, sec: st.sec
        })));
      });
    });
    const sh = sheet_('세트');
    writeRows_(sh, sh.getLastRow() + 1, rows);
    return true;
  });
}

function deleteSession(id) {
  return withLock_(function () { deleteWhere_('수업', 'id', id); deleteWhere_('세트', 'sessionId', id); return true; });
}

function saveInbody(r) { return withLock_(function () { upsert_('인바디', r); return true; }); }
function deleteInbody(id) { return withLock_(function () { deleteWhere_('인바디', 'id', id); return true; }); }

function saveAssessment(a) {
  return withLock_(function () {
    upsert_('평가', { id: a.id, memberId: a.memberId, date: a.date, itemsJson: JSON.stringify(a.items || {}), memo: a.memo });
    return true;
  });
}
function deleteAssessment(id) { return withLock_(function () { deleteWhere_('평가', 'id', id); return true; }); }

function saveCatalog(list) {
  return withLock_(function () { replaceAll_('운동목록', (list || []).map(function (c) { return toRow_('운동목록', c); })); return true; });
}

function saveSettings(obj) {
  return withLock_(function () {
    const rows = Object.keys(obj || {}).map(function (k) { return toRow_('설정', { key: k, valueJson: JSON.stringify(obj[k]) }); });
    replaceAll_('설정', rows);
    return true;
  });
}

// 운동 이름 바꾸기·합치기: 세트 탭의 운동 이름과 부위를 일괄 변경
function renameExercise(fromName, toName, part) {
  return withLock_(function () {
    const sh = sheet_('세트');
    const n = sh.getLastRow() - 1;
    if (n < 1) return true;
    const ks = keys_('세트');
    const ci = ks.indexOf('exercise'), pi = ks.indexOf('part');
    const rg = sh.getRange(2, 1, n, ks.length);
    const vals = rg.getDisplayValues();
    const norm = function (s) { return String(s || '').toLowerCase().replace(/\s+/g, ''); };
    const k = norm(fromName);
    let changed = false;
    vals.forEach(function (r) { if (norm(r[ci]) === k) { r[ci] = toName; r[pi] = part || r[pi]; changed = true; } });
    if (changed) { rg.setNumberFormat('@'); rg.setValues(vals); }
    return true;
  });
}
