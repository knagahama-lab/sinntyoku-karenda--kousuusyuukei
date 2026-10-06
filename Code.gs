/**
 * カレンダー工数管理（EnkinGT風・追加コストゼロ版）
 * Googleカレンダーの予定 → 工数ログ → タスク予実・集計・ボード／ガント／日報
 */
const SH = { CONF: '設定', TASK: 'タスク', LOG: '工数ログ', SUM: '集計', SIM: '移管シミュレーション', DAILY: '日報' };
const MEMBER_ROWS = 30, CAT_ROWS = 15, TASK_ROWS_INIT = 200;
const STATUSES = ['未着手', '進行中', '確認待ち', '完了'];
const LEVELS = ['大', '中', '小'];
const TASK_ID_RE = /\b[Tt]-(\d{1,5})\b/;
const PROJ_RE = /[@＠]([^\s　@＠#＃【】\[\]［］]+)/;
const CLR = { head: '#1f3864', sub: '#d9e1f2', yel: '#fff200', gry: '#f2f2f2', line: '#808080' };
// タスクシートの列（1始まり）
const TC = { id: 1, proj: 2, name: 3, owner: 4, dept: 5, cat: 6, status: 7, prio: 8, start: 9, due: 10, plan: 11, actual: 12, rate: 13, remain: 14, judge: 15, last: 16, memo: 17, calEv: 18, parent: 19, level: 20 };
const TASK_HEAD = ['タスクID', '案件', 'タスク名', '担当者', '部署', 'カテゴリ', 'ステータス', '優先度', '開始日', '期限', '予定工数(h)', '実績工数(h)', '消化率', '残工数(h)', '判定', '最終作業日', 'メモ', 'カレンダー予定ID', '親タスクID', '階層'];
const LOG_HEAD = ['日付', '開始', '終了', '時間(h)', '氏名', '部署', 'カテゴリ', '案件', 'タスクID', '件名', '判定方法', 'イベントID', 'カレンダーID', '上位タスクID'];
const DAILY_HEAD = ['日付', '氏名', '合計(h)', '業務内容（カレンダーから自動）', '所感・コメント', '更新日時'];
const FIELD_COL = { proj: TC.proj, name: TC.name, owner: TC.owner, cat: TC.cat, status: TC.status, prio: TC.prio, start: TC.start, due: TC.due, plan: TC.plan, memo: TC.memo, parent: TC.parent };

// ================= メニュー =================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('工数管理')
    .addItem('① 初期設定（シート作成）', 'setup')
    .addSeparator()
    .addItem('カレンダーから取り込む（直近）', 'importRecent')
    .addItem('全期間を取り込み直す', 'importAll')
    .addSeparator()
    .addItem('タスク用カレンダーを今すぐ同期', 'syncTaskCalendars')
    .addItem('業務分掌を大・中・小タスクに反映', 'importGyomu')
    .addItem('毎朝6時の自動取り込みをON（タスク同期も）', 'installTrigger')
    .addItem('自動取り込みをOFF', 'removeTrigger')
    .addSeparator()
    .addItem('ボード／ガント画面のURLを表示', 'showWebAppUrl')
    .addToUi();
}

// タスク名を入れたらIDを自動採番
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SH.TASK || e.range.getLastRow() < 2) return;
  assignTaskIds_(sh);
}

// ================= Webアプリ =================
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('工数・タスク管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAppData() {
  const cfg = readConfig_();
  const viewer = viewer_();
  const me = cfg.allMembers.find(m => viewer && m.calId.toLowerCase() === viewer);
  return {
    members: cfg.allMembers.map(m => ({ name: m.name, dept: m.dept, calId: m.calId })),
    isAdmin: isAdmin_(cfg), viewer: viewer,
    depts: cfg.depts,
    cats: cfg.cats.map(c => c.name),
    statuses: STATUSES,
    tasks: readTasks_(),
    lastImport: cfg.lastImport,
    today: ymd_(new Date()),
    me: me ? me.name : ''
  };
}

function updateTask(id, patch) {
  return withLock_(() => {
    const sh = sheet_(SH.TASK);
    const row = findTaskRow_(sh, id);
    if (!row) throw new Error('タスクが見つかりません: ' + id);
    writeTaskFields_(sh, row, patch || {});
    SpreadsheetApp.flush();
    return readTasks_();
  });
}

function createTask(p) {
  return withLock_(() => {
    const id = insertTask_(sheet_(SH.TASK), p, '');
    SpreadsheetApp.flush();
    return { id: id, tasks: readTasks_() };
  });
}

function insertTask_(sh, p, calEv) {
  const last = Math.max(sh.getLastRow(), 1);
  const v = last > 1 ? sh.getRange(2, 1, last - 1, 3).getValues() : [];
  const idx = v.findIndex(r => !r[0] && !String(r[2]).trim());
  const row = idx >= 0 ? idx + 2 : last + 1;
  const id = fmtId_(maxTaskNo_(v.map(r => r[0])) + 1);
  sh.getRange(row, TC.id).setValue(id);
  writeTaskFormulas_(sh, row);
  writeTaskFields_(sh, row, Object.assign({ status: '未着手', prio: '中' }, p || {}));
  if (calEv) sh.getRange(row, TC.calEv).setValue(calEv);
  return id;
}

function getDashboard(fromStr, toStr) {
  const cfg = readConfig_();
  const from = parseYmd_(fromStr), toEx = addDays_(parseYmd_(toStr), 1);
  const logs = readLogs_(from, toEx);
  const d1 = cfg.depts[0], d2 = cfg.depts[1];
  const cat = {}, person = {}, proj = {}, week = {};
  let total = 0, uncl = 0;
  const addDept = (o, l) => { if (l.dept === d1) o.a += l.h; else if (l.dept === d2) o.b += l.h; else o.o += l.h; };
  logs.forEach(l => {
    total += l.h;
    if (l.cat === '未分類') uncl += l.h;
    addDept(cat[l.cat] || (cat[l.cat] = { cat: l.cat, a: 0, b: 0, o: 0 }), l);
    const p = person[l.name] || (person[l.name] = { name: l.name, dept: l.dept, h: 0, uncl: 0, days: {} });
    p.h += l.h; p.days[l.ymd] = 1;
    if (l.cat === '未分類') p.uncl += l.h;
    const pk = l.proj || '（案件なし）';
    proj[pk] = (proj[pk] || 0) + l.h;
    const wk = ymd_(weekStart_(l.date));
    addDept(week[wk] || (week[wk] = { wk: wk, a: 0, b: 0, o: 0 }), l);
  });
  const parentOf = {};
  readTasks_().forEach(t => { parentOf[t.id] = t.parent; });
  const taskHours = {};
  logs.forEach(l => {
    let id = l.task, guard = 0;
    while (id && guard++ < 10) { taskHours[id] = (taskHours[id] || 0) + l.h; id = parentOf[id]; }
  });
  const order = cfg.cats.map(c => c.name).concat(['未分類']);
  const transfer = {};
  cfg.cats.forEach(c => { transfer[c.name] = c.transfer; });
  const pos = n => { const i = order.indexOf(n); return i < 0 ? 999 : i; };
  return {
    total: total, uncl: uncl, d1: d1, d2: d2, monthDays: cfg.monthDays, wage: cfg.wage,
    bizDays: bizDays_(from, addDays_(toEx, -1)),
    cats: Object.keys(cat).map(k => Object.assign(cat[k], { transfer: !!transfer[k] })).sort((x, y) => pos(x.cat) - pos(y.cat)),
    persons: Object.keys(person).map(k => { const p = person[k]; return { name: p.name, dept: p.dept, h: p.h, uncl: p.uncl, days: Object.keys(p.days).length }; }).sort((a, b) => b.h - a.h),
    projs: Object.keys(proj).map(k => ({ proj: k, h: proj[k] })).sort((a, b) => b.h - a.h),
    weeks: Object.keys(week).sort().map(k => week[k]),
    taskHours: taskHours
  };
}

function getDaily(name, ymdStr) {
  const d = parseYmd_(ymdStr);
  const rows = readLogs_(d, addDays_(d, 1)).filter(l => l.name === name)
    .sort((a, b) => a.start - b.start)
    .map(l => ({ s: hm_(l.start), e: hm_(l.end), h: l.h, cat: l.cat, proj: l.proj, task: l.task, title: l.title }));
  const saved = findDaily_(name, ymdStr);
  return { rows: rows, comment: saved ? saved.comment : '', updated: saved ? saved.updated : '' };
}

function saveDaily(name, ymdStr, comment) {
  return withLock_(() => {
    const day = getDaily(name, ymdStr);
    const total = Math.round(day.rows.reduce((s, r) => s + r.h, 0) * 100) / 100;
    const summary = day.rows.map(r => r.s + '-' + r.e + ' [' + r.cat + '] ' + r.title).join('\n');
    const sh = sheet_(SH.DAILY);
    const f = findDaily_(name, ymdStr);
    const row = f ? f.row : sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 6).setValues([[parseYmd_(ymdStr), name, total, summary, comment, new Date()]]);
    return { updated: Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd HH:mm') };
  });
}

// ================= カレンダー取り込み =================
function importRecent() {
  const cfg = readConfig_();
  const today = startOfDay_(new Date());
  let from = addDays_(today, -cfg.recentDays);
  if (cfg.startDate && cfg.startDate > from) from = cfg.startDate;
  if (sheet_(SH.LOG).getLastRow() < 2) from = cfg.startDate || addDays_(today, -28);
  if (cfg.taskCals.length) { try { syncTaskCalendars(); } catch (e) { console.error(e); } }
  return runImport_(readConfig_(), from, addDays_(today, 1));
}

function importAll() {
  const cfg = readConfig_();
  const today = startOfDay_(new Date());
  return runImport_(cfg, cfg.startDate || addDays_(today, -28), addDays_(today, 1));
}

function runImport_(cfg, from, toEx, onlyCalIds) {
  const tz = tz_();
  const taskMap = taskMap_();
  const fresh = [], okCals = {}, status = {};
  const only = onlyCalIds ? onlyCalIds.map(c => c.toLowerCase()) : null;
  cfg.members.filter(mb => !only || only.indexOf(mb.calId.toLowerCase()) >= 0).forEach(mb => {
    try {
      let n = 0, pageToken = null;
      do {
        const opt = { timeMin: from.toISOString(), timeMax: toEx.toISOString(), singleEvents: true, orderBy: 'startTime', maxResults: 2500 };
        if (pageToken) opt.pageToken = pageToken;
        const res = Calendar.Events.list(mb.calId, opt);
        (res.items || []).forEach(ev => {
          const row = eventToRow_(ev, mb, cfg, taskMap);
          if (row) { fresh.push(row); n++; }
        });
        pageToken = res.nextPageToken;
      } while (pageToken);
      okCals[mb.calId] = true;
      status[mb.row] = 'OK ' + n + '件（' + Utilities.formatDate(new Date(), tz, 'MM/dd HH:mm') + '）';
    } catch (e) {
      status[mb.row] = 'NG: ' + String(e.message || e).slice(0, 80) + '（カレンダー共有を確認）';
    }
  });

  // 今回取り込み直した期間×カレンダー以外の既存ログは残す
  const sh = sheet_(SH.LOG);
  const n = sh.getLastRow() - 1;
  const keep = n > 0 ? sh.getRange(2, 1, n, LOG_HEAD.length).getValues()
    .filter(r => r[0] !== '' && !(r[0] instanceof Date && r[0] >= from && r[0] < toEx && okCals[r[12]])) : [];
  const all = keep.concat(fresh).sort((a, b) => (a[0] - b[0]) || String(a[4]).localeCompare(String(b[4])) || (a[1] - b[1]));
  if (n > 0) sh.getRange(2, 1, n, LOG_HEAD.length).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, LOG_HEAD.length).setValues(all);

  const conf = cfg.sheet;
  Object.keys(status).forEach(r => conf.getRange(Number(r), 5).setValue(status[r]));
  conf.getRange('L12').setValue(new Date()).setNumberFormat('yyyy/mm/dd hh:mm');
  const ng = Object.keys(status).filter(r => status[r].indexOf('NG') === 0).length;
  const msg = fresh.length + '件を取り込みました（' + ymd_(from) + '〜' + ymd_(addDays_(toEx, -1)) + '）' +
    (ng ? ' ／ 失敗 ' + ng + '人（設定シートE列を確認）' : '');
  try { SpreadsheetApp.getActive().toast(msg, '工数管理', 8); } catch (e) {}
  return msg;
}

function eventToRow_(ev, mb, cfg, taskMap) {
  if (ev.status === 'cancelled' || !ev.start || !ev.start.dateTime) return null; // 終日予定は除外
  if (excludeReason_(ev, cfg)) return null;
  const s = new Date(ev.start.dateTime), e = new Date(ev.end.dateTime);
  const h = Math.round((e - s) / 36000) / 100;
  if (!(h > 0) || h > 24) return null;
  const c = classifyEvent_(ev, cfg, taskMap);
  const day = new Date(s.getFullYear(), s.getMonth(), s.getDate());
  return [day, s, e, h, mb.name, mb.dept, c.cat, c.proj, c.taskId, ev.summary || '（非公開の予定）', c.how, ev.id, mb.calId, ancestors_(taskMap, c.taskId)];
}

// 取り込まない理由（空文字＝取り込む）
function excludeReason_(ev, cfg) {
  if (ev.eventType && ['outOfOffice', 'workingLocation', 'birthday'].indexOf(ev.eventType) >= 0) return '不在・勤務場所';
  const me = (ev.attendees || []).find(a => a.self);
  if (me && me.responseStatus === 'declined') return '辞退';
  const title = ev.summary || '';
  if (cfg.excludes.some(k => title.indexOf(k) >= 0)) return '除外キーワード';
  return '';
}

// 優先順位：管理画面で付けたタスク → 件名/説明のタスクID → 管理画面で付けたカテゴリ → タグ → キーワード
function classifyEvent_(ev, cfg, taskMap) {
  const title = ev.summary || '', desc = ev.description || '';
  const xp = (ev.extendedProperties && ev.extendedProperties.private) || {};
  const idm = xp.kousuuTask ? String(xp.kousuuTask).match(TASK_ID_RE) : (title + ' ' + desc).match(TASK_ID_RE);
  const linked = !idm && taskMap._ev ? taskMap._ev[ev.recurringEventId || ev.id] : null; // タスク用カレンダーと同じ予定
  const taskId = idm ? fmtId_(Number(idm[1])) : (linked ? linked.id : '');
  const task = taskId ? taskMap[taskId] : null;
  let cat, how;
  if (task && task.cat) { cat = task.cat; how = xp.kousuuTask ? '手動（タスク）' : linked ? 'タスク予定' : 'タスクID'; }
  else if (xp.kousuuCat && cfg.cats.some(c => c.name === xp.kousuuCat)) { cat = xp.kousuuCat; how = '手動'; }
  else {
    const c = classify_(title, cfg.cats);
    cat = c[0]; how = c[1] + (taskId && !task ? '（ID不明）' : '');
  }
  const pm = title.match(PROJ_RE);
  return { taskId: taskId, cat: cat, how: how, proj: (task && task.proj) || (pm ? pm[1] : '') };
}

// 【カテゴリ】/#カテゴリ のタグ → 件名キーワード → 未分類
function classify_(title, cats) {
  const names = cats.map(c => c.name);
  const re = /[【\[［]([^】\]］]+)[】\]］]|[#＃]([^\s　#＃@＠【】\[\]［］]+)/g;
  let m;
  while ((m = re.exec(title)) !== null) {
    const t = (m[1] || m[2]).trim();
    if (names.indexOf(t) >= 0) return [t, 'タグ'];
  }
  const hit = cats.find(c => c.kws.some(k => title.indexOf(k) >= 0));
  if (hit) return [hit.name, 'キーワード'];
  return ['未分類', '未分類'];
}

// ================= メンバー管理（Web管理画面） =================
function getAdminData() {
  const cfg = readConfig_();
  const props = PropertiesService.getScriptProperties();
  return {
    isAdmin: isAdmin_(cfg), owner: ownerEmail_(), depts: cfg.depts, taskCals: cfg.taskCals.join(','), lastTaskSync: cfg.lastTaskSync,
    taskTrigger: ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'syncTaskCalendars'),
    members: cfg.sheet.getRange(5, 1, MEMBER_ROWS, 5).getValues().map((r, i) => {
      const calId = String(r[2]).trim();
      return { row: 5 + i, name: String(r[0]).trim(), dept: String(r[1]).trim(), calId: calId, on: r[3] === true,
        result: String(r[4] || ''), access: calId ? (props.getProperty('ROLE:' + calId.toLowerCase()) || '未確認') : '' };
    }).filter(m => m.name || m.calId)
  };
}

function saveMember(m) {
  const cfg = readConfig_();
  assertAdmin_(cfg);
  const calId = String(m.calId || '').trim(), name = String(m.name || '').trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(calId)) throw new Error('メールアドレスの形式が正しくありません');
  if (!name) throw new Error('氏名を入力してください');
  withLock_(() => {
    const sh = cfg.sheet;
    const v = sh.getRange(5, 1, MEMBER_ROWS, 3).getValues();
    const dup = v.findIndex((r, i) => String(r[2]).trim().toLowerCase() === calId.toLowerCase() && 5 + i !== m.row);
    if (dup >= 0) throw new Error('このメールアドレスは登録済みです（' + v[dup][0] + '）');
    let row = Number(m.row) || 0;
    if (row && (row < 5 || row >= 5 + MEMBER_ROWS)) throw new Error('行番号が不正です');
    if (!row) {
      const i = v.findIndex(r => !String(r[0]).trim() && !String(r[2]).trim());
      if (i < 0) throw new Error('登録できるのは' + MEMBER_ROWS + '人までです');
      row = 5 + i;
    }
    sh.getRange(row, 1, 1, 4).setValues([[name, String(m.dept || '').trim(), calId, m.on !== false]]);
  });
  const access = testAccess_(calId);
  return { access: access, data: getAdminData() };
}

function removeMember(row) {
  const cfg = readConfig_();
  assertAdmin_(cfg);
  row = Number(row);
  if (!(row >= 5 && row < 5 + MEMBER_ROWS)) throw new Error('行番号が不正です');
  withLock_(() => cfg.sheet.getRange(row, 1, 1, 5).setValues([['', '', '', false, '']]));
  return getAdminData();
}

function testMember(calId) {
  assertAdmin_(readConfig_());
  return { access: testAccess_(String(calId).trim()), data: getAdminData() };
}

function importMember(calId) {
  const cfg = readConfig_();
  assertAdmin_(cfg);
  const today = startOfDay_(new Date());
  const msg = runImport_(cfg, cfg.startDate || addDays_(today, -28), addDays_(today, 1), [String(calId)]);
  return { msg: msg, data: getAdminData() };
}

// 管理者アカウントが相手のカレンダーをどこまで操作できるか確認（カレンダーリストに追加される）
function testAccess_(calId) {
  let label;
  try {
    let entry;
    try { entry = Calendar.CalendarList.get(calId); } catch (e) { entry = Calendar.CalendarList.insert({ id: calId, hidden: true }); }
    label = { owner: '編集可', writer: '編集可', reader: '閲覧のみ', freeBusyReader: '空き時間のみ' }[entry.accessRole] || String(entry.accessRole);
  } catch (e) {
    label = 'アクセス不可';
  }
  PropertiesService.getScriptProperties().setProperty('ROLE:' + calId.toLowerCase(), label);
  return label;
}

// ================= カレンダー予定の表示・修正 =================
function getEvents(calId, fromYmd, toYmd) {
  const cfg = readConfig_();
  const mb = member_(cfg, calId);
  assertCanAccess_(cfg, mb.calId);
  const taskMap = taskMap_();
  const from = parseYmd_(fromYmd), toEx = addDays_(parseYmd_(toYmd), 1);
  const out = [];
  let pageToken = null;
  do {
    const opt = { timeMin: from.toISOString(), timeMax: toEx.toISOString(), singleEvents: true, orderBy: 'startTime', maxResults: 2500 };
    if (pageToken) opt.pageToken = pageToken;
    const res = Calendar.Events.list(mb.calId, opt);
    (res.items || []).forEach(ev => {
      if (ev.status === 'cancelled' || !ev.start || !ev.start.dateTime) return;
      const c = classifyEvent_(ev, cfg, taskMap);
      const xp = (ev.extendedProperties && ev.extendedProperties.private) || {};
      out.push({ id: ev.id, title: ev.summary || '', desc: ev.description || '',
        start: localIso_(ev.start.dateTime), end: localIso_(ev.end.dateTime),
        organizer: isOrganizer_(ev, mb.calId), task: c.taskId, cat: c.cat, how: c.how, proj: c.proj,
        manualTask: xp.kousuuTask || '', manualCat: xp.kousuuCat || '',
        excluded: excludeReason_(ev, cfg), link: ev.htmlLink || '' });
    });
    pageToken = res.nextPageToken;
  } while (pageToken);
  const role = PropertiesService.getScriptProperties().getProperty('ROLE:' + mb.calId.toLowerCase()) || '';
  return { calId: mb.calId, name: mb.name, access: role, events: out };
}

// p: {title, desc, start:'yyyy-MM-ddTHH:mm', end, task, cat}。task/catは件名を変えずに予定へ記録（非公開の拡張プロパティ）
function updateEvent(calId, eventId, p) {
  const cfg = readConfig_();
  const mb = member_(cfg, calId);
  assertCanAccess_(cfg, mb.calId);
  const ev = Calendar.Events.get(mb.calId, eventId);
  const oldStart = localIso_(ev.start.dateTime || ev.start.date), oldEnd = localIso_(ev.end.dateTime || ev.end.date);
  const res = { extendedProperties: { private: { kousuuTask: p.task || '', kousuuCat: p.cat || '' } } };
  const changed = (p.title != null && p.title !== (ev.summary || '')) || (p.desc != null && p.desc !== (ev.description || '')) ||
    (p.start && p.start !== oldStart) || (p.end && p.end !== oldEnd);
  let note = '';
  if (changed) {
    if (!isOrganizer_(ev, mb.calId)) note = '主催者ではない予定のため、件名・時間・説明は変更できません（分類だけ保存しました）';
    else {
      if (p.title != null) res.summary = p.title;
      if (p.desc != null) res.description = p.desc;
      if (p.start) res.start = { dateTime: p.start + ':00', timeZone: tz_() };
      if (p.end) res.end = { dateTime: p.end + ':00', timeZone: tz_() };
      if ((p.end || oldEnd) <= (p.start || oldStart)) throw new Error('終了は開始より後にしてください');
    }
  }
  Calendar.Events.patch(res, mb.calId, eventId, { sendUpdates: 'none' });
  reimportDays_(cfg, mb, [oldStart, p.start || oldStart]);
  return { note: note };
}

function createEvent(calId, p) {
  const cfg = readConfig_();
  const mb = member_(cfg, calId);
  assertCanAccess_(cfg, mb.calId);
  if (!p.title) throw new Error('件名を入力してください');
  if (!p.start || !p.end || p.end <= p.start) throw new Error('開始・終了を正しく入力してください');
  Calendar.Events.insert({
    summary: p.title, description: p.desc || '',
    start: { dateTime: p.start + ':00', timeZone: tz_() }, end: { dateTime: p.end + ':00', timeZone: tz_() },
    extendedProperties: { private: { kousuuTask: p.task || '', kousuuCat: p.cat || '' } }
  }, mb.calId, { sendUpdates: 'none' });
  reimportDays_(cfg, mb, [p.start]);
  return { note: '' };
}

function reimportDays_(cfg, mb, isoList) {
  if (!mb.on) return;
  const days = isoList.filter(Boolean).map(s => parseYmd_(s.slice(0, 10))).sort((a, b) => a - b);
  runImport_(cfg, days[0], addDays_(days[days.length - 1], 1), [mb.calId]);
}

// ================= 権限 =================
function viewer_() { try { return String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { return ''; } }
function ownerEmail_() { try { return String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) { return ''; } }
function isAdmin_(cfg) { const v = viewer_(); return !!v && (v === ownerEmail_() || cfg.admins.indexOf(v) >= 0); }
function assertAdmin_(cfg) { if (!isAdmin_(cfg)) throw new Error('管理者のみ操作できます（設定シートL13の管理者メールに追加してください）'); }
function assertCanAccess_(cfg, calId) {
  if (isAdmin_(cfg)) return;
  if (viewer_() && viewer_() === calId.toLowerCase()) return;
  throw new Error('このメンバーのカレンダーを見る・修正する権限がありません（本人か管理者のみ）');
}
function member_(cfg, calId) {
  const mb = cfg.allMembers.find(m => m.calId.toLowerCase() === String(calId).toLowerCase());
  if (!mb) throw new Error('登録されていないメンバーです: ' + calId);
  return mb;
}
function isOrganizer_(ev, calId) { return !ev.organizer || ev.organizer.self === true || String(ev.organizer.email || '').toLowerCase() === calId.toLowerCase(); }
function localIso_(s) { return s ? Utilities.formatDate(new Date(s), tz_(), "yyyy-MM-dd'T'HH:mm") : ''; }
function taskMap_() {
  const m = {};
  Object.defineProperty(m, '_ev', { value: {}, enumerable: false });
  readTasks_().forEach(t => { m[t.id] = t; if (t.calEv) m._ev[t.calEv] = t; });
  return m;
}

// ================= タスク用カレンダー → タスク自動追加 =================
// 共有カレンダー（例：営業やる事）に入った予定を1件＝1タスクとして「タスク」シートへ追加・更新する。
// カレンダー側が正：タスク名・開始日・期限・案件。システム側が正：ステータス・予定工数・担当者（空欄時のみ補完）。
function syncTaskCalendars() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return '他の同期が実行中です';
  try {
    const cfg = readConfig_();
    if (!cfg.taskCals.length) return 'タスク用カレンダーIDが未設定です（設定シートL14 または Web画面の管理タブ）';
    const msg = syncTaskCalendars_(cfg);
    try { SpreadsheetApp.getActive().toast(msg, '工数管理', 8); } catch (e) {}
    return msg;
  } finally { lock.releaseLock(); }
}

function syncTaskCalendars_(cfg) {
  const sh = sheet_(SH.TASK);
  ensureTaskColumns_(sh);
  const today = startOfDay_(new Date());
  let from = addDays_(today, -90);
  if (cfg.startDate && cfg.startDate < from) from = cfg.startDate;
  const to = addDays_(today, 366);
  const memberByMail = {};
  cfg.allMembers.forEach(m => { if (m.calId) memberByMail[m.calId.toLowerCase()] = m.name; });
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, TASK_HEAD.length).getValues() : [];
  const byEv = {};
  rows.forEach((r, i) => { const k = String(r[TC.calEv - 1] || ''); if (k) byEv[k] = { row: i + 2, r: r }; });
  let added = 0, updated = 0, closed = 0;
  const errs = [];
  cfg.taskCals.forEach(calId => {
    try {
      let pageToken = null;
      do {
        const opt = { timeMin: from.toISOString(), timeMax: to.toISOString(), singleEvents: false, showDeleted: true, maxResults: 2500 };
        if (pageToken) opt.pageToken = pageToken;
        const res = Calendar.Events.list(calId, opt);
        (res.items || []).forEach(ev => {
          const cur = byEv[ev.id];
          if (ev.status === 'cancelled') {
            if (cur && cur.row > 0 && cur.r[TC.status - 1] !== '完了') {
              writeTaskFields_(sh, cur.row, { status: '完了', memo: '[カレンダーで削除] ' + cur.r[TC.memo - 1] });
              closed++;
            }
            return;
          }
          if (!ev.start) return;
          const t = eventToTask_(ev, cfg, memberByMail);
          if (!cur) {
            insertTask_(sh, { name: t.name, proj: t.proj, owner: t.owner, cat: t.cat, start: t.start, due: t.due, plan: t.plan, memo: t.memo, status: t.done ? '完了' : '未着手' }, ev.id);
            byEv[ev.id] = { row: -1, r: [] };
            added++;
          } else if (cur.row > 0) {
            const r = cur.r, patch = {};
            if (String(r[TC.name - 1]) !== t.name) patch.name = t.name;
            if (ymd_(r[TC.start - 1]) !== t.start) patch.start = t.start;
            if (ymd_(r[TC.due - 1]) !== t.due) patch.due = t.due;
            if (t.proj && String(r[TC.proj - 1]) !== t.proj) patch.proj = t.proj;
            if (t.done && r[TC.status - 1] !== '完了') patch.status = '完了';
            if (!String(r[TC.owner - 1]).trim() && t.owner) patch.owner = t.owner;
            if (Object.keys(patch).length) { writeTaskFields_(sh, cur.row, patch); updated++; }
          }
        });
        pageToken = res.nextPageToken;
      } while (pageToken);
    } catch (e) {
      errs.push(calId.slice(0, 24) + '…: ' + String(e.message || e).slice(0, 80));
    }
  });
  const msg = 'タスク同期：追加 ' + added + '件・更新 ' + updated + '件' + (closed ? '・削除→完了 ' + closed + '件' : '') +
    (errs.length ? ' ／ 失敗：' + errs.join(' / ') : '');
  cfg.sheet.getRange('L15').setValue(Utilities.formatDate(new Date(), tz_(), 'MM/dd HH:mm') + ' ' + msg);
  return msg;
}

function eventToTask_(ev, cfg, memberByMail) {
  const title = String(ev.summary || '（件名なし）').trim();
  const allDay = !!ev.start.date;
  const s = allDay ? parseYmd_(ev.start.date) : new Date(ev.start.dateTime);
  let e = allDay ? addDays_(parseYmd_(ev.end.date), -1) : new Date(ev.end.dateTime);
  if (e < s) e = s;
  const desc = stripHtml_(ev.description || '');
  const pw = desc.match(/(?:予定)?工数\s*[:：]?\s*(\d+(?:\.\d+)?)/);
  const plan = pw ? Number(pw[1]) : (allDay ? '' : Math.round((new Date(ev.end.dateTime) - new Date(ev.start.dateTime)) / 36000) / 100);
  // 担当者：主催者→作成者→参加者の順に、登録メンバーのメールと一致した人
  const mails = [ev.organizer && ev.organizer.email, ev.creator && ev.creator.email]
    .concat((ev.attendees || []).filter(a => !a.resource && a.responseStatus !== 'declined').map(a => a.email));
  const owner = mails.map(x => memberByMail[String(x || '').toLowerCase()]).find(Boolean) || '';
  const pm = title.match(PROJ_RE);
  const cat = classify_(title, cfg.cats)[0];
  return {
    name: title.replace(PROJ_RE, '').replace(/\s+/g, ' ').trim() || title,
    proj: pm ? pm[1] : '', owner: owner, cat: cat === '未分類' ? '' : cat,
    start: ymd_(s), due: ymd_(e), plan: plan, memo: desc.slice(0, 300),
    done: /^\s*(✓|✔|☑|✅|【完了】|\[完了\]|済[ 　:：])/.test(title)
  };
}

function installTaskTriggers_(cals) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncTaskCalendars').forEach(t => ScriptApp.deleteTrigger(t));
  if (!cals.length) return '';
  const ng = [];
  cals.forEach(c => {
    try { ScriptApp.newTrigger('syncTaskCalendars').forUserCalendar(c).onEventUpdated().create(); } catch (e) { ng.push(c); }
  });
  if (ng.length) {
    ScriptApp.newTrigger('syncTaskCalendars').timeBased().everyMinutes(15).create();
    return 'タスク用カレンダーは15分ごとに同期します。';
  }
  return 'タスク用カレンダーに予定が追加・変更されると、自動でタスクに反映します。';
}

function saveTaskCalendars(ids) {
  const cfg = readConfig_();
  assertAdmin_(cfg);
  const list = splitList_(ids);
  list.forEach(c => { if (c.indexOf('@') < 0) throw new Error('カレンダーIDの形式が正しくありません: ' + c); });
  cfg.sheet.getRange('L14').setValue(list.join(','));
  SpreadsheetApp.flush();
  const trig = installTaskTriggers_(list);
  const msg = list.length ? syncTaskCalendars() : 'タスク用カレンダーの連携を解除しました';
  return { msg: msg + (trig ? '\n' + trig : ''), data: getAdminData() };
}

function runTaskSync() {
  assertAdmin_(readConfig_());
  return { msg: syncTaskCalendars(), data: getAdminData() };
}

function ensureConfigRows_(sh) {
  if (String(sh.getRange('K14').getValue()).trim()) return;
  sh.getRange('K14:K15').setValues([['タスク用カレンダーID（カンマ区切り）'], ['最終タスク同期']]);
  box_(sh.getRange('K14:K15'), CLR.sub).setFontWeight('bold');
  box_(sh.getRange('L14'), CLR.yel);
  box_(sh.getRange('L15'), CLR.gry);
}

function ensureTaskColumns_(sh) {
  const head = sh.getRange(1, 1, 1, TASK_HEAD.length).getValues()[0];
  [[TC.calEv, 160, 'タスク用カレンダーから自動追加されたタスクの予定ID（自動入力・編集しない）'],
   [TC.parent, 85, '上位のタスクID。業務分掌の小タスク・中タスクにぶら下げると、実績が上位にも合算されます'],
   [TC.level, 50, '大／中／小＝業務分掌から作った定常業務。空欄＝個別タスク']].forEach(x => {
    if (String(head[x[0] - 1]).trim()) return;
    head_(sh.getRange(1, x[0]), [TASK_HEAD[x[0] - 1]]);
    sh.setColumnWidth(x[0], x[1]);
    sh.getRange(1, x[0]).setNote(x[2]);
  });
  const lh = sheet_(SH.LOG).getRange(1, LOG_HEAD.length);
  if (!String(lh.getValue()).trim()) {
    head_(lh, [LOG_HEAD[LOG_HEAD.length - 1]]);
    lh.setNote('このログのタスクの上位（中・大タスク）のID。上位タスクの実績合算に使います');
  }
}

// 既存シートの数式を最新版に書き換える（親子合算・定常業務の判定を反映）
function refreshTaskFormulas_(sh) {
  const n = Math.max(sh.getLastRow(), TASK_ROWS_INIT + 1) - 1;
  const f1 = [], f2 = [];
  for (let r = 2; r < 2 + n; r++) { const f = taskFormulas_(r); f1.push([f.dept]); f2.push(f.tail); }
  sh.getRange(2, TC.dept, n, 1).setFormulas(f1);
  sh.getRange(2, TC.actual, n, 5).setFormulas(f2);
}

function ancestors_(taskMap, id) {
  const out = [];
  let t = id ? taskMap[id] : null, guard = 0;
  while (t && t.parent && guard++ < 10) { out.push(t.parent); t = taskMap[t.parent]; }
  return out.length ? ',' + out.join(',') + ',' : '';
}

function stripHtml_(s) {
  return String(s).replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').trim();
}

// ================= 業務分掌 → 大・中・小タスク =================
// 出典：営業部_業務分掌_260602.xlsx（営業課の業務。資材購買課・計画業務の一覧と「廃止」は除外）
// [大タスク, 中タスク, 小タスク, 担当区分, 月間工数{氏名:h}, 新規追加見込{客先:h}, 購買課h, 資材購買課が担当中(1/0), 具体的作業[], 補足]
const GYOMU_SRC = '営業部_業務分掌_260602';
const GYOMU = [
  ["製造側：客先対応","客先窓口","客先定期訪問（製造側）・資料作成・議事録作成含む","営業",{"野村":12.0,"赤羽":20.0,"永浜":4.0},{"コナミ":12.0,"エキサイト":12.0,"ニューギン":12.0},0,0,[],""],
  ["製造側：客先対応","客先窓口","客先からの連絡1ST対応～旗振り","営業",{"野村":4.0,"赤羽":12.0,"永浜":12.0},{"コナミ":6.0,"エキサイト":6.0,"ニューギン":6.0},0,0,[],""],
  ["製造側：客先対応","客先窓口","4M変更、EOL対応（藤商事様に対して）","営業",{"丹羽":1.0,"赤羽":8.0,"永浜":8.0},{"コナミ":1.0,"エキサイト":1.0,"ニューギン":1.0},0,0,[],""],
  ["製造側：客先対応","客先窓口","客先固定資産の確認","営業",{"赤羽":1.0,"永浜":1.0},{},5.0,1,[],"写真等の入手は、現場。指定フォーマットへの落とし込みは営業事務。"],
  ["製造側：客先対応","客先窓口","営業情報入手・社内展開（フォーキャスト・製造計画）","営業",{"赤羽":5.0,"永浜":5.0},{"コナミ":5.0,"エキサイト":5.0,"ニューギン":5.0},5.0,1,["販売フォーキャスト情報","生産計画","部材発注フォーキャスト情報"],""],
  ["製造側：客先対応","客先窓口","構成表の入手依頼、編集依頼と社内展開","営業",{"赤羽":4.0,"永浜":3.0},{"コナミ":4.0,"エキサイト":4.0,"ニューギン":4.0},4.0,1,[],""],
  ["製造側：客先対応","客先窓口","支給品の支給日の確認依頼、引き取り","資材購買",{"赤羽":2.0,"永浜":5.0},{"コナミ":2.0,"エキサイト":2.0,"ニューギン":2.0},8.0,1,[],""],
  ["製造側：客先対応","客先窓口","各種ビジネス文書の作成（案内・交渉・謝罪・通知）","営業、営業事務",{"丹羽":4.0,"野村":1.0,"赤羽":9.0,"永浜":8.0},{"コナミ":6.0,"エキサイト":6.0,"ニューギン":6.0},0,0,[],""],
  ["製造側：客先対応","見積もり業務","量産見積り（客先へ提示する、判断）・交渉","営業",{"丹羽":1.0,"野村":1.0,"赤羽":14.0,"永浜":5.0},{"コナミ":8.0,"エキサイト":8.0,"ニューギン":8.0},0,0,["藤商事向け見積書作成"],""],
  ["製造側：客先対応","見積もり業務","設計見積り（客先へ提示する、判断）・交渉","営業",{"丹羽":1.0,"野村":2.0,"赤羽":12.0,"永浜":5.0},{"コナミ":5.0,"エキサイト":5.0,"ニューギン":5.0},0,0,[],""],
  ["製造側：客先対応","見積もり業務","契約書の作成/確認/リーガルチェック/交渉　（製造側）","営業",{"野村":1.0},{"コナミ":1.0,"エキサイト":1.0,"ニューギン":1.0},0,0,[],""],
  ["製造側：客先対応","売上業務","受注管理（注文書の確認や注文催促、ミス修正依頼）","営業事務",{"野村":2.0,"赤羽":5.0,"永浜":4.0},{"コナミ":4.0,"エキサイト":4.0,"ニューギン":4.0},0,1,["修正依頼の試作で開発への依頼は営業で引き取ります"],""],
  ["製造側：客先対応","売上業務","売上管理（売上処理）←F社製造分のみ","営業事務",{"永浜":4.0},{"コナミ":1.0,"エキサイト":1.0,"ニューギン":1.0},0,1,[],"K10売上入力、受領書と販売管理（出荷金額）の照合"],
  ["製造側：客先対応","売上業務","請求確認（藤商事様）","営業事務",{"赤羽":1.0,"永浜":2.0},{"コナミ":1.0,"エキサイト":1.0,"ニューギン":1.0},0,1,[],"客先に発行する請求確認書"],
  ["製造側：客先対応","売上業務","客先出荷調整","営業",{"永浜":4.0},{"コナミ":2.0,"エキサイト":2.0,"ニューギン":2.0},0,1,[],"社内の出荷調整は、計画業務だが、客先窓口対応は営業。"],
  ["製造側：客先対応","売上業務","客先出荷伝票発行、検収依頼","営業事務",{"永浜":8.0},{"コナミ":0.5,"エキサイト":0.5,"ニューギン":0.5},0,1,[],"仮伝票発行、客先提出。（あと追い伝票発行時）"],
  ["製造側：客先対応","売上業務","客先検収確認","営業事務",{"赤羽":2.0,"永浜":2.0},{"コナミ":0.25,"エキサイト":0.25,"ニューギン":0.25},0,1,[],""],
  ["製造側：客先対応","売上業務","入金実績確認","営業事務",{"赤羽":1.0,"永浜":2.0},{"コナミ":0.25,"エキサイト":0.25,"ニューギン":0.25},0,1,[],""],
  ["製造側：客先対応","支払業務","支払い資料作成","営業",{"赤羽":1.0},{},0,0,[],""],
  ["製造側：客先対応","販路拡大業務","販路開拓・関係継続（製造側）","営業",{"丹羽":12.0,"野村":6.0,"赤羽":5.0,"永浜":5.0},{},0,0,[],""],
  ["製造側：客先対応","販路拡大業務","商材セールス(製品）","営業",{"丹羽":12.0,"野村":6.0,"赤羽":6.0,"永浜":6.0},{},0,0,[],""],
  ["製造側：客先対応","販路拡大業務","資料準備（販促系）","営業",{"野村":8.0,"早瀬":60.0},{},0,0,["DAVINCIパンフレット","商材ﾁﾗｼ","会社紹介","打合せ・提案用資料準備"],""],
  ["製造側：社内対応","利益計画管理","製造予算立案","営業",{"丹羽":3.0,"野村":1.0,"赤羽":6.0},{},0,0,[],""],
  ["製造側：社内対応","利益計画管理","利益計画（月毎の確認･修正・実績入力・達成のための呼びかけ）","営業",{"丹羽":8.0,"赤羽":10.0},{},0,0,[],""],
  ["製造側：社内対応","利益計画管理","製造側戦略立案/シミュレーション","営業",{"丹羽":2.0,"野村":10.0,"赤羽":10.0},{},0,0,[],""],
  ["製造側：社内対応","利益計画管理","事業別経営会議資料作成","営業",{"丹羽":8.0,"野村":4.0,"赤羽":8.0},{},0,0,[],""],
  ["製造側：社内対応","マルチ","社内旗振り（受注⇒社内管理⇒成果物納品⇒請求⇒売上）\n※新案件に対して。プロデュース業務ではない。⇒お客様のご要望を社内リソースを使って実現し、成果物を納品する。","営業",{"赤羽":10.0,"早瀬":10.0,"永浜":10.0},{},0,0,[],""],
  ["製造側：社内対応","マルチ","出張（工場見学・展示会）","営業",{"野村":4.0,"赤羽":2.0},{},0,0,[],""],
  ["製造側：社内対応","マルチ","全防連","営業",{"赤羽":2.0},{},0,0,[],""],
  ["製造側：社内対応","マルチ","古物商","営業",{"赤羽":1.0},{},0,0,[],""],
  ["製造側：社内対応","社内打合せ・その他","来客（商社・ベンダー）","営業",{"野村":2.0,"赤羽":1.0},{},0,0,[],""],
  ["製造側：社内対応","社内打合せ・その他","製造側社内打ち合わせ","全般",{"野村":32.0,"赤羽":36.0,"永浜":24.0},{},0,0,[],""],
  ["製造側：社内対応","社内打合せ・その他","初品出荷立ち合い/部材納品・移動","資材購買",{"早瀬":2.0,"永浜":10.0},{},0,0,["見本機/量産 初回納品","試作納品/受取","注残部材納品"],""],
  ["ソフト側：客先対応","客先窓口","客先定期訪問（ソフト側）・資料作成・議事録作成含む","営業",{"丹羽":8.0,"野村":6.0,"早瀬":8.0},{},0,0,[],""],
  ["ソフト側：客先対応","客先窓口","客先からの連絡1ST対応～旗振り","営業",{"丹羽":8.0,"野村":10.0,"早瀬":8.0},{},0,0,[],""],
  ["ソフト側：客先対応","客先窓口","営業情報入手・社内展開（機種情報・業界情報）","営業",{"丹羽":16.0,"野村":10.0,"早瀬":10.0},{},0,0,["藤商事ぱちんこラインナップ確認","藤商事パチスロラインナップ確認"],""],
  ["ソフト側：客先対応","見積もり業務","開発見積り（客先へ提示する、判断）・交渉（打合せ・電話・メール）","営業",{"丹羽":2.0,"野村":22.0,"早瀬":12.0},{},0,0,["藤商事向け見積書作成","見積交渉"],""],
  ["ソフト側：客先対応","見積もり業務","契約書の作成/確認/リーガルチェック/交渉　（開発側）","営業",{"丹羽":3.0,"野村":4.0,"早瀬":10.0},{},0,0,["フィールズ様保守契約"],""],
  ["ソフト側：客先対応","見積もり業務","契約業務","営業、営業事務",{"丹羽":3.0,"野村":4.0,"早瀬":1.0,"田中沙":5.0},{},0,0,["契約申請","管理表反映","契約書保管(Redmine)","契約書保管(フォルダ)","契約書保管(総務提出分)","ドキュサイン"],""],
  ["ソフト側：客先対応","売上業務","日報コード作成","営業事務",{"野村":0.5,"田中沙":3.0},{},0,0,["コード取得","勤次郎コード登録","総務へ紐づけ依頼","管理表3種・フォルダに反映"],""],
  ["ソフト側：客先対応","売上業務","販売管理（工数集計・社外開発費管理）","営業事務",{"野村":8.0,"早瀬":5.0,"田中沙":10.0},{},0,0,["工数集計（人別も）","経理工数報告書の作成・回覧"],""],
  ["ソフト側：客先対応","売上業務","売上管理（売上処理）〔F社・D社・FIS社・開発製品・回路設計・業務保守・ロイヤリティ・アプリ・ゲーム〕","営業事務",{"丹羽":16.0,"野村":4.0,"田中沙":35.0},{},0,0,["納品書・請求書の発行回覧・フォルダ保管","K10売上入力","納品書請求書・納品物の発送","納品物の準備依頼","販売台数報告の押印回覧・売上処理"],""],
  ["ソフト側：客先対応","売上業務","客先検収確認","営業事務",{"野村":0.5,"早瀬":1.0,"田中沙":1.0},{},0,0,["受領書の押印依頼","受領書フォルダ保管"],""],
  ["ソフト側：客先対応","売上業務","入金実績確認","営業事務",{"野村":1.0,"田中沙":1.0},{},0,0,["F社D社フィールズ入金額照合","承認回覧・保管"],""],
  ["ソフト側：客先対応","支払業務","支払い資料作成","営業",{"丹羽":8.0},{},0,0,[],""],
  ["ソフト側：客先対応","支払業務","支払依頼（外注費）","営業事務",{"丹羽":1.0,"野村":1.0,"早瀬":1.0,"田中沙":3.0},{},0,0,["外注費支払処理"],""],
  ["ソフト側：客先対応","販路拡大業務","販路開拓・関係継続（開発側）","営業",{"丹羽":16.0},{},0,0,[],""],
  ["ソフト側：客先対応","販路拡大業務","商材セールス","営業",{"丹羽":4.0},{},0,0,[],""],
  ["ソフト側：客先対応","販路拡大業務","戦略立案/シミュレーション","営業",{"丹羽":4.0,"野村":2.0},{},0,0,[],""],
  ["ソフト側：社内対応","利益計画管理","開発予算立案（ソフト）","営業",{"丹羽":3.0,"野村":12.0},{},0,0,[],""],
  ["ソフト側：社内対応","利益計画管理","利益計画（月毎の確認･修正・実績入力・達成のための呼びかけ）","営業",{"丹羽":8.0,"野村":20.0,"早瀬":5.0},{},0,0,["社内資料の整備含む"],""],
  ["ソフト側：社内対応","利益計画管理","事業別経営会議資料作成","営業",{"丹羽":8.0,"野村":2.0},{},0,0,[],""],
  ["ソフト側：社内対応","利益計画管理","社外発注用稟議起案＆却下判断・見積もり精査","営業",{"丹羽":8.0,"野村":4.0},{},0,0,[],""],
  ["ソフト側：社内対応","利益計画管理","社外開発検収管理（検査結果報告書など）","営業事務",{"丹羽":2.0,"野村":0.5,"田中沙":10.0},{},0,0,["外注検査結果送付","検査結果の保管","支払処理","検収と請求の進捗確認"],""],
  ["ソフト側：社内対応","社内打合せ・その他","監査","営業、営業事務",{"野村":0.5,"田中沙":1.0},{},0,0,[],""],
  ["ソフト側：社内対応","社内打合せ・その他","外注打合せ（来客・Web）","営業",{"田中沙":5.0},{},0,0,[],""],
  ["ソフト側：社内対応","社内打合せ・その他","社内打ち合わせ","営業",{"丹羽":4.0,"野村":4.0,"早瀬":4.0,"田中沙":5.0},{},0,0,[],""],
  ["ソフト側：社内対応","社内打合せ・その他","開発の業務の手伝い","営業",{"野村":2.0,"早瀬":20.0},{},0,0,[],""],
  ["出張対応","出張","東京などへの出張","営業",{"丹羽":12.0,"野村":12.0,"早瀬":12.0},{},0,0,[],""],
  ["事業部内共通業務","総務・事務","クレカ対応","",{"田中沙":3.0},{},0,0,["購入サイト登録","購入直前画面押印依頼","購入手続き","購入確定画面押印依頼","適格請求書ダウンロード","PDF結合フォルダ保管","支払処理の添付資料依頼とチェック"],""],
  ["事業部内共通業務","総務・事務","社用車管理","",{"田中沙":3.0},{},0,0,["車両選定","稟議申請","契約申請","契約書類のやり取り","タイヤ交換","点検対応","修理・支払対応"],"運転日報補充"],
  ["事業部内共通業務","総務・事務","消耗品","",{"田中沙":2.0},{},0,0,["在庫チェック","稟議と購入","補充・配布"],""],
  ["事業部内共通業務","総務・事務","支払処理","",{"田中沙":3.0},{},0,0,["毎月処理分","依頼分"],""],
  ["事業部内共通業務","総務・事務","その他事務","営業事務",{"永浜":7.0,"田中沙":10.0},{},0,0,["購読本管理","管理シール発行","荷物発送","送り状補充","郵便物配布","代理押印と回覧","契約書保管"],""],
  ["事業部内共通業務","アプリ・ゲーム・ロイヤリティ","アプリ（Apple/Google)国内外4種","営業事務",{"田中沙":3.0},{},0,0,["データ保存","手数料相殺処理","K10売上処理","入金自実績確認"],""],
  ["事業部内共通業務","アプリ・ゲーム・ロイヤリティ","アプリ四半期（１社は毎月対応）","営業事務",{"田中沙":10.0},{},0,0,["PDF結合","印章申請","実績報告書の送付(13社)","支払処理(支払発生全て）","客先メールのやり取り"],""],
  ["事業部内共通業務","アプリ・ゲーム・ロイヤリティ","ゲーム（各国分7件）","営業、営業事務",{"田中沙":5.0},{},0,0,["データ保存","相殺処理","売上入力・回覧","分配金の支払処理"],""],
  ["事業部内共通業務","知財","＋１知財活動","開発",{"丹羽":1.0,"田中沙":5.0},{},0,0,["活動内容検討と案内","アイデア選別","客先ヒアリング対応","明細書確認依頼","譲渡確認書処理","発明届出書作成","知財アイデア説明会"],"特許補償金の売上・請求書発行"],
  ["事業部内共通業務","営業システム","営業用システム構築・ルール検討","営業、営業事務",{"野村":4.0,"早瀬":10.0,"永浜":42.0},{},0,0,["予実管理修正","楽楽販売構築・修正/メンテナンス","見積書/注文書管理システム"],""],
  ["事業部内共通業務","総務・事務","固定資産の確認（社内）","",{"田中沙":2.0},{},0,0,[],""],
  ["事業部内共通業務","総務・事務","固定資産の確認（協力企業）","",{},{},0,0,[],""]
];
const GYOMU_OTHER = ['その他', 'その他', 'その他（上記に当てはまらない業務）', '', {}, {}, 0, 0, [], '業務分掌のどれにも当てはまらない仕事はここに付けてください'];
// 中タスク＝集計カテゴリ。並び順＝キーワード判定の優先順
const GYOMU_CATS = [
  ['客先窓口', '訪問,客先,問合せ,問い合わせ,議事録,4M,EOL,構成表,フォーキャスト,支給品', '×'],
  ['見積もり業務', '見積,契約,リーガル,ドキュサイン', '×'],
  ['売上業務', '受注,注文,売上,請求,出荷,検収,入金,販売管理,納品,日報コード', '○'],
  ['支払業務', '支払,外注費', '×'],
  ['販路拡大業務', '販路,セールス,商談,販促,パンフ,展示会,提案', '×'],
  ['利益計画管理', '予算,利益計画,経営会議,戦略,稟議,シミュレーション', '×'],
  ['マルチ', '旗振り,全防連,古物商', '×'],
  ['社内打合せ・その他', '打合せ,打ち合わせ,会議,MTG,ミーティング,定例,来客,監査', '×'],
  ['出張', '出張', '×'],
  ['総務・事務', 'クレカ,社用車,消耗品,郵便,押印,固定資産,発送', '×'],
  ['アプリ・ゲーム・ロイヤリティ', 'アプリ,ゲーム,ロイヤリティ,Apple', '×'],
  ['知財', '知財,特許,発明', '×'],
  ['営業システム', 'システム,楽楽販売,Redmine', '×'],
  ['その他', '', '×']
];

function importGyomu() {
  const ui = SpreadsheetApp.getUi();
  const ans = ui.alert('業務分掌を大・中・小タスクに反映',
    '「' + GYOMU_SRC + '」の営業課の業務 ' + GYOMU.length + ' 件を、大タスク → 中タスク → 小タスクとして「タスク」シートに登録します。' +
    '\n（2回目以降は予定工数・メモだけ更新し、重複登録はしません）' +
    '\n\n設定シートのカテゴリも「中タスク」（客先窓口・見積もり業務・売上業務…）に置き換えますか？' +
    '\n　はい＝置き換える（集計・移管シミュレーションが中タスク別になります）' +
    '\n　いいえ＝カテゴリはそのまま', ui.ButtonSet.YES_NO_CANCEL);
  if (ans !== ui.Button.YES && ans !== ui.Button.NO) return;
  const msg = withLock_(() => applyGyomu_(ans === ui.Button.YES));
  ui.alert(msg);
}

function applyGyomu_(replaceCats) {
  const sh = sheet_(SH.TASK);
  ensureTaskColumns_(sh);
  refreshTaskFormulas_(sh);
  const cfg = readConfig_();
  if (replaceCats) writeGyomuCats_(cfg.sheet);
  const names = cfg.allMembers.map(m => m.name);
  const ownerOf = p => names.find(n => n.replace(/\s|　/g, '').indexOf(p) === 0) || p;

  // 木構造に組み立て（大 → 中 → 小）
  const tree = [];
  GYOMU.concat([GYOMU_OTHER]).forEach(g => {
    let b = tree.find(x => x.name === g[0]);
    if (!b) tree.push(b = { name: g[0], mids: [] });
    let m = b.mids.find(x => x.name === g[1]);
    if (!m) b.mids.push(m = { name: g[1], items: [] });
    m.items.push(g);
  });
  const planOf = g => Object.keys(g[4]).reduce((s, k) => s + g[4][k], 0);
  const memoOf = g => {
    const fmt = o => Object.keys(o).map(k => k + ' ' + o[k] + 'h').join('・');
    return [g[3] && '担当区分：' + g[3], g[7] && '【資材購買課が担当中】',
      Object.keys(g[4]).length && '月間工数：' + fmt(g[4]), Object.keys(g[5]).length && '新規追加見込（月）：' + fmt(g[5]),
      g[6] && '購買課：' + g[6] + 'h／月', g[8].length && '具体的作業：' + g[8].join('／'), g[9] && '補足：' + g[9]]
      .filter(Boolean).join('\n');
  };

  const exist = {};
  readTasks_().forEach(t => { if (t.level) exist[t.level + '|' + t.parent + '|' + t.name] = t; });
  const n0 = Math.max(sh.getLastRow() - 1, 0);
  const v = n0 ? sh.getRange(2, 1, n0, 3).getValues() : [];
  let maxNo = maxTaskNo_(v.map(r => r[0])), lastUsed = -1;
  v.forEach((r, i) => { if (r[0] || String(r[2]).trim()) lastUsed = i; });
  const startRow = lastUsed + 3;
  const newRows = [];
  let updated = 0;
  const put = (lv, parent, name, f) => {
    const t = exist[lv + '|' + parent + '|' + name];
    if (t) {
      writeTaskFields_(sh, findTaskRow_(sh, t.id), { plan: f.plan, cat: f.cat, memo: f.memo });
      updated++;
      return t.id;
    }
    const id = fmtId_(++maxNo), r = startRow + newRows.length, fm = taskFormulas_(r);
    const row = new Array(TASK_HEAD.length).fill('');
    row[TC.id - 1] = id; row[TC.name - 1] = name; row[TC.owner - 1] = f.owner || ''; row[TC.dept - 1] = fm.dept;
    row[TC.cat - 1] = f.cat || ''; row[TC.status - 1] = '進行中'; row[TC.prio - 1] = '中'; row[TC.plan - 1] = f.plan;
    fm.tail.forEach((x, i) => { row[TC.actual - 1 + i] = x; });
    row[TC.memo - 1] = f.memo || ''; row[TC.parent - 1] = parent; row[TC.level - 1] = lv;
    newRows.push(row);
    return id;
  };
  let nb = 0, nm = 0, ns = 0;
  tree.forEach(b => {
    const bPlan = b.mids.reduce((s, m) => s + m.items.reduce((t, g) => t + planOf(g), 0), 0);
    const bid = put('大', '', b.name, { plan: bPlan, cat: '', memo: '業務分掌（' + GYOMU_SRC + '）の大項目。予定工数＝配下の月間工数の合計' });
    nb++;
    b.mids.forEach(m => {
      const mPlan = m.items.reduce((t, g) => t + planOf(g), 0);
      const mid = put('中', bid, m.name, { plan: mPlan, cat: m.name, memo: '予定工数＝配下の月間工数の合計' });
      nm++;
      m.items.forEach(g => {
        const who = Object.keys(g[4]).filter(k => g[4][k] > 0);
        put('小', mid, g[2], { plan: planOf(g), cat: m.name, owner: who.length === 1 ? ownerOf(who[0]) : '', memo: memoOf(g) });
        ns++;
      });
    });
  });
  if (newRows.length) {
    const rg = sh.getRange(startRow, 1, newRows.length, TASK_HEAD.length);
    rg.setValues(newRows);
    box_(rg);
    sh.getRange(startRow, TC.plan, newRows.length, 2).setNumberFormat('0.0');
    sh.getRange(startRow, TC.rate, newRows.length, 1).setNumberFormat('0%');
    sh.getRange(startRow, TC.remain, newRows.length, 1).setNumberFormat('0.0');
    sh.getRange(startRow, TC.last, newRows.length, 1).setNumberFormat('yyyy/mm/dd');
    sh.getRange(startRow, TC.memo, newRows.length, 1).setWrap(false);
    newRows.forEach((r, i) => {
      const bg = r[TC.level - 1] === '大' ? '#1f3864' : r[TC.level - 1] === '中' ? '#d9e1f2' : null;
      if (bg) sh.getRange(startRow + i, TC.id, 1, TC.name).setBackground(bg).setFontColor(bg === '#1f3864' ? '#ffffff' : '#000000').setFontWeight('bold');
    });
  }
  sh.getRange(2, TC.level, Math.max(sh.getLastRow() - 1, 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(LEVELS, true).setAllowInvalid(true).build());
  return '業務分掌を反映しました：大 ' + nb + '・中 ' + nm + '・小 ' + ns + ' 件（新規 ' + newRows.length + '件／更新 ' + updated + '件）' +
    (replaceCats ? '\nカテゴリを中タスク（' + GYOMU_CATS.length + '種）に置き換えました。' : '') +
    '\n\nカレンダーの予定を小タスクに付けると、中・大タスクにも実績が合算されます（Web画面「カレンダー」タブ、または件名にタスクID）。';
}

function writeGyomuCats_(conf) {
  const rows = [];
  for (let i = 0; i < CAT_ROWS; i++) rows.push(GYOMU_CATS[i] || ['', '', '']);
  conf.getRange(5, 7, CAT_ROWS, 3).setValues(rows);
}

// ================= トリガー =================
function installTrigger() {
  removeTrigger(true);
  ScriptApp.newTrigger('importRecent').timeBased().everyDays(1).atHour(6).create();
  saveSsId_();
  const t = installTaskTriggers_(readConfig_().taskCals);
  SpreadsheetApp.getUi().alert('毎朝6時台にカレンダーを自動取り込みします。' + (t ? '\n' + t : ''));
}

function removeTrigger(silent) {
  ScriptApp.getProjectTriggers().filter(t => ['importRecent', 'syncTaskCalendars'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  if (silent !== true) SpreadsheetApp.getUi().alert('自動取り込みをOFFにしました。');
}

function showWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    ui.alert('まだWebアプリとして公開されていません。\nApps Scriptエディタ →「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」で公開してください。');
    return;
  }
  const html = HtmlService.createHtmlOutput('<p style="font-family:sans-serif;font-size:13px">次のURLをブックマーク／チームに共有してください。</p><p style="font-family:sans-serif;font-size:13px;word-break:break-all"><a href="' + url + '" target="_blank">' + url + '</a></p>').setWidth(520).setHeight(150);
  ui.showModalDialog(html, 'ボード／ガント画面');
}

// ================= 初期設定 =================
function setup() {
  const ss = SpreadsheetApp.getActive();
  saveSsId_();
  const builders = [[SH.CONF, buildConfig_], [SH.TASK, buildTasks_], [SH.LOG, buildLog_], [SH.SUM, buildSummary_], [SH.SIM, buildSim_], [SH.DAILY, buildDaily_]];
  const made = [], skipped = [];
  builders.forEach(b => { if (ss.getSheetByName(b[0])) skipped.push(b[0]); else { ss.insertSheet(b[0], ss.getSheets().length); made.push(b[0]); } });
  builders.forEach(b => { if (made.indexOf(b[0]) >= 0) b[1](ss.getSheetByName(b[0])); });
  ['シート1', 'Sheet1'].forEach(n => { const d = ss.getSheetByName(n); if (d && d.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(d); });
  ss.setActiveSheet(ss.getSheetByName(SH.CONF));
  SpreadsheetApp.getUi().alert('初期設定',
    '作成: ' + (made.join('、') || 'なし') + (skipped.length ? '\n既存のためスキップ: ' + skipped.join('、') : '') +
    '\n\n次に「設定」シートでメンバーとカテゴリを入力し、メニュー「カレンダーから取り込む」を実行してください。',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function buildConfig_(sh) {
  title_(sh, '設定', '黄色セルを入力。メンバーには自分のカレンダーを、このシートの所有者へ「予定の変更」（閲覧だけなら「予定の詳細をすべて表示」）で共有してもらってください。Web画面の「管理」タブからも登録できます。');
  head_(sh.getRange('A4:E4'), ['氏名', '部署', 'カレンダーID（メールアドレス）', '取込', '最終取込結果']);
  box_(sh.getRange(5, 1, MEMBER_ROWS, 3), CLR.yel);
  box_(sh.getRange(5, 4, MEMBER_ROWS, 1), CLR.yel).insertCheckboxes().setHorizontalAlignment('center');
  box_(sh.getRange(5, 5, MEMBER_ROWS, 1), CLR.gry).setFontSize(9);
  let me = '';
  try { me = Session.getActiveUser().getEmail(); } catch (e) {}
  sh.getRange('A5:D5').setValues([['（自分の名前）', '営業課', me, true]]);
  sh.getRange(5, 2, MEMBER_ROWS, 1).setDataValidation(rangeRule_(sh.getRange('L8:L9')));

  head_(sh.getRange('G4:I4'), ['カテゴリ', '件名キーワード（カンマ区切り）', '移管候補']);
  const cats = [
    ['見積もり取得', '見積', '○'], ['仕様確認', '仕様,図面', '×'], ['納期調整', '納期,督促', '○'],
    ['発注処理', '発注,注文書', '×'], ['価格交渉', '価格,単価,値引', '○'], ['仕入先選定・評価', '仕入先,ベンダー', '×'],
    ['顧客対応・問合せ', '問合せ,問い合わせ,客先,来社,訪問', '×'], ['在庫・入荷確認', '在庫,入荷', '×'],
    ['請求・支払処理', '請求,支払', '×'], ['資料作成', '資料,提案書,報告書', '×'],
    ['社内会議', '会議,定例,MTG,ミーティング,打合せ,打ち合わせ', '×'], ['その他', '', '×']
  ];
  const rows = [];
  for (let i = 0; i < CAT_ROWS; i++) rows.push(cats[i] || ['', '', '']);
  box_(sh.getRange(5, 7, CAT_ROWS, 3), CLR.yel).setValues(rows);
  sh.getRange(5, 9, CAT_ROWS, 1).setDataValidation(listRule_(['○', '×'])).setHorizontalAlignment('center');

  head_(sh.getRange('K4:L4'), ['項目', '値']);
  const today = startOfDay_(new Date());
  const st = [['取込開始日', addDays_(today, -28)], ['再取込する直近日数', 14], ['除外キーワード', '休憩,昼食,ランチ,私用,通院,休暇,有給,移動'],
    ['部署1（移管元）', '営業課'], ['部署2（移管先）', '資材購買課'], ['月の営業日数', 20], ['人件費単価（円/時間）', 3000], ['最終取込', ''], ['管理者メール（カンマ区切り）', me], ['タスク用カレンダーID（カンマ区切り）', ''], ['最終タスク同期', '']];
  sh.getRange(5, 11, st.length, 2).setValues(st);
  box_(sh.getRange(5, 11, st.length, 1), CLR.sub).setFontWeight('bold');
  box_(sh.getRange(5, 12, st.length, 1), CLR.yel).setHorizontalAlignment('left');
  box_(sh.getRange(12, 12), CLR.gry);
  box_(sh.getRange(15, 12), CLR.gry);
  sh.getRange('L5').setNumberFormat('yyyy/mm/dd');

  const guide = [
    '■ カレンダー予定の書き方（どれか1つでOK。上ほど優先）',
    '1) タスクIDを件名か説明に入れる　例「T-0012 A社見積作成」→ タスクの案件・カテゴリで自動集計',
    '2) 件名に【カテゴリ】か #カテゴリ　例「【納期調整】△△部品の督促」',
    '3) 何もしなくても、件名にキーワード（見積・納期…）があれば自動判定。判定できない予定は「未分類」',
    '・案件は「@案件名」で付けられます　例「【見積もり取得】@A社更新 仕様確認」',
    '・終日予定／辞退した予定／不在・勤務場所の予定／除外キーワードを含む予定は取り込みません',
    '・Web画面「カレンダー」タブで、件名を変えずにカテゴリ・タスクを付け直せます（最優先で反映）'
  ];
  guide.forEach((g, i) => sh.getRange(22 + i, 7).setValue(g).setFontSize(i ? 9 : 10).setFontWeight(i ? 'normal' : 'bold').setFontColor(i ? '#333333' : CLR.head));
  [130, 100, 240, 45, 220, 16, 130, 280, 70, 16, 150, 260].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(4);
}

function buildTasks_(sh) {
  const conf = sheet_(SH.CONF);
  const N = TASK_ROWS_INIT;
  head_(sh.getRange(1, 1, 1, TASK_HEAD.length), TASK_HEAD);
  sh.setFrozenRows(1); sh.setFrozenColumns(3);
  box_(sh.getRange(2, 1, N, TASK_HEAD.length));
  [TC.proj, TC.name, TC.owner, TC.cat, TC.status, TC.prio, TC.start, TC.due, TC.plan, TC.memo].forEach(c => sh.getRange(2, c, N, 1).setBackground(CLR.yel));
  [TC.id, TC.dept, TC.actual, TC.rate, TC.remain, TC.judge, TC.last].forEach(c => sh.getRange(2, c, N, 1).setBackground(CLR.gry));
  const f1 = [], f2 = [];
  for (let r = 2; r < 2 + N; r++) { const f = taskFormulas_(r); f1.push([f.dept]); f2.push(f.tail); }
  sh.getRange(2, TC.dept, N, 1).setFormulas(f1);
  sh.getRange(2, TC.actual, N, 5).setFormulas(f2);
  sh.getRange(2, TC.owner, N, 1).setDataValidation(rangeRule_(conf.getRange('A5:A34')));
  sh.getRange(2, TC.cat, N, 1).setDataValidation(rangeRule_(conf.getRange('G5:G19')));
  sh.getRange(2, TC.status, N, 1).setDataValidation(listRule_(STATUSES));
  sh.getRange(2, TC.prio, N, 1).setDataValidation(listRule_(['高', '中', '低']));
  sh.getRange(2, TC.start, N, 2).setNumberFormat('yyyy/mm/dd');
  sh.getRange(2, TC.last, N, 1).setNumberFormat('yyyy/mm/dd');
  sh.getRange(2, TC.plan, N, 2).setNumberFormat('0.0');
  sh.getRange(2, TC.rate, N, 1).setNumberFormat('0%');
  sh.getRange(2, TC.remain, N, 1).setNumberFormat('0.0');
  sh.getRange(2, TC.id, N, 1).setHorizontalAlignment('center');

  const t = startOfDay_(new Date());
  sh.getRange(2, 1, 3, 4).setValues([
    ['T-0001', 'A社 更新案件', 'A社向け見積作成', '（自分の名前）'],
    ['T-0002', 'A社 更新案件', '部品の納期確認', '（自分の名前）'],
    ['T-0003', '社内', '月次報告資料の作成', '（自分の名前）']]);
  sh.getRange(2, TC.cat, 3, 6).setValues([
    ['見積もり取得', '進行中', '高', t, addDays_(t, 7), 6],
    ['納期調整', '未着手', '中', addDays_(t, 3), addDays_(t, 10), 3],
    ['資料作成', '未着手', '低', addDays_(t, 5), addDays_(t, 14), 4]]);

  const jr = sh.getRange(2, TC.judge, N, 1);
  sh.setConditionalFormatRules([['期限超過', '#ffc7ce', '#9c0006'], ['工数超過', '#ffc7ce', '#9c0006'], ['注意', '#ffeb9c', '#9c5700'], ['完了', '#e7e6e6', '#7f7f7f'], ['定常業務', '#ddebf7', '#1f3864']]
    .map(x => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(x[0]).setBackground(x[1]).setFontColor(x[2]).setRanges([jr]).build()));
  [80, 130, 220, 100, 90, 120, 80, 55, 90, 90, 80, 80, 60, 70, 75, 90, 200, 160, 85, 50].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A1').setNote('タスク名を入力するとIDが自動で付きます。カレンダーの予定の件名か説明に「T-0001」のようにIDを書くと、その時間がこのタスクの実績になります。');
}

function taskFormulas_(r) {
  const H = '工数ログ!$D:$D', I = '工数ログ!$I:$I', A = '工数ログ!$A:$A', N = '工数ログ!$N:$N';
  return {
    dept: '=IF(D' + r + '="","",IFERROR(VLOOKUP(D' + r + ',設定!$A$5:$B$34,2,FALSE),""))',
    tail: [
      '=IF(A' + r + '="","",SUMIFS(' + H + ',' + I + ',A' + r + ')+SUMIFS(' + H + ',' + N + ',"*,"&A' + r + '&",*"))',
      '=IF(OR(A' + r + '="",N(K' + r + ')=0),"",L' + r + '/K' + r + ')',
      '=IF(OR(A' + r + '="",K' + r + '=""),"",K' + r + '-L' + r + ')',
      '=IF(C' + r + '="","",IF(T' + r + '<>"","定常業務",IF(G' + r + '="完了","完了",IF(AND(J' + r + '<>"",J' + r + '<TODAY()),"期限超過",IF(AND(N(K' + r + ')>0,L' + r + '>K' + r + '),"工数超過",IF(AND(N(K' + r + ')>0,L' + r + '>=K' + r + '*0.8),"注意","順調"))))))',
      '=IF(A' + r + '="","",IFERROR(1/(1/MAX(MAXIFS(' + A + ',' + I + ',A' + r + '),MAXIFS(' + A + ',' + N + ',"*,"&A' + r + '&",*"))),""))'
    ]
  };
}

function writeTaskFormulas_(sh, r) {
  const f = taskFormulas_(r);
  sh.getRange(r, TC.dept).setFormula(f.dept);
  sh.getRange(r, TC.actual, 1, 5).setFormulas([f.tail]);
}

function buildLog_(sh) {
  head_(sh.getRange(1, 1, 1, LOG_HEAD.length), LOG_HEAD);
  sh.setFrozenRows(1);
  sh.getRange('A2:A').setNumberFormat('yyyy/mm/dd');
  sh.getRange('B2:C').setNumberFormat('HH:mm');
  sh.getRange('D2:D').setNumberFormat('0.00');
  [85, 55, 55, 60, 100, 90, 120, 120, 75, 300, 90, 120, 200].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A1').setNote('カレンダー取り込みで自動更新されます。手で編集しないでください（次回の取り込みで上書きされます）。');
}

function buildSummary_(sh) {
  title_(sh, '業務カテゴリ別 工数集計', '工数ログから自動集計。期間（黄色）は上書きすれば任意期間で集計できます。');
  sh.getRange(3, 1, 5, 2).setValues([
    ['集計開始日', '=IF(COUNT(工数ログ!A2:A)=0,TODAY(),MIN(工数ログ!A2:A))'],
    ['集計終了日', '=IF(COUNT(工数ログ!A2:A)=0,TODAY(),MAX(工数ログ!A2:A))'],
    ['期間の営業日数', '=NETWORKDAYS(B3,B4)'],
    ['月の営業日数', '=設定!L10'],
    ['月換算係数', '=IF(B5=0,0,B6/B5)']]);
  box_(sh.getRange(3, 1, 5, 1), CLR.sub).setFontWeight('bold');
  box_(sh.getRange(3, 2, 2, 1), CLR.yel).setNumberFormat('yyyy/mm/dd');
  box_(sh.getRange(5, 2, 3, 1), CLR.gry);
  sh.getRange('B7').setNumberFormat('0.00');

  head2_(sh, 'A9:A10', '業務カテゴリ');
  head2_(sh, 'B9:D9', '期間合計（時間）');
  head2_(sh, 'E9:G9', '月換算（時間／月）');
  head2_(sh, 'H9:H10', '移管\n候補');
  head2_(sh, 'I9:I10', '部署1のうち\n移管候補\n(時間/月)', CLR.yel, '#000000');
  sh.getRange('B10:G10').setValues([['=設定!L8', '=設定!L9', '合計', '=設定!L8', '=設定!L9', '合計']]);
  head_(sh.getRange('B10:G10'), null, CLR.sub, '#000000');
  sh.setRowHeight(10, 44);

  const rows = [];
  for (let i = 0; i < 16; i++) {
    const r = 11 + i;
    const s = col => '=IF($A' + r + '="","",SUMIFS(工数ログ!$D:$D,工数ログ!$F:$F,' + col + '$10,工数ログ!$G:$G,$A' + r + ',工数ログ!$A:$A,">="&$B$3,工数ログ!$A:$A,"<="&$B$4))';
    rows.push([
      i < 15 ? '=IF(設定!G' + (5 + i) + '="","",設定!G' + (5 + i) + ')' : '未分類',
      s('B'), s('C'),
      '=IF($A' + r + '="","",B' + r + '+C' + r + ')',
      '=IF($A' + r + '="","",B' + r + '*$B$7)',
      '=IF($A' + r + '="","",C' + r + '*$B$7)',
      '=IF($A' + r + '="","",D' + r + '*$B$7)',
      i < 15 ? '=IF($A' + r + '="","",設定!I' + (5 + i) + ')' : '',
      '=IF(H' + r + '="○",E' + r + ',0)']);
  }
  sh.getRange(11, 1, 16, 9).setValues(rows);
  box_(sh.getRange(11, 2, 16, 6), CLR.gry).setNumberFormat('0.0;-0.0;"-"');
  box_(sh.getRange(11, 1, 16, 1));
  box_(sh.getRange(11, 8, 16, 1)).setHorizontalAlignment('center');
  box_(sh.getRange(11, 9, 16, 1), CLR.yel).setNumberFormat('0.0;-0.0;"-"');
  const tot = ['合計'];
  'BCDEFG'.split('').forEach(c => tot.push('=SUM(' + c + '11:' + c + '26)'));
  tot.push('', '=SUM(I11:I26)');
  sh.getRange(27, 1, 1, 9).setValues([tot]);
  box_(sh.getRange(27, 1, 1, 9), CLR.sub).setFontWeight('bold').setNumberFormat('0.0;-0.0;"-"');
  [150, 80, 80, 80, 80, 80, 80, 55, 100].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.insertChart(sh.newChart().asBarChart().addRange(sh.getRange('A10:C26')).setStacked().setNumHeaders(1)
    .setPosition(3, 11, 0, 0).setOption('title', 'カテゴリ別工数（期間合計・時間）').setOption('width', 560).setOption('height', 440).build());
}

function buildSim_(sh) {
  title_(sh, '部署1 → 部署2 業務移管シミュレーション', '移管する業務に「○」、移管先での所要率（今の時間に対する比率）を入力してください。');
  sh.getRange('A3:B3').setValues([['人件費単価（円／時間）', '=設定!L11']]);
  box_(sh.getRange('A3'), CLR.sub).setFontWeight('bold');
  box_(sh.getRange('B3'), CLR.gry).setNumberFormat('#,##0');
  const sig = '+0.0;-0.0;"-"', yen = '+#,##0;-#,##0;"-"';
  sh.getRange('A5:B5').setValues([['移管元の削減（時間／月）', '=E25']]);
  sh.getRange('D5:F5').setValues([['移管先の増加（時間／月）', '', '=F25']]);
  sh.getRange('G5:H6').setValues([['会社全体（時間／月）', '=G25'], ['会社全体（円／年）', '=I25']]);
  [sh.getRange('A5'), sh.getRange('D5'), sh.getRange('G5:G6')].forEach(r => box_(r, CLR.sub).setFontWeight('bold'));
  box_(sh.getRange('B5'), CLR.gry).setNumberFormat(sig);
  box_(sh.getRange('F5'), CLR.gry).setNumberFormat(sig);
  box_(sh.getRange('H5'), CLR.yel).setNumberFormat(sig).setFontWeight('bold');
  box_(sh.getRange('H6'), CLR.yel).setNumberFormat(yen).setFontWeight('bold');

  head2_(sh, 'A8:A9', '業務カテゴリ');
  head2_(sh, 'B8:B8', '現在');
  head2_(sh, 'C8:D8', '移管設定（入力）', CLR.yel, '#000000');
  head2_(sh, 'E8:G8', '移管後の変化（時間／月）');
  head2_(sh, 'H8:I8', '金額換算（円）');
  sh.getRange('B9:I9').setValues([['部署1の\n工数\n(時間/月)', '移管\nする', '移管先の\n所要率', '移管元', '移管先', '全体効果', '月額', '年額']]);
  head_(sh.getRange('B9:I9'), null, CLR.sub, '#000000');
  sh.setRowHeight(9, 46);
  const rows = [];
  for (let i = 0; i < 15; i++) {
    const r = 10 + i, off = 'OR($A' + r + '="",$C' + r + '<>"○")';
    rows.push(['=集計!A' + (11 + i), '=IF($A' + r + '="","",集計!E' + (11 + i) + ')', '=IF($A' + r + '="","",集計!H' + (11 + i) + ')',
      (i === 0 || i === 4) ? 0.6 : 0.7,
      '=IF(' + off + ',0,-$B' + r + ')', '=IF(' + off + ',0,$B' + r + '*$D' + r + ')', '=E' + r + '+F' + r,
      '=G' + r + '*$B$3', '=H' + r + '*12']);
  }
  sh.getRange(10, 1, 15, 9).setValues(rows);
  box_(sh.getRange(10, 1, 15, 1));
  box_(sh.getRange(10, 2, 15, 1), CLR.gry).setNumberFormat('0.0;-0.0;"-"');
  box_(sh.getRange(10, 3, 15, 1), CLR.yel).setHorizontalAlignment('center').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['○', '×'], true).setAllowInvalid(true).build());
  box_(sh.getRange(10, 4, 15, 1), CLR.yel).setNumberFormat('0%');
  box_(sh.getRange(10, 5, 15, 2), CLR.gry).setNumberFormat(sig);
  box_(sh.getRange(10, 7, 15, 1), CLR.yel).setNumberFormat(sig).setFontWeight('bold');
  box_(sh.getRange(10, 8, 15, 2), CLR.gry).setNumberFormat(yen);
  sh.getRange(25, 1, 1, 9).setValues([['合計', '=SUM(B10:B24)', '', '', '=SUM(E10:E24)', '=SUM(F10:F24)', '=SUM(G10:G24)', '=SUM(H10:H24)', '=SUM(I10:I24)']]);
  box_(sh.getRange(25, 1, 1, 9), CLR.sub).setFontWeight('bold');
  sh.getRange('B25').setNumberFormat('0.0'); sh.getRange('E25:G25').setNumberFormat(sig); sh.getRange('H25:I25').setNumberFormat(yen);
  sh.getRange('A27').setValue('※ 所要率の初期値（60〜70%）は仮置き。移管先の担当者にヒアリングして実態に合わせてください。マイナス＝削減（効率化）。').setFontSize(9).setFontColor('#595959');
  sh.getRange('D9').setNote('移管先が担当すると今の何%の時間で済むか。例：60% → 40時間が24時間になる');
  [150, 85, 55, 75, 75, 75, 85, 95, 105].forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function buildDaily_(sh) {
  head_(sh.getRange(1, 1, 1, DAILY_HEAD.length), DAILY_HEAD);
  sh.setFrozenRows(1);
  sh.getRange('A2:A').setNumberFormat('yyyy/mm/dd');
  sh.getRange('C2:C').setNumberFormat('0.0');
  sh.getRange('F2:F').setNumberFormat('yyyy/mm/dd hh:mm');
  sh.getRange('D2:E').setWrap(true).setVerticalAlignment('top');
  [85, 100, 60, 380, 320, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

// ================= 読み書きヘルパ =================
function readConfig_() {
  const sh = sheet_(SH.CONF);
  const allMembers = sh.getRange(5, 1, MEMBER_ROWS, 4).getValues()
    .map((r, i) => ({ row: 5 + i, name: String(r[0]).trim(), dept: String(r[1]).trim(), calId: String(r[2]).trim(), on: r[3] === true }))
    .filter(x => x.name);
  const cats = sh.getRange(5, 7, CAT_ROWS, 3).getValues().filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), kws: splitList_(r[1]), transfer: String(r[2]).trim() === '○' }));
  ensureConfigRows_(sh);
  const s = sh.getRange(5, 12, 11, 1).getValues().map(r => r[0]);
  return {
    sheet: sh, allMembers: allMembers, members: allMembers.filter(x => x.calId && x.on), cats: cats,
    startDate: s[0] instanceof Date ? startOfDay_(s[0]) : null,
    recentDays: Number(s[1]) || 14, excludes: splitList_(s[2]),
    depts: [String(s[3]).trim(), String(s[4]).trim()],
    monthDays: Number(s[5]) || 20, wage: Number(s[6]) || 0,
    lastImport: s[7] instanceof Date ? Utilities.formatDate(s[7], tz_(), 'yyyy/MM/dd HH:mm') : '',
    admins: splitList_(s[8]).map(x => x.toLowerCase()),
    taskCals: splitList_(s[9]), lastTaskSync: String(s[10] || '')
  };
}

function readTasks_() {
  const sh = sheet_(SH.TASK);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const out = [];
  sh.getRange(2, 1, n, TASK_HEAD.length).getValues().forEach(r => {
    if (!r[0] || !String(r[2]).trim()) return;
    out.push({
      id: String(r[0]), proj: String(r[1]), name: String(r[2]), owner: String(r[3]), dept: String(r[4]), cat: String(r[5]),
      status: STATUSES.indexOf(r[6]) >= 0 ? r[6] : '未着手', prio: String(r[7] || ''),
      start: ymd_(r[8]), due: ymd_(r[9]), plan: num_(r[10]), actual: num_(r[11]) || 0,
      judge: String(r[14] || ''), last: ymd_(r[15]), memo: String(r[16] || ''), calEv: String(r[17] || ''),
      parent: String(r[18] || ''), level: LEVELS.indexOf(r[19]) >= 0 ? r[19] : ''
    });
  });
  return out;
}

function readLogs_(from, toEx) {
  const sh = sheet_(SH.LOG);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 11).getValues()
    .filter(r => r[0] instanceof Date && r[0] >= from && r[0] < toEx)
    .map(r => ({ date: r[0], ymd: ymd_(r[0]), start: r[1], end: r[2], h: Number(r[3]) || 0, name: String(r[4]), dept: String(r[5]),
      cat: String(r[6] || '未分類'), proj: String(r[7] || ''), task: String(r[8] || ''), title: String(r[9] || '') }));
}

function writeTaskFields_(sh, row, p) {
  Object.keys(FIELD_COL).forEach(k => {
    if (!(k in p)) return;
    let v = p[k];
    if (k === 'start' || k === 'due') v = v ? parseYmd_(v) : '';
    else if (k === 'plan') v = (v === '' || v == null) ? '' : Number(v);
    else if (k === 'parent' && v && String(sh.getRange(row, TC.id).getValue()) === String(v)) throw new Error('自分自身を親タスクにはできません');
    else v = v == null ? '' : String(v);
    sh.getRange(row, FIELD_COL[k]).setValue(v);
  });
}

function findTaskRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}

function assignTaskIds_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const v = sh.getRange(2, 1, n, 3).getValues();
  let max = maxTaskNo_(v.map(r => r[0])), changed = false;
  v.forEach(r => { if (!r[0] && String(r[2]).trim()) { r[0] = fmtId_(++max); changed = true; } });
  if (changed) sh.getRange(2, 1, n, 1).setValues(v.map(r => [r[0]]));
}

function findDaily_(name, ymdStr) {
  const sh = sheet_(SH.DAILY);
  const n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const v = sh.getRange(2, 1, n, 6).getValues();
  for (let i = 0; i < v.length; i++) {
    if (ymd_(v[i][0]) === ymdStr && String(v[i][1]) === name) {
      return { row: i + 2, comment: String(v[i][4] || ''), updated: v[i][5] instanceof Date ? Utilities.formatDate(v[i][5], tz_(), 'yyyy/MM/dd HH:mm') : '' };
    }
  }
  return null;
}

function maxTaskNo_(ids) { return ids.reduce((m, id) => { const x = String(id).match(/T-(\d+)/); return x ? Math.max(m, Number(x[1])) : m; }, 0); }
function fmtId_(n) { return 'T-' + ('0000' + n).slice(-Math.max(4, String(n).length)); }
function splitList_(v) { return String(v || '').split(/[,、，\n]/).map(s => s.trim()).filter(Boolean); }
function num_(x) { return (x === '' || x == null || isNaN(Number(x))) ? null : Math.round(Number(x) * 100) / 100; }
function tz_() { return Session.getScriptTimeZone() || 'Asia/Tokyo'; }
function ymd_(d) { return d instanceof Date ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') : ''; }
function hm_(d) { return d instanceof Date ? Utilities.formatDate(d, tz_(), 'HH:mm') : ''; }
function parseYmd_(s) { const m = String(s).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/); if (!m) throw new Error('日付の形式が不正です: ' + s); return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])); }
function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays_(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function weekStart_(d) { return addDays_(d, -((d.getDay() + 6) % 7)); }
function bizDays_(from, to) { let n = 0; for (let d = startOfDay_(from); d <= to; d = addDays_(d, 1)) if (d.getDay() % 6 !== 0) n++; return n; }

function ss_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) {}
  if (ss) return ss;
  const id = PropertiesService.getScriptProperties().getProperty('SSID');
  if (!id) throw new Error('スプレッドシートを開いてメニュー「① 初期設定」を一度実行してください。');
  return SpreadsheetApp.openById(id);
}
function saveSsId_() { PropertiesService.getScriptProperties().setProperty('SSID', SpreadsheetApp.getActive().getId()); }
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('「' + name + '」シートがありません。メニュー「工数管理 → ① 初期設定」を実行してください。');
  return sh;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---- 書式ヘルパ ----
function title_(sh, t, sub) {
  sh.getRange('A1').setValue(t).setFontSize(14).setFontWeight('bold').setFontColor(CLR.head);
  sh.getRange('A2').setValue(sub).setFontSize(9).setFontColor('#595959');
}
function head_(rng, values, bg, fg) {
  if (values) rng.setValues([values]);
  return rng.setBackground(bg || CLR.head).setFontColor(fg || '#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, true, true, CLR.line, SpreadsheetApp.BorderStyle.SOLID);
}
function head2_(sh, a1, text, bg, fg) {
  const r = sh.getRange(a1);
  if (r.getNumRows() > 1 || r.getNumColumns() > 1) r.merge();
  r.getCell(1, 1).setValue(text);
  head_(r, null, bg, fg);
}
function box_(rng, bg) {
  rng.setBorder(true, true, true, true, true, true, CLR.line, SpreadsheetApp.BorderStyle.SOLID).setVerticalAlignment('middle');
  if (bg) rng.setBackground(bg);
  return rng;
}
function listRule_(arr) { return SpreadsheetApp.newDataValidation().requireValueInList(arr, true).setAllowInvalid(false).build(); }
function rangeRule_(rng) { return SpreadsheetApp.newDataValidation().requireValueInRange(rng, true).setAllowInvalid(true).build(); }
