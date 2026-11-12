// 页面层：只调用 core 的公开 API；所有数据只存在本机 IndexedDB，不连任何外部地址
import * as XLSX from '../vendor/xlsx.mjs';
import { Store, ATT, MODES, STATUS_LIST, KIND_LABEL, courseLabel, presentCount, totalJie, migrate, SESSION_TAGS, sessionTeacher } from './core/store.js';
import { teacherSummary, studentSummary, absenceStreaks, streakKey } from './core/query.js';
import { readRosterSheet, planImport, writeRosterTemplate } from './core/roster.js';
import { buildStats, layoutSheet, writeMonthlyWorkbook, writeAllDataWorkbook, precheck, consumeLabel, EXPORT_FORMATS } from './core/export.js';
import { makeBackup, readBackup } from './core/backup.js';
import { fmtJie, fmtNum, keshi, isHalfStep } from './core/jie.js';
import { matchStudent } from './core/pinyin.js';
import { dayOf, fmtDay, monthRange, isDay } from './core/dates.js';
import * as db from './db.js';
import { makePinRecord, checkPin, isValidPin } from './pin.js';
import { esc, sheet, confirmDlg, alertDlg, promptDlg, toast, pinPad, askPin } from './ui.js';

export const APP_VERSION = '2.0.0';
const AUTO_LOCK_MS = 3 * 60 * 1000;
const XLSX_ACCEPT = '.xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel';
const $app = document.getElementById('app');
const S = { store: null, pinRec: null, unlocked: false, view: { name: 'home' }, hiddenAt: 0, persisted: null, integrity: [],
  saveChain: Promise.resolve(), lastExport: null, wxDismissed: false, draft: null };

const isStandalone = () => window.matchMedia?.('(display-mode: standalone)').matches || window.navigator.standalone === true;
const isWeChat = () => /MicroMessenger/i.test(navigator.userAgent);
const today = () => dayOf(new Date());
const st = () => S.store.state;
const stuName = (id) => st().students[id]?.name ?? '（已删除学生）';
const thisMonth = () => today().slice(0, 7);
const prevMonth = () => { const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1); return dayOf(d).slice(0, 7); };
const fmtTime = (iso) => { const d = new Date(iso); return `${dayOf(d)} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`; };

// ---------- 启动 ----------
async function boot() {
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js').catch(() => {});
  try { S.persisted = navigator.storage?.persisted ? await navigator.storage.persisted() : null;
    if (S.persisted === false && navigator.storage?.persist) S.persisted = await navigator.storage.persist(); } catch { S.persisted = null; }
  S.lastExport = await db.kvGet('lastBackup').catch(() => null);
  const rec = await db.loadData();
  if (rec && (rec.state.schema ?? 1) < 2) {
    // V1 → V2 升级：先把原样数据存一份快照，再补齐新字段（不删改已有数据）
    await db.addSnapshot('升级 V2 前自动快照', rec.state, rec.ops);
    migrate(rec.state);
    await db.saveData(rec.state, rec.ops);
  } else if (rec) migrate(rec.state);
  S.store = rec ? new Store({ state: rec.state, ops: rec.ops }) : new Store();
  if (!rec) await db.saveData(S.store.state, S.store.ops);
  attach();
  S.integrity = S.store.verifyIntegrity();
  const snaps = await db.listSnapshots();
  if (!snaps.length || dayOf(snaps[0].createdAt) !== today()) await db.addSnapshot('每日自动快照', S.store.state, S.store.ops);
  S.pinRec = await db.kvGet('pin');
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) S.hiddenAt = Date.now();
    else if (S.unlocked && S.pinRec && Date.now() - S.hiddenAt > AUTO_LOCK_MS) { S.unlocked = false; render(); }
  });
  render();
}
function attach() {
  S.store.onCommit = (op) => {
    const { state, ops } = S.store;
    S.saveChain = S.saveChain.then(() => db.saveData(state, ops))
      .catch((e) => alertDlg('保存失败', '数据没有写入手机存储：' + (e?.message ?? e) + '\n请立即导出完整备份。'));
    if (['importRoster', 'undo'].includes(op.kind)) db.addSnapshot(`${KIND_LABEL[op.kind]}后`, state, ops).catch(() => {});
  };
}
const flush = () => S.saveChain;

// ---------- 路由与渲染 ----------
function go(name, params = {}) { S.view = { name, ...params }; render(); window.scrollTo(0, 0); }
function render() {
  if (!S.pinRec) return renderSetPin();
  if (!S.unlocked) return renderLock();
  const v = VIEWS[S.view.name] ?? VIEWS.home;
  $app.innerHTML = wxBanner() + v(S.view);
  AFTER[S.view.name]?.(S.view);
}
function wxBanner() {
  if (!isWeChat()) return '';
  return `<div class="wx" id="wx-tip"><b>请在 Safari 中打开</b>点右上角「···」→「在 Safari 中打开」（安卓选「在浏览器打开」）。在微信里使用，数据可能丢失，也不能添加到主屏幕。</div>`;
}
function renderSetPin() {
  let first = '';
  const step2 = async (p) => {
    if (p !== first) { renderSetPin(); toast('两次输入不一致，请重新设置'); return null; }
    S.pinRec = await makePinRecord(p); await db.kvSet('pin', S.pinRec); S.unlocked = true; render(); return null;
  };
  $app.innerHTML = wxBanner() + '<div id="pinbox"></div>';
  pinPad(document.getElementById('pinbox'), { title: '设置 6 位密码', sub: '用于打开 App 和修改已锁账的月份。请牢记。',
    onDone: async (p) => { if (!isValidPin(p)) return '请输入 6 位数字'; first = p; pinPad(document.getElementById('pinbox'), { title: '再输入一次', sub: '确认密码', onDone: step2 }); return null; } });
}
function renderLock() {
  $app.innerHTML = wxBanner() + '<div id="pinbox"></div>';
  pinPad(document.getElementById('pinbox'), { title: '签课', sub: '输入密码', onDone: async (p) => {
    if (await checkPin(p, S.pinRec)) { S.unlocked = true; render(); return null; }
    return '密码不对，请重试';
  } });
}
const pinCheck = (title) => askPin(title, (p) => checkPin(p, S.pinRec));

const top = (title, { back = null, backId = '', right = '' } = {}) => `<div class="top">${back ? `<button class="iconbtn back" data-go="${back}" ${backId ? `data-id="${backId}"` : ''}>‹ 返回</button>` : '<span></span>'}
  <h1>${esc(title)}</h1>${right || '<span class="spacer"></span>'}</div>`;
const VIEWS = {}, AFTER = {};

// ---------- 首页 ----------
function backupDue() {
  if (!Object.keys(st().sessions).length) return false;
  return !S.lastExport || Date.now() - Date.parse(S.lastExport) > 7 * 864e5;
}
VIEWS.home = () => {
  const courses = S.store.courses();
  const todayS = S.store.sessions({ from: today(), to: today() });
  const done = new Set(todayS.map((s) => s.courseId));
  const banners = [
    !isStandalone() && !isWeChat() ? `<div class="banner" id="a2hs"><b>请添加到主屏幕使用</b>点 Safari 底部的「分享」→「添加到主屏幕」，以后从主屏图标打开。只在浏览器标签页里用，系统可能清理数据。</div>` : '',
    S.integrity.length ? `<div class="banner"><b>数据自检发现问题</b>${esc(S.integrity.slice(0, 2).join('；'))}。请导出完整备份并联系管理员。</div>` : '',
    backupDue() ? `<div class="banner" id="backup-due"><b>该备份了</b>超过 7 天没有导出完整备份。<button class="linkbtn" data-go="backup">去备份 ›</button></div>` : '',
  ].join('');
  const list = courses.length ? `<div class="list">${courses.map((c) => `<button class="row" data-act="start-roll" data-id="${c.id}">
      <div class="main"><div class="t">${esc(courseLabel(c))}</div><div class="s">${c.studentIds.length} 人 · 默认 ${fmtNum(c.defaultJie)} 节${done.has(c.id) ? ' · <b class="ok">今天已签</b>' : ''}</div></div><span class="chev">签课 ›</span></button>`).join('')}</div>`
    : `<div class="empty">还没有课程。<br>先「导入名单」，再「建课程」。</div>`;
  return `${top('签课', { right: '<button class="iconbtn" id="btn-menu" data-go="menu">菜单</button>' })}${banners}
  <div class="page">
    ${!Object.keys(st().students).length ? `<div class="card"><b>第一步：导入学生名单</b><p class="hint">把「学生总名单」Excel 通过微信或 AirDrop 存到「文件」，再从这里导入。</p><button class="btn primary" data-go="import">导入名单</button><button class="btn ghost" data-go="help">看使用教程</button></div>` : ''}
    ${streakHtml()}
    <div class="section">课程（点一下就签课）</div>${list}
    <div class="section">今天已签 ${todayS.length} 次</div>
    ${todayS.length ? `<div class="list">${todayS.map(sessRow).join('')}</div>` : '<div class="empty">今天还没签课</div>'}
  </div>
  <div class="fab-wrap"><button class="btn primary big" id="btn-roll" data-act="pick-course">签课</button></div>`;
};
function streakHtml() {
  const list = absenceStreaks(st());
  if (!list.length) return '';
  return `<div class="section err-text">连续缺课提醒（同一门课连续 2 次及以上请假/缺勤）</div><div class="list" id="streaks">${list.slice(0, 20).map((x) => `<button class="row err" data-go="byStudent" data-id="${x.studentId}"><div class="main"><div class="t">${esc(x.name)} · 连续 ${x.streak} 次</div><div class="s">${esc(x.course)} · ${esc(x.dates.slice(-3).join('、'))}</div></div><span class="chev">›</span></button>`).join('')}</div>`;
}
const tagText = (s) => (s.tag ? ` · ${s.tag}` : '') + (s.subTeacher ? ` · 代课：${s.subTeacher}` : '');
const sessRow = (s) => `<button class="row" data-go="session" data-id="${s.id}"><div class="main"><div class="t">${esc(courseLabel(st().courses[s.courseId]))}</div>
  <div class="s">${fmtDay(s.date)} · 到课 ${presentCount(s)} 人${leaveAbsText(s)} · 共 ${fmtNum(totalJie(s))} 节${esc(tagText(s))}${s.edits ? ` · 改过 ${s.edits} 次` : ''}</div></div>${S.store.isLocked(s.date.slice(0, 7)) ? '<span class="tag">已锁账</span>' : '<span class="tag">已签</span>'}</button>`;
const leaveAbsText = (s) => { const l = s.records.filter((r) => r.status === ATT.leave).length, a = s.records.filter((r) => r.status === ATT.absent).length; return (l ? ` · 请假 ${l}` : '') + (a ? ` · 缺勤 ${a}` : ''); };

// ---------- 菜单 ----------
VIEWS.menu = () => `${top('菜单', { back: 'home' })}<div class="page"><div class="list menu">
  ${[['sessions', '签课记录', '按月份查看、修改'], ['students', '学生名单', `${S.store.students().length} 人在读`], ['import', '导入名单', '从 Excel 导入或合并'], ['courses', '课程模板', `${S.store.courses().length} 门`],
    ['byTeacher', '按老师查课时', '本月每位老师上了多少节'], ['byStudent', '按学生查次数', '到课、请假、缺勤次数'],
    ['export', '月底导出课时统计表', '导出前自检，导出后锁账'], ['backup', '备份与恢复', S.lastExport ? '上次备份 ' + fmtTime(S.lastExport) : '还没备份过'], ['log', '操作记录', `${S.store.ops.length} 条`],
    ['help', '使用教程', '一页看懂'], ['settings', '设置', '校区、密码、锁账月份']]
    .map(([v, t, s]) => `<button class="row" data-go="${v}" id="m-${v}"><div class="main"><div class="t">${t}</div><div class="s">${esc(s)}</div></div><span class="chev">›</span></button>`).join('')}
  </div><p class="hint">版本 ${APP_VERSION} · 数据只存在这台手机，不上传。</p></div>`;

// ---------- 课程模板 ----------
VIEWS.courses = () => {
  const all = S.store.courses({ includeArchived: true });
  return `${top('课程模板', { back: 'menu', right: '<button class="iconbtn back" id="btn-add-course" data-act="edit-course">新建</button>' })}<div class="page">
  <p class="hint">固定好年级、科目、老师和学生，首页点一下就能签课。一对一每个学生每个科目建一门。</p>
  ${all.length ? `<div class="list">${all.map((c) => `<button class="row" data-act="edit-course" data-id="${c.id}"><div class="main"><div class="t">${esc(courseLabel(c))}${c.archived ? '（已停用）' : ''}</div><div class="s">${c.studentIds.length} 人：${esc(c.studentIds.map(stuName).join('、')) || '未选学生'}</div></div><span class="chev">›</span></button>`).join('')}</div>` : '<div class="empty">还没有课程，点右上角「新建」</div>'}</div>`;
};
VIEWS.course = ({ id }) => {
  const c = S.draft;
  const studs = S.store.students({ includeInactive: S.draft._all });
  const q = S.draft._q ?? '';
  const sel = new Set(c.studentIds);
  const chosen = c.studentIds.map((sid) => st().students[sid]).filter(Boolean);
  return `${top(id ? '修改课程' : '新建课程', { back: 'courses' })}<div class="page" id="course-form">
  <div class="grid2"><label class="field"><span>校区</span><input class="input" id="c-campus" value="${esc(c.campus)}" placeholder="如 曲江校区"></label>
  <label class="field"><span>年级</span><input class="input" id="c-grade" value="${esc(c.grade)}" placeholder="如 高二"></label></div>
  <div class="toggle" id="c-mode">${MODES.map((m) => `<button data-mode="${m}" class="${c.mode === m ? 'on' : ''}">${m}</button>`).join('')}</div>
  <div class="grid2"><label class="field"><span>科目</span><input class="input" id="c-subject" value="${esc(c.subject)}" placeholder="如 数学"></label>
  <label class="field"><span>教师</span><input class="input" id="c-teacher" value="${esc(c.teacher)}" placeholder="教师姓名"></label></div>
  <div class="grid2"><label class="field"><span>课程名称（可不填）</span><input class="input" id="c-name" value="${esc(c.name)}" placeholder="如 高二数学班"></label>
  <label class="field"><span>每次默认节数（1节=2课时）</span><input class="input" id="c-jie" inputmode="decimal" value="${fmtNum(c.defaultJie)}"></label></div>
  <div class="section">已选学生 ${chosen.length} 人${c.mode === '一对一' ? '（一对一只能选 1 人）' : ''}</div>
  <div class="chips">${chosen.map((s) => `<button class="chip on" data-act="toggle-stu" data-id="${s.id}">${esc(s.name)} ✕</button>`).join('') || '<span class="hint">从下面勾选</span>'}</div>
  <div class="section">从名单里选</div>
  <div class="grid2"><input class="input" id="c-q" type="search" autocomplete="off" placeholder="搜姓名/拼音首字母/编号" value="${esc(q)}"><button class="btn" data-act="course-all">${S.draft._all ? '只看在读' : '显示全部（含停课）'}</button></div>
  <p class="hint">只能选学生总名单里已有的学生；名单里没有的，请先把他加进总名单 Excel 再导入。</p>
  <div id="pick-list">${coursePickHtml(studs, q, sel)}</div>
  ${id ? `<button class="btn mt" data-act="archive-course" data-id="${id}">${st().courses[id].archived ? '恢复使用' : '停用这门课程'}</button>` : ''}
  </div><div class="fab-wrap"><button class="btn primary big" id="btn-save-course" data-act="save-course">保存课程</button></div>`;
};
function coursePickHtml(studs, q, sel) {
  const shown = studs.filter((s) => matchStudent(s, q));
  return `<div class="list pick">${shown.slice(0, 300).map((s) => `<label class="row"><input type="checkbox" data-act="toggle-stu" data-id="${s.id}" ${sel.has(s.id) ? 'checked' : ''}><div class="main"><div class="t">${esc(s.name)}</div><div class="s">${esc([s.grade, s.school, s.code, s.status !== '在读' ? s.status : ''].filter(Boolean).join(' · '))}</div></div></label>`).join('') || '<div class="empty">名单里没有匹配的学生</div>'}</div>
  ${shown.length > 300 ? '<p class="hint">只显示前 300 人，请用搜索缩小范围。</p>' : ''}`;
}
function readCourseForm() {
  const g = (i) => document.getElementById(i)?.value ?? '';
  Object.assign(S.draft, { campus: g('c-campus'), grade: g('c-grade'), subject: g('c-subject'), teacher: g('c-teacher'), name: g('c-name'), defaultJie: Number(g('c-jie')) });
}

// ---------- 签课 ----------
function newRollDraft(course, date = today()) {
  return { courseId: course.id, date, note: '', editId: null, tag: '', subTeacher: '',
    records: course.studentIds.filter((id) => st().students[id]).map((studentId) => ({ studentId, status: ATT.present, jie: course.defaultJie })) };
}
VIEWS.roll = () => {
  const d = S.draft; const c = st().courses[d.courseId];
  const n = d.records.filter((r) => r.status === ATT.present).length;
  const sum = d.records.reduce((a, r) => a + (r.status === ATT.present ? r.jie : 0), 0);
  const locked = S.store.isLocked(d.date.slice(0, 7));
  return `${top(d.editId ? '修改签课' : '签课', { back: d.editId ? 'session' : 'home', backId: d.editId ?? '' })}<div class="page" id="roll">
  <div class="card"><div class="t-strong">${esc(courseLabel(c))}</div>
  <label class="field"><span>上课日期</span><input class="input" type="date" id="r-date" value="${d.date}"></label>
  ${locked ? `<div class="banner inline"><b>${+d.date.slice(5, 7)}月已锁账</b>保存时需要输入密码。</div>` : ''}
  <div class="hint">默认全部到课。缺的人点「请假」或「缺勤」；节数可按人改（每次 ±0.5 节）。</div></div>
  <div class="list">${d.records.map((r, i) => `<div class="row roll-row ${r.status !== ATT.present ? 'off' : ''}" data-i="${i}">
    <div class="main"><div class="t">${esc(stuName(r.studentId))}</div><div class="s">${r.status === ATT.present ? fmtJie(r.jie) : r.status}</div></div>
    <div class="seg3">${[ATT.present, ATT.leave, ATT.absent].map((s) => `<button data-act="set-status" data-i="${i}" data-s="${s}" class="${r.status === s ? 'on' : ''}">${s}</button>`).join('')}</div>
    <div class="stepper ${r.status !== ATT.present ? 'hidden' : ''}"><button data-act="jie" data-i="${i}" data-d="-0.5">−</button><span class="num">${fmtNum(r.jie)}</span><button data-act="jie" data-i="${i}" data-d="0.5">＋</button></div>
  </div>`).join('') || '<div class="empty">这门课程还没有学生，请到「课程模板」里添加</div>'}</div>
  <p class="hint">只能签课程里的学生。要加人，请到「课程模板」从总名单里选。</p>
  <div class="field"><span class="lbl">这次课的标记</span><div class="seg3" id="r-tag">${SESSION_TAGS.map((t) => `<button data-act="set-tag" data-t="${t}" class="${(d.tag ?? '') === t ? 'on' : ''}">${t || '正常上课'}</button>`).join('')}</div></div>
  <label class="field"><span>授课老师（代课时改成代课老师；导出时课时记在他名下）</span><input class="input" id="r-sub" value="${esc(d.subTeacher || c.teacher)}" placeholder="${esc(c.teacher)}"></label>
  <label class="field"><span>备注（可不填）</span><input class="input" id="r-note" value="${esc(d.note)}" placeholder="如 调到周六上"></label>
  </div><div class="fab-wrap"><div class="sum" id="roll-sum">到课 ${n} 人 · 请假 ${d.records.filter((r) => r.status === ATT.leave).length} · 缺勤 ${d.records.filter((r) => r.status === ATT.absent).length} · 共 ${fmtNum(sum)} 节</div>
  <button class="btn primary big" id="btn-save-roll" data-act="save-roll">${d.editId ? '保存修改' : '完成签课'}</button></div>`;
};
async function saveRoll() {
  const d = S.draft;
  d.date = document.getElementById('r-date').value; d.note = document.getElementById('r-note').value; d.subTeacher = document.getElementById('r-sub').value.trim();
  if (!isDay(d.date)) return alertDlg('日期不对', '请选择上课日期');
  if (!d.records.some((r) => r.status === ATT.present) && !(await confirmDlg('没有人到课', '这次课全部请假或缺勤，确定保存吗？', '仍保存'))) return;
  const input = { courseId: d.courseId, date: d.date, records: d.records, note: d.note, tag: d.tag ?? '', subTeacher: d.subTeacher === st().courses[d.courseId].teacher ? '' : d.subTeacher };
  let pinOk = false;
  const needPin = S.store.isLocked(d.date.slice(0, 7)) || (d.editId && S.store.isLocked(st().sessions[d.editId].date.slice(0, 7)));
  if (needPin) { if (!(await pinCheck('该月已锁账，输入密码继续'))) return; pinOk = true; }
  try {
    if (d.editId) {
      const reason = await promptDlg('修改原因（必填）', { placeholder: '如 学生乙实际请假，记错了', okLabel: '保存修改' });
      if (reason == null) return;
      if (!reason.trim()) return alertDlg('请填写原因', '已签的课修改时必须写原因，会留在操作记录里。');
      S.store.editSession(d.editId, input, { reason, pinOk });
      await flush(); toast('已修改，改动已留痕'); go('session', { id: d.editId });
    } else {
      const dup = S.store.sessionsOn(d.courseId, d.date);
      if (dup.length) {
        const v = await sheet({ title: '同一个班同一天已经签过', message: `${fmtDay(d.date)} 这门课已签 ${dup.length} 次（到课 ${presentCount(dup[0])} 人）。是不是重复了？`,
          buttons: [{ label: '查看已签的那次', value: 'view', kind: 'primary' }, { label: '确实再上了一次，继续保存', value: 'save' }, { label: '取消', value: null }] });
        if (v === 'view') return go('session', { id: dup[0].id });
        if (v !== 'save') return;
      }
      const s = S.store.addSession(input, { pinOk });
      await flush(); toast('签课完成，已锁定'); go('session', { id: s.id });
    }
  } catch (e) { alertDlg('没有保存', e.message); }
}
VIEWS.session = ({ id }) => {
  const s = st().sessions[id];
  if (!s) return `${top('签课详情', { back: 'sessions' })}<div class="page"><div class="empty">这条签课已删除</div></div>`;
  const c = st().courses[s.courseId];
  const hist = S.store.historyOf(id).slice().reverse();
  return `${top('签课详情', { back: 'home' })}<div class="page" id="session">
  <div class="card"><div class="t-strong">${esc(courseLabel(c))}</div><div class="hint">${fmtDay(s.date)}${esc(tagText(s))} · 🔒 已签完锁定${S.store.isLocked(s.date.slice(0, 7)) ? ' · 本月已锁账' : ''}</div>
  <div class="grid3"><div><div class="k">到课</div><div class="v">${presentCount(s)} 人</div></div><div><div class="k">请假/缺勤</div><div class="v">${s.records.filter((r) => r.status === ATT.leave).length} / ${s.records.filter((r) => r.status === ATT.absent).length}</div></div><div><div class="k">合计</div><div class="v">${fmtJie(totalJie(s))}</div></div></div></div>
  <div class="list">${s.records.map((r) => `<div class="row ${r.status !== ATT.present ? 'off' : ''}"><div class="main"><div class="t">${esc(stuName(r.studentId))}</div></div><span>${r.status === ATT.present ? fmtJie(r.jie) : r.status}</span></div>`).join('')}</div>
  ${s.note ? `<p class="hint">备注：${esc(s.note)}</p>` : ''}
  <div class="section">留痕（${hist.length} 条）</div>
  <div class="list">${hist.map((o) => `<div class="row"><div class="main"><div class="t wrap">${esc(o.summary)}</div><div class="s">${fmtTime(o.at)}</div></div></div>`).join('')}</div>
  <button class="btn mt" id="btn-edit-session" data-act="edit-session" data-id="${id}">修改（需填原因）</button>
  <button class="btn danger" data-act="del-session" data-id="${id}">删除这次签课</button></div>`;
};
VIEWS.sessions = ({ month = thisMonth() }) => {
  const { from, to } = monthRange(month);
  const list = S.store.sessions({ from, to }).reverse();
  return `${top('签课记录', { back: 'menu' })}<div class="page"><label class="field"><span>月份</span><input class="input" type="month" id="s-month" value="${month}"></label>
  <p class="hint">${S.store.isLocked(month) ? '🔒 本月已锁账，修改需要密码' : '本月未锁账'} · 共 ${list.length} 次</p>
  ${list.length ? `<div class="list">${list.map(sessRow).join('')}</div>` : '<div class="empty">这个月没有签课</div>'}</div>`;
};

// ---------- 学生名单 ----------
VIEWS.students = ({ all = false, q = '' }) => {
  return `${top('学生名单', { back: 'menu' })}<div class="page">
  <div class="grid2"><input class="input" id="stu-q" type="search" autocomplete="off" placeholder="搜姓名/拼音首字母/编号" value="${esc(q)}"><button class="btn" data-act="stu-all">${all ? '只看在读' : '显示全部（含停课）'}</button></div>
  <div id="stu-list">${stuListHtml(all, q)}</div>
  <p class="hint">新学生请先加进「学生总名单」Excel，再「导入名单」合并；这里只能修改已有学生。</p>
  <button class="btn ghost" data-go="import">导入 / 合并名单 Excel</button></div>`;
};
function stuListHtml(all, q) {
  const list = S.store.students({ includeInactive: all }).filter((s) => matchStudent(s, q));
  return `<p class="hint">${all ? '全部' : '在读'} ${list.length} 人</p>
  <div class="list">${list.slice(0, 400).map((s) => `<button class="row" data-act="edit-stu" data-id="${s.id}"><div class="main"><div class="t">${esc(s.name)}${s.status !== '在读' ? ` <span class="tag">${esc(s.status)}</span>` : ''}</div><div class="s">${esc([s.code, s.grade, s.school, s.manager].filter(Boolean).join(' · '))}</div></div></button>`).join('') || '<div class="empty">名单里没有匹配的学生</div>'}</div>`;
}
async function editStudent(id) {
  if (!id || !st().students[id]) return alertDlg('请从总名单导入', '新学生必须先加进「学生总名单」Excel，再「导入名单」。');
  const s = st().students[id];
  const f = (k, label) => `<label class="field"><span>${label}</span><input class="input" data-k="${k}" value="${esc(s[k] ?? '')}"></label>`;
  const v = await sheet({ title: id ? '修改学生' : '新增学生', html: `${f('name', '姓名')}${f('code', '学生编号（没有写 待补）')}<div class="grid2">${f('grade', '年级')}${f('school', '学校')}</div>
    <div class="grid2">${f('cohort', '届别')}${f('manager', '负责学管')}</div><label class="field"><span>状态</span><select class="input" data-k="status">${[...new Set([...STATUS_LIST, s.status])].map((x) => `<option ${x === s.status ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select></label>`,
    buttons: [{ label: '保存', kind: 'primary', value: (w) => Object.fromEntries([...w.querySelectorAll('[data-k]')].map((e) => [e.dataset.k, e.value])) }, { label: '取消', value: null }] });
  if (!v) return;
  try { S.store.editStudent(id, v); await flush(); toast('已保存'); render(); } catch (e) { alertDlg('没有保存', e.message); }
}

// ---------- 导入名单 ----------
VIEWS.import = () => {
  const p = S.importPlan;
  const rowCls = (r) => (r.action === 'error' ? 'err' : r.warnings.length ? 'warn' : '');
  const act = { add: '新增', update: '更新', same: '不变', error: '不导入' };
  return `${top('导入名单', { back: 'menu' })}<div class="page" id="import">
  <div class="card"><b>怎么导入</b><p class="hint">1. 电脑或微信里把「学生总名单」Excel 发到手机，点开后选「存储到文件」。<br>2. 点下面按钮，从「文件」App 里选这个 Excel。<br>3. 看预览，红色行有错不导入，黄色行仅提醒；确认后导入。再次导入会按学生编号合并（编号写「待补」的按姓名），不会删人。</p>
  <label class="btn primary" id="lbl-import">选择 Excel 文件<input type="file" id="f-import" accept="${XLSX_ACCEPT}" class="hidden"></label>
  <button class="btn ghost" data-act="dl-template">下载空白导入模板</button></div>
  ${p ? `<div class="card" id="import-preview"><b>预览：${esc(p.fileName)}（${esc(p.sheetName)} 页）</b>
    <div class="grid3"><div><div class="k">新增</div><div class="v" id="imp-add">${p.summary.add}</div></div><div><div class="k">更新</div><div class="v">${p.summary.update}</div></div><div><div class="k">有错不导入</div><div class="v err-text" id="imp-err">${p.summary.error}</div></div></div>
    <p class="hint">不变 ${p.summary.same} 行，共 ${p.summary.total} 行。</p></div>
    <div class="tbl-wrap"><table><thead><tr><th>行</th><th>处理</th><th>姓名</th><th>学生编号</th><th>年级</th><th>状态</th><th>说明</th></tr></thead><tbody>
    ${p.rows.filter((r) => r.action !== 'same').map((r) => `<tr class="${rowCls(r)}"><td>${r.rowNo}</td><td>${act[r.action]}</td><td>${esc(r.data.name)}</td><td>${esc(r.data.code)}</td><td>${esc(r.data.grade)}</td><td>${esc(r.data.status)}</td><td>${esc([...r.errors, ...r.warnings, ...r.diff].join('；'))}</td></tr>`).join('') || '<tr><td colspan="7">全部不变</td></tr>'}
    </tbody></table></div>
    <button class="btn primary mt" id="btn-do-import" data-act="do-import" ${p.summary.add + p.summary.update ? '' : 'disabled'}>确认导入 ${p.summary.add + p.summary.update} 人</button>
    <button class="btn" data-act="cancel-import">取消</button>` : ''}
  ${lastImportUndo()}</div>`;
};
function lastImportUndo() {
  const op = [...S.store.ops].reverse().find((o) => o.kind === 'importRoster');
  if (!op || S.store.ops.some((o) => o.kind === 'undo' && o.ref === op.id)) return '';
  return `<div class="section">上次导入</div><div class="list"><div class="row"><div class="main"><div class="t wrap">${esc(op.summary)}</div><div class="s">${fmtTime(op.at)}</div></div><button class="btn small" id="btn-undo-import" data-act="undo-import" data-id="${op.id}">整批撤销</button></div></div>`;
}
async function onImportFile(file) {
  try {
    const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()));
    const sh = readRosterSheet(XLSX, wb);
    S.importPlan = { ...planImport(sh, st()), fileName: file.name };
  } catch (e) { S.importPlan = null; alertDlg('读不了这个文件', e.message); }
  render();
}

// ---------- 月底导出 ----------
VIEWS.export = ({ month = prevMonthDefault(), campus = '', format = st().settings.exportFormat ?? 'weiduo' }) => {
  const { from, to } = monthRange(month);
  const { issues, stats } = precheck(st(), { from, to, campus });
  const L = layoutSheet(stats, format);
  const aoa = L.aoa.slice(L.headerRow);
  const cols = format === 'jing' ? [0, 1, 2, 3, 4, 7] : [0, 1, 2, 3, 4, 5];
  const campuses = [...new Set(Object.values(st().courses).map((c) => c.campus).filter(Boolean))];
  const locked = S.store.isLocked(month);
  return `${top('月底导出', { back: 'menu' })}<div class="page" id="export">
  <div class="grid2"><label class="field"><span>月份</span><input class="input" type="month" id="e-month" value="${month}"></label>
  <label class="field"><span>校区</span><select class="input" id="e-campus"><option value="">全部校区</option>${campuses.map((x) => `<option ${x === campus ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select></label></div>
  <label class="field"><span>表格格式</span><select class="input" id="e-format">${Object.entries(EXPORT_FORMATS).map(([k, v]) => `<option value="${k}" ${k === format ? 'selected' : ''}>${esc(v)}</option>`).join('')}</select></label>
  <div class="card"><div class="grid3"><div><div class="k">签课</div><div class="v">${stats.sessions.length} 次</div></div><div><div class="k">学生行</div><div class="v">${stats.studentRows}</div></div><div><div class="k">合计</div><div class="v">${fmtJie(stats.totalJie)}</div></div></div>
  <p class="hint">${locked ? '🔒 本月已锁账' : '导出后会锁定本月，再改要输密码'}</p></div>
  <div class="section">导出前自检</div>
  <div id="precheck">${issues.length ? `<div class="list">${issues.map((i) => `<button class="row warn" data-go="session" data-id="${i.sessionId}"><div class="main"><div class="t wrap"><b>${esc(i.type)}</b>　${esc(i.text)}</div><div class="s">${esc(i.where)}</div></div><span class="chev">›</span></button>`).join('')}</div>`
    : '<div class="card ok-card" id="precheck-ok">✓ 没发现问题：无 0 人到课、无节数异常、学生都在名单里</div>'}</div>
  <div class="section">预览（${esc(EXPORT_FORMATS[format])}；单价、总价留空${L.title ? '；第 1 行标题「' + esc(L.title) + '」' : ''}）</div>
  <div class="tbl-wrap"><table id="export-preview"><thead><tr>${cols.map((i) => `<th>${esc(aoa[0][i])}</th>`).join('')}</tr></thead><tbody>
  ${aoa.slice(1).map((r) => `<tr>${cols.map((i) => `<td>${esc(r[i])}</td>`).join('')}</tr>`).join('') || '<tr><td colspan="6">这个月没有签课</td></tr>'}</tbody></table></div>
  </div><div class="fab-wrap"><button class="btn primary big" id="btn-export" data-act="do-export" ${stats.sessions.length ? '' : 'disabled'}>导出 Excel${issues.length ? `（有 ${issues.length} 条提醒）` : ''}</button></div>`;
};
const prevMonthDefault = () => (new Date().getDate() <= 10 ? prevMonth() : thisMonth());
async function doExport() {
  const month = S.view.month ?? prevMonthDefault(), campus = S.view.campus ?? '', format = S.view.format ?? st().settings.exportFormat ?? 'weiduo';
  const { from, to } = monthRange(month);
  const { issues, stats } = precheck(st(), { from, to, campus });
  if (issues.length && !(await confirmDlg(`自检有 ${issues.length} 条提醒`, issues.slice(0, 6).map((i) => `· ${i.where}：${i.text}`).join('\n') + (issues.length > 6 ? '\n……' : '') + '\n\n建议先点开逐条核对。仍要导出吗？', '仍导出'))) return;
  const org = st().settings.org ?? '';
  const name = format === 'jing' ? `${org}${campus}${month.replace('-', '')}课时费表.xlsx` : `${org}${campus}${month.slice(0, 4)}年${+month.slice(5)}月课时统计.xlsx`;
  await saveFile(writeMonthlyWorkbook(XLSX, st(), stats, { format }), name, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  if (!S.store.isLocked(month)) { S.store.lockMonth(month, '导出课时统计表后自动锁定'); await flush(); toast(`已导出，${+month.slice(5)}月已锁账`); }
  else toast('已导出');
  render();
}

// ---------- 备份与恢复 ----------
VIEWS.backup = () => `${top('备份与恢复', { back: 'menu' })}<div class="page" id="backup">
  <div class="card"><b>一键备份</b><p class="hint">数据只在这台手机里。删除主屏图标、清除 Safari 数据或换手机都会丢数据。每周至少备份一次，存到「文件」→ iCloud 云盘，或发给自己的微信。</p>
  <button class="btn primary" id="btn-backup" data-act="backup">导出完整备份（JSON）</button>
  <button class="btn" data-act="backup-xlsx">另存一份全部数据 Excel（给人看）</button>
  <p class="hint">${S.lastExport ? '上次备份：' + fmtTime(S.lastExport) : '还没备份过'}</p></div>
  <div class="card"><b>恢复</b><p class="hint">选之前导出的 .json 备份文件。恢复会用备份<b>替换</b>手机上现有的全部数据（替换前会自动在本机留一份快照）。</p>
  <label class="btn danger" id="lbl-restore">从备份恢复<input type="file" id="f-restore" accept=".json,application/json" class="hidden"></label></div>
  <div class="section">本机自动快照（换手机无效，仅防误操作）</div><div class="list" id="snaps"><div class="empty">读取中…</div></div></div>`;
AFTER.backup = async () => {
  const snaps = await db.listSnapshots();
  const el = document.getElementById('snaps'); if (!el) return;
  el.innerHTML = snaps.map((s) => `<div class="row"><div class="main"><div class="t">${esc(s.reason)}</div><div class="s">${fmtTime(s.createdAt)} · ${s.counts.students} 名学生 · ${s.counts.sessions} 次签课</div></div><button class="btn small" data-act="restore-snap" data-id="${s.id}">恢复</button></div>`).join('') || '<div class="empty">暂无</div>';
};
async function doBackup() {
  await flush();
  await saveFile(await makeBackup(S.store, APP_VERSION), `签课备份-${dayOf(new Date())}.json`, 'application/json');
  S.lastExport = new Date().toISOString(); await db.kvSet('lastBackup', S.lastExport);
  toast('备份已导出，请存到「文件」或 iCloud'); render();
}
async function replaceAll(state, ops, label) {
  await db.addSnapshot('恢复前自动快照', S.store.state, S.store.ops);
  migrate(state); S.store = new Store({ state, ops }); attach();
  await db.saveData(state, ops); S.integrity = S.store.verifyIntegrity();
  toast(label); go('home');
}
async function onRestoreFile(file) {
  try {
    const { meta, state, ops } = await readBackup(new Uint8Array(await file.arrayBuffer()));
    const ok = await confirmDlg('用这份备份替换现有数据？', `备份时间：${fmtTime(meta.exportedAt)}\n学生 ${meta.counts.students} 人，课程 ${meta.counts.courses} 门，签课 ${meta.counts.sessions} 次。\n\n手机上现有：学生 ${Object.keys(st().students).length} 人，签课 ${Object.keys(st().sessions).length} 次。`, '替换并恢复', true);
    if (!ok) return;
    if (!(await pinCheck('输入密码确认恢复'))) return;
    await replaceAll(state, ops, '已从备份恢复');
  } catch (e) { alertDlg('恢复失败', e.message); }
}

// ---------- V2 查询 ----------
VIEWS.byTeacher = ({ month = thisMonth(), id = '' }) => {
  const { from, to } = monthRange(month);
  const list = teacherSummary(st(), { from, to });
  return `${top('按老师查课时', { back: 'menu' })}<div class="page" id="by-teacher"><label class="field"><span>月份</span><input class="input" type="month" id="q-month" data-v="byTeacher" value="${month}"></label>
  <p class="hint">授课节数：每次课按这次课的节数计；学生消课：到课学生节数之和。代课的课时算代课老师。</p>
  ${list.length ? list.map((x) => `<div class="card"><button class="row plain" data-act="q-open" data-k="${esc(x.teacher)}"><div class="main"><div class="t-strong">${esc(x.teacher)}</div>
    <div class="s">上课 ${x.sessions} 次 · 授课 ${fmtJie(x.teachJie)} · 学生消课 ${fmtJie(x.studentJie)}${x.substitute ? ` · 其中代课 ${x.substitute} 次` : ''}</div></div><span class="chev">${S.qOpen === x.teacher ? '收起' : '明细'}</span></button>
    ${S.qOpen === x.teacher ? `<div class="list">${x.items.map((i) => `<button class="row" data-go="session" data-id="${i.sessionId}"><div class="main"><div class="t">${fmtDay(i.date)} ${esc(i.course)}</div><div class="s">到课 ${i.present} 人 · ${fmtJie(i.teachJie)}${i.tag ? ' · ' + i.tag : ''}${i.substitute ? ' · 代课' : ''}</div></div></button>`).join('')}</div>` : ''}</div>`).join('') : '<div class="empty">这个月没有签课</div>'}</div>`;
};
VIEWS.byStudent = ({ month = thisMonth(), id = '' }) => {
  const { from, to } = monthRange(month);
  const list = studentSummary(st(), { from, to });
  const red = new Set(absenceStreaks(st()).map((x) => x.studentId));
  const open = id || S.qOpenStu;
  return `${top('按学生查次数', { back: 'menu' })}<div class="page" id="by-student"><label class="field"><span>月份</span><input class="input" type="month" id="q-month" data-v="byStudent" value="${month}"></label>
  <p class="hint">红色：同一门课连续 2 次及以上请假或缺勤。</p>
  ${list.length ? `<div class="list">${list.map((x) => `<button class="row ${red.has(x.studentId) ? 'err' : ''}" data-act="q-open-stu" data-k="${x.studentId}"><div class="main"><div class="t">${esc(x.name)}</div>
    <div class="s">到课 ${x.present} 次（${fmtJie(x.jie)}）· 请假 ${x.leave} 次 · 缺勤 ${x.absent} 次</div></div><span class="chev">${open === x.studentId ? '收起' : '明细'}</span></button>
    ${open === x.studentId ? x.items.map((i) => `<button class="row sub" data-go="session" data-id="${i.sessionId}"><div class="main"><div class="t">${fmtDay(i.date)} ${esc(i.course)}</div><div class="s">${i.status === ATT.present ? fmtJie(i.jie) : i.status}${i.tag ? ' · ' + i.tag : ''}</div></div></button>`).join('') : ''}`).join('')}</div>` : '<div class="empty">这个月没有签课</div>'}</div>`;
};

// ---------- 操作记录 / 设置 / 教程 ----------
VIEWS.log = () => `${top('操作记录', { back: 'menu' })}<div class="page"><p class="hint">所有改动都在这里，最新的在上面。</p><div class="list">${[...S.store.ops].reverse().slice(0, 500).map((o) => `<div class="row"><div class="main"><div class="t wrap">${esc(o.summary)}</div><div class="s">${fmtTime(o.at)} · ${esc(KIND_LABEL[o.kind] ?? o.kind)}</div></div></div>`).join('') || '<div class="empty">暂无</div>'}</div></div>`;
VIEWS.settings = () => {
  const lm = st().settings.lockedMonths ?? [];
  return `${top('设置', { back: 'menu' })}<div class="page"><div class="list">
  <button class="row" data-act="set-campus"><div class="main"><div class="t">本机默认校区</div><div class="s">${esc(st().settings.campus || '未设置')}（新建课程时自动带入）</div></div></button>
  <button class="row" data-act="set-org"><div class="main"><div class="t">机构名称（导出文件名前缀）</div><div class="s">${esc(st().settings.org || '未设置')}，例如填「维多」导出为「维多2026年11月课时统计.xlsx」</div></div></button>
  <button class="row" data-act="change-pin"><div class="main"><div class="t">修改密码</div></div></button></div>
  <div class="section">已锁账的月份</div><div class="list">${lm.map((m) => `<div class="row"><div class="main"><div class="t">${m}</div></div><button class="btn small" data-act="unlock-month" data-m="${m}">解锁</button></div>`).join('') || '<div class="empty">暂无</div>'}</div>
  <p class="hint">存储永久保留：${S.persisted === true ? '已获准' : S.persisted === false ? '未获准（请定期备份）' : '未知'}。页面禁止任何外部连接，数据不离开手机。</p></div>`;
};
VIEWS.help = () => `${top('使用教程', { back: S.unlocked ? 'menu' : 'home' })}<div class="page help" id="help">
  <p>例子：<b>高二数学班，11月8日，到课 5 人、1 人请假，每人 1 节</b>。名字都是虚构的。</p>
  <h3>1. 加到桌面（只做一次）</h3><p>用 <b>Safari</b> 打开链接（在微信里先点右上角「···」→「在 Safari 中打开」）→ 点底部「分享」→「添加到主屏幕」→ 以后从桌面图标「签课」打开，第一次设 6 位密码。</p>
  <h3>2. 导入名单</h3><p>把「学生总名单」Excel 通过微信或 AirDrop 发到手机，点开后「存储到文件」。在 App 里：菜单 →「导入名单」→「选择 Excel 文件」→ 看预览（红色行有错不导入）→「确认导入」。导错了可以「整批撤销」。</p>
  <h3>3. 建课程模板（每个班建一次）</h3><p>菜单 →「课程模板」→「新建」：年级 高二、班课、科目 数学、教师、课程名称「高二数学班」、默认 1 节 → 勾选学生甲～学生己 6 人 →「保存课程」。一对一每个学生每科建一门。</p>
  <h3>4. 签课</h3><p>首页点「高二数学班」→ 日期默认今天，改成 11月8日 → 默认全部到课，把学生己点成「请假」（没来也没请假的点「缺勤」）→ 底部显示「到课 5 人 · 请假 1 · 共 5 节」→「完成签课」。某人只上半节就点「−」改成 0.5 节（=1课时）。同一个班同一天再签会提醒，防止重复。</p>
  <h3>5. 改错</h3><p>签完就锁定。要改：点那次签课 →「修改（需填原因）」→ 改好后填写原因 → 保存。每次改动都留在「留痕」和「操作记录」里。</p>
  <h3>6. 月底导出</h3><p>菜单 →「月底导出课时统计表」→ 选月份 → 看「导出前自检」（0 人到课、节数异常、学生不在名单里会逐条列出，点开核对）→「导出 Excel」。表格式与原「课时统计」一样：本例学生甲一行写「1节2课时」，课次写「5人1次」；消课单价、总价留空。导出后该月自动<b>锁账</b>，再改要输密码。</p>
  <h3>V2 新功能</h3><p><b>补课/调课</b>：签课页「这次课的标记」点「补课」或「调课」，导出时备注写「补课1次」。<b>代课</b>：签课页把「授课老师」改成代课老师，导出时这次课时记在代课老师名下，备注写原任课老师。<b>按老师查课时</b>、<b>按学生查次数</b>：菜单里选月份查看，点开看明细。<b>连续缺课提醒</b>：同一门课连续 2 次请假或缺勤，首页和学生查询里标红。</p>
  <h3>7. 备份（每周一次）</h3><p>菜单 →「备份与恢复」→「导出完整备份」→ 存到「文件」→ iCloud 云盘，或发给自己的微信。换手机时在新手机上「从备份恢复」。不要删除桌面图标、不要清除 Safari 数据。</p>
  <p class="hint">数据只存在这台手机，不上传，不同步到别的手机。</p></div>`;

// ---------- 文件保存（iPhone 走分享面板，可存到「文件」或发微信） ----------
async function saveFile(bytes, name, mime) {
  const file = new File([bytes], name, { type: mime });
  if (navigator.canShare?.({ files: [file] })) {
    try { await navigator.share({ files: [file], title: name }); return; } catch (e) { if (e?.name === 'AbortError') return; }
  }
  const url = URL.createObjectURL(file);
  const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

// ---------- 事件 ----------
document.addEventListener('click', async (e) => {
  const g = e.target.closest('[data-go]');
  if (g) { e.preventDefault(); return go(g.dataset.go, g.dataset.id ? { id: g.dataset.id } : {}); }
  const b = e.target.closest('[data-act]'); if (!b) return;
  const act = b.dataset.act, id = b.dataset.id;
  try {
    switch (act) {
      case 'pick-course': {
        const cs = S.store.courses();
        if (!cs.length) return alertDlg('还没有课程', '先到 菜单 →「课程模板」建一门课程。');
        if (cs.length === 1) { S.draft = newRollDraft(cs[0]); return go('roll'); }
        const v = await sheet({ title: '选课程', html: `<div class="list">${cs.map((c) => `<button class="row" data-pick="${c.id}"><div class="main"><div class="t">${esc(courseLabel(c))}</div></div></button>`).join('')}</div>`,
          buttons: [{ label: '取消', value: null }], onMount: (w) => w.addEventListener('click', (ev) => { const p = ev.target.closest('[data-pick]'); if (p) { w.remove(); S.draft = newRollDraft(st().courses[p.dataset.pick]); go('roll'); } }) });
        return v;
      }
      case 'start-roll': S.draft = newRollDraft(st().courses[id]); return go('roll');
      case 'set-status': { const r = S.draft.records[+b.dataset.i]; r.status = b.dataset.s; if (r.status === ATT.present && !r.jie) r.jie = st().courses[S.draft.courseId].defaultJie; return keepRoll(); }
      case 'jie': { const r = S.draft.records[+b.dataset.i]; const n = Math.round((r.jie + Number(b.dataset.d)) * 2) / 2; if (n >= 0.5 && n <= 10) r.jie = n; return keepRoll(); }
      case 'save-roll': return saveRoll();
      case 'set-tag': S.draft.tag = b.dataset.t; return keepRoll();
      case 'edit-session': { const s = st().sessions[id]; S.draft = { courseId: s.courseId, date: s.date, note: s.note ?? '', editId: id, tag: s.tag ?? '', subTeacher: s.subTeacher ?? '', records: s.records.map((r) => ({ ...r, jie: r.status === ATT.present ? r.jie : st().courses[s.courseId]?.defaultJie ?? 1 })) }; return go('roll'); }
      case 'del-session': {
        const s = st().sessions[id];
        let pinOk = false;
        if (S.store.isLocked(s.date.slice(0, 7))) { if (!(await pinCheck('该月已锁账，输入密码继续'))) return; pinOk = true; }
        const reason = await promptDlg('删除原因（必填）', { placeholder: '如 重复录入', okLabel: '删除' });
        if (reason == null) return;
        S.store.deleteSession(id, { reason, pinOk }); await flush(); toast('已删除，记录已留痕'); return go('home');
      }
      case 'edit-course': {
        const c = id ? st().courses[id] : null;
        S.draft = { campus: c?.campus ?? st().settings.campus ?? '', grade: c?.grade ?? '', mode: c?.mode ?? '班课', subject: c?.subject ?? '', teacher: c?.teacher ?? '', name: c?.name ?? '', defaultJie: c?.defaultJie ?? 1, studentIds: [...(c?.studentIds ?? [])], _all: false, _q: '' };
        return go('course', { id });
      }
      case 'toggle-stu': {
        readCourseForm();
        const set = S.draft.studentIds; const i = set.indexOf(id);
        if (i >= 0) set.splice(i, 1); else if (S.draft.mode === '一对一') S.draft.studentIds = [id]; else set.push(id);
        return keepScroll();
      }
      case 'course-all': readCourseForm(); S.draft._all = !S.draft._all; return keepScroll();
      case 'save-course': {
        readCourseForm();
        const { _all, _q, ...input } = S.draft;
        if (!isHalfStep(input.defaultJie) || input.defaultJie <= 0) return alertDlg('默认节数不对', '填 0.5 的倍数，如 1、1.5、2');
        if (!input.studentIds.length && !(await confirmDlg('还没选学生', '先保存，稍后再选学生？', '先保存'))) return;
        S.view.id ? S.store.editCourse(S.view.id, input) : S.store.addCourse(input);
        await flush(); toast('课程已保存'); return go('courses');
      }
      case 'archive-course': { const c = st().courses[id]; S.store.editCourse(id, { archived: !c.archived }); await flush(); return go('courses'); }
      case 'edit-stu': return editStudent(id);
      case 'stu-all': return go('students', { all: !S.view.all, q: S.view.q ?? '' });
      case 'dl-template': return saveFile(writeRosterTemplate(XLSX), '学生名单导入模板.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      case 'do-import': {
        const p = S.importPlan;
        const op = S.store.applyImport(p, p.fileName); S.importPlan = null; await flush();
        toast(op?.summary ?? '已导入', op ? { label: '撤销', run: async () => { S.store.undo(op.id); await flush(); toast('已撤销整批导入'); render(); } } : null);
        return render();
      }
      case 'cancel-import': S.importPlan = null; return render();
      case 'undo-import': { if (!(await confirmDlg('撤销整批导入？', '这次导入新增的学生会删除，更新的恢复原样。', '整批撤销', true))) return; S.store.undo(id); await flush(); toast('已撤销'); return render(); }
      case 'do-export': return doExport();
      case 'backup': return doBackup();
      case 'backup-xlsx': return saveFile(writeAllDataWorkbook(XLSX, st()), `签课全部数据-${dayOf(new Date())}.xlsx`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      case 'restore-snap': {
        const s = await db.getSnapshot(id);
        if (!(await confirmDlg('恢复到这份快照？', `${s.reason} ${fmtTime(s.createdAt)}，当前数据会先自动留一份快照。`, '恢复', true))) return;
        if (!(await pinCheck('输入密码确认恢复'))) return;
        return replaceAll(s.state, s.ops, '已恢复到快照');
      }
      case 'set-campus': { const v = await promptDlg('本机默认校区', { value: st().settings.campus ?? '' }); if (v == null) return; S.store.setSettings({ campus: v.trim() }); await flush(); return render(); }
      case 'q-open': S.qOpen = S.qOpen === b.dataset.k ? null : b.dataset.k; return keepScroll();
      case 'q-open-stu': S.qOpenStu = S.qOpenStu === b.dataset.k ? null : b.dataset.k; S.view.id = ''; return keepScroll();
      case 'set-org': { const v = await promptDlg('机构名称', { value: st().settings.org ?? '' }); if (v == null) return; S.store.setSettings({ org: v.trim() }); await flush(); return render(); }
      case 'unlock-month': { if (!(await pinCheck(`解锁 ${b.dataset.m}`))) return; S.store.unlockMonth(b.dataset.m); await flush(); toast('已解锁'); return render(); }
      case 'change-pin': {
        if (!(await pinCheck('先输入当前密码'))) return;
        const p1 = await promptDlg('新密码（6 位数字）', { type: 'password', inputmode: 'numeric' }); if (p1 == null) return;
        if (!isValidPin(p1)) return alertDlg('密码不对', '请输入 6 位数字');
        const p2 = await promptDlg('再输入一次新密码', { type: 'password', inputmode: 'numeric' }); if (p2 !== p1) return alertDlg('两次不一致', '密码没有修改');
        S.pinRec = await makePinRecord(p1); await db.kvSet('pin', S.pinRec); return toast('密码已修改');
      }
    }
  } catch (err) {
    if (err?.code === 'LOCKED') return alertDlg('已锁账', err.message);
    alertDlg('出错了', err?.message ?? String(err));
  }
});
function keepScroll() { const y = window.scrollY; render(); window.scrollTo(0, y); }
function keepRoll() { const d = document.getElementById('r-date'); if (d) S.draft.date = d.value; const n = document.getElementById('r-note'); if (n) S.draft.note = n.value; const t = document.getElementById('r-sub'); if (t) S.draft.subTeacher = t.value.trim(); keepScroll(); }
document.addEventListener('change', (e) => {
  const t = e.target;
  if (t.id === 'f-import' && t.files?.[0]) onImportFile(t.files[0]);
  else if (t.id === 'f-restore' && t.files?.[0]) onRestoreFile(t.files[0]);
  else if (t.id === 's-month' && t.value) go('sessions', { month: t.value });
  else if (t.id === 'q-month' && t.value) go(t.dataset.v, { month: t.value });
  else if (t.id === 'e-month' && t.value) go('export', { month: t.value, campus: S.view.campus ?? '', format: S.view.format });
  else if (t.id === 'e-campus') go('export', { month: S.view.month ?? prevMonthDefault(), campus: t.value, format: S.view.format });
  else if (t.id === 'e-format') { if (st().settings.exportFormat !== t.value) S.store.setSettings({ exportFormat: t.value }); go('export', { month: S.view.month ?? prevMonthDefault(), campus: S.view.campus ?? '', format: t.value }); }
  else if (t.id === 'r-date' && S.draft) { S.draft.date = t.value; keepRoll(); }
  else if (t.matches?.('#c-mode button')) {}
});
document.addEventListener('click', (e) => {
  const m = e.target.closest('#c-mode button'); if (!m) return;
  readCourseForm(); S.draft.mode = m.dataset.mode; if (S.draft.mode === '一对一') S.draft.studentIds = S.draft.studentIds.slice(0, 1); keepScroll();
});
// 实时搜索：中文输入法选字期间（composition）不搜索、不重绘；只重绘结果列表，输入框节点和内容保持不变
let qTimer = null, composing = false;
function runSearch(t) {
  if (t.id === 'c-q' && S.draft) {
    S.draft._q = t.value;
    const el = document.getElementById('pick-list');
    if (el) el.innerHTML = coursePickHtml(S.store.students({ includeInactive: S.draft._all }), t.value, new Set(S.draft.studentIds));
  } else if (t.id === 'stu-q') {
    S.view.q = t.value;
    const el = document.getElementById('stu-list');
    if (el) el.innerHTML = stuListHtml(!!S.view.all, t.value);
  }
}
const isSearch = (t) => t && (t.id === 'c-q' || t.id === 'stu-q');
document.addEventListener('compositionstart', (e) => { if (isSearch(e.target)) { composing = true; clearTimeout(qTimer); } });
document.addEventListener('compositionend', (e) => { if (isSearch(e.target)) { composing = false; const t = e.target; clearTimeout(qTimer); qTimer = setTimeout(() => runSearch(t), 0); } });
document.addEventListener('input', (e) => {
  const t = e.target;
  if (!isSearch(t)) return;
  if (composing || e.isComposing) return;
  clearTimeout(qTimer); qTimer = setTimeout(() => runSearch(t), 150);
});

boot().catch((e) => { $app.innerHTML = `<div class="boot">打开失败：${esc(e?.message ?? e)}<br>请换 Safari 打开，或联系管理员。</div>`; });
