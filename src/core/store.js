// 数据核心：名单、课程（班级）、签课记录；所有改动经 commit 写入操作记录，可撤销
import { clone, deepEqual } from './util.js';
import { isHalfStep } from './jie.js';
import { isDay } from './dates.js';

export const STATUS_LIST = ['在读', '停课', '毕业', '结课'];
export const MODES = ['班课', '一对一'];
export const ATT = { present: '到课', absent: '缺勤', leave: '请假' };
export const KIND_LABEL = {
  addStudent: '新增学生', editStudent: '修改学生', importRoster: '导入名单', addCourse: '新建课程', editCourse: '修改课程',
  archiveCourse: '停用课程', addSession: '签课', editSession: '修改签课', deleteSession: '删除签课', undo: '撤销', settings: '修改设置',
  lockMonth: '锁定月份', unlockMonth: '解锁月份',
};
const COLLS = ['students', 'courses', 'sessions'];
export const SCHEMA = 2;
export const SESSION_TAGS = ['', '补课', '调课'];
export const emptyState = () => ({ schema: SCHEMA, students: {}, courses: {}, sessions: {}, settings: { campus: '', lockedMonths: [] } });

const defaultUuid = () => globalThis.crypto.randomUUID();
export const isPendingCode = (c) => !c || /^待补/.test(String(c).trim());

export class Store {
  constructor({ state = emptyState(), ops = [], now = () => new Date(), uuid = defaultUuid } = {}) {
    this.state = state; this.ops = ops; this.now = now; this.uuid = uuid; this.onCommit = null;
  }
  stamp() { return this.now().toISOString(); }

  /** changes: [{ coll, id, after }]（after=null 表示删除）；记录 before 以便撤销 */
  commit(kind, summary, changes, extra = {}) {
    const real = [];
    for (const c of changes) {
      const before = clone(this.cur(c.coll, c.id));
      const after = c.after == null ? null : clone(c.after);
      if (deepEqual(before, after)) continue;
      real.push({ coll: c.coll, id: c.id, before, after });
    }
    if (!real.length && !extra.allowEmpty) return null;
    for (const c of real) this.put(c.coll, c.id, c.after);
    const op = { id: this.uuid(), at: this.stamp(), kind, summary, changes: real, ...(extra.ref ? { ref: extra.ref } : {}), ...(extra.reason ? { reason: extra.reason } : {}) };
    this.ops.push(op);
    this.onCommit?.(op);
    return op;
  }

  // ---------- 学生 ----------
  students({ includeInactive = false } = {}) {
    return Object.values(this.state.students).filter((s) => includeInactive || s.status === '在读')
      .sort((a, b) => (a.seq ?? 1e9) - (b.seq ?? 1e9) || a.name.localeCompare(b.name, 'zh-CN'));
  }
  normStudent(input, prev = null) {
    const name = String(input.name ?? '').trim();
    if (!name) throw new Error('姓名不能为空');
    const status = String(input.status ?? '在读').trim() || '在读';
    const s = { ...(prev ?? {}), name, code: String(input.code ?? '').trim() || '待补', school: String(input.school ?? '').trim(),
      grade: String(input.grade ?? '').trim(), cohort: String(input.cohort ?? '').trim(), status, manager: String(input.manager ?? '').trim(),
      note: String(input.note ?? prev?.note ?? '').trim() };
    if (input.seq != null && input.seq !== '') s.seq = Number(input.seq) || prev?.seq;
    if (!isPendingCode(s.code)) {
      const dup = Object.values(this.state.students).find((x) => x.code === s.code && x.id !== prev?.id);
      if (dup) throw new Error(`学生编号 ${s.code} 已被「${dup.name}」使用`);
    }
    return s;
  }
  addStudent(input) {
    const s = this.normStudent(input);
    s.id = this.uuid(); s.createdAt = this.stamp(); s.updatedAt = s.createdAt;
    this.commit('addStudent', `新增学生 ${s.name}`, [{ coll: 'students', id: s.id, after: s }]);
    return s;
  }
  editStudent(id, input) {
    const prev = this.state.students[id]; if (!prev) throw new Error('学生不存在');
    const s = this.normStudent({ ...prev, ...input }, prev); s.updatedAt = this.stamp();
    return this.commit('editStudent', `修改学生 ${s.name}`, [{ coll: 'students', id, after: s }]);
  }
  /** 导入名单：plan 来自 roster.planImport；整批一个操作，可一次撤销 */
  applyImport(plan, fileName = '') {
    const changes = [];
    for (const r of plan.rows) {
      if (r.action === 'add') { const id = this.uuid(); changes.push({ coll: 'students', id, after: { ...r.data, id, createdAt: this.stamp(), updatedAt: this.stamp() } }); }
      else if (r.action === 'update') { const prev = this.state.students[r.matchId]; changes.push({ coll: 'students', id: r.matchId, after: { ...prev, ...r.data, id: prev.id, note: prev.note ?? '', updatedAt: this.stamp() } }); }
    }
    const nAdd = plan.rows.filter((r) => r.action === 'add').length, nUpd = plan.rows.filter((r) => r.action === 'update').length;
    return this.commit('importRoster', `导入名单${fileName ? ' ' + fileName : ''}：新增 ${nAdd} 人，更新 ${nUpd} 人`, changes, { allowEmpty: true });
  }

  // ---------- 课程（班级） ----------
  courses({ includeArchived = false } = {}) {
    return Object.values(this.state.courses).filter((c) => includeArchived || !c.archived)
      .sort((a, b) => courseLabel(a).localeCompare(courseLabel(b), 'zh-CN'));
  }
  normCourse(input, prev = null) {
    const c = { ...(prev ?? {}) };
    for (const k of ['campus', 'grade', 'subject', 'teacher', 'name']) c[k] = String(input[k] ?? prev?.[k] ?? '').trim();
    c.mode = input.mode ?? prev?.mode ?? '班课';
    if (!MODES.includes(c.mode)) throw new Error('班制只能是 班课 或 一对一');
    if (!c.teacher) throw new Error('请填写教师');
    if (!c.subject) throw new Error('请填写科目');
    c.defaultJie = Number(input.defaultJie ?? prev?.defaultJie ?? 1);
    if (!isHalfStep(c.defaultJie) || c.defaultJie <= 0 || c.defaultJie > 10) throw new Error('默认节数须为 0.5 的倍数（0.5～10）');
    const ids = [...new Set(input.studentIds ?? prev?.studentIds ?? [])];
    for (const sid of ids) if (!this.state.students[sid]) throw new Error('课程里有不存在的学生');
    if (c.mode === '一对一' && ids.length > 1) throw new Error('一对一课程只能有 1 个学生（同一学生不同科目请分别建课程）');
    c.studentIds = ids;
    c.archived = !!(input.archived ?? prev?.archived ?? false);
    return c;
  }
  addCourse(input) {
    const c = this.normCourse(input); c.id = this.uuid(); c.createdAt = this.stamp(); c.updatedAt = c.createdAt;
    this.commit('addCourse', `新建课程 ${courseLabel(c)}`, [{ coll: 'courses', id: c.id, after: c }]);
    return c;
  }
  editCourse(id, input) {
    const prev = this.state.courses[id]; if (!prev) throw new Error('课程不存在');
    const c = this.normCourse(input, prev); c.updatedAt = this.stamp();
    const kind = input.archived === true && !prev.archived ? 'archiveCourse' : 'editCourse';
    return this.commit(kind, `${KIND_LABEL[kind]} ${courseLabel(c)}`, [{ coll: 'courses', id, after: c }]);
  }

  // ---------- 签课（一次上课） ----------
  sessions({ from = '0000-00-00', to = '9999-99-99', courseId = null } = {}) {
    return Object.values(this.state.sessions).filter((s) => s.date >= from && s.date <= to && (!courseId || s.courseId === courseId))
      .sort((a, b) => a.date.localeCompare(b.date) || a.createdAt.localeCompare(b.createdAt));
  }
  normSession(input, prev = null) {
    const courseId = input.courseId ?? prev?.courseId;
    const course = this.state.courses[courseId]; if (!course) throw new Error('请选择课程');
    const date = input.date ?? prev?.date;
    if (!isDay(date)) throw new Error('日期不对');
    const seen = new Set();
    const records = (input.records ?? prev?.records ?? []).map((r) => {
      if (!this.state.students[r.studentId]) throw new Error('签课里有不存在的学生');
      if (seen.has(r.studentId)) throw new Error('同一学生重复'); seen.add(r.studentId);
      const status = Object.values(ATT).includes(r.status) ? r.status : ATT.present;
      const jie = status === ATT.present ? Number(r.jie) : 0;
      if (!isHalfStep(jie) || jie > 10) throw new Error(`${this.state.students[r.studentId].name} 的节数须为 0.5 的倍数`);
      if (status === ATT.present && jie === 0) throw new Error(`${this.state.students[r.studentId].name} 到课但节数为 0，请改为缺勤或填写节数`);
      return { studentId: r.studentId, status, jie };
    });
    if (!records.length) throw new Error('没有学生');
    const tag = input.tag ?? prev?.tag ?? '';
    if (!SESSION_TAGS.includes(tag)) throw new Error('标记只能是 补课 或 调课');
    let subTeacher = String(input.subTeacher ?? prev?.subTeacher ?? '').trim();
    if (subTeacher === course.teacher) subTeacher = '';
    return { ...(prev ?? {}), courseId, date, records, note: String(input.note ?? prev?.note ?? '').trim(), tag, subTeacher };
  }
  // ---------- 月度锁账 ----------
  isLocked(month) { return (this.state.settings.lockedMonths ?? []).includes(month); }
  /** 锁定月份内的增删改须先输 PIN（pinOk=true） */
  guardMonth(date, pinOk) { const m = String(date).slice(0, 7); if (this.isLocked(m) && !pinOk) { const e = new Error(`${+m.slice(5)}月已锁账，修改需要输入密码`); e.code = 'LOCKED'; e.month = m; throw e; } }
  lockMonth(month, why = '导出后锁定') {
    if (this.isLocked(month)) return null;
    const lm = [...(this.state.settings.lockedMonths ?? []), month].sort();
    return this.commit('lockMonth', `锁定 ${month}（${why}）`, [{ coll: 'meta', id: 'settings', after: { ...this.state.settings, lockedMonths: lm } }]);
  }
  unlockMonth(month) {
    if (!this.isLocked(month)) return null;
    return this.commit('unlockMonth', `解锁 ${month}`, [{ coll: 'meta', id: 'settings', after: { ...this.state.settings, lockedMonths: this.state.settings.lockedMonths.filter((m) => m !== month) } }]);
  }

  addSession(input, { pinOk = false } = {}) {
    this.guardMonth(input.date, pinOk);
    const s = this.normSession(input); s.id = this.uuid(); s.createdAt = this.stamp(); s.updatedAt = s.createdAt;
    this.commit('addSession', `签课 ${s.date} ${courseLabel(this.state.courses[s.courseId])}${s.tag ? '（' + s.tag + '）' : ''}${s.subTeacher ? '（代课：' + s.subTeacher + '）' : ''}：到课 ${presentCount(s)} 人`, [{ coll: 'sessions', id: s.id, after: s }]);
    return s;
  }
  /** 签完即锁定：修改必须填原因，原因与改动写进操作记录 */
  editSession(id, input, { reason = '', pinOk = false } = {}) {
    const prev = this.state.sessions[id]; if (!prev) throw new Error('签课记录不存在');
    reason = String(reason).trim(); if (!reason) throw new Error('修改已签的课要填写原因');
    this.guardMonth(prev.date, pinOk); if (input.date) this.guardMonth(input.date, pinOk);
    const s = this.normSession(input, prev); s.updatedAt = this.stamp(); s.edits = (prev.edits ?? 0) + 1;
    const d = diffText(this, prev, s);
    if (d === '无变化' && s.note === prev.note) throw new Error('没有改动');
    return this.commit('editSession', `修改签课 ${s.date} ${courseLabel(this.state.courses[s.courseId])}：${d}；原因：${reason}`, [{ coll: 'sessions', id, after: s }], { reason });
  }
  deleteSession(id, { reason = '', pinOk = false } = {}) {
    const prev = this.state.sessions[id]; if (!prev) throw new Error('签课记录不存在');
    reason = String(reason).trim(); if (!reason) throw new Error('删除已签的课要填写原因');
    this.guardMonth(prev.date, pinOk);
    return this.commit('deleteSession', `删除签课 ${prev.date} ${courseLabel(this.state.courses[prev.courseId])}；原因：${reason}`, [{ coll: 'sessions', id, after: null }], { reason });
  }
  /** 某条签课的全部留痕 */
  historyOf(sessionId) { return this.ops.filter((o) => o.changes.some((c) => c.coll === 'sessions' && c.id === sessionId)); }
  /** 同一课程同一天已有的签课 */
  sessionsOn(courseId, date) { return this.sessions({ from: date, to: date, courseId }); }

  setSettings(patch) {
    return this.commit('settings', '修改设置', [{ coll: 'meta', id: 'settings', after: { ...this.state.settings, ...patch } }]);
  }

  // ---------- 撤销 ----------
  /** 可撤销：最近的、未被撤销过的非撤销操作，且其后改动未覆盖同一条数据 */
  undoable() {
    const undone = new Set(this.ops.filter((o) => o.kind === 'undo').map((o) => o.ref));
    for (let i = this.ops.length - 1; i >= 0; i--) {
      const op = this.ops[i];
      if (op.kind === 'undo' || op.kind === 'restore' || undone.has(op.id)) continue;
      return op;
    }
    return null;
  }
  canUndo(op) {
    for (const c of op.changes) if (!deepEqual(this.cur(c.coll, c.id), c.after)) return false;
    return true;
  }
  cur(coll, id) { return coll === 'meta' ? this.state[id] ?? null : this.state[coll][id] ?? null; }
  put(coll, id, v) {
    if (coll === 'meta') { this.state[id] = v; return; }
    if (v == null) delete this.state[coll][id]; else this.state[coll][id] = v;
  }
  undo(opId) {
    const op = this.ops.find((o) => o.id === opId); if (!op) throw new Error('找不到该操作');
    if (this.ops.some((o) => o.kind === 'undo' && o.ref === opId)) throw new Error('该操作已撤销过');
    if (!this.canUndo(op)) throw new Error('之后又改动过同一条数据，不能直接撤销');
    if (op.changes.some((c) => c.coll === 'sessions')) throw new Error('签课记录不能撤销，请用「修改」并填写原因');
    return this.commit('undo', `撤销：${op.summary}`, op.changes.map((c) => ({ coll: c.coll, id: c.id, after: c.before })), { ref: opId, allowEmpty: true });
  }

  /** 自检：引用都存在、数值合法 */
  verifyIntegrity() {
    const issues = [];
    for (const c of Object.values(this.state.courses)) for (const sid of c.studentIds) if (!this.state.students[sid]) issues.push(`课程 ${courseLabel(c)} 引用了不存在的学生`);
    for (const s of Object.values(this.state.sessions)) {
      if (!this.state.courses[s.courseId]) issues.push(`签课 ${s.date} 引用了不存在的课程`);
      for (const r of s.records) { if (!this.state.students[r.studentId]) issues.push(`签课 ${s.date} 引用了不存在的学生`); if (!isHalfStep(r.jie)) issues.push(`签课 ${s.date} 节数不合法`); }
    }
    return issues;
  }
}

export const presentCount = (s) => s.records.filter((r) => r.status === ATT.present && r.jie > 0).length;
export const totalJie = (s) => s.records.reduce((a, r) => a + (r.status === ATT.present ? r.jie : 0), 0);
export function courseLabel(c) {
  if (!c) return '（已删除课程）';
  return c.name || [c.campus, c.grade, c.subject, c.mode, c.teacher].filter(Boolean).join(' · ');
}
/** V1 数据（schema 1）升级到 V2：补齐新字段，不删除、不改动任何已有数据 */
export function migrate(state) {
  if (!state || typeof state !== 'object') throw new Error('数据格式不对');
  const from = state.schema ?? 1;
  if (from > SCHEMA) throw new Error('数据来自更新的版本，请先更新 App');
  state.settings ??= { campus: '' };
  state.settings.lockedMonths ??= [];
  for (const k of ['students', 'courses', 'sessions']) state[k] ??= {};
  for (const s of Object.values(state.sessions)) { s.tag ??= ''; s.subTeacher ??= ''; }
  state.schema = SCHEMA;
  return { from, to: SCHEMA, changed: from !== SCHEMA };
}
export const sessionTeacher = (state, s) => s.subTeacher || state.courses[s.courseId]?.teacher || '';
function diffText(store, a, b) {
  const parts = [];
  if (a.date !== b.date) parts.push(`日期 ${a.date}→${b.date}`);
  if ((a.tag ?? '') !== (b.tag ?? '')) parts.push(`标记 ${a.tag || '正常'}→${b.tag || '正常'}`);
  if ((a.subTeacher ?? '') !== (b.subTeacher ?? '')) parts.push(`授课老师 ${a.subTeacher || '原任课'}→${b.subTeacher || '原任课'}`);
  const ma = new Map(a.records.map((r) => [r.studentId, r]));
  for (const r of b.records) {
    const o = ma.get(r.studentId); const nm = store.state.students[r.studentId]?.name ?? '?';
    if (!o) parts.push(`${nm} 加入`);
    else if (o.status !== r.status || o.jie !== r.jie) parts.push(`${nm} ${o.status}${o.jie}节→${r.status}${r.jie}节`);
    ma.delete(r.studentId);
  }
  for (const r of ma.values()) parts.push(`${store.state.students[r.studentId]?.name ?? '?'} 移出`);
  return parts.join('，') || '无变化';
}
const COLL_OK = new Set([...COLLS, 'meta']);
export const isValidColl = (c) => COLL_OK.has(c);
