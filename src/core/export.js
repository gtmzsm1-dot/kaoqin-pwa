// 月度导出：Sheet1 与现有「课时统计」同版式；另附「明细」页（一行 = 一个学生的一次上课，全数值列、不合并）
import { ATT, courseLabel, STATUS_LIST } from './store.js';
import { fmtJie, keshi, roundHalf } from './jie.js';

export const SHEET1_HEADER = (label) => ['教师姓名', '所带学生', '年级', '班型', label, '课次', '消课单价', '消课总价', '课时费单价', '教师课时费', '合计', '备注', '人工核对'];
export const DETAIL_HEADER = ['日期', '校区', '教师', '年级', '班型', '科目', '课程', '学生', '学生编号', '出勤', '节数', '课时数', '备注'];
const zh = (a, b) => String(a).localeCompare(String(b), 'zh-CN');

export function consumeLabel(from, to) {
  const m1 = +from.slice(5, 7), m2 = +to.slice(5, 7);
  return from.slice(0, 7) === to.slice(0, 7) ? `${m1}月消课量` : `${from.slice(0, 4) !== to.slice(0, 4) ? from.slice(0, 4) + '年' : ''}${m1}月-${m2}月消课量`;
}

/** 课次：每次课的到课人数分组 → 「N人M次」，人数多的在前 */
export function sessionHistogram(sessions) {
  const h = new Map();
  for (const s of sessions) { const n = s.records.filter((r) => r.status === ATT.present && r.jie > 0).length; if (n) h.set(n, (h.get(n) ?? 0) + 1); }
  return [...h.entries()].sort((a, b) => b[0] - a[0]).map(([n, m]) => `${n}人${m}次`);
}

/** 计算统计：groups = [{ teacher, courses:[{ course, rows:[{student, jie}], hist }] }] */
export function buildStats(state, { from, to, campus = '' }) {
  const sessions = Object.values(state.sessions).filter((s) => s.date >= from && s.date <= to)
    .filter((s) => !campus || (state.courses[s.courseId]?.campus ?? '') === campus)
    .sort((a, b) => a.date.localeCompare(b.date) || a.createdAt.localeCompare(b.createdAt));
  const byCourse = new Map();
  for (const s of sessions) { if (!byCourse.has(s.courseId)) byCourse.set(s.courseId, []); byCourse.get(s.courseId).push(s); }
  const teachers = new Map();
  for (const [cid, list] of byCourse) {
    const course = state.courses[cid]; if (!course) continue;
    const jie = new Map();
    for (const s of list) for (const r of s.records) if (r.status === ATT.present && r.jie > 0) jie.set(r.studentId, (jie.get(r.studentId) ?? 0) + r.jie);
    if (!jie.size) continue;
    const order = [...course.studentIds.filter((id) => jie.has(id)), ...[...jie.keys()].filter((id) => !course.studentIds.includes(id))
      .sort((a, b) => zh(state.students[a]?.name ?? '', state.students[b]?.name ?? ''))];
    const rows = order.map((id) => ({ student: state.students[id], jie: roundHalf(jie.get(id)) }));
    const t = course.teacher || '（未填教师）';
    if (!teachers.has(t)) teachers.set(t, []);
    teachers.get(t).push({ course, rows, hist: sessionHistogram(list), sessions: list.length });
  }
  const groups = [...teachers.entries()].sort((a, b) => zh(a[0], b[0]))
    .map(([teacher, courses]) => ({ teacher, courses: courses.sort((a, b) => zh(a.course.grade, b.course.grade) || zh(a.course.mode, b.course.mode) || zh(courseLabel(a.course), courseLabel(b.course))) }));
  const totalJie = groups.reduce((a, g) => a + g.courses.reduce((b, c) => b + c.rows.reduce((x, r) => x + r.jie, 0), 0), 0);
  return { from, to, campus, sessions, groups, totalJie, studentRows: groups.reduce((a, g) => a + g.courses.reduce((b, c) => b + c.rows.length, 0), 0) };
}

/** Sheet1 的二维数组与合并区域 */
export function sheet1Aoa(stats) {
  const aoa = [SHEET1_HEADER(consumeLabel(stats.from, stats.to))];
  const merges = [];
  for (const g of stats.groups) {
    const tStart = aoa.length;
    for (const c of g.courses) {
      const cStart = aoa.length;
      c.rows.forEach((r, i) => {
        aoa.push([aoa.length === tStart ? g.teacher : '', r.student?.name ?? '（已删除学生）', i === 0 ? c.course.grade : '', i === 0 ? c.course.mode : '',
          fmtJie(r.jie), c.hist[i] ?? '', '', '', '', '', '', '', '']);
      });
      if (c.rows.length > 1) for (const col of [2, 3]) merges.push({ s: { r: cStart, c: col }, e: { r: aoa.length - 1, c: col } });
    }
    if (aoa.length - tStart > 1) merges.push({ s: { r: tStart, c: 0 }, e: { r: aoa.length - 1, c: 0 } });
  }
  return { aoa, merges };
}

export function detailAoa(state, stats) {
  const aoa = [DETAIL_HEADER];
  for (const s of stats.sessions) {
    const c = state.courses[s.courseId]; if (!c) continue;
    for (const r of s.records) {
      const st = state.students[r.studentId];
      aoa.push([s.date, c.campus, c.teacher, c.grade, c.mode, c.subject, courseLabel(c), st?.name ?? '（已删除学生）', st?.code ?? '', r.status, r.jie, keshi(r.jie), s.note ?? '']);
    }
  }
  return aoa;
}

export function writeMonthlyWorkbook(XLSX, state, stats) {
  const wb = XLSX.utils.book_new();
  const { aoa, merges } = sheet1Aoa(stats);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = merges;
  ws['!cols'] = [10.17, 10.17, 7.33, 7.33, 11.33, 8.33, 10.17, 10.17, 12.83, 12.83, 8.33, 16.17, 23.17].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const det = XLSX.utils.aoa_to_sheet(detailAoa(state, stats));
  for (let r = 1; ; r++) { const a = det[XLSX.utils.encode_cell({ r, c: 0 })]; if (!a) break; a.t = 'd'; a.v = new Date(a.v + 'T00:00:00Z'); a.z = 'yyyy-mm-dd'; delete a.w; }
  det['!cols'] = [11, 8, 10, 7, 7, 7, 26, 10, 13, 6, 6, 7, 16].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(wb, det, '明细');
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellDates: true });
}

/** 全部数据 Excel（随完整备份一起导出，便于人工查看；恢复以 JSON 为准） */
export function writeAllDataWorkbook(XLSX, state) {
  const wb = XLSX.utils.book_new();
  const stu = [['序号', '姓名', '学生编号', '学校', '年级', '届别', '状态', '负责学管', '备注']];
  for (const s of Object.values(state.students).sort((a, b) => (STATUS_LIST.indexOf(a.status) - STATUS_LIST.indexOf(b.status)) || (a.seq ?? 1e9) - (b.seq ?? 1e9) || zh(a.name, b.name)))
    stu.push([s.seq ?? '', s.name, s.code, s.school, s.grade, s.cohort, s.status, s.manager, s.note ?? '']);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(stu), '学生名单');
  const crs = [['课程', '校区', '年级', '班制', '科目', '教师', '默认节数', '学生人数', '学生', '状态']];
  for (const c of Object.values(state.courses).sort((a, b) => zh(courseLabel(a), courseLabel(b))))
    crs.push([courseLabel(c), c.campus, c.grade, c.mode, c.subject, c.teacher, c.defaultJie, c.studentIds.length,
      c.studentIds.map((id) => state.students[id]?.name ?? '').join('、'), c.archived ? '已停用' : '使用中']);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(crs), '课程');
  const all = buildStats(state, { from: '0000-01-01', to: '9999-12-31' });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(detailAoa(state, all)), '签课明细');
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}

/** 导出前自检：逐条列出 0 人到课、节数异常、学生不在名单里、同班同天多次 */
export function precheck(state, { from, to, campus = '' }) {
  const issues = [];
  const stats = buildStats(state, { from, to, campus });
  const seen = new Map();
  for (const s of stats.sessions) {
    const c = state.courses[s.courseId];
    const where = `${s.date} ${courseLabel(c)}`;
    if (!c) { issues.push({ type: '课程不存在', where, sessionId: s.id, text: '这次签课对应的课程已不存在' }); continue; }
    const key = s.courseId + '|' + s.date;
    if (seen.has(key)) issues.push({ type: '同班同天多次', where, sessionId: s.id, text: '同一课程同一天签了不止一次，请确认不是重复' });
    else seen.set(key, s.id);
    const present = s.records.filter((r) => r.status === ATT.present && r.jie > 0);
    if (!present.length) issues.push({ type: '0 人到课', where, sessionId: s.id, text: '这次课没有人到课（全部请假/缺勤）' });
    const limit = Math.max(4, (c.defaultJie ?? 1) * 2);
    for (const r of s.records) {
      const st = state.students[r.studentId];
      const nm = st?.name ?? '（已删除学生）';
      if (r.status === ATT.present && (r.jie > limit || r.jie <= 0 || Math.abs(r.jie * 2 - Math.round(r.jie * 2)) > 1e-9))
        issues.push({ type: '节数异常', where, sessionId: s.id, text: `${nm} 记了 ${r.jie} 节（课程默认 ${c.defaultJie} 节）` });
      if (r.status !== ATT.present) continue;
      if (!st) issues.push({ type: '学生不在名单里', where, sessionId: s.id, text: `${nm}：名单里已没有这个学生` });
      else if (st.status !== '在读') issues.push({ type: '学生不在名单里', where, sessionId: s.id, text: `${nm}：名单状态是「${st.status}」，不是在读` });
      else if (!c.studentIds.includes(r.studentId)) issues.push({ type: '学生不在名单里', where, sessionId: s.id, text: `${nm}：不在这门课程的学生名单里` });
    }
  }
  return { issues, stats };
}
