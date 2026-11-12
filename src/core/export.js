// 月度导出：只有 Sheet1 一页，与现有「课时统计」同版式。明细（一行 = 一个学生的一次上课）只在数据备份 Excel 里
import { ATT, courseLabel, STATUS_LIST, sessionTeacher } from './store.js';
import { fmtJie, keshi, roundHalf } from './jie.js';
import { writeStyledXlsx } from './xlsxw.js';

export const SHEET1_HEADER = (label) => ['教师姓名', '所带学生', '年级', '班型', label, '课次', '消课单价', '消课总价', '课时费单价', '教师课时费', '合计', '备注', '人工核对'];
export const DETAIL_HEADER = ['日期', '校区', '教师', '年级', '班型', '科目', '课程', '学生', '学生编号', '出勤', '节数', '课时数', '标记', '原任课教师', '备注'];
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
  // 代课：课时记在这次实际授课的老师名下，所以按「课程 + 授课老师」分组
  const byKey = new Map();
  for (const s of sessions) { const k = s.courseId + '|' + sessionTeacher(state, s); if (!byKey.has(k)) byKey.set(k, []); byKey.get(k).push(s); }
  const teachers = new Map();
  for (const [key, list] of byKey) {
    const cid = key.split('|')[0];
    const course = state.courses[cid]; if (!course) continue;
    const t = key.slice(cid.length + 1) || '（未填教师）';
    const jie = new Map(), tags = new Map();
    for (const s of list) for (const r of s.records) if (r.status === ATT.present && r.jie > 0) {
      jie.set(r.studentId, (jie.get(r.studentId) ?? 0) + r.jie);
      if (s.tag) { const m = tags.get(r.studentId) ?? {}; m[s.tag] = (m[s.tag] ?? 0) + 1; tags.set(r.studentId, m); }
    }
    if (!jie.size) continue;
    const order = [...course.studentIds.filter((id) => jie.has(id)), ...[...jie.keys()].filter((id) => !course.studentIds.includes(id))
      .sort((a, b) => zh(state.students[a]?.name ?? '', state.students[b]?.name ?? ''))];
    const sub = t !== course.teacher ? `代课（原任课：${course.teacher}）` : '';
    const rows = order.map((id) => ({ student: state.students[id], jie: roundHalf(jie.get(id)),
      note: [sub, ...Object.entries(tags.get(id) ?? {}).map(([k, n]) => `${k}${n}次`)].filter(Boolean).join('；') }));
    if (!teachers.has(t)) teachers.set(t, []);
    teachers.get(t).push({ course, rows, hist: sessionHistogram(list), sessions: list.length, first: list[0].date + list[0].createdAt, substitute: !!sub });
  }
  const groups = [...teachers.entries()].sort((a, b) => zh(a[0], b[0]))
    .map(([teacher, courses]) => ({ teacher, courses: courses.sort((a, b) => (a.course.mode === '一对一') - (b.course.mode === '一对一') || a.first.localeCompare(b.first) || zh(courseLabel(a.course), courseLabel(b.course))) }));
  const totalJie = groups.reduce((a, g) => a + g.courses.reduce((b, c) => b + c.rows.reduce((x, r) => x + r.jie, 0), 0), 0);
  return { from, to, campus, sessions, groups, totalJie, studentRows: groups.reduce((a, g) => a + g.courses.reduce((b, c) => b + c.rows.length, 0), 0) };
}

export function detailAoa(state, stats) {
  const aoa = [DETAIL_HEADER];
  for (const s of stats.sessions) {
    const c = state.courses[s.courseId]; if (!c) continue;
    for (const r of s.records) {
      const st = state.students[r.studentId];
      aoa.push([s.date, c.campus, sessionTeacher(state, s), c.grade, c.mode, c.subject, courseLabel(c), st?.name ?? '（已删除学生）', st?.code ?? '', r.status, r.jie, keshi(r.jie), s.tag ?? '', s.subTeacher ? c.teacher : '', s.note ?? '']);
    }
  }
  return aoa;
}

// 样式照原表：「课时统计」Sheet1（默认）与「(精)课时费表」
const F11 = { name: '宋体', sz: 11 };
const ST = {
  wdHead: { font: { name: '宋体', sz: 12 }, fill: 'FF4874CB', border: 'thin', h: 'center', v: 'center' },
  wdCell: { font: F11, border: 'thin', h: 'center', v: 'center' },
  wdNum: { font: F11, border: 'thin', h: 'center', v: 'center', numFmt: '#,##0.00' },
  wdCheck: { font: { name: '宋体', sz: 10 }, border: 'thin', h: 'center', v: 'center' },
  jHead: { font: F11, fill: 'FFFFC000', border: 'thin', h: 'center', v: 'center' },
  jCell: { font: F11, border: 'thin', h: 'center', v: 'center' },
  jNum: { font: F11, border: 'thin', h: 'center', v: 'center', numFmt: '#,##0.00;[Red]\\(#,##0.00\\);\\-' },
  dHead: { font: { name: '宋体', sz: 11, b: true }, fill: 'FFD9E1F2', border: 'thin', h: 'center', v: 'center' },
  dCell: { font: F11, border: 'thin', v: 'center' },
  dDate: { font: F11, border: 'thin', v: 'center', numFmt: 'yyyy-mm-dd' },
};
export const EXPORT_FORMATS = { weiduo: '课时统计表（默认，同 10 月课时统计 Sheet1）', jing: '(精)课时费表格式' };
const encodeMerge = (r1, c1, r2, c2) => `${String.fromCharCode(65 + c1)}${r1 + 1}:${String.fromCharCode(65 + c2)}${r2 + 1}`;

/** 二维数组 + 合并（按格式） */
export function layoutSheet(stats, format = 'weiduo') {
  const label = consumeLabel(stats.from, stats.to);
  const jing = format === 'jing';
  const header = jing ? ['教师姓名', '所带学生', '年级', '班型', label, '消课单价', '消课总价', '课次', '课时费单价', '教师课时费', '合计', '备注', '备注'] : SHEET1_HEADER(label);
  const aoa = [];
  const ym = stats.from.slice(0, 4) + stats.from.slice(5, 7);
  if (jing) aoa.push([`${ym}课时费表`, ...Array(12).fill('')]);
  const top = aoa.length;
  aoa.push(header);
  const merges = jing ? ['A1:M1'] : [];
  for (const g of stats.groups) {
    const tStart = aoa.length;
    for (const c of g.courses) {
      const cStart = aoa.length;
      c.rows.forEach((r, i) => {
        const base = [aoa.length === tStart ? g.teacher : '', r.student?.name ?? '（已删除学生）', i === 0 ? c.course.grade : '', i === 0 ? c.course.mode : '', fmtJie(r.jie)];
        aoa.push(jing ? [...base, '', '', c.hist[i] ?? '', '', '', '', r.note ?? '', ''] : [...base, c.hist[i] ?? '', '', '', '', '', '', r.note ?? '', '']);
      });
      if (c.rows.length > 1) for (const col of [2, 3]) merges.push(encodeMerge(cStart, col, aoa.length - 1, col));
    }
    if (aoa.length - tStart > 1) { merges.push(encodeMerge(tStart, 0, aoa.length - 1, 0)); if (jing) merges.push(encodeMerge(tStart, 10, aoa.length - 1, 10)); }
  }
  return { aoa, merges, headerRow: top, title: jing ? `${ym}课时费表` : null, sheetName: jing ? ym : 'Sheet1' };
}
/** 兼容旧接口（Sheet1 默认格式，合并为对象） */
export function sheet1Aoa(stats) {
  const { aoa, merges } = layoutSheet(stats, 'weiduo');
  const dec = (m) => { const [a, b] = m.split(':'); const p = (x) => ({ c: x.charCodeAt(0) - 65, r: +x.slice(1) - 1 }); return { s: p(a), e: p(b) }; };
  return { aoa, merges: merges.map(dec) };
}

export function writeMonthlyWorkbook(_XLSX, state, stats, { format = 'weiduo' } = {}) {
  const L = layoutSheet(stats, format);
  const jing = format === 'jing';
  const numCols = jing ? new Set([5, 6, 8, 9, 10]) : new Set([6, 7]);
  const rows = L.aoa.map((r, ri) => r.map((v, ci) => {
    if (ri <= L.headerRow) return { v, s: jing ? ST.jHead : ST.wdHead };
    if (!jing && ci === 12) return { v, s: ST.wdCheck };
    return { v, s: numCols.has(ci) ? (jing ? ST.jNum : ST.wdNum) : (jing ? ST.jCell : ST.wdCell) };
  }));
  const cols = jing
    ? [[1, 2, 10.3076923076923], [3, 3, 8.15384615384615], [4, 4, 20], [5, 5, 12.3846153846154], [6, 7, 10.3076923076923], [8, 8, 9], [9, 10, 12.7692307692308], [11, 11, 9.61538461538461], [12, 12, 9.15384615384615], [13, 13, 6]]
    : [[1, 2, 11], [3, 4, 8.15384615384615], [5, 5, 12.2307692307692], [6, 6, 9.23076923076923], [7, 8, 11], [9, 10, 13.6153846153846], [11, 11, 9.23076923076923], [12, 12, 17], [13, 13, 24]];
  const main = { name: L.sheetName, rows, merges: L.merges, cols: cols.map(([min, max, width]) => ({ min, max, width })), heights: jing ? {} : { 1: 17.6 }, defaultRowHeight: 16.8, fitWidth: true };
  // V2.0.1：月度课时统计表只有这一页（同原表 Sheet1），不再附「明细」页；明细只在「设置 → 数据备份（非统计表）」
  return writeStyledXlsx([main]);
}

/** 文件名：表名用作文件名（原表没有标题行，导出也不加） */
export const monthlyFileName = (org, campus, month) => `${org ?? ''}${campus ?? ''}${month.slice(0, 4)}年${+month.slice(5, 7)}月课时统计.xlsx`;
export const dataBackupFileName = (day) => `签课数据备份（非统计表）-${day}.xlsx`;

/** 全部数据 Excel（随完整备份一起导出，便于人工查看；恢复以 JSON 为准） */
export function writeAllDataWorkbook(XLSX, state) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['这是签课 App 的数据备份（非统计表）'], ['内容：学生名单、课程、签课明细，只给人查看。'], ['月度课时统计表请在 App 菜单 →「导出月度课时统计表」导出。'], ['换手机恢复数据请用「备份与恢复」里的 JSON 完整备份。']]), '说明（非统计表）');
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
