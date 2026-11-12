// V2 查询：按老师查课时、按学生查次数、连续缺课提醒
import { ATT, courseLabel, sessionTeacher } from './store.js';
import { roundHalf } from './jie.js';
const zh = (a, b) => String(a).localeCompare(String(b), 'zh-CN');
const inRange = (state, from, to) => Object.values(state.sessions).filter((s) => s.date >= from && s.date <= to)
  .sort((a, b) => a.date.localeCompare(b.date) || a.createdAt.localeCompare(b.createdAt));

/** 按老师：上课次数、授课节数（每次课按到课学生里的最大节数计）、学生消课合计；明细按次 */
export function teacherSummary(state, { from, to }) {
  const map = new Map();
  for (const s of inRange(state, from, to)) {
    const t = sessionTeacher(state, s) || '（未填教师）';
    const present = s.records.filter((r) => r.status === ATT.present && r.jie > 0);
    const teachJie = present.reduce((m, r) => Math.max(m, r.jie), 0);
    const stuJie = present.reduce((a, r) => a + r.jie, 0);
    if (!map.has(t)) map.set(t, { teacher: t, sessions: 0, teachJie: 0, studentJie: 0, substitute: 0, items: [] });
    const x = map.get(t);
    x.sessions++; x.teachJie += teachJie; x.studentJie += stuJie; if (s.subTeacher) x.substitute++;
    x.items.push({ sessionId: s.id, date: s.date, course: courseLabel(state.courses[s.courseId]), present: present.length, teachJie, studentJie: stuJie, tag: s.tag ?? '', substitute: !!s.subTeacher });
  }
  return [...map.values()].map((x) => ({ ...x, teachJie: roundHalf(x.teachJie), studentJie: roundHalf(x.studentJie) })).sort((a, b) => zh(a.teacher, b.teacher));
}

/** 按学生：到课次数、节数、请假次数、缺勤次数 */
export function studentSummary(state, { from, to }) {
  const map = new Map();
  for (const s of inRange(state, from, to)) for (const r of s.records) {
    if (!map.has(r.studentId)) map.set(r.studentId, { studentId: r.studentId, name: state.students[r.studentId]?.name ?? '（已删除学生）', present: 0, jie: 0, leave: 0, absent: 0, items: [] });
    const x = map.get(r.studentId);
    if (r.status === ATT.present) { x.present++; x.jie += r.jie; } else if (r.status === ATT.leave) x.leave++; else x.absent++;
    x.items.push({ sessionId: s.id, date: s.date, course: courseLabel(state.courses[s.courseId]), status: r.status, jie: r.jie, tag: s.tag ?? '' });
  }
  return [...map.values()].sort((a, b) => zh(a.name, b.name));
}

/** 连续缺课：同一门课最近连续 n 次（含）请假或缺勤 */
export function absenceStreaks(state, n = 2) {
  const byCourse = new Map();
  for (const s of Object.values(state.sessions)) { if (!byCourse.has(s.courseId)) byCourse.set(s.courseId, []); byCourse.get(s.courseId).push(s); }
  const out = [];
  for (const [cid, list] of byCourse) {
    const c = state.courses[cid]; if (!c || c.archived) continue;
    list.sort((a, b) => b.date.localeCompare(a.date) || b.createdAt.localeCompare(a.createdAt)); // 最新在前
    const ids = new Set(list.flatMap((s) => s.records.map((r) => r.studentId)));
    for (const sid of ids) {
      let streak = 0, last = null; const dates = [];
      for (const s of list) {
        const r = s.records.find((x) => x.studentId === sid);
        if (!r) continue;
        if (r.status === ATT.present) break;
        streak++; dates.push(`${s.date}${r.status}`); last ??= s.date;
      }
      if (streak >= n) out.push({ studentId: sid, name: state.students[sid]?.name ?? '（已删除学生）', courseId: cid, course: courseLabel(c), streak, last, dates: dates.reverse() });
    }
  }
  return out.sort((a, b) => b.streak - a.streak || b.last.localeCompare(a.last));
}
export const streakKey = (x) => x.courseId + '|' + x.studentId;
