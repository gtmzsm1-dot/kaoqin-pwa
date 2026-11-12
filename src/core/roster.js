// 名单导入：读取「学生名单」页（第 4 行表头），逐行校验，按学生编号合并；编号为「待补」的按姓名合并
import { STATUS_LIST, isPendingCode } from './store.js';

export const ROSTER_COLS = ['序号', '姓名', '学生编号', '学校', '年级', '届别', '状态', '负责学管'];
const FIELD = { 序号: 'seq', 姓名: 'name', 学生编号: 'code', 编号: 'code', 学校: 'school', 年级: 'grade', 届别: 'cohort', 状态: 'status', 负责学管: 'manager', 学管: 'manager' };
const CMP = ['name', 'code', 'school', 'grade', 'cohort', 'status', 'manager'];
const LABEL = { name: '姓名', code: '学生编号', school: '学校', grade: '年级', cohort: '届别', status: '状态', manager: '负责学管', seq: '序号' };
const cell = (v) => String(v ?? '').replace(/\s+/g, ' ').trim();

/** 从工作簿取出名单表格：返回 { sheetName, headerRow(1 起), header, rows:[{rowNo, values}] } */
export function readRosterSheet(XLSX, wb) {
  const order = [...wb.SheetNames].sort((a, b) => (b === '学生名单') - (a === '学生名单'));
  for (const name of order) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: true, blankrows: true });
    for (let i = 0; i < Math.min(aoa.length, 12); i++) {
      const hdr = aoa[i].map(cell);
      if (hdr.includes('姓名') && (hdr.includes('学生编号') || hdr.includes('编号') || hdr.includes('状态'))) {
        const map = {};
        hdr.forEach((h, j) => { if (FIELD[h] && map[FIELD[h]] == null) map[FIELD[h]] = j; });
        const rows = [];
        for (let r = i + 1; r < aoa.length; r++) {
          const v = {}; for (const [f, j] of Object.entries(map)) v[f] = cell(aoa[r][j]);
          rows.push({ rowNo: r + 1, values: v });
        }
        return { sheetName: name, headerRow: i + 1, header: hdr, map, rows };
      }
    }
  }
  throw new Error('没找到名单表头（需要有「姓名」「学生编号」「状态」等列，建议用导入模板）');
}

/** 对照现有学生，生成导入计划 */
export function planImport(sheet, state) {
  const existing = Object.values(state.students);
  const byCode = new Map(existing.filter((s) => !isPendingCode(s.code)).map((s) => [s.code, s]));
  const byName = new Map(); for (const s of existing) { if (!byName.has(s.name)) byName.set(s.name, []); byName.get(s.name).push(s); }
  const seenCode = new Map(), seenPendingName = new Map(), usedIds = new Set();
  const rows = [];
  for (const { rowNo, values } of sheet.rows) {
    if (!Object.values(values).some((x) => x)) continue; // 空行
    const errors = [], warnings = [];
    const data = { name: values.name ?? '', code: values.code || '待补', school: values.school ?? '', grade: values.grade ?? '', cohort: values.cohort ?? '',
      status: values.status || '在读', manager: values.manager ?? '' };
    if (values.seq && /^\d+$/.test(values.seq)) data.seq = +values.seq;
    if (!data.name) errors.push('姓名为空');
    else if (data.name.length > 20) errors.push('姓名过长');
    if (!STATUS_LIST.includes(data.status)) warnings.push(`状态「${data.status}」不是 ${STATUS_LIST.join('/')}，照原样导入（不算在读，默认不显示）`);
    if (!values.status) warnings.push('状态为空，按「在读」导入');
    const pending = isPendingCode(data.code);
    if (!pending) {
      if (seenCode.has(data.code)) errors.push(`学生编号与第 ${seenCode.get(data.code)} 行重复`); else seenCode.set(data.code, rowNo);
    } else if (data.name) {
      if (seenPendingName.has(data.name)) errors.push(`与第 ${seenPendingName.get(data.name)} 行同名且都没有编号，无法区分`); else seenPendingName.set(data.name, rowNo);
    }
    let match = null;
    if (!errors.length) {
      if (!pending) match = byCode.get(data.code) ?? null;
      if (!match && data.name) {
        const cands = (byName.get(data.name) ?? []).filter((s) => (pending ? true : isPendingCode(s.code)) && !usedIds.has(s.id));
        if (cands.length > 1) errors.push('现有名单里有多个同名学生，请补学生编号');
        else if (cands.length === 1) match = cands[0];
      }
      if (match && usedIds.has(match.id)) errors.push('与本文件另一行对应同一学生');
    }
    let action = 'error', diff = [];
    if (!errors.length) {
      if (!match) action = 'add';
      else {
        usedIds.add(match.id);
        if (pending && !isPendingCode(match.code)) data.code = match.code; // 不用「待补」覆盖已有编号
        diff = CMP.filter((k) => (match[k] ?? '') !== (data[k] ?? '')).map((k) => `${LABEL[k]}：${match[k] || '空'}→${data[k] || '空'}`);
        action = diff.length || (data.seq != null && data.seq !== match.seq) ? 'update' : 'same';
      }
    }
    rows.push({ rowNo, data, action, errors, warnings, matchId: match?.id ?? null, diff });
  }
  const count = (a) => rows.filter((r) => r.action === a).length;
  return { sheetName: sheet.sheetName, rows, summary: { total: rows.length, add: count('add'), update: count('update'), same: count('same'), error: count('error') } };
}

/** 导入模板（只含虚构示例） */
export function writeRosterTemplate(XLSX) {
  const aoa = [
    ['学生总名单（导入模板）'],
    ['第 4 行是表头，从第 5 行起每行一个学生；状态只能填 在读/停课/毕业/结课；没有编号写「待补」。下面是虚构示例，请删掉后填写。'],
    [],
    ROSTER_COLS,
    [1, '学生甲', 'TEST0001', '示例小学', '五年级', '小2028届', '在读', '学管A'],
    [2, '学生乙', 'TEST0002', '示例中学', '初一', '初2029届', '在读', '学管A'],
    [3, '学生丙', '待补', '示例中学', '初二', '初2028届', '停课', '学管B'],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 7 } }];
  ws['!cols'] = [6, 10, 13, 12, 8, 10, 7, 10].map((w) => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '学生名单');
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}
