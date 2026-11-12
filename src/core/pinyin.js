// 拼音首字母（本地计算，不联网）：用系统内置的中文拼音排序规则定位声母分界字，另加常见多音姓氏修正
const BOUND = '阿八嚓哒妸发旮哈讥咔垃妈拏噢妑七呥仨他屲夕丫帀';
const LET = 'ABCDEFGHJKLMNOPQRSTWXYZ';
const SURNAME = { 曾: 'Z', 单: 'S', 翟: 'Z', 仇: 'Q', 解: 'X', 区: 'O', 查: 'Z', 朴: 'P', 尉: 'Y', 盖: 'G', 缪: 'M', 乐: 'Y', 覃: 'Q', 长: 'C', 重: 'C', 行: 'X', 藏: 'Z', 会: 'H', 柏: 'B', 秘: 'B' };
let col = null;
try { col = new Intl.Collator('zh-CN-u-co-pinyin'); } catch { col = null; }
const cache = new Map();
function charInitial(ch) {
  if (/[a-z0-9]/i.test(ch)) return ch.toLowerCase();
  if (!/[\u3400-\u9fff]/.test(ch) || !col) return '';
  if (cache.has(ch)) return cache.get(ch);
  let i = -1;
  for (let k = 0; k < BOUND.length; k++) { if (col.compare(ch, BOUND[k]) >= 0) i = k; else break; }
  const r = i < 0 ? '' : LET[i].toLowerCase();
  cache.set(ch, r);
  return r;
}
/** 「学生甲」→「xsj」；姓氏多音字取姓氏读音（只对第 1 个字） */
export function initials(s) {
  const a = [...String(s ?? '')];
  return a.map((ch, i) => (i === 0 && SURNAME[ch] ? SURNAME[ch].toLowerCase() : charInitial(ch))).join('');
}
/** 学生搜索：姓名、拼音首字母、学生编号、年级、学校、学管 */
export function matchStudent(s, q) {
  q = String(q ?? '').trim().toLowerCase().replace(/\s+/g, '');
  if (!q) return true;
  if ((s.name ?? '').includes(q) || String(s.code ?? '').toLowerCase().includes(q)) return true;
  if ([s.grade, s.school, s.manager].some((x) => (x ?? '').includes(q))) return true;
  return /^[a-z]+$/.test(q) && initials(s.name).includes(q);
}
