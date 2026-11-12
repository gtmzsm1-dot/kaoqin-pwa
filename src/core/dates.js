// 日期：一律用本地日期字符串 YYYY-MM-DD
const p2 = (n) => String(n).padStart(2, '0');
export const dayOf = (d = new Date()) => { const x = new Date(d); return `${x.getFullYear()}-${p2(x.getMonth() + 1)}-${p2(x.getDate())}`; };
export const isDay = (s) => /^\d{4}-\d{2}-\d{2}$/.test(s) && !Number.isNaN(Date.parse(s + 'T00:00:00'));
export const monthRange = (ym) => { const [y, m] = ym.split('-').map(Number); const last = new Date(y, m, 0).getDate(); return { from: `${y}-${p2(m)}-01`, to: `${y}-${p2(m)}-${p2(last)}` }; };
export const monthOf = (day) => day.slice(0, 7);
export const weekdayOf = (day) => new Date(day + 'T12:00:00').getDay(); // 0=周日
export const WEEK = ['日', '一', '二', '三', '四', '五', '六'];
export const fmtDay = (day) => `${+day.slice(5, 7)}月${+day.slice(8, 10)}日 周${WEEK[weekdayOf(day)]}`;
/** Excel 日期序列号 / Date / 字符串 → YYYY-MM-DD，失败 null */
export function toDay(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : dayOf(v);
  if (typeof v === 'number') { const d = new Date(Math.round((v - 25569) * 864e5)); return `${d.getUTCFullYear()}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}`; }
  const s = String(v).trim().replace(/[./年]/g, '-').replace(/月/g, '-').replace(/日/g, '');
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return null;
  const r = `${m[1]}-${p2(m[2])}-${p2(m[3])}`;
  return isDay(r) ? r : null;
}
