// 节与课时：1 节 = 2 课时，步长 0.5 节
export const KESHI_PER_JIE = 2;
export const isHalfStep = (n) => typeof n === 'number' && Number.isFinite(n) && n >= 0 && Math.abs(n * 2 - Math.round(n * 2)) < 1e-9;
export const roundHalf = (n) => Math.round(n * 2) / 2;
export const fmtNum = (n) => { const r = Math.round(n * 100) / 100; return String(r); };
export const keshi = (jie) => roundHalf(jie) * KESHI_PER_JIE;
/** 「1节2课时」「0.5节1课时」 */
export const fmtJie = (jie) => `${fmtNum(jie)}节${fmtNum(keshi(jie))}课时`;
/** 解析「3节」「3节6课时」「1.5」等为节数，失败返回 null */
export function parseJie(v) {
  if (typeof v === 'number') return isHalfStep(v) ? v : null;
  const s = String(v ?? '').trim();
  if (!s) return null;
  let m = s.match(/^(\d+(?:\.\d+)?)\s*节/);
  if (m) { const n = +m[1]; return isHalfStep(n) ? n : null; }
  m = s.match(/^(\d+(?:\.\d+)?)\s*课时$/);
  if (m) { const n = +m[1] / KESHI_PER_JIE; return isHalfStep(n) ? n : null; }
  if (/^\d+(\.\d+)?$/.test(s)) { const n = +s; return isHalfStep(n) ? n : null; }
  return null;
}
/** 金额（元）保留两位，避免浮点误差 */
export const yuan = (n) => Math.round((Number(n) || 0) * 100) / 100;
export const fmtYuan = (n) => (n == null ? '' : '¥' + yuan(n).toLocaleString('zh-CN', { minimumFractionDigits: 0, maximumFractionDigits: 2 }));
