// 完整备份：一个 JSON 文件（数据 + 操作记录 + 校验和）；恢复前校验
import { canonical, sha256Hex } from './util.js';
import { Store } from './store.js';

export const BACKUP_FORMAT = 'kaoqin-pwa-backup';
export const BACKUP_VERSION = 1;

export async function makeBackup(store, appVersion) {
  const { state, ops } = store;
  const meta = { format: BACKUP_FORMAT, version: BACKUP_VERSION, appVersion, exportedAt: store.stamp(),
    counts: { students: Object.keys(state.students).length, courses: Object.keys(state.courses).length, sessions: Object.keys(state.sessions).length, ops: ops.length },
    checksum: await sha256Hex(canonical({ state, ops })) };
  return new TextEncoder().encode(JSON.stringify({ meta, state, ops }));
}

export async function readBackup(bytes) {
  let obj;
  try { obj = JSON.parse(typeof bytes === 'string' ? bytes : new TextDecoder().decode(bytes)); } catch { throw new Error('文件无法解析，不是本 App 的备份'); }
  const { meta, state, ops } = obj ?? {};
  if (meta?.format !== BACKUP_FORMAT) throw new Error('不是本 App 的备份文件');
  if (meta.version > BACKUP_VERSION) throw new Error('备份来自更新的版本，请先更新 App');
  if (!state || !Array.isArray(ops)) throw new Error('备份内容不完整');
  if (await sha256Hex(canonical({ state, ops })) !== meta.checksum) throw new Error('校验和不一致，备份可能已损坏');
  const issues = new Store({ state, ops }).verifyIntegrity();
  if (issues.length) throw new Error('数据自检失败：' + issues.slice(0, 3).join('；'));
  return { meta, state, ops };
}
