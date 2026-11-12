// IndexedDB：kv（数据、PIN、设置）、snapshots（本机自动快照，保留最近 20 份）
const DB_NAME = 'kaoqin-pwa';
export const SNAPSHOT_KEEP = 20;
let dbp = null;
function open() {
  if (dbp) return dbp;
  dbp = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains('kv')) db.createObjectStore('kv');
      if (!db.objectStoreNames.contains('snapshots')) db.createObjectStore('snapshots', { keyPath: 'id' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbp;
}
function tx(stores, mode, fn) {
  return open().then((db) => new Promise((resolve, reject) => {
    const t = db.transaction(stores, mode);
    let result;
    Promise.resolve(fn(t)).then((r) => { result = r; });
    t.oncomplete = () => resolve(result);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error ?? new Error('写入被中止'));
  }));
}
const req2p = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
export const kvGet = (k) => tx(['kv'], 'readonly', (t) => req2p(t.objectStore('kv').get(k)));
export const kvSet = (k, v) => tx(['kv'], 'readwrite', (t) => { t.objectStore('kv').put(v, k); });
export const saveData = (state, ops) => kvSet('data', { state, ops, savedAt: new Date().toISOString() });
export const loadData = () => kvGet('data');

export async function addSnapshot(reason, state, ops) {
  const id = new Date().toISOString() + '-' + Math.random().toString(36).slice(2, 6);
  await tx(['snapshots'], 'readwrite', (t) => { t.objectStore('snapshots').put({ id, createdAt: new Date().toISOString(), reason, state, ops,
    counts: { students: Object.keys(state.students).length, sessions: Object.keys(state.sessions).length } }); });
  const all = await listSnapshots();
  if (all.length > SNAPSHOT_KEEP) await tx(['snapshots'], 'readwrite', (t) => { for (const d of all.slice(SNAPSHOT_KEEP)) t.objectStore('snapshots').delete(d.id); });
  return id;
}
export async function listSnapshots() {
  const all = await tx(['snapshots'], 'readonly', (t) => req2p(t.objectStore('snapshots').getAll()));
  return all.sort((a, b) => b.id.localeCompare(a.id));
}
export const getSnapshot = (id) => tx(['snapshots'], 'readonly', (t) => req2p(t.objectStore('snapshots').get(id)));
