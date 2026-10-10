// Armazenamento local (IndexedDB). Tudo fica no navegador do usuário — sem servidor.
const DB_NAME = 'contratofacil';
const DB_VERSION = 1;
export const STORES = ['modelos', 'contratos', 'cadastros', 'clausulas', 'config'];

let dbPromise = null;
const listeners = new Set();

function open() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    if (!('indexedDB' in window)) { reject(new Error('Seu navegador não permite salvar dados (IndexedDB indisponível).')); return; }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      for (const s of STORES) if (!db.objectStoreNames.contains(s)) db.createObjectStore(s, { keyPath: s === 'config' ? 'key' : 'id' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('Falha ao abrir o banco de dados local.'));
    req.onblocked = () => reject(new Error('Feche outras abas do ContratoFácil e recarregue a página.'));
  });
  return dbPromise;
}

function tx(store, mode, fn) {
  return open().then((db) => new Promise((resolve, reject) => {
    const t = db.transaction(store, mode);
    const s = t.objectStore(store);
    let result;
    Promise.resolve(fn(s)).then((r) => { result = r; });
    t.oncomplete = () => resolve(result);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error || new Error('Operação cancelada'));
  }));
}

const req2p = (req) => new Promise((resolve, reject) => { req.onsuccess = () => resolve(req.result); req.onerror = () => reject(req.error); });

export const db = {
  get: (store, id) => tx(store, 'readonly', (s) => req2p(s.get(id))),
  all: (store) => tx(store, 'readonly', (s) => req2p(s.getAll())),
  put: async (store, obj) => {
    const r = await tx(store, 'readwrite', (s) => req2p(s.put(obj)));
    emit(store, obj);
    return r;
  },
  del: async (store, id) => {
    await tx(store, 'readwrite', (s) => req2p(s.delete(id)));
    emit(store, { id, deleted: true });
  },
  clear: async (store) => {
    await tx(store, 'readwrite', (s) => req2p(s.clear()));
    emit(store, null);
  },
};

function emit(store, obj) {
  for (const l of listeners) { try { l(store, obj); } catch (e) { console.error(e); } }
}

export function onChange(fn) {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

export function uid(prefix = '') {
  const r = crypto.getRandomValues(new Uint32Array(2));
  return prefix + Date.now().toString(36) + r[0].toString(36) + r[1].toString(36).slice(0, 4);
}

/** Pede ao navegador para não apagar os dados automaticamente. */
export async function persistStorage() {
  try {
    if (navigator.storage && navigator.storage.persist) {
      const already = await navigator.storage.persisted();
      return already || (await navigator.storage.persist());
    }
  } catch { /* ignore */ }
  return false;
}

export async function storageEstimate() {
  try {
    if (navigator.storage && navigator.storage.estimate) return await navigator.storage.estimate();
  } catch { /* ignore */ }
  return null;
}
