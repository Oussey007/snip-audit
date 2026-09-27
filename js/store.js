// Stockage local des pièces (IndexedDB du volet) : les PDF ne quittent jamais le poste.
const DB = "snip-audit-pieces", ST = "files";
let dbp = null;
function db() {
  if (!dbp) dbp = new Promise((res, rej) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(ST, { keyPath: "name" });
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
  return dbp;
}
async function tx(mode, fn) {
  const d = await db();
  return new Promise((res, rej) => {
    const t = d.transaction(ST, mode); const s = t.objectStore(ST);
    const out = fn(s);
    t.oncomplete = () => res(out && out.result !== undefined ? out.result : out);
    t.onerror = () => rej(t.error);
  });
}
// Mémoire vive en secours si IndexedDB est indisponible (navigation privée, etc.)
const mem = new Map();
let idbOk = true;

export async function putFiles(files) {
  const recs = files.map(f => ({ name: f.name, path: f.webkitRelativePath || f.name, size: f.size, blob: f, added: Date.now() }));
  recs.forEach(r => mem.set(r.name, r));
  if (idbOk) {
    try { await tx("readwrite", s => { recs.forEach(r => s.put(r)); }); }
    catch (e) { idbOk = false; console.warn("IndexedDB indisponible, stockage en mémoire", e); }
  }
  return recs.length;
}
export async function listFiles() {
  if (idbOk) {
    try {
      const all = await tx("readonly", s => s.getAll());
      all.forEach(r => { if (!mem.has(r.name)) mem.set(r.name, r); });
    } catch (e) { idbOk = false; }
  }
  return [...mem.values()].map(r => ({ name: r.name, path: r.path, size: r.size }));
}
export async function getFile(name) {
  if (mem.has(name) && mem.get(name).blob) return mem.get(name).blob;
  if (idbOk) {
    try { const r = await tx("readonly", s => s.get(name)); if (r) { mem.set(name, r); return r.blob; } } catch (e) {}
  }
  return null;
}
export async function clearFiles() {
  mem.clear();
  if (idbOk) { try { await tx("readwrite", s => s.clear()); } catch (e) {} }
}
