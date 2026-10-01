// Pièces justificatives : enregistrées DANS le classeur Excel (feuille masquée « _Pieces »), comme DataSnipper.
// Chaque classeur a donc ses propres pièces ; rien n'est conservé dans le volet d'un classeur à l'autre.
let XL = null;
const cache = new Map();          // nom -> Blob (session en cours uniquement)
export function init(xl) {
  XL = xl;
  try { indexedDB.deleteDatabase("snip-audit-pieces"); } catch (e) {}   // ancienne version : copies locales supprimées
}
function toB64(file) {
  return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(",")[1] || ""); r.onerror = () => rej(r.error); r.readAsDataURL(file); });
}
async function fromB64(b64) {
  try { return await (await fetch("data:application/pdf;base64," + b64)).blob(); }
  catch (e) { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return new Blob([u], { type: "application/pdf" }); }
}
export async function putFiles(files, onProgress) {
  let added = 0, replaced = 0;
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    onProgress && onProgress(i + 1, files.length, f.name);
    const r = await XL.addPiece(f.name, f.webkitRelativePath || f.name, f.size, await toB64(f));
    cache.set(f.name, f);
    r.replaced ? replaced++ : added++;
  }
  return { added, replaced };
}
export async function listFiles() { return (await XL.listPieces()).map(p => ({ name: p.name, path: p.path || p.name, size: p.size, added: p.added })); }
export async function getFile(name) {
  if (cache.has(name)) return cache.get(name);
  const b64 = await XL.getPiece(name); if (!b64) return null;
  const blob = await fromB64(b64); cache.set(name, blob); return blob;
}
export async function deleteFile(name) { cache.delete(name); return XL.deletePiece(name); }
export async function clearFiles() { cache.clear(); return XL.clearPieces(); }
