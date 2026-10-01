// Excel simulé pour les tests hors Office (test.html). Même interface que excel-bridge.js.
export const SNIP_SHEET = "_Snips";
export const COLORS = { texte: "#DDEBF7", somme: "#E4DFEC", valide: "#C6EFCE", exception: "#FFC7CE" };
const M = window.__MOCK__;           // { sheets: {nom: {cells:{A1:{v,f}}}}, liens:[...], sel:{sheet,cell} }
let cb = null;
const colIdx = c => c.split("").reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const colName = i => { let s = ""; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
const parse = a => { const m = /^([A-Z]+)(\d+)$/.exec(a); return { c: colIdx(m[1]), r: +m[2] - 1 }; };
const cell = (sh, a) => (M.sheets[sh] = M.sheets[sh] || { cells: {} }).cells[a] = (M.sheets[sh].cells[a] || { v: "", f: "" });

export async function ready() { return { host: "Excel" }; }
export function onSelection(f) { cb = f; }
window.__mock_select = (sheet, a) => { M.sel = { sheet, cell: a }; cb && cb(); };
export async function getSelection() {
  const { sheet, cell: a } = M.sel; const p = parse(a); const c = cell(sheet, a);
  const hdr = cell(sheet, colName(p.c) + "1");
  const rowValues = []; for (let i = 0; i < 16; i++) rowValues.push(cell(sheet, colName(i) + (p.r + 1)).v);
  return { sheet, cell: a, row: p.r, col: p.c, value: c.v, formula: c.f || c.v, header: hdr.v, rowValues };
}
export async function readLiens() {
  const map = new Map();
  for (const l of M.liens) { if (!map.has(l.ref)) map.set(l.ref, []); map.get(l.ref).push(l); }
  return map;
}
export async function readSnips() { return (M.snips || []).map(s => ({ ...s })); }
export async function saveSnip(rec, writeValue) {
  M.snips = (M.snips || []).filter(s => !(s.sheet === rec.sheet && s.cell === rec.cell && s.type !== "auto"));
  const c = cell(rec.sheet, rec.cell);
  const hasFormula = typeof c.f === "string" && c.f.startsWith("=");
  let written = false;
  if (writeValue !== undefined && writeValue !== null) {
    if (rec.type === "valide" || rec.type === "exception") { if (c.v === "" || c.v === null) { c.v = writeValue; written = true; } }
    else if (!hasFormula) { c.v = writeValue; written = true; }
  }
  c.fill = COLORS[rec.type];
  M.snips.push({ ...rec });
  M.log = (M.log || []).concat([{ op: "save", rec, written }]);
  return { written, hasFormula };
}
export async function deleteSnip(sheet, a) {
  const n = (M.snips || []).length;
  M.snips = (M.snips || []).filter(s => !(s.sheet === sheet && s.cell === a && s.type !== "auto"));
  delete cell(sheet, a).fill;
  return M.snips.length < n;
}
export async function selectCell(sheet, a) { window.__mock_select(sheet, a); }

// ---------- Pièces dans le classeur (simulation)
export const PIECE_SHEET = "_Pieces";
M.pieces = M.pieces || [];
export async function listPieces() { return M.pieces.map((p, i) => ({ row: i + 1, name: p.name, path: p.path, size: p.size, chunks: 1, added: p.added })); }
export async function addPiece(name, path, size, b64) {
  const ex = M.pieces.find(p => p.name === name);
  if (ex) Object.assign(ex, { path, size, b64 }); else M.pieces.push({ name, path, size, b64, added: new Date().toLocaleString("fr-FR") });
  return { replaced: !!ex };
}
export async function getPiece(name) { const p = M.pieces.find(x => x.name === name); return p ? p.b64 : null; }
export async function deletePiece(name) { const n = M.pieces.length; M.pieces = M.pieces.filter(p => p.name !== name); return M.pieces.length < n; }
export async function clearPieces() { const n = M.pieces.length; M.pieces = []; return n; }
