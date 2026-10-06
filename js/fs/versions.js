// Versions successives d'une plaquette : instantané des montants lus et du texte de l'annexe, comparaison avec la version
// précédente, et reprise des validations de l'auditeur pour tout ce qui n'a pas bougé.
import { normLabel } from "./pcg.js?v=10";

const nz = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
const num = v => (v === null || v === undefined || v === "") ? "" : Math.round(Number(v) * 100) / 100;

// Clés stables (indépendantes de la page et de la position)
export const keyP = r => `S|${r.id}|${r.cle}`;
export const keyA = r => `A|${r.type}|${nz(r.note)}|${nz(r.ligne)}|${nz(r.colonne)}|${r.controle}|${nz(r.reference)}`;

export function snapshot(rowsP, rowsA, lines) {
  const items = [];
  for (const r of rowsP) items.push({ type: "Comptes annuels", key: keyP(r), lib: `${r.sectionLib} – ${r.poste} (${r.colonne})`, v1: num(r.plaquette), v2: num(r.balance), page: r.page ?? "" });
  const seen = new Map();
  for (const r of rowsA) { let k = keyA(r); const n = (seen.get(k) || 0) + 1; seen.set(k, n); if (n > 1) k += "|" + n; items.push({ type: "Annexe – tableaux", key: k, lib: `${r.note} – ${r.ligne} (${r.colonne}) · ${r.controle}`, v1: num(r.annexe), v2: num(r.ref), page: r.page ?? "" }); }
  lines.forEach((l, i) => items.push({ type: "Annexe – texte", key: `T|${i}`, lib: l.text, v1: "", v2: "", page: l.page }));
  return items;
}
export function hashItems(items) {
  let h = 2166136261; const s = items.map(i => `${i.key}=${i.v1}/${i.v2}/${i.type === "Annexe – texte" ? nz(i.lib) : ""}`).join("\n");
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; }
  return h.toString(16);
}

// Lignes d'en-tête / pied de page répétées et numéros de page : ignorées dans la comparaison du texte
const isBoiler = t => /^page \d+$|^\d{1,3}$|voir rapport|^periode du|comptes annuels$/.test(t);

// Comparaison du texte (plus longue sous-séquence commune sur les lignes normalisées), regroupée par blocs
function textDiff(oldL, newL) {
  const a = oldL.map(x => nz(x.lib)).map(t => isBoiler(t) ? "" : t), b = newL.map(x => nz(x.lib)).map(t => isBoiler(t) ? "" : t);
  const A = a.map((t, i) => [t, i]).filter(x => x[0]), B = b.map((t, i) => [t, i]).filter(x => x[0]);
  const n = A.length, m = B.length; const L = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = A[i][0] === B[j][0] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
  const ops = []; let i = 0, j = 0;
  while (i < n || j < m) {
    if (i < n && j < m && A[i][0] === B[j][0]) { ops.push(["=", A[i][1], B[j][1]]); i++; j++; }
    else if (j < m && (i >= n || L[i][j + 1] >= L[i + 1][j])) { ops.push(["+", null, B[j][1]]); j++; }
    else { ops.push(["-", A[i][1], null]); i++; }
  }
  const blocks = []; let cur = null;
  for (const o of ops) {
    if (o[0] === "=") { cur = null; continue; }
    if (!cur) { cur = { old: [], neu: [] }; blocks.push(cur); }
    if (o[0] === "-") cur.old.push(oldL[o[1]]); else cur.neu.push(newL[o[2]]);
  }
  return { blocks, unchangedNew: new Set(ops.filter(o => o[0] === "=").map(o => o[2])) };
}

// Différences entre deux instantanés
export function compare(prevItems, newItems) {
  const P = new Map(prevItems.filter(i => i.type !== "Annexe – texte").map(i => [i.key, i]));
  const N = new Map(newItems.filter(i => i.type !== "Annexe – texte").map(i => [i.key, i]));
  const changes = [], status = new Map();
  for (const [k, n] of N) {
    const p = P.get(k);
    if (!p) { status.set(k, { s: "Nouveau" }); if (n.v1 !== "" || n.v2 !== "") changes.push({ type: n.type, element: n.lib, ancien: "", nouveau: fmtPair(n), statut: "Ajouté", page: n.page, key: k }); continue; }
    if (String(p.v1) === String(n.v1) && String(p.v2) === String(n.v2)) { status.set(k, { s: "Inchangé" }); continue; }
    status.set(k, { s: "Modifié", old: p });
    changes.push({ type: n.type, element: n.lib, ancien: fmtPair(p), nouveau: fmtPair(n), statut: describe(p, n), page: n.page, key: k });
  }
  for (const [k, p] of P) if (!N.has(k) && (p.v1 !== "" || p.v2 !== "")) changes.push({ type: p.type, element: p.lib, ancien: fmtPair(p), nouveau: "", statut: "Supprimé", page: "", key: k });
  const td = textDiff(prevItems.filter(i => i.type === "Annexe – texte"), newItems.filter(i => i.type === "Annexe – texte"));
  for (const b of td.blocks) changes.push({ type: "Annexe – texte", element: b.neu.length ? `p.${b.neu[0].page}` : `ancienne p.${b.old[0].page}`, ancien: b.old.map(x => x.lib).join("\n"), nouveau: b.neu.map(x => x.lib).join("\n"), statut: b.old.length && b.neu.length ? "Modifié" : b.neu.length ? "Ajouté" : "Supprimé", page: b.neu[0]?.page ?? "", textIdx: b.neu.map(x => Number(x.key.slice(2))) });
  return { changes, status, textChanged: td.blocks.length > 0, unchangedText: td.unchangedNew };
}
const f2 = v => v === "" ? "—" : Number(v).toLocaleString("fr-FR", { maximumFractionDigits: 2 });
const fmtPair = i => i.type === "Comptes annuels" ? `plaquette ${f2(i.v1)} · balance ${f2(i.v2)}` : `annexe ${f2(i.v1)} · référence ${f2(i.v2)}`;
function describe(p, n) {
  const a = String(p.v1) !== String(n.v1), b = String(p.v2) !== String(n.v2);
  if (n.type === "Comptes annuels") return a && b ? "Modifié (plaquette et balance)" : a ? "Modifié (plaquette)" : "Modifié (balance / FEC)";
  return a && b ? "Modifié (annexe et référence)" : a ? "Modifié (annexe)" : "Modifié (référence)";
}

// Lecture / écriture de la feuille d'historique (valeurs brutes)
export const VERS_HEAD = ["Version", "Date", "Plaquette", "Empreinte", "Type", "Clé", "Libellé", "Valeur 1", "Valeur 2", "Page"];
export function parseVersions(values) {
  if (!values || values.length < 2) return [];
  const by = new Map();
  for (const r of values.slice(1)) {
    if (!r[0]) continue; const v = Number(r[0]);
    if (!by.has(v)) by.set(v, { version: v, date: r[1], pdf: r[2], hash: r[3], items: [] });
    by.get(v).items.push({ type: r[4], key: r[5], lib: String(r[6] ?? ""), v1: r[7] === "" ? "" : Number(r[7]), v2: r[8] === "" ? "" : Number(r[8]), page: r[9] });
  }
  return [...by.values()].sort((a, b) => a.version - b.version);
}
export function versionRows(v) { return v.items.map(i => [v.version, v.date, v.pdf, v.hash, i.type, i.key, i.lib, i.v1, i.v2, i.page]); }
