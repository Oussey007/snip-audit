// Lecture d'un FEC (fichier des écritures comptables, art. A.47 A-1 du LPF) et reconstitution de la balance.
// Module sans dépendance au navigateur ni à Excel : testable seul.

export const FEC_COLS = ["JournalCode", "JournalLib", "EcritureNum", "EcritureDate", "CompteNum", "CompteLib", "CompAuxNum", "CompAuxLib",
  "PieceRef", "PieceDate", "EcritureLib", "Debit", "Credit", "EcritureLet", "DateLet", "ValidDate", "Montantdevise", "Idevise"];

// Décodage : UTF-8 si le fichier est valide en UTF-8, sinon Windows-1252 (cas le plus fréquent des logiciels français)
export function decode(buf) {
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  try { return { text: new TextDecoder("utf-8", { fatal: true }).decode(u8).replace(/^\uFEFF/, ""), encoding: "UTF-8" }; }
  catch (e) { return { text: new TextDecoder("windows-1252").decode(u8), encoding: "Windows-1252" }; }
}

const norm = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");

// Montant FEC : virgule ou point décimal, sans séparateur de milliers (on tolère espaces et signe)
export function fecAmount(s) {
  let t = String(s ?? "").trim().replace(/[\s\u00a0\u202f]/g, "");
  if (!t) return 0;
  if (t.includes(",") && t.includes(".")) t = t.lastIndexOf(",") > t.lastIndexOf(".") ? t.replace(/\./g, "").replace(",", ".") : t.replace(/,/g, "");
  else t = t.replace(",", ".");
  const v = Number(t);
  return isFinite(v) ? v : NaN;
}
const r2 = v => Math.round(v * 100) / 100;

// Analyse complète : renvoie { info, controles, balance:[{compte, lib, an, debit, credit, solde}], resultat }
export function parseFEC(buf) {
  const { text, encoding } = decode(buf);
  const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
  if (lines.length < 2) throw new Error("fichier vide ou illisible");
  const head = lines[0];
  const sep = ["\t", "|", ";"].map(s => [s, head.split(s).length]).sort((a, b) => b[1] - a[1])[0][0];
  const cols = head.split(sep).map(c => c.trim());
  const idx = {}; FEC_COLS.forEach(c => { idx[c] = cols.findIndex(x => norm(x) === norm(c)); });
  // Variante autorisée : Montant + Sens au lieu de Debit / Credit
  const iMont = cols.findIndex(x => norm(x) === "montant"), iSens = cols.findIndex(x => norm(x) === "sens");
  const manquantes = ["JournalCode", "JournalLib", "EcritureNum", "EcritureDate", "CompteNum", "CompteLib", "EcritureLib"].filter(c => idx[c] < 0);
  const modeSens = (idx.Debit < 0 || idx.Credit < 0) && iMont >= 0 && iSens >= 0;
  if (!modeSens && (idx.Debit < 0 || idx.Credit < 0)) manquantes.push("Debit/Credit");
  if (manquantes.length) throw new Error("colonnes FEC absentes : " + manquantes.join(", ") + " (séparateur détecté : " + ({ "\t": "tabulation", "|": "barre verticale", ";": "point-virgule" }[sep]) + ")");

  const comptes = new Map(), journaux = new Map(), ecritures = new Map();
  let totD = 0, totC = 0, nb = 0, dMin = "99999999", dMax = "00000000", erreursMontant = 0;
  const lignesErr = [];
  for (let i = 1; i < lines.length; i++) {
    const f = lines[i].split(sep);
    if (f.length < 5) continue;
    const g = k => (idx[k] >= 0 ? (f[idx[k]] ?? "").trim() : "");
    let d, c;
    if (modeSens) { const m = fecAmount(f[iMont]); const s = (f[iSens] || "").trim().toUpperCase(); d = /^(D|\+1|1)$/.test(s) ? m : 0; c = /^(C|-1)$/.test(s) ? m : 0; }
    else { d = fecAmount(g("Debit")); c = fecAmount(g("Credit")); }
    if (isNaN(d) || isNaN(c)) { erreursMontant++; if (lignesErr.length < 5) lignesErr.push(i + 1); continue; }
    nb++; totD += d; totC += c;
    const jc = g("JournalCode"), jl = g("JournalLib"), num = g("EcritureNum"), date = g("EcritureDate").replace(/\D/g, "").slice(0, 8);
    if (date && date < dMin) dMin = date; if (date && date > dMax) dMax = date;
    const cpt = g("CompteNum");
    if (!journaux.has(jc)) journaux.set(jc, { code: jc, lib: jl, lignes: 0, an: false });
    const J = journaux.get(jc); J.lignes++;
    if (!comptes.has(cpt)) comptes.set(cpt, { compte: cpt, lib: g("CompteLib"), lignesAN: [], debit: 0, credit: 0, anD: 0, anC: 0 });
    const C = comptes.get(cpt);
    C._rows = C._rows || []; C._rows.push([jc, date, d, c]);
    const k = jc + "\u0001" + num;
    const E = ecritures.get(k) || { d: 0, c: 0 }; E.d += d; E.c += c; ecritures.set(k, E);
  }
  // Journal d'à-nouveaux : libellé explicite, sinon journal dont toutes les écritures sont au 1er jour de l'exercice
  const isAN = J => /a.?nouveau|^an$|report.*nouv|ouverture/i.test(J.lib.normalize("NFD").replace(/[\u0300-\u036f]/g, "")) || /^(an|ran|ouv)$/i.test(J.code);
  for (const J of journaux.values()) J.an = isAN(J);
  if (![...journaux.values()].some(J => J.an)) {
    for (const J of journaux.values()) {
      const dates = new Set(); for (const C of comptes.values()) for (const r of C._rows) if (r[0] === J.code) dates.add(r[1]);
      if (dates.size === 1 && dates.has(dMin) && /0101$/.test(dMin)) J.an = true;
    }
  }
  const anCodes = new Set([...journaux.values()].filter(J => J.an).map(J => J.code));
  const balance = [];
  for (const C of comptes.values()) {
    let an = 0, md = 0, mc = 0;
    for (const [jc, , d, c] of C._rows) { if (anCodes.has(jc)) an += d - c; else { md += d; mc += c; } }
    balance.push({ compte: C.compte, lib: C.lib, an: r2(an), debit: r2(md), credit: r2(mc), solde: r2(an + md - mc) });
  }
  balance.sort((a, b) => a.compte.localeCompare(b.compte));
  let deseq = 0, ecartsDeseq = 0;
  for (const E of ecritures.values()) { const x = r2(E.d - E.c); if (x !== 0) { deseq++; ecartsDeseq += Math.abs(x); } }
  const sum = pfx => r2(balance.filter(b => pfx.some(p => b.compte.startsWith(p))).reduce((a, b) => a + b.solde, 0));
  const resultat = r2(-sum(["6", "7"]));
  const fmt = d => d && d.length === 8 ? `${d.slice(6, 8)}/${d.slice(4, 6)}/${d.slice(0, 4)}` : d;
  const controles = [
    { controle: "Colonnes du FEC", statut: "OK", detail: `${cols.length} colonnes, séparateur ${({ "\t": "tabulation", "|": "barre verticale", ";": "point-virgule" }[sep])}, encodage ${encoding}` },
    { controle: "Lignes lues", statut: erreursMontant ? "Anomalie" : "OK", detail: `${nb.toLocaleString("fr-FR")} lignes d'écritures` + (erreursMontant ? ` ; ${erreursMontant} ligne(s) au montant illisible (ex. lignes ${lignesErr.join(", ")}) ignorée(s)` : "") },
    { controle: "Période des écritures", statut: "Info", detail: `du ${fmt(dMin)} au ${fmt(dMax)}` },
    { controle: "Équilibre global débit / crédit", statut: r2(totD - totC) === 0 ? "OK" : "Anomalie", detail: `Débit ${r2(totD).toLocaleString("fr-FR")} · Crédit ${r2(totC).toLocaleString("fr-FR")} · écart ${r2(totD - totC).toLocaleString("fr-FR")}` },
    { controle: "Équilibre de chaque écriture", statut: deseq ? "Anomalie" : "OK", detail: deseq ? `${deseq} écriture(s) déséquilibrée(s) prise(s) isolément (par journal et n° d'écriture), pour ${r2(ecartsDeseq).toLocaleString("fr-FR")} € d'écarts cumulés. Non bloquant si l'équilibre global est respecté, mais non conforme à la norme du FEC.` : "toutes les écritures sont équilibrées" },
    { controle: "Journal d'à-nouveaux", statut: anCodes.size ? "OK" : "Anomalie", detail: anCodes.size ? [...journaux.values()].filter(J => J.an).map(J => `${J.code} – ${J.lib}`).join(" ; ") : "aucun journal d'à-nouveaux identifié : les soldes d'ouverture ne peuvent pas être isolés" },
    { controle: "Résultat de l'exercice (classes 6 et 7)", statut: "Info", detail: resultat.toLocaleString("fr-FR", { minimumFractionDigits: 2 }) + " €" }
  ];
  return { info: { lignes: nb, encoding, sep, journaux: [...journaux.values()].map(({ code, lib, lignes, an }) => ({ code, lib, lignes, an })), debut: dMin, fin: dMax },
    controles, balance, resultat };
}
