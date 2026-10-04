// Rapprochement plaquette <-> balance (FEC) et contrôles arithmétiques de la plaquette.
import { comptesDe, controleAffectation, SECTION_LABEL } from "./pcg.js?v=8";

const r2 = v => Math.round(v * 100) / 100;
const eur = v => (v ?? 0).toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const COLS = { actif: [["brut", "Brut"], ["amort", "Amort. / dépréc."], ["net", "Net"]], passif: [["n", "Montant"]], cr: [["n", "Montant"]] };
const N1 = { actif: "net1", passif: "n1", cr: "n1" };
// côté « débit » d'un poste : actif et charges +1, passif et produits -1 (pour reconnaître les reclassements)
const sideSign = p => p.section === "actif" ? 1 : p.section === "passif" ? -1 : (p.sens || 1);

function evalFormule(f, get) {
  let total = 0; const re = /([+-]?)\s*([A-Z0-9_]+)/g; let m; const ids = [];
  while ((m = re.exec(f))) { const v = get(m[2]); ids.push(m[2]); if (v !== null && v !== undefined) total += (m[1] === "-" ? -1 : 1) * v; }
  return { total, ids };
}

export function tieOut(postes, balance, matched, fec) {
  const byId = new Map(postes.map(p => [p.id, p]));
  const exp = new Map();          // valeurs attendues d'après la balance : {brut, amort, net} ou {n}
  const det = new Map();          // comptes retenus par poste
  const notes = new Map();
  // 1) postes simples
  for (const p of postes.filter(p => !p.total)) {
    const sg = p.section === "actif" ? 1 : p.section === "passif" ? -1 : (p.sens || 1);
    const cs = p.regle ? comptesDe(p.regle, balance) : [];
    const ca = p.section === "actif" && p.amort ? comptesDe(p.amort, balance) : [];
    const brut = r2(sg * cs.reduce((a, b) => a + b.solde, 0));
    if (p.section === "actif") { const amort = r2(-ca.reduce((a, b) => a + b.solde, 0)); exp.set(p.id, { brut, amort, net: r2(brut - amort) }); }
    else exp.set(p.id, { n: brut });
    det.set(p.id, [...cs.map(b => ({ ...b, role: "" })), ...ca.map(b => ({ ...b, role: "amort." }))]);
  }
  // 2) postes absents de la plaquette avec une ligne de repli (ex. « Droits du concédant » présentés en « Avances conditionnées »)
  for (const p of postes.filter(p => !p.total && p.repli)) {
    const e = exp.get(p.id); const nz = Object.values(e).some(v => Math.abs(v) > 0.004);
    if (!matched.has(p.id) && nz && byId.has(p.repli)) {
      const t = exp.get(p.repli); for (const k of Object.keys(t)) t[k] = r2(t[k] + (e[k] || 0));
      det.get(p.repli).push(...det.get(p.id));
      notes.set(p.repli, (notes.get(p.repli) ? notes.get(p.repli) + " " : "") + `Inclut « ${p.libelle} » (${eur(e.n ?? e.brut)}), poste absent de la plaquette.`);
      for (const k of Object.keys(e)) e[k] = 0; det.set(p.id, []);
    }
  }
  // 3) totaux (ordre de déclaration : les totaux utilisent des totaux déjà calculés)
  for (const p of postes.filter(p => p.total)) {
    const keys = p.section === "actif" ? ["brut", "amort", "net"] : ["n"];
    const v = {}; for (const k of keys) v[k] = r2(evalFormule(p.formule, id => exp.get(id)?.[k] ?? 0).total);
    exp.set(p.id, v);
  }
  // 4) lignes de rapprochement
  const rows = [];
  for (const p of postes) {
    const line = matched.get(p.id); const e = exp.get(p.id);
    for (const [k, lib] of COLS[p.section]) {
      const pv = line && line.values[k] !== undefined ? line.values[k] : null;
      const bv = e[k];
      if (pv === null && Math.abs(bv) < 0.5 && !line) continue;
      if (pv === null && Math.abs(bv) < 0.5) continue;
      const nKids = p.total ? evalFormule(p.formule, id => exp.get(id)?.[k]).ids.filter(id => matched.get(id)?.values?.[k] !== undefined).length : 1;
      const tol = p.total ? Math.max(1, Math.ceil(nKids / 2)) : 1;
      const ecart = r2(bv - (pv ?? 0));
      let statut;
      if (pv === null) statut = line ? "Non renseigné dans la plaquette" : "Absent de la plaquette";
      else if (Math.abs(ecart) < 0.005) statut = "Concordant";
      else if (Math.abs(ecart) <= tol) statut = "Concordant (arrondi)";
      else statut = "Écart";
      const comptes = (det.get(p.id) || []).filter(b => Math.abs(b.solde) > 0.004 && (p.section !== "actif" || (k === "amort" ? b.role : (k === "brut" ? !b.role : true))));
      rows.push({
        section: p.section, sectionLib: SECTION_LABEL[p.section], id: p.id, poste: p.libelle, total: !!p.total, colonne: lib, cle: k,
        plaquette: pv, balance: r2(bv), ecart, statut, commentaire: notes.get(p.id) && k !== "amort" ? notes.get(p.id) : "",
        comptes: p.total ? "" : comptes.slice(0, 25).map(b => `${b.compte} ${b.lib} (${eur(b.solde)})`).join(" ; ") + (comptes.length > 25 ? ` ; … (${comptes.length} comptes)` : ""),
        page: line ? line.page : null, box: line ? line.boxes[k] || null : null, libellePlaquette: line ? line.label : "",
        n1: line ? (line.values[N1[p.section]] ?? null) : null, n1box: line ? line.boxes[N1[p.section]] || null : null, side: sideSign(p),
        // signe comptable pour les cumuls : débit positif (actif, charges), crédit négatif (passif, produits, amortissements)
        sg: p.section === "actif" ? (k === "amort" ? -1 : 1) : p.section === "passif" ? -1 : (p.sens || 1), nature: p.total ? "Total" : "Poste"
      });
    }
  }
  // 5) reclassements probables : deux écarts de postes simples qui se compensent
  // un reclassement ne modifie pas les totaux : on ne le retient que si le total général (bilan) ou le résultat (compte de résultat) concorde
  const okTotal = (id, k) => { const r = rows.find(x => x.id === id && x.cle === k); return !r || r.statut.startsWith("Concordant"); };
  const famOk = { bilan: okTotal("A_TG", "net") && okTotal("P_TG", "n"), cr: okTotal("R_BEN", "n") };
  const ec = rows.filter(r => r.statut === "Écart" && !r.total && r.cle !== "amort" && r.id !== "P_RES" && famOk[r.section === "cr" ? "cr" : "bilan"]);
  const used = new Set();
  for (let i = 0; i < ec.length; i++) for (let j = i + 1; j < ec.length; j++) {
    const a = ec[i], b = ec[j]; if (used.has(a) || used.has(b)) continue;
    const fam = s => s === "cr" ? "cr" : "bilan"; if (fam(a.section) !== fam(b.section)) continue;
    if (Math.abs(a.ecart * a.side + b.ecart * b.side) <= 2 && Math.abs(a.ecart) > 2) {
      used.add(a); used.add(b);
      a.statut = b.statut = "Écart – reclassement probable";
      a.commentaire = (a.commentaire ? a.commentaire + " " : "") + `Écart compensé par « ${b.poste} » (${eur(b.ecart)}) : la plaquette présente ces comptes dans un autre poste que le plan de regroupement. Classement à apprécier.`;
      b.commentaire = (b.commentaire ? b.commentaire + " " : "") + `Écart compensé par « ${a.poste} » (${eur(a.ecart)}) : classement différent du plan de regroupement. Classement à apprécier.`;
    }
  }
  for (const r of rows.filter(r => r.total && r.statut === "Écart")) {
    const kids = evalFormule(byId.get(r.id).formule, () => 0).ids;
    if (rows.some(x => kids.includes(x.id) && x.cle === r.cle && x.statut.startsWith("Écart"))) r.commentaire = "Découle des écarts sur les postes qui le composent.";
  }
  // 6) contrôles de la plaquette et de l'affectation des comptes
  const controles = [];
  for (const p of postes) {
    const line = matched.get(p.id); if (!line) continue;
    if (p.section === "actif" && line.values.brut !== undefined) {
      const net = r2((line.values.brut ?? 0) - (line.values.amort ?? 0));
      if (line.values.net !== undefined && Math.abs(net - line.values.net) > 1) controles.push({ type: "Plaquette", controle: `${p.libelle} : brut − amortissements ≠ net`, statut: "Anomalie", detail: `${eur(line.values.brut)} − ${eur(line.values.amort ?? 0)} = ${eur(net)}, net affiché ${eur(line.values.net)}`, page: line.page, box: line.boxes.net });
    }
    if (p.total) for (const [k] of [...COLS[p.section], [N1[p.section]]]) {
      if (line.values[k] === undefined) continue;
      const f = evalFormule(p.formule, id => matched.get(id)?.values?.[k]);
      const n = f.ids.filter(id => matched.get(id)?.values?.[k] !== undefined).length;
      const d = r2(line.values[k] - f.total);
      if (Math.abs(d) > Math.max(1, Math.ceil(n / 2))) controles.push({ type: "Plaquette", controle: `${p.libelle} (${k === N1[p.section] ? "N-1" : COLS[p.section].find(c => c[0] === k)?.[1] || k}) : le total ne correspond pas à la somme des lignes`, statut: "Anomalie", detail: `Somme des lignes ${eur(f.total)}, total affiché ${eur(line.values[k])}, écart ${eur(d)}`, page: line.page, box: line.boxes[k] });
    }
  }
  const nbArith = controles.length;
  controles.unshift({ type: "Plaquette", controle: "Additions et soldes de la plaquette (totaux, brut − amortissements = net)", statut: nbArith ? "Anomalie" : "OK", detail: nbArith ? `${nbArith} anomalie(s) détaillée(s) ci-dessous` : "tous les totaux et soldes lus sont cohérents (à ±1 € d'arrondi par ligne)" });
  const resPl = matched.get("P_RES")?.values?.n, resCR = matched.get("R_BEN")?.values?.n;
  if (resPl !== undefined || resCR !== undefined) controles.push({ type: "Plaquette", controle: "Résultat : bilan passif = compte de résultat = FEC", statut: [resPl, resCR].every(v => v === undefined || Math.abs(v - fec.resultat) <= 1) ? "OK" : "Anomalie", detail: `Bilan ${resPl !== undefined ? eur(resPl) : "—"} · Compte de résultat ${resCR !== undefined ? eur(resCR) : "—"} · FEC ${eur(fec.resultat)}` });
  const aff = controleAffectation(postes, balance);
  controles.push({ type: "Mapping", controle: "Comptes non affectés à un poste", statut: aff.nonAffectes.length ? "Anomalie" : "OK", detail: aff.nonAffectes.length ? aff.nonAffectes.map(b => `${b.compte} ${b.lib} (${eur(b.solde)})`).join(" ; ") + " → compléter la feuille Mapping" : "tous les comptes soldés sont affectés" });
  controles.push({ type: "Mapping", controle: "Comptes affectés à plusieurs postes", statut: aff.doublons.length ? "Anomalie" : "OK", detail: aff.doublons.length ? aff.doublons.map(b => `${b.compte} → ${b.postes.join(" / ")}`).join(" ; ") : "aucun" });
  const lus = postes.filter(p => matched.has(p.id)).length;
  controles.push({ type: "Plaquette", controle: "Lecture des états financiers", statut: "Info", detail: `${lus} lignes de la plaquette reconnues et rattachées à un poste` });
  const stats = { concordants: rows.filter(r => r.statut.startsWith("Concordant")).length, ecarts: rows.filter(r => r.statut.startsWith("Écart")).length, autres: rows.filter(r => !r.statut.startsWith("Concordant") && !r.statut.startsWith("Écart")).length };
  return { rows, controles, stats };
}
