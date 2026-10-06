// Contrôles de l'annexe : calculs internes des tableaux, concordance avec les comptes annuels et avec la balance (FEC).
import { comptesDe, normLabel } from "./pcg.js?v=10";

const r2 = v => Math.round(v * 100) / 100;
const eur = v => (v ?? 0).toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const COLNAME = { debut: "Début d'exercice", aug: "Augmentations", dim: "Diminutions", fin: "Fin d'exercice", ouv: "Ouverture", dot: "Dotations", repu: "Reprises utilisées", repnu: "Reprises non utilisées", rep: "Reprises", clo: "Clôture", brut: "Montant brut", m1: "À un an au plus", p1: "À plus d'un an", p15: "De 1 à 5 ans", p5: "À plus de 5 ans", montant: "Montant", expl: "Exploitation", fin_: "Financier", exc: "Exceptionnel" };
const cn = k => COLNAME[k] || (k.charAt(0).toUpperCase() + k.slice(1));

const TYPE_SIGN = { immobilisations: 1, amortissements: -1, provisions: -1, creances: 1, dettes: -1, par: 1, cap: -1, cca: 1, pca: -1, effectif: 1, ca: -1, autre: 1 };
const FACT_SIGN = { totalBilan: 1, benefice: -1, perte: -1, totalDettes: -1, totalCreances: 1, capital: -1, effectif: 1 };
const GROUPS = {
  incorp: { re: /^immobilisations incorporelles/, postes: ["A_FE", "A_RD", "A_CBL", "A_FC", "A_AII", "A_AAII"], lib: "immobilisations incorporelles" },
  corp: { re: /^immobilisations corporelles(?! en cours)/, postes: ["A_TER", "A_CONS", "A_ITMOI", "A_AIC", "A_ICEC", "A_AAIC"], lib: "immobilisations corporelles" },
  fin: { re: /^immobilisations financieres/, postes: ["A_PME", "A_AP", "A_CRP", "A_ATI", "A_PRET", "A_AIF"], lib: "immobilisations financières" },
  total: { re: /^(actif immobilise|total)/, postes: null, lib: "actif immobilisé" }
};
GROUPS.total.postes = [...GROUPS.incorp.postes, ...GROUPS.corp.postes, ...GROUPS.fin.postes];

const PROV_LINES = [
  [/litige/, "1511"], [/garantie/, "1512"], [/marches a terme/, "1513"], [/amende|penalite/, "1514"], [/pertes? de change/, "1515"],
  [/pension|retraite|depart|fin de carriere|idr/, "153"], [/impot/, "155"], [/renouvellement/, "156"], [/gros entretien|grosses reparations|grandes revisions/, "1572"],
  [/conges a payer/, "1582"], [/autres provisions pour risques|autres provisions/, "1516 1518 154 157 -1572 158 -1582"],
  [/clients|comptes clients|creances clients/, "491"], [/stocks|en.cours/, "39"], [/autres creances|comptes de tiers/, "495 496"], [/valeurs mobilieres/, "59"],
  [/titres|participations|immobilisations financieres/, "296 297"], [/immobilisations/, "290 291 293"]
];
const DETTE_POSTES = [
  [/obligataires convertibles/, "P_EOC"], [/autres emprunts obligataires/, "P_AEO"], [/etablissements? de credit|a l origine|concours bancaires/, "P_EEC"],
  [/groupe et associes|financieres? divers/, "P_EDFD"], [/avances et acomptes recus/, "P_AAR"], [/fournisseurs/, "P_FOUR"],
  [/dettes sur immobilisations/, "P_DIMMO"], [/fiscales et sociales|personnel|securite sociale|organismes sociaux|impots sur les benefices|valeur ajoutee|autres impots|obligations cautionnees|etat et autres/, "P_DFS"],
  [/produits constates d avance/, "P_PCA"], [/autres dettes/, "P_AD"]
];
const CREANCE_POSTES = (r) => {
  const t = r.nl;
  if (/clients/.test(t)) return "A_CLI";
  if (/charges constatees d avance/.test(t)) return "A_CCA";
  if (/capital souscrit/.test(t)) return "A_CSANV";
  if (r.ctx === "immobilise") { if (/prets/.test(t)) return "A_PRET"; if (/rattachees/.test(t)) return "A_CRP"; if (/autres/.test(t)) return "A_AIF"; return null; }
  if (/^autres|personnel|securite sociale|organismes sociaux|impots|taxe|tva|valeur ajoutee|groupe et associes|debiteurs divers|etat et autres/.test(t)) return "A_AC";
  return null;
};

export function checkAnnex(annex, ctx) {
  const { postes, matched, balance, statements } = ctx;
  const out = [];
  const V = (id, k) => matched.get(id)?.values?.[k];
  const posteLib = id => postes.find(p => p.id === id)?.libelle || id;
  const sumP = (ids, k) => { let s = 0, any = false; for (const id of ids) { const v = V(id, k); if (v !== undefined) { s += v; any = true; } } return any ? s : null; };
  const refBox = (ids, k) => { for (const id of ids) { const l = matched.get(id); if (l && l.values[k] !== undefined) return { page: l.page, box: l.boxes[k] }; } return {}; };
  const regleOf = (ids, field) => ids.map(id => postes.find(p => p.id === id)?.[field]).filter(Boolean).join(" ");
  const fec = (regle, f, sg) => { if (!regle.trim()) return { v: 0, cs: [] }; const cs = comptesDe(regle, balance); return { v: r2(sg * cs.reduce((a, b) => a + b[f], 0)), cs }; };
  const lst = (cs, f, sg) => cs.filter(b => Math.abs(b[f]) > 0.004).slice(0, 12).map(b => `${b.compte} ${b.lib} (${eur(sg * b[f])})`).join(" ; ");

  const add = (o) => {
    const ecart = o.annexe === null || o.ref === null ? null : r2(o.ref - o.annexe);
    const tol = o.tol ?? 1;
    let statut = o.statut;
    if (!statut) {
      if (o.annexe === null) statut = "Absent de l'annexe";
      else if (o.ref === null) statut = "Référence introuvable";
      else if (Math.abs(ecart) < 0.005) statut = "Concordant";
      else if (Math.abs(ecart) <= tol) statut = "Concordant (arrondi)";
      else statut = o.analyser ? "Écart – à analyser" : "Écart";
    }
    if (statut.startsWith("Concordant")) o.commentaire = o.commentaireOK || "";
    // signe comptable de la ligne (débit +, crédit −), appliqué au montant de l'annexe et à sa référence
    const sg = o.sg ?? (TYPE_SIGN[o.type] ?? 1);
    out.push({ ...o, ecart, statut, sg });
  };
  const cellOf = (r, k) => r.cells[k] || null;
  const ann = (t, r, k, extra) => { const c = cellOf(r, k); return { note: t.title, type: t.type, page: t.page, ligne: r.label, colonne: cn(k), annexe: c ? c.value : null, box: c ? c.box : null, ...extra }; };

  for (const t of annex.tables) {
    const rows = t.rows.filter(r => !r.empty);
    const totIdx = rows.findIndex(r => /^total\b/.test(r.nl) || /^actif immobilise/.test(r.nl));
    const body = rows.filter((r, i) => !r.info && (totIdx < 0 || i < totIdx) && !/^total\b/.test(r.nl));
    const keys = [...new Set(rows.flatMap(r => Object.keys(r.cells)))];

    // ---- 1) calculs horizontaux
    const horiz = (r, plus, minus, res, lib) => {
      if (!cellOf(r, res) && !plus.concat(minus).some(k => cellOf(r, k))) return;
      const v = k => cellOf(r, k)?.value ?? 0;
      const calc = r2(plus.reduce((a, k) => a + v(k), 0) - minus.reduce((a, k) => a + v(k), 0));
      add({ ...ann(t, r, res), controle: "Calcul", reference: lib, ref: calc, tol: 2 });
    };
    if (t.type === "immobilisations" || t.type === "amortissements") rows.filter(r => !r.info).forEach(r => horiz(r, ["debut", "aug"], ["dim"], "fin", "Début + augmentations − diminutions"));
    if (t.type === "provisions") rows.filter(r => !r.info && (totIdx < 0 || rows.indexOf(r) <= totIdx)).forEach(r => horiz(r, ["ouv", "dot"], ["repu", "repnu", "rep"], "clo", "Ouverture + dotations − reprises"));
    if (t.type === "creances") rows.filter(r => !r.info).forEach(r => horiz(r, ["m1", "p1"], [], "brut", "Somme des échéances"));
    if (t.type === "dettes") rows.filter(r => !r.info && rows.indexOf(r) <= (totIdx < 0 ? rows.length : totIdx)).forEach(r => horiz(r, ["m1", "p15", "p5"], [], "brut", "Somme des échéances"));

    // ---- 2) calculs verticaux
    if (t.type === "immobilisations" || t.type === "amortissements") {
      let dash = [], subs = [];
      for (const r of rows.filter(r => !r.info)) {
        if (r.dash) { dash.push(r); continue; }
        const isTot = /^(actif immobilise|total)/.test(r.nl);
        const parts = isTot ? (subs.length ? subs : dash) : dash;
        if (parts.length) for (const k of keys) {
          if (!cellOf(r, k) && !parts.some(p => cellOf(p, k))) continue;
          add({ ...ann(t, r, k), controle: "Calcul", reference: `Somme des lignes (${parts.length})`, ref: r2(parts.reduce((a, p) => a + (cellOf(p, k)?.value ?? 0), 0)), tol: Math.max(1, Math.ceil(parts.length / 2)) });
        }
        if (isTot) { dash = []; subs = []; } else { subs.push(r); dash = []; }
      }
    } else if (totIdx >= 0) {
      const tot = rows[totIdx];
      for (const k of keys) {
        if (!cellOf(tot, k) && !body.some(p => cellOf(p, k))) continue;
        add({ ...ann(t, tot, k), controle: "Calcul", reference: `Somme des lignes (${body.length})`, ref: r2(body.reduce((a, p) => a + (cellOf(p, k)?.value ?? 0), 0)), tol: Math.max(1, Math.ceil(body.length / 2)) });
      }
      // répartition des dotations et reprises (exploitation / financières / exceptionnelles)
      if (t.type === "provisions") {
        const rep = rows.slice(totIdx + 1).filter(r => /^(exploitation|financieres?|exceptionnelles?)$/.test(r.nl));
        if (rep.length) {
          const sumK = ks => r2(rep.reduce((a, r) => a + ks.reduce((b, k) => b + (cellOf(r, k)?.value ?? 0), 0), 0));
          if (cellOf(tot, "dot")) add({ ...ann(t, tot, "dot"), ligne: "Total (répartition des dotations)", controle: "Calcul", reference: "Répartition exploitation + financières + exceptionnelles", ref: sumK(["dot"]), tol: 2, commentaire: "La répartition des dotations par nature ne correspond pas au total des dotations du tableau." });
          const repK = ["repu", "repnu", "rep"].filter(k => cellOf(tot, k));
          if (repK.length) add({ note: t.title, type: t.type, page: t.page, ligne: "Total (répartition des reprises)", colonne: "Reprises", annexe: r2(repK.reduce((a, k) => a + cellOf(tot, k).value, 0)), box: cellOf(tot, repK[0]).box, controle: "Calcul", reference: "Répartition exploitation + financières + exceptionnelles", ref: sumK(["repu", "repnu", "rep"]), tol: 2 });
        }
      }
    }

    // ---- 3) concordance avec les comptes annuels et la balance
    if (t.type === "immobilisations" || t.type === "amortissements") {
      const amort = t.type === "amortissements";
      for (const r of rows.filter(r => !r.dash && !r.info)) {
        const g = Object.values(GROUPS).find(g => g.re.test(r.nl)); if (!g) continue;
        const k = amort ? "amort" : "brut";
        if (cellOf(r, "fin") || sumP(g.postes, k) !== null) {
          const rb = refBox(g.postes, k);
          add({ ...ann(t, r, "fin"), controle: "Comptes annuels", reference: `Bilan actif – ${amort ? "amortissements" : "brut"} des ${g.lib}`, ref: sumP(g.postes, k) ?? 0, refPage: rb.page, refBox: rb.box, tol: 2 });
        }
        const regle = regleOf(g.postes, amort ? "amort" : "regle");
        const sg = amort ? -1 : 1;
        for (const [col, f, s, lib] of [["debut", "an", sg, "solde d'ouverture"], ["aug", amort ? "credit" : "debit", 1, amort ? "mouvements crédit" : "mouvements débit"], ["dim", amort ? "debit" : "credit", 1, amort ? "mouvements débit" : "mouvements crédit"]]) {
          const x = fec(regle, f, s);
          if (!cellOf(r, col) && Math.abs(x.v) < 0.5) continue;
          add({ ...ann(t, r, col), controle: "Balance (FEC)", reference: `FEC – ${lib} des comptes d'${amort ? "amortissement des " : ""}${g.lib}`, ref: x.v, analyser: true, tol: 2, commentaire: `Comptes : ${lst(x.cs, f, s)}. Les mouvements du FEC incluent aussi les virements de poste à poste et les corrections, qui peuvent expliquer un écart.` });
        }
        if (amort && /^(actif immobilise|total)/.test(r.nl)) {
          const cr = matched.get("R_DAI");
          const x = fec(postes.find(p => p.id === "R_DAI")?.regle || "6811 6812", "solde", 1);
          add({ ...ann(t, r, "aug"), controle: "Comptes annuels", reference: "Compte de résultat – dotations aux amortissements sur immobilisations", ref: cr?.values?.n ?? null, refPage: cr?.page, refBox: cr?.boxes?.n, tol: 2, commentaire: `Composition de la ligne du compte de résultat : ${lst(x.cs, "solde", 1)}.` });
        }
      }
    }
    if (t.type === "provisions") {
      const hasDep = rows.some(r => /clients|stocks|creances|titres|immobilisations|valeurs mobilieres/.test(r.nl));
      for (const r of body) {
        const m = PROV_LINES.find(([re]) => re.test(r.nl)); if (!m) continue;
        for (const [col, f, s, lib] of [["ouv", "an", -1, "solde d'ouverture"], ["dot", "credit", 1, "mouvements crédit"], ["clo", "solde", -1, "solde de clôture"]]) {
          const x = fec(m[1], f, s);
          if (!cellOf(r, col) && Math.abs(x.v) < 0.5) continue;
          add({ ...ann(t, r, col), controle: "Balance (FEC)", reference: `FEC – ${lib} (comptes ${m[1].replace(/ -\d+/g, "")})`, ref: x.v, analyser: true, tol: 2, commentaire: `Comptes : ${lst(x.cs, f, s) || "aucun"}.` });
        }
      }
      if (totIdx >= 0 && !hasDep) {
        const tot = rows[totIdx];
        for (const [col, k, lib] of [["clo", "n", "N"], ["ouv", "n1", "N-1"]]) {
          const l = matched.get("P_TPRC"); if (!l || (!cellOf(tot, col) && l.values[k] === undefined)) continue;
          add({ ...ann(t, tot, col), controle: "Comptes annuels", reference: `Bilan passif – total provisions pour risques et charges (${lib})`, ref: l.values[k] ?? 0, refPage: l.page, refBox: l.boxes[k], tol: 2 });
        }
      }
    }
    if (t.type === "creances" || t.type === "dettes") {
      const map = new Map();
      for (const r of body) {
        const id = t.type === "creances" ? CREANCE_POSTES(r) : (DETTE_POSTES.find(([re]) => re.test(r.nl)) || [])[1];
        if (!id || !cellOf(r, "brut")) continue;
        if (!map.has(id)) map.set(id, []); map.get(id).push(r);
      }
      const k = t.type === "creances" ? "brut" : "n";
      const sec = t.type === "creances" ? "Bilan actif" : "Bilan passif";
      const lignes = [];
      for (const [id, rs] of map) {
        const l = matched.get(id);
        const o = { note: t.title, type: t.type, page: t.page, ligne: rs.map(r => r.label).join(" + "), colonne: cn("brut"), annexe: r2(rs.reduce((a, r) => a + r.cells.brut.value, 0)), box: rs[0].cells.brut.box, controle: "Comptes annuels", reference: `${sec} – ${posteLib(id)}${t.type === "creances" ? " (brut)" : ""}`, ref: l?.values?.[k] ?? 0, refPage: l?.page, refBox: l?.boxes?.[k], tol: Math.max(1, rs.length), side: 1 };
        add(o); lignes.push(out[out.length - 1]);
      }
      // reclassement entre lignes de l'état si le total concorde avec le bilan
      const tot = totIdx >= 0 ? rows[totIdx] : null;
      const totBil = t.type === "dettes" ? matched.get("P_TD")?.values?.n : null;
      if (t.type === "dettes" && tot && cellOf(tot, "brut")) {
        const l = matched.get("P_TD");
        add({ ...ann(t, tot, "brut"), controle: "Comptes annuels", reference: "Bilan passif – total dettes", ref: totBil ?? null, refPage: l?.page, refBox: l?.boxes?.n, tol: 2 });
        // échéances : renvois du bilan passif « dont à plus / à moins d'un an »
        for (const [re, cols, lib] of [[/^dont a moins d un an/, ["m1"], "à moins d'un an"], [/^dont a plus d un an/, ["p15", "p5", "p1"], "à plus d'un an"]]) {
          const sl = (statements.passif || []).find(x => re.test(x.nl.replace(/^\d+ /, "")));
          if (!sl || sl.values.n === undefined) continue;
          const v = cols.filter(c => cellOf(tot, c)); if (!v.length) continue;
          add({ note: t.title, type: t.type, page: t.page, ligne: "Total", colonne: lib, annexe: r2(v.reduce((a, c) => a + cellOf(tot, c).value, 0)), box: cellOf(tot, v[0]).box, controle: "Comptes annuels", reference: `Bilan passif – renvoi « dont ${lib} »`, ref: sl.values.n, refPage: sl.page, refBox: sl.boxes.n, tol: 2 });
        }
      }
      const totOk = t.type !== "dettes" || (tot && cellOf(tot, "brut") && totBil !== undefined && Math.abs(cellOf(tot, "brut").value - totBil) <= 2);
      const ec = lignes.filter(o => o.statut === "Écart");
      if (totOk) for (let i = 0; i < ec.length; i++) for (let j = i + 1; j < ec.length; j++) {
        const a = ec[i], b = ec[j]; if (a.statut !== "Écart" || b.statut !== "Écart") continue;
        if (Math.abs(a.ecart + b.ecart) <= 2) {
          a.statut = b.statut = "Écart – reclassement probable";
          a.commentaire = `Écart compensé par « ${b.ligne} » (${eur(b.ecart)}) : un même montant est classé dans une ligne différente du bilan.`;
          b.commentaire = `Écart compensé par « ${a.ligne} » (${eur(a.ecart)}) : un même montant est classé dans une ligne différente du bilan.`;
        }
      }
    }
    if (["par", "cap", "cca", "pca"].includes(t.type) && totIdx >= 0) {
      const tot = rows[totIdx];
      const tv = r2(Object.values(tot.cells).reduce((a, c) => a + c.value, 0));
      const box = Object.values(tot.cells)[0].box;
      const base = { note: t.title, type: t.type, page: t.page, ligne: tot.label, colonne: "Total", annexe: tv, box };
      if (t.type === "cca" || t.type === "pca") {
        const id = t.type === "cca" ? "A_CCA" : "P_PCA"; const k = t.type === "cca" ? "brut" : "n"; const l = matched.get(id);
        add({ ...base, controle: "Comptes annuels", reference: `${t.type === "cca" ? "Bilan actif" : "Bilan passif"} – ${posteLib(id)}`, ref: l?.values?.[k] ?? 0, refPage: l?.page, refBox: l?.boxes?.[k], tol: 2 });
      } else {
        // comptes de charges à payer / produits à recevoir par définition du PCG (familles 428, 438, 448, 468 selon le sens du solde)
        const regle = t.type === "par" ? "418 4098 428D 438D 448D 468D 5187 5188D 2768" : "408 4198 428C 438C 448C 468C 5186 5188C 16884";
        const sg = t.type === "par" ? 1 : -1;
        const x = fec(regle, "solde", sg);
        add({ ...base, controle: "Balance (FEC)", reference: `FEC – comptes de ${t.type === "par" ? "produits à recevoir" : "charges à payer"} (${regle.replace(/[DC]\b/g, "").split(" ").slice(0, 6).join(", ")}…)`, ref: x.v, analyser: true, tol: 2, commentaire: `Comptes du FEC : ${lst(x.cs, "solde", sg)}. Un écart peut venir de comptes présentés en ${t.type === "par" ? "produits à recevoir" : "charges à payer"} sans en avoir le numéro (ou l'inverse) : à analyser.` });
      }
    }
  }

  // ---- 4) montants cités dans le texte
  const tableTot = (type, k) => { const t = annex.tables.find(t => t.type === type); if (!t) return null; const r = t.rows.find(r => /^total\b/.test(r.nl)); return r?.cells?.[k] ? r.cells[k] : null; };
  for (const f of annex.facts) {
    const base = { note: "Texte de l'annexe", type: "texte", page: f.page, ligne: f.texte.slice(0, 140), colonne: "", annexe: f.value, box: f.box, controle: "Texte", sg: FACT_SIGN[f.kind] ?? 1 };
    const ref = (id, k, lib) => { const l = matched.get(id); add({ ...base, reference: lib, ref: l?.values?.[k] ?? null, refPage: l?.page, refBox: l?.boxes?.[k], tol: 1 }); };
    if (f.kind === "totalBilan") ref("A_TG", "net", "Bilan – total général (net)");
    if (f.kind === "benefice") ref("P_RES", "n", "Bilan passif – résultat de l'exercice");
    if (f.kind === "perte") { const l = matched.get("P_RES"); add({ ...base, annexe: -f.value, reference: "Bilan passif – résultat de l'exercice", ref: l?.values?.n ?? null, refPage: l?.page, refBox: l?.boxes?.n, tol: 1 }); }
    if (f.kind === "totalDettes") ref("P_TD", "n", "Bilan passif – total dettes");
    if (f.kind === "totalCreances") { const c = tableTot("creances", "brut"); const t = annex.tables.find(t => t.type === "creances"); add({ ...base, reference: "État des créances – total brut", ref: c ? c.value : null, refPage: t?.page, refBox: c?.box, tol: 1 }); }
    if (f.kind === "capital") {
      ref("P_CAP", "n", "Bilan passif – capital");
      if (f.titres && f.nominal) add({ ...base, ligne: f.texte.slice(0, 140), colonne: "Nombre de titres × nominal", reference: `${f.titres.toLocaleString("fr-FR")} titres × ${eur(f.nominal)}`, ref: r2(f.titres * f.nominal), controle: "Calcul", tol: 1 });
    }
    if (f.kind === "effectif") { const c = tableTot("effectif", Object.keys(annex.tables.find(t => t.type === "effectif")?.rows.find(r => /^total/.test(r.nl))?.cells || {})[0]); if (c) add({ ...base, reference: "Tableau des effectifs – total", ref: c.value, tol: 0.4 }); }
  }
  const stats = { concordants: out.filter(o => o.statut.startsWith("Concordant")).length, ecarts: out.filter(o => o.statut.startsWith("Écart")).length, autres: out.filter(o => !o.statut.startsWith("Concordant") && !o.statut.startsWith("Écart")).length };
  return { rows: out, stats };
}
