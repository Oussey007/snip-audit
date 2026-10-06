// Lecture de l'annexe d'une plaquette : pages, tableaux (en-têtes, colonnes, lignes) et montants cités dans le texte.
import { analysePage, toAmount } from "./plaquette.js?v=10";
import { normLabel } from "./pcg.js?v=10";

const START_TITLES = /^(annexe|regles et methodes comptables|faits caracteristiques|faits marquants|notes sur le bilan|notes sur le compte de resultat|autres informations|engagements|evenements posterieurs|informations complementaires|complements d informations|notes annexes)/;
const STOP_TITLES = /^(soldes intermediaires|detail des comptes|bilan detaille|compte de resultat detaille|liasse|liste simplifiee|liste des immobilisations|rapport|tableau de financement|tableau des flux|attestation)/;
const HEADER_WORDS = /\b(debut|fin|augmentation|augmentations|diminution|diminutions|ouverture|cloture|dotations?|reprises?|montant|brut|echeance|an|ans|exercice|exploitation|financieres?|exceptionnel(le)?s?|charges|produits|effectif|employe|valeur|net|n 1|cumul)\b/;

// Titre de page : lignes du haut, sans montant
function pageTitle(pg) {
  return pg.lines.filter(l => l.y > pg.height * 0.84 && !l.amounts.length).map(l => normLabel(l.label)).filter(Boolean);
}

// Pages de l'annexe : après le compte de résultat, jusqu'au premier titre de fin (SIG, détail des comptes, liasse…)
export function annexPages(pages, statementPages) {
  const A = pages.map(analysePage);
  const last = Math.max(0, ...Object.values(statementPages || {}).flat());
  const out = []; let started = false;
  for (const pg of A) {
    if (pg.n <= last) continue;
    const t = pageTitle(pg);
    if (t.some(x => STOP_TITLES.test(x)) && started) break;
    if (t.some(x => STOP_TITLES.test(x))) continue;
    if (t.some(x => START_TITLES.test(x))) started = true;
    if (started && pg.lines.length > 4) out.push(pg);
  }
  return out;
}

// Colonnes d'en-tête : sur chaque ligne, segments de mots rapprochés ; puis regroupement des segments alignés d'une ligne à l'autre
function headerColumns(hdrLines) {
  const segs = [];
  hdrLines.forEach((l, li) => {
    const ts = l.toks.slice().sort((a, b) => a.x0 - b.x0); let cur = null;
    for (const t of ts) {
      if (cur && t.x0 - cur.x1 < 7) { cur.x1 = t.x1; cur.words.push(t.t); }
      else { cur = { x0: t.x0, x1: t.x1, words: [t.t], li }; segs.push(cur); }
    }
  });
  const cols = [];
  for (const sg of segs.sort((a, b) => a.li - b.li || a.x0 - b.x0)) {
    const c = sg.li > 0 && cols.find(c => { const ov = Math.min(c.x1, sg.x1) - Math.max(c.x0, sg.x0); return ov > 0.4 * Math.min(c.x1 - c.x0, sg.x1 - sg.x0) || Math.abs((c.x0 + c.x1) / 2 - (sg.x0 + sg.x1) / 2) < 12; });
    if (c && !c.lines.has(sg.li)) { c.x0 = Math.min(c.x0, sg.x0); c.x1 = Math.max(c.x1, sg.x1); c.label += " " + sg.words.join(" "); c.lines.add(sg.li); }
    else cols.push({ x0: sg.x0, x1: sg.x1, label: sg.words.join(" "), lines: new Set([sg.li]) });
  }
  return cols.map(({ x0, x1, label }) => ({ x0, x1, label })).sort((a, b) => a.x0 - b.x0);
}

// Rôle d'une colonne selon son libellé
function colRole(type, lab) {
  const t = normLabel(lab);
  if (type === "immobilisations" || type === "amortissements") {
    if (/debut/.test(t)) return "debut"; if (/augment/.test(t)) return "aug"; if (/diminut/.test(t)) return "dim"; if (/fin/.test(t)) return "fin";
  }
  if (type === "provisions") {
    if (/ouverture|debut/.test(t)) return "ouv"; if (/non utilis/.test(t)) return "repnu"; if (/utilis/.test(t)) return "repu"; if (/reprise|diminut/.test(t)) return "rep"; if (/dotation|augment/.test(t)) return "dot"; if (/cloture|fin/.test(t)) return "clo";
  }
  if (type === "creances" || type === "dettes") {
    if (/brut|montant/.test(t) && !/echeance/.test(t)) return "brut";
    if (/plus de cinq|plus de 5/.test(t)) return "p5";
    if (/cinq ans au plus|5 ans au plus|et cinq|et 5/.test(t)) return "p15";
    if (/plus d un an|plus d 1 an|plus de 1 an/.test(t)) return "p1";
    if (/au plus|moins d un an|1 an au plus/.test(t)) return "m1";
  }
  if (type === "cca" || type === "pca") { if (/exploitation/.test(t)) return "expl"; if (/financ/.test(t)) return "fin"; if (/exception/.test(t)) return "exc"; }
  return /montant|total/.test(t) ? "montant" : t.slice(0, 20) || "montant";
}

function tableType(title, hdr) {
  const t = normLabel(title), h = normLabel(hdr);
  if (/amortissement/.test(t) && /debut|fin/.test(h)) return "amortissements";
  if (/immobilisation/.test(t) && /debut|fin/.test(h) && /augment|acquisit/.test(h)) return "immobilisations";
  if (/provision|depreciation/.test(t) && /ouverture|cloture|dotation/.test(h)) return "provisions";
  if (/ouverture/.test(h) && /cloture/.test(h)) return "provisions";
  if (/dette/.test(t) || (/echeance/.test(h) && /cinq|5 ans/.test(h))) return "dettes";
  if (/creance/.test(t) || (/echeance/.test(h) && /un an/.test(h))) return "creances";
  if (/produits a recevoir/.test(t)) return "par";
  if (/charges a payer/.test(t)) return "cap";
  if (/charges constatees d avance/.test(t)) return "cca";
  if (/produits constates d avance/.test(t)) return "pca";
  if (/effectif/.test(t) || /effectif/.test(h)) return "effectif";
  if (/chiffre d affaires|ventilation/.test(t)) return "ca";
  return "autre";
}

const isHeaderLine = l => !l.amounts.length && l.toks.length && l.lx0 > 150 && HEADER_WORDS.test(normLabel(l.label));
const isParagraph = l => !l.amounts.length && l.label.length > 70;

// Tableaux d'une page
export function pageTables(pg) {
  const tables = []; let cur = null; const L = pg.lines.filter(l => l.y < pg.height * 0.9 && l.y > pg.height * 0.05);
  for (let i = 0; i < L.length; i++) {
    const l = L[i];
    if (isHeaderLine(l)) {
      const hdr = [l]; while (i + 1 < L.length && isHeaderLine(L[i + 1]) && hdr[hdr.length - 1].y - L[i + 1].y < 20) hdr.push(L[++i]);
      // titre : dernière ligne de texte à gauche au-dessus de l'en-tête (hors phrase longue)
      let title = ""; for (let k = L.indexOf(hdr[0]) - 1; k >= 0; k--) { const x = L[k]; if (!x.amounts.length && x.lx0 < 120 && x.label.length <= 70 && hdr[0].y - x.y < 140 - 0) { title = x.label; break; } }
      const hdrText = hdr.map(h => h.label).join(" ");
      const type = tableType(title, hdrText);
      cur = { page: pg.n, title, type, header: hdrText, cols: (cs => cs.map(c => ({ ...c, role: cs.length === 1 ? "montant" : colRole(type, c.label) })))(headerColumns(hdr)), rows: [], y0: hdr[0].y, ctx: "" };
      tables.push(cur); continue;
    }
    // tableau sans ligne d'en-tête reconnue (ex. « Produits à recevoir » + « Montant ») : un titre suivi de lignes chiffrées
    if (!cur && l.amounts.length && i > 0) {
      let title = ""; for (let k = i - 1; k >= 0; k--) { const x = L[k]; if (!x.amounts.length && x.lx0 < 120) { title = x.label; break; } }
      cur = { page: pg.n, title, type: tableType(title, ""), header: "", cols: [{ x0: l.amounts[l.amounts.length - 1].x0, x1: l.amounts[l.amounts.length - 1].x1, label: "Montant", role: "montant" }], rows: [], y0: l.y, ctx: "" };
      tables.push(cur);
    }
    if (!cur) continue;
    if (isParagraph(l) && cur.rows.some(r => Object.keys(r.cells).length)) { cur = null; continue; }
    if (!l.amounts.length) {
      const nl = normLabel(l.label);
      if (/actif immobilise|creances de l actif immobilise/.test(nl) && /creance|:$/.test(l.label.toLowerCase() + ":")) cur.ctx = "immobilise";
      if (/actif circulant/.test(nl)) cur.ctx = "circulant";
      cur.rows.push({ label: l.label, nl, cells: {}, y: l.y, lx0: l.lx0, ctx: cur.ctx, empty: true });
      continue;
    }
    // libellé sur deux lignes : « Dettes fournisseurs et comptes » / « rattachés 872 106 »
    let label = l.label; const prev = cur.rows[cur.rows.length - 1];
    if (prev && prev.empty && (!label || /^[a-zà-ÿ]/.test(label)) && prev.y - l.y < 16) { label = (prev.label + " " + label).trim(); cur.rows.pop(); }
    if (!label && prev) label = prev.label;
    const cells = {};
    for (const a of l.amounts) {
      const center = (a.x0 + a.x1) / 2;
      let best = null, bd = 1e9;
      for (const c of cur.cols) {
        const inside = center >= c.x0 - 30 && center <= c.x1 + 30;
        const d = inside ? Math.abs(a.x1 - c.x1) : 1000 + Math.abs(center - (c.x0 + c.x1) / 2);
        if (d < bd) { bd = d; best = c; }
      }
      if (!best) continue;
      let role = best.role; if (cells[role]) role = role + "_2";
      cells[role] = { value: a.value, raw: a.raw, box: [a.x0 - 1, a.y - a.h * 0.25, a.x1 + 1, a.y + a.h * 0.9] };
    }
    cur.rows.push({ label, nl: normLabel(label), cells, y: l.y, lx0: l.lx0, ctx: cur.ctx, dash: /^[-–•]/.test(l.label.trim()) || /^[-–•]/.test(label.trim()) });
  }
  // libellés poursuivis sur la ligne suivante sans montant (« Autres provisions pour risques » / « et charges »)
  for (const t of tables) for (let i = t.rows.length - 2; i >= 0; i--) {
    const r = t.rows[i], n = t.rows[i + 1];
    if (!r.empty && n.empty && /^[a-zà-ÿ]/.test(n.label) && r.y - n.y < 16) { r.label += " " + n.label; r.nl = normLabel(r.label); t.rows.splice(i + 1, 1); }
  }
  for (const t of tables) for (const r of t.rows) r.info = /^(dont|\(\*)/.test(r.label.trim().toLowerCase()) || /^\(\d\) dont/.test(r.label.toLowerCase());
  // un tableau sans en-tête doit compter au moins deux lignes chiffrées (sinon : phrase se terminant par un montant)
  return tables.filter(t => { const n = t.rows.filter(r => Object.keys(r.cells).length).length; return n >= (t.header ? 1 : 2); });
}

// Montants cités dans le texte : total du bilan, résultat, total des créances / dettes, capital
export function textFacts(pg) {
  const out = [];
  pg.lines.forEach((l, i) => {
    const toks = l.toks; const txt = normLabel(l.label);
    const amountAfter = (re) => {
      const m = re.exec(txt); if (!m) return null;
      // premier montant à droite du mot-clé
      const words = l.label.split(/\s+/); let acc = "", idx = -1;
      for (let k = 0; k < toks.length; k++) { acc = normLabel(toks.slice(0, k + 1).map(t => t.t).join(" ")); if (acc.length >= m.index + m[0].length - 1) { idx = k + 1; break; } }
      for (let k = Math.max(0, idx); k < toks.length; k++) if (toks[k].num && /\d{3}|\d,\d/.test(toks[k].t)) return { value: toAmount(toks[k].t), raw: toks[k].t, box: [toks[k].x0 - 1, l.y - l.h * 0.25, toks[k].x1 + 1, l.y + l.h * 0.9] };
      return null;
    };
    const push = (kind, re, label) => { const a = amountAfter(re); if (a && a.value !== null) out.push({ kind, label, page: pg.n, texte: l.label, ...a }); };
    push("totalBilan", /total (est|s eleve) (de|a)/, "Total du bilan cité dans l'annexe");
    if (/benefice de|perte de/.test(txt)) push(/perte de/.test(txt) ? "perte" : "benefice", /(benefice|perte) de/, "Résultat cité dans l'annexe");
    push("totalCreances", /total des creances.*?s eleve a/, "Total des créances cité dans l'annexe");
    push("totalDettes", /total des dettes.*?s eleve a/, "Total des dettes cité dans l'annexe");
    const raw = l.label.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[\u00a0\u202f]/g, " ");
    const cap = /capital social d.un montant de ([\d .,]+?)\s*(?:€|euros?)?\s*(?:decompose|divise) en ([\d .]+?) (?:titres|actions|parts)[^\d]*?nominale de ([\d .,]+)/.exec(raw);
    if (cap) { const a = amountAfter(/capital social d un montant de/); if (a) out.push({ kind: "capital", label: "Capital social cité dans l'annexe", page: pg.n, texte: l.label, ...a, titres: toAmount(cap[2]), nominal: toAmount(cap[3].replace(/[.,]$/, "")) }); }
    const eff = /effectif moyen[^0-9]*(\d+) personnes/.exec(txt);
    if (eff) out.push({ kind: "effectif", label: "Effectif moyen cité dans l'annexe", page: pg.n, texte: l.label, value: Number(eff[1]), box: null });
  });
  return out;
}

export function readAnnex(pages, statementPages) {
  const P = annexPages(pages, statementPages);
  return { pages: P.map(p => p.n), tables: P.flatMap(pageTables), facts: P.flatMap(textFacts) };
}
