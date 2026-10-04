// Lecture des états financiers d'une plaquette PDF (bilan actif, bilan passif, compte de résultat).
// Entrée : pages au format { n, width, height, items:[{str, x, y, w, h}] } (coordonnées PDF, origine en bas à gauche),
// obtenues avec pdf.js (getTextContent). Module sans dépendance au navigateur : testable seul.
import { normLabel } from "./pcg.js?v=7";

export const SECTION_TITLES = {
  actif: ["bilan actif", "bilan actif suite", "actif", "bilan - actif", "bilan (actif)"],
  passif: ["bilan passif", "bilan passif suite", "passif", "bilan - passif", "bilan (passif)", "bilan passif avant repartition"],
  cr: ["compte de resultat", "compte de resultat suite", "compte de resultat en liste", "compte de resultat de l exercice"]
};

// Découpe un élément pdf.js en mots positionnés
function words(items) {
  const out = [];
  for (const it of items) {
    const s = it.str || ""; if (!s.trim()) continue;
    const cw = s.length ? (it.w || 0) / s.length : 0;
    const re = /\S+/g; let m;
    while ((m = re.exec(s))) out.push({ t: m[0], x0: it.x + m.index * cw, x1: it.x + (m.index + m[0].length) * cw, y: it.y, h: it.h || 8 });
  }
  return out;
}
// Regroupe les mots en lignes (même ordonnée)
function lines(ws) {
  ws = ws.slice().sort((a, b) => b.y - a.y || a.x0 - b.x0);
  const L = [];
  for (const w of ws) {
    const l = L.find(l => Math.abs(l.y - w.y) < Math.max(1.5, Math.min(l.h, w.h) * 0.35));
    if (l) { l.w.push(w); l.h = Math.max(l.h, w.h); } else L.push({ y: w.y, h: w.h, w: [w] });
  }
  L.forEach(l => l.w.sort((a, b) => a.x0 - b.x0));
  return L.sort((a, b) => b.y - a.y);
}

const START = /^[-−(]?\d{1,3}(?:[.,]\d+)?\)?%?$/;             // premier groupe d'un montant
const GROUP = /^\d{3}(?:[.,]\d+)?\)?%?$/;                    // groupe de milliers suivant
const WHOLE = /^[-−(]?\d{1,3}(?:[\s\u00a0\u202f.]\d{3})+(?:,\d+)?\)?$|^[-−(]?\d+(?:,\d+)?\)?$/;
const FOOT = /^\(\d{1,2}\)$|^\([a-z]\)$/i;                  // renvois (1), (a)

export function toAmount(raw) {
  let s = String(raw).replace(/[\s\u00a0\u202f]/g, "").replace("−", "-");
  const neg = /^\(.*\)$/.test(s) || s.startsWith("-");
  s = s.replace(/[()\-%]/g, "");
  if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, "");
  s = s.replace(",", ".");
  const v = Number(s); if (!isFinite(v) || s === "") return null;
  return neg ? -v : v;
}

// Fusionne les groupes de chiffres d'un même montant (« 1 » « 659 » « 554 » -> 1 659 554)
function mergeNumbers(ws) {
  const out = [];
  for (const w of ws) {
    const p = out[out.length - 1];
    const gap = p ? w.x0 - p.x1 : 99;
    if (p && p.num && !p.closed && GROUP.test(w.t) && gap < Math.min(6, w.h * 0.75) && /\d$/.test(p.t) && !/[.,]\d+$/.test(p.t)) {
      p.t += " " + w.t; p.x1 = w.x1; if (/\)$/.test(w.t)) p.closed = true; continue;
    }
    const num = (START.test(w.t) || WHOLE.test(w.t)) && !FOOT.test(w.t);
    out.push({ ...w, num, closed: !num });
  }
  return out;
}

// Analyse une page : lignes avec libellé et montants
export function analysePage(pg) {
  const L = lines(words(pg.items)).map(l => {
    const toks = mergeNumbers(l.w);
    // les montants sont la série de nombres en fin de ligne
    let k = toks.length; while (k > 0 && (toks[k - 1].num || FOOT.test(toks[k - 1].t))) k--;
    const lab = toks.slice(0, k), am = toks.slice(k).filter(t => t.num && !/%$/.test(t.t));
    // un nombre isolé collé au libellé (ex. « Total I ») reste dans le libellé si la ligne n'a pas de colonne de montants
    return {
      y: l.y, h: l.h, label: lab.map(t => t.t).join(" "), toks: toks.map(t => ({ t: t.t, x0: t.x0, x1: t.x1, num: t.num })), lx0: lab.length ? lab[0].x0 : null, lx1: lab.length ? lab[lab.length - 1].x1 : null,
      amounts: am.map(t => ({ raw: t.t, value: toAmount(t.t), x0: t.x0, x1: t.x1, y: l.y, h: l.h })).filter(a => a.value !== null)
    };
  });
  return { n: pg.n, width: pg.width, height: pg.height, lines: L };
}

// Titre de la page : ligne du haut de page égale à un titre connu (« Bilan actif », « Compte de résultat (suite) »…)
export function sectionOf(page) {
  const top = page.lines.filter(l => l.y > page.height * 0.7 && !l.amounts.length);
  for (const l of top) {
    const t = normLabel(l.label);
    for (const [sec, titles] of Object.entries(SECTION_TITLES)) if (titles.includes(t)) return { sec, weak: !t.includes(" ") };
  }
  return null;
}
// Repli si aucun titre reconnu : pages contenant le plus de libellés de postes de la section
function sectionByContent(pages, postes, sec) {
  const keys = new Set(postes.filter(p => p.section === sec).flatMap(p => [p.libelle, ...(p.alias || [])].map(normLabel)));
  const sc = pages.map(pg => ({ n: pg.n, s: pg.lines.filter(l => l.amounts.length && keys.has(normLabel(l.label))).length }));
  const max = Math.max(0, ...sc.map(x => x.s));
  return max >= 5 ? sc.filter(x => x.s >= max * 0.6).map(x => x.n) : [];
}

// Colonnes de montants : regroupement des bords droits (montants alignés à droite)
function anchors(page) {
  const xs = page.lines.flatMap(l => l.amounts.map(a => a.x1)).sort((a, b) => a - b);
  const cl = [];
  for (const x of xs) { const c = cl[cl.length - 1]; if (c && x - c.max < 8) { c.max = x; c.n++; c.sum += x; } else cl.push({ min: x, max: x, n: 1, sum: x }); }
  return cl.filter(c => c.n >= 2).map(c => c.sum / c.n);
}
const ROLES = { actif: ["brut", "amort", "net", "net1"], passif: ["n", "n1", "x"], cr: ["n", "n1", "x"] };

// Rôle de chaque colonne : en-têtes (Brut / Amortissements / Net) si présents, sinon ordre gauche -> droite
function roles(page, sec, anc) {
  if (sec === "actif" && anc.length !== 4) {
    const hdr = page.lines.filter(l => l.y > page.height * 0.6).flatMap(l => l.label ? [l] : []);
    const pos = {};
    for (const l of hdr) { const t = normLabel(l.label); if (/^brut/.test(t)) pos.brut = (l.lx0 + l.lx1) / 2; if (/^amort|deprec/.test(t)) pos.amort = pos.amort ?? (l.lx0 + l.lx1) / 2; }
    if (pos.brut && anc.length === 3) {
      // 3 colonnes : brut, net N, net N-1 (pas d'amortissements) ou brut, amort, net
      return pos.amort ? ["brut", "amort", "net"] : ["brut", "net", "net1"];
    }
  }
  return ROLES[sec];
}

// Extrait les lignes de chaque état : [{section, page, label, nl, values:{brut,amort,net,net1|n,n1}, boxes:{...}}]
export function extractStatements(pages, postes) {
  const res = { actif: [], passif: [], cr: [], pages: {} };
  const A = pages.map(analysePage);
  const found = A.map(p => [p.n, sectionOf(p)]);
  // un titre d'un seul mot (« ACTIF ») ne compte que si aucune page n'a un titre complet (« Bilan actif ») pour cette section
  const strong = new Set(found.filter(([, f]) => f && !f.weak).map(([, f]) => f.sec));
  const secOf = new Map(found.map(([n, f]) => [n, f && (!f.weak || !strong.has(f.sec)) ? f.sec : null]));
  for (const sec of ["actif", "passif", "cr"]) {
    if ([...secOf.values()].includes(sec) || !postes) continue;
    for (const n of sectionByContent(A.filter(p => !secOf.get(p.n)), postes, sec)) secOf.set(n, sec);
  }
  for (const page of A) {
    const sec = secOf.get(page.n); if (!sec) continue;
    (res.pages[sec] = res.pages[sec] || []).push(page.n);
    const anc = anchors(page); const rl = roles(page, sec, anc);
    let pending = null;
    for (const l of page.lines) {
      if (l.y > page.height * 0.9 || l.y < page.height * 0.04) continue;      // en-têtes et pieds de page
      if (!l.amounts.length) { pending = l.label ? l : null; if (l.label) res[sec].push({ section: sec, page: page.n, label: l.label, nl: normLabel(l.label), values: {}, boxes: {}, header: true, y: l.y }); continue; }
      let label = l.label, nl = normLabel(label);
      if (!nl && pending) { label = pending.label; nl = normLabel(label); }
      const values = {}, boxes = {};
      for (const a of l.amounts) {
        let best = -1, d = 1e9; anc.forEach((x, i) => { const dd = Math.abs(x - a.x1); if (dd < d) { d = dd; best = i; } });
        if (best < 0 || d > 14) continue;
        const role = rl[best]; if (!role || role === "x") continue;
        values[role] = a.value; boxes[role] = [a.x0 - 1, a.y - a.h * 0.25, a.x1 + 1, a.y + a.h * 0.9];
      }
      res[sec].push({ section: sec, page: page.n, label, nl, values, boxes, y: l.y });
      pending = null;
    }
  }
  return res;
}

// Associe les lignes lues aux postes, dans l'ordre du document (gère les libellés répétés : « Variations de stock »…)
export function matchPostes(postes, statements) {
  const out = new Map();
  for (const sec of ["actif", "passif", "cr"]) {
    const P = postes.filter(p => p.section === sec).map(p => ({ p, keys: [p.libelle, ...(p.alias || [])].map(normLabel).filter(Boolean) }));
    let cursor = 0;
    const lignes = statements[sec] || [];
    for (const l of lignes) {
      if (!l.nl) continue;
      let found = -1;
      const ok = (key) => l.nl === key || (l.nl.startsWith(key + " ") && /^( ?(i|ii|iii|iv|v|vi|vii|viii|ix|x|n|net|i ii|v vi|vii viii|i a vi|i a v))*$/.test(l.nl.slice(key.length).replace(/[^a-z ]/g, " ").replace(/\s+/g, " ")));
      for (let i = cursor; i < P.length && found < 0; i++) if (!out.has(P[i].p.id) && P[i].keys.some(ok)) found = i;
      if (found < 0) for (let i = 0; i < cursor && found < 0; i++) if (!out.has(P[i].p.id) && P[i].keys.some(ok)) found = i;
      if (found < 0) continue;
      if (l.header && !Object.keys(l.values).length) {
        // libellé sans montant : poste présent mais nul (on le garde pour l'ordre, sans figer le poste si un montant suit)
        cursor = found; continue;
      }
      out.set(P[found].p.id, l); cursor = found + 1;
    }
  }
  return out;
}
