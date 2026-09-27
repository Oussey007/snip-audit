// Extraction du texte dans une zone et interprétation des montants (formats français et anglais).

// Rectangle PDF normalisé [x1,y1,x2,y2] (origine en bas à gauche, unités PDF)
export function norm(r) { return [Math.min(r[0], r[2]), Math.min(r[1], r[3]), Math.max(r[0], r[2]), Math.max(r[1], r[3])]; }

// Boîte d'un élément de texte pdf.js en coordonnées PDF
export function itemBox(it) {
  const [a, b, c, d, e, f] = it.transform;
  const h = it.height || Math.hypot(c, d) || Math.hypot(a, b);
  const w = it.width || 0;
  return [e, f - h * 0.2, e + w, f + h * 0.85];
}
function overlapRatio(box, r) {
  const ix = Math.max(0, Math.min(box[2], r[2]) - Math.max(box[0], r[0]));
  const iy = Math.max(0, Math.min(box[3], r[3]) - Math.max(box[1], r[1]));
  const area = Math.max(1e-6, (box[2] - box[0]) * (box[3] - box[1]));
  return (ix * iy) / area;
}

// Texte contenu dans le rectangle (éléments recouverts à ≥ 50 %), rangé par lignes
export function textInRect(items, rect) {
  const r = norm(rect);
  const hits = [];
  for (const it of items) {
    if (!it.str || !it.str.trim()) continue;
    const box = itemBox(it);
    if (overlapRatio(box, r) >= 0.5) {
      // découpe partielle horizontale : ne garde que les caractères situés dans le cadre
      let s = it.str;
      const w = box[2] - box[0];
      if (w > 0 && s.length > 1 && (box[0] < r[0] - 1 || box[2] > r[2] + 1)) {
        const cw = w / s.length;
        const i0 = Math.max(0, Math.floor((r[0] - box[0]) / cw + 0.3));
        const i1 = Math.min(s.length, Math.ceil((r[2] - box[0]) / cw - 0.3));
        s = s.slice(i0, i1);
      }
      // ignore les doublons superposés (texte « gras » imprimé deux fois)
      if (s.trim() && !hits.some(q => q.s === s && Math.abs(q.x - box[0]) < 1.5 && Math.abs(q.y - box[1]) < 1.5)) hits.push({ s, x: box[0], y: box[1], h: box[3] - box[1] });
    }
  }
  hits.sort((p, q) => (Math.abs(p.y - q.y) < Math.max(p.h, q.h) * 0.5 ? p.x - q.x : q.y - p.y));
  const lines = [];
  for (const h of hits) {
    const last = lines[lines.length - 1];
    if (last && Math.abs(last.y - h.y) < h.h * 0.5) last.t.push(h.s); else lines.push({ y: h.y, t: [h.s] });
  }
  return lines.map(l => l.t.join(" ").replace(/\s+/g, " ").trim()).join("\n").trim();
}

const DATE_RE = /\b\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}\b/g;
const NUM_RE = /[-−]?\s?(?:\d{1,4}(?:[   .]\d{3})+|\d+)(?:[,.]\d{1,})?/g;

// Convertit une chaîne en nombre : « 1 234,56 », « 1.234,56 », « 1,234.56 », « 1234.56 », « -12 »
export function toNumber(raw) {
  let s = String(raw).trim().replace(/[€$£%]|EUR/gi, "").replace(/[   ]/g, "").replace("−", "-");
  if (!s) return null;
  const neg = /^\(.*\)$/.test(s) || s.startsWith("-");
  s = s.replace(/[()\-]/g, "");
  if (!/^[\d.,]+$/.test(s)) return null;
  const lc = s.lastIndexOf(","), ld = s.lastIndexOf(".");
  if (lc > -1 && ld > -1) {
    if (lc > ld) s = s.replace(/\./g, "").replace(",", "."); else s = s.replace(/,/g, "");
  } else if (lc > -1) {
    // virgule = séparateur décimal (usage français) ; plusieurs virgules = séparateurs de milliers
    s = s.split(",").length > 2 ? s.replace(/,/g, "") : s.replace(",", ".");
  } else if (ld > -1) {
    const parts = s.split(".");
    if (parts.length > 2 || (parts[1].length === 3 && parts[0].length <= 3 && parts[0] !== "0")) s = s.replace(/\./g, "");
  }
  const v = Number(s);
  if (!isFinite(v)) return null;
  return neg ? -v : v;
}

// Tous les montants d'un texte (les dates sont ignorées)
export function numbersIn(text) {
  const t = String(text).replace(DATE_RE, " ");
  const out = [];
  for (const m of t.matchAll(NUM_RE)) { const v = toNumber(m[0]); if (v !== null) out.push(v); }
  return out;
}

// Valeur d'un snip texte : nombre si la zone ne contient qu'un montant, sinon texte
export function valueFromText(text) {
  const t = text.replace(/\n/g, " ").trim();
  const nums = numbersIn(text.split("\n").join(" | "));
  const letters = t.replace(/EUR|HT|TTC/gi, "").replace(/[\d\s.,€$£%\-−()|:]/g, "");
  if (nums.length === 1 && letters.length === 0) return nums[0];
  return t;
}

export function round2(v) { return Math.round(v * 100) / 100; }

// Variantes d'écriture d'une valeur de cellule pour la recherche dans la pièce
export function searchVariants(v) {
  if (v === null || v === undefined || v === "") return [];
  if (typeof v === "number" && isFinite(v)) {
    if (Math.abs(v) < 1) return [];
    const a = Math.abs(v);
    const fixed = a.toFixed(2);
    const [ent, dec] = fixed.split(".");
    const grp = ent.replace(/\B(?=(\d{3})+(?!\d))/g, " ");
    const set = new Set([`${grp},${dec}`, `${ent},${dec}`, `${ent}.${dec}`, `${ent.replace(/\B(?=(\d{3})+(?!\d))/g, ",")}.${dec}`]);
    if (Number.isInteger(a)) { set.add(ent); set.add(grp); }
    return [...set].map(s => s.replace(/\s/g, ""));
  }
  const s = String(v).trim();
  if (s.length < 4 || s.length > 60 || s === "n.d." || s.startsWith("=")) return [];
  return [s.replace(/\s/g, "").toLowerCase()];
}
