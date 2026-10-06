// Liseuse double : à gauche le montant pointé (plaquette ou annexe), à droite la source du contrôle (onglets par source :
// états financiers, balance FEC, feuilles de cadrage, plaquette N-1, calcul). Utilisée dans le volet et dans la fenêtre détachée.
import { Viewer } from "./viewer.js?v=10";

const ORDER = ["États financiers", "Balance (FEC)", "Feuille immos", "Feuille emprunts", "Plaquette N-1", "Calcul (Σ)", "Référence"];
const rank = k => { const i = ORDER.findIndex(o => k.startsWith(o)); return i < 0 ? ORDER.length : i; };
export const srcKey = s => String(s || "Source").replace(/\s+p\.\s*\d+$/, "").trim() || "Source";
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

// zones : snips de la cellule portant un panneau (G / D) ; manual : snips manuels de la cellule (affichés à gauche)
export function buildModel(zones, manual = []) {
  const L = zones.filter(z => z.pane === "G");
  const groups = [];
  for (const z of zones.filter(z => z.pane === "D")) {
    const key = srcKey(z.source || z.text), id = key + "|" + z.file;
    let g = groups.find(x => x.id === id);
    if (!g) groups.push(g = { id, key, file: z.file, zones: [] });
    g.zones.push(z);
  }
  groups.sort((a, b) => rank(a.key) - rank(b.key));
  const seen = {}; groups.forEach(g => { seen[g.key] = (seen[g.key] || 0) + 1; });
  groups.forEach(g => { g.label = g.key + (seen[g.key] > 1 ? ` · ${g.file.replace(/\.pdf$/i, "").slice(0, 16)}` : ""); });
  const strip = z => ({ file: z.file, page: z.page, rect: z.rect, type: z.type, cell: z.cell, source: z.source || "", text: z.text || "" });
  return { L: L.map(strip), groups: groups.map(g => ({ ...g, zones: g.zones.map(strip) })), manual: manual.map(strip) };
}

export class Split {
  // opts : { getFile(name) → Blob|null, onRect(viewer, page, rect), onActive(viewer), onPage(viewer, n, t), onZoom(viewer, s), layout: "auto"|"side"|"stack" }
  constructor(root, opts = {}) {
    this.root = root; this.o = opts; this.layout = opts.layout || "auto"; this.model = null; this.activeKey = null; this.double = false;
    root.classList.add("split");
    root.innerHTML = `<section class="pane L"><div class="ph"><span class="pt">Montant pointé</span><span class="pl"></span></div><div class="viewer" tabindex="0"></div></section>
      <section class="pane R" hidden><div class="ph rtabs"></div><div class="viewer" tabindex="0"></div></section>`;
    const mk = el => new Viewer(el, {
      onRect: (page, rect) => this.o.onRect && this.o.onRect(this._v(el), page, rect),
      onPage: (n, t) => this.o.onPage && this.o.onPage(this._v(el), n, t),
      onZoom: s => this.o.onZoom && this.o.onZoom(this._v(el), s),
      onActivate: v => this.setActive(v)
    });
    this.left = mk(root.querySelector(".pane.L .viewer"));
    this.right = mk(root.querySelector(".pane.R .viewer"));
    this.active = this.left;
    root.querySelector(".pane.L .ph").hidden = true;
    this._applyLayout();
    if (window.ResizeObserver) new ResizeObserver(() => this._applyLayout()).observe(root);
  }
  _v(el) { return el === this.left?.el ? this.left : this.right; }
  setActive(v) {
    this.active = v;
    this.root.querySelectorAll(".pane").forEach(p => p.classList.toggle("active", this.double && p.contains(v.el)));
    this.o.onActive && this.o.onActive(v);
  }
  setLayout(mode) { this.layout = mode; this._applyLayout(); }
  _applyLayout() {
    const side = this.layout === "side" || (this.layout === "auto" && this.root.clientWidth >= 760);
    this.root.classList.toggle("side", side); this.root.classList.toggle("stack", !side);
  }
  setDrawMode(type) { this.left.setDrawMode(type); this.right.setDrawMode(type); }

  // Mode simple : une seule liseuse (pièces justificatives, snips manuels)
  single() {
    if (!this.double) return;
    this.double = false; this.model = null;
    this.root.querySelector(".pane.R").hidden = true;
    this.root.classList.remove("double");
    this.root.querySelector(".pane.L .ph").hidden = true;
    this.setActive(this.left);
  }
  _setDouble(on) {
    this.double = on;
    this.root.classList.toggle("double", on);
    this.root.querySelector(".pane.R").hidden = !on;
    this.root.querySelector(".pane.L .ph").hidden = false;
  }

  async _open(v, file, page, rect, hls) {
    const blob = await this.o.getFile(file);
    if (!blob) { v.empty(`Pièce « ${esc(file)} » absente du classeur.`); return false; }
    await v.open(file, blob);
    v.setHighlights(hls);
    v.goto(page, rect);
    return true;
  }
  _extra(file) { return (this.model?.manual || []).filter(z => z.file === file).map(z => ({ page: z.page, rect: z.rect, cls: z.type, label: ({ texte: "T", somme: "Σ", valide: "✓", exception: "✗" }[z.type] || "◆") + " " + (z.cell || "") })); }

  // model = buildModel(...) ; affiche le montant à gauche et la première source (ou la source déjà choisie) à droite
  async show(model) {
    this.model = model;
    const hasR = model.groups.length > 0;
    this._setDouble(true);
    this.root.querySelector(".pane.R").hidden = !hasR;
    this.root.classList.toggle("double", hasR);
    const L = model.L.length ? model.L : model.manual;
    const lab = this.root.querySelector(".pane.L .pl");
    if (L.length) {
      const z = L[0];
      lab.textContent = z.source ? `· ${z.source}` : `· ${z.file} p.${z.page}`;
      const hls = model.L.filter(x => x.file === z.file).map(x => ({ page: x.page, rect: x.rect, cls: "auto", focus: true })).concat(this._extra(z.file));
      await this._open(this.left, z.file, z.page, z.rect, hls);
    } else { lab.textContent = ""; this.left.empty("Aucun montant de la plaquette n'est lié à cette cellule."); }
    if (hasR) {
      const g = model.groups.find(x => x.key === this.activeKey) || model.groups[0];
      await this.openGroup(g.id);
    }
  }
  _tabs() {
    const el = this.root.querySelector(".rtabs"); const m = this.model;
    el.innerHTML = `<span class="pt">Source</span>` + m.groups.map(g => {
      const pages = [...new Set(g.zones.map(z => z.page))];
      return `<button class="rtab ${g.id === this.activeId ? "active" : ""}" data-id="${esc(g.id)}" title="${esc(g.file)}">${esc(g.label)}${g.zones.length > 1 ? ` <span class="n">×${g.zones.length}</span>` : ""}${pages.length === 1 && !/Balance/.test(g.key) ? ` <span class="n">p.${pages[0]}</span>` : ""}</button>`;
    }).join("");
    el.querySelectorAll(".rtab").forEach(b => b.onclick = () => this.openGroup(b.dataset.id));
  }
  async openGroup(id) {
    const g = this.model.groups.find(x => x.id === id); if (!g) return;
    this.activeId = g.id; this.activeKey = g.key; this._tabs();
    const first = g.zones.slice().sort((a, b) => a.page - b.page || b.rect[3] - a.rect[3])[0];
    const hls = g.zones.map(z => ({ page: z.page, rect: z.rect, cls: "auto", focus: true })).concat(this._extra(g.file));
    await this._open(this.right, g.file, first.page, first.rect, hls);
  }
}
