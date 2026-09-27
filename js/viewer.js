// Visionneuse PDF (PDF.js embarqué) avec surlignages et tracé de zones (snips).
import * as pdfjsLib from "../vendor/pdfjs/pdf.min.mjs";
import { itemBox, norm } from "./extract.js?v=2";

const VENDOR = new URL("../vendor/pdfjs/", import.meta.url).href;
pdfjsLib.GlobalWorkerOptions.workerSrc = VENDOR + "pdf.worker.min.mjs";

export class Viewer {
  constructor(el, cb = {}) {
    this.el = el; this.cb = cb;
    this.doc = null; this.name = null; this.pages = []; this.scale = 1; this.fit = true;
    this.hls = []; this.textCache = new Map(); this.drawType = null; this.loadToken = 0;
    this.io = new IntersectionObserver(es => es.forEach(e => { if (e.isIntersecting) this._render(+e.target.dataset.n); }), { root: el, rootMargin: "400px 0px" });
    el.addEventListener("scroll", () => this._reportPage());
    window.addEventListener("resize", () => { if (this.fit && this.doc) this.fitWidth(); });
    this.empty("Chargez les pièces (📁) puis sélectionnez une cellule de la table.");
  }
  empty(msg) { this.el.innerHTML = `<div class="empty">${msg}</div>`; this.doc = null; this.name = null; this.pages = []; }

  async open(name, blob) {
    if (this.name === name && this.doc) return;
    const token = ++this.loadToken;
    const data = new Uint8Array(await blob.arrayBuffer());
    const doc = await pdfjsLib.getDocument({ data, standardFontDataUrl: VENDOR + "standard_fonts/", isEvalSupported: false }).promise;
    if (token !== this.loadToken) { doc.destroy(); return; }
    if (this.doc) this.doc.destroy();
    this.doc = doc; this.name = name; this.textCache.clear(); this.hls = [];
    this.pages = [];
    for (let n = 1; n <= doc.numPages; n++) {
      const p = await doc.getPage(n);
      this.pages.push({ n, page: p, base: p.getViewport({ scale: 1 }), el: null, rendered: 0 });
    }
    this._layout(true);
  }
  get numPages() { return this.pages.length; }

  _availWidth() { return Math.max(200, this.el.clientWidth - 24); }
  fitWidth() {
    if (!this.pages.length) return;
    this.fit = true; this.scale = this._availWidth() / this.pages[0].base.width; this._layout(false);
  }
  zoom(f) { if (!this.pages.length) return; this.fit = false; this.scale = Math.min(5, Math.max(0.3, this.scale * f)); this._layout(false); }

  _layout(first) {
    const keepPage = first ? 1 : this.currentPage();
    if (first && this.fit) this.scale = this._availWidth() / this.pages[0].base.width;
    this.io.disconnect(); this.el.innerHTML = "";
    this.el.classList.toggle("drawing", !!this.drawType);
    for (const pg of this.pages) {
      const vp = pg.page.getViewport({ scale: this.scale });
      const d = document.createElement("div"); d.className = "page"; d.dataset.n = pg.n;
      d.style.width = vp.width + "px"; d.style.height = vp.height + "px";
      d.innerHTML = `<span class="lbl">p.${pg.n}</span><div class="overlay"></div>`;
      pg.el = d; pg.vp = vp; pg.rendered = 0;
      this._bindDraw(pg);
      this.el.appendChild(d); this.io.observe(d);
    }
    this._drawHls();
    this.goto(keepPage, null, false);
    this.cb.onZoom && this.cb.onZoom(this.scale);
  }

  async _render(n) {
    const pg = this.pages[n - 1]; if (!pg || pg.rendered === this.scale) return;
    pg.rendered = this.scale;
    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    const vp = pg.page.getViewport({ scale: this.scale * dpr });
    const c = document.createElement("canvas"); c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
    try {
      await pg.page.render({ canvasContext: c.getContext("2d"), viewport: vp }).promise;
      const old = pg.el.querySelector("canvas"); if (old) old.remove();
      pg.el.insertBefore(c, pg.el.firstChild);
    } catch (e) { pg.rendered = 0; }
  }

  currentPage() {
    if (!this.pages.length) return 1;
    const top = this.el.scrollTop + 20;
    let cur = 1; for (const pg of this.pages) { if (pg.el && pg.el.offsetTop <= top) cur = pg.n; }
    return cur;
  }
  _reportPage() { this.cb.onPage && this.cb.onPage(this.currentPage(), this.pages.length); }

  goto(n, rect, smooth = true) {
    const pg = this.pages[Math.min(Math.max(1, n || 1), this.pages.length) - 1]; if (!pg) return;
    let y = pg.el.offsetTop - 8;
    if (rect) { const vr = pg.vp.convertToViewportRectangle(norm(rect)); y = pg.el.offsetTop + Math.min(vr[1], vr[3]) - this.el.clientHeight / 3; }
    this.el.scrollTo({ top: Math.max(0, y), behavior: smooth ? "smooth" : "auto" });
    setTimeout(() => this._reportPage(), 50);
  }

  // Surlignages : [{page, rect:[x1,y1,x2,y2] (PDF), cls, label, focus}]
  setHighlights(list) { this.hls = list || []; this._drawHls(); }
  _drawHls() {
    for (const pg of this.pages) { if (pg.el) pg.el.querySelectorAll(".hl:not(.draft)").forEach(x => x.remove()); }
    for (const h of this.hls) {
      const pg = this.pages[h.page - 1]; if (!pg || !pg.el) continue;
      const vr = pg.vp.convertToViewportRectangle(norm(h.rect));
      const d = document.createElement("div"); d.className = `hl ${h.cls}` + (h.focus ? " focus" : "");
      Object.assign(d.style, { left: Math.min(vr[0], vr[2]) - 2 + "px", top: Math.min(vr[1], vr[3]) - 2 + "px", width: Math.abs(vr[2] - vr[0]) + 4 + "px", height: Math.abs(vr[3] - vr[1]) + 4 + "px" });
      if (h.label) d.innerHTML = `<span class="tag">${h.label}</span>`;
      pg.el.querySelector(".overlay").appendChild(d);
    }
  }

  setDrawMode(type) { this.drawType = type; this.el.classList.toggle("drawing", !!type); }
  _bindDraw(pg) {
    const ov = pg.el.querySelector(".overlay");
    let start = null, draft = null;
    const pos = e => { const r = ov.getBoundingClientRect(); return [e.clientX - r.left, e.clientY - r.top]; };
    ov.addEventListener("pointerdown", e => {
      if (!this.drawType || e.button !== 0) return;
      e.preventDefault(); ov.setPointerCapture(e.pointerId);
      start = pos(e); draft = document.createElement("div"); draft.className = "hl draft"; ov.appendChild(draft);
    });
    ov.addEventListener("pointermove", e => {
      if (!start) return; const p = pos(e);
      Object.assign(draft.style, { left: Math.min(start[0], p[0]) + "px", top: Math.min(start[1], p[1]) + "px", width: Math.abs(p[0] - start[0]) + "px", height: Math.abs(p[1] - start[1]) + "px" });
    });
    ov.addEventListener("pointerup", e => {
      if (!start) return; const p = pos(e); const s = start; start = null; draft.remove();
      if (Math.abs(p[0] - s[0]) < 4 || Math.abs(p[1] - s[1]) < 4) return;
      const a = pg.vp.convertToPdfPoint(s[0], s[1]), b = pg.vp.convertToPdfPoint(p[0], p[1]);
      this.cb.onRect && this.cb.onRect(pg.n, norm([a[0], a[1], b[0], b[1]]));
    });
  }

  async textItems(n) {
    if (!this.textCache.has(n)) { const tc = await this.pages[n - 1].page.getTextContent(); this.textCache.set(n, tc.items); }
    return this.textCache.get(n);
  }
  // Image d'une zone (pour l'OCR), rendue à haute résolution
  async regionCanvas(n, rect, scale = 3) {
    const pg = this.pages[n - 1]; const r = norm(rect);
    const vp = pg.page.getViewport({ scale });
    const full = vp.convertToViewportRectangle(r);
    const x = Math.floor(Math.min(full[0], full[2])), y = Math.floor(Math.min(full[1], full[3]));
    const w = Math.ceil(Math.abs(full[2] - full[0])), h = Math.ceil(Math.abs(full[3] - full[1]));
    const c = document.createElement("canvas"); c.width = w + 16; c.height = h + 16;
    const ctx = c.getContext("2d"); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height);
    ctx.translate(-x + 8, -y + 8);
    await pg.page.render({ canvasContext: ctx, viewport: vp }).promise;
    return c;
  }
  // Recherche d'une valeur (variantes sans espaces) dans le texte de la pièce
  async search(variants, maxHits = 12) {
    const hits = [];
    if (!variants.length || !this.pages.length) return hits;
    for (const pg of this.pages) {
      const items = await this.textItems(pg.n);
      for (const it of items) {
        const s = (it.str || "").replace(/[\s  ]/g, "").toLowerCase();
        if (!s) continue;
        if (variants.some(v => s === v || (v.length >= 4 && s.includes(v)))) {
          hits.push({ page: pg.n, rect: itemBox(it) });
          if (hits.length >= maxHits) return hits;
        }
      }
    }
    return hits;
  }
}
