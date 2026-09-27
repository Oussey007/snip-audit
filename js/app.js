// Volet « Pièces justificatives » : visionneuse liée aux cellules + snips façon DataSnipper.
import { Viewer } from "./viewer.js?v=2";
import * as store from "./store.js?v=2";
import { textInRect, numbersIn, valueFromText, round2, searchVariants } from "./extract.js?v=2";

const $ = s => document.querySelector(s);
const XL = window.__MOCK__ ? await import("./excel-mock.js?v=2") : await import("./excel-bridge.js?v=2");

const S = { files: [], liens: new Map(), snips: [], idx: new Map(), zones: [], sel: null, ref: null, tool: null, tabs: [], busy: false };
const LABEL = { texte: "T", somme: "Σ", valide: "✓", exception: "✗", auto: "◆" };
// Index des snips par cellule (manuels d'abord, puis zones automatiques)
function indexSnips() {
  S.idx = new Map();
  for (const s of S.snips) { const k = s.sheet + "|" + s.cell; if (!S.idx.has(k)) S.idx.set(k, []); S.idx.get(k).push(s); }
  for (const l of S.idx.values()) l.sort((a, b) => (a.type === "auto") - (b.type === "auto"));
}
async function loadSnips() { S.snips = await XL.readSnips(); indexSnips(); }
const manual = () => S.snips.filter(s => s.type !== "auto");
const TYPE_FR = { texte: "texte", somme: "somme", valide: "validation", exception: "exception" };

function msg(t, kind = "") { const m = $("#msg"); m.textContent = t; m.className = "msg " + kind; }
const fmt = v => typeof v === "number" ? v.toLocaleString("fr-FR", { minimumFractionDigits: Number.isInteger(v) ? 0 : 2, maximumFractionDigits: 2 }) : String(v ?? "");
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const base = p => String(p).split(/[\\/]/).pop();

const viewer = new Viewer($("#viewer"), {
  onRect: (page, rect) => createSnip(page, rect).catch(e => msg("Snip impossible : " + e.message, "err")),
  onPage: (n, t) => { $("#pageInfo").textContent = `p. ${n} / ${t}`; },
  onZoom: s => { $("#zoomInfo").textContent = Math.round(s * 100) + " %"; }
});

// ---------- Pièces (stockées localement)
async function refreshFiles() {
  S.files = (await store.listFiles()).sort((a, b) => a.path.localeCompare(b.path, "fr"));
  $("#docCount").textContent = `${S.files.length} pièce${S.files.length > 1 ? "s" : ""}`;
  const sel = $("#docSelect"); const cur = viewer.name;
  sel.innerHTML = `<option value="">— choisir une pièce —</option>` + S.files.map(f => `<option value="${esc(f.name)}" ${f.name === cur ? "selected" : ""}>${esc(f.path)}</option>`).join("");
}
async function addFiles(list) {
  const pdfs = [...list].filter(f => /\.pdf$/i.test(f.name));
  if (!pdfs.length) return msg("Aucun PDF trouvé dans la sélection.", "err");
  msg(`Chargement de ${pdfs.length} pièce(s)…`);
  await store.putFiles(pdfs);
  await refreshFiles();
  msg(`${pdfs.length} pièce(s) chargée(s). Elles restent sur ce poste.`, "ok");
  onSelection();
}
$("#pickDir").addEventListener("change", e => addFiles(e.target.files));
$("#pickFiles").addEventListener("change", e => addFiles(e.target.files));
$("#docSelect").addEventListener("change", e => { if (e.target.value) openDoc(e.target.value, 1); });

async function openDoc(name, page = 1, rect = null) {
  const blob = await store.getFile(name);
  if (!blob) { msg(`Pièce « ${name} » non chargée : cliquez sur 📁 Pièces et sélectionnez le dossier « pièces triées ».`, "err"); return false; }
  await viewer.open(name, blob);
  $("#docSelect").value = name;
  paintHighlights(rect ? { page, rect } : null);
  viewer.goto(page, rect);
  renderTabs();
  return true;
}

// ---------- Surlignages : snips de la pièce ouverte + résultats de recherche
let searchHits = [];
function paintHighlights(focus) {
  const list = manual().filter(s => s.file === viewer.name && !S.zones.includes(s)).map(s => ({
    page: s.page, rect: s.rect, cls: s.type, label: `${LABEL[s.type]} ${s.cell}`,
    focus: !!focus && focus.page === s.page && focus.rect.join() === s.rect.join()
  }));
  // zones de la cellule sélectionnée (encadrées et mises en évidence)
  S.zones.filter(z => z.file === viewer.name).forEach(z => list.push({ page: z.page, rect: z.rect, cls: z.type, focus: true,
    label: z.type === "auto" ? (z.text || "Donnée") : `${LABEL[z.type]} ${z.cell}` }));
  searchHits.forEach(h => list.push({ page: h.page, rect: h.rect, cls: "search" }));
  viewer.setHighlights(list);
}

// ---------- Onglets des pièces de l'écriture sélectionnée (onglet « Liens »)
function shortRole(r) {
  const t = r.toLowerCase();
  if (t.startsWith("facture")) return "Facture";
  if (t.includes("bon de commande")) return "BC";
  if (t.includes("devis") || t.includes("proposition")) return "Devis";
  if (t.includes("contrat")) return "Contrat";
  if (t.includes("bon pour accord")) return "BPA";
  if (t.includes("livraison") || t.includes("récupération") || t.includes("edl")) return "EDL";
  if (t.includes("mail")) return t.includes("arrêt") ? "Arrêt" : "Mail";
  if (t.includes("géoloc")) return "Carte";
  if (t.includes("extrait")) return "Cpte client";
  if (t.includes("remise")) return t.includes("total") ? "Remise (total)" : "Remise";
  if (t.includes("relevé")) return "Banque";
  if (t.includes("od")) return "OD";
  return r.slice(0, 14);
}
function roleOf(file) { const t = S.tabs.find(x => x.file === file); return t ? shortRole(t.role) : file.replace(/\.pdf$/i, "").slice(0, 18); }
function renderTabs() {
  const el = $("#tabs");
  const zf = [...new Set(S.zones.map(z => z.file))];
  const zoneBar = zf.length ? `<div class="zbar">◆ Donnée dans : ` + zf.map((f, i) => { const n = S.zones.filter(z => z.file === f).length;
      return `<button class="zchip ${f === viewer.name ? "active" : ""}" data-z="${i}" title="${esc(f)}">${esc(roleOf(f))}${n > 1 ? " ×" + n : ""}</button>`; }).join("") + `</div>` : "";
  if (!S.tabs.length && !zoneBar) { el.innerHTML = ""; return; }
  el.innerHTML = zoneBar + S.tabs.map((t, i) => `<button class="tab ${t.file === viewer.name ? "active" : ""}" data-i="${i}" title="${esc(t.role)} – ${esc(t.file)}">${esc(shortRole(t.role))} <span class="p">p.${t.page}</span></button>`).join("");
  el.querySelectorAll(".tab").forEach(b => b.onclick = () => { const t = S.tabs[+b.dataset.i]; openDoc(t.file, t.page); });
  el.querySelectorAll(".zchip").forEach(b => b.onclick = () => { const f = zf[+b.dataset.z]; const z = S.zones.find(x => x.file === f); openDoc(f, z.page, z.rect); });
}

// ---------- Sélection d'une cellule dans Excel
async function onSelection() {
  let sel;
  try { sel = await XL.getSelection(); } catch (e) { return; }
  S.sel = sel;
  const zones = S.idx.get(sel.sheet + "|" + sel.cell) || [];
  const snip = zones.find(z => z.type !== "auto");
  const hdr = String(sel.header || "").replace(/\n/g, " ");
  const nfiles = new Set(zones.map(z => z.file)).size;
  $("#cellInfo").innerHTML = `<b>${esc(sel.sheet)}!${esc(sel.cell)}</b> ${hdr ? "· " + esc(hdr.slice(0, 60)) : ""}<br>${esc(fmt(sel.value)).slice(0, 120) || "<span class='muted'>(vide)</span>"}` +
    (snip ? `<br><span style="color:var(--${snip.type})">● Snip ${TYPE_FR[snip.type]} – ${esc(snip.file)} p.${snip.page}</span>` : "") +
    (zones.length && !snip ? `<br><span style="color:var(--auto)">◆ ${zones.length} zone${zones.length > 1 ? "s" : ""} dans ${nfiles} pièce${nfiles > 1 ? "s" : ""}</span>` : "");

  // Écriture de la ligne → pièces liées
  const ref = (sel.rowValues || []).map(v => String(v ?? "").trim()).find(v => S.liens.has(v)) || null;
  S.ref = ref; S.tabs = ref ? S.liens.get(ref) : [];
  S.zones = zones; searchHits = [];

  // 1) la cellule a des zones (snip manuel ou zone automatique) → pièce + zone(s) encadrée(s)
  if (zones.length) {
    const z = zones[0];
    if (await openDoc(z.file, z.page, z.rect)) msg(nfiles > 1 ? `Donnée localisée dans ${nfiles} pièces : cliquez sur les pastilles ◆ pour passer de l'une à l'autre.` : "");
    return;
  }
  // 2) la cellule est un lien « ▶ … p.X » → cette pièce, cette page
  const m = /HYPERLINK\(\s*"([^"]+)"\s*[,;]\s*"[^"]*?p\.(\d+)/i.exec(String(sel.formula || ""));
  if (m) { await openDoc(base(m[1]), +m[2]); return; }
  // 3) autre cellule d'une écriture → pièce principale (ou celle déjà ouverte si elle concerne l'écriture)
  if (ref) {
    const already = S.tabs.find(t => t.file === viewer.name);
    const t = already || S.tabs[0];
    if (!already) { if (!(await openDoc(t.file, t.page))) return; } else { paintHighlights(null); renderTabs(); }
  } else { paintHighlights(null); renderTabs(); }
  // 4) recherche de la valeur de la cellule dans la pièce ouverte
  if (viewer.name) {
    searchHits = await viewer.search(searchVariants(sel.value));
    paintHighlights(null);
    if (searchHits.length) { viewer.goto(searchHits[0].page, searchHits[0].rect); msg(`Valeur trouvée ${searchHits.length} fois dans la pièce (surlignée en jaune).`); }
    else if (ref) msg("Pas de zone liée à cette cellule (donnée de calcul ou pièce non fournie).");
  }
}

// ---------- Création d'un snip
function setTool(type) {
  S.tool = S.tool === type ? null : type;
  document.querySelectorAll(".tool").forEach(b => b.classList.toggle("active", b.dataset.type === S.tool));
  viewer.setDrawMode(S.tool);
  if (S.tool) msg(viewer.name ? `Snip ${TYPE_FR[S.tool]} : encadrez la zone sur la pièce. Elle sera liée à la cellule sélectionnée dans Excel.` : "Ouvrez d'abord une pièce.", viewer.name ? "" : "err");
  else msg("");
}
document.querySelectorAll(".tool").forEach(b => b.addEventListener("click", () => setTool(b.dataset.type)));

async function createSnip(page, rect) {
  if (S.busy) return; S.busy = true;
  try {
    const sel = await XL.getSelection();
    if (sel.sheet === XL.SNIP_SHEET) throw new Error("sélectionnez une cellule de la table, pas la feuille _Snips");
    const type = S.tool;
    let text = textInRect(await viewer.textItems(page), rect);
    let ocr = false;
    if (!text && (type === "texte" || type === "somme")) {
      msg("Zone scannée : lecture OCR en cours…");
      const { ocrCanvas } = await import("./ocr.js?v=2");
      text = await ocrCanvas(await viewer.regionCanvas(page, rect)); ocr = true;
    }
    let value;
    if (type === "texte") { if (!text) throw new Error("aucun texte lisible dans la zone"); value = valueFromText(text); }
    else if (type === "somme") { const n = numbersIn(text); if (!n.length) throw new Error("aucun montant dans la zone"); value = round2(n.reduce((a, b) => a + b, 0)); }
    else value = type === "valide" ? "✓" : "✗";
    const rec = { id: (crypto.randomUUID ? crypto.randomUUID() : String(Date.now())), sheet: sel.sheet, cell: sel.cell, type, file: viewer.name, page, rect: rect.map(x => Math.round(x * 100) / 100), value, text: text + (ocr ? " [OCR]" : ""), date: new Date().toLocaleString("fr-FR") };
    const res = await XL.saveSnip(rec, value);
    await loadSnips();
    paintHighlights({ page, rect: rec.rect });
    let t = `Snip ${TYPE_FR[type]} → ${sel.sheet}!${sel.cell}`;
    if (type === "texte" || type === "somme") t += ` = ${fmt(value)}`;
    if (!res.written && res.hasFormula) t += " (la cellule contient une formule : valeur non écrite, zone liée)";
    if (ocr) t += " · OCR : vérifiez la valeur";
    msg(t, "ok");
    $("#cellInfo").lastChild && onSelectionInfoOnly();
  } finally { S.busy = false; }
}
async function onSelectionInfoOnly() { const keep = searchHits; await onSelection(); searchHits = keep; }

$("#btnDelSnip").addEventListener("click", async () => {
  const sel = await XL.getSelection();
  const ok = await XL.deleteSnip(sel.sheet, sel.cell);
  await loadSnips(); paintHighlights(null);
  msg(ok ? `Snip supprimé de ${sel.sheet}!${sel.cell}.` : "Aucun snip sur la cellule sélectionnée.", ok ? "ok" : "err");
});

// ---------- Liste des snips
function renderSnipList() {
  const el = $("#snipList");
  const L = manual(); const nauto = S.snips.length - L.length;
  const head = nauto ? `<p class="muted">${nauto} zones automatiques liées aux cellules de la table (non listées).</p>` : "";
  if (!L.length) { el.innerHTML = head + `<p class="muted">Aucun snip manuel. Sélectionnez une cellule, choisissez un outil (T, Σ, ✓, ✗) puis encadrez la zone sur la pièce.</p>`; return; }
  el.innerHTML = head + L.map((s, i) => `<div class="snip ${s.type}" data-i="${i}"><b>${LABEL[s.type]} ${esc(s.sheet)}!${esc(s.cell)}</b> ${s.type === "texte" || s.type === "somme" ? "= " + esc(fmt(s.value)) : ""}<div class="f">${esc(s.file)} · p.${s.page} · ${esc(s.date)}</div></div>`).join("");
  el.querySelectorAll(".snip").forEach(d => d.onclick = async () => { const s = L[+d.dataset.i]; await XL.selectCell(s.sheet, s.cell); await openDoc(s.file, s.page, s.rect); });
}
$("#btnSnips").addEventListener("click", async () => { await loadSnips(); renderSnipList(); $("#snipPanel").hidden = false; });
$("#closeSnips").addEventListener("click", () => { $("#snipPanel").hidden = true; });

// ---------- Navigation
$("#prevPage").onclick = () => viewer.goto(viewer.currentPage() - 1);
$("#nextPage").onclick = () => viewer.goto(viewer.currentPage() + 1);
$("#zoomIn").onclick = () => viewer.zoom(1.2);
$("#zoomOut").onclick = () => viewer.zoom(1 / 1.2);
$("#zoomFit").onclick = () => viewer.fitWidth();

// ---------- Démarrage
(async () => {
  try {
    const info = await XL.ready();
    if (!info || info.host !== (window.Office && Office.HostType ? Office.HostType.Excel : "Excel")) { msg("Ce volet doit être ouvert depuis Excel.", "err"); }
    await refreshFiles();
    try { S.liens = await XL.readLiens(); } catch (e) { S.liens = new Map(); }
    try { await loadSnips(); } catch (e) { S.snips = []; }
    XL.onSelection(() => onSelection().catch(e => msg(e.message, "err")));
    msg(S.files.length ? `${S.liens.size} écritures liées · ${manual().length} snip(s) · ${S.snips.length - manual().length} zones automatiques.` : "Cliquez sur 📁 Pièces et sélectionnez le dossier « pièces triées ».");
    await onSelection();
    window.__app_ready = true;
  } catch (e) { msg("Erreur au démarrage : " + e.message, "err"); window.__app_error = e.message; }
})();
window.__app = { S, viewer, createSnip, onSelection, setTool, openDoc };
