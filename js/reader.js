/* global Office */
// Fenêtre détachée : reçoit d'Excel (volet) la cellule sélectionnée et ses zones, demande les PDF manquants,
// et renvoie au volet les zones encadrées (snips), que le volet enregistre dans le classeur.
import { Split } from "./split.js?v=10";
import { textInRect } from "./extract.js?v=10";

const $ = s => document.querySelector(s);
const files = new Map(), pending = new Map();
let send = () => {}, tool = null;
const msg = (t, k = "") => { const m = $("#msg"); m.textContent = t; m.className = "msg " + k; };

function getFile(name) {
  if (files.has(name)) return Promise.resolve(files.get(name));
  if (!pending.has(name)) {
    let res; const p = new Promise(r => { res = r; });
    pending.set(name, { parts: [], got: 0, p, res });
    send({ t: "need", name });
    msg(`Chargement de « ${name} »…`);
  }
  return pending.get(name).p;
}
async function b64ToBlob(b64) {
  try { return await (await fetch("data:application/pdf;base64," + b64)).blob(); }
  catch (e) { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return new Blob([u], { type: "application/pdf" }); }
}

const split = new Split($("#reader"), {
  layout: "auto", getFile,
  onRect: async (v, page, rect) => {
    if (!tool) return;
    const text = textInRect(await v.textItems(page), rect);
    send({ t: "snip", tool, file: v.name, page, rect: rect.map(x => Math.round(x * 100) / 100), text });
    msg("Snip transmis à Excel…");
  },
  onPage: (v, n, t) => { if (v === split.active) $("#pageInfo").textContent = `p. ${n} / ${t}`; },
  onZoom: (v, s) => { if (v === split.active) $("#zoomInfo").textContent = Math.round(s * 100) + " %"; },
  onActive: v => { $("#pageInfo").textContent = v.numPages ? `p. ${v.currentPage()} / ${v.numPages}` : "–"; $("#zoomInfo").textContent = Math.round(v.scale * 100) + " %"; }
});
split.left.empty("Sélectionnez une cellule de la feuille de pointage dans Excel.");

async function onMessage(m) {
  if (!m || !m.t) return;
  if (m.t === "chunk") {
    const q = pending.get(m.name); if (!q) return;
    q.parts[m.i] = m.data; q.got++;
    if (q.got % 10 === 0) msg(`Chargement de « ${m.name} » : ${Math.round(100 * q.got / m.n)} %`);
    if (q.got === m.n) { const blob = await b64ToBlob(q.parts.join("")); files.set(m.name, blob); pending.delete(m.name); q.res(blob); msg(""); }
  } else if (m.t === "missing") { const q = pending.get(m.name); if (q) { pending.delete(m.name); q.res(null); } }
  else if (m.t === "show") {
    $("#cell").innerHTML = m.head || "";
    if (m.model) { await split.show(m.model); msg(m.note || "", m.kind || ""); }
    else msg(m.note || "Pas de zone liée à cette cellule.", m.kind || "");
  } else if (m.t === "msg") msg(m.text, m.kind || "");
  else if (m.t === "tool") setTool(m.tool, true);
}

function setTool(t, quiet) {
  tool = tool === t && !quiet ? null : t;
  document.querySelectorAll(".tool").forEach(b => b.classList.toggle("active", b.dataset.type === tool));
  split.setDrawMode(tool);
  if (!quiet) { send({ t: "tool", tool }); msg(tool ? "Encadrez la zone sur l'une des deux liseuses : elle sera liée à la cellule sélectionnée dans Excel." : ""); }
}
document.querySelectorAll(".tool").forEach(b => b.onclick = () => setTool(b.dataset.type));
$("#prevPage").onclick = () => split.active.goto(split.active.currentPage() - 1);
$("#nextPage").onclick = () => split.active.goto(split.active.currentPage() + 1);
$("#zoomIn").onclick = () => split.active.zoom(1.2);
$("#zoomOut").onclick = () => split.active.zoom(1 / 1.2);
$("#zoomFit").onclick = () => split.active.fitWidth();
$("#layout").onclick = () => split.setLayout($("#reader").classList.contains("side") ? "stack" : "side");
$("#attach").onclick = () => send({ t: "close" });

// Canal de communication
if (window.opener && window.name === "liseuse") {
  send = m => window.opener.postMessage(m, "*");
  window.addEventListener("message", e => { if (e.source === window.opener) onMessage(e.data); });
  send({ t: "ready" });
} else if (window.Office) {
  Office.onReady(() => {
    send = m => Office.context.ui.messageParent(JSON.stringify(m));
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, a => { try { onMessage(JSON.parse(a.message)); } catch (e) { msg("Message illisible : " + e.message, "err"); } });
    send({ t: "ready" });
  });
} else msg("Ouvrez cette fenêtre depuis le volet « Pièces & snips » d'Excel.", "err");
window.__reader_split = split;
