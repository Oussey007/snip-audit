// Volet « Pointage plaquette » : FEC -> balance -> rapprochement avec le bilan et le compte de résultat de la plaquette.
import * as pdfjsLib from "../vendor/pdfjs/pdf.min.mjs";
import { parseFEC } from "./fs/fec.js?v=7";
import { DEFAULT_POSTES, comptesDe, SECTION_LABEL } from "./fs/pcg.js?v=7";
import { extractStatements, matchPostes } from "./fs/plaquette.js?v=7";
import { tieOut } from "./fs/tieout.js?v=7";
import { readAnnex } from "./fs/annexe.js?v=7";
import { checkAnnex } from "./fs/annexcheck.js?v=7";
import { annexPages } from "./fs/annexe.js?v=7";
import { snapshot, hashItems, compare, parseVersions, versionRows, VERS_HEAD, keyP, keyA } from "./fs/versions.js?v=7";
import { SEUILS, CHECKLIST, STATUTS, categorie, annexLines, verifyCitation, instructionExcel, dossierConversation, parseReponse } from "./fs/annexai.js?v=7";

export const SHEETS = { versions: "_Versions", modifications: "Modifications plaquette", pointage: "Pointage plaquette", annexe: "Pointage annexe", texte: "Annexe texte", conformite: "Conformité annexe", incoherences: "Incohérences annexe", controles: "Contrôles plaquette", balance: "Balance FEC", mapping: "Mapping PCG" };
const VENDOR = new URL("../vendor/pdfjs/", import.meta.url).href;
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const eur = v => (v ?? 0).toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const STATUS_COLOR = { "Concordant": "#C6EFCE", "Concordant (arrondi)": "#E2F0D9", "Écart": "#FFC7CE", "Écart – reclassement probable": "#FCE4D6", "Absent de la plaquette": "#FFEB9C", "Écart – à analyser": "#FFEB9C", "Absent de l'annexe": "#FFEB9C", "Référence introuvable": "#FFEB9C", "Citation vérifiée": "#C6EFCE", "Citation vérifiée (extraits)": "#E2F0D9", "Citation introuvable": "#FFC7CE", "Trop courte pour être vérifiée": "#FFEB9C", "Statut non reconnu": "#FFC7CE", "Inchangé": "#C6EFCE", "Modifié": "#FCE4D6", "Nouveau": "#DDEBF7", "Ajouté": "#DDEBF7", "Supprimé": "#FFC7CE", "Non renseigné dans la plaquette": "#FFEB9C", "OK": "#C6EFCE", "Anomalie": "#FFC7CE", "Info": "#DDEBF7" };

let A = null;                   // dépendances fournies par app.js
const F = { fec: null, fecName: "", last: null };

function status(t, kind = "") { const el = $("#fsStatus"); el.textContent = t; el.className = "msg " + kind; A.msg(t, kind); }

// ---------- Mapping : feuille du classeur si elle existe (corrections de l'utilisateur), sinon plan par défaut
const MAP_HEAD = ["ID", "Section", "Poste", "Libellés reconnus dans la plaquette (séparés par « | »)", "Comptes", "Amort. / dépréc. (actif)", "Sens (CR : -1 = produit)", "Formule (totaux)", "Si absent de la plaquette, reporter sur"];
async function loadMapping() {
  const v = await A.XL.readSheet(SHEETS.mapping);
  if (v && v.length) {
    const h = v.findIndex(r => String(r[0]).trim() === "ID");
    if (h >= 0) {
      const postes = v.slice(h + 1).filter(r => String(r[0]).trim() && ["actif", "passif", "cr"].includes(String(r[1]).trim())).map(r => {
        const libs = String(r[3] || "").split("|").map(s => s.trim()).filter(Boolean);
        const p = { id: String(r[0]).trim(), section: String(r[1]).trim(), libelle: String(r[2]).trim(), alias: libs, regle: String(r[4] || "").trim(), amort: String(r[5] || "").trim() || undefined, sens: Number(r[6]) || undefined, repli: String(r[8] || "").trim() || undefined };
        if (String(r[7] || "").trim()) { p.total = true; p.formule = String(r[7]).trim(); }
        return p;
      });
      if (postes.length) return { postes, fromSheet: true };
    }
  }
  return { postes: DEFAULT_POSTES, fromSheet: false };
}
async function writeMapping(postes) {
  await A.XL.writeTable(SHEETS.mapping, {
    title: "Plan de regroupement : comptes → postes des comptes annuels",
    subtitle: ["Modifiable : l'outil relit cette feuille à chaque pointage. Section : actif, passif ou cr.",
      "Comptes : préfixes séparés par des espaces. 40C = comptes 40… à solde créditeur, 411D = à solde débiteur, -409 = exclure, RES = résultat (classes 6 et 7).",
      "Libellés : ajoutez le libellé exact de votre plaquette s'il n'est pas reconnu (séparateur « | »)."],
    header: MAP_HEAD,
    rows: postes.map(p => [p.id, p.section, p.libelle, [p.libelle, ...(p.alias || [])].filter((x, i, a) => a.indexOf(x) === i).join(" | "), p.regle || "", p.amort || "", p.sens || "", p.formule || "", p.repli || ""]),
    widths: [70, 50, 260, 360, 260, 140, 70, 260, 120], wrapCols: [3],
    fills: postes.map((p, i) => p.total ? { row: i, color: "#F2F2F2" } : null).filter(Boolean)
  });
}

// ---------- Balance : FEC chargé, sinon feuille « Balance FEC » déjà présente dans le classeur
async function balanceFromSheet() {
  const v = await A.XL.readSheet(SHEETS.balance); if (!v) return null;
  const h = v.findIndex(r => String(r[0]).trim() === "Compte"); if (h < 0) return null;
  const bal = v.slice(h + 1).filter(r => String(r[0]).trim()).map(r => ({ compte: String(r[0]).trim(), lib: String(r[1] || ""), an: +r[2] || 0, debit: +r[3] || 0, credit: +r[4] || 0, solde: Math.round((+r[5] || 0) * 100) / 100 }));
  const resultat = Math.round(-bal.filter(b => /^[67]/.test(b.compte)).reduce((a, b) => a + b.solde, 0) * 100) / 100;
  return { balance: bal, resultat, controles: [{ controle: "Source de la balance", statut: "Info", detail: `feuille « ${SHEETS.balance} » du classeur (FEC non rechargé)` }] };
}

// ---------- Texte de la plaquette (pdf.js)
async function pdfPages(name) {
  const blob = await A.store.getFile(name); if (!blob) throw new Error(`pièce « ${name} » introuvable dans le classeur`);
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(await blob.arrayBuffer()), standardFontDataUrl: VENDOR + "standard_fonts/", isEvalSupported: false }).promise;
  const pages = [];
  for (let n = 1; n <= doc.numPages; n++) {
    if (n % 5 === 0) status(`Lecture de la plaquette : page ${n} / ${doc.numPages}…`);
    const p = await doc.getPage(n); const vp = p.getViewport({ scale: 1 }); const tc = await p.getTextContent();
    pages.push({ n, width: vp.width, height: vp.height, items: tc.items.filter(i => i.str !== undefined).map(i => ({ str: i.str, x: i.transform[4], y: i.transform[5], w: i.width, h: i.height || Math.hypot(i.transform[2], i.transform[3]) })) });
  }
  doc.destroy();
  return pages;
}

// ---------- Versions : reprise des validations de l'auditeur pour les lignes inchangées
function oldValidations(values, firstHeader, cols) {
  const m = new Map(); const h = values ? values.findIndex(r => String(r[0]).trim() === firstHeader) : -1; if (h < 0) return m;
  for (const r of values.slice(h + 1)) { const k = String(r[cols.key] ?? "").trim(); if (k) m.set(k, { a: r[cols.a], b: r[cols.b], valid: String(r[cols.valid] ?? "").trim() }); }
  return m;
}
const same = (x, y) => (x === "" || x === null || x === undefined ? "" : Math.round(Number(x) * 100) / 100) === (y === "" || y === null || y === undefined ? "" : Math.round(Number(y) * 100) / 100);
function suivi(old, a, b, prevV, newV, fmtOld) {
  if (!old) return { valid: "", suivi: prevV ? `Nouveau (V${newV})` : "", s: prevV ? "Nouveau" : "" };
  if (same(old.a, a) && same(old.b, b)) return { valid: old.valid, suivi: `Inchangé${prevV ? " depuis V" + prevV : ""}${old.valid ? " – validation reprise" : ""}`, s: "Inchangé" };
  return { valid: "", suivi: `Modifié (avant : ${fmtOld(old)}) – à revoir${old.valid ? " ; ancienne validation : " + old.valid : ""}`, s: "Modifié" };
}
const keysA = rows => { const seen = new Map(); return rows.map(r => { let k = keyA(r); const n = (seen.get(k) || 0) + 1; seen.set(k, n); return n > 1 ? k + "|" + n : k; }); };

// ---------- Lancement
async function run() {
  const pdf = $("#fsPdf").value;
  if (!pdf) return status("Choisissez la plaquette (PDF) parmi les pièces du classeur, ou ajoutez-la avec « + PDF ».", "err");
  $("#fsRun").disabled = true;
  try {
    let src = F.fec;
    if (!src) { src = await balanceFromSheet(); if (!src) throw new Error("chargez d'abord le FEC (📄 FEC)"); }
    status("Lecture du plan de regroupement…");
    const map = await loadMapping();
    if (!map.fromSheet) await writeMapping(map.postes);
    status("Lecture de la plaquette…");
    const pages = await pdfPages(pdf);
    const st = extractStatements(pages, map.postes);
    const missing = ["actif", "passif", "cr"].filter(s => !(st.pages[s] || []).length);
    if (missing.length === 3) throw new Error("aucun bilan ni compte de résultat reconnu dans ce PDF (titres de pages « Bilan actif », « Bilan passif », « Compte de résultat » attendus)");
    const matched = matchPostes(map.postes, st);
    const t = tieOut(map.postes, src.balance, matched, src);
    // Annexe : calculs des tableaux, concordance avec les comptes annuels et la balance
    status("Lecture et contrôle de l'annexe…");
    const annex = readAnnex(pages, st.pages);
    const ca = checkAnnex(annex, { postes: map.postes, matched, balance: src.balance, statements: st });
    const lines = annexLines(annexPages(pages, st.pages));
    // Versions : comparaison avec le pointage précédent de ce classeur
    const hist = parseVersions(await A.XL.readSheet(SHEETS.versions));
    const prev = hist[hist.length - 1] || null;
    const snap = snapshot(t.rows, ca.rows, lines.filter(l => !l.tab));
    const hash = hashItems(snap);
    const identical = prev && prev.hash === hash;
    const version = identical ? prev.version : (prev ? prev.version + 1 : 1);
    const cmp = prev && !identical ? compare(prev.items, snap) : null;
    const prevV = prev ? (identical ? (hist[hist.length - 2]?.version || null) : prev.version) : null;
    const oldP = oldValidations(await A.XL.readSheet(SHEETS.pointage), "Section", { a: 3, b: 4, valid: 12, key: 14 });
    const oldA = oldValidations(await A.XL.readSheet(SHEETS.annexe), "Note / tableau", { a: 3, b: 6, valid: 12, key: 14 });
    const f2 = v => v === "" || v === null || v === undefined ? "—" : Number(v).toLocaleString("fr-FR", { maximumFractionDigits: 2 });
    const supP = t.rows.map(r => suivi(oldP.get(keyP(r)), r.plaquette, r.balance, prevV ?? (oldP.size ? "préc." : null), version, o => `plaquette ${f2(o.a)} · balance ${f2(o.b)}`));
    const kA = keysA(ca.rows);
    const supA = ca.rows.map((r, i) => suivi(oldA.get(kA[i]), r.annexe, r.ref, prevV ?? (oldA.size ? "préc." : null), version, o => `annexe ${f2(o.a)} · référence ${f2(o.b)}`));
    const nRepris = [...supP, ...supA].filter(x => x.s === "Inchangé" && x.valid).length, nRevoir = [...supP, ...supA].filter(x => x.s === "Modifié").length;
    status("Écriture des feuilles de résultat…");
    // Balance
    if (F.fec) {
      const posteOf = new Map();
      for (const p of map.postes.filter(p => !p.total)) for (const reg of [p.regle, p.amort]) if (reg) for (const b of comptesDe(reg.replace(/\bRES\b/g, ""), src.balance)) posteOf.set(b.compte, (posteOf.get(b.compte) ? posteOf.get(b.compte) + " / " : "") + `${SECTION_LABEL[p.section]} – ${p.libelle}`);
      await A.XL.writeTable(SHEETS.balance, {
        title: "Balance reconstituée depuis le FEC", subtitle: [`Fichier : ${F.fecName} · ${src.info.lignes.toLocaleString("fr-FR")} lignes · résultat ${eur(src.resultat)} € · soldes : débit positif, crédit négatif`],
        header: ["Compte", "Libellé", "Solde d'ouverture (à-nouveaux)", "Mouvements débit", "Mouvements crédit", "Solde de clôture", "Poste des comptes annuels"],
        rows: src.balance.map(b => [b.compte, b.lib, b.an, b.debit, b.credit, b.solde, posteOf.get(b.compte) || ""]),
        numCols: [2, 3, 4, 5], widths: [70, 240, 110, 110, 110, 110, 320]
      });
    }
    // Pointage
    const rows = t.rows;
    const res = await A.XL.writeTable(SHEETS.pointage, {
      title: "Pointage de la plaquette avec la balance",
      subtitle: [`Plaquette : ${pdf} (version V${version}) · Balance : ${F.fec ? F.fecName : "feuille " + SHEETS.balance} · ${new Date().toLocaleString("fr-FR")}`,
        `${t.stats.concordants} montant(s) concordant(s), ${t.stats.ecarts} écart(s), ${t.stats.autres} autre(s). Cliquez sur un montant de la plaquette (colonne D ou K) pour le voir encadré dans le PDF.`],
      header: ["Section", "Poste", "Colonne", "Plaquette N", "Balance (FEC)", "Écart", "Statut", "Commentaire", "Libellé lu dans la plaquette", "Page PDF", "Plaquette N-1", "Comptes retenus", "Validation auditeur", "Suivi des versions", "Clé"],
      rows: rows.map((r, i) => [r.sectionLib, r.poste, r.colonne, r.plaquette, r.balance, null, r.statut, r.commentaire, r.libellePlaquette, r.page, r.n1, r.comptes, supP[i].valid, supP[i].suivi, keyP(r)]),
      formulas: { 5: (rr, r) => r[3] === null || r[3] === "" ? `=E${rr}` : `=E${rr}-D${rr}` },
      numCols: [3, 4, 5, 10], widths: [90, 250, 95, 95, 105, 85, 150, 300, 220, 55, 95, 420, 160, 260, 40], wrapCols: [7, 13],
      fills: rows.flatMap((r, i) => [r.total ? { row: i, color: "#F2F2F2" } : null, { row: i, col: 6, ncol: 1, color: STATUS_COLOR[r.statut] || "#FFFFFF" }, supP[i].s ? { row: i, col: 13, ncol: 1, color: STATUS_COLOR[supP[i].s] } : null]).filter(Boolean)
    });
    const zones = [];
    rows.forEach((r, i) => {
      const row = res.firstRow + i;
      if (r.box) zones.push({ cell: `D${row}`, file: pdf, page: r.page, rect: r.box, value: r.plaquette, text: `Plaquette p.${r.page} – ${r.poste} (${r.colonne})` });
      if (r.n1box) zones.push({ cell: `K${row}`, file: pdf, page: r.page, rect: r.n1box, value: r.n1, text: `Plaquette p.${r.page} – ${r.poste} (N-1)` });
    });
    await A.XL.replaceAutoZones(SHEETS.pointage, zones);
    const resA = await A.XL.writeTable(SHEETS.annexe, {
      title: "Pointage de l'annexe",
      subtitle: [`Plaquette : ${pdf} (version V${version}) · annexe lue pages ${annex.pages.length ? annex.pages[0] + " à " + annex.pages[annex.pages.length - 1] : "—"} · ${annex.tables.length} tableau(x) reconnu(s)`,
        `${ca.stats.concordants} contrôle(s) concordant(s), ${ca.stats.ecarts} écart(s). Contrôle : Calcul = additions du tableau ; Comptes annuels = bilan ou compte de résultat ; Balance (FEC) = soldes et mouvements des comptes ; Texte = montant cité dans une phrase.`,
        "Cliquez sur un montant (colonne D : annexe, colonne G : comptes annuels) pour le voir encadré dans le PDF."],
      header: ["Note / tableau", "Ligne", "Colonne", "Montant annexe", "Contrôle", "Référence", "Montant de référence", "Écart", "Statut", "Commentaire", "Page annexe", "Page référence", "Validation auditeur", "Suivi des versions", "Clé"],
      rows: ca.rows.map((r, i) => [r.note, r.ligne, r.colonne, r.annexe, r.controle, r.reference, r.ref, null, r.statut, r.commentaire || "", r.page, r.refPage || "", supA[i].valid, supA[i].suivi, kA[i]]),
      formulas: { 7: (rr, r) => (r[3] === null || r[6] === null) ? "" : `=G${rr}-D${rr}` },
      numCols: [3, 6, 7], widths: [150, 250, 105, 100, 95, 260, 110, 85, 150, 380, 60, 60, 160, 260, 40], wrapCols: [1, 5, 9, 13],
      fills: ca.rows.flatMap((r, i) => [{ row: i, col: 8, ncol: 1, color: STATUS_COLOR[r.statut] || "#FFFFFF" }, supA[i].s ? { row: i, col: 13, ncol: 1, color: STATUS_COLOR[supA[i].s] } : null]).filter(Boolean)
    });
    const zonesA = [];
    ca.rows.forEach((r, i) => {
      const row = resA.firstRow + i;
      if (r.box) zonesA.push({ cell: `D${row}`, file: pdf, page: r.page, rect: r.box, value: r.annexe, text: `Annexe p.${r.page} – ${r.note} – ${r.ligne}`.slice(0, 200) });
      if (r.refBox && r.refPage) zonesA.push({ cell: `G${row}`, file: pdf, page: r.refPage, rect: r.refBox, value: r.ref, text: `${r.reference} (p.${r.refPage})`.slice(0, 200) });
    });
    await A.XL.replaceAutoZones(SHEETS.annexe, zonesA);
    t.controles.push({ type: "Annexe", controle: "Lecture de l'annexe", statut: annex.tables.length ? "Info" : "Anomalie", detail: annex.tables.length ? `pages ${annex.pages.join(", ")} · tableaux reconnus : ${annex.tables.map(x => `${x.title || x.type} (p.${x.page})`).join(" ; ")}` : "aucun tableau d'annexe reconnu" });

    // Contrôles
    const ctr = [...(src.controles || []).map(c => ({ type: "FEC", ...c })), ...t.controles];
    const resC = await A.XL.writeTable(SHEETS.controles, {
      title: "Contrôles du FEC, du plan de regroupement et de la plaquette", subtitle: [`Plaquette : ${pdf}`],
      header: ["Type", "Contrôle", "Statut", "Détail", "Page PDF"],
      rows: ctr.map(c => [c.type, c.controle, c.statut, c.detail, c.page || ""]),
      widths: [70, 330, 80, 560, 60], wrapCols: [1, 3],
      fills: ctr.map((c, i) => ({ row: i, col: 2, ncol: 1, color: STATUS_COLOR[c.statut] || "#FFFFFF" }))
    });
    await A.XL.replaceAutoZones(SHEETS.controles, ctr.map((c, i) => c.page && c.box ? { cell: `B${resC.firstRow + i}`, file: pdf, page: c.page, rect: c.box, text: c.controle } : null).filter(Boolean));
    // Modifications par rapport à la version précédente
    const boxOf = new Map(); t.rows.forEach(r => boxOf.set(keyP(r), { page: r.page, box: r.box })); ca.rows.forEach((r, i) => boxOf.set(kA[i], { page: r.page, box: r.box }));
    const textLines = lines.filter(l => !l.tab);
    const ch = cmp ? cmp.changes : [];
    const resM = await A.XL.writeTable(SHEETS.modifications, {
      title: "Modifications par rapport à la version précédente de la plaquette",
      subtitle: [!prev ? `Première version enregistrée (V1) : ${pdf}. La comparaison sera faite au prochain pointage d'une nouvelle version de la plaquette ou du FEC.`
        : identical ? `Version V${version} : aucune différence avec la version pointée le ${prev.date} (${prev.pdf}). Les validations ont été reprises.`
        : `Version V${version} (${pdf}) comparée à V${prev.version} (${prev.pdf}, pointée le ${prev.date}) : ${ch.length} modification(s).`,
        prev && !identical ? `Lignes de pointage : ${nRepris} validation(s) reprise(s) car inchangées, ${nRevoir} ligne(s) à revoir (colonne « Suivi des versions » des feuilles de pointage).` : "",
        "Cliquez sur la colonne D pour voir l'élément dans la nouvelle version du PDF."].filter(Boolean),
      header: ["Type", "Élément", "Version précédente", "Nouvelle version", "Statut", "Page (nouvelle version)"],
      rows: ch.map(c => [c.type, c.element, c.ancien, c.nouveau, c.statut, c.page]),
      widths: [120, 330, 330, 330, 170, 70], wrapCols: [1, 2, 3],
      fills: ch.map((c, i) => ({ row: i, col: 4, ncol: 1, color: STATUS_COLOR[c.statut.split(" ")[0]] || "#FCE4D6" }))
    });
    const zonesM = [];
    ch.forEach((c, i) => {
      let z = c.key ? boxOf.get(c.key) : null;
      if (c.textIdx && c.textIdx.length) { const ls = c.textIdx.map(k => textLines[k]).filter(Boolean); if (ls.length) z = { page: ls[0].page, box: [Math.min(...ls.map(l => l.box[0])), Math.min(...ls.filter(l => l.page === ls[0].page).map(l => l.box[1])), Math.max(...ls.map(l => l.box[2])), Math.max(...ls.filter(l => l.page === ls[0].page).map(l => l.box[3]))] }; }
      if (z && z.page && z.box) zonesM.push({ cell: `D${resM.firstRow + i}`, file: pdf, page: z.page, rect: z.box, text: `Nouvelle version p.${z.page}` });
    });
    await A.XL.replaceAutoZones(SHEETS.modifications, zonesM);
    if (!identical) {
      const nv = { version, date: new Date().toLocaleString("fr-FR"), pdf, hash, items: snap };
      const keep = [...hist.slice(-2), nv];
      await A.XL.writeRaw(SHEETS.versions, [VERS_HEAD, ...keep.flatMap(versionRows)]);
    }
    await A.reloadSnips();
    const effTab = (() => { const tb = annex.tables.find(x => x.type === "effectif"); const r = tb?.rows.find(r => /^total/.test(r.nl)); const c = r && Object.values(r.cells)[0]; return c ? c.value : null; })();
    const effTxt = annex.facts.find(f => f.kind === "effectif")?.value;
    const allTxt = pages.slice(0, 25).flatMap(p => p.items.map(i => i.str)).join(" ");
    const meta = { bilan: matched.get("A_TG")?.values?.net ?? null, ca: matched.get("R_CA")?.values?.n ?? null, effectif: effTxt ?? effTab ?? null,
      societe: null,
      cloture: (/exercice clos le (\d{2}\/\d{2}\/\d{4})/i.exec(allTxt) || [])[1] || (src.info ? src.info.fin.replace(/(\d{4})(\d{2})(\d{2})/, "$3/$2/$1") : "") };
    F.cache = { pdf, pages, st, matched, postes: map.postes, balance: src.balance, rows, lines };
    F.last = { version: { n: version, prev: prev ? prev.version : null, identical, changes: ch.length, repris: nRepris, revoir: nRevoir, textChanged: !!(cmp && cmp.textChanged) }, meta, pdf, rows, firstRow: res.firstRow, controles: ctr, stats: t.stats, pages: st.pages, missing, annex: { rows: ca.rows, firstRow: resA.firstRow, stats: ca.stats, pages: annex.pages } };
    renderResult();
    await A.XL.activateSheet(SHEETS.pointage, `D${res.firstRow}`);
    const nAno = ctr.filter(c => c.statut === "Anomalie").length;
    const vtxt = prev ? (identical ? ` Version V${version} identique à la précédente.` : ` Version V${version} : ${ch.length} modification(s) depuis V${prev.version}, ${nRepris} validation(s) reprise(s), ${nRevoir} ligne(s) à revoir.`) : "";
    status(`Pointage terminé.${vtxt} Comptes annuels : ${t.stats.concordants} concordant(s), ${t.stats.ecarts} écart(s). Annexe : ${ca.stats.concordants} concordant(s), ${ca.stats.ecarts} écart(s). ${nAno} anomalie(s) de contrôle.` + (missing.length ? ` Section(s) non trouvée(s) : ${missing.map(s => SECTION_LABEL[s]).join(", ")}.` : ""), t.stats.ecarts || ca.stats.ecarts || nAno || missing.length ? "err" : "ok");
  } catch (e) { status("Pointage interrompu : " + e.message, "err"); console.error(e); }
  finally { $("#fsRun").disabled = false; }
}

// ---------- Analyse du texte de l'annexe (Claude for Excel ou conversation Claude)
async function annexContext(pdf) {
  if (!F.cache || F.cache.pdf !== pdf) {
    const map = await loadMapping(); const pages = await pdfPages(pdf); const st = extractStatements(pages, map.postes);
    F.cache = { pdf, pages, st, matched: matchPostes(map.postes, st), postes: map.postes, balance: F.fec?.balance || (await balanceFromSheet())?.balance || [], rows: F.last?.rows || [] };
  }
  if (!F.cache.lines) F.cache.lines = annexLines(annexPages(F.cache.pages, F.cache.st.pages));
  return F.cache;
}
const headerIdx = (v, first) => v ? v.findIndex(r => String(r[0]).trim() === first) : -1;

async function prepareAI() {
  const pdf = $("#fsPdf").value || F.last?.pdf;
  if (!F.last || F.last.pdf !== pdf) return status("Lancez d'abord le pointage (étape 3) sur cette plaquette : l'analyse s'appuie sur ses résultats.", "err");
  try {
    $("#aiPrep").disabled = true; status("Préparation de l'analyse de l'annexe…");
    const C = await annexContext(pdf); const lines = C.lines;
    if (!lines.length) throw new Error("aucune page d'annexe reconnue dans la plaquette");
    const m = F.last.meta;
    if (!m.societe) { const d = lines.map(l => /D[ée]signation de la soci[ée]t[ée]\s*:\s*(.+)/i.exec(l.text)).find(Boolean); m.societe = d ? d[1].trim() : pdf.replace(/\.pdf$/i, ""); }
    const cat = categorie(m, SEUILS);
    // 1) texte de l'annexe
    const resT = await A.XL.writeTable(SHEETS.texte, { title: "Texte de l'annexe (une ligne du PDF par ligne)", subtitle: [`Plaquette : ${pdf} · ${lines.length} lignes · cliquez sur un texte pour le voir dans le PDF`], header: ["ID", "Page", "Texte"], rows: lines.map(l => [l.id, l.page, l.text]), widths: [60, 45, 900] });
    await A.XL.replaceAutoZones(SHEETS.texte, lines.map((l, i) => ({ cell: `C${resT.firstRow + i}`, file: pdf, page: l.page, rect: l.box, text: `Annexe p.${l.page}` })));
    // 2) liste de contrôle (les réponses déjà présentes sont conservées)
    const old = await A.XL.readSheet(SHEETS.conformite); const oh = headerIdx(old, "ID"); const prev = new Map();
    if (oh >= 0) for (const r of old.slice(oh + 1)) if (String(r[0]).trim()) prev.set(String(r[0]).trim(), r);
    const applicable = niv => cat.code === "micro" ? "Annexe non requise" : niv === "Socle" ? "Oui" : cat.code === "petite" ? "Facultatif (annexe simplifiée)" : "Oui";
    const ids = new Set(CHECKLIST.map(c => c[0]));
    const extra = [...prev.entries()].filter(([id]) => !ids.has(id)).map(([, r]) => r);   // points ajoutés par le cabinet
    // nouvelle version de la plaquette : un point dont la citation n'existe plus, ou un point « Absent » alors que le texte a changé, est à revoir
    const V = F.last.version;
    const revise = (p) => {
      if (!p || !String(p[5] || "").trim() || !V || !V.prev || V.identical) return p;
      const st = String(p[5]).trim(); const cit = String(p[7] || "").trim(); const q = p.slice();
      const gone = cit && verifyCitation(cit, lines, Number(p[8]) || null).statut === "Citation introuvable";
      const absentStale = !cit && V.textChanged && ["Absent", "Non applicable", "À vérifier"].includes(st);
      if (gone || absentStale) { q[5] = "À revoir"; q[6] = `[V${V.prev} : ${st}] ${p[6] || ""}`.trim(); q[9] = ""; }
      return q;
    };
    for (const [k, p] of prev) prev.set(k, revise(p));
    const rowsC = [...CHECKLIST.map(c => { const p = prev.get(c[0]); return [c[0], c[1], p ? p[2] || c[2] : c[2], c[3], applicable(c[3]), p?.[5] ?? "", p?.[6] ?? "", p?.[7] ?? "", p?.[8] ?? "", "", p?.[10] ?? ""]; }), ...extra.map(r => [...r.slice(0, 9), "", r[10] ?? ""])];
    const eur0 = v => v == null ? "?" : Math.round(v).toLocaleString("fr-FR");
    const resC = await A.XL.writeTable(SHEETS.conformite, {
      title: "Conformité de l'annexe – analyse par Claude, citations vérifiées par l'outil",
      subtitle: [`Plaquette : ${pdf} · Société : ${m.societe} · Exercice clos le ${m.cloture}`,
        `Catégorie : ${cat.lib} → ${cat.annexe} (bilan ${eur0(m.bilan)} €, chiffre d'affaires ${eur0(m.ca)} €, effectif ${m.effectif ?? "?"} ; seuils du décret n° 2024-152 : ne pas dépasser 2 des 3 seuils, sur deux exercices consécutifs à confirmer).`,
        `Statuts : ${STATUTS.join(", ")}. Colonnes F à I : remplies par Claude. Colonne J : vérification automatique de la citation. Colonne K : votre conclusion.`,
        "Liste de contrôle par défaut à valider par le cabinet : vous pouvez modifier les colonnes A à D ou ajouter des lignes (ID unique)."],
      header: ["ID", "Thème", "Point de contrôle", "Niveau", "Applicable", "Statut", "Justification", "Citation exacte", "Page", "Vérification de la citation", "Conclusion de l'auditeur"],
      rows: rowsC, widths: [45, 120, 330, 70, 110, 90, 320, 320, 45, 150, 200], wrapCols: [2, 6, 7]
    });
    const oldI = await A.XL.readSheet(SHEETS.incoherences); const ih = headerIdx(oldI, "N°");
    const rowsI = ih >= 0 ? oldI.slice(ih + 1).filter(r => r.some(x => String(x).trim())).map(r => [...r.slice(0, 7), "", r[8] ?? ""]) : [];
    const resI = await A.XL.writeTable(SHEETS.incoherences, {
      title: "Incohérences relevées dans l'annexe", subtitle: [`Plaquette : ${pdf} · colonnes A à G remplies par Claude, H par l'outil (vérification de la citation), I par l'auditeur.`],
      header: ["N°", "Constat", "Citation exacte", "Page", "Montant cité", "Montant de référence", "Source de la référence", "Vérification de la citation", "Conclusion de l'auditeur"],
      rows: rowsI, numCols: [4, 5], widths: [35, 330, 330, 45, 100, 110, 220, 150, 200], wrapCols: [1, 2, 6]
    });
    F.ai = { pdf, rowC: resC.firstRow, rowI: resI.firstRow + rowsI.length, cat };
    const meta = { ...m, cat, rowC: resC.firstRow, rowI: resI.firstRow + rowsI.length };
    F.ai.instruction = instructionExcel(meta);
    const comptes = (F.last.rows || []).map(r => [`${r.sectionLib} – ${r.poste}`, r.colonne, r.plaquette ?? "", r.balance]);
    F.ai.dossier = dossierConversation(meta, lines, comptes, C.balance, CHECKLIST.map(c => [...c, applicable(c[3])]));
    $("#aiOut").value = F.ai.instruction; $("#aiBox").hidden = false;
    await A.XL.activateSheet(SHEETS.conformite, `F${resC.firstRow}`);
    status(`Analyse préparée (${lines.length} lignes d'annexe, ${rowsC.length} points de contrôle). Copiez la consigne dans Claude for Excel, puis cliquez sur « Vérifier les citations » quand Claude a terminé.`, "ok");
  } catch (e) { status("Préparation impossible : " + e.message, "err"); console.error(e); }
  finally { $("#aiPrep").disabled = false; }
}

async function copyText(t) {
  try { await navigator.clipboard.writeText(t); return true; }
  catch (e) { const ta = $("#aiOut"); ta.value = t; ta.focus(); ta.select(); try { return document.execCommand("copy"); } catch (e2) { return false; } }
}

async function verifyAI() {
  const pdf = $("#fsPdf").value || F.ai?.pdf || F.last?.pdf;
  if (!pdf) return status("Choisissez la plaquette analysée.", "err");
  try {
    $("#aiVerif").disabled = true; status("Vérification des citations dans la plaquette…");
    const C = await annexContext(pdf);
    const sum = { conf: {}, verif: {}, nonConf: [], inco: [] };
    // Conformité
    const v = await A.XL.readSheet(SHEETS.conformite); const h = headerIdx(v, "ID");
    if (h < 0) throw new Error(`feuille « ${SHEETS.conformite} » absente : cliquez d'abord sur « Préparer l'analyse »`);
    const outC = [], fills = [], zones = [];
    v.slice(h + 1).forEach((r, i) => {
      const row = h + 2 + i; if (!String(r[0]).trim()) { outC.push([""]); return; }
      const st = String(r[5] || "").trim(); const cit = String(r[7] || "").trim();
      let res;
      if (!st) res = { statut: "" };
      else if (st === "À revoir") res = { statut: "En attente : à revoir par Claude" };
      else if (!STATUTS.includes(st)) res = { statut: "Statut non reconnu" };
      else if (!cit) res = { statut: st === "Absent" || st === "Non applicable" ? "Sans objet (pas de citation)" : "Sans citation – à justifier" };
      else res = verifyCitation(cit, C.lines, Number(r[8]) || null);
      const lab = res.statut + (res.autrePage ? ` (trouvée p.${res.page})` : "");
      outC.push([lab]);
      if (res.statut) fills.push({ cell: `J${row}`, color: STATUS_COLOR[res.statut] || (res.statut.startsWith("Sans citation") ? "#FFEB9C" : "#F2F2F2") });
      if (res.box) zones.push({ cell: `H${row}`, file: pdf, page: res.page, rect: res.box, text: `Citation ${r[0]} – annexe p.${res.page}` });
      if (st) { sum.conf[st] = (sum.conf[st] || 0) + 1; sum.verif[res.statut] = (sum.verif[res.statut] || 0) + 1; }
      if (st && (!["Conforme", "Non applicable"].includes(st) || /introuvable|non reconnu|à justifier/.test(lab))) sum.nonConf.push({ id: r[0], point: r[2], statut: st, justif: r[6], verif: lab, row, page: res.page, box: res.box, sheet: SHEETS.conformite });
    });
    if (outC.length) await A.XL.writeCells(SHEETS.conformite, h + 2, 9, outC, fills);
    await A.XL.replaceAutoZones(SHEETS.conformite, zones);
    // Incohérences
    const w = await A.XL.readSheet(SHEETS.incoherences); const hi = headerIdx(w, "N°"); const outI = [], fI = [], zI = [];
    if (hi >= 0) w.slice(hi + 1).forEach((r, i) => {
      const row = hi + 2 + i; if (!String(r[1] || "").trim()) { outI.push([""]); return; }
      const res = String(r[2] || "").trim() ? verifyCitation(r[2], C.lines, Number(r[3]) || null) : { statut: "Sans citation – à justifier" };
      const lab = res.statut + (res.autrePage ? ` (trouvée p.${res.page})` : "");
      outI.push([lab]); fI.push({ cell: `H${row}`, color: STATUS_COLOR[res.statut] || "#FFEB9C" });
      if (res.box) zI.push({ cell: `C${row}`, file: pdf, page: res.page, rect: res.box, text: `Incohérence ${r[0]} – annexe p.${res.page}` });
      sum.inco.push({ id: r[0], point: r[1], statut: "Incohérence", justif: `${r[4] !== "" ? "cité " + r[4] : ""}${r[5] !== "" ? " · référence " + r[5] : ""}`, verif: lab, row, page: res.page, box: res.box, sheet: SHEETS.incoherences });
    });
    if (outI.length) await A.XL.writeCells(SHEETS.incoherences, hi + 2, 7, outI, fI);
    await A.XL.replaceAutoZones(SHEETS.incoherences, zI);
    await A.reloadSnips();
    F.aiResult = sum; renderAI();
    const intro = Object.entries(sum.verif).filter(([k]) => /introuvable|non reconnu|à justifier/.test(k)).reduce((a, [, n]) => a + n, 0) + sum.inco.filter(x => /introuvable|à justifier/.test(x.verif)).length;
    if (!Object.keys(sum.conf).length && !sum.inco.length) status("Aucune réponse de Claude dans les feuilles pour l'instant : collez la consigne dans Claude for Excel, ou importez la réponse d'une conversation.", "err");
    else status(`Vérification terminée : ${Object.values(sum.conf).reduce((a, b) => a + b, 0)} point(s) renseigné(s), ${sum.nonConf.length} non conforme(s) ou à vérifier, ${sum.inco.length} incohérence(s). ${intro ? intro + " citation(s) introuvable(s) ou manquante(s) : à contrôler." : "Toutes les citations ont été retrouvées dans la plaquette."}`, intro ? "err" : "ok");
  } catch (e) { status("Vérification impossible : " + e.message, "err"); console.error(e); }
  finally { $("#aiVerif").disabled = false; }
}

async function importAI() {
  try {
    const rep = parseReponse($("#aiIn").value);
    const v = await A.XL.readSheet(SHEETS.conformite); const h = headerIdx(v, "ID");
    if (h < 0) throw new Error("cliquez d'abord sur « Préparer l'analyse »");
    const byId = new Map(rep.controles.map(c => [c.id, c]));
    const block = v.slice(h + 1).map(r => { const c = byId.get(String(r[0]).trim()); return c ? [c.statut, c.justification, c.citation, c.page] : [r[5] ?? "", r[6] ?? "", r[7] ?? "", r[8] ?? ""]; });
    if (block.length) await A.XL.writeCells(SHEETS.conformite, h + 2, 5, block);
    const w = await A.XL.readSheet(SHEETS.incoherences); const hi = headerIdx(w, "N°");
    const n0 = hi >= 0 ? w.slice(hi + 1).filter(r => String(r[1] || "").trim()).length : 0;
    if (rep.incoherences.length) await A.XL.writeCells(SHEETS.incoherences, hi + 2 + n0, 0, rep.incoherences.map((c, i) => [n0 + i + 1, c.constat, c.citation, c.page, c.montant_cite, c.montant_reference, c.source]));
    $("#aiIn").value = "";
    status(`Réponse importée : ${rep.controles.length} point(s), ${rep.incoherences.length} incohérence(s). Vérification des citations…`, "ok");
    await verifyAI();
  } catch (e) { status("Import impossible : " + e.message, "err"); }
}

function renderAI() {
  const R = F.aiResult; const el = $("#aiResult"); if (!R) { el.innerHTML = ""; return; }
  const items = [...R.nonConf, ...R.inco];
  el.innerHTML = `<p class="muted">${Object.entries(R.conf).map(([k, n]) => `${esc(k)} : ${n}`).join(" · ") || "aucun statut"}</p>` +
    items.map((x, i) => `<div class="snip ${/introuvable|non reconnu|à justifier/.test(x.verif) ? "exception" : "somme"}" data-ai="${i}"><b>${esc(x.id)} – ${esc(x.statut)}</b><div>${esc(String(x.point).slice(0, 160))}</div><div class="f">${esc(String(x.justif || "").slice(0, 200))}</div><div class="f">${esc(x.verif)}</div></div>`).join("");
  el.querySelectorAll("[data-ai]").forEach(d => d.onclick = async () => { const x = items[+d.dataset.ai]; await A.XL.activateSheet(x.sheet, `${x.sheet === SHEETS.conformite ? "H" : "C"}${x.row}`); if (x.page) await A.openDoc(F.ai?.pdf || F.last?.pdf || $("#fsPdf").value, x.page, x.box); });
}

function renderResult() {
  const L = F.last; if (!L) return;
  const ec = L.rows.map((r, i) => ({ r, i })).filter(x => x.r.statut !== "Concordant" && x.r.statut !== "Concordant (arrondi)");
  const an = L.controles.filter(c => c.statut === "Anomalie");
  const pg = s => (L.pages[s] || []).join(", ") || "—";
  const V = L.version;
  $("#fsResult").innerHTML = (V && V.prev ? `<div class="cellinfo" data-mod="1" style="cursor:pointer"><b>Version V${V.n}</b> ${V.identical ? "identique à la version précédente" : `comparée à V${V.prev} : <b>${V.changes}</b> modification(s) · ${V.repris} validation(s) reprise(s) · <b style="color:var(--exception)">${V.revoir}</b> ligne(s) à revoir`} <span class="muted">(feuille « ${SHEETS.modifications} »)</span></div>` : "") + `<p class="muted">Pages lues : bilan actif p.${pg("actif")} · passif p.${pg("passif")} · compte de résultat p.${pg("cr")}</p>
    <p><b>${L.stats.concordants}</b> concordant(s) · <b style="color:var(--exception)">${L.stats.ecarts}</b> écart(s) · ${an.length} anomalie(s)</p>` +
    (ec.length ? `<div class="fsh">Comptes annuels : points à examiner</div>` + ec.map(x => `<div class="snip exception" data-i="${x.i}"><b>${esc(x.r.poste)}</b> <span class="muted">${esc(x.r.colonne)}</span><div>${esc(x.r.statut)} : ${esc(eur(x.r.ecart))}</div>${x.r.commentaire ? `<div class="f">${esc(x.r.commentaire)}</div>` : ""}</div>`).join("") : "") +
    (L.annex ? `<div class="fsh">Annexe (pages ${L.annex.pages.join(", ") || "—"}) : ${L.annex.stats.concordants} concordant(s) · ${L.annex.stats.ecarts} écart(s)</div>` + L.annex.rows.map((r, i) => ({ r, i })).filter(x => !x.r.statut.startsWith("Concordant")).map(x => `<div class="snip ${x.r.statut.includes("analyser") ? "somme" : "exception"}" data-a="${x.i}"><b>${esc(x.r.note)}</b> <span class="muted">p.${x.r.page}</span><div>${esc(x.r.ligne)} · ${esc(x.r.colonne)}</div><div>${esc(x.r.statut)} : ${esc(eur(x.r.ecart))} <span class="muted">(${esc(x.r.reference)})</span></div>${x.r.commentaire ? `<div class="f">${esc(x.r.commentaire.slice(0, 220))}</div>` : ""}</div>`).join("") : "") +
    (an.length ? `<div class="fsh">Anomalies de contrôle</div>` + an.map((c, k) => `<div class="snip somme" data-c="${k}"><b>${esc(c.controle)}</b><div class="f">${esc(c.detail)}</div></div>`).join("") : "");
  $("#fsResult").querySelectorAll("[data-i]").forEach(d => d.onclick = async () => {
    const x = L.rows[+d.dataset.i]; const row = L.firstRow + +d.dataset.i;
    await A.XL.activateSheet(SHEETS.pointage, `D${row}`);
    if (x.page) await A.openDoc(L.pdf, x.page, x.box);
  });
  const md = $("#fsResult").querySelector("[data-mod]"); if (md) md.onclick = () => A.XL.activateSheet(SHEETS.modifications, "D1");
  $("#fsResult").querySelectorAll("[data-a]").forEach(d => d.onclick = async () => {
    const x = L.annex.rows[+d.dataset.a]; await A.XL.activateSheet(SHEETS.annexe, `D${L.annex.firstRow + +d.dataset.a}`);
    if (x.page) await A.openDoc(L.pdf, x.page, x.box);
  });
  $("#fsResult").querySelectorAll("[data-c]").forEach(d => d.onclick = async () => { const c = an[+d.dataset.c]; if (c.page) await A.openDoc(L.pdf, c.page, c.box); });
}

async function fillPdfSelect() {
  const files = (await A.store.listFiles()).filter(f => /\.pdf$/i.test(f.name));
  const sel = $("#fsPdf"); const cur = sel.value;
  const guess = cur || (files.find(f => /plaquette|comptes?.annuels|etats?.financiers|bilan/i.test(f.name)) || {}).name || (A.viewer.name || "");
  sel.innerHTML = `<option value="">— choisir la plaquette —</option>` + files.map(f => `<option value="${esc(f.name)}" ${f.name === guess ? "selected" : ""}>${esc(f.path)}</option>`).join("");
}

export function initFS(deps) {
  A = deps;
  $("#btnFS").addEventListener("click", async () => { document.querySelectorAll("aside.panel").forEach(p => p.hidden = true); $("#fsPanel").hidden = false; await fillPdfSelect(); renderResult(); });
  $("#closeFS").addEventListener("click", () => { $("#fsPanel").hidden = true; });
  $("#fsFec").addEventListener("change", async e => {
    const f = e.target.files[0]; e.target.value = ""; if (!f) return;
    try {
      status(`Lecture du FEC « ${f.name} »…`);
      F.fec = parseFEC(await f.arrayBuffer()); F.fecName = f.name;
      const an = F.fec.controles.filter(c => c.statut === "Anomalie").length;
      $("#fecInfo").innerHTML = `<b>${esc(f.name)}</b><br>${F.fec.info.lignes.toLocaleString("fr-FR")} lignes · ${F.fec.balance.length} comptes · résultat ${eur(F.fec.resultat)} €` + (an ? `<br><span style="color:var(--exception)">${an} anomalie(s) du FEC (détail dans la feuille ${SHEETS.controles})</span>` : "");
      status("FEC chargé. Il n'est pas enregistré dans le classeur : seule la balance reconstituée le sera.", "ok");
    } catch (err) { F.fec = null; $("#fecInfo").textContent = "Aucun FEC chargé."; status("FEC illisible : " + err.message, "err"); }
  });
  $("#fsAddPdf").addEventListener("change", async e => { const fl = [...e.target.files]; e.target.value = ""; if (!fl.length) return; await A.addFiles(fl); $("#fsPdf").value = ""; await fillPdfSelect(); if (fl[0]) $("#fsPdf").value = fl[0].name; });
  $("#fsRun").addEventListener("click", run);
  $("#aiPrep").addEventListener("click", prepareAI);
  $("#aiVerif").addEventListener("click", verifyAI);
  $("#aiImport").addEventListener("click", importAI);
  $("#aiCopyExcel").addEventListener("click", async () => { if (!F.ai) return status("Cliquez d'abord sur « Préparer l'analyse ».", "err"); $("#aiOut").value = F.ai.instruction; status(await copyText(F.ai.instruction) ? "Consigne copiée : collez-la dans le volet Claude for Excel (Ctrl+V), puis envoyez." : "Copie automatique impossible : sélectionnez le texte ci-dessous (Ctrl+A, Ctrl+C).", "ok"); });
  $("#aiCopyChat").addEventListener("click", async () => { if (!F.ai) return status("Cliquez d'abord sur « Préparer l'analyse ».", "err"); $("#aiOut").value = F.ai.dossier; status(await copyText(F.ai.dossier) ? "Dossier copié : collez-le dans une conversation Claude, puis collez la réponse JSON ci-dessous et cliquez sur « Importer »." : "Copie automatique impossible : sélectionnez le texte ci-dessous (Ctrl+A, Ctrl+C).", "ok"); });
  $("#fsResetMap").addEventListener("click", async () => { await writeMapping(DEFAULT_POSTES); status(`Feuille « ${SHEETS.mapping} » réinitialisée avec le plan de regroupement par défaut.`, "ok"); });
}
export { F as __fsState };
