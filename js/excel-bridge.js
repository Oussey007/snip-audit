/* global Office, Excel */
// Passerelle Excel (Office.js). Toutes les données des snips sont stockées DANS le classeur (feuille masquée « _Snips »),
// ce qui les rend portables (Windows, Mac, Excel en ligne) et auditables.
export const SNIP_SHEET = "_Snips";
const SNIP_HEAD = ["ID", "Feuille", "Cellule", "Type", "Fichier", "Page", "X1", "Y1", "X2", "Y2", "Valeur", "Texte extrait", "Date", "Remplissage initial"];
export const COLORS = { texte: "#DDEBF7", somme: "#E4DFEC", valide: "#C6EFCE", exception: "#FFC7CE" };
export const BORDER = { texte: "#2F80ED", somme: "#7030A0", valide: "#1E9E4A", exception: "#E02424" };

export function ready() {
  return new Promise(res => Office.onReady(info => res(info)));
}
export function onSelection(cb) {
  let t = null;
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => { clearTimeout(t); t = setTimeout(cb, 120); });
}
const clean = a => (a || "").split("!").pop().replace(/\$/g, "");
const sheetOf = a => { const p = (a || "").split("!"); return p.length > 1 ? p[0].replace(/^'|'$/g, "").replace(/''/g, "'") : null; };

export async function getSelection() {
  return Excel.run(async ctx => {
    const r = ctx.workbook.getSelectedRange();
    r.load("address,rowIndex,columnIndex");
    const ws = r.worksheet; ws.load("name");
    const c = r.getCell(0, 0); c.load("address,values,formulas,numberFormat");
    await ctx.sync();
    const hdr = ws.getCell(0, r.columnIndex); hdr.load("values");
    const row = ws.getRangeByIndexes(r.rowIndex, 0, 1, 16); row.load("values");
    await ctx.sync();
    return { sheet: ws.name, cell: clean(c.address), row: r.rowIndex, col: r.columnIndex, value: c.values[0][0], formula: c.formulas[0][0], header: String(hdr.values[0][0] || ""), rowValues: row.values[0] };
  });
}

async function sheetValues(ctx, name) {
  const ws = ctx.workbook.worksheets.getItemOrNullObject(name);
  await ctx.sync();
  if (ws.isNullObject) return null;
  const u = ws.getUsedRangeOrNullObject(true); u.load("values,rowCount");
  await ctx.sync();
  return u.isNullObject ? [] : u.values;
}

// Onglet « Liens » : Référence écriture | Ordre | Rôle | Fichier | Dossier | Chemin relatif | Page
export async function readLiens() {
  return Excel.run(async ctx => {
    const v = await sheetValues(ctx, "Liens");
    const map = new Map();
    if (!v || v.length < 2) return map;
    const h = v[0].map(x => String(x).toLowerCase());
    const col = k => h.findIndex(x => x.includes(k));
    const cR = col("référence") >= 0 ? col("référence") : 0, cO = col("ordre"), cRole = col("rôle"), cF = col("fichier"), cP = col("page");
    for (const r of v.slice(1)) {
      const ref = String(r[cR] || "").trim(); if (!ref || cF < 0 || !r[cF]) continue;
      if (!map.has(ref)) map.set(ref, []);
      map.get(ref).push({ ordre: cO >= 0 ? +r[cO] || 0 : 0, role: cRole >= 0 ? String(r[cRole]) : "Pièce", file: String(r[cF]).trim(), page: cP >= 0 ? +r[cP] || 1 : 1 });
    }
    for (const l of map.values()) l.sort((a, b) => a.ordre - b.ordre);
    return map;
  });
}

export async function readSnips() {
  return Excel.run(async ctx => {
    const v = await sheetValues(ctx, SNIP_SHEET);
    if (!v || v.length < 2) return [];
    return v.slice(1).filter(r => r[0]).map(r => ({ id: String(r[0]), sheet: String(r[1]), cell: String(r[2]), type: String(r[3]), file: String(r[4]), page: +r[5], rect: [+r[6], +r[7], +r[8], +r[9]], value: r[10], text: String(r[11] || ""), date: String(r[12] || ""), prevFill: String(r[13] || "") }));
  });
}

async function ensureSnipSheet(ctx) {
  let ws = ctx.workbook.worksheets.getItemOrNullObject(SNIP_SHEET);
  await ctx.sync();
  if (ws.isNullObject) {
    ws = ctx.workbook.worksheets.add(SNIP_SHEET);
    ws.getRangeByIndexes(0, 0, 1, SNIP_HEAD.length).values = [SNIP_HEAD];
    ws.getRangeByIndexes(0, 0, 1, SNIP_HEAD.length).format.font.bold = true;
    ws.visibility = Excel.SheetVisibility.hidden;
    await ctx.sync();
  }
  return ws;
}

// Écrit (ou remplace) le snip d'une cellule, met à jour la cellule et son format.
export async function saveSnip(rec, writeValue) {
  return Excel.run(async ctx => {
    const target = ctx.workbook.worksheets.getItem(rec.sheet).getRange(rec.cell);
    target.load("values,formulas,format/fill/color");
    const ws = await ensureSnipSheet(ctx);
    const u = ws.getUsedRange(); u.load("values,rowCount");
    await ctx.sync();
    const rows = u.values;
    let idx = rows.findIndex((r, i) => i > 0 && r[1] === rec.sheet && r[2] === rec.cell && r[3] !== "auto");
    const prevFill = idx > 0 ? rows[idx][13] : (target.format.fill.color || "");
    const hasFormula = typeof target.formulas[0][0] === "string" && target.formulas[0][0].startsWith("=");
    const current = target.values[0][0];
    let written = false;
    if (writeValue !== undefined && writeValue !== null) {
      if (rec.type === "valide" || rec.type === "exception") {
        if (current === "" || current === null) { target.values = [[writeValue]]; written = true; }
      } else if (!hasFormula) { target.values = [[writeValue]]; written = true; }
    }
    target.format.fill.color = COLORS[rec.type];
    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(e => { const b = target.format.borders.getItem(e); b.style = "Continuous"; b.color = BORDER[rec.type]; b.weight = "Medium"; });
    const line = [rec.id, rec.sheet, rec.cell, rec.type, rec.file, rec.page, ...rec.rect.map(x => Math.round(x * 100) / 100), rec.value === undefined ? "" : rec.value, (rec.text || "").slice(0, 2000), rec.date, prevFill];
    const at = idx > 0 ? idx : rows.length;
    ws.getRangeByIndexes(at, 0, 1, line.length).values = [line];
    await ctx.sync();
    return { written, hasFormula };
  });
}

export async function deleteSnip(sheet, cell) {
  return Excel.run(async ctx => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(SNIP_SHEET);
    await ctx.sync(); if (ws.isNullObject) return false;
    const u = ws.getUsedRange(); u.load("values"); await ctx.sync();
    const idx = u.values.findIndex((r, i) => i > 0 && r[1] === sheet && r[2] === cell && r[3] !== "auto");
    if (idx < 1) return false;
    const prev = u.values[idx][13];
    const target = ctx.workbook.worksheets.getItem(sheet).getRange(cell);
    if (prev && prev !== "#FFFFFF") target.format.fill.color = prev; else target.format.fill.clear();
    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(e => { target.format.borders.getItem(e).style = "None"; });
    ws.getRangeByIndexes(idx, 0, 1, SNIP_HEAD.length).delete(Excel.DeleteShiftDirection.up);
    await ctx.sync();
    return true;
  });
}

export async function selectCell(sheet, cell) {
  return Excel.run(async ctx => { const ws = ctx.workbook.worksheets.getItem(sheet); ws.activate(); ws.getRange(cell).select(); await ctx.sync(); });
}

// ---------- Pièces enregistrées DANS le classeur (feuille « _Pieces », masquée de façon permanente)
// Une ligne par pièce : A nom | B chemin | C taille (octets) | D nb de morceaux | E ajoutée le | F… contenu PDF en base64 (morceaux de 32 000 caractères)
export const PIECE_SHEET = "_Pieces";
const PH = ["Nom", "Chemin", "Taille", "Morceaux", "Ajoutée le"], META = PH.length, CH = 32000, SLICE = 60;
async function pieceSheet(ctx, create) {
  let ws = ctx.workbook.worksheets.getItemOrNullObject(PIECE_SHEET);
  await ctx.sync();
  if (ws.isNullObject) {
    if (!create) return null;
    ws = ctx.workbook.worksheets.add(PIECE_SHEET);
    ws.getRangeByIndexes(0, 0, 1, META).values = [PH];
    ws.visibility = Excel.SheetVisibility.veryHidden;
    await ctx.sync();
  }
  return ws;
}
async function pieceRows(ctx, ws) {
  const u = ws.getUsedRangeOrNullObject(true); u.load("rowCount"); await ctx.sync();
  if (u.isNullObject || u.rowCount < 2) return [];
  const r = ws.getRangeByIndexes(1, 0, u.rowCount - 1, META); r.load("values"); await ctx.sync();
  return r.values.map((v, i) => ({ row: i + 1, name: String(v[0] || ""), path: String(v[1] || ""), size: +v[2] || 0, chunks: +v[3] || 0, added: String(v[4] || "") })).filter(x => x.name);
}
export async function listPieces() {
  return Excel.run(async ctx => { const ws = await pieceSheet(ctx, false); return ws ? pieceRows(ctx, ws) : []; });
}
export async function addPiece(name, path, size, b64) {
  const parts = []; for (let i = 0; i < b64.length; i += CH) parts.push("~" + b64.slice(i, i + CH));
  return Excel.run(async ctx => {
    const ws = await pieceSheet(ctx, true);
    const rows = await pieceRows(ctx, ws);
    const ex = rows.find(x => x.name === name);
    let row;
    if (ex) { row = ex.row; ws.getRange(`${row + 1}:${row + 1}`).clear(); }
    else { const u = ws.getUsedRange(true); u.load("rowCount"); await ctx.sync(); row = u.rowCount; }
    ws.getRangeByIndexes(row, 0, 1, META).values = [[name, path, size, parts.length, new Date().toLocaleString("fr-FR")]];
    await ctx.sync();
    for (let i = 0; i < parts.length; i += SLICE) {
      const sl = parts.slice(i, i + SLICE);
      ws.getRangeByIndexes(row, META + i, 1, sl.length).values = [sl];
      await ctx.sync();
    }
    return { replaced: !!ex };
  });
}
export async function getPiece(name) {
  return Excel.run(async ctx => {
    const ws = await pieceSheet(ctx, false); if (!ws) return null;
    const p = (await pieceRows(ctx, ws)).find(x => x.name === name); if (!p) return null;
    let out = "";
    for (let i = 0; i < p.chunks; i += SLICE) {
      const r = ws.getRangeByIndexes(p.row, META + i, 1, Math.min(SLICE, p.chunks - i)); r.load("values"); await ctx.sync();
      out += r.values[0].map(s => String(s).replace(/^~/, "")).join("");
    }
    return out;
  });
}
export async function deletePiece(name) {
  return Excel.run(async ctx => {
    const ws = await pieceSheet(ctx, false); if (!ws) return false;
    const p = (await pieceRows(ctx, ws)).find(x => x.name === name); if (!p) return false;
    ws.getRange(`${p.row + 1}:${p.row + 1}`).delete(Excel.DeleteShiftDirection.up);
    await ctx.sync(); return true;
  });
}
export async function clearPieces() {
  return Excel.run(async ctx => {
    const ws = await pieceSheet(ctx, false); if (!ws) return 0;
    const n = (await pieceRows(ctx, ws)).length;
    ws.visibility = Excel.SheetVisibility.hidden; await ctx.sync();
    ws.delete(); await ctx.sync(); return n;
  });
}

// ---------- Pointage de la plaquette : écriture de tableaux et zones automatiques
const colLetter = i => { let s = ""; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
export { colLetter };

// Lit les valeurs d'une feuille (null si absente)
export async function readSheet(name) {
  return Excel.run(async ctx => sheetValues(ctx, name));
}

// Écrit un tableau dans une feuille (créée ou vidée) : opts = { title, subtitle, header, rows, widths, numCols, formulas:{col: fn(rowIndex)}, fills:[{row, color}], startRow }
export async function writeTable(name, opts) {
  return Excel.run(async ctx => {
    let ws = ctx.workbook.worksheets.getItemOrNullObject(name);
    await ctx.sync();
    if (ws.isNullObject) ws = ctx.workbook.worksheets.add(name);
    else { ws.getRange().clear(); }
    let r0 = 0;
    if (opts.title) { const t = ws.getRange("A1"); t.values = [[opts.title]]; t.format.font.bold = true; t.format.font.size = 13; t.format.font.color = "#1F3864"; r0 = 1; }
    if (opts.subtitle && opts.subtitle.length) { ws.getRangeByIndexes(r0, 0, opts.subtitle.length, 1).values = opts.subtitle.map(s => [s]); ws.getRangeByIndexes(r0, 0, opts.subtitle.length, 1).format.font.color = "#4B5563"; r0 += opts.subtitle.length; }
    if (r0) r0++;
    const nc = opts.header.length;
    const h = ws.getRangeByIndexes(r0, 0, 1, nc); h.values = [opts.header];
    h.format.font.bold = true; h.format.font.color = "#FFFFFF"; h.format.fill.color = "#1F3864"; h.format.wrapText = true;
    const rows = opts.rows;
    if (rows.length) {
      const body = ws.getRangeByIndexes(r0 + 1, 0, rows.length, nc);
      const vals = rows.map((r, i) => r.map((v, j) => (opts.formulas && opts.formulas[j]) ? opts.formulas[j](r0 + 2 + i, r) : (v === null || v === undefined ? "" : v)));
      body.values = vals;
      for (const j of opts.numCols || []) ws.getRangeByIndexes(r0 + 1, j, rows.length, 1).numberFormat = rows.map(() => ["#,##0.00;-#,##0.00;0.00"]);
      body.format.verticalAlignment = "Top";
    }
    (opts.widths || []).forEach((w, j) => { ws.getRangeByIndexes(0, j, 1, 1).format.columnWidth = w; });
    for (const f of opts.fills || []) ws.getRangeByIndexes(r0 + 1 + f.row, f.col ?? 0, 1, f.ncol ?? nc).format.fill.color = f.color;
    for (const j of opts.wrapCols || []) if (rows.length) ws.getRangeByIndexes(r0 + 1, j, rows.length, 1).format.wrapText = true;
    ws.freezePanes.freezeRows(r0 + 1);
    await ctx.sync();
    return { headerRow: r0 + 1, firstRow: r0 + 2 };
  });
}

// Remplace les zones automatiques d'une feuille : zones = [{cell, file, page, rect, value, text}]
export async function replaceAutoZones(sheetName, zones) {
  return Excel.run(async ctx => {
    const ws = await ensureSnipSheet(ctx);
    const u = ws.getUsedRange(); u.load("values,rowCount"); await ctx.sync();
    const keep = u.values.filter((r, i) => i === 0 || !(r[3] === "auto" && r[1] === sheetName));
    const date = new Date().toLocaleString("fr-FR");
    const add = zones.map((z, i) => [`fs-${sheetName}-${i}-${Date.now()}`, sheetName, z.cell, "auto", z.file, z.page, ...z.rect.map(x => Math.round(x * 100) / 100), z.value ?? "", String(z.text || "").slice(0, 2000), date, ""]);
    const all = keep.concat(add);
    ws.getRange().clear();
    ws.getRangeByIndexes(0, 0, all.length, SNIP_HEAD.length).values = all.map(r => r.slice(0, SNIP_HEAD.length).concat(Array(Math.max(0, SNIP_HEAD.length - r.length)).fill("")));
    ws.getRangeByIndexes(0, 0, 1, SNIP_HEAD.length).format.font.bold = true;
    await ctx.sync();
    return add.length;
  });
}

export async function activateSheet(name, cell) {
  return Excel.run(async ctx => { const ws = ctx.workbook.worksheets.getItem(name); ws.activate(); if (cell) ws.getRange(cell).select(); await ctx.sync(); });
}

// Écrit un bloc de valeurs à partir d'une cellule (ligne 1-based, colonne 0-based) et colore des cellules
export async function writeCells(name, row, col, values, fills = []) {
  return Excel.run(async ctx => {
    const ws = ctx.workbook.worksheets.getItem(name);
    if (values.length) ws.getRangeByIndexes(row - 1, col, values.length, values[0].length).values = values;
    for (const f of fills) ws.getRange(f.cell).format.fill.color = f.color;
    await ctx.sync();
  });
}

// Feuille de données brutes (historique des versions), masquée
export async function writeRaw(name, values, hidden = true) {
  return Excel.run(async ctx => {
    let ws = ctx.workbook.worksheets.getItemOrNullObject(name); await ctx.sync();
    if (ws.isNullObject) { ws = ctx.workbook.worksheets.add(name); await ctx.sync(); }
    ws.getRange().clear();
    const CH = 2000;
    for (let i = 0; i < values.length; i += CH) { const part = values.slice(i, i + CH); ws.getRangeByIndexes(i, 0, part.length, part[0].length).values = part; await ctx.sync(); }
    if (hidden) ws.visibility = Excel.SheetVisibility.veryHidden;
    await ctx.sync();
  });
}
