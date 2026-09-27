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
