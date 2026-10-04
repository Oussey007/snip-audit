// Analyse du texte de l'annexe par Claude : texte de l'annexe, liste de contrôle, consignes, import de la réponse
// et vérification des citations (chaque conclusion doit s'appuyer sur un passage réellement présent dans le PDF).
import { normLabel } from "./pcg.js?v=8";

// Seuils de taille (décret n° 2024-152 du 28 février 2024, exercices ouverts à compter du 1er janvier 2024) : modifiables dans la feuille Paramètres
export const SEUILS = { micro: { bilan: 450000, ca: 900000, effectif: 10 }, petite: { bilan: 7500000, ca: 15000000, effectif: 50 }, moyenne: { bilan: 25000000, ca: 50000000, effectif: 250 } };

// Catégorie : ne pas dépasser deux des trois seuils (sur deux exercices consécutifs : l'outil ne voit que l'exercice N, à confirmer)
export function categorie(d, seuils = SEUILS) {
  const depasse = s => [d.bilan > s.bilan, d.ca > s.ca, d.effectif > s.effectif].filter(Boolean).length;
  const ok = s => d.bilan != null && d.ca != null && d.effectif != null ? depasse(s) < 2 : null;
  if (ok(seuils.micro)) return { code: "micro", lib: "micro-entreprise", annexe: "dispensée d'annexe" };
  if (ok(seuils.petite)) return { code: "petite", lib: "petite entreprise", annexe: "annexe simplifiée (petites entreprises)" };
  if (ok(seuils.moyenne)) return { code: "moyenne", lib: "moyenne entreprise", annexe: "annexe de base complète" };
  if (ok(seuils.moyenne) === false) return { code: "grande", lib: "grande entreprise", annexe: "annexe de base complète" };
  return { code: "?", lib: "catégorie non déterminée (donnée manquante)", annexe: "à déterminer" };
}

// Liste de contrôle par défaut : à valider et compléter par le cabinet (feuille « Conformité annexe », colonnes A à D)
// Niveau « Socle » : attendu dans toute annexe ; « Complète » : annexe de base (moyennes et grandes entreprises), facultatif en annexe simplifiée.
export const CHECKLIST = [
  ["M01", "Règles et méthodes", "Le référentiel comptable appliqué est indiqué (règlement ANC n° 2014-03 et ses modifications).", "Socle"],
  ["M02", "Règles et méthodes", "Les hypothèses de base sont rappelées (continuité d'exploitation, permanence des méthodes, indépendance des exercices) et sont cohérentes avec les faits décrits dans l'annexe.", "Socle"],
  ["M03", "Règles et méthodes", "Les méthodes d'évaluation des postes significatifs du bilan sont décrites : immobilisations (modes et durées d'amortissement), stocks, créances et dépréciations, provisions.", "Socle"],
  ["M04", "Règles et méthodes", "Tout changement de méthode comptable ou de présentation est mentionné, justifié, et son incidence sur la comparabilité avec l'exercice précédent est indiquée.", "Socle"],
  ["M05", "Règles et méthodes", "Les dérogations aux principes comptables éventuelles sont mentionnées et justifiées, avec leur incidence.", "Socle"],
  ["F01", "Faits caractéristiques", "Les faits caractéristiques de l'exercice sont décrits avec leur incidence chiffrée, et les montants cités concordent avec les comptes.", "Socle"],
  ["F02", "Faits caractéristiques", "La situation de continuité d'exploitation est cohérente : en cas d'incertitude (pertes, procédures, difficultés de trésorerie), une information est donnée.", "Socle"],
  ["B01", "Notes sur le bilan", "Les mouvements de l'actif immobilisé (valeurs brutes) sont présentés.", "Socle"],
  ["B02", "Notes sur le bilan", "Les amortissements et dépréciations de l'actif immobilisé sont présentés (mouvements de l'exercice).", "Socle"],
  ["B03", "Notes sur le bilan", "Les provisions sont détaillées par nature avec leurs mouvements, et chaque provision significative est expliquée (objet, montant).", "Socle"],
  ["B04", "Notes sur le bilan", "Les créances et les dettes sont ventilées par échéance (à plus d'un an, à plus de cinq ans), de façon cohérente avec les échéanciers ou étalements décrits dans l'annexe.", "Socle"],
  ["B05", "Notes sur le bilan", "Les dettes garanties par des sûretés réelles sont indiquées (ou l'absence est explicite).", "Socle"],
  ["B06", "Notes sur le bilan", "Les engagements financiers hors bilan sont présentés : crédit-bail, cautions et garanties, engagements de retraite non provisionnés, autres engagements.", "Socle"],
  ["B07", "Notes sur le bilan", "La composition du capital (nombre et valeur nominale des titres) et ses mouvements sont indiqués.", "Socle"],
  ["B08", "Notes sur le bilan", "Les charges à payer, produits à recevoir, charges et produits constatés d'avance sont détaillés.", "Socle"],
  ["B09", "Notes sur le bilan", "Les avances et crédits alloués aux dirigeants sont indiqués (ou l'absence est explicite).", "Socle"],
  ["R01", "Notes sur le compte de résultat", "Le chiffre d'affaires est ventilé par catégorie d'activité et par marché géographique ; tout tableau annoncé est effectivement présent.", "Complète"],
  ["R02", "Notes sur le compte de résultat", "Les produits et charges exceptionnels (ou d'importance ou d'incidence exceptionnelle) sont expliqués.", "Socle"],
  ["R03", "Notes sur le compte de résultat", "L'impôt sur les bénéfices est ventilé entre résultat courant et résultat exceptionnel.", "Complète"],
  ["R04", "Notes sur le compte de résultat", "Le montant des honoraires du commissaire aux comptes est indiqué, s'il en existe un.", "Complète"],
  ["A01", "Autres informations", "L'effectif moyen est indiqué et cohérent entre le texte et le tableau.", "Socle"],
  ["A02", "Autres informations", "Les événements postérieurs à la clôture significatifs sont mentionnés (ou l'absence est explicite).", "Socle"],
  ["A03", "Autres informations", "L'identité de l'entreprise consolidante est indiquée lorsque la société est consolidée.", "Complète"],
  ["A04", "Autres informations", "Les transactions significatives avec les parties liées non conclues aux conditions normales de marché sont indiquées.", "Complète"],
  ["A05", "Autres informations", "La rémunération des organes d'administration, de direction et de surveillance est indiquée (sauf si elle permet d'identifier un membre).", "Complète"],
  ["A06", "Autres informations", "La date et l'organe d'arrêté des comptes sont indiqués et plausibles (postérieurs à la clôture, cohérents avec les autres dates du dossier).", "Socle"]
];
export const STATUTS = ["Conforme", "Incomplet", "Absent", "Incohérent", "Non applicable", "À vérifier"];

// Lignes de texte de l'annexe (une par ligne du PDF) avec leur zone, pour l'analyse et la vérification des citations
export function annexLines(annexPagesAnalysed) {
  const out = [];
  for (const pg of annexPagesAnalysed) {
    let k = 0;
    for (const l of pg.lines) {
      if (l.y > pg.height * 0.93 || l.y < pg.height * 0.03) continue;
      const amounts = l.amounts.map(a => a.raw).join("   ");
      const text = (l.label + (amounts ? "   " + amounts : "")).trim(); if (!text) continue;
      const xs = [l.lx0, l.lx1, ...l.amounts.flatMap(a => [a.x0, a.x1])].filter(v => v != null);
      out.push({ id: `p${pg.n}-${String(++k).padStart(2, "0")}`, page: pg.n, text, tab: l.amounts.length > 0, box: [Math.min(...xs) - 1, l.y - l.h * 0.3, Math.max(...xs) + 1, l.y + l.h * 0.95] });
    }
  }
  return out;
}

// Normalisation pour la recherche des citations : minuscules, sans accents ni ponctuation, chiffres conservés
const nz = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[’']/g, " ").replace(/[^a-z0-9]+/g, " ").trim();

// Vérifie qu'une citation figure dans l'annexe : exacte, ou par morceaux (citation tronquée par « … »), sur une même page
export function verifyCitation(citation, lines, pageHint) {
  const parts = String(citation || "").split(/\.\.\.|…|\[\.\.\.\]/).map(nz).filter(p => p.length >= 12);
  if (!parts.length) return { statut: nz(citation) ? "Trop courte pour être vérifiée" : "Sans citation" };
  const pages = [...new Set(lines.map(l => l.page))].sort((a, b) => (a === pageHint ? -1 : b === pageHint ? 1 : a - b));
  for (const p of pages) {
    const L = lines.filter(l => l.page === p);
    let txt = "", map = [];
    L.forEach((l, i) => { const t = nz(l.text); if (!t) return; if (txt) { txt += " "; map.push(i); } for (let k = 0; k < t.length; k++) map.push(i); txt += t; });
    let from = 0, first = -1, last = -1, ok = true;
    for (const part of parts) { const at = txt.indexOf(part, from); if (at < 0) { ok = false; break; } if (first < 0) first = at; last = at + part.length - 1; from = at + part.length; }
    if (ok) {
      const i0 = map[first], i1 = map[last]; const sel = L.slice(i0, i1 + 1);
      const box = [Math.min(...sel.map(l => l.box[0])), Math.min(...sel.map(l => l.box[1])), Math.max(...sel.map(l => l.box[2])), Math.max(...sel.map(l => l.box[3]))];
      return { statut: parts.length > 1 ? "Citation vérifiée (extraits)" : "Citation vérifiée", page: p, box, autrePage: pageHint && p !== pageHint };
    }
  }
  return { statut: "Citation introuvable" };
}

// Consigne pour Claude for Excel (le classeur contient déjà toutes les feuilles nécessaires)
export function instructionExcel(meta) {
  return `Tu es auditeur financier. Analyse l'annexe des comptes annuels de ${meta.societe || "la société"} (exercice clos le ${meta.cloture || "?"}) à partir des feuilles de ce classeur :
- « Annexe texte » : le texte de l'annexe, une ligne du PDF par ligne (colonne A = identifiant, B = page, C = texte) ;
- « Pointage plaquette » : bilan et compte de résultat (D = montant de la plaquette, E = montant de la balance) ;
- « Pointage annexe » : contrôles chiffrés déjà réalisés sur les tableaux de l'annexe ;
- « Balance FEC » : soldes de tous les comptes (F = solde de clôture, débit positif).
Catégorie retenue : ${meta.cat.lib} → ${meta.cat.annexe}. Les points de niveau « Complète » sont facultatifs pour une petite entreprise : s'ils sont absents, mets « Non applicable ».

TÂCHE 1 – Feuille « Conformité annexe », tableau qui commence ligne ${meta.rowC} (en-têtes ligne ${meta.rowC - 1}). Pour chaque ligne dont la colonne F (Statut) est vide ou « À revoir », renseigne :
F Statut : ${STATUTS.join(" / ")} ;
G Justification : une ou deux phrases, avec les montants utilisés et leur source (feuille, poste ou compte) ;
H Citation exacte : recopie MOT POUR MOT, sans rien reformuler, le passage de la colonne C de « Annexe texte » qui fonde ta conclusion (lignes consécutives, 300 caractères au plus ; si tu coupes, utilise « … ») ; vide si Absent ;
I Page : la page de cette citation (colonne B de « Annexe texte »).
Ne modifie pas les colonnes A à E, ni J et K.

TÂCHE 2 – Feuille « Incohérences annexe », à partir de la ligne ${meta.rowI} : ajoute une ligne par incohérence trouvée entre le texte de l'annexe et les comptes (bilan, compte de résultat, Balance FEC), ou entre deux passages de l'annexe. Exemples : montant cité différent du montant comptabilisé ; numéro de compte cité différent de celui de la balance ; provision ou opération comptabilisée mais non expliquée ; explication qui ne couvre qu'une partie d'un montant ; date incohérente ; tableau annoncé mais absent ; texte qui semble repris d'un exercice antérieur.
Colonnes : A n° (1, 2, 3…), B Constat, C Citation exacte (mot pour mot, comme ci-dessus), D Page, E Montant cité, F Montant de référence, G Source de la référence (feuille + poste ou compte). Ne remplis pas H et I.

RÈGLES : n'invente rien. Chaque conclusion s'appuie sur une citation exacte de « Annexe texte » ou sur des montants des autres feuilles. En cas de doute, statut « À vérifier ». Ne modifie aucune autre feuille. Termine par un résumé de 5 lignes au plus.`;
}

// Dossier complet pour une conversation Claude classique (sans Claude for Excel) : la réponse attendue est un bloc JSON
export function dossierConversation(meta, lines, comptes, balance, checklist) {
  const tab = (rows) => rows.map(r => r.join(" | ")).join("\n");
  return `Tu es auditeur financier. Analyse l'annexe des comptes annuels de ${meta.societe || "la société"} (exercice clos le ${meta.cloture || "?"}). Catégorie retenue : ${meta.cat.lib} → ${meta.cat.annexe} ; les points de niveau « Complète » sont facultatifs pour une petite entreprise (« Non applicable » s'ils sont absents).

1) Pour chaque point de la LISTE DE CONTRÔLE, donne : statut (${STATUTS.join(", ")}), justification (1 ou 2 phrases, avec montants et source), citation exacte (recopiée MOT POUR MOT depuis le TEXTE DE L'ANNEXE ci-dessous, 300 caractères au plus, « … » si tu coupes ; vide si Absent) et page.
2) Liste les INCOHÉRENCES entre le texte de l'annexe et les comptes ou la balance, ou entre deux passages de l'annexe (montant cité différent, compte cité erroné, provision non expliquée, date incohérente, tableau annoncé absent, texte repris d'un exercice antérieur…).
N'invente rien ; en cas de doute, statut « À vérifier ».

Réponds UNIQUEMENT par un bloc JSON de cette forme :
{"controles":[{"id":"M01","statut":"Conforme","justification":"…","citation":"…","page":10}],
 "incoherences":[{"constat":"…","citation":"…","page":15,"montant_cite":13107,"montant_reference":173484.40,"source":"Bilan actif – dépréciation clients / compte 491600"}]}

=== LISTE DE CONTRÔLE (id | thème | point | niveau | applicable) ===
${tab(checklist.map(c => [c[0], c[1], c[2], c[3], c[4]]))}

=== TEXTE DE L'ANNEXE ([page] texte) ===
${lines.map(l => `[${l.page}] ${l.text}`).join("\n")}

=== COMPTES ANNUELS (poste | colonne | plaquette | balance) ===
${tab(comptes)}

=== BALANCE (compte | libellé | solde de clôture, débit positif) ===
${tab(balance.filter(b => Math.abs(b.solde) > 0.004).map(b => [b.compte, b.lib, b.solde.toFixed(2)]))}
`;
}

// Lecture de la réponse JSON d'une conversation Claude (tolère le texte autour et les blocs ```json)
export function parseReponse(txt) {
  const s = String(txt || "");
  const m = /```(?:json)?\s*([\s\S]*?)```/.exec(s); let body = m ? m[1] : s;
  const a = body.indexOf("{"), b = body.lastIndexOf("}"); if (a < 0 || b < a) throw new Error("aucun bloc JSON trouvé dans la réponse");
  const j = JSON.parse(body.slice(a, b + 1));
  const controles = (j.controles || []).map(c => ({ id: String(c.id || "").trim(), statut: String(c.statut || "").trim(), justification: String(c.justification || ""), citation: String(c.citation || ""), page: Number(c.page) || "" })).filter(c => c.id);
  const incoherences = (j.incoherences || []).map(c => ({ constat: String(c.constat || ""), citation: String(c.citation || ""), page: Number(c.page) || "", montant_cite: c.montant_cite ?? "", montant_reference: c.montant_reference ?? "", source: String(c.source || "") })).filter(c => c.constat);
  return { controles, incoherences };
}
export { nz as normCitation };
