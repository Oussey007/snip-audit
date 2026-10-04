// Plan de regroupement par défaut : comptes PCG -> postes des comptes annuels (modèle en liste, règlement ANC 2014-03).
// Chaque poste est modifiable ensuite dans la feuille « Mapping » du classeur.
//
// Syntaxe d'une règle de comptes (séparateur : espace ou point-virgule) :
//   208       comptes commençant par 208
//   40C       comptes commençant par 40 dont le solde est créditeur (D = débiteur)
//   -409      exclut les comptes commençant par 409
//   RES       résultat de l'exercice (classes 6 et 7)
// Signe : actif et charges = solde débiteur positif ; passif et produits = solde créditeur positif.

const P = (section, id, libelle, regle, opts = {}) => ({ section, id, libelle, regle, ...opts });
const T = (section, id, libelle, formule, opts = {}) => ({ section, id, libelle, formule, total: true, ...opts });

export const DEFAULT_POSTES = [
  // ---------------- ACTIF (colonnes : brut, amortissements et dépréciations, net)
  P("actif", "A_CSNA", "Capital souscrit non appelé", "109"),
  P("actif", "A_FE", "Frais d'établissement", "201", { amort: "2801" }),
  P("actif", "A_RD", "Frais de recherche et de développement", "203", { amort: "2803", alias: ["Frais de développement"] }),
  P("actif", "A_CBL", "Concessions, brevets et droits similaires", "205", { amort: "2805 2905", alias: ["Concessions, brvts, licences, logiciels, drts & val.similaires", "Concessions, brevets, licences, logiciels, droits et valeurs similaires"] }),
  P("actif", "A_FC", "Fonds commercial", "206 207", { amort: "2807 2906 2907" }),
  P("actif", "A_AII", "Autres immobilisations incorporelles", "208 232", { amort: "2808 2908 2932" }),
  P("actif", "A_AAII", "Avances et acomptes sur immobilisations incorporelles", "237", { alias: ["Immobilisations incorporelles en cours, avances et acomptes"] }),
  P("actif", "A_TER", "Terrains", "211 212", { amort: "2811 2812 2911" }),
  P("actif", "A_CONS", "Constructions", "213 214", { amort: "2813 2814 2913 2914" }),
  P("actif", "A_ITMOI", "Installations techniques, matériel et outillage industriels", "215", { amort: "2815 2915" }),
  P("actif", "A_AIC", "Autres immobilisations corporelles", "218 22 -229", { amort: "2818 282 2918" }),
  P("actif", "A_ICEC", "Immobilisations corporelles en cours", "231", { amort: "2931" }),
  P("actif", "A_AAIC", "Avances et acomptes", "238", { alias: ["Avances et acomptes sur immobilisations corporelles"] }),
  P("actif", "A_PME", "Participations (mise en équivalence)", "", { alias: ["Participations évaluées selon la méthode de mise en équivalence"] }),
  P("actif", "A_AP", "Autres participations", "261 266", { amort: "2961 2966", alias: ["Participations"] }),
  P("actif", "A_CRP", "Créances rattachées aux participations", "267 268", { amort: "2967 2968", alias: ["Créances rattachées à des participations"] }),
  P("actif", "A_ATI", "Autres titres immobilisés", "271 272 273", { amort: "2971 2972 2973" }),
  P("actif", "A_PRET", "Prêts", "274", { amort: "2974" }),
  P("actif", "A_AIF", "Autres immobilisations financières", "275 276", { amort: "2975 2976" }),
  T("actif", "A_TAI", "Total actif immobilisé", "A_FE+A_RD+A_CBL+A_FC+A_AII+A_AAII+A_TER+A_CONS+A_ITMOI+A_AIC+A_ICEC+A_AAIC+A_PME+A_AP+A_CRP+A_ATI+A_PRET+A_AIF", { alias: ["TOTAL (I)", "Total I"] }),
  P("actif", "A_MP", "Matières premières et autres approvisionnements", "31 32", { amort: "391 392", alias: ["Matières premières, approvisionnements"] }),
  P("actif", "A_ENC", "En-cours de production (biens et services)", "33 34", { amort: "393 394", alias: ["En-cours de production de biens", "En cours de production de biens et services"] }),
  P("actif", "A_PIF", "Produits intermédiaires et finis", "35", { amort: "395" }),
  P("actif", "A_MARCH", "Marchandises", "37", { amort: "397" }),
  P("actif", "A_AAV", "Avances et acomptes versés sur commandes", "4091D"),
  P("actif", "A_CLI", "Clients et comptes rattachés", "411D 413D 416D 417D 418D", { amort: "491" }),
  P("actif", "A_AC", "Autres créances", "40D -4091 419D 42D 43D 44D 45D -4562 46D 47D -476 -477 48D -481 -486 -487 -488", { amort: "495 496" }),
  P("actif", "A_CSANV", "Capital souscrit et appelé, non versé", "4562"),
  P("actif", "A_VMP", "Valeurs mobilières de placement", "50 -509", { amort: "590" }),
  P("actif", "A_DISP", "Disponibilités", "51D 53 54 58D"),
  P("actif", "A_CCA", "Charges constatées d'avance", "486"),
  T("actif", "A_TAC", "Total actif circulant", "A_MP+A_ENC+A_PIF+A_MARCH+A_AAV+A_CLI+A_AC+A_CSANV+A_VMP+A_DISP+A_CCA", { alias: ["TOTAL (II)", "Total II"] }),
  P("actif", "A_FEE", "Frais d'émission d'emprunt à étaler", "4816", { alias: ["Frais d'émission d'emprunt"] }),
  P("actif", "A_PRO", "Primes de remboursement des obligations", "169", { alias: ["Primes de remboursement des emprunts"] }),
  P("actif", "A_ECA", "Ecarts de conversion actif", "476", { alias: ["Écarts de conversion actif et différences d'évaluation"] }),
  T("actif", "A_TG", "Total général actif", "A_CSNA+A_TAI+A_TAC+A_FEE+A_PRO+A_ECA", { alias: ["TOTAL GENERAL", "Total général (I à VI)", "TOTAL ACTIF"] }),

  // ---------------- PASSIF
  P("passif", "P_CAP", "Capital", "101 108", { alias: ["Capital social ou individuel", "Capital individuel"] }),
  P("passif", "P_PRIM", "Primes d'émission, de fusion, d'apport, ...", "104", { alias: ["Primes d'émission, de fusion, d'apport"] }),
  P("passif", "P_ECR", "Ecart de réévaluation", "105", { alias: ["Ecarts de réévaluation"] }),
  P("passif", "P_RL", "Réserve légale", "1061"),
  P("passif", "P_RS", "Réserves statutaires ou contractuelles", "1063"),
  P("passif", "P_RR", "Réserves réglementées", "1062 1064"),
  P("passif", "P_AR", "Autres réserves", "1068"),
  P("passif", "P_RAN", "Report à nouveau", "11"),
  P("passif", "P_RES", "Résultat de l'exercice (bénéfice ou perte)", "RES 12", { alias: ["RESULTAT DE L'EXERCICE", "Résultat de l'exercice"] }),
  P("passif", "P_SUBV", "Subventions d'investissement", "13"),
  P("passif", "P_PROVR", "Provisions réglementées", "14"),
  T("passif", "P_TCP", "Total capitaux propres", "P_CAP+P_PRIM+P_ECR+P_RL+P_RS+P_RR+P_AR+P_RAN+P_RES+P_SUBV+P_PROVR", { alias: ["TOTAL (I)"] }),
  P("passif", "P_TP", "Produits des émissions de titres participatifs", "1671", { alias: ["Produit des émissions de titres participatifs"] }),
  P("passif", "P_AVC", "Avances conditionnées", "1674"),
  P("passif", "P_DCON", "Droits du concédant", "229", { repli: "P_AVC" }),
  T("passif", "P_TAFP", "Total autres fonds propres", "P_TP+P_AVC+P_DCON"),
  P("passif", "P_PRR", "Provisions pour risques", "151"),
  P("passif", "P_PRC", "Provisions pour charges", "153 154 155 156 157 158"),
  T("passif", "P_TPRC", "Total provisions pour risques et charges", "P_PRR+P_PRC"),
  P("passif", "P_EOC", "Emprunts obligataires convertibles", "161"),
  P("passif", "P_AEO", "Autres emprunts obligataires", "163"),
  P("passif", "P_EEC", "Emprunts et dettes auprès des établissements de crédit", "164 16884 512C 514C 517C 5186C 519"),
  P("passif", "P_EDFD", "Emprunts et dettes financières diverses", "165 166 1675 1678 168 -16884 17 455C", { alias: ["Emprunts et dettes financières divers"] }),
  P("passif", "P_AAR", "Avances et acomptes reçus sur commandes en cours", "4191C"),
  P("passif", "P_FOUR", "Dettes fournisseurs et comptes rattachés", "40C -404 -405 -4084 -409"),
  P("passif", "P_DFS", "Dettes fiscales et sociales", "42C 43C 44C"),
  P("passif", "P_DIMMO", "Dettes sur immobilisations et comptes rattachés", "404C 405C 4084C"),
  P("passif", "P_AD", "Autres dettes", "411C 413C 416C 417C 418C 4196C 4197C 4198C 45C -455 -4562 46C 47C -476 -477 48C -481 -486 -487 509C"),
  P("passif", "P_PCA", "Produits constatés d'avance", "487"),
  T("passif", "P_TD", "Total dettes", "P_EOC+P_AEO+P_EEC+P_EDFD+P_AAR+P_FOUR+P_DFS+P_DIMMO+P_AD+P_PCA", { alias: ["TOTAL (IV)"] }),
  P("passif", "P_ECP", "Ecarts de conversion passif", "477", { alias: ["Écart de conversion passif et différences d'évaluation"] }),
  T("passif", "P_TG", "Total général passif", "P_TCP+P_TAFP+P_TPRC+P_TD+P_ECP", { alias: ["TOTAL GENERAL", "Total général (I à V)", "TOTAL PASSIF"] }),

  // ---------------- COMPTE DE RÉSULTAT (en liste)
  P("cr", "R_VM", "Ventes de marchandises", "707 7097", { sens: -1 }),
  P("cr", "R_PVB", "Production vendue (biens)", "701 702 703 7091 7092 7093", { sens: -1, alias: ["Production vendue biens"] }),
  P("cr", "R_PVS", "Production vendue (services)", "704 705 706 708 7094 7095 7096 7098", { sens: -1, alias: ["Production vendue services"] }),
  T("cr", "R_CA", "Chiffre d'affaires net", "R_VM+R_PVB+R_PVS", { alias: ["Chiffres d'affaires nets", "Montant net du chiffre d'affaires"] }),
  P("cr", "R_PS", "Production stockée", "713", { sens: -1 }),
  P("cr", "R_PI", "Production immobilisée", "72", { sens: -1 }),
  P("cr", "R_SUB", "Subventions d'exploitation", "74", { sens: -1 }),
  P("cr", "R_REP", "Reprises sur provisions (et amortissements), transferts de charges", "781 791", { sens: -1, alias: ["Reprises sur amortissements et provisions, transferts de charges", "Reprises sur dépréciations, provisions (et amortissements), transferts de charges"] }),
  P("cr", "R_AP", "Autres produits", "75 -755", { sens: -1 }),
  T("cr", "R_T1", "Total des produits d'exploitation (I)", "R_CA+R_PS+R_PI+R_SUB+R_REP+R_AP", { alias: ["Total I", "Total des produits d'exploitation"] }),
  P("cr", "R_AM", "Achats de marchandises", "607 6087 6097", { alias: ["Achats de marchandises (y compris droits de douane)"] }),
  P("cr", "R_VSM", "Variations de stock (marchandises)", "6037", { alias: ["Variations de stock", "Variation de stock (marchandises)"] }),
  P("cr", "R_AMP", "Achats de matières premières et autres approvisionnements", "601 602 6081 6082 6091 6092"),
  P("cr", "R_VSMP", "Variations de stock (matières premières)", "6031 6032", { alias: ["Variations de stock", "Variation de stock (matières premières et approvisionnements)"] }),
  P("cr", "R_AACE", "Autres achats et charges externes", "604 605 606 608 -6081 -6082 -6087 609 -6091 -6092 -6097 61 62"),
  P("cr", "R_IMP", "Impôts, taxes et versements assimilés", "63"),
  P("cr", "R_SAL", "Salaires et traitements", "641 644"),
  P("cr", "R_CS", "Charges sociales", "645 646 647 648", { alias: ["Cotisations sociales"] }),
  P("cr", "R_DAI", "Dotations aux amortissements sur immobilisations", "6811 6812", { alias: ["Sur immobilisations : dotations aux amortissements", "Dotations aux amortissements"] }),
  P("cr", "R_DDI", "Dotations aux dépréciations sur immobilisations", "6816", { alias: ["Sur immobilisations : dotations aux dépréciations"] }),
  P("cr", "R_DDAC", "Dotations aux dépréciations sur actif circulant", "6817", { alias: ["Sur actif circulant : dotations aux dépréciations"] }),
  P("cr", "R_DPRC", "Dotations aux provisions pour risques et charges", "6815", { alias: ["Pour risques et charges : dotations aux provisions"] }),
  P("cr", "R_AC", "Autres charges", "65 -655"),
  T("cr", "R_T2", "Total des charges d'exploitation (II)", "R_AM+R_VSM+R_AMP+R_VSMP+R_AACE+R_IMP+R_SAL+R_CS+R_DAI+R_DDI+R_DDAC+R_DPRC+R_AC", { alias: ["Total II", "Total des charges d'exploitation"] }),
  T("cr", "R_REX", "Résultat d'exploitation (I-II)", "R_T1-R_T2", { alias: ["RESULTAT D'EXPLOITATION", "1 - RESULTAT D'EXPLOITATION (I - II)"] }),
  P("cr", "R_QPB", "Bénéfice attribué ou perte transférée", "755", { sens: -1 }),
  P("cr", "R_QPP", "Perte supportée ou bénéfice transféré", "655"),
  P("cr", "R_PFP", "Produits financiers de participations", "761", { sens: -1, alias: ["De participation"] }),
  P("cr", "R_PFVM", "Produits des autres valeurs mobilières et créances de l'actif immobilisé", "762", { sens: -1, alias: ["D'autres valeurs mobilières et créances de l'actif immobilisé"] }),
  P("cr", "R_PFI", "Autres intérêts et produits assimilés", "763 764 765 768", { sens: -1 }),
  P("cr", "R_RPF", "Reprises sur provisions et dépréciations financières et transferts de charges", "786 796", { sens: -1, alias: ["Reprises sur provisions et dépréciations et transferts de charges", "Reprises sur dépréciations et provisions, transferts de charges"] }),
  P("cr", "R_DPC", "Différences positives de change", "766", { sens: -1 }),
  P("cr", "R_PNC", "Produits nets sur cessions de valeurs mobilières de placement", "767", { sens: -1, alias: ["Produits nets sur cessions de valeurs mobilièrers de placement"] }),
  T("cr", "R_T5", "Total des produits financiers (V)", "R_PFP+R_PFVM+R_PFI+R_RPF+R_DPC+R_PNC", { alias: ["Total V"] }),
  P("cr", "R_DAF", "Dotations financières aux amortissements, dépréciations et provisions", "686", { alias: ["Dotations aux amortissements, aux dépréciations et aux provisions"] }),
  P("cr", "R_INT", "Intérêts et charges assimilées", "661 664 665 668"),
  P("cr", "R_DNC", "Différences négatives de change", "666"),
  P("cr", "R_CNC", "Charges nettes sur cessions de valeurs mobilières de placement", "667"),
  T("cr", "R_T6", "Total des charges financières (VI)", "R_DAF+R_INT+R_DNC+R_CNC", { alias: ["Total VI"] }),
  T("cr", "R_RFI", "Résultat financier (V-VI)", "R_T5-R_T6", { alias: ["RESULTAT FINANCIER", "2 - RÉSULTAT FINANCIER (V - VI)"] }),
  T("cr", "R_RCAI", "Résultat courant avant impôts", "R_REX+R_QPB-R_QPP+R_RFI", { alias: ["RESULTAT COURANT avant impôts", "3 - RÉSULTAT COURANT AVANT IMPÔTS"] }),
  P("cr", "R_PEG", "Produits exceptionnels sur opérations de gestion", "771", { sens: -1, alias: ["Sur opérations de gestion"] }),
  P("cr", "R_PEC", "Produits exceptionnels sur opérations en capital", "775 777 778", { sens: -1, alias: ["Sur opérations en capital"] }),
  P("cr", "R_RPE", "Reprises exceptionnelles sur provisions et dépréciations, transferts de charges", "787 797", { sens: -1, alias: ["Reprises sur provisions et dépréciation et transferts de charges", "Reprises sur provisions et dépréciations et transferts de charges", "Reprises sur dépréciations et provisions, transferts de charges"] }),
  T("cr", "R_T7", "Total des produits exceptionnels (VII)", "R_PEG+R_PEC+R_RPE", { alias: ["Total produits exceptionnels"] }),
  P("cr", "R_CEG", "Charges exceptionnelles sur opérations de gestion", "671", { alias: ["Sur opérations de gestion"] }),
  P("cr", "R_CEC", "Charges exceptionnelles sur opérations en capital", "675 678", { alias: ["Sur opérations en capital"] }),
  P("cr", "R_DAE", "Dotations exceptionnelles aux amortissements, dépréciations et provisions", "687", { alias: ["Dotations aux amortissements, aux dépréciations et aux provisions"] }),
  T("cr", "R_T8", "Total des charges exceptionnelles (VIII)", "R_CEG+R_CEC+R_DAE", { alias: ["Total charges exceptionnelles"] }),
  T("cr", "R_REXC", "Résultat exceptionnel (VII-VIII)", "R_T7-R_T8", { alias: ["RESULTAT EXCEPTIONNEL", "4 - RÉSULTAT EXCEPTIONNEL (VII - VIII)"] }),
  P("cr", "R_PART", "Participation des salariés aux résultats (IX)", "691", { alias: ["Participation des salariés aux résultats", "Participation des salariés aux résultats de l'entreprise"] }),
  P("cr", "R_IS", "Impôts sur les bénéfices (X)", "695 696 697 698 699", { alias: ["Impôts sur les bénéfices"] }),
  T("cr", "R_TP", "Total des produits (I+III+V+VII)", "R_T1+R_QPB+R_T5+R_T7", { alias: ["Total des produits", "TOTAL DES PRODUITS"] }),
  T("cr", "R_TC", "Total des charges (II+IV+VI+VIII+IX+X)", "R_T2+R_QPP+R_T6+R_T8+R_PART+R_IS", { alias: ["Total des charges", "TOTAL DES CHARGES"] }),
  T("cr", "R_BEN", "Bénéfice ou perte", "R_TP-R_TC", { alias: ["BENEFICE OU PERTE", "5 - BÉNÉFICE OU PERTE", "Bénéfice ou perte (total des produits - total des charges)"] })
];

// ---------- Règles de comptes
export function parseRegle(txt) {
  const inc = [], exc = []; let res = false;
  for (const tok of String(txt || "").split(/[\s;,]+/).filter(Boolean)) {
    if (/^RES$/i.test(tok)) { res = true; continue; }
    const m = /^(-?)(\d+)([DC]?)$/i.exec(tok);
    if (!m) throw new Error(`règle de comptes illisible : « ${tok} »`);
    (m[1] ? exc : inc).push({ p: m[2], f: m[3].toUpperCase() });
  }
  return { inc, exc, res };
}
// Comptes de la balance retenus par une règle
export function comptesDe(regle, balance) {
  const r = typeof regle === "string" ? parseRegle(regle) : regle;
  const out = [];
  for (const b of balance) {
    const c = b.compte;
    if (r.res && (c[0] === "6" || c[0] === "7")) { out.push(b); continue; }
    if (r.exc.some(e => c.startsWith(e.p))) continue;
    if (r.inc.some(i => c.startsWith(i.p) && (!i.f || (i.f === "D" ? b.solde > 0 : b.solde < 0)))) out.push(b);
  }
  return out;
}

export const SECTION_SIGN = { actif: 1, passif: -1 };
export const SECTION_LABEL = { actif: "Bilan actif", passif: "Bilan passif", cr: "Compte de résultat" };

// Contrôle d'affectation : chaque compte de bilan (classes 1 à 5) et de gestion (6, 7) doit être affecté une fois et une seule
export function controleAffectation(postes, balance) {
  const vu = new Map();
  const note = (b, p, fam) => { const k = fam + "|" + b.compte; if (!vu.has(k)) vu.set(k, { b, fam, postes: [] }); vu.get(k).postes.push(p.libelle); };
  for (const p of postes) {
    if (p.total) continue;
    const fam = p.section === "cr" ? "cr" : "bilan";
    for (const reg of [p.regle, p.amort]) {
      if (!reg) continue;
      const r = parseRegle(reg); const rr = { ...r, res: false };
      for (const b of comptesDe(rr, balance)) note(b, p, fam);
    }
  }
  const nonAffectes = [], doublons = [];
  for (const b of balance) {
    if (Math.abs(b.solde) < 0.005) continue;
    const cl = b.compte[0];
    const fams = cl >= "1" && cl <= "5" ? ["bilan"] : (cl === "6" || cl === "7") ? ["cr"] : [];
    for (const fam of fams) {
      const v = vu.get(fam + "|" + b.compte);
      if (!v) nonAffectes.push({ ...b, famille: fam });
      else if (v.postes.length > 1) doublons.push({ ...b, famille: fam, postes: v.postes });
    }
  }
  return { nonAffectes, doublons };
}

// Normalisation des libellés pour la reconnaissance dans la plaquette
export function normLabel(s) {
  return String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase()
    .replace(/\((\d{1,2}|[a-z]|\*)\)/g, " ").replace(/\*/g, " ")
    .replace(/^[\s\-–•]+/, "").replace(/[^a-z0-9]+/g, " ").trim();
}
