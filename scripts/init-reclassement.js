/**
 * init-reclassement.js
 * Génère hospifinance-it-data/data/reclassement.json depuis les fichiers Excel de nomenclature DSI.
 *
 * Usage :
 *   node scripts/init-reclassement.js --ref <Referentiel_v6.xlsx> --plan <Plan_v5.xlsx> [--out <path>] [--dry]
 *
 * Arguments :
 *   --ref   Chemin vers Referentiel_Fournisseurs_DSI_v6.xlsx  (requis ou par défaut)
 *   --plan  Chemin vers Plan_reclassement_DSI_COMPLET_v5.xlsx (requis ou par défaut)
 *   --out   Fichier de sortie (défaut : ./data/reclassement.json)
 *   --dry   Dry-run : affiche les stats sans écrire le fichier
 */

'use strict';
const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

// ── Parsing des arguments ────────────────────────────────────────────────────

const argv = process.argv.slice(2);
const getArg  = (name) => { const i = argv.indexOf(`--${name}`); return i !== -1 && argv[i + 1] ? argv[i + 1] : null; };
const hasFlag = (name) => argv.includes(`--${name}`);

const REF_PATH  = getArg('ref')  ?? path.join(__dirname, '..', '..', '..', 'Résultats', 'Referentiel_Fournisseurs_DSI_v6.xlsx');
const PLAN_PATH = getArg('plan') ?? path.join(__dirname, '..', '..', '..', 'Résultats', 'Plan_reclassement_DSI_COMPLET_v5.xlsx');
const OUT_PATH  = getArg('out')  ?? path.join(__dirname, '..', 'data', 'reclassement.json');
const DRY_RUN   = hasFlag('dry');

// ── Helpers ──────────────────────────────────────────────────────────────────

const toStr   = (v) => (v == null || v === '-') ? '' : String(v).trim();
const toFloat = (v) => {
  if (v == null || v === '-' || v === '') return 0;
  return parseFloat(String(v).replace(/\s/g, '').replace(',', '.')) || 0;
};
const toInt   = (v) => Math.round(toFloat(v));
const toBool  = (v) => ['oui', 'true', '1', 'corrigé'].includes(toStr(v).toLowerCase());

/** Normalise un compte : '61526100' ou 'H61526100' → 'H61526100' */
const normaliseCompte = (c) => {
  const s = toStr(c).replace(/^H/i, '');
  return s && /^\d+/.test(s) ? `H${s}` : '';
};

/** Ordre canonique des familles */
const FAMILLE_ORDER = [
  'Infrastructures',
  'Applications',
  'Support et services utilisateurs',
  'Cybersécurité',
  'Data et pilotage',
  'Prestations externes récurrentes',
  'Hors périmètre DSI',
];

/**
 * Mapping abréviation → nom complet de famille
 * Couvre les formats courts utilisés dans le "Récapitulatif des corrections"
 */
const ABBREV_TO_FAMILLE = {
  'infra':                    'Infrastructures',
  'infrastructures':          'Infrastructures',
  'appli':                    'Applications',
  'applications':             'Applications',
  'ssu':                      'Support et services utilisateurs',
  'support et services':      'Support et services utilisateurs',
  'support':                  'Support et services utilisateurs',
  'cyber':                    'Cybersécurité',
  'cybersecurite':            'Cybersécurité',
  'data':                     'Data et pilotage',
  'data et pilotage':         'Data et pilotage',
  'prestations':              'Prestations externes récurrentes',
  'hors perimetre dsi':       'Hors périmètre DSI',
  'hors perimetre':           'Hors périmètre DSI',
  'hors':                     'Hors périmètre DSI',
};

/**
 * Parse une chaîne "Abbrev > sous-catégorie" → { familleN1, sousCatN2 }
 * Utilisé pour les lignes "Récapitulatif des corrections" du référentiel.
 */
function parseCategorie(str) {
  const parts = str.split('>').map(s => s.trim());
  const raw   = (parts[0] || '').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '') // retire accents pour la clé
    .replace(/[^a-z\s]/g, '').trim();
  const famille  = ABBREV_TO_FAMILLE[raw] || parts[0].trim();
  const sousCat  = parts[1] || '';
  return { familleN1: famille, sousCatN2: sousCat };
}

/** Nettoie un statut : supprime emoji et whitespace résiduel */
const normaliseStatut = (raw) => {
  const s = toStr(raw)
    .replace(/[\u{1F300}-\u{1FFFF}\u{2600}-\u{27BF}✅✓⚠❌🔴🟡🟢]/gu, '')
    .trim();
  if (!s || s.toLowerCase() === 'ok') return 'OK';
  if (/v.rifier/i.test(s))           return 'A vérifier';
  if (/recat/i.test(s))              return 'Recatégorisé';
  if (/note/i.test(s))               return 'Note';
  if (/corriger/i.test(s))           return 'A corriger';
  return s;
};

/**
 * Parse une chaîne de conditions pipe-séparées en tableau d'objets.
 * Ex: "COMPTE=61526100 | DESIGNATION CONTIENT=support" →
 *     [{champ:'COMPTE', operateur:'=', valeur:'61526100'}, ...]
 */
const parseConditions = (regleStr) => {
  if (!regleStr) return [];
  return regleStr.split('|').map(part => {
    part = part.trim();
    if (!part || part.toUpperCase() === 'DEFAUT') {
      return { champ: 'DEFAUT', operateur: '=', valeur: 'DEFAUT' };
    }
    const contientMatch = part.match(/^([A-Z_]+)\s+CONTIENT=(.+)$/i);
    if (contientMatch) {
      return { champ: contientMatch[1].trim().toUpperCase(), operateur: 'CONTIENT', valeur: contientMatch[2].trim() };
    }
    const eqMatch = part.match(/^([A-Z_]+)=(.+)$/i);
    if (eqMatch) {
      return { champ: eqMatch[1].trim().toUpperCase(), operateur: '=', valeur: eqMatch[2].trim() };
    }
    return { champ: 'DESIGNATION', operateur: 'CONTIENT', valeur: part };
  }).filter(c => c.champ === 'DEFAUT' || c.valeur);
};

// ── Parseurs par onglet ──────────────────────────────────────────────────────

/**
 * 5. Nomenclature → nomenclature[]
 * Colonnes : [0]# | [1]Famille N1 | [2]Sous-catégorie N2 | [3]Périmètre | [4]Description
 */
function parseNomenclature(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const groupes = new Map();

  for (const row of rows) {
    const num = Number(row?.[0]);
    if (!Number.isInteger(num) || num <= 0) continue;

    const famille    = toStr(row[1]);
    const sousCat    = toStr(row[2]);
    const perimetre  = toStr(row[3]);
    const description = toStr(row[4]);

    if (!famille || !sousCat) continue;
    if (!groupes.has(famille)) groupes.set(famille, []);
    groupes.get(famille).push({ label: sousCat, perimetre, description });
  }

  return FAMILLE_ORDER
    .filter(f => groupes.has(f))
    .map(f => ({ famille: f, sousCategoriesDisponibles: groupes.get(f) }));
}

/**
 * 1. Fournisseurs → referentielFournisseurs[]
 * Colonnes : [0]# | [1]Fournisseur | [2]Famille N1 | [3]Sous-cat N2
 *            [4]Justification | [5]Vol.2025 | [6]Vol.2026 | [7]Vol.total | [8]Statut | [9]Note
 */
function parseFournisseurs(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const result = [];

  for (const row of rows) {
    const num = Number(row?.[0]);
    if (!Number.isInteger(num) || num <= 0) continue;

    const nomFourn = toStr(row[1]);
    if (nomFourn.toUpperCase() === 'TOTAL' || !nomFourn) continue;

    let familleN1 = toStr(row[2]);
    let sousCatN2 = toStr(row[3]);

    if (familleN1.includes('>')) {
      // Section "Récapitulatif des corrections" : col[2]=ancienne cat, col[3]=nouvelle cat
      const corrected = parseCategorie(toStr(row[3]));
      familleN1 = corrected.familleN1;
      sousCatN2 = corrected.sousCatN2;
    }

    result.push({
      id:            num,
      fournisseur:   nomFourn,
      familleN1,
      sousCatN2,
      justification: toStr(row[4]),
      vol2025:       toFloat(row[5]),
      vol2026:       toFloat(row[6]),
      volTotal:      toFloat(row[7]),
      statut:        normaliseStatut(toStr(row[8])),
      note:          toStr(row[9]),
    });
  }

  return result;
}

/**
 * 2. Multi-nature → reglesMultiNature[]
 * Colonnes : [0]# | [1]Fournisseur | [2]Famille N1 | [3]Sous-cat N2
 *            [4]Règle prioritaire | [5]Comptes observés | [6]Famille alt | [7]Sous-cat alt
 *            [8]Commentaire | [9]Validé ?
 */
function parseMultiNature(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const result = [];

  for (const row of rows) {
    const num = Number(row?.[0]);
    if (!Number.isInteger(num) || num <= 0) continue;

    result.push({
      id:          num,
      priorite:    num,
      fournisseur: toStr(row[1]),
      familleN1:   toStr(row[2]),
      sousCatN2:   toStr(row[3]),
      conditions:  parseConditions(toStr(row[4])),
      comptesObs:  toStr(row[5]),
      familleAlt:  toStr(row[6]),
      sousCatAlt:  toStr(row[7]),
      commentaire: toStr(row[8]),
      valide:      toBool(row[9]),
    });
  }

  return result;
}

/**
 * 4. Mots-clés → reglesMosCles[]
 * Colonnes : [0]# | [1]Mot-clé | [2]Famille N1 | [3]Sous-cat N2
 */
function parseMosCles(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const result = [];

  for (const row of rows) {
    const num = Number(row?.[0]);
    if (!Number.isInteger(num) || num <= 0) continue;

    const motCle = toStr(row[1]).toUpperCase();
    if (!motCle) continue;

    result.push({
      id:        num,
      priorite:  num,
      motCle,
      familleN1: toStr(row[2]),
      sousCatN2: toStr(row[3]),
    });
  }

  return result;
}

/**
 * 3. Mapping comptes (v6) → tableau de base
 * Colonnes : [0]# | [1]Compte | [2]Libellé | [3]Famille N1 | [4]Sous-cat N2 | [5]Type | [6]Hétérogène
 */
function parseMappingV6(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const result = [];

  for (const row of rows) {
    const num = Number(row?.[0]);
    if (!Number.isInteger(num) || num <= 0) continue;

    const compte = normaliseCompte(toStr(row[1]));
    if (!compte) continue;

    result.push({
      compte,
      libelleCompte: toStr(row[2]),
      familleDefaut: toStr(row[3]),
      sousCatDefaut: toStr(row[4]),
      type:          toStr(row[5]),
      heterogene:    toBool(row[6]),
    });
  }

  return result;
}

/**
 * Plan v5 — onglet "4. Mapping compte" → Map<compteNaked, {totalObserve, nbCommandes, commentaire}>
 * Colonnes : [0]label | [1]Compte | [2]Libellé | [3]Famille | [4]Sous-cat | [5]Type
 *            [6]Hétérogène | [7]Total observé | [8]Nb cmds | [9]Commentaire
 * On s'arrête dès une ligne qui contient 'CAPEX' ou 'CLASSE 2' dans n'importe quelle colonne.
 */
function parseMappingV5(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const result = new Map();

  for (const row of rows) {
    if (!row) continue;
    const rowStr = row.map(c => toStr(c)).join('|').toUpperCase();
    // Arrêt sur la ligne CAPEX
    if (rowStr.includes('CAPEX') || rowStr.includes('CLASSE 2') || rowStr.includes('CLASSE2')) break;

    // La colonne compte peut être en [1] (index 1-based du spec) mais avec header:1 c'est déjà 0-based
    const compteRaw = toStr(row[1]);
    const compte    = compteRaw.replace(/^H/i, '').trim();
    if (!compte || !/^\d{6,}/.test(compte)) continue;

    result.set(compte, {
      totalObserve: toFloat(row[7]),
      nbCommandes:  toInt(row[8]),
      commentaire:  toStr(row[9]),
    });
  }

  return result;
}

/**
 * Jointure v6 + v5 : enrichit chaque ligne v6 avec les données volumétriques v5.
 */
function joinMapping(v6, v5Map) {
  return v6.map(entry => {
    const naked  = entry.compte.replace(/^H/i, '');
    const v5data = v5Map.get(naked) ?? { totalObserve: 0, nbCommandes: 0, commentaire: '' };
    return { ...entry, ...v5data };
  });
}

// ── Main ─────────────────────────────────────────────────────────────────────

function main() {
  // Vérification existence des fichiers
  if (!fs.existsSync(REF_PATH)) {
    console.error(`\n❌ Fichier référentiel introuvable :\n   ${REF_PATH}`);
    console.error('\nUsage : node scripts/init-reclassement.js --ref <path_v6.xlsx> --plan <path_v5.xlsx>\n');
    process.exit(1);
  }
  if (!fs.existsSync(PLAN_PATH)) {
    console.warn(`\n⚠️  Fichier plan v5 introuvable (enrichissement volumétrique désactivé) :\n   ${PLAN_PATH}`);
  }

  console.log('\n🔄 Lecture des fichiers Excel...');
  console.log(`   Référentiel v6 : ${path.basename(REF_PATH)}`);
  console.log(`   Plan v5        : ${path.basename(PLAN_PATH)}`);

  const rbRef  = XLSX.readFile(REF_PATH,  { cellDates: false, raw: false });
  const rbPlan = fs.existsSync(PLAN_PATH)
    ? XLSX.readFile(PLAN_PATH, { cellDates: false, raw: false })
    : null;

  // Vérification des onglets
  const required = ['5. Nomenclature', '1. Fournisseurs', '2. Multi-nature', '4. Mots-clés', '3. Mapping comptes'];
  for (const sheet of required) {
    if (!rbRef.Sheets[sheet]) {
      console.error(`❌ Onglet manquant dans v6 : "${sheet}"`);
      console.error(`   Onglets disponibles : ${rbRef.SheetNames.join(', ')}`);
      process.exit(1);
    }
  }

  console.log('\n📊 Parsing en cours...');

  const nomenclature            = parseNomenclature(rbRef.Sheets['5. Nomenclature']);
  const referentielFournisseurs = parseFournisseurs(rbRef.Sheets['1. Fournisseurs']);
  const reglesMultiNature       = parseMultiNature(rbRef.Sheets['2. Multi-nature']);
  const reglesMosCles           = parseMosCles(rbRef.Sheets['4. Mots-clés']);
  const mappingV6               = parseMappingV6(rbRef.Sheets['3. Mapping comptes']);
  const mappingV5               = rbPlan?.Sheets?.['4. Mapping compte']
    ? parseMappingV5(rbPlan.Sheets['4. Mapping compte'])
    : new Map();
  const mappingComptes          = joinMapping(mappingV6, mappingV5);

  const output = {
    version:   'v6',
    updatedAt: new Date().toISOString().split('T')[0],
    nomenclature,
    referentielFournisseurs,
    reglesMultiNature,
    reglesMosCles,
    mappingComptes,
  };

  // ── Rapport ────────────────────────────────────────────────────────────────
  const totalSousCats = nomenclature.reduce((s, f) => s + f.sousCategoriesDisponibles.length, 0);
  console.log('\n✅ Génération terminée :');
  console.log(`   nomenclature           : ${nomenclature.length} familles, ${totalSousCats} sous-catégories`);
  console.log(`   referentielFournisseurs: ${referentielFournisseurs.length} entrées`);
  console.log(`   reglesMultiNature      : ${reglesMultiNature.length} règles`);
  console.log(`   reglesMosCles          : ${reglesMosCles.length} règles`);
  console.log(`   mappingComptes         : ${mappingComptes.length} comptes`);
  console.log(`   mappingV5 enrichi      : ${mappingV5.size} comptes depuis plan v5`);

  // Warnings
  const aVerifier = referentielFournisseurs.filter(f => f.statut === 'A vérifier');
  const aCorrect  = referentielFournisseurs.filter(f => f.statut === 'A corriger' || f.familleN1 === 'A corriger');
  const noValide  = reglesMultiNature.filter(r => !r.valide);
  const sansConds = reglesMultiNature.filter(r => !r.conditions || r.conditions.length === 0);

  if (aVerifier.length) {
    console.warn(`\n⚠️  ${aVerifier.length} fournisseur(s) "À vérifier" :`);
    aVerifier.slice(0, 10).forEach(f => console.warn(`   #${f.id} — ${f.fournisseur}`));
    if (aVerifier.length > 10) console.warn(`   ... et ${aVerifier.length - 10} autres`);
  }
  if (aCorrect.length) {
    console.warn(`\n⚠️  ${aCorrect.length} fournisseur(s) avec famille malformée (format "A > B") :`);
    aCorrect.forEach(f => console.warn(`   #${f.id} — ${f.fournisseur}`));
  }
  if (noValide.length) {
    console.warn(`\nℹ️  ${noValide.length} règle(s) multi-nature non validées (valide: false)`);
  }
  if (sansConds.length) {
    console.warn(`\n⚠️  ${sansConds.length} règle(s) multi-nature sans conditions parsées :`);
    sansConds.forEach(r => console.warn(`   #${r.id} — ${r.fournisseur}`));
  }

  if (DRY_RUN) {
    console.log('\n🔍 Mode dry-run — fichier NON écrit.');
    console.log(JSON.stringify(output, null, 2).slice(0, 2000) + '\n...[tronqué]');
    return;
  }

  // Écriture
  fs.mkdirSync(path.dirname(OUT_PATH), { recursive: true });
  fs.writeFileSync(OUT_PATH, JSON.stringify(output, null, 2), 'utf8');
  console.log(`\n💾 Fichier écrit : ${OUT_PATH}`);
  console.log('   Rechargez l\'application ou redémarrez le serveur pour prendre en compte les changements.\n');
}

main();
