#!/usr/bin/env node
/**
 * Script d'import initial — EXTRACTION COMMANDES SAGE → JSON
 * Usage :
 *   node scripts/import-sage-2026.js --input "EXTRACTION COMMANDES 2021-2026.xlsx" --sheet 2026 --output ./data/
 *
 * Sorties :
 *   data/opex.json          Fournisseurs OPEX groupés par (Fournisseur + Compte ordonnateur)
 *   data/opex-orders.json   Toutes les lignes OPEX en commandes détaillées
 *   data/capex.json         Fournisseurs CAPEX groupés
 *   data/capex-orders.json  Toutes les lignes CAPEX
 *   data/import-log.json    Statistiques et erreurs
 */

import { readFileSync, writeFileSync, existsSync } from 'fs';
import { resolve, join } from 'path';
import { fileURLToPath } from 'url';
import * as XLSX from 'xlsx';

const __dirname = fileURLToPath(new URL('.', import.meta.url));

// ─── Arguments CLI ────────────────────────────────────────────────────────────
const args = process.argv.slice(2);
const getArg = (name) => {
  const idx = args.indexOf(name);
  return idx !== -1 ? args[idx + 1] : null;
};

const inputFile  = getArg('--input');
const sheetName  = getArg('--sheet') || '2026';
const outputDir  = getArg('--output') || join(__dirname, '..', 'data');

if (!inputFile) {
  console.error('Usage: node import-sage-2026.js --input <fichier.xlsx> [--sheet 2026] [--output ./data/]');
  process.exit(1);
}

const inputPath = resolve(inputFile);
if (!existsSync(inputPath)) {
  console.error(`Fichier introuvable : ${inputPath}`);
  process.exit(1);
}

// ─── Mapping compte → famille analytique ─────────────────────────────────────
const COMPTE_TO_FAMILLE = {
  'H60625100': 'Support et services utilisateurs',
  'H60625211': 'Hors périmètre DSI',
  'H61325100': 'Infrastructures',
  'H61525400': 'Applications',
  'H61526100': 'Applications',
  'H61841500': 'Prestations externes récurrentes',
  'H62281000': 'Prestations externes récurrentes',
  'I62281000': 'Prestations externes récurrentes',
  'H62610000': 'Infrastructures',
  'H62630000': 'Hors périmètre DSI',
  'H62631000': 'Infrastructures',
  'H62650000': 'Infrastructures',
  'H62882000': 'Hors périmètre DSI',
  'H65100000': 'Applications',
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
const parseSageDate = (val) => {
  if (!val || val === '-') return '';
  if (val instanceof Date) return val.toISOString().split('T')[0];
  if (typeof val === 'number') {
    // Numéro de série Excel
    const date = XLSX.SSF.parse_date_code(val);
    if (date) return `${date.y}-${String(date.m).padStart(2,'0')}-${String(date.d).padStart(2,'0')}`;
  }
  return '';
};

const num = (val) => {
  if (val === '-' || val === undefined || val === null || val === '') return 0;
  return Number(val) || 0;
};

const mapSageStatus = (etat, montantEngageNonRecu) => {
  if (String(etat).trim() === 'Soldée') return 'Facturée';
  return num(montantEngageNonRecu) > 0 ? 'Commandée' : 'Livrée';
};

const getLineType = (compte) => {
  if (!compte) return null;
  const c = String(compte).toUpperCase().trim();
  if (c.startsWith('H6') || c.startsWith('I6')) return 'OPEX';
  if (c.startsWith('H2')) return 'CAPEX';
  return null;
};

const genId = () => `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

// ─── Lecture du fichier XLSX ──────────────────────────────────────────────────
console.log(`\n📂 Lecture : ${inputPath}`);
console.log(`📋 Onglet  : ${sheetName}\n`);

const workbook = XLSX.readFile(inputPath, { cellDates: false, raw: false });
const sheet    = workbook.Sheets[sheetName];

if (!sheet) {
  console.error(`Onglet "${sheetName}" introuvable. Onglets disponibles : ${workbook.SheetNames.join(', ')}`);
  process.exit(1);
}

const rows = XLSX.utils.sheet_to_json(sheet, { defval: '-', header: 0 });
console.log(`✅ ${rows.length} lignes lues\n`);

// ─── Noms de colonnes (position par index pour robustesse) ───────────────────
// On travaille sur les valeurs des cellules par index de colonne (A=0, B=1…)
// sheet_to_json avec header:0 utilise la première ligne comme clés
// On va re-parser avec header:1 pour utiliser les indices
const rowsRaw = XLSX.utils.sheet_to_json(sheet, { defval: '-', header: 1 });
const header  = rowsRaw[0]; // Ligne d'entête
const dataRows = rowsRaw.slice(1); // Données

// ─── Filtrage et classification ───────────────────────────────────────────────
const log = { total: dataRows.length, filteredIT: 0, opex: 0, capex: 0, skipped: 0, errors: [] };

const opexGroups  = new Map(); // clé: "Fournisseur||Compte"
const capexGroups = new Map();
const opexOrders  = [];
const capexOrders = [];

for (let i = 0; i < dataRows.length; i++) {
  const r = dataRows[i];

  try {
    // Indices 0-based selon spec (col[0]=Code projet … col[24]=Montant réalisé)
    const codeGestionnaire = String(r[2]  || '').trim(); // col[2]  Code Gestionnaire Commande
    const codeUF           = num(r[5]);                  // col[5]  Code UF
    const noCommande       = String(r[6]  || '').trim(); // col[6]  N° de Commande
    const noLigne          = String(r[7]  || '').trim(); // col[7]  N° ligne commande
    const designation      = String(r[8]  || '').trim(); // col[8]  Désignation ligne
    const noMarche         = num(r[9]);                  // col[9]  N° du marché
    const datePassation    = parseSageDate(r[10]);       // col[10] Date de passation
    const dateImputation   = parseSageDate(r[11]);       // col[11] Date imputation
    const dateReception    = parseSageDate(r[12]);       // col[12] Date réception effective
    const compte           = String(r[13] || '').trim(); // col[13] Compte ordonnateur
    const libelleCompte    = String(r[14] || '').trim(); // col[14] Libellé compte ordonnateur
    const fournisseur      = String(r[15] || '').trim(); // col[15] Fournisseur
    const etat             = String(r[16] || '').trim(); // col[16] Etat ligne commande
    const typeCmd          = String(r[17] || '').trim(); // col[17] Type de commande
    const montantEngage    = num(r[18]);                 // col[18] Montant engagé
    const montantEngNonRec = num(r[19]);                 // col[19] Montant engagé non reçu
    const montantMandateNet= num(r[22]);                 // col[22] Montant mandaté net
    const montantRealise   = num(r[24]);                 // col[24] Montant réalisé

    // Filtre gestionnaire IT
    if (codeGestionnaire !== 'IT') { log.skipped++; continue; }
    log.filteredIT++;

    const lineType = getLineType(compte);
    if (!lineType) { log.skipped++; continue; }

    const groupKey  = `${fournisseur}||${compte}`;
    const famille   = COMPTE_TO_FAMILLE[compte] || 'Hors périmètre DSI';
    const status    = mapSageStatus(etat, montantEngNonRec);
    const reference = `${noCommande}/${noLigne}`;

    if (lineType === 'OPEX') {
      log.opex++;

      // Groupement fournisseur OPEX
      if (!opexGroups.has(groupKey)) {
        opexGroups.set(groupKey, {
          id:               genId(),
          supplier:         fournisseur,
          category:         libelleCompte,
          compteOrdonnateur: compte,
          familleAnalytique: famille,
          budgetAnnuel:     0,
          depenseActuelle:  0,
          engagement:       0,
          montantRealise:   0,
          nbCommandes:      0,
          codeUF:           codeUF || 250,
          notes:            '',
        });
      }
      const grp = opexGroups.get(groupKey);
      grp.depenseActuelle += montantMandateNet;
      grp.engagement      += montantEngNonRec;
      grp.montantRealise  += montantRealise;
      grp.nbCommandes     += 1;

      // Commande OPEX
      opexOrders.push({
        id:               genId(),
        parentId:         grp.id,
        description:      designation,
        montant:          montantEngage,
        status,
        dateCommande:     datePassation,
        dateFacture:      dateImputation,
        dateReception,
        reference,
        numeroMarche:     noMarche,
        typeCommande:     typeCmd,
        etatSage:         etat,
        compteOrdonnateur: compte,
        notes:            '',
      });

    } else { // CAPEX
      log.capex++;

      if (!capexGroups.has(groupKey)) {
        capexGroups.set(groupKey, {
          id:               genId(),
          project:          `${fournisseur} — ${libelleCompte}`,
          enveloppe:        famille,
          compteOrdonnateur: compte,
          budgetTotal:      0,
          depense:          montantMandateNet,
          engagement:       montantEngNonRec,
          montantRealise:   montantRealise,
          status:           'En cours',
          startDate:        datePassation,
          endDate:          '',
          notes:            '',
        });
      } else {
        const grp = capexGroups.get(groupKey);
        grp.depense    += montantMandateNet;
        grp.engagement += montantEngNonRec;
        grp.montantRealise += montantRealise;
      }
      const grpCapex = capexGroups.get(groupKey);

      capexOrders.push({
        id:               genId(),
        parentId:         grpCapex.id,
        description:      designation,
        montant:          montantEngage,
        status,
        dateCommande:     datePassation,
        dateFacture:      dateImputation,
        dateReception,
        reference,
        numeroMarche:     noMarche,
        typeCommande:     typeCmd,
        etatSage:         etat,
        compteOrdonnateur: compte,
        notes:            '',
      });
    }

  } catch (err) {
    log.errors.push({ ligne: i + 2, erreur: err.message });
  }
}

// ─── Arrondi des montants ─────────────────────────────────────────────────────
const round2 = (v) => Math.round(v * 100) / 100;
const opexSuppliers = [...opexGroups.values()].map(s => ({
  ...s,
  depenseActuelle: round2(s.depenseActuelle),
  engagement:      round2(s.engagement),
  montantRealise:  round2(s.montantRealise),
}));
const capexProjects = [...capexGroups.values()].map(p => ({
  ...p,
  depense:        round2(p.depense),
  engagement:     round2(p.engagement),
  montantRealise: round2(p.montantRealise),
}));

// ─── Écriture des fichiers ────────────────────────────────────────────────────
const write = (filename, data) => {
  const filepath = join(outputDir, filename);
  writeFileSync(filepath, JSON.stringify(data, null, 2), 'utf8');
  console.log(`  ✅ ${filename} — ${data.length} entrées`);
};

console.log(`\n📝 Écriture des fichiers dans ${outputDir} :\n`);
write('opex.json',         opexSuppliers);
write('opex-orders.json',  opexOrders);
write('capex.json',        capexProjects);
write('capex-orders.json', capexOrders);

const importLog = {
  date: new Date().toISOString(),
  source: inputPath,
  sheet: sheetName,
  stats: {
    totalLignes:   log.total,
    filtreesIT:    log.filteredIT,
    lignesOPEX:    log.opex,
    lignesCAPEX:   log.capex,
    ignorees:      log.skipped,
    fournisseursOPEX: opexSuppliers.length,
    fournisseursCAPEX: capexProjects.length,
    commandesOPEX: opexOrders.length,
    commandesCAPEX: capexOrders.length,
    erreurs:       log.errors.length,
  },
  erreurs: log.errors,
};
write('import-log.json', [importLog]);

// ─── Résumé ───────────────────────────────────────────────────────────────────
console.log(`
📊 Résumé :
   Lignes totales    : ${log.total}
   Filtrées (IT)     : ${log.filteredIT}
   Fournisseurs OPEX : ${opexSuppliers.length}
   Commandes OPEX    : ${opexOrders.length}
   Projets CAPEX     : ${capexProjects.length}
   Commandes CAPEX   : ${capexOrders.length}
   Ignorées          : ${log.skipped}
   Erreurs           : ${log.errors.length}
`);

if (log.errors.length > 0) {
  console.warn('⚠️  Erreurs détectées :');
  log.errors.forEach(e => console.warn(`   Ligne ${e.ligne}: ${e.erreur}`));
}
