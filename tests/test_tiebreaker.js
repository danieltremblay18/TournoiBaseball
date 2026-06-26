// ============================================================================
//  Test hors-ligne du bris d'égalité (Art. 42.11)
// ============================================================================
//
// Le projet n'a pas de cadre de test : ce fichier charge les fonctions PURES de
// classement de TournoiBaseball_Script.gs (celles qui n'appellent pas
// SpreadsheetApp) et les exécute sur un scénario fabriqué, sans Google Sheets.
//
// Pourquoi : la logique du bris d'égalité (égalités circulaires, ratios
// RA/manches, récursion) est la partie la plus délicate du script et la plus
// facile à casser par mégarde. Ce test vérifie qu'un changement de score qui
// laisse les fiches V-D identiques peut quand même faire BASCULER l'ordre du
// pool via la Priorité 2 (ratio RA/DefInn).
//
// Lancer :   node tests/test_tiebreaker.js
// (Exécutable depuis n'importe quel dossier — le chemin du .gs est résolu
//  relativement à ce fichier.)
//
// Note : on charge le .gs par eval. C'est volontaire — ça teste le code RÉEL
// du script, pas une copie qui pourrait diverger. Aucune fonction n'est exécutée
// au chargement ; seules les déclarations de fonctions/constantes le sont.
// ============================================================================

var fs   = require('fs');
var path = require('path');

var GS_PATH = path.join(__dirname, '..', 'TournoiBaseball_Script.gs');
var src = fs.readFileSync(GS_PATH, 'utf8');

// Stubs pour les globals Apps Script (jamais réellement appelés par les
// fonctions pures testées ici, mais référencés dans le fichier).
var Logger = { log: function () {} };
var TOTAL_INNINGS = 6;

// Définit toutes les fonctions/constantes du script dans cette portée.
eval(src);

// --- Aides de construction d'un scénario ------------------------------------

// Une partie complète : 6 manches, fin "Normal" -> manches symétriques 6/6.
function game(pool, a, b, sa, sb) {
  var winner = sa > sb ? a : (sb > sa ? b : '');
  return {
    pool: pool, local: a, visiteur: b,
    scoreLocal: sa, scoreVisiteur: sb,
    winner: winner, type: 'Normal',
    offLocal: 6, defLocal: 6, offVisiteur: 6, defVisiteur: 6
  };
}

var teams = ['Kamouraska', 'RiviereDuLoup', 'Temiscouata', 'Montmagny'];

// Égalité circulaire à 3 (chacune 2-1) : K bat R, R bat T, T bat K ; toutes
// battent Montmagny (0-3). La marge de la partie Témiscouata-Kamouraska est
// paramétrable pour provoquer la bascule.
function buildGames(temiScoreVsKam) {
  return [
    game(1, 'Kamouraska',    'RiviereDuLoup', 10, 0),
    game(1, 'RiviereDuLoup', 'Temiscouata',    5, 4),
    game(1, 'Temiscouata',   'Kamouraska', temiScoreVsKam, 1),
    game(1, 'Kamouraska',    'Montmagny',      7, 0),
    game(1, 'RiviereDuLoup', 'Montmagny',      7, 0),
    game(1, 'Temiscouata',   'Montmagny',      7, 0)
  ];
}

// --- Exécution + vérification -----------------------------------------------

function rankOf(stats, team) {
  for (var i = 0; i < stats.length; i++) {
    if (stats[i].team === team) { return stats[i].rank; }
  }
  return -1;
}

function show(label, stats) {
  console.log('\n=== ' + label + ' ===');
  stats.forEach(function (s) {
    console.log(
      '  ' + s.rank + '. ' + s.team.padEnd(15) +
      ' V-D ' + s.v + '-' + s.d +
      ' | RA ' + s.ra + ' RS ' + s.rs +
      ' | RA/DefInn ' + s.raRatio.toFixed(3)
    );
  });
}

var before = calculatePoolStandings(buildGames(2), teams);
var after  = calculatePoolStandings(buildGames(9), teams);

show('AVANT (Temiscouata-Kamouraska 2-1)', before);
show('APRES (meme partie corrigee a 9-1)', after);

// Attendu : V-D inchangées (toujours 2-1 pour les trois meneuses), mais l'ordre
// bascule via le ratio RA/DefInn -> Kamouraska passe 1re -> 2e, Temiscouata 2e -> 1re.
var checks = [
  ['Kamouraska 1re AVANT',     rankOf(before, 'Kamouraska')  === 1],
  ['Temiscouata 2e AVANT',     rankOf(before, 'Temiscouata') === 2],
  ['Temiscouata 1re APRES',    rankOf(after,  'Temiscouata') === 1],
  ['Kamouraska 2e APRES',      rankOf(after,  'Kamouraska')  === 2],
  ['RiviereDuLoup reste 3e',   rankOf(before, 'RiviereDuLoup') === 3 &&
                               rankOf(after,  'RiviereDuLoup') === 3]
];

console.log('\n--- Vérifications : bris d\'égalité ---');
var allOk = true;
checks.forEach(function (c) {
  console.log((c[1] ? '  OK   ' : '  ÉCHEC') + ' : ' + c[0]);
  if (!c[1]) { allOk = false; }
});

// --- Test de régression : exemple officiel du PDF (Note 2, Art. 42.11) -------
// QC bat RS 6-3, CN bat QC 6-4, RS bat CN 10-8. Égalité à 3 (chacune 1-1).
//   P1 (tête-à-tête V-D) : tous à 0 -> pas de séparation.
//   P2 (RA/DefInn) : QC 9/12=0.75 isolé 1er ; RS 14/12 et CN 14/12 à égalité.
//   P3 (RS/OffInn) sur la portée ORIGINALE (3 équipes) : RS 13/12 < CN 14/12
//     -> CN 2e, RS 3e. Résultat attendu : QC > CN > RS.
// Le bug Note 2 (recommencer à P1 sur {RS,CN}) donnerait QC > RS > CN, car RS a
// battu CN en tête-à-tête.
var pdfGames = [
  game(1, 'QC', 'RS',  6, 3),
  game(1, 'CN', 'QC',  6, 4),
  game(1, 'RS', 'CN', 10, 8)
];
var pdfOrder = orderTeams(['QC', 'RS', 'CN'], pdfGames, false);

console.log('\n--- Vérification : régression PDF (Note 2) ---');
console.log('  Ordre produit : ' + pdfOrder.join(' > ') + ' (attendu : QC > CN > RS)');
var pdfOk = pdfOrder.join(' > ') === 'QC > CN > RS';
console.log((pdfOk ? '  OK   ' : '  ÉCHEC') + ' : QC > CN > RS (continue à P3, ne recommence pas à P1)');
if (!pdfOk) { allOk = false; }

// --- Test : Priorité 4 (manuelle) après épuisement des ratios P2/P3 -----------
// Art. 42.11, Note 2 : on CONTINUE aux priorités suivantes (pas de retour à la
// P1) ; une fois P2 et P3 épuisées sans départage, la priorité suivante est la
// P4 (« manches avec l'avance au pointage »), NON automatisable (feuille de
// pointage) -> drapeau de vérification manuelle.
//
// 4 équipes à égalité. P1 (sur les 4) sépare {B,C} (fiche 2-1) au-dessus de
// {A,D} (fiche 1-2). {B,C} CONTINUE à P2/P3 sur la portée des 4 : B et C ont des
// ratios défensif ET offensif IDENTIQUES (RA 8, RS 13, 18 manches chacun) -> P2
// et P3 ne séparent rien -> P4 -> drapeau manuel + ordre alphabétique provisoire.
// {A,D} restent séparables par P2 (D RA 11 < A RA 13) -> D avant A.
//
// Le duel direct est gagné par C (5-3) : l'ANCIEN code bogué (retour-à-P1
// re-restreint au duel) aurait donné C > B SANS drapeau. Le code correct donne
// B > C (alphabétique) AVEC drapeau. Les deux assertions ensemble le prouvent.
var p4Games = [
  game(1, 'C', 'B', 5, 3), // C bat B  <- duel direct (NON consulté : on s'arrête à P4)
  game(1, 'B', 'A', 6, 1), // B bat A
  game(1, 'B', 'D', 4, 2), // B bat D
  game(1, 'C', 'A', 6, 2), // C bat A
  game(1, 'D', 'C', 3, 2), // D bat C
  game(1, 'A', 'D', 5, 1)  // A bat D
];
var p4Order = orderTeams(['A', 'B', 'C', 'D'], p4Games, false);

console.log('\n--- Vérification : Priorité 4 manuelle (Note 2) ---');
console.log('  Ordre produit : ' + p4Order.join(' > ') + ' (attendu : B > C > D > A)');
console.log('  __needsManualCheck : ' + (p4Order.__needsManualCheck === true));
var p4OrderOk = p4Order.join(' > ') === 'B > C > D > A';
var p4FlagOk  = p4Order.__needsManualCheck === true;
console.log((p4OrderOk ? '  OK   ' : '  ÉCHEC') +
            ' : B > C alphabétique (PAS le duel C>B) ; D > A par P2');
console.log((p4FlagOk ? '  OK   ' : '  ÉCHEC') +
            ' : drapeau vérification manuelle levé (P4 atteinte, pas de retour à P1)');
if (!p4OrderOk || !p4FlagOk) { allOk = false; }

// --- Test de isRowComplete (gate du recalcul live onEdit) -------------------
// Colonnes H..O : scoreA, scoreB, local, manches, retraits, type, suppTie.
function rc(scoreA, scoreB, local, manches, retraits, type, suppTie) {
  return isRowComplete({ scoreA: scoreA, scoreB: scoreB, local: local,
                         manches: manches, retraits: retraits, type: type,
                         suppTie: suppTie });
}

var rowChecks = [
  // Partie normale complète (6 colonnes remplies) -> complète.
  ['Normale complète',
   rc(5, 3, 'Kamouraska', 6, 0, 'Normal') === true],
  // Score 0 et Retraits 0 doivent compter comme REMPLIS (pas comme vides).
  ['Score 0 / retraits 0 = rempli',
   rc(0, 0, 'Kamouraska', 6, 0, 'Normal') === true],
  // Journée de pluie : partie écourtée à 5 manches, explicitement saisie.
  ['Pluie (5 manches) complète',
   rc(4, 2, 'Kamouraska', 5, 0, 'Normal') === true],
  // Forfait : seuls scores + type requis (manches/retraits/locale vides).
  ['Forfait sans manches/locale',
   rc(7, 0, '', '', '', 'Forfait') === true],
  // Saisie partielle : un seul score -> incomplète (pas de recalcul).
  ['Un seul score = incomplète',
   rc(5, '', 'Kamouraska', 6, 0, 'Normal') === false],
  // Manches manquantes sur partie normale -> incomplète (le cas pluie l'exige).
  ['Manches manquante = incomplète',
   rc(5, 3, 'Kamouraska', '', 0, 'Normal') === false],
  // Équipe Locale manquante sur partie normale -> incomplète.
  ['Locale manquante = incomplète',
   rc(5, 3, '', 6, 0, 'Normal') === false],
  // Type de fin manquant -> incomplète (on ne sait pas si c'est un forfait).
  ['Type manquant = incomplète',
   rc(5, 3, 'Kamouraska', 6, 0, '') === false],
  // Supplémentaires AVEC pointage régl. (col O) -> complète.
  ['Suppl. avec pointage régl. = complète',
   rc(7, 5, 'Kamouraska', 7, 0, 'Supplémentaires', 5) === true],
  // Supplémentaires SANS pointage régl. -> incomplète (Note 4 non applicable).
  ['Suppl. sans pointage régl. = incomplète',
   rc(7, 5, 'Kamouraska', 7, 0, 'Supplémentaires', '') === false],
  // Pointage régl. = 0 doit compter comme REMPLI (nul 0-0 après le réglementaire).
  ['Suppl. pointage régl. 0 = complète',
   rc(2, 1, 'Kamouraska', 7, 0, 'Supplémentaires', 0) === true]
];

console.log('\n--- Vérifications : isRowComplete (gate onEdit) ---');
rowChecks.forEach(function (c) {
  console.log((c[1] ? '  OK   ' : '  ÉCHEC') + ' : ' + c[0]);
  if (!c[1]) { allOk = false; }
});

// --- Test de gameIsSupp (drapeau Note 4 / manches supplémentaires) ----------
var suppChecks = [
  ['Type Supplémentaires détecté', gameIsSupp({ type: 'Supplémentaires' }) === true],
  ['Type Normal non détecté',      gameIsSupp({ type: 'Normal' }) === false],
  ['Objet vide non détecté',       gameIsSupp({}) === false]
];

console.log('\n--- Vérifications : gameIsSupp (Note 4) ---');
suppChecks.forEach(function (c) {
  console.log((c[1] ? '  OK   ' : '  ÉCHEC') + ' : ' + c[0]);
  if (!c[1]) { allOk = false; }
});

// --- Test de calculateInnings (manches réglementaires + victoire locale) -----
// Signature : calculateInnings(scoreLocal, scoreVis, manchesCompletes, retraits,
//             typeFin, homeKnown, regulation).
// Vérifie : (1) toute victoire LOCALE crédite la locale d'une dernière manche
// offensive incomplète et retranche la dernière manche défensive de la visiteuse
// (pas de manche défensive offerte) ; (2) le crédit défensif d'un Mercy suit les
// manches PRÉVUES (5 lors d'une journée écourtée), pas un 6 figé.
function approx(a, b) { return Math.abs(a - b) < 1e-9; }
function innEq(inn, oL, dL, oV, dV) {
  return approx(inn.offLocal, oL) && approx(inn.defLocal, dL) &&
         approx(inn.offVisiteur, oV) && approx(inn.defVisiteur, dV);
}
var innChecks = [
  // Victoire locale complète : la locale menait, n'a pas frappé en bas -> 5/6/6/5.
  ['Victoire locale complète 7-5 (K6 L0 prév6) -> 5/6/6/5',
   innEq(calculateInnings(7, 5, 6, 0, 'Normal', true, 6), 5, 6, 6, 5)],
  // Walk-off bas 6e avec 2 retraits -> 5⅔ / 6 / 6 / 5⅔ (inchangé).
  ['Walk-off 6-5 (K6 L2) -> 5⅔/6/6/5⅔',
   innEq(calculateInnings(6, 5, 6, 2, 'Normal', true, 6), 5 + 2 / 3, 6, 6, 5 + 2 / 3)],
  // Mercy locale, partie prévue 5 manches : def locale = 5 (pas 6).
  ['Mercy locale 10-0 prévues=5 (K5 L0) -> 4/5/5/4',
   innEq(calculateInnings(10, 0, 5, 0, 'Mercy', true, 5), 4, 5, 5, 4)],
  // Mercy visiteuse, prévues 5 : def visiteuse (gagnante) = 5.
  ['Mercy visiteuse 0-12 prévues=5 (K5 L0) -> 5/5/5/5',
   innEq(calculateInnings(0, 12, 5, 0, 'Mercy', true, 5), 5, 5, 5, 5)],
  // Victoire visiteuse complète : symétrique 6/6/6/6 (la locale a frappé jusqu'au bout).
  ['Victoire visiteuse 3-7 (K6 L0) -> 6/6/6/6',
   innEq(calculateInnings(3, 7, 6, 0, 'Normal', true, 6), 6, 6, 6, 6)],
  // Équipe locale non indiquée (homeKnown=false) : repli symétrique.
  ['Locale inconnue 7-5 (homeKnown=false) -> 6/6/6/6',
   innEq(calculateInnings(7, 5, 6, 0, 'Normal', false, 6), 6, 6, 6, 6)],
  // Forfait gagné par la locale (prévues 6) : gagnant 0 off / 6 déf, perdant 6 off / 0 déf.
  ['Forfait locale 6-0 (prév6) -> 0/6/6/0',
   innEq(calculateInnings(6, 0, 6, 0, 'Forfait', true, 6), 0, 6, 6, 0)],
  // Forfait gagné par la visiteuse, journée écourtée (prévues 5) : miroir 5/0/0/5.
  ['Forfait visiteuse 0-5 (prév5) -> 5/0/0/5',
   innEq(calculateInnings(0, 5, 5, 0, 'Forfait', true, 5), 5, 0, 0, 5)]
];

console.log('\n--- Vérifications : calculateInnings (manches prévues + victoire locale) ---');
innChecks.forEach(function (c) {
  console.log((c[1] ? '  OK   ' : '  ÉCHEC') + ' : ' + c[0]);
  if (!c[1]) { allOk = false; }
});

// --- Test Note 4 : exclusion des manches supplémentaires du ratio ------------
// Partie allée en supplémentaires : score FINAL 7-5 (la locale gagne en suppl.),
// mais NULLE 5-5 au terme des 6 manches réglementaires. computeTeamStats doit :
//  - afficher PP/PC = totaux RÉELS (7 et 5, suppl. incluses) ;
//  - calculer les ratios sur les manches RÉGULIÈRES seulement (5 marqués / 5
//    alloués sur 6 manches → 5/6), surtout PAS sur le score final (7/7 et 5/7).
// C'est exactement la Note 4 de l'Art. 42.11.
function suppGame(local, vis, finalLoc, finalVis, tie, reg) {
  return {
    pool: 1, local: local, visiteur: vis,
    scoreLocal: finalLoc, scoreVisiteur: finalVis,
    winner: finalLoc > finalVis ? local : vis, type: 'Supplémentaires',
    // Manches RÉELLES (suppl. incluses) — journal Résultats, jamais le ratio.
    offLocal: 7, defLocal: 7, offVisiteur: 7, defVisiteur: 7,
    // Base RÉGULIÈRE (Note 4) : pointage nul X sur `reg` manches, pour les deux.
    regRsLocal: tie, regRsVisiteur: tie,
    regOffLocal: reg, regDefLocal: reg, regOffVisiteur: reg, regDefVisiteur: reg,
    suppNeedsTie: false
  };
}
var sgStats = computeTeamStats('X', [suppGame('X', 'Y', 7, 5, 5, 6)], false);
var note4Checks = [
  ['PP affiché = 7 (score final réel)',          sgStats.rs === 7],
  ['PC affiché = 5 (score final réel)',          sgStats.ra === 5],
  ['RO = 5/6 (régulier, PAS 7/7 du score final)', approx(sgStats.rsRatio, 5 / 6)],
  ['RD = 5/6 (régulier, PAS 5/7 du score final)', approx(sgStats.raRatio, 5 / 6)],
  ['MO du ratio = 6 (régulières, PAS 7 réelles)', approx(sgStats.offInn, 6)],
  ['MD du ratio = 6 (régulières, PAS 7 réelles)', approx(sgStats.defInn, 6)]
];

console.log('\n--- Vérifications : Note 4 (exclusion des suppl. du ratio) ---');
note4Checks.forEach(function (c) {
  console.log((c[1] ? '  OK   ' : '  ÉCHEC') + ' : ' + c[0]);
  if (!c[1]) { allOk = false; }
});

// --- Test Note 5 : forfaits exclus des ratios du « Meilleur deuxième » -------
// Un forfait gagné par la locale (pointage officiel 6-0) crédite le gagnant de
// 0 manche offensive / 6 défensives, le perdant de 6 off / 0 déf (Art. 42.11).
// computeTeamStats doit :
//  - excludeForfaitRatios=false (Étape A / tableau de pool) : la partie compte
//    dans les ratios (rsNum/raNum + offInn/defInn) ;
//  - excludeForfaitRatios=true (Étapes B/C) : la partie est EXCLUE EN ENTIER des
//    ratios — autant les points que les manches (sinon le numérateur garderait
//    des points sans dénominateur).
// La fiche V-D et les PP/PC réels, eux, comptent dans les deux cas.
function forfaitGame(local, vis) {
  return {
    pool: 1, local: local, visiteur: vis,
    scoreLocal: 6, scoreVisiteur: 0, winner: local, type: 'Forfait',
    offLocal: 0, defLocal: 6, offVisiteur: 6, defVisiteur: 0
  };
}
var fWin  = computeTeamStats('W', [forfaitGame('W', 'L')], false); // gagnant, inclus
var fWinX = computeTeamStats('W', [forfaitGame('W', 'L')], true);  // gagnant, exclu
var fLose = computeTeamStats('L', [forfaitGame('W', 'L')], false); // perdant, inclus
var note5Checks = [
  ['Inclus : gagnant rsNum=6 / raNum=0',      fWin.rsNum === 6 && fWin.raNum === 0],
  ['Inclus : gagnant offInn=0 / defInn=6',    approx(fWin.offInn, 0) && approx(fWin.defInn, 6)],
  ['Inclus : perdant offInn=6 / defInn=0',    approx(fLose.offInn, 6) && approx(fLose.defInn, 0)],
  ['Exclu : ratios vidés (rsNum/raNum/MO/MD = 0)',
   fWinX.rsNum === 0 && fWinX.raNum === 0 &&
   approx(fWinX.offInn, 0) && approx(fWinX.defInn, 0)],
  ['Exclu : fiche V-D conservée (1-0)',       fWinX.v === 1 && fWinX.d === 0],
  ['Exclu : PP/PC réels conservés (6/0)',     fWinX.rs === 6 && fWinX.ra === 0]
];

console.log('\n--- Vérifications : Note 5 (forfaits exclus du « Meilleur deuxième ») ---');
note5Checks.forEach(function (c) {
  console.log((c[1] ? '  OK   ' : '  ÉCHEC') + ' : ' + c[0]);
  if (!c[1]) { allOk = false; }
});

// --- Test : forçage admin du « 2e de pool » (resolveSecondRepresentative) ----
// L'admin peut forcer le représentant d'un pool au Meilleur 2e (Étape B) en
// inscrivant « 2 » à côté d'une équipe (col « Forcer 2e »). Préséance sur le 2e
// automatique, SAUF si la marque vise la 1re équipe (déjà qualifiée Étape C) ->
// ignorée + avertissement. Marques multiples valides -> ambigu -> ignoré.
function standing(team, rank) { return { team: team, rank: rank }; }
var poolStd = [standing('A', 1), standing('B', 2), standing('C', 3), standing('D', 4)];
var rsr = resolveSecondRepresentative;
var rsrChecks = [
  ['Aucun forçage -> 2e automatique (B), non forcé',
   rsr(poolStd, []).team === 'B' && rsr(poolStd, []).forced === false],
  ['Forçage sur le 3e (C) -> C forcé',
   rsr(poolStd, ['C']).team === 'C' && rsr(poolStd, ['C']).forced === true],
  ['Forçage sur le 2e lui-même (B) -> B (no-op)',
   rsr(poolStd, ['B']).team === 'B' && rsr(poolStd, ['B']).forced === true],
  ['Forçage sur la 1re (A) -> ignoré, repli B + avertissement',
   (function () { var r = rsr(poolStd, ['A']);
     return r.team === 'B' && r.forced === false && r.warning !== ''; })()],
  ['Deux forçages valides (C,D) -> ambigu, repli B + avertissement',
   (function () { var r = rsr(poolStd, ['C', 'D']);
     return r.team === 'B' && r.forced === false && r.warning !== ''; })()],
  ['Marque hors pool ignorée (Z) -> 2e automatique (B)',
   rsr(poolStd, ['Z']).team === 'B' && rsr(poolStd, ['Z']).forced === false]
];

console.log('\n--- Vérifications : forçage admin du 2e (resolveSecondRepresentative) ---');
rsrChecks.forEach(function (c) {
  console.log((c[1] ? '  OK   ' : '  ÉCHEC') + ' : ' + c[0]);
  if (!c[1]) { allOk = false; }
});

if (!allOk) {
  console.error('\nÉCHEC : un test ne se comporte pas comme attendu.');
  process.exit(1);
}
console.log('\nSUCCÈS : bris d\'égalité (P2) + isRowComplete + gameIsSupp + Note 4 + Note 5 + forçage 2e OK.');
