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

// --- Test de isRowComplete (gate du recalcul live onEdit) -------------------
// Colonnes H..M : scoreA, scoreB, local, manches, retraits, type.
function rc(scoreA, scoreB, local, manches, retraits, type) {
  return isRowComplete({ scoreA: scoreA, scoreB: scoreB, local: local,
                         manches: manches, retraits: retraits, type: type });
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
   rc(5, 3, 'Kamouraska', 6, 0, '') === false]
];

console.log('\n--- Vérifications : isRowComplete (gate onEdit) ---');
rowChecks.forEach(function (c) {
  console.log((c[1] ? '  OK   ' : '  ÉCHEC') + ' : ' + c[0]);
  if (!c[1]) { allOk = false; }
});

if (!allOk) {
  console.error('\nÉCHEC : un test ne se comporte pas comme attendu.');
  process.exit(1);
}
console.log('\nSUCCÈS : bris d\'égalité (Priorité 2) + isRowComplete OK.');
