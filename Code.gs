/**
 * ACTION MSP · FEMAS Hauts-de-France
 * Script du « réservoir » Google Sheets.
 * À coller dans : Extensions > Apps Script, puis Déployer > Nouveau déploiement > Application Web
 *   - Exécuter en tant que : Moi
 *   - Qui a accès : Tout le monde
 * Aucune donnée de patient n’est enregistrée : seulement des totaux, des textes de fiches et des réponses anonymes au quiz.
 */

var TABS = {
  MSP:    ['Code', 'Maison de santé', 'Commune', 'Référent', 'E-mail', 'Créé le', 'Mis à jour le'],
  ETAT:   ['Code', 'Données 1 (ne pas modifier)', 'Données 2', 'Données 3', 'Données 4', 'Mis à jour le'],
  JAUGE:  ['Code', 'Maison de santé', 'Objectif', 'Patients repérés', 'Dont fumeurs', 'Souhaitent arrêter', 'Orientés / accompagnés', 'Mis à jour le'],
  STAND:  ['Code', 'Maison de santé', 'Date du stand', 'Personnes rencontrées', 'Fumeurs', 'Mesures du CO', 'Documentations', 'Orientations', 'Mis à jour le'],
  ARS:    ['Code', 'Maison de santé', 'Actions', 'Type de fiche', 'Titre', 'Objectifs prévus', 'Étapes faites', 'Rubriques remplies', 'Mis à jour le'],
  BILAN:  ['Code', 'Maison de santé', 'Actions menées', 'Prévu / réalisé', 'Ce qui a marché', 'À améliorer', 'Et l’an prochain', 'Envoyé le', 'Mis à jour le'],
  FICHES: ['Code · fiche', 'Code', 'Maison de santé', 'Fiche', 'Titre', 'Responsable', 'Diagnostic', 'Objectif général', 'Objectifs spécifiques', 'Objectifs opérationnels', 'Public cible', 'Acteurs MSP', 'Partenaires', 'Description', 'Calendrier', 'Moyens', 'Budget', 'Indicateurs', 'Évaluation', 'Constats', 'Envoyée le', 'Mis à jour le'],
  AVIS:   ['Code', 'Maison de santé', 'Satisfaction /5', 'Gain de temps', 'Prochaine action', 'Remarque', 'Mis à jour le'],
  QUIZ:   ['Horodatage', 'Code', 'Score /8', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Souhaite en parler']
};

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(TABS[name]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, TABS[name].length).setFontWeight('bold').setBackground('#3558a2').setFontColor('#ffffff');
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* Empêche qu’un texte commençant par = + - @ soit lu comme une formule. */
function t_(v) {
  v = String(v == null ? '' : v).slice(0, 5000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function clean_(code) {
  return String(code || '').toUpperCase().replace(/[^A-Z0-9-]/g, '').slice(0, 40);
}

/* Remplace la ligne dont la colonne A vaut « code », ou l’ajoute. */
function upsert_(name, code, values) {
  var sh = sheet_(name);
  var data = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === code) {
      sh.getRange(i + 1, 1, 1, values.length).setValues([values]);
      return;
    }
  }
  sh.appendRow(values);
}

function find_(name, code) {
  var sh = sheet_(name);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) if (String(data[i][0]) === code) return data[i];
  return null;
}

/* ---------- lectures ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || '';
  var code = clean_(p.code);
  try {
    if (action === 'load') {
      var row = find_('ETAT', code);
      return out_({ ok: !!row, state: row ? JSON.parse([row[1], row[2], row[3], row[4]].map(function (x) { return String(x || '').replace(/^~/, ''); }).join('')) : null });
    }
    if (action === 'quizstats') {
      var rows = sheet_('QUIZ').getDataRange().getValues().slice(1).filter(function (r) { return String(r[1]) === code; });
      var n = rows.length, q = [0, 0, 0, 0, 0, 0, 0, 0], tot = 0, parler = 0;
      rows.forEach(function (r) {
        tot += Number(r[2]) || 0;
        for (var k = 0; k < 8; k++) if (r[3 + k] === 1 || r[3 + k] === '1' || r[3 + k] === true) q[k]++;
        if (r[11] === 1 || r[11] === '1' || r[11] === true) parler++;
      });
      return out_({ ok: true, n: n, avg: n ? tot / n : 0, correct: q.map(function (x) { return n ? x / n : 0; }), parler: parler });
    }
    if (action === 'count') {
      return out_({ ok: true, count: Math.max(0, sheet_('MSP').getLastRow() - 1) });
    }
    if (action === 'exists') {
      return out_({ ok: true, exists: !!find_('ETAT', code) });
    }
    return out_({ ok: true, service: 'Action MSP' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/* ---------- écritures ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var code = clean_(body.code);
    var now = new Date();
    if (!code) return out_({ ok: false, error: 'code manquant' });

    if (body.action === 'save') {
      var s = body.state || {};
      var pf = s.profil || {};
      var msp = s.msp || '';
      var prev = find_('MSP', code);
      upsert_('MSP', code, [code, t_(msp), t_(pf.commune), t_(pf.ref), t_(pf.email), prev ? prev[5] : now, now]);
      var js = JSON.stringify(s), ch = [];
      if (js.length > 180000) return out_({ ok: false, error: 'données trop volumineuses' });
      for (var c = 0; c < 4; c++) ch.push('~' + js.slice(c * 45000, (c + 1) * 45000));
      upsert_('ETAT', code, [code].concat(ch).concat([now]));
      var sm = body.summary || {};
      if (sm.jauge) upsert_('JAUGE', code, [code, t_(msp), sm.jauge.goal, sm.jauge.rep, sm.jauge.fum, sm.jauge.arr, sm.jauge.ori, now]);
      if (sm.stand && sm.stand.n) upsert_('STAND', code, [code, t_(msp), t_(sm.stand.date), sm.stand.n, sm.stand.fum, sm.stand.co, sm.stand.doc, sm.stand.ori, now]);
      if (sm.ars) upsert_('ARS', code, [code, t_(msp), t_(sm.ars.actions), t_(sm.ars.mode), t_(sm.ars.titre), t_(sm.ars.prevu), t_(sm.ars.etapes), t_(sm.ars.filled), now]);
      if (sm.bilan) upsert_('BILAN', code, [code, t_(msp), t_(sm.bilan.actions), t_(sm.bilan.evaluation), t_(sm.bilan.marche), t_(sm.bilan.ameliorer), t_(sm.bilan.suite), t_(sm.bilan.sentAt), now]);
      if (sm.fiches) sm.fiches.forEach(function (f) {
        var v = f.vals || {}, k = String(f.key || '');
        upsert_('FICHES', code + ' · ' + k, [code + ' · ' + k, code, t_(msp), k === 'commune' ? 'Fiche commune' : 'Action ' + k].concat(['titre','responsable','diagnostic','objGeneral','objSpec','objOp','public','acteurs','partenaires','description','calendrier','moyens','budget','indicateurs','evaluation','constats'].map(function (x) { return t_(v[x]); })).concat([t_(sm.sentAt), now]));
      });
      if (sm.avis) upsert_('AVIS', code, [code, t_(msp), t_(sm.avis.note), t_(sm.avis.temps), t_(sm.avis.refaire), t_(sm.avis.com), now]);
      return out_({ ok: true, savedAt: now.toISOString() });
    }

    if (body.action === 'quiz') {
      var a = (body.answers || []).slice(0, 8);
      while (a.length < 8) a.push('');
      sheet_('QUIZ').appendRow([now, code, Number(body.score) || 0].concat(a.map(function (x) { return x === true ? 1 : (x === false ? 0 : ''); })).concat([body.parler ? 1 : 0]));
      return out_({ ok: true });
    }

    return out_({ ok: false, error: 'action inconnue' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* À lancer une fois depuis l’éditeur (bouton ▶ Exécuter) pour créer les onglets et autoriser le script. */
function installer() {
  Object.keys(TABS).forEach(sheet_);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var first = ss.getSheets()[0];
  if (first.getName() === 'Feuille 1' || first.getName() === 'Sheet1') ss.deleteSheet(first);
}
