// Noms exacts des feuilles dans le Google Sheet
const SHEET_DISTRICT        = 'Base district';
const SHEET_REGION          = 'Base région';
const SHEET_ACTEURS_DISTRICT = 'Acteurs District';
const SHEET_ACTEURS_REGION   = 'Acteurs Région';

// ============================================================
//  Point d'entrée POST (formulaires HTML)
function doPost(e) {
  try {

    let payload;

    // ===== CAS 1 : JSON envoyé par fetch =====
    if (e.postData && e.postData.contents) {

      const content = e.postData.contents;

      if (content.startsWith('{')) {
        payload = JSON.parse(content);

      } else if (e.parameter.payload) {
        payload = JSON.parse(e.parameter.payload);

      } else {
        return ContentService
          .createTextOutput(JSON.stringify({
            status: 'error',
            message: 'Payload invalide.'
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }

    }

    // ===== CAS 2 : formulaire iframe =====
    else if (e.parameter.payload) {
      payload = JSON.parse(e.parameter.payload);

    } else {
      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'error',
          message: 'Aucune donnée reçue.'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let result;

    if (payload.action === 'updateDistrict') {
      result = updateDistrict(payload);

    } else if (payload.action === 'updateRegion') {
      result = updateRegion(payload);

    } else if (payload.action === 'saveActeursDistrict') {
      result = saveActeursDistrict(payload);

    } else if (payload.action === 'saveActeursRegion') {
      result = saveActeursRegion(payload);

    }  else if (payload.action === 'saveActeursCommunautaires') {
//      result = saveActeursCommunautaires(payload);

    } else {
      result = {
        status: 'error',
        message: 'Action inconnue : ' + payload.action
      };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
//  Point d'entrée GET (lecture des données existantes)
// ============================================================
function doGet(e) {
  try {
    const action = e.parameter.action;
    let result;

    if (action === 'getDistrict') {
      result = getDistrictData(e.parameter.region, e.parameter.district);
    } else if (action === 'getRegion') {
      result = getRegionData(e.parameter.region, e.parameter.district);
    } else if (action === 'getActeursDistrict') {
      result = getActeursDistrict(e.parameter.region, e.parameter.district);
    } else if (action === 'getActeursRegion') {
      result = getActeursRegion(e.parameter.region);

    } else if (action === 'getActeursCommunautaires') {
//      result = getActeursCommunautaires(e.parameter.region, e.parameter.district);

    } else {
      result = { status: 'error', message: 'Action GET inconnue.' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
//  Mise à jour feuille "Base district"
//  Colonnes : A=REGION  B=DISTRICTS  C=RIPOSTE  D=NB CSB cas
//             E=NB Fkt  F=NB AC  G=Voiture  H=Moto
// ============================================================
function updateDistrict(payload) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DISTRICT);

  if (!sheet) {
    return { status: 'error', message: 'Feuille "' + SHEET_DISTRICT + '" introuvable.' };
  }

  const data     = sheet.getDataRange().getValues();
  const region   = (payload.region   || '').trim().toUpperCase();
  const district = (payload.district || '').trim().toUpperCase();

  let targetRow = -1;

  // Chercher la ligne dont col A = région ET col B = district (ligne 1 = en-tête)
  for (let i = 1; i < data.length; i++) {
    const rowRegion   = String(data[i][0] || '').trim().toUpperCase();
    const rowDistrict = String(data[i][1] || '').trim().toUpperCase();
    if (rowRegion === region && rowDistrict === district) {
      targetRow = i + 1; // +1 car getValues() est 0-indexé mais setValues() est 1-indexé
      break;
    }
  }

  if (targetRow === -1) {
    return {
      status: 'error',
      message: 'District "' + payload.district + '" non trouvé dans la région "' + payload.region + '".'
    };
  }

  // Écriture des colonnes C à H
  sheet.getRange(targetRow, 3).setValue(payload.riposte  || '');  // C
  sheet.getRange(targetRow, 4).setValue(payload.nb_csb   || '');  // D
  sheet.getRange(targetRow, 5).setValue(payload.nb_fkt   || '');  // E
  sheet.getRange(targetRow, 6).setValue(payload.nb_ac    || '');  // F
  sheet.getRange(targetRow, 7).setValue(payload.voiture  || '');  // G
  sheet.getRange(targetRow, 8).setValue(payload.moto     || '');  // H

  return {
    status: 'ok',
    message: 'District ' + payload.district + ' mis à jour (ligne ' + targetRow + ').'
  };
}

// ============================================================
//  Mise à jour feuille "Base région"
//  Colonnes : A=REGION  B=DISTRICT  C=Voiture (V)  D=Moto (M)
//
//  IMPORTANT : La feuille "Base région" doit maintenant contenir
//  une colonne B "District" insérée entre Région et Voiture.
//  Les colonnes Voiture et Moto sont donc en C et D.
//
//  Pour la voiture : on cherche la ligne dont col A = région
//                    ET col B = district_voiture, puis on écrit
//                    "V" en colonne C.
//  Pour la moto    : on cherche la ligne dont col A = région
//                    ET col B = district_moto, puis on écrit
//                    "M" en colonne D.
//
//  payload attendu :
//    { action, region,
//      voiture: 'V'|'', district_voiture: '...',
//      moto:    'M'|'', district_moto:    '...' }
// ============================================================
function updateRegion(payload) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_REGION);

  if (!sheet) {
    return { status: 'error', message: 'Feuille "' + SHEET_REGION + '" introuvable.' };
  }

  // Relire les données fraîches à chaque opération
  const region = (payload.region || '').trim().toUpperCase();
  const messages = [];
  let hasError = false;

  // ── Diagnostic : retourner les 5 premières lignes brutes si action=debugRegion
  // (utile pour vérifier la structure réelle de la feuille)
  if (payload.debug) {
    const raw = sheet.getDataRange().getValues();
    const preview = raw.slice(0, 6).map((row, i) => 'L'+(i+1)+': ['+row.map(c=>'«'+c+'»').join(', ')+']');
    return { status: 'debug', preview: preview };
  }

  // --- Mise à jour Voiture (colonne C) ---
  if (payload.voiture !== undefined && payload.district_voiture) {
    const districtVoiture = payload.district_voiture.trim().toUpperCase();

    // Relire à chaque fois pour éviter les effets de bord
    const dataV = sheet.getDataRange().getValues();
    let rowV = -1;
    for (let i = 1; i < dataV.length; i++) {
      const rr = String(dataV[i][0] || '').trim().toUpperCase();
      const dd = String(dataV[i][1] || '').trim().toUpperCase();
      if (rr === region && dd === districtVoiture) {
        rowV = i + 1;
        break;
      }
    }
    if (rowV === -1) {
      messages.push('❌ District voiture "' + payload.district_voiture + '" non trouvé (région="' + region + '").');
      hasError = true;
    } else {
      sheet.getRange(rowV, 3).setValue('V');  // colonne C
      messages.push('✓ Voiture → ligne ' + rowV + ', col C (district: ' + payload.district_voiture + ')');
    }
  }

  // --- Mise à jour Moto (colonne D) ---
  if (payload.moto !== undefined && payload.district_moto) {
    const districtMoto = payload.district_moto.trim().toUpperCase();

    // Relire à chaque fois pour éviter les effets de bord
    const dataM = sheet.getDataRange().getValues();
    let rowM = -1;
    for (let i = 1; i < dataM.length; i++) {
      const rr = String(dataM[i][0] || '').trim().toUpperCase();
      const dd = String(dataM[i][1] || '').trim().toUpperCase();
      if (rr === region && dd === districtMoto) {
        rowM = i + 1;
        break;
      }
    }
    if (rowM === -1) {
      messages.push('❌ District moto "' + payload.district_moto + '" non trouvé (région="' + region + '").');
      hasError = true;
    } else {
      sheet.getRange(rowM, 4).setValue('M');  // colonne D
      messages.push('✓ Moto → ligne ' + rowM + ', col D (district: ' + payload.district_moto + ')');
    }
  }

  return {
    status:  hasError ? 'error' : 'ok',
    message: messages.join(' | ') || 'Aucune modification effectuée.'
  };
}

// ============================================================
//  Lecture des données existantes d'un district
// ============================================================
function getDistrictData(region, district) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DISTRICT);
  if (!sheet) return { status: 'error', message: 'Feuille introuvable.' };

  const data     = sheet.getDataRange().getValues();
  const r        = (region   || '').trim().toUpperCase();
  const d        = (district || '').trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    const rr = String(data[i][0] || '').trim().toUpperCase();
    const dd = String(data[i][1] || '').trim().toUpperCase();
    if (rr === r && dd === d) {
      return {
        status: 'ok',
        region:   data[i][0],
        district: data[i][1],
        riposte:  data[i][2],
        nb_csb:   data[i][3],
        nb_fkt:   data[i][4],
        nb_ac:    data[i][5],
        voiture:  data[i][6],
        moto:     data[i][7]
      };
    }
  }

  return { status: 'empty', message: 'Aucune donnée enregistrée pour ce district.' };
}

// ============================================================
//  Lecture des données existantes d'une région
//  Retourne toutes les lignes de la région (voiture + moto
//  peuvent être sur des lignes / districts différents).
//  Si district est fourni, retourne uniquement cette ligne.
// ============================================================
function getRegionData(region, district) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_REGION);
  if (!sheet) return { status: 'error', message: 'Feuille introuvable.' };

  const data = sheet.getDataRange().getValues();
  const r    = (region   || '').trim().toUpperCase();
  const d    = (district || '').trim().toUpperCase();

  // Si un district précis est demandé
  if (d) {
    for (let i = 1; i < data.length; i++) {
      const rr = String(data[i][0] || '').trim().toUpperCase();
      const dd = String(data[i][1] || '').trim().toUpperCase();
      if (rr === r && dd === d) {
        return {
          status:   'ok',
          region:   data[i][0],
          district: data[i][1],
          voiture:  data[i][2],
          moto:     data[i][3]
        };
      }
    }
    return { status: 'empty', message: 'Aucune donnée pour ce district dans la région.' };
  }

  // Sinon, retourner toutes les lignes de la région
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const rr = String(data[i][0] || '').trim().toUpperCase();
    if (rr === r) {
      rows.push({
        district: data[i][1],
        voiture:  data[i][2],
        moto:     data[i][3]
      });
    }
  }

  if (rows.length === 0) {
    return { status: 'empty', message: 'Aucune donnée enregistrée pour cette région.' };
  }

  return { status: 'ok', region: region, rows: rows };
}

// ============================================================
//  Sauvegarde acteurs district dans "Acteurs District"
//  Colonnes : A=Région  B=District  C=Nom et Prénom
//             D=Poste   E=CIN       F=Num M'vola
//
//  Stratégie : supprimer toutes les lignes existantes pour
//  cette région+district, puis réécrire les nouvelles lignes.
//  payload.acteurs = [{ nom, poste, cin, mvola }, ...]
// ============================================================
function saveActeursDistrict(payload) {
  // Dans le backend, au début de saveActeursDistrict :
console.log('saveActeursDistrict reçu:', JSON.stringify(payload));
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACTEURS_DISTRICT);
  if (!sheet) {
    return { status: 'error', message: 'Feuille "' + SHEET_ACTEURS_DISTRICT + '" introuvable.' };
  }

  const region   = (payload.region   || '').trim().toUpperCase();
  const district = (payload.district || '').trim().toUpperCase();
  const acteurs  = payload.acteurs || [];

  // Si aucun acteur, on ne fait rien (pas d'erreur)
  if (acteurs.length === 0) {
    return { status: 'ok', message: 'Aucun acteur à enregistrer.' };
  }

  // S'assurer que la feuille a au moins une ligne d'en-tête
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Région', 'District', 'Nom et Prénom', 'Poste', 'CIN', 'Num M\'vola']);
  }

  // Supprimer les lignes existantes pour cette région+district (de bas en haut)
  const allData = sheet.getDataRange().getValues();
  for (let i = allData.length - 1; i >= 1; i--) {
    const rr = String(allData[i][0] || '').trim().toUpperCase();
    const dd = String(allData[i][1] || '').trim().toUpperCase();
    if (rr === region && dd === district) {
      sheet.deleteRow(i + 1);
    }
  }

  // Ajouter les nouvelles lignes via appendRow (plus sûr que getLastRow+getRange)
  acteurs.forEach(function(a) {
    sheet.appendRow([
      payload.region   || '',
      payload.district || '',
      a.nom    || '',
      a.poste  || '',
      a.cin    || '',
      a.mvola  || ''
    ]);
  });

  return {
    status:  'ok',
    message: acteurs.length + ' acteur(s) district enregistré(s) pour ' + payload.district + '.'
  };
}

// ============================================================
//  Sauvegarde acteurs région dans "Acteurs Région"
//  Colonnes : A=Région  B=Nom et Prénom  C=Poste
//             D=CIN     E=Num M'vola
//
//  payload.acteurs = [{ nom, poste, cin, mvola }, ...]
// ============================================================
function saveActeursRegion(payload) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACTEURS_REGION);
  if (!sheet) {
    return { status: 'error', message: 'Feuille "' + SHEET_ACTEURS_REGION + '" introuvable.' };
  }

  const region  = (payload.region || '').trim().toUpperCase();
  const acteurs = payload.acteurs || [];

  // S'assurer que la feuille a au moins une ligne d'en-tête
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Région', 'Nom et Prénom', 'Poste', 'CIN', 'Num M\'vola']);
  }

  // Supprimer les lignes existantes pour cette région (de bas en haut)
  const allData = sheet.getDataRange().getValues();
  for (let i = allData.length - 1; i >= 1; i--) {
    const rr = String(allData[i][0] || '').trim().toUpperCase();
    if (rr === region) {
      sheet.deleteRow(i + 1);
    }
  }

  // Ajouter les nouvelles lignes via appendRow
  acteurs.forEach(function(a) {
    sheet.appendRow([
      payload.region || '',
      a.nom   || '',
      a.poste || '',
      a.cin   || '',
      a.mvola || ''
    ]);
  });

  return {
    status:  'ok',
    message: acteurs.length + ' acteur(s) région enregistré(s) pour ' + payload.region + '.'
  };
}

// ============================================================
//  Lecture acteurs district
// ============================================================
function getActeursDistrict(region, district) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACTEURS_DISTRICT);
  if (!sheet) return { status: 'error', message: 'Feuille introuvable.' };

  const data = sheet.getDataRange().getValues();
  const r    = (region   || '').trim().toUpperCase();
  const d    = (district || '').trim().toUpperCase();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const rr = String(data[i][0] || '').trim().toUpperCase();
    const dd = String(data[i][1] || '').trim().toUpperCase();
    if (rr === r && dd === d) {
      rows.push({
        nom:   data[i][2],
        poste: data[i][3],
        cin:   data[i][4],
        mvola: data[i][5]
      });
    }
  }

  if (rows.length === 0) {
    return { status: 'empty', message: 'Aucun acteur enregistré pour ce district.' };
  }
  return { status: 'ok', acteurs: rows };
}

// ============================================================
//  Lecture acteurs région
// ============================================================
function getActeursRegion(region) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACTEURS_REGION);
  if (!sheet) return { status: 'error', message: 'Feuille introuvable.' };

  const data = sheet.getDataRange().getValues();
  const r    = (region || '').trim().toUpperCase();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const rr = String(data[i][0] || '').trim().toUpperCase();
    if (rr === r) {
      rows.push({
        nom:   data[i][1],
        poste: data[i][2],
        cin:   data[i][3],
        mvola: data[i][4]
      });
    }
  }

  if (rows.length === 0) {
    return { status: 'empty', message: 'Aucun acteur enregistré pour cette région.' };
  }
  return { status: 'ok', acteurs: rows };
}

// ============================================================
//  FONCTION DE TEST – à exécuter directement dans Apps Script
//  pour vérifier que les feuilles sont accessibles et que
//  l'écriture fonctionne. Supprimer après validation.
// ============================================================
function testActeurs() {
  const results = [];

  // Test district
  const r1 = saveActeursDistrict({
    region:   'ANALAMANGA',
    district: 'AMBOHIDRATRIMO',
    acteurs:  [
      { nom: 'TEST Nom', poste: 'TEST Poste', cin: '000000000', mvola: '034 00 000 00' }
    ]
  });
  results.push('District: ' + JSON.stringify(r1));

  // Test région
  const r2 = saveActeursRegion({
    region:  'ANALAMANGA',
    acteurs: [
      { nom: 'TEST Nom Région', poste: 'TEST Poste Région', cin: '111111111', mvola: '034 11 111 11' }
    ]
  });
  results.push('Région: ' + JSON.stringify(r2));

  Logger.log(results.join('\n'));
  return results;
}

// ══════════════════════════════════════════════════════
//  Nom de la feuille cible
// ══════════════════════════════════════════════════════
const SHEET_ACTEURS_COMMUNAUTAIRES = 'Acteurs Communautaires';

// ══════════════════════════════════════════════════════
//  Sauvegarde des acteurs communautaires
//  Colonnes :
//    A = Région
//    B = District
//    C = CSB
//    D = Fokontany
//    E = Nom et Prénom
//    F = Poste
//    G = CIN
//    H = Num M'vola
//
//  Stratégie : supprimer toutes les lignes existantes pour
//  cette région+district, puis réécrire les nouvelles lignes.
//
//  payload = {
//    action: 'saveActeursCommunautaires',
//    region, district,
//    acteurs: [{ csb, fokontany, nom, poste, cin, mvola }, …]
//  }
// ══════════════════════════════════════════════════════
function saveActeursCommunautaires(payload) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACTEURS_COMMUNAUTAIRES);

  if (!sheet) {
    return {
      status:  'error',
      message: 'Feuille "' + SHEET_ACTEURS_COMMUNAUTAIRES + '" introuvable. Créez-la d\'abord dans le Google Sheet.'
    };
  }

  const region   = (payload.region   || '').trim().toUpperCase();
  const district = (payload.district || '').trim().toUpperCase();
  const acteurs  = payload.acteurs   || [];

  // Créer l'en-tête si la feuille est vide
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Région', 'District', 'CSB', 'Fokontany', 'Nom et Prénom', 'Poste', 'CIN', "Num M'vola"]);
  }

  // Supprimer les lignes existantes pour cette région+district (de bas en haut)
  const allData = sheet.getDataRange().getValues();
  for (let i = allData.length - 1; i >= 1; i--) {
    const rr = String(allData[i][0] || '').trim().toUpperCase();
    const dd = String(allData[i][1] || '').trim().toUpperCase();
    if (rr === region && dd === district) {
      sheet.deleteRow(i + 1);
    }
  }

  // Ajouter les nouvelles lignes
  acteurs.forEach(function(a) {
    sheet.appendRow([
      payload.region   || '',
      payload.district || '',
      a.csb        || '',
      a.fokontany  || '',
      a.nom        || '',
      a.poste      || '',
      a.cin        || '',
      a.mvola      || ''
    ]);
  });

  return {
    status:  'ok',
    message: acteurs.length + ' acteur(s) communautaire(s) enregistré(s) pour ' + payload.district + '.'
  };
}

// ══════════════════════════════════════════════════════
//  Lecture des acteurs communautaires
//  Retourne tous les acteurs pour une région+district
// ══════════════════════════════════════════════════════
function getActeursCommunautaires(region, district) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACTEURS_COMMUNAUTAIRES);

  if (!sheet) {
    return {
      status:  'error',
      message: 'Feuille "' + SHEET_ACTEURS_COMMUNAUTAIRES + '" introuvable.'
    };
  }

  const data = sheet.getDataRange().getValues();
  const r    = (region   || '').trim().toUpperCase();
  const d    = (district || '').trim().toUpperCase();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const rr = String(data[i][0] || '').trim().toUpperCase();
    const dd = String(data[i][1] || '').trim().toUpperCase();
    if (rr === r && dd === d) {
      rows.push({
        csb:       data[i][2],
        fokontany: data[i][3],
        nom:       data[i][4],
        poste:     data[i][5],
        cin:       data[i][6],
        mvola:     data[i][7]
      });
    }
  }

  if (rows.length === 0) {
    return {
      status:  'empty',
      message: 'Aucun acteur communautaire enregistré pour ce district.'
    };
  }

  return { status: 'ok', acteurs: rows };
}
