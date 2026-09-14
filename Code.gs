/**
 * Control d'Assistència — backend (Google Apps Script)
 *
 * Fulls de càlcul que fa servir aquest script:
 *  - Config     : paràmetres del curs (dies lectius, pesos de cada franja, llindar d'alerta)
 *  - Alumnes    : llistat d'alumnes (Classe | Alumne)
 *  - Professors : emails autoritzats a entrar a la web (llista blanca)
 *  - Registre   : una fila per alumne i dia amb alguna marca (falta o retard)
 *
 * Executa la funció setup() UNA SOLA VEGADA des de l'editor d'Apps Script
 * per crear aquests fulls amb les capçaleres i valors per defecte.
 */

const SHEET_CONFIG = 'Config';
const SHEET_ALUMNES = 'Alumnes';
const SHEET_PROFESSORS = 'Professors';
const SHEET_REGISTRE = 'Registre';

/**
 * Demanar la Spreadsheet activa a Google té un cost real de xarxa cada vegada
 * que es fa. Aquesta funció la demana un sol cop per execució i la reutilitza,
 * en comptes que cada funció la torni a demanar pel seu compte.
 */
let _ss = null;
function ss_() {
  if (!_ss) _ss = SpreadsheetApp.getActiveSpreadsheet();
  return _ss;
}

/**
 * Numeret que es guarda de forma permanent (PropertiesService) i que s'incrementa
 * cada vegada que es guarda un canvi a "Registre". getResum() l'inclou en la clau
 * del seu caché, així que en incrementar-lo, qualsevol resultat de Resum guardat
 * en caché queda automàticament invalidat (obsolet) sense haver-lo d'esborrar.
 */
function versioRegistre_() {
  return PropertiesService.getScriptProperties().getProperty('REGISTRE_VERSIO') || '0';
}
function incrementarVersioRegistre_() {
  const props = PropertiesService.getScriptProperties();
  const actual = Number(props.getProperty('REGISTRE_VERSIO') || '0');
  props.setProperty('REGISTRE_VERSIO', String(actual + 1));
}

function doGet() {
  if (!usuariAutoritzat_()) {
    const emailDetectat = emailActual_() || '(cap email detectat — potser cal compartir la Sheet amb aquest usuari)';
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:24px;">' +
      '🔒 Accés denegat. Aquesta aplicació és només per a professors del centre.<br>' +
      'Si creus que hauries de tenir accés, demana que afegeixin el teu email a la pestanya "Professors".<br><br>' +
      '<b>Email detectat pel sistema:</b> ' + emailDetectat +
      '</p>'
    ).setTitle("Control d'Assistència");
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle("Control d'Assistència")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function emailActual_() {
  try {
    return (Session.getActiveUser().getEmail() || '').toLowerCase();
  } catch (e) {
    return '';
  }
}

/**
 * Si la pestanya "Professors" no té cap email, es deixa entrar a tothom
 * (per no bloquejar-vos abans de configurar-la). En quant hi hagi almenys
 * un email a la llista, només aquests podran entrar-hi.
 */
function usuariAutoritzat_() {
  const sh = ss_().getSheetByName(SHEET_PROFESSORS);
  if (!sh) return true;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return true;

  const emails = sh.getRange(2, 1, lastRow - 1, 1).getValues()
    .map((r) => String(r[0] || '').toLowerCase().trim())
    .filter(String);

  if (emails.length === 0) return true;

  return emails.indexOf(emailActual_()) !== -1;
}

/**
 * Crea (si no existeixen) els fulls Config, Alumnes, Professors i Registre.
 * Es pot tornar a executar sense por: no esborra dades ja introduïdes.
 */
function setup() {
  const ss = ss_();

  let cfg = ss.getSheetByName(SHEET_CONFIG);
  if (!cfg) {
    cfg = ss.insertSheet(SHEET_CONFIG);
    cfg.getRange(1, 1, 5, 3).setValues([
      ['Clau', 'Valor', 'Descripció'],
      ['TOTAL_DIES_LECTIUS', 177, 'Dies lectius totals del curs (per calcular el %)'],
      ['PES_MATI', 2 / 3, 'Quina part d\'un dia compta una falta de matí'],
      ['PES_TARDA', 1 / 3, 'Quina part d\'un dia compta una falta de tarda (matí+tarda = dia complet)'],
      ['UMBRAL_ALERTA_PCT', 10, 'A partir de quin % es marca l\'alumne en alerta al Resum'],
    ]);
    cfg.setFrozenRows(1);
    cfg.autoResizeColumns(1, 3);
  }

  let al = ss.getSheetByName(SHEET_ALUMNES);
  if (!al) {
    al = ss.insertSheet(SHEET_ALUMNES);
    al.getRange(1, 1, 4, 3).setValues([
      ['Classe', 'Alumne', 'No promocionar (repetidor)'],
      ['P3A', 'Alumne Exemple 1', false],
      ['P3A', 'Alumna Exemple 2', false],
      ['P4B', 'Alumne Exemple 3', false],
    ]);
    al.setFrozenRows(1);
    al.autoResizeColumns(1, 3);
  } else if (al.getRange(1, 3).getValue() === '') {
    // Migració d'un full antic sense la columna de repetidors.
    al.getRange(1, 3).setValue('No promocionar (repetidor)');
  }

  let promo = ss.getSheetByName('Promocio');
  if (!promo) {
    promo = ss.insertSheet('Promocio');
    promo.getRange(1, 1, 10, 2).setValues([
      ['Classe actual', 'Classe nova (buit = surt del centre, ex. 6è que es gradua)'],
      ['P3A', 'P4A'], ['P3B', 'P4B'],
      ['P4A', 'P5A'], ['P4B', 'P5B'],
      ['P5A', '1rA'], ['P5B', '1rB'],
      ['...', '...'],
      ['6èA', ''], ['6èB', ''],
    ]);
    promo.setFrozenRows(1);
    promo.autoResizeColumns(1, 2);
  }

  let prof = ss.getSheetByName(SHEET_PROFESSORS);
  if (!prof) {
    prof = ss.insertSheet(SHEET_PROFESSORS);
    prof.getRange(1, 1, 1, 2).setValues([['Email', 'Nom (opcional)']]);
    const meuEmail = emailActual_();
    if (meuEmail) {
      prof.getRange(2, 1, 1, 2).setValues([[meuEmail, 'Afegit automàticament en fer setup()']]);
    }
    prof.setFrozenRows(1);
    prof.autoResizeColumns(1, 2);
  }

  let reg = ss.getSheetByName(SHEET_REGISTRE);
  if (!reg) {
    reg = ss.insertSheet(SHEET_REGISTRE);
    reg.getRange(1, 1, 1, 11).setValues([
      ['Data', 'Classe', 'Alumne', 'Matí', 'Tarda', 'Justificada matí', 'Justificada tarda', 'Retard matí', 'Retard tarda', 'Actualitzat per', 'Última actualització'],
    ]);
    reg.setFrozenRows(1);
    reg.autoResizeColumns(1, 11);
  }

  SpreadsheetApp.getUi().alert(
    'Fulls creats correctament.\n\n' +
    'Recorda: afegeix els emails dels professors autoritzats a la pestanya "Professors" ' +
    'abans de compartir l\'enllaç amb tothom.\n\n' +
    'Ja pots anar a Implementar > Nova implementació.'
  );
}

/**
 * NOMÉS PER FER PROVES DE RENDIMENT: substitueix el contingut de la pestanya
 * "Alumnes" per dades falses (uns quants centenars d'alumnes repartits en
 * moltes classes), per comprovar que la web va igual de ràpida amb un
 * col·legi sencer que amb 3 alumnes d'exemple.
 *
 * Executa-la manualment des de l'editor quan vulguis fer la prova, i
 * torna a executar setup() (o esborra la pestanya "Alumnes" a mà) per
 * tornar a les dades reals després.
 */
function generarAlumnesDeProva() {
  const cursos = ['P3', 'P4', 'P5', '1r', '2n', '3r', '4t', '5è', '6è'];
  const grups = ['A', 'B'];
  const alumnesPerClasse = 25;

  const files = [];
  cursos.forEach((curs) => {
    grups.forEach((grup) => {
      const classe = curs + grup;
      for (let i = 1; i <= alumnesPerClasse; i++) {
        files.push([classe, `Alumne ${classe}-${i}`]);
      }
    });
  });

  const sh = ss_().getSheetByName(SHEET_ALUMNES);
  const filesActuals = Math.max(sh.getLastRow() - 1, 0);
  if (filesActuals > 0) sh.getRange(2, 1, filesActuals, 2).clearContent();
  sh.getRange(2, 1, files.length, 2).setValues(files);

  SpreadsheetApp.getUi().alert(
    `Generats ${files.length} alumnes de prova en ${cursos.length * grups.length} classes.`
  );
}

/**
 * NOMÉS PER FER PROVES: omple la pestanya "Registre" amb faltes i retards
 * aleatoris per a tots els alumnes, repartits en els últims `diesEnrere` dies
 * lectius (caps de setmana exclosos). Serveix per veure el Resum, els %, les
 * alertes i el cercador amb dades que semblin reals, sense haver de passar
 * llista dia a dia a mà.
 *
 * Substitueix TOT el contingut actual de "Registre". Executa-la manualment
 * des de l'editor. Per tornar a buit, esborra el contingut de "Registre" a mà.
 */
function generarRegistreDeProva(diesEnrere) {
  diesEnrere = diesEnrere || 60;

  const alSh = ss_().getSheetByName(SHEET_ALUMNES);
  const alValues = alSh.getRange(2, 1, Math.max(alSh.getLastRow() - 1, 0), 2).getValues()
    .filter((r) => r[0] && r[1]);

  if (!alValues.length) {
    SpreadsheetApp.getUi().alert('No hi ha alumnes a la pestanya "Alumnes". Genera\'ls primer amb generarAlumnesDeProva().');
    return;
  }

  const dates = [];
  const cursor = new Date();
  cursor.setHours(0, 0, 0, 0);
  while (dates.length < diesEnrere) {
    const diaSetmana = cursor.getDay();
    if (diaSetmana !== 0 && diaSetmana !== 6) {
      dates.push(Utilities.formatDate(cursor, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    }
    cursor.setDate(cursor.getDate() - 1);
  }

  let email;
  try {
    email = Session.getActiveUser().getEmail() || 'prova';
  } catch (e) {
    email = 'prova';
  }
  const ara = new Date();

  const files = [];
  alValues.forEach(([classe, alumne]) => {
    dates.forEach((data) => {
      const rand = Math.random();
      let mati = false, tarda = false, retardMati = false, retardTarda = false;
      let justificadaMati = false, justificadaTarda = false;

      if (rand < 0.03) {
        mati = true; tarda = true; // dia complet
      } else if (rand < 0.07) {
        mati = true;
      } else if (rand < 0.09) {
        tarda = true;
      } else if (rand < 0.12) {
        retardMati = true;
      } else if (rand < 0.13) {
        retardTarda = true;
      }

      if (mati) justificadaMati = Math.random() < 0.5;
      if (tarda) justificadaTarda = Math.random() < 0.5;

      if (mati || tarda || retardMati || retardTarda) {
        files.push([data, classe, alumne, mati, tarda, justificadaMati, justificadaTarda, retardMati, retardTarda, email, ara]);
      }
    });
  });

  const sh = ss_().getSheetByName(SHEET_REGISTRE);
  const filesActuals = Math.max(sh.getLastRow() - 1, 0);
  if (filesActuals > 0) sh.getRange(2, 1, filesActuals, 11).clearContent();
  if (files.length) {
    sh.getRange(2, 1, files.length, 11).setValues(files);
  }

  SpreadsheetApp.getUi().alert(
    `Generats ${files.length} registres de falta/retard de prova, ` +
    `repartits en ${dates.length} dies lectius (caps de setmana exclosos) i ${alValues.length} alumnes.`
  );
}

/**
 * Promoció de curs de final d'any: mou cada alumne de la seva classe actual
 * a la següent, seguint el mapa que hi hagi escrit a la pestanya "Promocio"
 * (Classe actual -> Classe nova; deixa "Classe nova" buida per als cursos
 * que es graduen i surten del centre, com ara 6è).
 *
 * Els alumnes marcats amb la casella "No promocionar (repetidor)" a la
 * pestanya "Alumnes" es queden a la mateixa classe i no es toquen.
 *
 * Executa-la manualment UN SOL COP quan comenci el curs nou. Abans de fer-ho:
 *  1. Revisa/omple bé la pestanya "Promocio" amb els noms reals de les teves classes.
 *  2. Marca la casella de repetidors als alumnes que calgui.
 *  3. Si vols conservar l'històric de faltes de l'any que acaba, copia el
 *     contingut de "Registre" a una altra pestanya (ex. "Registre 2025-26")
 *     abans d'esborrar-lo, i actualitza TOTAL_DIES_LECTIUS a "Config" pel curs nou.
 */
function promocionarCurs() {
  const promoSh = ss_().getSheetByName('Promocio');
  if (!promoSh) {
    SpreadsheetApp.getUi().alert('No trobo la pestanya "Promocio". Executa setup() primer.');
    return;
  }

  const mapa = {};
  promoSh.getRange(2, 1, Math.max(promoSh.getLastRow() - 1, 0), 2).getValues().forEach(([actual, nova]) => {
    if (actual) mapa[String(actual).trim()] = String(nova || '').trim();
  });

  const alSh = ss_().getSheetByName(SHEET_ALUMNES);
  const numFiles = Math.max(alSh.getLastRow() - 1, 0);
  if (!numFiles) {
    SpreadsheetApp.getUi().alert('No hi ha alumnes a la pestanya "Alumnes".');
    return;
  }

  const dades = alSh.getRange(2, 1, numFiles, 3).getValues();
  let promocionats = 0;
  let repetidors = 0;
  let graduats = 0;
  let senseMapa = 0;
  const filesPerEliminar = [];

  for (let i = 0; i < dades.length; i++) {
    const [classe, alumne, noPromocionar] = dades[i];
    if (!alumne) continue;

    if (noPromocionar === true) {
      repetidors++;
      continue;
    }

    if (!(classe in mapa)) {
      senseMapa++;
      continue;
    }

    const classeNova = mapa[classe];
    if (!classeNova) {
      filesPerEliminar.push(i + 2); // es gradua / surt del centre
      graduats++;
    } else {
      alSh.getRange(i + 2, 1).setValue(classeNova);
      promocionats++;
    }
  }

  filesPerEliminar
    .sort((a, b) => b - a)
    .forEach((fila) => alSh.deleteRow(fila));

  // Reinicia les caselles de "repetidor" de cara al curs vinent.
  const numFilesFinal = Math.max(alSh.getLastRow() - 1, 0);
  if (numFilesFinal) {
    alSh.getRange(2, 3, numFilesFinal, 1).setValue(false);
  }

  SpreadsheetApp.getUi().alert(
    'Promoció completada:\n\n' +
    `• ${promocionats} alumnes promocionats de classe\n` +
    `• ${repetidors} repetidors mantinguts a la mateixa classe\n` +
    `• ${graduats} alumnes graduats/eliminats (classe nova buida al mapa)\n` +
    (senseMapa ? `• ⚠️ ${senseMapa} alumnes amb una classe que no és a la pestanya "Promocio" (no s'han tocat)\n` : '') +
    '\nRevisa la pestanya "Alumnes" per confirmar que tot ha quedat bé.'
  );
}

/**
 * NOMÉS PER MESURAR RENDIMENT: crida getResum() (amb totes les classes, el cas
 * més pesat) i mostra quant triga en mil·lisegons. Útil per comprovar si val
 * la pena optimitzar-lo amb el volum de dades que tingueu.
 */
function mesurarRendiment() {
  const inici = new Date().getTime();
  const resultat = getResum('');
  const fi = new Date().getTime();

  SpreadsheetApp.getUi().alert(
    `getResum() ha trigat ${fi - inici} ms\n` +
    `(${resultat.length} alumnes processats)`
  );
}

function getConfig_() {
  const sh = ss_().getSheetByName(SHEET_CONFIG);
  const values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), 2).getValues();
  const cfg = {};
  values.forEach(([clau, valor]) => {
    if (clau) cfg[clau] = valor;
  });
  return cfg;
}

function getClasses() {
  const sh = ss_().getSheetByName(SHEET_ALUMNES);
  const values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), 1).getValues();
  const set = new Set(values.map((r) => r[0]).filter(String));
  return Array.from(set).sort();
}

function clauDia_(data, alumne) {
  return data + '||' + alumne;
}

function nomDiaSetmana_(dataISO) {
  const dies = ['Diumenge', 'Dilluns', 'Dimarts', 'Dimecres', 'Dijous', 'Divendres', 'Dissabte'];
  const parts = dataISO.split('-').map(Number);
  const d = new Date(parts[0], parts[1] - 1, parts[2]);
  return { nom: dies[d.getDay()], capDeSetmana: d.getDay() === 0 || d.getDay() === 6 };
}

/**
 * Retorna tots els alumnes d'una classe amb les marques que ja existeixin per a `data`
 * (si encara no s'ha passat llista aquell dia, totes les marques surten a false).
 */
function getGraella(classe, data) {
  const alSh = ss_().getSheetByName(SHEET_ALUMNES);
  const alValues = alSh.getRange(2, 1, Math.max(alSh.getLastRow() - 1, 0), 2).getValues();
  const alumnes = alValues
    .filter((r) => r[0] === classe && r[1])
    .map((r) => r[1])
    .sort((a, b) => a.localeCompare(b, 'ca'));

  const marques = {};
  const regSh = ss_().getSheetByName(SHEET_REGISTRE);
  const lastRow = regSh.getLastRow();
  if (lastRow >= 2) {
    const regValues = regSh.getRange(2, 1, lastRow - 1, 9).getValues();
    regValues.forEach(([rData, rClasse, rAlumne, mati, tarda, justificadaMati, justificadaTarda, retardMati, retardTarda]) => {
      if (rClasse === classe && formatDataISO_(rData) === data) {
        marques[rAlumne] = {
          mati: mati === true,
          tarda: tarda === true,
          justificadaMati: justificadaMati === true,
          justificadaTarda: justificadaTarda === true,
          retardMati: retardMati === true,
          retardTarda: retardTarda === true,
        };
      }
    });
  }

  const diaInfo = nomDiaSetmana_(data);

  return {
    diaSetmana: diaInfo.nom,
    capDeSetmana: diaInfo.capDeSetmana,
    alumnes: alumnes.map((alumne) => {
      const m = marques[alumne] || { mati: false, tarda: false, justificadaMati: false, justificadaTarda: false, retardMati: false, retardTarda: false };
      return {
        alumne,
        mati: m.mati,
        tarda: m.tarda,
        justificadaMati: m.justificadaMati,
        justificadaTarda: m.justificadaTarda,
        retardMati: m.retardMati,
        retardTarda: m.retardTarda,
      };
    }),
  };
}

/**
 * files = [{ alumne, mati, tarda, justificadaMati, justificadaTarda, retardMati, retardTarda }, ...]
 * per a una classe i data concretes. Actualitza, crea o esborra la fila corresponent a cada alumne segons calgui.
 */
function guardarGraella(classe, data, files) {
  if (!classe || !data || !files) throw new Error('Falten dades.');

  const sh = ss_().getSheetByName(SHEET_REGISTRE);
  const lastRow = sh.getLastRow();
  const numRows = Math.max(lastRow - 1, 0);
  const values = numRows ? sh.getRange(2, 1, numRows, 11).getValues() : [];

  const indexPerClau = {};
  values.forEach((row, i) => {
    if (row[1] === classe && formatDataISO_(row[0]) === data) {
      indexPerClau[clauDia_(data, row[2])] = i + 2; // fila real al full
    }
  });

  let email;
  try {
    email = Session.getActiveUser().getEmail() || 'desconegut';
  } catch (e) {
    email = 'desconegut';
  }
  const ara = new Date();

  const filesPerEsborrar = [];

  files.forEach((f) => {
    const teAlgunaMarca = f.mati || f.tarda || f.retardMati || f.retardTarda;
    const clau = clauDia_(data, f.alumne);
    const filaExistent = indexPerClau[clau];

    if (!teAlgunaMarca) {
      if (filaExistent) filesPerEsborrar.push(filaExistent);
      return;
    }

    const novaFila = [
      data, classe, f.alumne,
      !!f.mati, !!f.tarda,
      !!f.justificadaMati, !!f.justificadaTarda,
      !!f.retardMati, !!f.retardTarda,
      email, ara,
    ];

    if (filaExistent) {
      sh.getRange(filaExistent, 1, 1, 11).setValues([novaFila]);
    } else {
      sh.appendRow(novaFila);
    }
  });

  filesPerEsborrar
    .sort((a, b) => b - a)
    .forEach((fila) => sh.deleteRow(fila));

  incrementarVersioRegistre_();

  return true;
}

/**
 * El cache d'Apps Script només accepta ~100KB per clau. Amb molts alumnes,
 * el JSON del Resum pot superar-ho, així que el comprimim (gzip + base64)
 * abans de guardar-lo, i el desfem en llegir-lo.
 */
function cacheGet_(cache, key) {
  const base64 = cache.get(key);
  if (!base64) return null;
  try {
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, 'application/x-gzip');
    return Utilities.ungzip(blob).getDataAsString();
  } catch (e) {
    return null;
  }
}
function cachePut_(cache, key, str, ttlSec) {
  try {
    const gzipBlob = Utilities.gzip(Utilities.newBlob(str));
    const base64 = Utilities.base64Encode(gzipBlob.getBytes());
    if (base64.length <= 100000) cache.put(key, base64, ttlSec);
  } catch (e) {
    // Si tot i comprimir-ho segueix sent massa gran, simplement no es guarda en caché.
  }
}

function getResum(classeFiltre) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'resum_v' + versioRegistre_() + '_' + (classeFiltre || '__totes__');
  const cacheValue = cacheGet_(cache, cacheKey);
  if (cacheValue) return JSON.parse(cacheValue);

  const cfg = getConfig_();
  const pesMati = Number(cfg.PES_MATI) || 0;
  const pesTarda = Number(cfg.PES_TARDA) || 0;
  const totalDies = Number(cfg.TOTAL_DIES_LECTIUS) || 0;
  const umbral = Number(cfg.UMBRAL_ALERTA_PCT) || 10;

  const alSh = ss_().getSheetByName(SHEET_ALUMNES);
  const alValues = alSh.getRange(2, 1, Math.max(alSh.getLastRow() - 1, 0), 2).getValues();

  const buit_ = () => ({ diesFalta: 0, diesJustificats: 0, retardsMati: 0, retardsTarda: 0, perDiaSetmana: [0, 0, 0, 0, 0, 0, 0] });
  const stats = {};
  const clau = (classe, alumne) => classe + ' ||| ' + alumne;

  alValues.forEach(([classe, alumne]) => {
    if (!alumne) return;
    if (classeFiltre && classe !== classeFiltre) return;
    stats[clau(classe, alumne)] = Object.assign({ classe, alumne }, buit_());
  });

  const regSh = ss_().getSheetByName(SHEET_REGISTRE);
  const lastRow = regSh.getLastRow();
  if (lastRow >= 2) {
    const regValues = regSh.getRange(2, 1, lastRow - 1, 9).getValues();
    regValues.forEach(([data, classe, alumne, mati, tarda, justificadaMati, justificadaTarda, retardMati, retardTarda]) => {
      if (!alumne) return;
      if (classeFiltre && classe !== classeFiltre) return;

      const k = clau(classe, alumne);
      if (!stats[k]) {
        stats[k] = Object.assign({ classe, alumne }, buit_());
      }

      // Si per algun motiu (ex. edició manual de la Sheet) una franja té marcada
      // alhora Falta i Retard, dona prioritat a la Falta i ignora el Retard,
      // per no comptar dues vegades la mateixa incidència.
      if (retardMati === true && mati !== true) stats[k].retardsMati += 1;
      if (retardTarda === true && tarda !== true) stats[k].retardsTarda += 1;

      if (mati === true) {
        if (justificadaMati === true) stats[k].diesJustificats += pesMati;
        else stats[k].diesFalta += pesMati;
      }
      if (tarda === true) {
        if (justificadaTarda === true) stats[k].diesJustificats += pesTarda;
        else stats[k].diesFalta += pesTarda;
      }

      if (mati === true || tarda === true) {
        const diaSetm = new Date(formatDataISO_(data)).getUTCDay();
        stats[k].perDiaSetmana[diaSetm] += 1;
      }
    });
  }

  const NOMS_DIA = ['Diumenge', 'Dilluns', 'Dimarts', 'Dimecres', 'Dijous', 'Divendres', 'Dissabte'];

  const resultat = Object.values(stats)
    .map((s) => {
      const total = s.diesFalta + s.diesJustificats;
      const pctNoJustificat = totalDies ? (s.diesFalta / totalDies) * 100 : 0;
      const pctJustificat = totalDies ? (s.diesJustificats / totalDies) * 100 : 0;
      const pctTotal = pctNoJustificat + pctJustificat;

      // Ràtio de justificació: quin % de les SEVES pròpies faltes estan justificades
      // (diferent de percentatgeJustificat, que és sobre el total de dies del curs).
      const ratioJustificacio = total > 0 ? arrodonir_((s.diesJustificats / total) * 100) : null;

      // Desglossament de tots els dies amb alguna falta, ordenats de més a menys freqüent.
      const diaMesFrequent = s.perDiaSetmana
        .map((n, i) => ({ nom: NOMS_DIA[i], n }))
        .filter((d) => d.n > 0)
        .sort((a, b) => b.n - a.n)
        .map((d) => `${d.nom} (${d.n})`)
        .join(', ') || '—';

      // Escala progressiva en 3 esglaons, relativa al llindar d'alerta configurat:
      // per sota de la meitat del llindar = verd, entre la meitat i el llindar = ambre,
      // a partir del llindar = vermell. Així el canvi de color comença abans del límit exacte.
      let nivell = 'ok';
      if (pctTotal >= umbral) nivell = 'alerta';
      else if (pctTotal >= umbral / 2) nivell = 'advertencia';

      return {
        classe: s.classe,
        alumne: s.alumne,
        diesFaltaNoJustificades: arrodonir_(s.diesFalta),
        diesJustificats: arrodonir_(s.diesJustificats),
        totalDies: arrodonir_(total),
        percentatgeNoJustificat: arrodonir_(pctNoJustificat),
        percentatgeJustificat: arrodonir_(pctJustificat),
        percentatge: arrodonir_(pctTotal),
        ratioJustificacio,
        alerta: pctTotal >= umbral,
        nivell,
        retardsMati: s.retardsMati,
        retardsTarda: s.retardsTarda,
        diaMesFrequent,
      };
    })
    .sort((a, b) => b.percentatge - a.percentatge);

  cachePut_(cache, cacheKey, JSON.stringify(resultat), 120); // 2 minuts

  return resultat;
}

function formatDataISO_(d) {
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return d;
}

function arrodonir_(n) {
  return Math.round(n * 100) / 100;
}
