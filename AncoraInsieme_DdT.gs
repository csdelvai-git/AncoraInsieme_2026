// ============================================================
//  ANCORA INSIEME — DdT Sheet Manager
//  Versione 1.0
//  
//  INSTALLAZIONE:
//  1. Apri il tuo Google Sheet
//  2. Estensioni → Apps Script
//  3. Incolla tutto questo codice
//  4. Salva (Ctrl+S)
//  5. Esegui setupFoglio() una sola volta per inizializzare
//  6. Autorizza le permissioni richieste da Google
// ============================================================

// ── CONFIGURAZIONE ──────────────────────────────────────────
// Modifica questi valori nel foglio "Config" dopo il setup,
// oppure direttamente qui prima di eseguire setupFoglio()

const CFG = {
  SOGLIA_DOPPIONE_MINUTI: 30,
  EMAIL_DDT: [],           // Verrà letto dal foglio Config
  NOME_TORNEO: "Ancora Insieme",
  LINK_SHEET: "",          // Verrà popolato automaticamente
};

// Nomi dei fogli
const FOGLI = {
  SEGNALAZIONI: "Segnalazioni_form",
  GIOCATORI:    "Giocatori",
  DA_CONFERMARE:"Da Confermare",
  VALIDATE:     "Validate",
  ESPORTATE:    "Esportate",
  CONFIG:       "Config",
};

// Colonne del foglio Segnalazioni (1-based)
const COL = {
  TIMESTAMP_INVIO:  1,  // A - quando è stato compilato il form
  DATA_GARA:        2,  // B - data fine gara
  ORA_GARA:         3,  // C - ora fine gara
  GIOCATORE_BLU1:   4,  // D
  GIOCATORE_BLU2:   5,  // E
  GOL_BLU:          6,  // F
  GOL_ROSSI:        7,  // G
  GIOCATORE_ROSSI1: 8,  // H
  GIOCATORE_ROSSI2: 9,  // I
  NOTE:            10,  // J
  SEGNALATORE:     11,  // K - chi ha segnalato
  EMAIL_SEGNALATORE:12, // L - email segnalatore
  STATO:           13,  // M - In attesa / Da confermare / Validata / Pronto
  DOPPIONE:        14,  // N - flag doppione
  VALIDATO_DA:     15,  // O - chi ha validato
  VALIDATO_IL:     16,  // P - quando
};

// Colori di stato
const COLORI = {
  IN_ATTESA:     "#FFFFFF",
  DA_CONFERMARE: "#FFF3CD",  // giallo
  VALIDATA:      "#D4EDDA",  // verde chiaro
  DOPPIONE:      "#FFE0B2",  // arancione
  HEADER:        "#1F4E8C",  // blu istituzionale
};


// ============================================================
//  SETUP INIZIALE — eseguire UNA SOLA VOLTA
// ============================================================
// Esegui in sequenza: setup1(), poi setup2(), poi setup3()
function setup1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _creaFoglioSeNonEsiste(ss, FOGLI.SEGNALAZIONI);
  _creaFoglioSeNonEsiste(ss, FOGLI.GIOCATORI);
  _creaFoglioSeNonEsiste(ss, FOGLI.DA_CONFERMARE);
  _creaFoglioSeNonEsiste(ss, FOGLI.VALIDATE);
  _creaFoglioSeNonEsiste(ss, FOGLI.ESPORTATE);
  _creaFoglioSeNonEsiste(ss, FOGLI.CONFIG);
  const def = ss.getSheetByName("Foglio1") || ss.getSheetByName("Sheet1");
  if (def) ss.deleteSheet(def);
  _setupGiocatori(ss);
  _setupDaConfermare(ss);
  _setupValidate(ss);
  _setupEsportate(ss);
  _setupConfig(ss);
  SpreadsheetApp.getUi().alert("\u2705 Passo 1/3 completato. Esegui ora setup2().");
}

function setup2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _setupSegnalazioni(ss);
  SpreadsheetApp.getUi().alert("\u2705 Passo 2/3 completato. Esegui ora setup3().");
}

function setup3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _setupTriggerForm(ss);
  SpreadsheetApp.getUi().alert("\u2705 Setup completato! Vai nel foglio Config e inserisci le email DdT.");
}

// Alias per retrocompatibilit\u00e0
function setupFoglio() { setup1(); }

function _creaFoglioSeNonEsiste(ss, nome) {
  if (!ss.getSheetByName(nome)) {
    ss.insertSheet(nome);
  }
}


// ============================================================
//  SETUP SINGOLI FOGLI
// ============================================================

function _setupSegnalazioni(ss) {
  const sh = ss.getSheetByName(FOGLI.SEGNALAZIONI);
  sh.clearFormats();

  // Header
  const headers = [
    "Timestamp Invio", "Data Gara", "Ora Gara",
    "Blu 1", "Blu 2", "Gol Blu", "Gol Rossi",
    "Rosso 1", "Rosso 2",
    "Note", "Stato", "⚠️ Doppione", "Validato da", "Validato il"
  ];
  const headerRow = sh.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setBackground(COLORI.HEADER);
  headerRow.setFontColor("#FFFFFF");
  headerRow.setFontWeight("bold");
  headerRow.setFontFamily("Arial");
  headerRow.setFontSize(10);

  // Larghezze colonne
  sh.setColumnWidth(1, 140);  // Timestamp invio
  sh.setColumnWidth(2, 100);  // Data gara
  sh.setColumnWidth(3, 80);   // Ora gara
  sh.setColumnWidth(4, 130);  // Blu1
  sh.setColumnWidth(5, 130);  // Blu2
  sh.setColumnWidth(6, 70);   // Gol Blu
  sh.setColumnWidth(7, 70);   // Gol Rossi
  sh.setColumnWidth(8, 130);  // Rosso1
  sh.setColumnWidth(9, 130);  // Rosso2
  sh.setColumnWidth(10, 200); // Note
  sh.setColumnWidth(11, 110); // Stato
  sh.setColumnWidth(12, 200); // Doppione
  sh.setColumnWidth(13, 120); // Validato da
  sh.setColumnWidth(14, 120); // Validato il

  sh.setFrozenRows(1);

  // Validazione dropdown per colonna Stato
  const statoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["In attesa", "Da confermare", "Validata"], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, COL.STATO, 200, 1).setDataValidation(statoRule);

  // Formattazione condizionale per stato
  const rules = sh.getConditionalFormatRules();

  // Verde = Validata
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$J2="Validata"`)
    .setBackground(COLORI.VALIDATA)
    .setRanges([sh.getRange("A2:M200")])
    .build());

  // Giallo = Da confermare
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$J2="Da confermare"`)
    .setBackground(COLORI.DA_CONFERMARE)
    .setRanges([sh.getRange("A2:M200")])
    .build());

  // Arancione colonna doppione se non vuota
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$K2<>""`)
    .setBackground(COLORI.DOPPIONE)
    .setRanges([sh.getRange("K2:K200")])
    .build());

  sh.setConditionalFormatRules(rules);
}

function _setupGiocatori(ss) {
  const sh = ss.getSheetByName(FOGLI.GIOCATORI);
  sh.clearFormats();

  const headers = ["Nome", "Sede", "ELO attuale", "Ultima importazione"];
  const headerRow = sh.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setBackground(COLORI.HEADER);
  headerRow.setFontColor("#FFFFFF");
  headerRow.setFontWeight("bold");
  headerRow.setFontFamily("Arial");
  headerRow.setFontSize(10);

  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 140);
  sh.setFrozenRows(1);

  // Nota istruzione
  sh.getRange("A2").setValue("← Importa qui i giocatori dal CSV dell'app (Nome, Sede)");
  sh.getRange("A2").setFontColor("#AAAAAA").setFontStyle("italic");
}

function _setupDaConfermare(ss) {
  const sh = ss.getSheetByName(FOGLI.DA_CONFERMARE);
  sh.clearFormats();

  const headers = [
    "Timestamp Invio", "Data/Ora Gara",
    "Blu 1", "Blu 2", "Gol Blu", "Gol Rossi",
    "Rosso 1", "Rosso 2",
    "Note", "Motivo", "Riga originale"
  ];
  const headerRow = sh.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setBackground("#E65100");
  headerRow.setFontColor("#FFFFFF");
  headerRow.setFontWeight("bold");
  headerRow.setFontFamily("Arial");
  headerRow.setFontSize(10);
  sh.setFrozenRows(1);
  sh.getRange("A2:K1000").setBackground(COLORI.DA_CONFERMARE);
}

function _setupValidate(ss) {
  const sh = ss.getSheetByName(FOGLI.VALIDATE);
  sh.clearFormats();

  const headers = [
    "Timestamp Invio", "Data/Ora Gara",
    "Blu 1", "Blu 2", "Gol Blu", "Gol Rossi",
    "Rosso 1", "Rosso 2",
    "Validato da", "Validato il"
  ];
  const headerRow = sh.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setBackground("#2E7D32");
  headerRow.setFontColor("#FFFFFF");
  headerRow.setFontWeight("bold");
  headerRow.setFontFamily("Arial");
  headerRow.setFontSize(10);
  sh.setFrozenRows(1);
  sh.getRange("A2:J1000").setBackground(COLORI.VALIDATA);

  // Bottone export (nota: in GSheet non ci sono veri bottoni,
  // usiamo il menu DdT per l'export)
  sh.getRange("L1").setValue("👆 Usa menu DdT → Esporta CSV validate");
  sh.getRange("L1").setFontColor("#888888").setFontStyle("italic").setFontSize(9);
}

function _setupEsportate(ss) {
  const sh = ss.getSheetByName(FOGLI.ESPORTATE);
  sh.clearFormats();

  const headers = [
    "Timestamp Invio", "Data/Ora Gara",
    "Blu 1", "Blu 2", "Gol Blu", "Gol Rossi",
    "Rosso 1", "Rosso 2",
    "Validato da", "Validato il", "Esportato il"
  ];
  const headerRow = sh.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setBackground("#455A64");
  headerRow.setFontColor("#FFFFFF");
  headerRow.setFontWeight("bold");
  headerRow.setFontFamily("Arial");
  headerRow.setFontSize(10);
  sh.setFrozenRows(1);
  sh.getRange("A2:K1000").setBackground("#ECEFF1");
}

function _setupConfig(ss) {
  const sh = ss.getSheetByName(FOGLI.CONFIG);
  sh.clearFormats();

  sh.getRange("A1").setValue("⚙️ CONFIGURAZIONE DdT — Ancora Insieme");
  sh.getRange("A1").setFontWeight("bold").setFontSize(13)
    .setFontColor(COLORI.HEADER).setFontFamily("Arial");

  const config = [
    ["", ""],
    ["Email collaboratori DdT (una per riga):", ""],
    ["Email 1", "ddt_ancorainsieme@gmail.com"],
    ["Email 2", ""],
    ["Email 3", ""],
    ["Email 4", ""],
    ["", ""],
    ["Soglia doppioni (minuti):", "30"],
    ["", ""],
    ["── ISTRUZIONI FORM ──────────────────", ""],
    ["", ""],
    ["Crea il Form Google con questi campi:", ""],
    ["Campo 1", "Data e ora di fine gara (Data e ora)"],
    ["Campo 2", "Squadra BLU — Giocatore 1 (Risposta breve)"],
    ["Campo 3", "Squadra BLU — Giocatore 2 (Risposta breve)"],
    ["Campo 4", "Gol segnati dalla squadra BLU (Numero intero)"],
    ["Campo 5", "Gol segnati dalla squadra ROSSA (Numero intero)"],
    ["Campo 6", "Squadra ROSSA — Giocatore 1 (Risposta breve)"],
    ["Campo 7", "Squadra ROSSA — Giocatore 2 (Risposta breve)"],
    ["Campo 8", "Note o segnalazioni (Paragrafo — facoltativo)"],
    ["", ""],
    ["Collega il form a questo Sheet:", "Risposte → Crea foglio di lavoro → Seleziona foglio esistente → Segnalazioni"],
    ["", ""],
    ["── ISTRUZIONI IMPORTAZIONE GIOCATORI ──", ""],
    ["", ""],
    ["Dall'app HTML:", "Esporta Classifica XLS → apri in Sheets → copia colonne Nome e Sede nel foglio Giocatori"],
  ];

  sh.getRange(2, 1, config.length, 2).setValues(config);
  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(2, 350);

  // Evidenzia le etichette
  for (let i = 0; i < config.length; i++) {
    const row = i + 2;
    if (config[i][0] && !config[i][0].startsWith("──") && config[i][1] !== "") {
      sh.getRange(row, 1).setFontWeight("bold").setFontColor("#333333");
    }
    if (config[i][0].startsWith("──")) {
      sh.getRange(row, 1, 1, 2).setBackground(COLORI.HEADER)
        .setFontColor("#FFFFFF").setFontWeight("bold");
    }
  }
}


// ============================================================
//  TRIGGER — chiamato automaticamente dal form
// ============================================================
function _setupTriggerForm(ss) {
  // Rimuove trigger esistenti per evitare duplicati
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onFormSubmit") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}

function onFormSubmit(e) {
  // Piccola pausa per assicurarsi che il form abbia scritto tutti i valori
  Utilities.sleep(2000);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(FOGLI.SEGNALAZIONI);

  // L'ultima riga compilata è quella appena arrivata
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  // Verifica che la riga abbia dati (non sovrascrivere se già processata)
  const statoAttuale = sh.getRange(lastRow, COL.STATO).getValue();
  if (statoAttuale && statoAttuale !== "") return;

  // Imposta stato "In attesa"
  sh.getRange(lastRow, COL.STATO).setValue("In attesa");

  // Controlla doppioni
  _controllaDoppioni(sh, lastRow);

  // Notifica email al DdT
  _notificaDdT(sh, lastRow);
}


// ============================================================
//  RILEVAMENTO DOPPIONI
// ============================================================
function _controllaDoppioni(sh, rigaTarget) {
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return; // Meno di 2 righe dati, niente da confrontare

  const datiTarget = sh.getRange(rigaTarget, 1, 1, COL.NOTE).getValues()[0];
  const dataGaraTarget = _parseData(datiTarget[COL.DATA_GARA - 1]);
  if (!dataGaraTarget) return;

  const giocatoriTarget = new Set([
    _normalizzaNome(datiTarget[COL.GIOCATORE_BLU1 - 1]),
    _normalizzaNome(datiTarget[COL.GIOCATORE_BLU2 - 1]),
    _normalizzaNome(datiTarget[COL.GIOCATORE_ROSSI1 - 1]),
    _normalizzaNome(datiTarget[COL.GIOCATORE_ROSSI2 - 1]),
  ]);

  const soglia = _leggiConfig("Soglia doppioni (minuti):") || CFG.SOGLIA_DOPPIONE_MINUTI;
  const sogliaMs = soglia * 60 * 1000;

  for (let r = 2; r < rigaTarget; r++) {
    const dati = sh.getRange(r, 1, 1, COL.NOTE).getValues()[0];
    const stato = sh.getRange(r, COL.STATO).getValue();
    if (stato === "Da confermare") continue; // Ignora già scartate

    const dataGara = _combinaDataOra(dati[COL.DATA_GARA - 1], dati[COL.ORA_GARA - 1]);
    if (!dataGara) continue;

    const diffMs = Math.abs(dataGaraTarget - dataGara);
    if (diffMs > sogliaMs) continue;

    const giocatori = new Set([
      _normalizzaNome(dati[COL.GIOCATORE_BLU1 - 1]),
      _normalizzaNome(dati[COL.GIOCATORE_BLU2 - 1]),
      _normalizzaNome(dati[COL.GIOCATORE_ROSSI1 - 1]),
      _normalizzaNome(dati[COL.GIOCATORE_ROSSI2 - 1]),
    ]);

    // Stesso gruppo di 4 giocatori (in qualsiasi ordine)
    const stessiGiocatori = [...giocatoriTarget].every(g => giocatori.has(g)) &&
                            giocatoriTarget.size === giocatori.size;

    if (stessiGiocatori) {
      const diffMin = Math.round(diffMs / 60000);
      sh.getRange(rigaTarget, COL.DOPPIONE)
        .setValue(`⚠️ Possibile doppione: riga ${r} (${diffMin} min di distanza)`);
      return;
    }
  }
}

function _parseData(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const d = new Date(val);
  return isNaN(d) ? null : d;
}

function _combinaDataOra(dataVal, oraVal) {
  const data = _parseData(dataVal);
  if (!data) return null;
  if (oraVal) {
    // oraVal può essere stringa "HH:mm" o "HH:mm:ss"
    const oraStr = oraVal.toString().trim();
    const parti = oraStr.split(":");
    if (parti.length >= 2) {
      data.setHours(parseInt(parti[0]) || 0);
      data.setMinutes(parseInt(parti[1]) || 0);
      data.setSeconds(parseInt(parti[2]) || 0);
    }
  }
  return data;
}

function _normalizzaNome(nome) {
  if (!nome) return "";
  return nome.toString().trim().toLowerCase();
}


// ============================================================
//  NOTIFICA EMAIL
// ============================================================
function _notificaDdT(sh, riga) {
  const emails = _leggiEmailDdT();
  if (!emails.length) return;

  const dati = sh.getRange(riga, 1, 1, COL.NOTE).getValues()[0];
  const doppione = sh.getRange(riga, COL.DOPPIONE).getValue();

  const dataGara = _combinaDataOra(dati[COL.DATA_GARA - 1], dati[COL.ORA_GARA - 1]);
  const dataFmt = dataGara instanceof Date
    ? Utilities.formatDate(dataGara, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
    : (dati[COL.DATA_GARA - 1] || "non specificata");

  const b1 = dati[COL.GIOCATORE_BLU1 - 1] || "?";
  const b2 = dati[COL.GIOCATORE_BLU2 - 1] || "?";
  const gb = dati[COL.GOL_BLU - 1] || "?";
  const gr = dati[COL.GOL_ROSSI - 1] || "?";
  const r1 = dati[COL.GIOCATORE_ROSSI1 - 1] || "?";
  const r2 = dati[COL.GIOCATORE_ROSSI2 - 1] || "?";
  const note = dati[COL.NOTE - 1] || "—";

  const allerta = doppione ? `\n⚠️ ATTENZIONE: ${doppione}\n` : "";

  const corpo =
    `Nuova partita segnalata nel torneo ${CFG.NOME_TORNEO}.\n` +
    `${allerta}\n` +
    `📅 Data/ora gara: ${dataFmt}\n` +
    `🔵 Squadra BLU:   ${b1} + ${b2}  (${gb} gol)\n` +
    `🔴 Squadra ROSSA: ${r1} + ${r2}  (${gr} gol)\n` +
    `📝 Note: ${note}\n\n` +
    `Apri il foglio per validare:\n${SpreadsheetApp.getActiveSpreadsheet().getUrl()}` +
    `#gid=${SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FOGLI.SEGNALAZIONI).getSheetId()}`;

  const oggetto = doppione
    ? `⚠️ [${CFG.NOME_TORNEO}] Nuova segnalazione — POSSIBILE DOPPIONE`
    : `⚽ [${CFG.NOME_TORNEO}] Nuova partita da validare`;

  emails.forEach(email => {
    try {
      MailApp.sendEmail({ to: email, subject: oggetto, body: corpo });
    } catch (err) {
      Logger.log("Errore invio email a " + email + ": " + err);
    }
  });
}

function _leggiEmailDdT() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FOGLI.CONFIG);
  if (!sh) return [];
  const emails = [];
  // Le email sono nelle righe 4-7 colonna B
  for (let r = 4; r <= 7; r++) {
    const val = sh.getRange(r, 2).getValue().toString().trim();
    if (val && val.includes("@")) emails.push(val);
  }
  return emails;
}

function _leggiConfig(chiave) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FOGLI.CONFIG);
  if (!sh) return null;
  const dati = sh.getDataRange().getValues();
  for (const row of dati) {
    if (row[0].toString().trim() === chiave) return row[1];
  }
  return null;
}


// ============================================================
//  AZIONI DdT — chiamate dal menu
// ============================================================

// Valida la riga selezionata
function validaRigaSelezionata() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FOGLI.SEGNALAZIONI);
  const riga = SpreadsheetApp.getActiveSheet().getActiveRange().getRow();

  if (SpreadsheetApp.getActiveSheet().getName() !== FOGLI.SEGNALAZIONI) {
    SpreadsheetApp.getUi().alert("Seleziona una riga nel foglio Segnalazioni.");
    return;
  }
  if (riga < 2) {
    SpreadsheetApp.getUi().alert("Seleziona una riga dati (non l'intestazione).");
    return;
  }

  const stato = sh.getRange(riga, COL.STATO).getValue();
  if (stato === "Validata") {
    SpreadsheetApp.getUi().alert("Questa partita è già validata.");
    return;
  }

  const utente = Session.getActiveUser().getEmail() || "DdT";
  const ora = new Date();

  sh.getRange(riga, COL.STATO).setValue("Validata");
  sh.getRange(riga, COL.VALIDATO_DA).setValue(utente);
  sh.getRange(riga, COL.VALIDATO_IL).setValue(Utilities.formatDate(ora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"));

  // Sposta nel foglio Validate
  _spostaRiga(sh, riga, FOGLI.VALIDATE, [utente, Utilities.formatDate(ora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")]);

  SpreadsheetApp.getUi().alert("✅ Partita validata e spostata nel foglio Validate.");
}

// Marca come "Da confermare" la riga selezionata
function marcaDaConfermare() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FOGLI.SEGNALAZIONI);

  if (SpreadsheetApp.getActiveSheet().getName() !== FOGLI.SEGNALAZIONI) {
    SpreadsheetApp.getUi().alert("Seleziona una riga nel foglio Segnalazioni.");
    return;
  }

  const riga = SpreadsheetApp.getActiveSheet().getActiveRange().getRow();
  if (riga < 2) {
    SpreadsheetApp.getUi().alert("Seleziona una riga dati (non l'intestazione).");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const risposta = ui.prompt(
    "Da confermare",
    "Inserisci il motivo (verrà mostrato nel foglio 'Da Confermare'):",
    ui.ButtonSet.OK_CANCEL
  );
  if (risposta.getSelectedButton() !== ui.Button.OK) return;
  const motivo = risposta.getResponseText().trim() || "Da verificare";

  sh.getRange(riga, COL.STATO).setValue("Da confermare");

  // Sposta nel foglio Da Confermare
  _spostaRigaDaConfermare(sh, riga, motivo);

  SpreadsheetApp.getUi().alert("⚠️ Partita spostata in 'Da Confermare'.");
}

// Esporta CSV delle partite validate
function esportaCSVValidate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shVal = ss.getSheetByName(FOGLI.VALIDATE);
  const shExp = ss.getSheetByName(FOGLI.ESPORTATE);

  const lastRow = shVal.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Nessuna partita validata da esportare.");
    return;
  }

  const dati = shVal.getRange(2, 1, lastRow - 1, 10).getValues();
  const oraExport = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  // Costruisce CSV
  const intestazione = "timestamp_invio,data_gara,blu1,blu2,gol_blu,gol_rossi,rossi1,rossi2,validato_da,validato_il";
  const righe = dati.map(r => r.map(c => {
    const val = c instanceof Date
      ? Utilities.formatDate(c, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      : c.toString();
    return `"${val.replace(/"/g, '""')}"`;
  }).join(","));

  const csv = intestazione + "\n" + righe.join("\n");

  // Salva come file nella cartella del Drive
  const nomeFile = `AncoraInsieme_Gare_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm")}.csv`;
  const file = DriveApp.createFile(nomeFile, csv, MimeType.CSV);
  const url = file.getUrl();

  // Sposta le righe nel foglio Esportate
  for (let i = dati.length - 1; i >= 0; i--) {
    const rigaVal = i + 2;
    const rigaEsp = shExp.getLastRow() + 1;
    const valori = shVal.getRange(rigaVal, 1, 1, 10).getValues()[0];
    shExp.getRange(rigaEsp, 1, 1, 11).setValues([[...valori, oraExport]]);
    shVal.deleteRow(rigaVal);
  }

  SpreadsheetApp.getUi().alert(
    `✅ Esportate ${dati.length} partite.\n\n` +
    `File salvato su Drive:\n${url}\n\n` +
    `Le partite sono state spostate nel foglio 'Esportate' e non saranno più esportabili.`
  );
}

// Importa giocatori da CSV (incolla testo CSV nella finestra di dialogo)
function importaGiocatori() {
  const ui = SpreadsheetApp.getUi();
  const risposta = ui.prompt(
    "Importa Giocatori",
    "Incolla il contenuto CSV (nome,sede — una riga per giocatore, senza intestazione):\n\nEs:\nCarlo,Levico\nSonia,Levico\nAndrea,Cantù",
    ui.ButtonSet.OK_CANCEL
  );
  if (risposta.getSelectedButton() !== ui.Button.OK) return;

  const testo = risposta.getResponseText().trim();
  if (!testo) return;

  const righe = testo.split("\n").map(r => r.split(",").map(c => c.trim()));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(FOGLI.GIOCATORI);

  // Svuota dati esistenti (mantieni header)
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 4).clearContent();

  const oraImport = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const datiDaScrivere = righe.map(r => [r[0] || "", r[1] || "", "", oraImport]);
  sh.getRange(2, 1, datiDaScrivere.length, 4).setValues(datiDaScrivere);

  // Aggiorna validazione a tendina nelle Segnalazioni
  _aggiornaValidazioneNomi(ss, sh);

  ui.alert(`✅ Importati ${datiDaScrivere.length} giocatori.`);
}

function _aggiornaValidazioneNomi(ss, shGiocatori) {
  const lastRow = shGiocatori.getLastRow();
  if (lastRow < 2) return;

  const nomi = shGiocatori.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => r[0].toString().trim())
    .filter(n => n);

  if (!nomi.length) return;

  const shSeg = ss.getSheetByName(FOGLI.SEGNALAZIONI);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(nomi, true)
    .setAllowInvalid(true)  // Permette nomi non in lista (con avviso)
    .setHelpText("Seleziona un giocatore dalla lista o scrivi il nome")
    .build();

  [COL.GIOCATORE_BLU1, COL.GIOCATORE_BLU2, COL.GIOCATORE_ROSSI1, COL.GIOCATORE_ROSSI2].forEach(col => {
    shSeg.getRange(2, col, 1000, 1).setDataValidation(rule);
  });
}

// Riesegui controllo doppioni su tutte le segnalazioni
function ricalcolaDoppioni() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FOGLI.SEGNALAZIONI);
  const lastRow = sh.getLastRow();
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert("Meno di 2 segnalazioni presenti.");
    return;
  }

  // Pulisci colonna doppioni
  sh.getRange(2, COL.DOPPIONE, lastRow - 1, 1).clearContent();

  for (let r = 3; r <= lastRow; r++) {
    _controllaDoppioni(sh, r);
  }
  SpreadsheetApp.getUi().alert("✅ Controllo doppioni completato.");
}


// ============================================================
//  UTILITY — spostamento righe tra fogli
// ============================================================

function _spostaRiga(shSorgente, riga, nomeDest, colAggiuntive) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shDest = ss.getSheetByName(nomeDest);
  const dati = shSorgente.getRange(riga, 1, 1, 11).getValues()[0];
  const rigaDest = shDest.getLastRow() + 1;
  shDest.getRange(rigaDest, 1, 1, dati.length).setValues([dati]);
  if (colAggiuntive) {
    shDest.getRange(rigaDest, dati.length + 1, 1, colAggiuntive.length).setValues([colAggiuntive]);
  }
  // Non cancelliamo la riga sorgente — manteniamo lo stato aggiornato in Segnalazioni
  // La colorazione condizionale mostrerà lo stato
}

function _spostaRigaDaConfermare(shSorgente, riga, motivo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shDest = ss.getSheetByName(FOGLI.DA_CONFERMARE);
  const dati = shSorgente.getRange(riga, 1, 1, 11).getValues()[0];
  const rigaDest = shDest.getLastRow() + 1;
  shDest.getRange(rigaDest, 1, 1, dati.length).setValues([dati]);
  shDest.getRange(rigaDest, dati.length + 1).setValue(motivo);
  shDest.getRange(rigaDest, dati.length + 2).setValue(`Riga ${riga} in Segnalazioni_form`);
}



// ============================================================
//  CONTROLLA NOMI — verifica i nomi dei giocatori
// ============================================================

function controllaSegnalazione() {
  controllaNomi(true);
}

function controllaNomi(includiRisultato) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSeg = ss.getSheetByName(FOGLI.SEGNALAZIONI);
  const shGio = ss.getSheetByName(FOGLI.GIOCATORI);

  if (!shSeg || !shGio) {
    SpreadsheetApp.getUi().alert("Fogli non trovati. Verifica il setup.");
    return;
  }

  const lastRowSeg = shSeg.getLastRow();
  if (lastRowSeg < 2) {
    SpreadsheetApp.getUi().alert("Nessuna segnalazione da controllare.");
    return;
  }

  // Carica elenco giocatori
  const lastRowGio = shGio.getLastRow();
  const giocatori = lastRowGio > 1
    ? shGio.getRange(2, 1, lastRowGio - 1, 1).getValues()
        .map(r => r[0].toString().trim())
        .filter(n => n && n !== "← Importa qui i giocatori dal CSV dell\'app (Nome, Sede)")
    : [];

  if (!giocatori.length) {
    SpreadsheetApp.getUi().alert("Nessun giocatore nel foglio Giocatori. Importali prima con il menu DdT.");
    return;
  }

  let controllate = 0, okCount = 0, problemiCount = 0;

  for (let r = 2; r <= lastRowSeg; r++) {
    const stato = shSeg.getRange(r, COL.STATO).getValue().toString().trim();
    if (stato !== "In attesa") continue;

    controllate++;
    const dati = shSeg.getRange(r, 1, 1, COL.EMAIL_SEGNALATORE).getValues()[0];
    const nomiDaVerificare = [
      { col: COL.GIOCATORE_BLU1,   valore: dati[COL.GIOCATORE_BLU1 - 1] },
      { col: COL.GIOCATORE_BLU2,   valore: dati[COL.GIOCATORE_BLU2 - 1] },
      { col: COL.GIOCATORE_ROSSI1, valore: dati[COL.GIOCATORE_ROSSI1 - 1] },
      { col: COL.GIOCATORE_ROSSI2, valore: dati[COL.GIOCATORE_ROSSI2 - 1] },
    ];

    let tuttiOk = true;
    const problemi = [];

    nomiDaVerificare.forEach(({ col, valore }) => {
      const nome = valore.toString().trim();
      if (!nome) { tuttiOk = false; problemi.push({ nome: "(vuoto)", suggerimenti: [] }); return; }

      // Cerca corrispondenza esatta (case insensitive)
      const trovato = giocatori.find(g => g.toLowerCase() === nome.toLowerCase());
      if (trovato) {
        // Correggi maiuscolo/minuscolo se necessario
        if (trovato !== nome) shSeg.getRange(r, col).setValue(trovato);
        return;
      }

      // Non trovato — calcola similarità
      tuttiOk = false;
      const candidati = _trovaCandidati(nome, giocatori, 3);
      problemi.push({ nome, suggerimenti: candidati });
    });

    // Verifica risultato
    const problemiRisultato = [];
    if (includiRisultato) {
      const golBlu  = parseInt(dati[COL.GOL_BLU - 1]) || 0;
      const golRossi = parseInt(dati[COL.GOL_ROSSI - 1]) || 0;
      const errRis = _verificaRisultato(golBlu, golRossi);
      if (errRis) {
        tuttiOk = false;
        problemiRisultato.push(errRis);
      }
    }

    if (tuttiOk) {
      shSeg.getRange(r, COL.STATO).setValue("Validata");
      _spostaRiga(shSeg, r, FOGLI.VALIDATE, [Session.getActiveUser().getEmail()||"DdT (auto)", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")]);
      okCount++;
    } else {
      shSeg.getRange(r, COL.STATO).setValue("Da confermare");
      problemiCount++;

      // Costruisce motivo completo
      const motivoNomi = problemi.map(p =>
        `"${p.nome}" non trovato${p.suggerimenti.length ? " — forse: " + p.suggerimenti.join(", ") : ""}`
      ).join(" | ");
      const motivoRis = problemiRisultato.join(" | ");
      const motivoCompleto = [motivoNomi, motivoRis].filter(Boolean).join(" || ");

      _spostaRigaDaConfermare(shSeg, r, motivoCompleto);

      // Notifica automatica al segnalante
      const emailSegnalante = dati[COL.EMAIL_SEGNALATORE - 1].toString().trim();
      if (emailSegnalante && emailSegnalante.includes("@")) {
        _notificaSegnalante(emailSegnalante, dati, problemi, problemiRisultato);
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    `✅ Controllo completato su ${controllate} segnalazioni "In attesa":\n\n` +
    `✔️ Pronte per validazione: ${okCount}\n` +
    `⚠️ Da confermare (nomi non riconosciuti): ${problemiCount}`
  );
}

// Calcola i candidati più simili con algoritmo di Levenshtein
function _trovaCandidati(nome, giocatori, maxCandidati) {
  const nomeNorm = nome.toLowerCase();
  const scored = giocatori.map(g => ({
    nome: g,
    score: _similarita(nomeNorm, g.toLowerCase())
  }));
  scored.sort((a, b) => b.score - a.score);
  return scored.slice(0, maxCandidati)
    .filter(s => s.score > 0.3)
    .map(s => s.nome);
}

// Similarità stringa (Jaro-Winkler semplificato)
function _similarita(s1, s2) {
  if (s1 === s2) return 1;
  if (!s1 || !s2) return 0;

  // Distanza di Levenshtein normalizzata
  const len1 = s1.length, len2 = s2.length;
  const matrix = [];
  for (let i = 0; i <= len1; i++) matrix[i] = [i];
  for (let j = 0; j <= len2; j++) matrix[0][j] = j;
  for (let i = 1; i <= len1; i++) {
    for (let j = 1; j <= len2; j++) {
      const cost = s1[i-1] === s2[j-1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i-1][j] + 1,
        matrix[i][j-1] + 1,
        matrix[i-1][j-1] + cost
      );
    }
  }
  const dist = matrix[len1][len2];
  const maxLen = Math.max(len1, len2);

  // Bonus se iniziano con le stesse lettere
  let bonus = 0;
  for (let k = 0; k < Math.min(3, len1, len2); k++) {
    if (s1[k] === s2[k]) bonus += 0.05; else break;
  }

  return Math.max(0, 1 - dist / maxLen + bonus);
}

// Verifica conformità risultato al regolamento
function _verificaRisultato(golBlu, golRossi) {
  if (isNaN(golBlu) || isNaN(golRossi)) return "Risultato non numerico";
  if (golBlu < 0 || golRossi < 0) return "Risultato negativo non valido";
  if (golBlu === golRossi) return `Risultato ${golBlu}-${golRossi}: i pareggi non sono ammessi`;
  const mx = Math.max(golBlu, golRossi);
  const mn = Math.min(golBlu, golRossi);
  if (mx < 10) return `Risultato ${golBlu}-${golRossi}: la vittoria deve essere a 10 gol`;
  if (mx === 10 && mn === 9) return `Risultato ${golBlu}-${golRossi}: scarto minimo 2 gol (es. 10-8)`;
  if (mx > 10 && (mx - mn) !== 2) return `Risultato ${golBlu}-${golRossi}: oltre i 10 gol lo scarto deve essere esattamente 2`;
  return null; // OK
}

// Email automatica al segnalante
function _notificaSegnalante(email, dati, problemi, problemiRisultato) {
  problemiRisultato = problemiRisultato || [];
  const segnalatore = dati[COL.SEGNALATORE - 1] || "partecipante";
  const dataGara = _combinaDataOra(dati[COL.DATA_GARA - 1], dati[COL.ORA_GARA - 1]);
  const dataFmt = dataGara instanceof Date
    ? Utilities.formatDate(dataGara, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
    : "data non specificata";

  const dettagliNomi = problemi.length ? "\n🔴 NOMI NON RICONOSCIUTI:\n" + problemi.map(p => {
    const sugg = p.suggerimenti.length
      ? `\n   Forse intendevi: ${p.suggerimenti.join(", ")}?`
      : "\n   (nessun giocatore simile trovato — verifica l\'iscrizione)";
    return `• "${p.nome}" non riconosciuto${sugg}`;
  }).join("\n") : "";

  const dettagliRisultato = problemiRisultato.length
    ? "\n\n⚽ RISULTATO NON CONFORME AL REGOLAMENTO:\n" + problemiRisultato.map(p => `• ${p}`).join("\n") +
      "\n\n   Ricorda: vittoria a 10 gol, scarto minimo 2 gol, no pareggi."
    : "";

  const corpo =
    `Ciao ${segnalatore},\n\n` +
    `grazie per aver segnalato la partita del ${dataFmt} nel torneo ${CFG.NOME_TORNEO}.\n\n` +
    `La segnalazione richiede alcune verifiche:` +
    `${dettagliNomi}` +
    `${dettagliRisultato}\n\n` +
    `Ti chiediamo di verificare e segnalare nuovamente la partita con i dati corretti,\n` +
    `oppure di contattare il Direttore di Torneo per chiarimenti.\n\n` +
    `I giocatori devono essere inseriti come: Cognome Nome (es. Rossi Mario)\n\n` +
    `Grazie,\nDirettore di Torneo — ${CFG.NOME_TORNEO}`;

  try {
    MailApp.sendEmail({
      to: email,
      subject: `[${CFG.NOME_TORNEO}] Partita in attesa di verifica — dati da correggere`,
      body: corpo
    });
  } catch (err) {
    Logger.log("Errore invio email segnalante " + email + ": " + err);
  }
}


// ============================================================
//  WEB APP — riceve dati dal form HTML esterno (SharePoint)
// ============================================================

function doPost(e) {
  try {
    // Accetta sia JSON che form fields
    let dati;
    if (e.postData && e.postData.type === 'application/json') {
      dati = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      dati = e.parameter;
    } else {
      dati = JSON.parse(e.postData.contents);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(FOGLI.SEGNALAZIONI);
    if (!sh) return _rispostaJSON({ ok: false, errore: "Foglio non trovato" });

    // Valida campi obbligatori
    const campiObbligatori = ['blu1','blu2','rossi1','rossi2','golBlu','golRossi','dataGara','segnalatore','emailSegnalatore'];
    const mancanti = campiObbligatori.filter(c => !dati[c] && dati[c] !== 0);
    if (mancanti.length) return _rispostaJSON({ ok: false, errore: `Campi mancanti: ${mancanti.join(', ')}` });

    // Validazione risultato
    const errRis = _verificaRisultato(parseInt(dati.golBlu), parseInt(dati.golRossi));
    if (errRis) return _rispostaJSON({ ok: false, errore: errRis });

    // Giocatori distinti
    if (new Set([dati.blu1,dati.blu2,dati.rossi1,dati.rossi2]).size !== 4) {
      return _rispostaJSON({ ok: false, errore: "I 4 giocatori devono essere distinti" });
    }

    // Timestamp invio
    const ora = new Date();
    const oraFmt = Utilities.formatDate(ora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    // Parsing data e ora gara
    let dataGaraFmt = dati.dataGara || '';
    let oraGaraFmt  = dati.oraGara  || '';

    // Scrive la riga nel foglio
    const nuovaRiga = [
      oraFmt,           // A - Timestamp invio
      dataGaraFmt,      // B - Data gara
      oraGaraFmt,       // C - Ora gara
      dati.blu1,        // D
      dati.blu2,        // E
      parseInt(dati.golBlu),   // F
      parseInt(dati.golRossi), // G
      dati.rossi1,      // H
      dati.rossi2,      // I
      dati.note || '',  // J
      dati.segnalatore, // K
      dati.emailSegnalatore, // L
      '',               // M - Stato (verrà impostato da onPost)
      '',               // N - Doppione
      '',               // O - Validato da
      '',               // P - Validato il
    ];

    const ultimaRiga = sh.getLastRow() + 1;
    sh.getRange(ultimaRiga, 1, 1, nuovaRiga.length).setValues([nuovaRiga]);

    // Pausa e poi imposta stato + controlla doppioni + notifica
    Utilities.sleep(500);
    sh.getRange(ultimaRiga, COL.STATO).setValue("In attesa");
    _controllaDoppioni(sh, ultimaRiga);
    _notificaDdT(sh, ultimaRiga);

    return _rispostaJSON({ ok: true, messaggio: "Partita registrata con successo!" });

  } catch(err) {
    return _rispostaJSON({ ok: false, errore: "Errore interno: " + err.message });
  }
}

// Risposta CORS-friendly per il form HTML
function _rispostaJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Gestisce anche le richieste GET (per test e per fornire l'elenco giocatori)
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(FOGLI.GIOCATORI);
  
  let giocatori = [];
  if (sh && sh.getLastRow() >= 2) {
    const dati = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
    giocatori = dati
      .map(r => ({ nome: r[0].toString().trim(), sede: r[1].toString().trim() }))
      .filter(g => g.nome && !g.nome.startsWith('\u2190'));
  }

  const output = ContentService
    .createTextOutput(JSON.stringify({ giocatori }))
    .setMimeType(ContentService.MimeType.JSON);
  
  return output;
}

// ============================================================
//  RESET COMPLETO
// ============================================================
function resetCompleto() {
  const ui = SpreadsheetApp.getUi();
  
  // Prima conferma
  const r1 = ui.alert(
    "⚠️ RESET COMPLETO",
    "Questa operazione cancellerà TUTTI i dati dal foglio:\n\n" +
    "• Tutte le segnalazioni\n" +
    "• Tutti i giocatori importati\n" +
    "• Tutte le partite validate\n" +
    "• Tutte le partite esportate\n" +
    "• Tutte le partite da confermare\n\n" +
    "Le intestazioni e la configurazione verranno mantenute.\n\n" +
    "Sei sicuro di voler continuare?",
    ui.ButtonSet.YES_NO
  );
  if (r1 !== ui.Button.YES) return;

  // Seconda conferma
  const r2 = ui.alert(
    "⚠️ ULTIMA CONFERMA",
    "I dati cancellati non potranno essere recuperati.\n\n" +
    "Confermi il reset completo?",
    ui.ButtonSet.YES_NO
  );
  if (r2 !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cancella dati (mantiene riga 1 = intestazioni)
  const foghliDaResettare = [
    FOGLI.SEGNALAZIONI,
    FOGLI.GIOCATORI,
    FOGLI.DA_CONFERMARE,
    FOGLI.VALIDATE,
    FOGLI.ESPORTATE
  ];

  foghliDaResettare.forEach(nomeFoglio => {
    const sh = ss.getSheetByName(nomeFoglio);
    if (!sh) return;
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
    }
  });

  // Reimposta nota nel foglio Giocatori
  ss.getSheetByName(FOGLI.GIOCATORI).getRange("A2")
    .setValue("← Importa qui i giocatori dal CSV dell'app (Nome, Sede)")
    .setFontColor("#AAAAAA").setFontStyle("italic");

  ui.alert("✅ Reset completato. Il foglio è pronto per i dati reali.");
}

// ============================================================
//  MENU DdT
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚽ DdT Ancora Insieme")
    .addItem("✅ Valida partita selezionata", "validaRigaSelezionata")
    .addItem("⚠️ Marca come Da confermare", "marcaDaConfermare")
    .addSeparator()
    .addItem("📥 Importa giocatori (CSV)", "importaGiocatori")
    .addItem("🔍 Controlla segnalazione (nomi + risultato)", "controllaSegnalazione")
    .addItem("🔄 Ricalcola doppioni", "ricalcolaDoppioni")
    .addSeparator()
    .addItem("📤 Esporta CSV partite validate", "esportaCSVValidate")
    .addSeparator()
    .addItem("🗑️ Reset completo (cancella tutti i dati)", "resetCompleto")
    .addToUi();
}
