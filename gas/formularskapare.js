/**
 * Skript 1: Formulärskapare
 * Skapar Google Formulär från fliken "Frågor" och bygger facitflikar automatiskt.
 *
 * Frågor-flikens kolumner (rad 1 = rubrikrad):
 *   A: TestID   B: Område   C: Frågetext   D: Rätt svar   E: Fel alt 1   F: Fel alt 2   G: Fel alt 3
 *
 * Vad scriptet gör:
 *   1. Grupperar frågor per TestID
 *   2. Skapar ett Google Formulär per TestID med slumpad alternativordning
 *   3. Skapar en facitflik (FACIT_<TestID>) med rätt svar i kolumn B
 *   4. Uppdaterar Testregister med facitflikens namn och svarssheetens namn
 *   5. Kopplar formulärsvar till kalkylarket
 *
 * Kör via menyval: 📝 Formulär → Skapa nytt formulär
 */

var FOLDER_ID = "DIN_MAPP_ID_HÄR";

// ── Meny ───────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📝 Formulär")
    .addItem("Skapa nytt formulär", "skapaFormular")
    .addItem("Hämta befintliga länkar", "hamtaBefintligaLankar")
    .addToUi();
}

// ── Huvudfunktion ──────────────────────────────────────────────────────────────
function skapaFormular() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fragSheet = ss.getSheetByName("Frågor");

  if (!fragSheet) {
    SpreadsheetApp.getUi().alert('Fliken "Frågor" hittades inte.');
    return;
  }

  var data = fragSheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('Inga frågor hittades i fliken "Frågor".');
    return;
  }

  // Gruppera frågor per TestID
  var tester = {};
  var testOrdning = [];
  for (var i = 1; i < data.length; i++) {
    var rad      = data[i];
    var testId   = String(rad[0]).trim();
    var omrade   = String(rad[1]).trim();
    var fragText = String(rad[2]).trim();
    var rattSvar = String(rad[3]).trim();
    var fel1     = String(rad[4]).trim();
    var fel2     = String(rad[5]).trim();
    var fel3     = String(rad[6]).trim();

    if (!testId || !fragText || !rattSvar) continue;

    if (!tester[testId]) {
      tester[testId] = { omrade: omrade, fragor: [] };
      testOrdning.push(testId);
    }
    tester[testId].fragor.push({
      text:     fragText,
      rattSvar: rattSvar,
      fel:      [fel1, fel2, fel3].filter(function(a) { return a !== ""; })
    });
  }

  if (testOrdning.length === 0) {
    SpreadsheetApp.getUi().alert("Hittade inga kompletta frågor (TestID + Frågetext + Rätt svar krävs).");
    return;
  }

  var ui       = SpreadsheetApp.getUi();
  var mapp     = DriveApp.getFolderById(FOLDER_ID);
  var skapade  = [];

  for (var t = 0; t < testOrdning.length; t++) {
    var testId = testOrdning[t];
    var info   = tester[testId];

    // ── Skapa formulär ─────────────────────────────────────────────────────────
    var formTitel = "Historia Hi1B – Självtest " + testId + " – " + info.omrade;
    var form      = FormApp.create(formTitel);
    form.setTitle(formTitel);
    form.setDescription("Historia Hi1B");
    form.setRequireLogin(true);
    form.setCollectEmail(true);
    form.setLimitOneResponsePerUser(false);
    form.setShowLinkToRespondAgain(true);

    // Lägg till frågor med slumpad alternativordning
    for (var f = 0; f < info.fragor.length; f++) {
      var fraga = info.fragor[f];
      var alts  = [fraga.rattSvar].concat(fraga.fel);
      alts      = shuffle(alts);

      var item    = form.addMultipleChoiceItem();
      item.setTitle(fraga.text);
      item.setRequired(true);

      var choices = [];
      for (var a = 0; a < alts.length; a++) {
        choices.push(item.createChoice(alts[a], alts[a] === fraga.rattSvar));
      }
      item.setChoices(choices);
    }

    // Flytta formulärfilen till mappen
    var formFil = DriveApp.getFileById(form.getId());
    mapp.addFile(formFil);
    DriveApp.getRootFolder().removeFile(formFil);

    // Koppla svar till detta kalkylark
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

    // ── Bygg facitflik ─────────────────────────────────────────────────────────
    var facitNamn = "FACIT_" + testId;
    var befintlig = ss.getSheetByName(facitNamn);
    if (befintlig) ss.deleteSheet(befintlig);

    var facitSheet = ss.insertSheet(facitNamn);
    for (var f2 = 0; f2 < info.fragor.length; f2++) {
      // Kolumn A: frågenummer, Kolumn B: rätt svar (text). Ingen rubrikrad.
      facitSheet.getRange(f2 + 1, 1).setValue(f2 + 1);
      facitSheet.getRange(f2 + 1, 2).setValue(info.fragor[f2].rattSvar);
    }

    // ── Uppdatera Testregister ─────────────────────────────────────────────────
    // Hitta svarssheetens namn (Google skapar den med formulärets titel + " (svar)")
    // Vi uppdaterar Testregister efter en kort paus så att svarsfliken hinner skapas.
    uppdateraTestregister(ss, testId, info.omrade, facitNamn, form.getId());

    skapade.push({ testId: testId, omrade: info.omrade, antal: info.fragor.length, lank: form.getPublishedUrl() });
  }

  // ── Uppdatera svarssheet-kolumnen i Testregister ───────────────────────────
  // Google skapar svarsfliken asynkront, så vi söker igenom befintliga flikar.
  Utilities.sleep(3000);
  uppdateraSvarssheets(ss);

  // ── Skriv alla formulärlänkar till fliken "Formulärlänkar" ────────────────
  sparaFormularlankar(ss, skapade);

  var meddelande = "Skapade " + skapade.length + " formulär.\n\nAlla länkar finns i fliken \"Formulärlänkar\".";
  SpreadsheetApp.getUi().alert(meddelande);
}

// ── Hjälp: spara formulärlänkar i egen flik ───────────────────────────────────
function sparaFormularlankar(ss, skapade) {
  var flikNamn = "Formulärlänkar";
  var flik = ss.getSheetByName(flikNamn);
  if (!flik) flik = ss.insertSheet(flikNamn);
  flik.clearContents();

  flik.getRange(1, 1, 1, 4).setValues([["TestID", "Område", "Antal frågor", "Formulärlänk"]]);
  flik.getRange(1, 1, 1, 4).setFontWeight("bold");

  for (var i = 0; i < skapade.length; i++) {
    var r = skapade[i];
    flik.appendRow([r.testId, r.omrade, r.antal, r.lank]);
  }

  // Gör länkkolumnen klickbar
  var lankRange = flik.getRange(2, 4, skapade.length, 1);
  lankRange.setFontColor("#1155CC");
}

// ── Hjälp: uppdatera Testregister (kolumn A–C) ────────────────────────────────
function uppdateraTestregister(ss, testId, omrade, facitNamn, formId) {
  var reg = ss.getSheetByName("Testregister");
  if (!reg) {
    reg = ss.insertSheet("Testregister");
    reg.getRange(1, 1, 1, 4).setValues([["TestID", "Område", "Facitflik", "Svarssheet"]]);
  }

  var data = reg.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === testId) {
      reg.getRange(i + 1, 2).setValue(omrade);
      reg.getRange(i + 1, 3).setValue(facitNamn);
      return;
    }
  }
  // Ny rad
  reg.appendRow([testId, omrade, facitNamn, ""]);
}

// ── Hjälp: matcha svarssheet-namn till Testregister kolumn D ─────────────────
function uppdateraSvarssheets(ss) {
  var reg = ss.getSheetByName("Testregister");
  if (!reg) return;

  var data   = reg.getDataRange().getValues();
  var flikar = ss.getSheets().map(function(s) { return s.getName(); });

  for (var i = 1; i < data.length; i++) {
    if (data[i][3]) continue; // redan ifylld

    var testId = String(data[i][0]).trim();
    var omrade = String(data[i][1]).trim();

    // Svarsfliken heter vanligen "Historia Hi1B – Självtest V01 – Område (svar)"
    var sokterm = testId;
    for (var j = 0; j < flikar.length; j++) {
      if (flikar[j].indexOf(sokterm) !== -1 && flikar[j].indexOf("svar") !== -1) {
        reg.getRange(i + 1, 4).setValue(flikar[j]);
        break;
      }
    }
  }
}

// ── Hämta länkar från befintliga formulär i mappen ───────────────────────────
function hamtaBefintligaLankar() {
  var mapp   = DriveApp.getFolderById(FOLDER_ID);
  var filer  = mapp.getFilesByType(MimeType.GOOGLE_FORMS);
  var skapade = [];

  while (filer.hasNext()) {
    var fil  = filer.next();
    var form = FormApp.openById(fil.getId());
    var titel = form.getTitle();

    // Försök plocka ut TestID ur titeln (formatet: "... – Självtest 1_1_U1 – ...")
    var match = titel.match(/Självtest\s+(\S+)\s+–\s+(.+)$/);
    var testId = match ? match[1] : titel;
    var omrade = match ? match[2].trim() : "";

    skapade.push({ testId: testId, omrade: omrade, antal: form.getItems().length, lank: form.getPublishedUrl() });
  }

  if (skapade.length === 0) {
    SpreadsheetApp.getUi().alert("Inga formulär hittades i mappen.");
    return;
  }

  skapade.sort(function(a, b) { return a.testId.localeCompare(b.testId, 'sv'); });
  sparaFormularlankar(SpreadsheetApp.getActiveSpreadsheet(), skapade);
  SpreadsheetApp.getUi().alert("Klart! " + skapade.length + " formulärlänkar sparade i fliken \"Formulärlänkar\".");
}

// ── Fisher-Yates-shuffle ───────────────────────────────────────────────────────
function shuffle(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = a[i]; a[i] = a[j]; a[j] = tmp;
  }
  return a;
}
