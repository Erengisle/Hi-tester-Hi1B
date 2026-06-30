/**
 * Skript 2: Resultathantering
 * Triggas av onFormSubmit. Rättar, loggar och uppdaterar klassöversikt.
 * Eleven får en länk till webbappen istället för ett Drive-dokument.
 *
 * Flikar i kalkylarket:
 *   Testregister  – A: TestID, B: Område, C: Facitflik, D: Svarssheet
 *   Elever        – A: Email, B: Namn, C: Token (unik länknyckel)
 *   RESULTAT_LOGG – A: Tidsstämpel, B: Email, C: TestID, D: Område,
 *                   E: Procent, F: Rätta, G: Totalt, H: Försök
 *   KLASSÖVERSIKT – A: Namn, B+: ett test per kolumn, sista: Medel
 *
 * Konstant att fylla i:
 *   WEBAPP_URL – URL till den publicerade webbappen (Distribuera → Ny distribution)
 */

var WEBAPP_URL = "DIN_WEBAPP_URL_HÄR";

var FARG_ROD  = "#FF9999";
var FARG_GUL  = "#FFEB99";
var FARG_GRON = "#99E699";

// ---------------------------------------------------------------------------
// Huvudtrigger
// ---------------------------------------------------------------------------

function onFormSubmit(e) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var svarSheet = e.range.getSheet();

  var testInfo = hittaTestInfo(ss, svarSheet.getName());
  if (!testInfo) {
    Logger.log("Inget testregister matchar sheet: " + svarSheet.getName());
    return;
  }

  var svarRad   = e.range.getRow();
  var timestamp = svarSheet.getRange(svarRad, 1).getValue();
  var email     = svarSheet.getRange(svarRad, 2).getValue().toString().toLowerCase().trim();

  var elevInfo = hittaElev(ss, email);
  if (!elevInfo) {
    Logger.log("Eleven hittades inte: " + email);
    return;
  }

  // Generera token om eleven saknar en
  var token = elevInfo.token;
  if (!token) {
    token = Utilities.getUuid();
    ss.getSheetByName("Elever").getRange(elevInfo.rad, 3).setValue(token);
    elevInfo.token = token;
    skickaValkommen(elevInfo.email, elevInfo.namn, token);
  }

  var facitData = hamtaFacit(ss, testInfo.facitflik);
  var svarData  = hamtaSvar(svarSheet, svarRad, facitData.length);
  var rattning  = ratta(svarData, facitData);

  loggaResultat(ss, timestamp, email, testInfo.testId, testInfo.omrade, rattning);
  uppdateraKlassoversikt(ss);
}

// ---------------------------------------------------------------------------
// Datahämtning
// ---------------------------------------------------------------------------

function hittaTestInfo(ss, svarssheetNamn) {
  var reg  = ss.getSheetByName("Testregister");
  if (!reg) { Logger.log("Fliken Testregister hittades inte"); return null; }
  var data = reg.getDataRange().getValues();
  // Rad 1 = rubrik, kolumn D (index 3) = Svarssheet
  Logger.log("Söker efter svarssheet: '" + svarssheetNamn + "'");
  for (var i = 1; i < data.length; i++) {
    var cellVarde = data[i][3] ? data[i][3].toString().trim() : "";
    Logger.log("Testregister rad " + (i+1) + " kolumn D: '" + cellVarde + "'");
    if (cellVarde === svarssheetNamn.trim()) {
      return { testId: data[i][0], omrade: data[i][1], facitflik: data[i][2] };
    }
  }
  return null;
}

function hittaElev(ss, email) {
  var elever = ss.getSheetByName("Elever");
  if (!elever) return null;
  var data   = elever.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase().trim() === email) {
      return { email: data[i][0], namn: data[i][1], token: data[i][2] || "", rad: i + 1 };
    }
  }
  return null;
}

function hittaElevMedToken(ss, token) {
  var elever = ss.getSheetByName("Elever");
  if (!elever) return null;
  var data   = elever.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === token) {
      return { email: data[i][0], namn: data[i][1], token: data[i][2] };
    }
  }
  return null;
}

function hamtaFacit(ss, facitflik) {
  var sheet = ss.getSheetByName(facitflik);
  if (!sheet) return [];
  // Kolumn A = frågenummer, kolumn B = rätt svar
  return sheet.getDataRange().getValues().map(function (r) { return r[1]; });
}

function hamtaSvar(svarSheet, rad, antal) {
  // Kolumn A = tidstämpel, B = email, C+ = svar
  return svarSheet.getRange(rad, 3, 1, antal).getValues()[0];
}

// Returnerar bästa resultat per test för en elev (används av webbappen)
// RESULTAT_LOGG: A=Tidsstämpel, B=Email, C=TestID, D=Område, E=Procent, F=Rätta, G=Totalt
function hamtaElevResultat(ss, email) {
  var logg = ss.getSheetByName("RESULTAT_LOGG");
  if (!logg) return [];
  var data  = logg.getDataRange().getValues();
  var basta = {};

  for (var i = 1; i < data.length; i++) {
    if (!data[i][1] || data[i][1].toString().toLowerCase().trim() !== email) continue;
    var testId  = data[i][2];
    var procent = parseInt(data[i][4]) || 0;
    if (!basta[testId] || procent > basta[testId].procent) {
      basta[testId] = {
        testId:  testId,
        omrade:  data[i][3],
        datum:   data[i][0],
        procent: procent,
        ratt:    data[i][5],
        totalt:  data[i][6]
      };
    }
  }

  return Object.keys(basta).map(function (k) { return basta[k]; });
}

// Returnerar aggregerad data för hela klassen (används av klassöversikts-webbappen)
function hamtaKlassData(ss) {
  var reg     = ss.getSheetByName("Testregister");
  var regData = reg ? reg.getDataRange().getValues() : [];
  var testOrdning = [];
  var testOmrade  = {};
  for (var i = 1; i < regData.length; i++) {
    var tid = regData[i][0] ? regData[i][0].toString().trim() : "";
    if (tid) {
      testOrdning.push(tid);
      testOmrade[tid] = regData[i][1] ? regData[i][1].toString() : "";
    }
  }

  var elever = ss.getSheetByName("Elever").getDataRange().getValues();
  var logg   = ss.getSheetByName("RESULTAT_LOGG").getDataRange().getValues();

  // basta[email][testId] = bästa procent
  var basta = {};
  for (var i = 1; i < logg.length; i++) {
    var email   = logg[i][1] ? logg[i][1].toString().toLowerCase().trim() : "";
    var testId  = logg[i][2] ? logg[i][2].toString() : "";
    var procent = parseInt(logg[i][4]) || 0;
    if (!email || !testId) continue;
    if (!basta[email]) basta[email] = {};
    if (!basta[email][testId] || procent > basta[email][testId]) {
      basta[email][testId] = procent;
    }
  }

  var studenter = [];
  for (var i = 1; i < elever.length; i++) {
    var email = elever[i][0] ? elever[i][0].toString().toLowerCase().trim() : "";
    var namn  = elever[i][1] ? elever[i][1].toString() : "";
    if (!email || !namn) continue;

    var elevData = basta[email] || {};
    var testResultat = {};
    var sum = 0, count = 0;

    for (var t = 0; t < testOrdning.length; t++) {
      var tid = testOrdning[t];
      var v   = elevData[tid] !== undefined ? elevData[tid] : null;
      testResultat[tid] = v;
      if (v !== null) { sum += v; count++; }
    }

    studenter.push({
      namn:         namn,
      testResultat: testResultat,
      totalt:       count > 0 ? Math.round(sum / count) : null
    });
  }

  // Klassmedel per test
  var klassStats = {};
  for (var t = 0; t < testOrdning.length; t++) {
    var tid = testOrdning[t];
    var s = 0, c = 0;
    for (var j = 0; j < studenter.length; j++) {
      var v = studenter[j].testResultat[tid];
      if (v !== null) { s += v; c++; }
    }
    klassStats[tid] = c > 0 ? Math.round(s / c) : null;
  }

  var ktSum = 0, ktCount = 0;
  for (var j = 0; j < studenter.length; j++) {
    if (studenter[j].totalt !== null) { ktSum += studenter[j].totalt; ktCount++; }
  }

  return {
    tester:     testOrdning,
    testOmrade: testOmrade,
    studenter:  studenter,
    klassStats: klassStats,
    klassTotal: ktCount > 0 ? Math.round(ktSum / ktCount) : null
  };
}

// ---------------------------------------------------------------------------
// Rättning
// ---------------------------------------------------------------------------

function ratta(svar, facit) {
  var ratt = 0;
  for (var i = 0; i < facit.length; i++) {
    var elevSvar  = svar[i] !== undefined ? svar[i].toString().trim().toLowerCase() : "";
    var rattSvar  = facit[i] !== undefined ? facit[i].toString().trim().toLowerCase() : "";
    var korrekt   = elevSvar === rattSvar;
    Logger.log("F" + (i+1) + ": elev='" + elevSvar + "' facit='" + rattSvar + "' → " + (korrekt ? "RÄTT" : "FEL"));
    if (korrekt) ratt++;
  }
  var procent = facit.length > 0 ? Math.round((ratt / facit.length) * 100) : 0;
  return { ratt: ratt, totalt: facit.length, procent: procent };
}

// ---------------------------------------------------------------------------
// Loggning
// RESULTAT_LOGG: A=Tidsstämpel, B=Email, C=TestID, D=Område, E=Procent, F=Rätta, G=Totalt
// ---------------------------------------------------------------------------

function loggaResultat(ss, timestamp, email, testId, omrade, rattning) {
  var logg = ss.getSheetByName("RESULTAT_LOGG");
  if (!logg) return;

  // Räkna antal tidigare försök på detta test för denna elev
  var data  = logg.getDataRange().getValues();
  var forsok = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().toLowerCase().trim() === email &&
        data[i][2].toString() === testId.toString()) {
      forsok++;
    }
  }

  logg.appendRow([
    timestamp, email, testId, omrade,
    rattning.procent, rattning.ratt, rattning.totalt, forsok + 1
  ]);
}

// ---------------------------------------------------------------------------
// Välkomstmejl (skickas bara första gången)
// ---------------------------------------------------------------------------

function skickaValkommen(email, namn, token) {
  var lank = WEBAPP_URL + "?token=" + token;
  GmailApp.sendEmail(email, "Dina resultat på självtester", "", {
    htmlBody:
      "<p>Hej " + namn + "!</p>" +
      "<p>Du kan följa dina resultat på:</p>" +
      "<p><a href='" + lank + "' style='font-size:16px;'>📊 Öppna min resultatsida</a></p>" +
      "<p>Spara gärna länken – den fungerar för alla dina inlämningar.</p>"
  });
}

// Skickar om resultatlänken till alla elever i Elever-fliken.
// Kör den här om elever fått en trasig länk (t.ex. när WEBAPP_URL saknades).
function skickaOmLankar() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var elever = ss.getSheetByName("Elever").getDataRange().getValues();
  var skickade = 0;
  for (var i = 1; i < elever.length; i++) {
    var email = elever[i][0] ? elever[i][0].toString().toLowerCase().trim() : "";
    var namn  = elever[i][1] ? elever[i][1].toString() : "";
    var token = elever[i][2] ? elever[i][2].toString() : "";
    if (!email || !namn || !token) continue;
    skickaValkommen(email, namn, token);
    skickade++;
    Utilities.sleep(300);
  }
  SpreadsheetApp.getUi().alert("Klart! Länk skickad till " + skickade + " elever.");
}

// ---------------------------------------------------------------------------
// Klassöversikt
// Byggs om helt från RESULTAT_LOGG vid varje anrop – robust och alltid korrekt
// ---------------------------------------------------------------------------

function uppdateraKlassoversikt(ss) {
  var oversikt = ss.getSheetByName("KLASSÖVERSIKT");
  if (!oversikt) oversikt = ss.insertSheet("KLASSÖVERSIKT");

  // Hämta testlista i ordning från Testregister
  var reg     = ss.getSheetByName("Testregister");
  var regData = reg ? reg.getDataRange().getValues() : [];
  var testOrdning = [];
  for (var i = 1; i < regData.length; i++) {
    var tid = regData[i][0] ? regData[i][0].toString().trim() : "";
    if (tid) testOrdning.push(tid);
  }

  var elever = ss.getSheetByName("Elever").getDataRange().getValues();
  var logg   = ss.getSheetByName("RESULTAT_LOGG").getDataRange().getValues();

  // Samla bästa resultat per elev och test (kolumn E = procent)
  var basta = {};
  for (var i = 1; i < logg.length; i++) {
    var email   = logg[i][1] ? logg[i][1].toString().toLowerCase().trim() : "";
    var testId  = logg[i][2] ? logg[i][2].toString() : "";
    var procent = parseInt(logg[i][4]) || 0;
    if (!email || !testId) continue;
    if (!basta[email]) basta[email] = {};
    if (!basta[email][testId] || procent > basta[email][testId]) {
      basta[email][testId] = procent;
    }
  }

  // Rubrikrad: Namn | Självtest 1 | Självtest 2 | … | Medel
  var rubrik = ["Namn"].concat(testOrdning).concat(["Medel"]);
  var rader  = [rubrik];

  for (var i = 1; i < elever.length; i++) {
    var email = elever[i][0] ? elever[i][0].toString().toLowerCase().trim() : "";
    var namn  = elever[i][1] ? elever[i][1].toString() : "";
    if (!email || !namn) continue;

    var elevData = basta[email] || {};
    var rad = [namn];
    var sum = 0, count = 0;

    for (var t = 0; t < testOrdning.length; t++) {
      var v = elevData[testOrdning[t]] !== undefined ? elevData[testOrdning[t]] : "";
      rad.push(v);
      if (v !== "") { sum += v; count++; }
    }
    rad.push(count > 0 ? Math.round(sum / count) : "");
    rader.push(rad);
  }

  // Skriv allt i ett anrop
  oversikt.clearContents();
  oversikt.clearFormats();
  oversikt.getRange(1, 1, rader.length, rubrik.length).setValues(rader);

  // Rubrikrad och sista kolumnen (Medel) i fetstil
  oversikt.getRange(1, 1, 1, rubrik.length).setFontWeight("bold");
  oversikt.getRange(1, rubrik.length, rader.length, 1).setFontWeight("bold");

  // Färgkodning av celler med procentvärden
  for (var r = 2; r <= rader.length; r++) {
    for (var c = 2; c <= rubrik.length; c++) {
      var v = rader[r - 1][c - 1];
      if (v === "" || isNaN(v)) continue;
      oversikt.getRange(r, c).setBackground(beraknaFarg(v));
    }
  }
}

// ---------------------------------------------------------------------------
// Färgkodning
// ---------------------------------------------------------------------------

function beraknaFarg(procent) {
  if (procent < 50) return FARG_ROD;
  if (procent < 75) return FARG_GUL;
  return FARG_GRON;
}

// Kör den här för att uppdatera KLASSÖVERSIKT-fliken manuellt
function uppdateraKlassoversiktManuellt() {
  uppdateraKlassoversikt(SpreadsheetApp.getActiveSpreadsheet());
}
