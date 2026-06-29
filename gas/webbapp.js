/**
 * Webbapp – elevens resultatsida
 *
 * Publicera via: Distribuera → Ny distribution → Webbapp
 *   Kör som: Mig (din Google-konto)
 *   Vem har åtkomst: Alla (så att eleverna kan öppna utan inloggning)
 *
 * Kopiera den publicerade URL:en till WEBAPP_URL i resultathantering.js
 */

function doGet(e) {
  var view  = e && e.parameter && e.parameter.view  ? e.parameter.view  : "";
  var token = e && e.parameter && e.parameter.token ? e.parameter.token : "";

  if (view === "klass") {
    return visaKlassoversikt();
  }

  if (!token) {
    return felSida("Ogiltig länk – ingen token angiven.");
  }

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var elevInfo = hittaElevMedToken(ss, token);

  if (!elevInfo) {
    return felSida("Länken känns inte igen. Kontrollera att du använder rätt länk.");
  }

  var resultat = hamtaElevResultat(ss, elevInfo.email);

  var template = HtmlService.createTemplateFromFile("webbapp1");
  template.namn    = elevInfo.namn;
  template.resultat = resultat;

  return template.evaluate()
    .setTitle(elevInfo.namn + " – Resultat på självtester")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function visaKlassoversikt() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var klassData = hamtaKlassData(ss);
  var template  = HtmlService.createTemplateFromFile("klassoversikt");
  template.klassData = klassData;
  return template.evaluate()
    .setTitle("Klassöversikt – Historia Hi1B")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function felSida(meddelande) {
  var html = "<html><body style='font-family:sans-serif;padding:40px;'>" +
             "<h2>⚠️ " + meddelande + "</h2>" +
             "<p>Kontakta din lärare om problemet kvarstår.</p>" +
             "</body></html>";
  return HtmlService.createHtmlOutput(html).setTitle("Fel");
}
