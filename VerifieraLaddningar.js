function kontrolleraAllaSessioner() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sessionSheet = ss.getSheetByName("Sessioner");
  var spotSheet = ss.getSheetByName("Spotpriser");
  
  if (!sessionSheet || !spotSheet) {
    SpreadsheetApp.getUi().alert("Hittade inte flikarna 'Sessioner' eller 'Spotpriser'!");
    return;
  }
  
  var sessionData = sessionSheet.getDataRange().getValues();
  var spotData = spotSheet.getDataRange().getValues();
  
  // 1. Bygg en snabbkarta över spotpriserna för kontroll
  var spotMap = {};
  for (var i = 1; i < spotData.length; i++) {
    var rawTime = spotData[i][0]; // Kolumn A (Tid)
    var pris = spotData[i][2];    // Kolumn C (Pris inkl. påslag+moms)
    if (rawTime) {
      // Formatera till en ren textnyckel för exakt matchning
      var formatKey = Utilities.formatDate(new Date(rawTime), 'Europe/Stockholm', "yyyy-MM-dd HH:mm");
      spotMap[formatKey] = pris;
    }
  }
  
  // 2. Skapa en ny flik för testresultatet om den inte finns
  var testSheet = ss.getSheetByName("TEST_RESULTAT");
  if (!testSheet) {
    testSheet = ss.insertSheet("TEST_RESULTAT");
  } else {
    testSheet.clear();
  }
  
  // Skriv rubriker till testfliken
  testSheet.appendRow(["Rad i Sessioner", "Starttid", "Lägenhet", "Debiterat Pris", "Förväntat Startpris", "Status"]);
  
  var felaktiga = 0;
  var korrekta = 0;
  
  // 3. Loopa igenom alla sessioner och jämför priser
  for (var j = 2; j < sessionData.length; j++) { // Startar på rad 3 (hoppar över rubriker)
    var radNummer = j + 1;
    var startTid = sessionData[j][0];   // Kolumn A
    var lgh = sessionData[j][2];        // Kolumn C
    var debiteratPris = parseFloat(sessionData[j][6]); // Kolumn G (Snittpris)
    var kwh = parseFloat(sessionData[j][4]); // Kolumn E (kWh)
    
    if (!startTid || isNaN(debiteratPris) || kwh === 0) continue;
    
    // Runda av starttiden nedåt till närmaste kvart för att efterlikna Plan B
    var d = new Date(startTid);
    var minuter = d.getMinutes();
    var avrundadeMinuter = Math.floor(minuter / 15) * 15;
    d.setMinutes(avrundadeMinuter);
    d.setSeconds(0);
    
    var checkKey = Utilities.formatDate(d, 'Europe/Stockholm', "yyyy-MM-dd HH:mm");
    var förväntatPris = spotMap[checkKey];
    
    if (förväntatPris !== undefined) {
      // Vi kollar om det debiterade priset matchar starttidens pris på öret
      var diff = Math.abs(debiteratPris - förväntatPris);
      
      if (diff < 0.05) {
        // Priset är EXAKT starttidens pris. Det betyder att Plan A MISSLYCKADES och Plan B kördes!
        testSheet.appendRow([radNummer, startTid, lgh, debiteratPris, förväntatPris, "⚠️ FEL (Låst till startpris)"]);
        felaktiga++;
      } else {
        // Priset är annorlunda, vilket betyder att den antingen räknat rätt eller kört den trasiga skalningsfaktorn
        testSheet.appendRow([radNummer, startTid, lgh, debiteratPris, förväntatPris, "Kvartssnitt eller Skalning använd"]);
        korrekta++;
      }
    }
  }
  
  // Snygga till fliken lite
  testSheet.getFilter() || testSheet.getDataRange().createFilter();
  testSheet.autoResizeColumns(1, 6);
  
  SpreadsheetApp.getUi().alert("Analysen är klar!\n\nAntal laddningar låsta till startpris: " + felaktiga + "\nÖvriga/Skalade: " + korrekta + "\n\nSe detaljer i fliken 'TEST_RESULTAT'.");
}