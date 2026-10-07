// ── Instructions sheet (Swedish user guide) ──────────────────────────────────
// Created by onOpen when missing. Run buildInstructionsSheet manually from the
// Apps Script editor to regenerate it after changing the text.

var INSTRUCTIONS_SHEET = 'Instruktioner';

function buildInstructionsSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(INSTRUCTIONS_SHEET) || ss.insertSheet(INSTRUCTIONS_SHEET, 0);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();

  var row = 1;
  var CHARS_PER_LINE = 95;

  function height(text) { return Math.max(22, Math.ceil(text.length / CHARS_PER_LINE) * 20); }
  function spacer(h) { sheet.setRowHeight(row, h); row++; }

  function title(text) {
    sheet.getRange(row, 1, 1, 2).merge().setValue(text)
      .setFontSize(20).setFontWeight('bold').setBackground(COLOR_HEADER).setFontColor('#FFFFFF')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(row, 60);
    row++;
    spacer(8);
  }
  function section(text) {
    spacer(10);
    sheet.getRange(row, 1, 1, 2).merge().setValue(text)
      .setFontSize(13).setFontWeight('bold').setBackground(COLOR_SUBTLE).setFontColor(COLOR_ENERGY)
      .setVerticalAlignment('middle');
    sheet.setRowHeight(row, 30);
    row++;
  }
  function paragraph(text) {
    sheet.getRange(row, 1, 1, 2).merge().setValue(text)
      .setFontSize(11).setFontColor('#374151').setWrap(true).setVerticalAlignment('top');
    sheet.setRowHeight(row, height(text));
    row++;
  }
  function numbered(num, text) {
    sheet.getRange(row, 1).setValue(num)
      .setFontSize(11).setFontWeight('bold').setFontColor('#1A56DB')
      .setHorizontalAlignment('center').setVerticalAlignment('top');
    sheet.getRange(row, 2).setValue(text)
      .setFontSize(11).setFontColor('#374151').setWrap(true).setVerticalAlignment('top');
    sheet.setRowHeight(row, height(text));
    row++;
  }
  // A bold label followed by an explanation, in one merged cell.
  function item(label, text) {
    var full = label + ' — ' + text;
    var rich = SpreadsheetApp.newRichTextValue().setText(full)
      .setTextStyle(0, label.length, SpreadsheetApp.newTextStyle()
        .setBold(true).setForegroundColor(COLOR_ENERGY).build())
      .build();
    sheet.getRange(row, 1, 1, 2).merge();
    sheet.getRange(row, 1).setRichTextValue(rich);
    sheet.getRange(row, 1, 1, 2).setFontSize(11).setWrap(true).setVerticalAlignment('top');
    sheet.setRowHeight(row, height(full));
    row++;
  }
  function qa(question, answer) {
    sheet.getRange(row, 1, 1, 2).merge().setValue(question)
      .setFontSize(11).setFontWeight('bold').setFontColor(COLOR_ENERGY);
    sheet.setRowHeight(row, 22);
    row++;
    paragraph(answer);
    spacer(6);
  }

  // ── Content ──
  title(ORG_NAME + ' — Laddboxrapport');

  paragraph('Det här kalkylarket beräknar vad varje lägenhet ska betala för laddning av elbil. '
    + 'Det hämtar laddningsdata från Zaptec och elpriser per kvart, och skapar rapporter för styrelsen och revisorn.');

  section('Vad gör verktyget?');
  numbered('1.', 'Hämtar alla avslutade laddningssessioner för föreningens laddboxar från Zaptec, med förbrukning mätt var 15:e minut.');
  numbered('2.', 'Hämtar spotpriset för el i prisområde SE3 för varje kvart. Priserna kommer från Nord Pool. '
    + 'Dagar som är äldre än cirka två månader hämtas från elprisetjustnu.se, eftersom Nord Pool tar betalt för äldre data.');
  numbered('3.', 'Multiplicerar energin i varje kvart med kvartens spotpris, plus överföringsavgift, energiskatt, påslag från leverantör och påslag från förening, och lägger på moms.');
  numbered('4.', 'Summerar kostnaden per lägenhet och skriver rapporterna.');

  section('Så här kör du en rapport');
  numbered('1.', 'Klicka på menyn Beräkna högst upp i kalkylarket.');
  numbered('2.', 'Välj förra månadens rapport, senaste kvartalets rapport eller en valfri period.');
  numbered('3.', 'Ange e-postadress och lösenord till Zaptec Portal i rutan som öppnas. Tryck Enter eller klicka på Hämta. '
    + 'Lösenordet sparas inte, men e-postadressen fylls i automatiskt nästa gång.');
  numbered('4.', 'Vänta medan data hämtas. Det tar oftast under en minut. När det är klart visas en sammanfattning, och eventuella varningar, i samma ruta.');
  paragraph('Tips: kör månadsrapporten några dagar in i den nya månaden. Laddboxar som tillfälligt varit offline hinner då skicka in sina sessioner.');

  section('Vilka flikar skapas?');
  item('Detaljerad rapport', 'förbrukning och kostnad per lägenhet, grupperad per elmätare, med antal sessioner, snittpris och första/senaste laddning. För styrelsens översikt.');
  paragraph('Laddboxarna kan sitta på flera elmätare. Längst ner i Detaljerad rapport finns Månadsstatistik per elmätare med en rad per månad och mätare: '
    + 'laddad energi och snittpris (spot + påslag från leverantör, exkl. moms). Jämför med elräkningen för samma månad och mätare. '
    + 'Laddboxarnas förbrukning ska vara mindre än mätarens, och räkningens snittpris bör ligga nära periodens snittpris.');
  item('Förenklad rapport',  'lägenhetsnummer, kWh och kostnad exkl. moms. Skickas till revisorn för fakturering.');
  item('Sessioner',          'varje enskild laddning med tider, förbrukning, kostnad och eventuell notering. Användbar för att kontrollera en enskild lägenhet.');
  item('Spotpriser',         'alla kvartspriser för perioden och vilken källa varje pris kom från.');
  item('Prisarkiv (dold)',   'sparade spotpriser, så att de inte behöver hämtas igen. Fylls automatiskt. Ändra inget här.');

  section('Hur perioden räknas');
  paragraph('En laddning räknas till den period då den avslutades. En laddning som startar 31 mars kl. 23 och slutar 1 april kl. 02 hamnar alltså i aprilrapporten. '
    + 'Varje kvart prissätts ändå med sitt eget spotpris, så timmen i mars får marspris. Ingen laddning räknas dubbelt eller faller bort mellan två rapporter.');

  section('Inställningar');
  paragraph('I fliken Inställningar finns fem värden som påverkar priset. Alla anges exklusive moms.');
  item('transmission_fee', 'överföringsavgift i öre/kWh som nätägaren tar ut av föreningen. Värdet står på nätägarens faktura. Uppdatera om elnätsavtalet ändras.');
  item('energy_tax',       'energiskatt i öre/kWh. Ändras varje år. Uppdatera i januari enligt Skatteverket (länk finns i fliken).');
  item('supplier_markup',    'påslag från leverantör i öre/kWh. Står på elhandelsfakturan.');
  item('association_markup', 'påslag från förening i öre/kWh, som föreningen själv lägger på laddningen. 0 om inget påslag.');
  item('vat_factor',       'momsfaktor. 1,25 betyder 25 % moms.');
  paragraph('Ändringar gäller från nästa rapportkörning.');

  section('Varningar och noteringar');
  item('Ej prissatt (kWh)', 'energi som saknar spotpris ingår inte i kostnaden. Raden markeras gul och det syns i sammanfattningen. Kör rapporten igen senare.');
  item('Gula rader i Sessioner', 'laddningen har en notering som bör kontrolleras, till exempel att laddboxen varit offline eller saknat mätpunkter. '
    + 'Energin har då fördelats jämnt över tiden mellan mätpunkterna.');
  item('Spotpriser saknas', 'ingen av priskällorna hade priser för dagen. Försök igen senare.');

  section('Vanliga frågor');
  qa('En lägenhet saknas i rapporten. Varför?',
     'Laddboxen har inte avslutat någon laddning under perioden, eller så saknas den i koppling mellan laddbox och lägenhet i scriptet. Kontakta den som administrerar scriptet om en aktiv laddbox saknas.');
  qa('Varför skiljer sig siffrorna från Zaptec-portalen?',
     'Små skillnader beror på avrundning. Större skillnader beror oftast på att laddningar räknas till den period då de avslutades.');
  qa('Vad händer om energiskatten ändras mitt i ett kvartal?',
     'Verktyget använder samma energiskatt för hela perioden. Kör då månadsrapporter och ändra värdet i Inställningar mellan körningarna.');
  qa('Kan jag spara en gammal rapport?',
     'Ja. Varje körning skriver över rapportflikarna, så kopiera kalkylarket (Arkiv → Skapa en kopia) innan du kör en ny rapport om du vill behålla den gamla.');

  section('Felsökning');
  item('Inloggningen misslyckades', 'kontrollera e-postadress och lösenord. Det är samma uppgifter som du loggar in med på portal.zaptec.com.');
  item('Fel vid hämtning av sessioner', 'Zaptec kan vara tillfälligt otillgängligt. Vänta några minuter och försök igen.');

  section('Tekniska detaljer');
  paragraph('Scriptet finns under Tillägg → Apps Script. Det är skrivet i JavaScript och uppdelat i flera filer. '
    + 'Kontakta styrelsen om något inte fungerar eller behöver ändras.');

  // ── Layout ──
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 720);
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(1);

  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });
  sheet.protect().setDescription('Instruktioner — bara läsning').setWarningOnly(true);
}
