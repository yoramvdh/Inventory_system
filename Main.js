/* eslint-disable require-jsdoc */
/* eslint-disable no-unused-vars */
/* eslint-disable max-len */
// project: Voorraadbeheersysteem
// functie: Een semi-automatisch voorraadbeheersysteem.
// Deze applicatie is ontwikkeld voor het anatomo-pathologisch labo van AZ Zeno.
// Naam: Yoram Vandenhouwe
// Datum  start aanmaak: 13/02/2024
// Datum einde aanmaak:
// Versie: 0.1


// vult de  huidige datum in de gegeven cell
function AddDate(sheet, rij, kolom) {
  var time = new Date();
  time = Utilities.formatDate(time, 'GMT+02:00', 'dd/MM/yy');
  sheet.getRange(rij, kolom).setValue(time);
}

// spoort via een trigger vervallen producten op en verplaatst een deze in het tabblad vervallen reagentia, stuurt ook mail naar alle MLT
function VervallenProducten() {
  // declaratie
  var voorraadbeheer = SpreadsheetApp.getActive().getSheetByName('voorraadbeheer');
  var vervallenReagentia = SpreadsheetApp.getActive().getSheetByName('vervallen reagentia');
  var row = 3; // begin van de tabel
  var vervallen = 0; // vervaldatum
  var huidigeCell = voorraadbeheer.getRange(row, 1).getValue();

  // zolang de rij niet leeg is overloopt de functie de tabel en vergelijk de datum met de vervaldatum

  while ( huidigeCell != '') {
    var houdbaarheidsDatum = voorraadbeheer.getRange(row, 9).getValue();
    var legecell =voorraadbeheer.getRange(rowv, 11).getValue();
    // als de houdbaarheid vervallen is
    if (houdbaarheidsDatum == vervallen && legecell == '' ) {
      AddDate(voorraadbeheer, row, 12); // plaatsen we in kolom 12 de datum

      // kopier de lijn en plaats deze is de sheet vervallen reagentia
      var vervallenLijn = voorraadbeheer.getRange(row, 1, 1, 15);
      var destRange = vervallenReagentia.getRange(vervallenReagentia.getLastRow()+1, 1);
      vervallenLijn.copyTo(destRange, {contentsOnly: false});
      vervallenLijn.clear();

      // stuurt een email als er een product vervallen is
      MailApp.sendEmail({to: 'yoram.vandenhouwe@azzenopathologie.net',
        subject: 'automatische mail-Vervallen product',
        htmlBody: 'Zie tablad Vervallen Producten   https://docs.google.com/spreadsheets/d/1e5JcF0UiyphQIlexmdSrpDlXzzvxRfpcHvutMr6iB14/edit#gid=1452909895 ',
      });
    }
    row = row +1; // ga naar de volgende niet lege rij
    huidigeCell = voorraadbeheer.getRange(row, 1).getValue();
  }

  var range = voorraadbeheer.getRange('A3:01100');
  range.sort(7); // op welke colom er gesorteerd wordt
}


