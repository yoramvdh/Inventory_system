/* eslint-disable require-jsdoc */
/* eslint-disable no-unused-vars */
/* eslint-disable max-len */
// project: Inventorymanagement system
// functie: Een semi-automatic Inventorymanagement system.
// This application is develloped for the pathology labo fo AZ Zeno.
// Name: Yoram Vandenhouwe
// Start of project: 13/02/2024
// Version: 0.1

/* Declaration*/

let row = 2; // start of the table
// Get SpreadsheetUrl
const sheetUrl = SpreadsheetApp.getActive().getUrl();
// Get all the sheets
const sheets = SpreadsheetApp.getActive().getSheets();
const voorraadbeheer = sheets[0];
const vervallenReagentia = sheets[3];
const opgebruikteReagentia = sheets[2];

// Create array to store all the links
const links = [];
// For each sheet in sheets add an array element to our array with the
// string of the URL for that sheet
sheets.forEach((sheet)=>links.push(sheetUrl+'#gid='+sheet.getSheetId()));
const voorraadbeheerlink = links[0];
const vervallenReagentialink = links[3];

const maxrange = voorraadbeheer.getRange('A3:01100');

// Function to add the current date in a given cell.
function addDate(sheet, row, column) {
  let time = new Date();
  time = Utilities.formatDate(time, 'GMT+02:00', 'dd/MM/yy');
  sheet.getRange(row, column).setValue(time);
}

// Add a custom menu to the active spreadsheet, including a separator and a sub-menu.
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Specialised Functions')
      .addItem('Total Supply', 'totalSupply')
      .addSeparator()
      .addItem('My menu item', 'myFunction')
      .addToUi();
}

/* Activates using a trigger in the Google App Script aplication. If the product is expired moves all data of this product to a separate sheet to store the data.
Then send a mail to all mail adresses in the config sheet. */
function expiredProduct() {
  // declaration
  const expired = 0; // Experationdate
  let currentcell = voorraadbeheer.getRange(row, 1).getValue();

  // While the row is not empty, check each row to see if the product is expired.
  while ( currentcell != '') {
    const experationdate = voorraadbeheer.getRange(row, 9).getValue();
    const emptycell =voorraadbeheer.getRange(row, 11).getValue();
    // If expired.
    if (experationdate == expired && emptycell == '' ) {
      addDate(voorraadbeheer, row, 12); // Add the date in colum 12.

      // Cuts the row and places the data in a new line in sheet 'vervallen reagentia'.
      const expiredProduct = voorraadbeheer.getRange(row, 1, 1, 15);
      const destRange = vervallenReagentia.getRange(vervallenReagentia.getLastRow()+1, 1);
      expiredProduct.copyTo(destRange, {contentsOnly: false});
      expiredProduct.clear();

      // Send a mail if a product is expired.
      MailApp.sendEmail({to: 'yoram.vandenhouwe@azzenopathologie.net',
        subject: 'automatic mail-Expired product',
        htmlBody: 'The product, '+ currentcell + ', has expired and was placed in the tab vervallen reagentia on the last row.For more information use the link:' + vervallenReagentialink,
      });
    }
    row = row +1; // While the row is not empty, check each row to see if the product is expired.
    currentcell = voorraadbeheer.getRange(row, 1).getValue();
  }
  maxrange.sort(7); // Sort complete range based on column.
}

// This function uses a trigger to find all product who are almoust expired and send a mail to specifiek users
function almoustExpiredProducts() {
  const almoustexpired = 14; // number of day befor the product expires
  let currentcell = voorraadbeheer.getRange(row, 1).getValue();

  // As long as the currentcell is not empty the function goes over the table and will compare each time the expirationdate with the number of days till it expires.

  while ( currentcell != '') {
    const expiredate = voorraadbeheer.getRange(row, 9).getValue();
    const alreadyused =voorraadbeheer.getRange(row, 11).getValue();
    if (expiredate == almoustexpired && alreadyused === '' ) { // Checks if the product is almost expired.
      // Sends a mail to the user.
      MailApp.sendEmail({to: 'yoram.vandenhouwe@azzenopathologie.net',
        subject: 'automatische mail- Bijna Vervallen product',
        htmlBody: 'Het product '+currentcell+' op rij '+row+' zal over 14 dagen vervallen.'+ voorraadbeheerlink,
      });
    }
    row = row +1; // Go the the next empty row.
    currentcell = voorraadbeheer.getRange(row, 1).getValue();
  }
}

// This function uses a trigger to find all product used up and moves the data to a seperate sheet.

function usedUp() {
  let currentcell = voorraadbeheer.getRange(row, 1).getValue();

  // As long as the currentcell is not empty the function goes over the table and will compare each time the expirationdate with the number of days till it expires.

  while ( currentcell != '') {
    const alreadyused =voorraadbeheer.getRange(row, 11).getValue();

    if ( alreadyused != '' ) {
      const currentrow = voorraadbeheer.getRange(row, 1, 1, 15);
      const destRange = opgebruikteReagentia.getRange(opgebruikteReagentia.getLastRow()+1, 1);
      currentrow.copyTo(destRange, {contentsOnly: false});
      currentrow.clear();
    }
    row = row +1; // ga naar de volgende niet lege rij
    currentcell = voorraadbeheer.getRange(row, 1).getValue();
  }

  maxrange.sort(7); // sort the whole range on column 7
}


function totalSupply() {
  const selObj = SpreadsheetApp.getActiveSheet().getSelection();
  const selectedText = selObj.getActiveRange().getValue();
  voorraadbeheer.activate();
  let currentcell = voorraadbeheer.getRange(row, 1).getValue();
  if (selectedText != '') {
    let totalitems = 0;
    let currentitem = 0;
    Logger.log(totalitems);
    Logger.log(currentitem);
    while ( currentcell != '') {
      // eslint-disable-next-line prefer-const
      let itemId =voorraadbeheer.getRange(row, 1).getValue();
      if ( itemId == selectedText ) {
        Logger.log(itemId);
        // eslint-disable-next-line prefer-const
        let maxitem = voorraadbeheer.getRange(row, 4).getValue();
        // eslint-disable-next-line prefer-const
        let useditem = voorraadbeheer.getRange(row, 6).getValue();
        currentitem = maxitem- useditem;
        Logger.log(currentitem);
        Logger.log(totalitems);
        totalitems = totalitems + currentitem;
        Logger.log(totalitems);
      }
      row = row +1; // ga naar de volgende niet lege rij
      currentcell = voorraadbeheer.getRange(row, 1).getValue();
    }
    SpreadsheetApp.getUi().alert('The total number of item '+selectedText+' is: '+ totalitems);
  } else {
    SpreadsheetApp.getUi().alert('Select the item you want to check the inventory for out of the list of in the minimum voorraad tab. Then run the function again.');
  }
}


function onEdit(e) {
  // declaration
  // the follow code is execute each time a change was made in google sheets
  // eslint-disable-next-line prefer-const
  let activerange = e.range;
  // eslint-disable-next-line prefer-const
  let activerow = activerange.getRow();
  // eslint-disable-next-line prefer-const


  // Check on edit if the number of items in that row is equal to the max number of items in that row. If this maches then a the current date is placed in column 11 ( "opgebruikt")
  if (SpreadsheetApp.getActiveSheet().getName() =='voorraadbeheer') {
    const maxitems = voorraadbeheer.getRange(activerow, 4).getValue();
    const currentstock = voorraadbeheer.getRange(activerow, 6).getValue();
    const emptycell =voorraadbeheer.getRange(activerow, 11).getValue();
    if (maxitems == currentstock && maxitems != 0 && emptycell == '') {
      addDate(voorraadbeheer, row, 11);
    }
  }
  // Checks on the active line when a product is used for the first time, when it is: add date in column 10 "ingebruikname"
  if (SpreadsheetApp.getActiveSheet().getName() =='voorraadbeheer') {
    const huidigAantal = voorraadbeheer.getRange(row, 6).getValue();
    const emptycell =voorraadbeheer.getRange(row, 10).getValue();

    if (huidigAantal != 0 && emptycell == '' ) {
      addDate(voorraadbeheer, row, 10);
    }
  }
}
