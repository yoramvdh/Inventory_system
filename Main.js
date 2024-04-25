/* eslint-disable prefer-const */
/* eslint-disable require-jsdoc */
/* eslint-disable no-unused-vars */
/* eslint-disable max-len */
// Project: Inventorymanagement system
// Function: a semi-automatic Inventorymanagement system.
// This application is developed for the pathology labo of AZ Zeno.
// Name: Yoram Vandenhouwe
// Start of project: 13/02/2024
// Version: 1

/* Declaration*/

let row = 2; // Start of the table.
// Get SpreadsheetUrl.
const sheetUrl = SpreadsheetApp.getActive().getUrl();
// Get all the sheets in order.
const sheets = SpreadsheetApp.getActive().getSheets();
const voorraadbeheer = sheets[0];
const minimumVoorraad = sheets[1];
const teBestellen = sheets[2];
const besteld = sheets[3];
const opgebruikteReagentia = sheets[4];
const vervallenReagentia = sheets[5];
const statistieken = sheets[6];
const configuration = sheets[7];
const jaarrapporten = sheets[8];

// Create array to store all the links.
const links = [];
// For each sheet in the spreadsheet add an array element to our array with the string of the URL for that sheet.
sheets.forEach((sheet)=>links.push(sheetUrl+'#gid='+sheet.getSheetId()));
const voorraadbeheerlink = links[0]; // Link to the voorraadbeheer tab.
const vervallenReagentialink = links[5]; // Link to the vervallen reagentia tab.

const maxrange = voorraadbeheer.getRange('A3:01100'); // Total range to sort items in the voorraadbeheer tab.

// Function to add the current date in a given cell.
function addDate(sheet, row, column) {
  let time = new Date(); // Create a new date variable.
  time = Utilities.formatDate(time, 'GMT+02:00', 'dd/MM/yy'); // Get local time in dd/MM/yy format.
  sheet.getRange(row, column).setValue(time); // Add date to given cell.
}

// Add a custom menu to the active spreadsheet, including a separator and a submenu. This is added when opening the spreadsheet.
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Specialised Functions') // Create a new option in the menu.
      .addItem('Total Supply of one item', 'totalSupply') // Add an item with the name of the function and then the link to the funtion.
      .addSeparator() // Adds a line between functions.
      .addItem('makeYearRapport', 'makeYearRapport')
      .addSeparator()
      .addItem('bestel', 'orderItems')
      .addSeparator()
      .addItem('toegekomen', 'itemArrived')
      .addSeparator()
      .addItem('buildGraphYear', 'buildGraphYear')
      .addToUi();
}

// Creates a series of functionalities where the conditions are checked each time a change is made to the spreadsheet.
function onEdit(e) {
  // Declaration
  let activerange = e.range; // Selects the active range.
  let activerow = activerange.getRow(); // Selects the active range.
  // This checks to see if the active sheet is the besteld tab. If there is an item in the list but no date of the below minimal supply add date in besteld column and message in below minimal supply column.
  if (SpreadsheetApp.getActiveSheet().getName() ==besteld.getName()) {
    const addItem = besteld.getRange(activerow, 2).getValue(); // Retrieves the name of the item in the active row.
    const datePresent =besteld.getRange(activerow, 1).getValue(); // Retrieves the date of the item in the active row.
    if (addItem != '' && datePresent =='') { // If an item is filled in the row but there is no date present:
      addDate(besteld, activerow, 4); // Add date in active row.
      besteld.getRange(activerow, 1).setValue('Besteld voor minimale hoeveelheid overschreden was.'); // Set message in active row where the date of below minimal supply should be.
    }
  }

  // Check on edit if the number of items in that row is equal to the max number of items in that row. If this maches then a the current date is placed in column 11 ( "opgebruikt").
  if (SpreadsheetApp.getActiveSheet().getName() ==voorraadbeheer.getName()) { // Check to see if the active cell is in the voorraadbeheer tab.
    const maxitems = voorraadbeheer.getRange(activerow, 4).getValue(); // Retrieves the total number of orderd items for the current line.
    const currentstock = voorraadbeheer.getRange(activerow, 6).getValue(); // Retrieves the current number of orderd items for the current line.
    const emptycell =voorraadbeheer.getRange(activerow, 11).getValue(); // Retrieves the date for when this item was used up, for the current line.
    if (maxitems == currentstock && maxitems != 0 && emptycell == '') { // If the maximum number of items is equal to the current number of used up items and the date is not already filled in:
      addDate(voorraadbeheer, activerow, 11); // Add the current date.
    }
  }
  // Checks on the active line when a product is used for the first time, when it is: add date in column 10 "ingebruikname".
  if (SpreadsheetApp.getActiveSheet().getName() == voorraadbeheer.getName()) {
    const huidigAantal = voorraadbeheer.getRange(activerow, 6).getValue(); // Retrieves the current number of orderd items for the current line.
    const emptycell =voorraadbeheer.getRange(activerow, 10).getValue(); // Retrieves the date for when this item was first used, for the current line.

    if (huidigAantal != 0 && emptycell == '' ) { // If the number of items is larger then 0 and the date for the first in use is not filled in:
      addDate(voorraadbeheer, activerow, 10); // Add the current date.
    }
  }
}
/* Activates using a trigger in the Google App Script application. If the product is expired, move all data of this product to a separate sheet to store the data.
Then send an email to all mail adresses in the config sheet. */
function expiredProduct() {
  // Declaration
  const expired = 0; // Expiration date.
  let currentcell = voorraadbeheer.getRange(row, 1).getValue(); // Retrieve the name of the item for the current row.

  // While the row is not empty, check each row to see if the product is expired:
  while ( currentcell != '') {
    const experationdate = voorraadbeheer.getRange(row, 9).getValue();
    const emptycell =voorraadbeheer.getRange(row, 11).getValue();
    // If expired there are 0 days let till expiration and the column for expired date is still empty:
    if (experationdate == expired && emptycell == '' ) {
      addDate(voorraadbeheer, row, 12); // Add the date in colum 12.

      // Cuts the row and places the data in a new line in sheet 'vervallen reagentia'.
      const expiredProduct = voorraadbeheer.getRange(row, 1, 1, 15);
      const destRange = vervallenReagentia.getRange(vervallenReagentia.getLastRow()+1, 1);
      expiredProduct.copyTo(destRange, {contentsOnly: false});
      expiredProduct.clear();
      // Get list of emails in the config tab.
      const emailList= getRowOfData(configuration, 8);
      // Send an email if a product is expired.
      MailApp.sendEmail({to: emailList,
        subject: 'automatic mail-Expired product',
        htmlBody: 'The product, '+ currentcell + ', has expired and was placed in the tab vervallen reagentia on the last row. For more information use the link:' + vervallenReagentialink,
      });
    }
    row = row +1; // While the row is not empty, check each row to see if the product is expired.
    currentcell = voorraadbeheer.getRange(row, 1).getValue();
  }
  maxrange.sort(7); // Sort complete range based on column.
}

// This function uses a trigger to find all products that are almost expired and send an email to specific users.
function almoustExpiredProducts() {
  const almoustexpired = 14; // Number of days before the product expires.
  let currentcell = voorraadbeheer.getRange(row, 1).getValue();

  // As long as the current cell is not empty the function goes over the table and will compare each time the expiration date with the number of days till it expires.

  while ( currentcell != '') {
    const expiredate = voorraadbeheer.getRange(row, 9).getValue();
    const alreadyusedUp =voorraadbeheer.getRange(row, 11).getValue();
    if (expiredate == almoustexpired && alreadyusedUp == '' ) { // If the product is 14 days befor expiration.
      // Get list of emails in the config tab.
      const emailList= getRowOfData(configuration, 9);
      // Sends an email to the users.
      MailApp.sendEmail({to: emailList,
        subject: 'automatische mail- Bijna Vervallen product',
        htmlBody: 'Het product '+currentcell+' op rij '+row+' zal over 14 dagen vervallen: '+ voorraadbeheerlink,
      });
    }
    row = row +1; // Go the the next empty row.
    currentcell = voorraadbeheer.getRange(row, 1).getValue();
  }
}

// This function uses a trigger to find all products used up and moves the data to a seperate sheet.

function usedUp() {
  let currentcell = voorraadbeheer.getRange(row, 1).getValue();

  // As long as the current cell is not empty the function loops over the table and will compare each time the expirationdate with the number of days till it expires.
  while ( currentcell != '') {
    const alreadyusedUp =voorraadbeheer.getRange(row, 11).getValue();

    if ( alreadyusedUp != '' ) { // If the product is not already used up:
      const currentrow = voorraadbeheer.getRange(row, 1, 1, 15);
      const destRange = opgebruikteReagentia.getRange(opgebruikteReagentia.getLastRow()+1, 1);
      currentrow.copyTo(destRange, {contentsOnly: false});
      currentrow.clear();
      voorraadbeheer.getRange(row, 9).setValue('=DAYS360(configuration!$B$2,H'+row+')');
    }
    row = row +1; // Go to the next not empty row
    currentcell = voorraadbeheer.getRange(row, 1).getValue();
  }

  maxrange.sort(7); // Sort the whole range on column 7
}

// Calculates the total supply of a selected item.
function totalSupply() {
  // start of itemarray
  const itemarr = [['Item name', 'Lotnumber', 'Expiration date ', 'Number of items']];
  // Get the value of the cell selected in the spreadsheet.
  const selectedText = configuration.getRange(4, 2).getValue();
  // Selects the first tab.
  voorraadbeheer.activate();
  // Get the name of the first item in the list.
  let currentcell = voorraadbeheer.getRange(row, 1).getValue();
  if (selectedText != '') { // Gives an error message if the selected cell is empty.
  // Initiates values.
    let totalitems = 0;
    let currentitem = 0;
    // While the current cell is not empty compare values with the selected text.
    while ( currentcell != '') {
      // eslint-disable-next-line prefer-const
      let itemId =voorraadbeheer.getRange(row, 1).getValue();
      if ( itemId == selectedText ) { // If the current cell has the same value as the selected text.
        // eslint-disable-next-line prefer-const
        // Calculate the current number of items by taking the maximum number of items and substracting the used items.
        const maxitem = voorraadbeheer.getRange(row, 4).getValue();
        // eslint-disable-next-line prefer-const
        let useditem = voorraadbeheer.getRange(row, 6).getValue();
        currentitem = maxitem- useditem;
        // Adding the number of used items in this instance to the number of total items.
        totalitems = totalitems + currentitem;
        const arr = []; // Generates a new array and adds the name of the item
        arr.push(selectedText);
        // Takes the expiration date out of the selected line. Transforms it to string then slices the string to keep relevant information.
        const expirationdateObj = voorraadbeheer.getRange(row, 8).getValue();
        const expirationdateString= expirationdateObj.toString();
        const expirationdate = expirationdateString.slice(0, 16);
        const lotNumber = voorraadbeheer.getRange(row, 3).getValue();
        // Adds items to array.
        arr.push(lotNumber);
        arr.push(expirationdate);
        arr.push(currentitem);
        // Adds the new array to the existing table.
        itemarr.push(arr);
      }
      row = row +1; // Go to the next line.
      currentcell = voorraadbeheer.getRange(row, 1).getValue();
    }
    // This function takes a name of an item, a number of items in total and an array to construct a table.
    createDoc(selectedText, totalitems, itemarr);
  } else { // Gives an error message if the selected cell is empty.
    SpreadsheetApp.getUi().alert('Select the item you want to check the inventory for out of the list of in the minimum voorraad tab. Then run the function again.');
  }
}
// This function take a name of an item, a number of items in total and an array to construct a table.
function createDoc(itemname, totalitems, itemarr) {
  // Style of the title
  const titleStyle = {};
  titleStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  titleStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  titleStyle[DocumentApp.Attribute.FONT_SIZE] = 20;
  titleStyle[DocumentApp.Attribute.BOLD] = true;
  // Style of the basic text
  const textStyle = {};
  textStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  textStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  textStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
  textStyle[DocumentApp.Attribute.BOLD] = false;
  // Style of the solution text
  const solutionStyle = {};
  solutionStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  solutionStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  solutionStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
  solutionStyle[DocumentApp.Attribute.BOLD] = true;
  solutionStyle[DocumentApp.Attribute.UNDERLINE] = true;
  // Style of the table
  const tableStyle = {};
  tableStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  tableStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  tableStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
  tableStyle[DocumentApp.Attribute.BOLD] = false;
  solutionStyle[DocumentApp.Attribute.UNDERLINE] = false;

  // Generates the time from google
  let time = new Date();
  time = Utilities.formatDate(time, 'GMT+02:00', 'dd/MM/yy HH:mm');

  // Generate the name of the document and get the ID
  const newdoc= DocumentApp.create('Item rapport : '+itemname+' '+ time);
  const docid= newdoc.getId();
  // Open a document by ID.
  const doc = DocumentApp.openById(docid);
  const body = doc.getBody();

  // Append a paragraph to the document body section directly.
  const title =body.appendParagraph(itemname);
  // Apply the custom style.
  title.setAttributes(titleStyle);
  // Creates a horizontal line.
  body.appendHorizontalRule();
  // Generates and adds the main body of text
  const standardtext= body.appendParagraph('This document was automatically generated and uses the data of 2024-Voorraadbeheer to calculate the quantity of the requested product. This rapport was generated on '+time+'.');
  // Apply the custom style.
  standardtext.setAttributes(textStyle);
  // Creates a paragraph break.
  body.appendParagraph('');
  // Generates and adds the solution line.
  const solution =body.appendParagraph('The total number of times item '+itemname+' is available is '+ totalitems+'.');
  // Apply the custom style.
  solution.setAttributes(solutionStyle);

  // Build a table from the array.
  const table =body.appendTable(itemarr);
  // Apply the custom style.
  table.setAttributes(tableStyle);
  // Searches a file in the drive, checks to see if the given foldername exist in the drive, if not, the script creates it, then moves the file in the folder.
  moveFile('Item rapports automatically generated', docid);
}


// Searches a file in the drive, checks to see if the given foldername exist in the drive, if not, the script creates it, then moves the file in the folder.
function moveFile(nameOfDestination, idOfFile) {
  let folderpresent = false; // Initiates parameter to check if folder is available
  let folderid = ''; // Initiates the ID string for the folder

  const folders = DriveApp.getFolders(); // Gets all folders form the users drive

  // Loops over all folders and compares the name to the name given, given as the first param
  while (folders.hasNext()) {
    const folder = folders.next();
    const foldername = folder.getName();

    if (foldername === nameOfDestination) {
    // If the folder is present, change folderpresent too true and get the ID
      folderpresent = true;
      folderid = folder.getId();
    }
  }
  // If the folder does not exist, create the folder and get the ID
  if (folderpresent === false) {
    folderid = DriveApp.createFolder(nameOfDestination).getId();
  }
  // Get the folder and move file to folder.
  const correctfolder = DriveApp.getFolderById(folderid);
  DriveApp.getFileById(idOfFile).moveTo(correctfolder);
  const fileUrl=DriveApp.getFileById(idOfFile).getUrl();

  SpreadsheetApp.getUi().alert('The file is available the '+nameOfDestination+' folder. ', 'This is the link to the file: \n '+fileUrl, SpreadsheetApp.getUi().ButtonSet.OK);
}

// This function gets all the data from 1 row from a selected sheet en concatinates its in a comma sepperated string.
function getRowOfData(sheet, rowNumber) {
  let columnNumber = 2; // Starting with column 2
  // Uses the coördinates to get the chosen value
  let currentcell = sheet.getRange(rowNumber, columnNumber).getValue();
  // Loops over the row untill a cell is empty en concatinates each new item with the previous items.
  let emailList= '';
  while ( currentcell != '') {
    // eslint-disable-next-line prefer-const
    let emailItem = sheet.getRange(rowNumber, columnNumber).getValue();


    emailList= emailList +','+ emailItem;

    columnNumber = columnNumber +1; // go to the next not empty row
    currentcell = sheet.getRange(rowNumber, columnNumber).getValue();
  }
  emailList=emailList.slice(1); // Removes the comma at the start of the string.

  return emailList; // Return the string.
}


// This function loops over a list in the tab minimum voorraad, for each item in the list it loops over the list in the tab voorraadbeheer and calculated the total supplies present for the item in minimum voorraad.
// The calculation is done by subtracting the used items from the total items. Next the item is compared to a list of minimum item kwantities, if the item is lower then the minimum than the te bestellen page and the besteld page are checked to see if the item is present in the first 30 lines. If the item is present then update the current supplies, if not present search the first empty line and add the itemname, the active supply and the current date.
function minimumSupply() {
  let rowMin= 2; // Begin row of the tab minimum voorraad.

  let itemId = minimumVoorraad.getRange(rowMin, 1).getValue(); // The name of the item on specified row on the minimum voorraad tab.

  // As long as the current cell is not empty the function goes over the table in minimum voorraad.
  while ( itemId != '') {
    let row = 2; // Begin row of the tab voorraad.
    // Loop over all items in the tab voorraadbeheer.
    let selectedText = voorraadbeheer.getRange(row, 1).getValue(); // The name of the item on specified row on the voorraadbeheer tab.

    let totalitems = 0; // Set the total counted items to 0.
    while (selectedText!= '') { // Loop over all items in the tab voorraadbeheer untill there is an empty row.
      if (itemId == selectedText) { // If the names of the items in minimum voorraad and in the vooraadbeheer tab match.
        // Calculate the current number of items by taking the maximum number of items and substracting the used items.
        const maxitem = voorraadbeheer.getRange(row, 4).getValue();
        // eslint-disable-next-line prefer-const
        let useditem = voorraadbeheer.getRange(row, 6).getValue();
        currentitem = maxitem- useditem;
        // Added the number of used items in this instance to the number of total items.
        totalitems = totalitems + currentitem;
      }
      row = row +1; // Go to the next not empty row in the voorraadbeheer tab.

      selectedText = voorraadbeheer.getRange(row, 1).getValue(); // Get the next name in the table.
    }
    // Sets the calculated number of supplies in the tabel next to the correct item.
    minimumVoorraad.getRange(rowMin, 3).setValue(totalitems);

    // eslint-disable-next-line prefer-const
    let minimumSupply = minimumVoorraad.getRange(rowMin, 4).getValue(); // Get the minimum supply for a product from the tab minimum voorraad.

    if (totalitems < minimumSupply ) { // If the current value is lower then the minimum supply.
      const resultTeBestellen = itemPresentInList(teBestellen, itemId, 2); // Look if the item is already present in the tab te bestellen and notate the row if so.
      const inListTeBestellen=resultTeBestellen.alreadyInList; // If values was in list then this is true.
      const itemRowTeBestellen=resultTeBestellen.rowWithItem;

      const resultBesteld = itemPresentInList(besteld, itemId, 2); // Look if the item is already present in the tab besteld and notate the row if so.
      const inListBesteld=resultBesteld.alreadyInList; // If values was in list then this is true.
      const itemRowBesteld=resultBesteld.rowWithItem;

      if (inListTeBestellen== true) { // If the item was in the list te bestellen then update the number of items to the current supply.
        teBestellen.getRange(itemRowTeBestellen, 4).setValue(totalitems);
      }
      if (inListBesteld== true) { // If the item was in the list besteld then update the number of items to the current supply.
        besteld.getRange(itemRowBesteld, 4).setValue(totalitems);
      }
      if (inListTeBestellen == false && inListBesteld== false ) { // If it is not present in both lists above then add it to the list of te bestellen items.
        let rowTeBes=2;
        let emptycell = teBestellen.getRange(rowTeBes, 2).getValue();

        while (emptycell!='') { // Find the next empty line of te bestellen.
          rowTeBes = rowTeBes +1;
          emptycell = teBestellen.getRange(rowTeBes, 2).getValue();
        }
        addDate(teBestellen, rowTeBes, 1); // Add the date.
        teBestellen.getRange(rowTeBes, 2).setValue(itemId); // Add the name of the item.
        let itemFirm = minimumVoorraad.getRange(rowMin, 2); // The name of the firm on specified row on the minimum voorraad tab.
        let firmRange = teBestellen.getRange(rowTeBes, 3); // Add firm to the tab te bestellen.
        itemFirm.copyTo(firmRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);


        teBestellen.getRange(rowTeBes, 4).setValue(totalitems); // Add the current number of items.
      }
    }


    rowMin = rowMin +1; // Go to the next not empty row

    itemId = minimumVoorraad.getRange(rowMin, 1).getValue(); // Get the name of the item of the row.
  }
}

// This function takes the name of the tab and the name of an item and the column number. It will then check if it can find the name in the specified column.
// Returns an array: alreadyInList is true when the item is present, false if not. rowWithItem: is 0 when not found, otherwise it will be the row number where the name was found.
function itemPresentInList(sheet, itemId, columnNumber) {
  let alreadyInList = false; // Set the default as false
  let rowWithItem =0; //
  for (let i = 2; i < 30; i++) { // Loop over the first 30 items in the given tab, in the given column.
    const itemInList = sheet.getRange(i, columnNumber).getValue(); // Get the name of the item in the list.
    if (itemInList == itemId) { // Compare the name of the item in the list to the given name in the function.
      alreadyInList = true; // If the same name is found set variables to true and note the number of the line.
      rowWithItem = i;
    }
  }
  return { // Return array with bolean to see if item is present and integer of the row where it was found ( 0 if not found)
    alreadyInList: alreadyInList,
    rowWithItem: rowWithItem,
  };
}

// Checks to see if items are checked of in the tab te bestellen. If checked, places removes the item from the te bestellen tab and places the item in the tab besteld. It also adds the date in the new tab en resets the checkmark in the old tab.
function orderItems() {
  for (let i = 2; i < 30; i++) { // Loop over the first 30 items in the given tab, in the given column.
    // eslint-disable-next-line prefer-const
    let orderd = teBestellen.getRange(i, 5).getValue(); // Get the value of the checkbox of the item in the list.
    if (orderd == true) { // Check if the item's checkbox was checked, if checked proceed with following code, if not go to next line.
      const currentrow = teBestellen.getRange(i, 1, 1, 4); // Get all data of the the current line.

      let rowBes=2; // The first row that will be used.
      let emptycell = besteld.getRange(rowBes, 2).getValue(); // The value of the name of the row.
      while (emptycell!='') { // Find the first empty line in the tab besteld.
        rowBes = rowBes +1;
        emptycell = besteld.getRange(rowBes, 2).getValue();
      }
      let destRange = besteld.getRange(rowBes, 1, 1, 4); // Select the range where an empty cell was found.
      currentrow.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false); // Copy all data to the new line in the besteld tab and clear the old line.
      currentrow.clear();


      addDate(besteld, rowBes, 5); // Add the date of this action.
      teBestellen.getRange(i, 5).setValue(false); // Reset the checkbox.
    }
  }
}
// Checks to see if items are checked of in the tab besteld. If checked, places removes the item from the besteld tab and places the item in the tab statistieken. It also adds the date in the new tab en resets the checkmark in the old tab.
function itemArrived() {
  for (let i = 2; i < 30; i++) { // Loop over the first 30 items in the given tab, in the given column.
    let orderd = besteld.getRange(i, 6).getValue(); // Get the value of the checkbox of the item in the list.
    if (orderd == true) { // Check if the item's checkbox was checked, if checked proceed with following code, if not go to next line.
      const currentrow = besteld.getRange(i, 1, 1, 5); // Get all data of the the current line.
      let destRange = statistieken.getRange(statistieken.getLastRow()+1, 1, 1, 5); // Select the range where an empty cell was found.
      let rowStat= statistieken.getLastRow()+1; // Get the value of this row.
      currentrow.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false); // Copy all data to the new line in the besteld tab and clear the old line.
      currentrow.clear();


      addDate(statistieken, rowStat, 6); // Add the date of this action.
      besteld.getRange(i, 6).setValue(false); // Reset the checkbox.
    }
  }
}
// Make a function to calculate the total number of times an item was used in the selected year. The calculation is based on the data from voorraadbeheer tab
// and the tab opgebruikte reagentia. Builds a graph with the 6 most used items.
function makeYearRapport() {
  const givenYearDate= configuration.getRange(12, 2).getValue();
  if (givenYearDate < 2100 && givenYearDate > 2000) { // Check if the date is in the yyyy format and is a possible date.
    const yearRapportColumn= jaarrapporten.getLastColumn()+1; // Gets the last empty column.
    jaarrapporten.getRange(1, yearRapportColumn).setValue(givenYearDate);
    let rowMin= 2; // Begin row of the tab minimum voorraad.
    let itemIdMin = minimumVoorraad.getRange(rowMin, 1).getValue(); // The name of the item on specified row on the minimum voorraad tab.
    // As long as the current cell is not empty the function goes over the table in minimum voorraad.
    while ( itemIdMin != '') { // Loop over all the items in the minimum voorraad tab.
      //
      // Loops over all items in the tab opgebruikteReagentia and counts the items with the correct name which are used up.
      let rowOpgbrRea = 2; // Begin row of the tab opgebruikteReagentia.
      let selectedTextopgbrRea = opgebruikteReagentia.getRange(rowOpgbrRea, 1).getValue(); // The name of the item on specified row on the opgebruikteReagentia tab.
      let totalitems = 0; // Initialize parameter; set the total counted items to 0.
      while (selectedTextopgbrRea!= '') { // Loop over all items in the tab opgebruikteReagentia untill there is an empty row.
        if (itemIdMin == selectedTextopgbrRea) { // If the names of the items in minimum voorraad and in the vooraadbeheer tab match.
          let expirationYear = opgebruikteReagentia.getRange(rowOpgbrRea, 11).getValue().getFullYear(); // Gets the year in yyyy format.
          if (expirationYear == givenYearDate ) { // Checks to see if the year is the same as the given year.
            // eslint-disable-next-line prefer-const
            let useditem = opgebruikteReagentia.getRange(rowOpgbrRea, 6).getValue(); // Gets the used items and adds them to the total.
            // Adding the number of used items in this instance to the number of total items.
            totalitems = totalitems + useditem;
          }
        }
        rowOpgbrRea = rowOpgbrRea +1; // Go to the next not empty row in the opgebruikteReagentia tab.

        selectedTextopgbrRea = opgebruikteReagentia.getRange(rowOpgbrRea, 1).getValue(); // Get the next name in the table.
      }
      // Loops over all items in the tab voorraadbeheer and counts the items with the correct name which have been started that year
      let rowVoorraad = 2; // Begin row of the tab voorraadbeheer.
      let selectedTextVoor = voorraadbeheer.getRange(rowVoorraad, 1).getValue(); // The name of the item on specified row on the voorraadbeheer tab.
      while (selectedTextVoor!= '') { // Loop over all items in the tab voorraadbeheer untill there is an empty row.
        if (itemIdMin == selectedTextVoor) { // If the names of the items in minimum voorraad and in the vooraadbeheer tab match.
          let initiationYear = voorraadbeheer.getRange(rowVoorraad, 10).getValue().getFullYear(); ; // Gets the year in YYYY format.
          if (initiationYear == givenYearDate ) { // Checks to see if the year is the same as the given year.
            // eslint-disable-next-line prefer-const
            let useditem = voorraadbeheer.getRange(rowVoorraad, 6).getValue(); // Gets the used items and adds them to the total.
            // Adding the number of used items in this instance to the number of total items.
            totalitems = totalitems + useditem;
          }
        }
        rowVoorraad = rowVoorraad +1; // Go to the next not empty row in the opgebruikteReagentia tab.

        selectedTextVoor = voorraadbeheer.getRange(rowVoorraad, 1).getValue(); // Get the next name in the table.
      }

      let rowYear= 2; // Begin row of the tab jaar rapport tab.
      let itemIdYear = jaarrapporten.getRange(rowYear, 1).getValue(); // The name of the item on specified row on the jaar rapport tab.

      let alreadyInList = false; // Sets the bolean to check if an item is already present in the list.
      while (itemIdYear != '') { // Loop over the list in jaar rapport tab.
        if (itemIdMin== itemIdYear) { // Check if the item is already in the list.
          alreadyInList = true; // Initialize parameter; if the item is already in the list, set this to true.
          jaarrapporten.getRange(rowYear, yearRapportColumn ).setValue(totalitems);
        }
        rowYear = rowYear +1; // Go to the next not empty row

        itemIdYear = jaarrapporten.getRange(rowYear, 1).getValue(); // Get the name of the item of the row.
      }
      // If the item is not on the jaar list, add it at the bottom.
      if (alreadyInList == false) {
        const destRange =jaarrapporten.getLastRow()+1;
        jaarrapporten.getRange(destRange, 1).setValue(itemIdMin);

        jaarrapporten.getRange(destRange, yearRapportColumn ).setValue(totalitems);
      }


      //
      rowMin = rowMin +1; // Go to the next not empty row
      itemIdMin = minimumVoorraad.getRange(rowMin, 1).getValue(); // Get the name of the item of the row.
    }
    let maxSortRange = jaarrapporten.getRange(jaarrapporten.getLastRow(), jaarrapporten.getLastColumn()).getA1Notation(); // Get A1 range natation for the last row and colum in the table.
    let sortRange = jaarrapporten.getRange('A2:'+maxSortRange); // Select the sorting range.
    // Sorts descending by the new column.
    sortRange.sort({column: yearRapportColumn, ascending: false});
    // Create a range to build a chart.
    let firstCell = jaarrapporten.getRange(2, yearRapportColumn); // Select the first row.
    let lastCell= jaarrapporten.getRange(8, yearRapportColumn); // Select the last row.
    firstCell= firstCell.getA1Notation(); // Change first row range to A1 notation.
    lastCell= lastCell.getA1Notation(); // Change last row range to A1 notation.

    const chart = jaarrapporten.newChart() // Make a new chart.
        .setChartType(chart.ChartType.BAR) // Make a barplot chart.
        .setOption('title', 'Top 5 most frequently used items of '+givenYearDate) // Give a title.
        .setOption('titleTextStyle.alignment', 'center') // Center the title.
        .setOption('hAxis.title', 'Number of items') // Give a name to the horizontal axis.
        .setOption('vAxis.title', 'Itemnames') // Give a name to the vertical axis.
        .addRange(jaarrapporten.getRange('A2:A8')) // Select the item name range.
        .addRange(jaarrapporten.getRange(firstCell+':'+lastCell)) // Select the range of the values.
        .setOption('colors', ['#1FD0E9']) // Give a color scheme.
        .setPosition(4, 4, 0, 0) // Set the position on the sheet.
        .build();


    jaarrapporten.insertChart(chart); // Adds the new chart.
  } else {
    SpreadsheetApp.getUi().alert('The year needs to be between 2000 and 2100.'); // Give the user an errormessage to give a correct date.
  }
}

// Calculates the averige time it takes between the order date and the date of the arrival. This is calculated based on the firm and uses the data of the tab statestieken.
// Builds a graph with all the firms and the average order length.
function averageOrderTime() {
  const miliSecondsDay = 1000 * 60 * 60 * 24; // Number of miliseconds in 1 day.
  let firmRow = 15; // First row of the table with firms.
  let firm = configuration.getRange(firmRow, 2).getValue(); // First firm of the table.
  while (firm != '') { // While there are firms in the table.
    let totalDaysTillArive =0; // Initialize parameter; total number of days for 1 firm.
    let numberOfOrders =0; // Initialize parameter; total number of orders for 1 firm.
    let averageOrderTime= 0; // Initialize parameter; average number of days for 1 firm.
    let statistiekenRow = 2; // Initialize parameter; first row of the tab statestieken.
    let statFirm = statistieken.getRange(statistiekenRow, 3).getValue(); //
    while (statFirm != '') {
      if (statFirm ==firm) { // If the firm in statistieken has the same name as the firm in the list in configuration.
        let startDate =statistieken.getRange(statistiekenRow, 5).getValue(); // Get the start date of the order.
        let endDate = statistieken.getRange(statistiekenRow, 6).getValue(); // Get the date of arrival of the order.
        let first =startDate.getTime(); // Get time in miliseconds sinds Epoc.
        let last =endDate.getTime(); // Get time in miliseconds sinds Epoc.
        let mili = last - first; // Calculate the defenrence in miliseconds.
        let daysTillArive= mili/ miliSecondsDay; // Divide the difference in miliseconds by the number of miliseconds in a day.
        totalDaysTillArive= totalDaysTillArive +daysTillArive; // Add to the total number of days of orders.
        numberOfOrders= numberOfOrders +1; // Total number of orders.
      }

      statistiekenRow = statistiekenRow +1; // Select the next value in the tab of statistieken.
      statFirm = statistieken.getRange(statistiekenRow, 3).getValue();
    }
    averageOrderTime =totalDaysTillArive/numberOfOrders; // Calculate the averige lenght of orders the selected firm.
    configuration.getRange(firmRow, 3).setValue(averageOrderTime); // Put the averige lenght of orders in the table.
    configuration.getRange(firmRow, 4).setValue(numberOfOrders); // Put the total number of orders in the table.


    firmRow = firmRow +1; // Select the next value in the tab of configuration.
    firm = configuration.getRange( firmRow, 2).getValue();
  }


  let allCharts =configuration.getCharts(); // Select all charts in the tab configuration and removes them.
  // eslint-disable-next-line guard-for-in
  for (let i in allCharts) {
    let chart = allCharts[i];
    configuration.removeChart(chart);
  }
  firmRow=firmRow-1; // End of the firm table.
  const chart = configuration.newChart() // Make a new chart.
      .setChartType(chart.ChartType.BAR) // Make a barplot chart.
      .setOption('title', 'Gemiddelde besteltijd bij firma\'s' ) // Give a title.
      .setOption('titleTextStyle.alignment', 'center') // Center the title.
      .setOption('hAxis.title', 'Firma\'s') // Give a name to the horizontal axis.
      .setOption('vAxis.title', 'Gemiddeld besteltijd') // Give a name to the vertical axis.
      .addRange(configuration.getRange('B15:C'+firmRow)) // Select the range.
      .setOption('colors', ['#1FD0E9']) // Give a color scheme.
      .setOption('height', 386)
      .setOption('width', 625)
      .setPosition(13, 1, 0, 3) // Set the position on the sheet.
      .build();

  configuration.insertChart(chart); // Adds the new chart.
}


