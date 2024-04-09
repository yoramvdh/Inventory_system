
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
const minimumVoorraad = sheets[1];
const teBestellen = sheets[2];
const besteld = sheets[3];
const opgebruikteReagentia = sheets[4];
const vervallenReagentia = sheets[5];
const statistieken = sheets[6];
const configuration = sheets[7];


// Create array to store all the links
const links = [];
// For each sheet in sheets add an array element to our array with the
// string of the URL for that sheet
sheets.forEach((sheet)=>links.push(sheetUrl+'#gid='+sheet.getSheetId()));
const voorraadbeheerlink = links[0];
const vervallenReagentialink = links[5];

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
      .addItem('Total Supply of one item', 'totalSupply')
      .addSeparator()
      .addItem('Calculate supplies', 'minimumSupply')
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
      // Get list of emails in the config tab.
      const emailList= getRowOfData(configuration, 8);
      // Send a mail if a product is expired.
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

// This function uses a trigger to find all product who are almoust expired and send a mail to specifiek users
function almoustExpiredProducts() {
  const almoustexpired = 14; // number of day befor the product expires
  let currentcell = voorraadbeheer.getRange(row, 1).getValue();

  // As long as the currentcell is not empty the function goes over the table and will compare each time the expirationdate with the number of days till it expires.

  while ( currentcell != '') {
    const expiredate = voorraadbeheer.getRange(row, 9).getValue();
    const alreadyused =voorraadbeheer.getRange(row, 11).getValue();
    if (expiredate == almoustexpired && alreadyused === '' ) { // Checks if the product is almost expired.
      // Get list of emails in the config tab.
      const emailList= getRowOfData(configuration, 9);
      // Sends a mail to the users.
      MailApp.sendEmail({to: emailList,
        subject: 'automatische mail- Bijna Vervallen product',
        htmlBody: 'Het product '+currentcell+' op rij '+row+' zal over 14 dagen vervallen: '+ voorraadbeheerlink,
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
    row = row +1; // go to the next not empty row
    currentcell = voorraadbeheer.getRange(row, 1).getValue();
  }

  maxrange.sort(7); // sort the whole range on column 7
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
  if (selectedText != '') { // Gives an error message if not the selected cell is empty.
  // Initiates values.
    let totalitems = 0;
    let currentitem = 0;
    // While the current cell is not empty compare values with the selected text.
    while ( currentcell != '') {
      // eslint-disable-next-line prefer-const
      let itemId =voorraadbeheer.getRange(row, 1).getValue();
      if ( itemId == selectedText ) { // if the the current cell has the same value the selected text.
        // eslint-disable-next-line prefer-const
        // Calculate the current number of items by taking the maxium number of items and subscracting the used items.
        const maxitem = voorraadbeheer.getRange(row, 4).getValue();
        // eslint-disable-next-line prefer-const
        let useditem = voorraadbeheer.getRange(row, 6).getValue();
        currentitem = maxitem- useditem;
        // added the number of used items in this instance to the number of total items.
        totalitems = totalitems + currentitem;
        const arr = []; // Generates a new array and adds the name of the item
        arr.push(selectedText);
        // Takes the expirationdate out from the selected line. Transforms it to string then slices the string to keep relevant information.
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
      row = row +1; // go to the next line.
      currentcell = voorraadbeheer.getRange(row, 1).getValue();
    }
    // This function take a name of an item, a number of item in total and an array to construct a table.
    createDoc(selectedText, totalitems, itemarr);
  } else { // Gives an error message if not the selected cell is empty.
    SpreadsheetApp.getUi().alert('Select the item you want to check the inventory for out of the list of in the minimum voorraad tab. Then run the function again.');
  }
}
// This function take a name of an item, a number of item in total and an array to construct a table.
function createDoc(itemname, totalitems, itemarr) {
  // style of the title
  const titleStyle = {};
  titleStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  titleStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  titleStyle[DocumentApp.Attribute.FONT_SIZE] = 20;
  titleStyle[DocumentApp.Attribute.BOLD] = true;
  // style of the basic text
  const textStyle = {};
  textStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  textStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  textStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
  textStyle[DocumentApp.Attribute.BOLD] = false;
  // style of the solution text
  const solutionStyle = {};
  solutionStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  solutionStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  solutionStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
  solutionStyle[DocumentApp.Attribute.BOLD] = true;
  solutionStyle[DocumentApp.Attribute.UNDERLINE] = true;
  // style of the table
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
  // Searches a file in the drive, check to see if the given foldername exist in the drive, if not the scipt creates it, then moves the file in the folder.
  moveFile('Item rapports automatically generated', docid);
}


// Searches a file in the drive, check to see if the given foldername exist in the drive, if not the script creates it, then moves the file in the folder.
function moveFile(nameOfDestination, idOfFile) {
  let folderpresent = false; // initiates parameter to check if folder is available
  let folderid = ''; // initiates the ID string for the folder

  const folders = DriveApp.getFolders(); // gets all folder form the users drive

  // loops over all folders and compares the name to the name given given as the first param
  while (folders.hasNext()) {
    const folder = folders.next();
    const foldername = folder.getName();

    if (foldername === nameOfDestination) {
    // if the folder is present,change folderpresent too true and get the id
      folderpresent = true;
      folderid = folder.getId();
    }
  }
  // if the folder does not exist create the folder and get the ID
  if (folderpresent === false) {
    folderid = DriveApp.createFolder(nameOfDestination).getId();
  }
  // get the folder and move file to folder.
  const correctfolder = DriveApp.getFolderById(folderid);
  DriveApp.getFileById(idOfFile).moveTo(correctfolder);
  const fileUrl=DriveApp.getFileById(idOfFile).getUrl();

  SpreadsheetApp.getUi().alert('The file is available the '+nameOfDestination+' folder. ', 'This is the link to the file: \n '+fileUrl, SpreadsheetApp.getUi().ButtonSet.OK);
}

// This function gets all the data from 1 row from a selected sheet en concatinates its in a comma sepperated string.
function getRowOfData(sheet, rowNumber) {
  let columnNumber = 2; // starting with column 2
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
  emailList=emailList.slice(1); // Remove the comma at the start of the string.

  return emailList; // return the string.
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
// This function loops over a list in the tab minimum voorraad, for each item in the list it loops over the list in the tab voorraadbeheer and calculated the total supplies present for the item in minimum voorraad.
// The calculation is done by suptracting the used items from the total items. Next the item is compared to a list of minimum item kwantities, if the item is lower then the minimum than the te bestellen page and the besteld page are checked to see if the item is present in the first 30 lines. If the item is present then update the current supplies, if not present search the first empty line and add the itemname, the active supply and the current date.
function minimumSupply() {
  let rowMin= 2; // Begin row of the tab minimum voorraad.

  let itemId = minimumVoorraad.getRange(rowMin, 1).getValue(); // The name of the item on specified row on the minimum voorraad tab.

  // As long as the currentcell is not empty the function goes over the table in minimum voorraad.
  while ( itemId != '') {
    let row = 2; // Begin row of the tab voorraad.
    // Loop over all items in the tab voorraadbeheer.
    let selectedText = voorraadbeheer.getRange(row, 1).getValue(); // The name of the item on specified row on the voorraadbeheer tab.

    let totalitems = 0; // Set the total counted items to 0.
    while (selectedText!= '') { // Loop over all items in the tab voorraadbeheer untill there is an empty row.
      if (itemId == selectedText) { // If the names of the items in minimum voorraad and in the vooraadbeheer tab match.
        // Calculate the current number of items by taking the maxium number of items and subscracting the used items.
        const maxitem = voorraadbeheer.getRange(row, 4).getValue();
        // eslint-disable-next-line prefer-const
        let useditem = voorraadbeheer.getRange(row, 6).getValue();
        currentitem = maxitem- useditem;
        // added the number of used items in this instance to the number of total items.
        totalitems = totalitems + currentitem;
      }
      row = row +1; // go to the next not empty row in the voorraadbeheer tab.

      selectedText = voorraadbeheer.getRange(row, 1).getValue(); // Get the next name in the table.
    }
    // Sets the calculated number of supplys in the tabel next to the correct item.
    minimumVoorraad.getRange(rowMin, 2).setValue(totalitems);

    // eslint-disable-next-line prefer-const
    let minimumSupply = minimumVoorraad.getRange(rowMin, 3).getValue(); // Get the minimum supply for a product from the tab minimum voorraad.

    if (totalitems < minimumSupply ) { // If the current value is lower then the minimul supply.
      const resultTeBestellen = itemPresentInList(teBestellen, itemId, 2); // Look if the item is already present in the tab te bestellen and notate the row if so.
      const inListTeBestellen=resultTeBestellen.alreadyInList; // If values was in list then this is true.
      const itemRowTeBestellen=resultTeBestellen.rowWithItem;

      const resultBesteld = itemPresentInList(besteld, itemId, 2); // Look if the item is already present in the tab besteld and notate the row if so.
      const inListBesteld=resultBesteld.alreadyInList; // If values was in list then this is true.
      const itemRowBesteld=resultBesteld.rowWithItem;

      if (inListTeBestellen== true) { // If the item was in the list te bestellen then update the number of items to the current supply.
        teBestellen.getRange(itemRowTeBestellen, 3).setValue(totalitems);
      }
      if (inListBesteld== true) { // If the item was in the list besteld then update the number of items to the current supply.
        besteld.getRange(itemRowBesteld, 3).setValue(totalitems);
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
        teBestellen.getRange(rowTeBes, 3).setValue(totalitems); // Add the current number of items.
      }
    }


    rowMin = rowMin +1; // Go to the next not empty row

    itemId = minimumVoorraad.getRange(rowMin, 1).getValue(); // Get the name of the item of the row.
  }
}

// This function takes the name of the tab and the name of and item and the column number. It will then check if it can find the name in the specified column.
// returns an array: alreadyInList is true when the item is present, false if not. rowWithItem: is 0 when not found, otherwise it will be the row number where the name was found.
function itemPresentInList(sheet, itemId, columnNumber) {
  let alreadyInList = false; // set the default as false
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
