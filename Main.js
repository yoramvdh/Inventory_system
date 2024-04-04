/* eslint-disable camelcase */
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
    row = row +1; // go to the next not empty row
    currentcell = voorraadbeheer.getRange(row, 1).getValue();
  }

  maxrange.sort(7); // sort the whole range on column 7
}

// Calculates the total supply of a selected item.
function totalSupply() {
  // start of itemarray
  const itemarr = [['Item name', 'Expiration date ', 'Number of items']];
  // Get the value of the cell selected in the spreadsheet.
  const selObj = SpreadsheetApp.getActiveSheet().getSelection();
  const selectedText = selObj.getActiveRange().getValue();
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
        const expirationdate_Obj = voorraadbeheer.getRange(row, 8).getValue();
        const expirationdate_String= expirationdate_Obj.toString();
        const expirationdate = expirationdate_String.slice(0, 16);
        // Adds items to array.
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
function moveFile(name_Of_Destination, id_Of_File) {
  let folderpresent = false; // initiates parameter to check if folder is available
  let folderid = ''; // initiates the ID string for the folder

  const folders = DriveApp.getFolders(); // gets all folder form the users drive

  // loops over all folders and compares the name to the name given given as the first param
  while (folders.hasNext()) {
    const folder = folders.next();
    const foldername = folder.getName();

    if (foldername === name_Of_Destination) {
    // if the folder is present,change folderpresent too true and get the id
      folderpresent = true;
      folderid = folder.getId();
    }
  }
  // if the folder does not exist create the folder and get the ID
  if (folderpresent === false) {
    folderid = DriveApp.createFolder(name_Of_Destination).getId();
  }
  // get the folder and move file to folder.
  const correctfolder = DriveApp.getFolderById(folderid);
  DriveApp.getFileById(id_Of_File).moveTo(correctfolder);
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


