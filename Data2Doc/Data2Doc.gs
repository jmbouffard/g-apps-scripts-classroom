/**************************************************************************************
 * Data2Doc.gs
 *
 * This Google Sheets script creates a document from the values 
 *  in columns "titre" and "nouvelles". The lines are selected based on the date
 *  that was selected by the user. Everytime the script is ran for a new date,
 *  a new document is created. A template is used to create the new document and
 *  the script fails if the template is not found.
 *
 * The script was developed using a specific Sheets document and would have to be
 *  adapted to use with a different format.
 *
 **************************************************************************************/

// Running the script with Today's date
//  This was done to bypass the date selection dialog which was using the
//  deprecated UiApp object 
function startAppToday() {
  var d = new Date();
  SpreadsheetApp.getUi().alert('Création du document pour le '+d);
  // Processing document
  readRows(d);
}

// Running the script with Tomorrow's date
//  This was done to bypass the date selection dialog which was using the
//  deprecated UiApp object 
function startAppTomorrow() {
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var now = new Date();
  var d = new Date(now.getTime() + MILLIS_PER_DAY);
  SpreadsheetApp.getUi().alert('Création du document pour le '+d);
  // Processing document
  readRows(d);
}

function startApp() {
  var app = UiApp.createApplication();
  var introtxt = "Sélectionnez la date de la liste à produire et\nappuyez sur Continuer pour lancer la création du document. \"Data2Doc result\" utilise le contenu des colonnes \"titre\" et \"nouvelles\"";
  //var textBoxA = app.createTextBox().setId('textBoxA').setName('Title');
  var currentdate = new Date();
  var label1 = app.createLabel(introtxt);
  var datebox = app.createDateBox().setId('datebox').setName('datebox').setValue(currentdate);
  //var datetxt = app.createTextBox().setId('datetxt').setName('datetxt').setText(currentdate).setVisible(false);
  var button = app.createButton("Continuer");
  app.add(label1);
  app.add(datebox);
  //app.add(datetxt);
  app.add(button);
  var handler_date = app.createServerHandler('dateChangeHandler');
  var handler_button = app.createServerHandler('continueHandler')
                          .addCallbackElement(datebox);
  datebox.addValueChangeHandler(handler_date);
  button.addClickHandler(handler_button);
  // Show app dialog
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);
};

function dateChangeHandler(eventInfo) {
  var app = UiApp.getActiveApplication();
  // Do stuff when date is changed
  //var datetxt = app.getElementById('datetxt');
  //atetxt.setText(eventInfo.parameter.datebox);
  return app;
};

function continueHandler(e) {
  var app = UiApp.getActiveApplication();
  var d = new Date(e.parameter.datebox);
  // Processing document
  readRows(d);
  app.close();
  return app;
};

function getFile(child_folder, filename) {
  //var child_folder = DriveApp.getFolder('Annonces_RHJ');
  //var filename = 'Template_RHJ';
  var parents = child_folder.getParents();
  var folder = parents[0];
  var files = folder.getFiles();
  //Logger.log("files found: "+files.length);
  for (var i=0; i<files.length; i++) {
    //Logger.log("file #"+i+": "+files[i].getName());
    if (files[i].getName() == filename) {
      //Logger.log("file found: "+files[i].getName());
      return files[i];
    }
  }
  return null;
}

/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows(docdate) {
  var app = UiApp.getActiveApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  // Current date
  if (!docdate) {
    docdate = new Date();
  }
  
  //var range = ss.getRangeByName("UsersRange");
  var sheetObjects = getRowsData(sheet, rows, 1);
  
  // Create the new document
  var docname = docdate.getYear()+'/'+(docdate.getMonth()+1)+'/'+docdate.getDate()+' - RHJ';
  var folder_iterator = DriveApp.getFoldersByName('Annonces_RHJ');
  var rhjfolder;
  if (folder_iterator.hasNext()) {
    rhjfolder = folder_iterator.next();
  } else {
    throw new Error("Dossier Annonces_RHJ introuvable");
  }
  Logger.log("doc name: "+docname);
  //var doc = DocumentApp.create(docname);
  //var rhjfile = DriveApp.getFileById(doc.getId());
  var templatefile_iterator = DriveApp.getFilesByName('Template_RHJ');
  var templatefile;
  if (templatefile_iterator.hasNext()) {
    templatefile = templatefile_iterator.next();
  } else {
    throw new Error("Fichier Template_RHJ introuvable");
  }
  var rhjfile = templatefile.makeCopy(docname, rhjfolder);
  //rhjfile.removeFromFolder(DriveApp.getRootFolder());
  //rhjfile.addToFolder(rhjfolder);
  var doc = DocumentApp.openById(rhjfile.getId());

  // Process document to find edit location
  var body = doc.getBody();
  var searchResult = body.findText('%Articles%');
  var elem = searchResult.getElement();
  var para = elem.getParent();
  var container = para.getParent();
  var index = container.getChildIndex(para);
  container.removeChild(para);
  var pos = 0;

  Logger.log("Document date: " + docdate);
  Logger.log("Reading spreadsheet with # rows = " + numRows);
  Logger.log("Reading spreadsheet with # sheetObjects = " + sheetObjects.length);
    
  for (var i = 1; i <= sheetObjects.length - 1; i++) {
    var row = values[i];
    // Add a paragraph to the document
    var line = sheetObjects[i];
    
    Logger.log("Item date: " + line.parution);
    
    if (docdate.getYear() == line.parution.getYear() &&
        docdate.getMonth() == line.parution.getMonth() &&
        docdate.getDate() == line.parution.getDate() ) {
      Logger.log("Add item to line: " + (index+(2*pos)));
      var title_para = container.insertParagraph(index+(2*pos), line.titre);
      title_para.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      var content_para = container.insertParagraph(index+(2*pos)+1, line.nouvelles);
      content_para.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      pos = pos + 1;
    }
  }
  
  // Save and close the document
  doc.saveAndClose();
  
};

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Export today's rows to document",
    functionName : "startAppToday"
  },{
    name : "Export tomorrow's rows to document",
    functionName : "startAppTomorrow"
  }];
  sheet.addMenu("Data2Doc", entries);
};

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getLastColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}


// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}
