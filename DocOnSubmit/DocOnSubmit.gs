/**************************************************************************************
 * DocOnSubmit.gs
 *
 * This Google Forms script appends a table with all answers to the form to a
 *  Docs document. The document name must be included as the answer to the first
 *  question of the form. If the document does not exist it will be created before
 *  being populated. The script runs everytime the form is submitted.
 *
 * The script was developed using a specific Forms document and would have to be
 *  adapted to use with a different format.
 *
 **************************************************************************************/

/**
 * Adds a custom menu to the active form, containing a help item.
 */
 function onOpen() {
  //Logger.log("onOpen called");
  var ui = FormApp.getUi();
  ui.createMenu('DocOnSubmit')
      .addItem('Info', 'onAideClicked')
      .addToUi();    
      /*.addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Second item', 'menuItem2'))
      .addToUi();*/
};

/**
 * Display help dialog.
 */
function onAideClicked() {
  FormApp.getUi().alert('Documents will be created in the Form\'s folder when user submits results!');
};

/**
 * Runs a folder and file creation test, not used in normal operation.
 */
function TestWriteFiles() {
  //See if content folder exists, if not create
  contentFolder = setContentFolder("Files");  
  //See if content file exists, if not create
  contentFile = setContentFile(contentFolder, "Test");
};

// Will run everytime the Form is submitted.
// Arguments:
//   - e: Contains all results from submitted form.
function onFormSubmit(e) {
  //Logger.log("authMode: "+e.authMode);
  //Logger.log("response: "+e.response);
  //Logger.log("source: "+e.source);
  var responses = e.response.getItemResponses();
  Logger.log("Number of responses: "+responses.length);
  
  var fileName = responses[0].getResponse();

  //See if content folder exists, if not create
  contentFolder = setContentFolder("Files");
  
  //See if content file exists, if not create
  contentFile = setContentFile(contentFolder, fileName);
  
  var document = DocumentApp.openById(contentFile.getId());
  var body = document.getBody();
  var table = body.appendTable();
  
  // Define a style with bold.
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  // Define a style with normal.
  var normalStyle = {};
  normalStyle[DocumentApp.Attribute.BOLD] = false;
  
  for (var i = 1; i < responses.length; ++i) {
    Logger.log("Q"+i+": "+responses[i].getItem().getTitle());
    Logger.log("A"+i+": "+responses[i].getResponse());
    var tr = table.appendTableRow();
    var td = tr.appendTableCell(responses[i].getItem().getTitle());
    td.setAttributes(boldStyle);
    //var tr = table.appendTableRow();
	// If item named "vid�o / photo :" then handle the value as a link to a Drive location.
    if (responses[i].getItem().getTitle() == "vid�o / photo :") {
      var td = tr.appendTableCell("https://drive.google.com/open?id="+responses[i].getResponse());
      // Define a style for links.
      var linkStyle = {};
      linkStyle[DocumentApp.Attribute.LINK_URL] = "https://drive.google.com/open?id="+responses[i].getResponse();
      td.setAttributes(linkStyle);
    } else {
      var td = tr.appendTableCell(responses[i].getResponse());
      td.setAttributes(normalStyle);
    }
  }
  
};

// Creates new file in provided folder.
// Arguments:
//   - folder: Folder to store the new file
//   - name: Name of new file
// Returns:
//   - Newly created File
function setContentFile(folder, name) {
  //See if content file exists
  var contentFolder = folder
  var contentFileName = name;
  var newFile;
  try {
    Logger.log("Looking for file: "+contentFileName);
    newFile = contentFolder.getFilesByName(contentFileName).next();
  }
  catch(e) {
    Logger.log("File: "+contentFileName+" does not exist, creating...");
    var newDocument = DocumentApp.create(contentFileName);
    // Set Title
    var body = newDocument.getBody();
    var titleParagraph = body.insertParagraph(0, "Preuves d'apprentissages - "+contentFileName);
    titleParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    titleParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    // Move file
    newFile = DriveApp.getFileById(newDocument.getId());
    contentFolder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);
  }
  
  Logger.log("Content file is: "+newFile.getUrl());
  return newFile;
};

// Creates new folder.
// Arguments:
//   - name: Name of new folder
// Returns:
//   - Newly created Folder
function setContentFolder(name) {
  //Locate current folder
  var thisFileId = FormApp.getActiveForm().getId();
  var thisFile = DriveApp.getFileById(thisFileId);
  var parentFolder = thisFile.getParents().next();
  
  //See if content folder exists
  var folder_name = name;
  var newFdr;
  try {
    Logger.log("Looking for folder: "+folder_name);
    newFdr = parentFolder.getFoldersByName(folder_name).next();
  }
  catch(e) {
    Logger.log("Folder: "+folder_name+" does not exist, creating...");
    newFdr = parentFolder.createFolder(folder_name);
  }
  
  Logger.log("Content folder is: "+newFdr.getUrl());
  return newFdr;
};
