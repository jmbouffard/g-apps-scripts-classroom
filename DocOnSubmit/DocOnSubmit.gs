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
  Logger.log("onOpen called");
  var ui = FormApp.getUi();
  ui.createMenu('DocOnSubmit')
      .addItem('Info', 'onAideClicked')
      .addItem('Install Trigger', 'createFormSubmitTrigger')
      .addToUi();
      /*.addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Second item', 'menuItem2'))
      .addToUi();*/
};

/**
 * Verify if trigger was install and install it otherwise.
 */
function createFormSubmitTrigger() {
  var form = FormApp.getActiveForm();
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length == 0) {
    FormApp.getUi().alert('Trigger installed on OnFormSubmit().');
    ScriptApp.newTrigger('onFormSubmit')
        .forForm(form)
        .onFormSubmit()
        .create();
  } else {
    FormApp.getUi().alert('Trigger alreay installed, no change were made.');
  }
}

/**
 * Display help dialog.
 */
function onAideClicked() {
  FormApp.getUi().alert('If documents are not created in the "Files" subfolder when\nsubmitting results, use the "Install Trigger" button.');
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
  
  var fileNames;
  if ((typeof responses[0].getResponse()) == "string")
  {
    // This path used if the filename is a dropbox selection (only one answer)
    fileNames = [responses[0].getResponse()];
  } else {
    // This path used if the filename is a set of checkbox selections (multiple answers possible)
    fileNames = responses[0].getResponse();
  }
  Logger.log("Number of names: "+fileNames.length);

  for (var n = 0; n < fileNames.length; ++n) {
    var fileName = fileNames[n];
    Logger.log("Writing file: "+fileName);
  
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
            
      // TEST: Check for response type (https://developers.google.com/apps-script/reference/forms/item-type)
      //Logger.log("Response Type: "+responses[i].getItem().getType());
      // If item named "vidéo / photo :" then handle the value as a link to a Drive location.
      if (responses[i].getItem().getTitle() == "vidéo / photo :") {
        var imageNames;
        if ((typeof responses[i].getResponse()) == "string")
        {
          // This path used if one image
          imageNames = [responses[i].getResponse()];
        } else {
          // This path used if multiple images
          imageNames = responses[i].getResponse();
        }
        Logger.log("Number of images submitted: "+imageNames.length);
        var td = tr.appendTableCell();
        for (var j = 0; j < imageNames.length; ++j) {
          //var par = td.appendParagraph("https://drive.google.com/open?id="+imageNames[j]);
          var par = td.insertParagraph(j,"https://drive.google.com/open?id="+imageNames[j]);
          // Define a style for links.
          var linkStyle = {};
          linkStyle[DocumentApp.Attribute.LINK_URL] = "https://drive.google.com/open?id="+imageNames[j];
          linkStyle[DocumentApp.Attribute.MARGIN_BOTTOM] = 6;
          par.setAttributes(linkStyle);
        }
      } else {
        var td = tr.appendTableCell(responses[i].getResponse());
        td.setAttributes(normalStyle);
      }
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

