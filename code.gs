/** -----------GLOBAL VARIABLES, CAN BE CALLED FROM ANY FILE ------------ **/
/** --------------------------------------------------------------------  **/
var ss = SpreadsheetApp.getActiveSpreadsheet();
var quotesheet = ss.getSheetByName("Quote"); // The name of the sheet where the quote is built
var productlistsheet = ss.getSheetByName("Product Finder"); // The name of the sheet where the list of products in chosen
var quoterange = quotesheet.getRange('B10:U');
var quotefullrange = quotesheet.getRange('A10:V');
var quoteNameRange = quotesheet.getRange('B7'); // The cell containting the company name
var companyName = quoteNameRange.getValue();
var quoteDateRange = quotesheet.getRange('U3'); // The cell containing the date
var quoteDate = quoteDateRange.getValue();

/** ---------------------------------------------------------------------- **/

function Trial() {
    var one = DriveApp.getFolderById("folder-id-1").getName();
    var two = DriveApp.getFolderById("folder-id-2").getName();
    var thr = DriveApp.getFolderById("folder-id-3").getName();
    var fiv = DriveApp.getFolderById('folder-id-4').getName();
}

/** 
 *================================================================================= =================================================================================
 *                                                                     Copy to, Clear & Load Quotes 
 *================================================================================= =================================================================================
 **/
function sortQuote() {
    quotesheet.getRange('B10:U').sort(20);
    quotesheet.getRange('B10:U').sort(2);
}

function CopyToQuote() {
    //copies selected line items from product finder in to the quote sheet

    //Set variables  
    var data = productlistsheet.getRange(3, 1, productlistsheet.getLastRow(), 20).getValues(); //num columns to copy is 20 = T
    var dest = [];
    var ui = SpreadsheetApp.getUi();

    //Check array for a ◉ in column [19] = T if true then write the contents to an array called 'dest'
    for (var i = 0; i < data.length; i++) {
        if (data[i][19] == "◉") {
            dest.push(data[i]);
        }
    }

    // if something has been written in the dest array then batch write it to quotesheet
    if (dest.length > 0) {
        //Get the range to place the data and then place the data there.| first row | first column | num rows | num columns
        quotesheet.getRange(quotesheet.getLastRow() + 1, 2, dest.length, dest[0].length).setValues(dest);
    }

    //Sort the quote into alphabetical order
    sortQuote();
    //Uncheck all selected products
    productlistsheet.getRange("T3:T").setValue("");

    //delete rows that were added
    try {
        quotesheet.deleteRows(quotesheet.getLastRow() - dest.length, dest.length);
    } catch (e) {
        Logger.log(e + 'if rows are out of bounds, it is because nothing was selected in Product list to copy');
    }

    // give pop-up explaining what has happened
    ui.alert("Items copied", "(" + dest.length + ") Items added to the Quote sheet.", ui.ButtonSet.OK);
    //                        list.setFormula("QUERY(Import!A2:V, \"Select A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,V where A contains \'\"& B1 &\"\' and A contains \'\"& P1 &\"\' and A contains \'\"& Q1 &\"\' \") ");
}

function ClearSelectedQuote() {
    //copies selected line items from product finder in to the quote sheet

    var ui = SpreadsheetApp.getUi();
    var dest = [];
    var dest2 = [];
    try {
        var quoteList = quotesheet.getRange('B10:B').getValues(); // the range to find
        var listLength = quoteList.filter(String).length; // number of non empty cells in the range (the number of rows of data that there are)
        var quoteRange = quotesheet.getRange(10, 2, listLength, 23); //the actual range of the quote (first row, first column, num rows, num columns)
        var quoteItems = quoteRange.getValues(); // Get the values of the whole quote.
    }

    //throw an error message if the quoterange is invalid
    catch (e) {
        ui.alert("No items in quote", "There are no items in the quote", ui.ButtonSet.OK);
        return
    }

    //Check that the user really wants to delete selected items
    var checkresponse = ui.alert("Are you sure?", "Do you wish to delete the contents of all the selected rows, this cannot be undone", ui.ButtonSet.YES_NO);

    //if the user does not want to continue then end 
    if (checkresponse != ui.Button.YES) {
        return;
    }

    //Check rows for a ◉ in column [22] = X if true then write the start and end rows into arrays dest and dest2
    for (var i = 0; i < quoteItems.length; i++) {
        if (quoteItems[i][22] == "◉") {
            var first = quoteRange.getCell(i + 1, 1).getA1Notation(); // 1  = B
            var last = quoteRange.getCell(i + 1, 20).getA1Notation(); // 20 = U
            dest.push(first);
            dest2.push(last);
        }
    }

    // loop through the selected rows and delete their values
    for (var u = 0; u < dest.length; u++) {
        quotesheet.getRange(dest[u] + ":" + dest2[u]).setValue("");
    }

    //De-select rows (get range based on 2 columns to the right of the quote)
    quotesheet.getRange(10, quotefullrange.getLastColumn() + 2, quotesheet.getLastRow() - 9, 1).setValue(""); //first row, first column, num rows, num columns

    // Sort the quote to remove gaps
    sortQuote();
}

function CopyToQuoteCustom() {
    //copies selected line items from product finder in to the quote sheet

    //Set variables 
    var customProductSheet = ss.getSheetByName("Custom Product Builder");
    var data = customProductSheet.getRange(2, 1, customProductSheet.getLastRow(), 20).getValues(); // (start row, start column, num rows, num columns) Last Column to copy is 20 = T
    var dest = [];
    var ui = SpreadsheetApp.getUi();

    //Check rows for a ◉ in column [19] = T if true then write the contents to an array called 'dest'
    for (var i = 0; i < data.length; i++) {
        if (data[i][19] == "◉") {
            dest.push(data[i]);
        }
    }

    // if something has been written in the dest array then batch write it to quotesheet
    if (dest.length > 0) {
        //Get the range to place the data and then place the data there.| first row | first column | num rows | num columns
        quotesheet.getRange(quotesheet.getLastRow() + 1, 2, dest.length, dest[0].length).setValues(dest);
    }

    //Sort the quote into alphabetical order
    sortQuote();
    //Uncheck all selected products
    customProductSheet.getRange("T2:T").setValue("");

    //delete rows that were added
    try {
        quotesheet.deleteRows(quotesheet.getLastRow() - dest.length, dest.length);
    } catch (e) {
        Logger.log(e + 'if rows are out of bounds, it is because nothing was selected in Product list to copy');
    }

    // give pop-up explaining what has happened
    ui.alert("Items copied", "(" + dest.length + ") Items added to the Quote sheet.", ui.ButtonSet.OK);
    //                        list.setFormula("QUERY(Import!A2:V, \"Select A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,V where A contains \'\"& B1 &\"\' and A contains \'\"& P1 &\"\' and A contains \'\"& Q1 &\"\' \") ");
}

function clearQuote() {
    // Clears the quote sheet of all data
    var ui = SpreadsheetApp.getUi();
    var check = ui.alert("ARE YOU SURE?", "Do you want to delete the contents of this quote? This cannot be undone.", ui.ButtonSet.YES_NO);
    if (check == ui.Button.YES) {
        quotefullrange.setValue(""); // Clear all line items
        quoteNameRange.setValue("COMPANY NAME"); // Reset the quote name
        quoteDateRange.setFormula('text(TODAY(),"dd/mm/yy")'); // Reset date to today
    }
}

function loadQuote() {
    //Loads existing quotes into the main quote sheet
    var lastColumn = quotesheet.getLastColumn();
    var ui = SpreadsheetApp.getUi();

    // if no valid quote selected then show message and exit script
    if (quotesheet.getRange(2, lastColumn - 4).getValue() == "#N/A") {
        ui.alert("No quote to load", "Select a quote to load from the dropdown menu.", ui.ButtonSet.OK);
        return;
    }

    var quoteURL = quotesheet.getRange(2, lastColumn - 4).getValue(); //=VLOOKUP(W2,Y3:Z,2,FALSE)
    var loadData = SpreadsheetApp.openByUrl(quoteURL).getRange("B10:U200").getValues();
    var loadCompanyName = SpreadsheetApp.openByUrl(quoteURL).getRange("B7").getValues();
    var loadDate = SpreadsheetApp.openByUrl(quoteURL).getRange("U3").getValue();
    try {
        var formattedDate = Utilities.formatDate(loadDate, "GMT", "dd/MM/YY");

        quoteDateRange.setFormula('text("' + formattedDate + '","dd/mm/yy")');
    } catch (e) {
        quoteDateRange.setValue(loadDate);
    }
    quoteNameRange.setValues(loadCompanyName);
    quoterange.setValues(loadData);

}

/** 
 *================================================================================= =================================================================================
 *                                                                        Print File Arrays
 *================================================================================= =================================================================================
 **/

function printAllLists() {
    printFileArray();
    printFileArray2();
    printFileArray3();
}

function printFileArray() { //Print the Quote list
    var folder = DriveApp.getFolderById("folder-id-1");
    var files = folder.getFiles();
    var fileList = [];
    //Loop though files and add names and urls to the array
    while (files.hasNext()) {
        var file = files.next();
        var fileName = file.getName();
        var fileUrl = file.getUrl();
        fileList = fileList.concat([
            [fileName, fileUrl]
        ]);
    }
    var lastColumn = quotesheet.getLastColumn();
    var firstCell = quotesheet.getRange(3, lastColumn - 5);
    var lastCell = firstCell.offset(fileList.length - 1, fileList[0].length - 1);
    var destinationRange = quotesheet.getRange(firstCell.getA1Notation() + ':' + lastCell.getA1Notation());
    quotesheet.getRange(3, lastColumn - 5, quotesheet.getLastRow() - 3, 2).setValue(""); //(row, column, numrows, numcolumns)
    destinationRange.setValues(fileList);
    Logger.log("Range" + destinationRange.getA1Notation());
}

function printFileArray2() { //Print the PDF list

    var folder = DriveApp.getFolderById("folder-id-3");
    var files = folder.getFiles();
    var fileList = [];
    //Loop though files and add names and urls to the array
    while (files.hasNext()) {
        var file = files.next();
        var fileName = file.getName();
        var fileUrl = file.getUrl();
        fileList = fileList.concat([
            [fileName, fileUrl]
        ]);
    }
    var lastColumn = quotesheet.getLastColumn();
    var firstCell = quotesheet.getRange(3, lastColumn - 3);
    var lastCell = firstCell.offset(fileList.length - 1, fileList[0].length - 1);
    var destinationRange = quotesheet.getRange(firstCell.getA1Notation() + ':' + lastCell.getA1Notation());

    // clear the existing list then print the new one
    quotesheet.getRange(3, lastColumn - 3, quotesheet.getLastRow() - 3, 2).setValue(""); //(row, column, numrows, numcolumns)                      
    destinationRange.setValues(fileList);
}

function printFileArray3() { //Print the templates list
    var ui = SpreadsheetApp.getUi();
    var folder = DriveApp.getFolderById("folder-id-2");
    var files = folder.getFolders();
    var fileList = [];
    //Loop though files and add names and urls to the array
    try {
        while (files.hasNext()) {
            var file = files.next();
            var fileName = file.getName();
            var fileUrl = file.getUrl();
            fileList = fileList.concat([
                [fileName, fileUrl]
            ]);
        }
        var lastColumn = quotesheet.getLastColumn();
        var firstCell = quotesheet.getRange(3, lastColumn - 1);
        var lastCell = firstCell.offset(fileList.length - 1, fileList[0].length - 1);
        var destinationRange = quotesheet.getRange(firstCell.getA1Notation() + ':' + lastCell.getA1Notation());
        //delete the existing list then print the new one
        quotesheet.getRange(3, lastColumn - 1, quotesheet.getLastRow() - 3, 2).setValue(""); //(row, column, numrows, numcolumns)
        destinationRange.setValues(fileList);
    } catch (error) {
        ui.alert("Cannot update templates folder list", "The templates folder list cannot be created \n" + error.message, ui.ButtonSet.OK)
    }
}

/** 
 *================================================================================= =================================================================================
 *                                                                      Generate PDF + Templates
 *================================================================================= =================================================================================
 **/

function generatePdf() {
    // --------========== CREATE THE PDF ==========--------------
    var ui = SpreadsheetApp.getUi();
    // ---=== CREATE AND MODIFY DUPLICATE QUOTE SHEET ===---
    //Sheets folder
    var existing = DriveApp.getFolderById("folder-id-1").getFilesByName(companyName + " Quote " + quoteDate).hasNext();


    //Check if a quote by the same name already exists, if it doesn't then contiue to save.
    if (existing == false) {

        //update the date IMPORTANT TO DO THIS BEFORE CREATING THE NEW FILES
        quoteDateRange.setFormula('text(TODAY(),"dd/mm/yy")');

        //Save out new spreadhseet & PDF
        createFiles();
    }
    //if a quote by the same name already exists then do the following
    else {
        //Show warning box about existing quote
        var pdfresponse = ui.alert("QUOTE ALREADY EXISTS", "Quote with this name already exists, Would you like to override the existing quote with this one?", ui.ButtonSet.YES_NO);
    }
    //if user chooses to override old files then do the following
    if (pdfresponse == ui.Button.YES) {

        //delete existing spreasheet
        DriveApp.getFolderById("folder-id-1").getFilesByName(companyName + " Quote " + quoteDate).next().setTrashed(true);

        //delete existing pdf
        try {
            DriveApp.getFolderById("folder-id-3").getFilesByName(companyName + " Quote " + quoteDate).next().setTrashed(true);
        } catch (e) {
            ui.alert("PDF file did not exist", " no PDF of the quote was found to override. Creating a new pdf...", ui.ButtonSet.OK);
        }

        //update the date IMPORTANT TO DO THIS BEFORE CREATING THE NEW FILES
        quoteDateRange.setFormula('text(TODAY(),"dd/mm/yy")');

        //Save out new spreadhseet & PDF
        createFiles();

    }

    // Update the list of available quotes to load
    printFileArray();
    printFileArray2();


}

function createFiles() {
    var ui = SpreadsheetApp.getUi();
    // ---=== CREATE AND MODIFY DUPLICATE QUOTE SHEET ===---
    // sort the data in the quote before saving
    sortQuote();

    // Create and name a new Spreadsheet
    var newSpreadsheet = SpreadsheetApp.create(companyName + " Quote " + quoteDate);

    //Copy the quote sheet to a new spreadsheet
    quotesheet.copyTo(newSpreadsheet);

    // Delete Columns 23 = V and all following
    var newquote = newSpreadsheet.getSheetByName('Copy of Quote');
    newquote.deleteColumns(23, newquote.getMaxColumns() - 22);

    // Delete Unused Rows
    var newColumn = newquote.getRange('B10:B').getValues();
    var newLength = newColumn.filter(String).length;
    var newmaxrows = newquote.getMaxRows();

    newquote.deleteRows(newLength + 10, newmaxrows - newLength - 10);

    //Replace the date formula with pure value so it doesn't change
    var setDate = newSpreadsheet.getSheetByName('Copy of Quote').getRange('T3');
    setDate.setValue(setDate.getValue());

    //copy the payment terms row to the bottom
    //newquote.getRange(8, 1, 1, newquote.getMaxColumns()).moveTo(newquote.getRange(newquote.getMaxRows(), 1, 1));


    //Select the default sheet that is made with the sheet
    newSpreadsheet.getSheetByName('Sheet1').activate();

    // Delete the selected useless sheet
    newSpreadsheet.deleteActiveSheet();

    // ---=== CREATE PDF FROM THE DUPLICATED QUOTE SHEET ===---

    // Find the file to make into a pdf
    var forPdf = DriveApp.getFileById(newSpreadsheet.getId());

    // Setup the file as a pdf and set its name.
    var theBlob = forPdf.getBlob().getAs('application/pdf').setName(companyName + " Quote " + quoteDate);

    // Folder id to save PDF in a folder.
    var pdfFolder = DriveApp.getFolderById("folder-id-3");

    //Folder id to save Sheets in a folder.
    var sheetsFolder = DriveApp.getFolderById("folder-id-1");

    //Create the PDF file in the pdf folder (important to do this BEFORE creating the CSV sheet)
    var pdffile = pdfFolder.createFile(theBlob);

    // get the url of the pdf
    var pdfurl = pdffile.getUrl();

    //Set sharing permissions for the pdf file
    pdffile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);


    //Create a copy of the duplicate quote sheet in the sheets folder and get its URL
    var newURL = DriveApp.getFileById(newSpreadsheet.getId()).makeCopy(companyName + " Quote " + quoteDate, sheetsFolder).getUrl();

    //Delete the original duplicate quote sheet
    DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);


    // Show the user the sharing link for the pdf
    ui.alert("PDF Created", "Sharable link for quote PDF:\n" + pdfurl, ui.ButtonSet.OK)
}

function exportTemplates() {
    var ui = SpreadsheetApp.getUi();
    //Does the file already exist?
    var existing = DriveApp.getFolderById("folder-id-2").getFoldersByName(companyName + ' templates ' + quoteDate).hasNext();
    //Check to see if the file already exists, if not create the file
    if (existing == false) {
        createZip();
    }

    //if the file already exists give the user the option to override it
    else {
        var templateresponse = ui.alert("Templates already exist", "A templates folder already exists for this quote, do you want to override it with a new templates folder", ui.ButtonSet.YES_NO);
    }

    //if user chooses to override then delete the existing template and create a new one
    if (templateresponse == ui.Button.YES) {

        //Delete the existing file
        DriveApp.getFolderById("folder-id-2").getFoldersByName(companyName + ' templates ' + quoteDate).next().setTrashed(true);

        //Create the new file
        createZip();
    }

    //Update the Lists of zip files
    printFileArray3();
}

function createZip() {
    var ui = SpreadsheetApp.getUi();
    //Pre-sort the data in the quote to make sure the correct number of values is returned
    sortQuote();

    //Get the correct number of used rows
    try {
        var quoteList = quotesheet.getRange('B10:B').getValues(); // the range to find
        var listLength = quoteList.filter(String).length; // number of non empty cells in the range (the number of rows of data that there are)
        var quoteRange = quotesheet.getRange(10, 2, listLength, 20); //the actual range of the quote (first row, first column, num rows, num columns)
        var quoteItems = quoteRange.getValues(); // Get the values of the whole quote.
    }

    //throw an error message if the quoterange is invalid
    catch (e) {
        ui.alert("No items in quote", "There are no templates in the quote, cannot create zip file \n" + e.message, ui.ButtonSet.OK);
        return
    }

    //get the folder where the zips will be saved
    var templatesFolder = DriveApp.getFolderById('folder-id-2');
    //make an empty array to place template files into
    var templates = [];
    //The folder where all templates are saved
    var sourcetemplates = DriveApp.getFolderById('folder-id-4');
    //make a 0 value for the number of items that will be skipped
    var skippednumber = 0;
    //make an empty array to place the names of all the lineitems that have no template
    var skippeditems = [];
    //Make a 0 value for the number of items that will be added
    var enterednumber = 0;
    //Make an empty array to place the names of the templates that were entered
    var entereditems = [];
    // the names of line items that should have no template
    var skipitems = ["ΩArtwork()", "ΩDe-rig()", "ΩInstall()", ""];

    // loop through all the rows of the quote
    for (var i = 0; quoteRange.getLastRow() - 9 > i; i++) {

        // for each row get the value in column that holds the template name
        var templatename = quoteItems[i][17];

        //if the lineitem is not supposed to have a template just skip it altogether
        if (skipitems.indexOf(String(templatename)) > -1 || templatename == "") {
            continue
        }

        // if the template is not in the templates array already then do the following
        if (String(templates).indexOf(String(templatename)) == -1) {

            //attempt to add the template to the array
            try {

                // get the template by its name and add the template to the array
                templates.push(sourcetemplates.getFilesByName(templatename).next().getName());
                //templates.push(sourcetemplates.getFilesByName(templatename).next());

                //add 1 to the entered number
                enterednumber++

                //add the name of the line item that will be zipped
                entereditems.push(templatename);
            }

            //if the above was unsuccessful (i.e. the template was not in the sourcefolder) then do the following
            catch (e) {
                skippeditems.push(templatename);
                skippednumber++
            }

        }

    } //end loop for quote rows
    Logger.log(entereditems);

    //try and create zip file in folder usings the list of templates   
    try {
        var templatecompanyfolder = templatesFolder.createFolder(companyName + ' templates ' + quoteDate);

        var templatecompanyfolderid = DriveApp.getFoldersByName(companyName + ' templates ' + quoteDate).next().getId();

        for (var u = 0; templates.length > u; u++) {
            sourcetemplates.getFilesByName(templates[u]).next().makeCopy(templatecompanyfolder);
        }


        var templateurl = templatecompanyfolder.getUrl();
        templatecompanyfolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        //append \n on to each item in the skippeditems array so that each one appears on a new line in the final message
        for (var y = 0; skippeditems.length > y; y++) {
            skippeditems[y] = skippeditems[y] + "\n"
        }

        // set skipped number & skipped items up to display message correctly
        if (skippednumber == 0) {
            skippednumber = "All templates in quote were found and zipped";
            skippeditems = ""
        } else {
            skippednumber = "[" + skippednumber + "] Templates could not be found\n\n";
            skippeditems = "Templates not found: \n [" + skippeditems + "]"
        }


        //Show a message saying what was exported.
        ui.alert("Templates Created", "[" + enterednumber + "] Templates added to zip \n" + skippednumber + skippeditems + "\n\n Sharing link for template.zip\n" + templateurl, ui.ButtonSet.OK);

    }

    //if error is thrown show a message box explaining error then end the function.
    catch (e) {
        ui.alert("Template creation failed", "Error desctiption:\n" + e.message + "\n File name:" + e.fileName + " Line No: " + e.lineNumber + "\n If error reads 'invalid argument' it is usually because there are no products in the quote list that have templates.", ui.ButtonSet.OK);
        return;
    }

}
