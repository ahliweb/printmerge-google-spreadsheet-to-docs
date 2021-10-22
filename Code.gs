// modified to spreadsheet with docs by unggul@ahliweb.com
// origin from https://spreadsheet.dev/mail-merge-from-google-sheets-to-google-slides

function mailMergeDocsFromSheets() {
  // Load data from the spreadsheet
  var dataRange = SpreadsheetApp.getActive().getDataRange();
  var sheetContents = dataRange.getValues();

  // Save the header in a variable called header
  var header = sheetContents.shift();

  // Create an array to save the data to be written back to the sheet.
  // We'll use this array to save links to Google Docs.
  var updatedContents = [];

  // Add the header to the array that will be written back
  // to the sheet.
  updatedContents.push(header);

  // For each row, see if the 22th column is empty.
  // If it is empty, it means that a slide deck hasn't been
  // created yet.
  sheetContents.forEach(function(row) {
    if(row[21] === "") {
      // Create a Google Slides presentation using
      // information from the row.
      var docs = createDocsFromRow(row);
      var docsId = docs.getId();
   
      // Create the Google Docs' URL using its Id.
      var docsUrl = `https://docs.google.com/document/d/${docsId}/edit`;

      // Add this URL to the 4th column of the row and add this row
      // to the updatedContents array to be written back to the sheet.
      row[21] = docsUrl;
      updatedContents.push(row);
    }
  });

  // Write the updated data back to the Google Sheets spreadsheet.
  dataRange.setValues(updatedContents);

}

function createDocsFromRow(row) {
 // Create a copy of the Slides template
 var deck = createCopyOfDocsTemplate();

 // Rename the deck using the nama and lastname of the student
 deck.setName(row[1] + " - " + row[0] + " - " + row[2]);

 // Replace template variables using the student's information.
 deck.replaceText("{{nama}}", row[0]);
 deck.replaceText("{{seri}}", row[1]);
 deck.replaceText("{{wa}}", row[2]);
 deck.replaceText("{{nilai}}", row[3]);
//
 deck.replaceText("{{data4}}", row[4]);
 deck.replaceText("{{data5}}", row[5]);
 deck.replaceText("{{data6}}", row[6]);
 deck.replaceText("{{data7}}", row[7]);
 deck.replaceText("{{data8}}", row[8]);
 deck.replaceText("{{data9}}", row[9]);
 deck.replaceText("{{data10}}", row[10]);
 deck.replaceText("{{data11}}", row[11]);
 deck.replaceText("{{data12}}", row[12]);
 deck.replaceText("{{data13}}", row[13]);
 deck.replaceText("{{data14}}", row[14]);
 deck.replaceText("{{data15}}", row[15]);
 deck.replaceText("{{data16}}", row[16]);
 deck.replaceText("{{data17}}", row[17]);
 deck.replaceText("{{data18}}", row[18]);
 deck.replaceText("{{data19}}", row[19]);
 deck.replaceText("{{data20}}", row[20]);
// you can add these variable

 deck.replaceText("{{link}}", row[21]);


 return deck;
}

function createCopyOfDocsTemplate() {
 //
 var TEMPLATE_ID = "1cyFn9Cr3bahZeYME46rt5u-PwxvUCWvBv-1U6V8d_Zw";

 // Create a copy of the file using DriveApp
 var copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy();

 // Load the copy using the DocumentApp.
 var docs = DocumentApp.openById(copy.getId());

 return docs;
}

function onOpen() {
 // Create a custom menu to make it easy to run the Mail Merge
 // script from the sheet.
 SpreadsheetApp.getUi().createMenu("⚙️ Admin")
   .addItem("Create Docs", "mailMergeDocsFromSheets")
   .addToUi();
}
