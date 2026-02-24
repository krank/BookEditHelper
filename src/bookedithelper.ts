type validReturnTypes = string | Date | number;

function Setup() {

  const prefix: string = "BookEditHelper.";
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("BookEditHelper")
    .addItem("Get word count from document URLs", `${prefix}GetWordCountFromDocumentUrls`)
    .addItem("Get page count from document URLs", `${prefix}GetPageCountFromDocumentUrls`)
    .addItem("Get document names from document URLs", `${prefix}GetDocumentNameFromDocumentUrls`)
    .addItem("Get last modified dates from document URLs", `${prefix}GetLastEditDateFromDocumentUrls`)
    .addToUi();
}

/* -----------------------------------------------------------------------------
   Menu wrappers
----------------------------------------------------------------------------- */

function GetWordCountFromDocumentUrls() {
  getFromDocumentUrls((doc) => getWordCount(doc).toString());
}

function GetPageCountFromDocumentUrls() {
  getFromDocumentUrls((doc => {
    const n = doc.getBlob().getDataAsString().split("/Contents").length - 1;
    return n.toString();
  }));
}

function GetDocumentNameFromDocumentUrls() {
  getFromDocumentUrls((doc) => doc.getName());
}

function GetLastEditDateFromDocumentUrls() {
  getFromDocumentUrls((doc) => {
    let t = DriveApp.getFileById(doc.getId()).getLastUpdated().toISOString();
    return new Date(t);
  });
}

/* -----------------------------------------------------------------------------
   Functionality
----------------------------------------------------------------------------- */

// https://stackoverflow.com/a/78375674
/**
 * Reads the word count of the file
 * @param {DocumentApp.Document} file
 * @return {number}
 */
function getWordCount(file: GoogleAppsScript.Document.Document) {
  // https://regex101.com/r/TpAuUt/1
  const nonWordRegex = /[\s]+(?:[\-—–]+\s+)*/g;
  const emptyFileRegex = /^[\s-—–]+$/i;
  const text = file.getBody().getText();
  if (emptyFileRegex.test(text)) {
    return 0;
  }
  const matches = file.getBody().getText().match(nonWordRegex);
  return !!matches ? matches.length : 0;
}

/**
 * Goes through current selection's first column, expecting urls to Google Docs. 
 * Runs a getter function on each url, inserts the result in last column of selection, on the same row
 * @param getterFunction the function to be run on every document
 * @returns 
 */
function getFromDocumentUrls(getterFunction: (doc: GoogleAppsScript.Document.Document) => validReturnTypes): void {
  const selectedRange = SpreadsheetApp.getSelection().getActiveRange();
  const ui = SpreadsheetApp.getUi();

  if (!selectedRange || selectedRange.getNumColumns() < 2)
  {
    ui.alert("No valid range selected");
    return;
  }

  const rightmostColumn = selectedRange.offset(0, selectedRange.getNumColumns() - 1);

  const rightmostColumnValues: validReturnTypes[][] = rightmostColumn.getValues();
  const selectedRangeValues: string[][] = selectedRange.getValues();

  for (let rowNum = 0; rowNum < selectedRangeValues.length; rowNum++) {
    const row = selectedRangeValues[rowNum];

    const url = row[0];
    try {
      const doc = DocumentApp.openByUrl(url);

      rightmostColumnValues[rowNum][0] = getterFunction(doc);
    } catch (e) {
      Logger.log(`{url} wasn't a valid docs url`)
    }
  }

  rightmostColumn.setValues(rightmostColumnValues);
}