/**
 * Creates an HtmlOutput object from an HTML file.
 * 
 * @return {HtmlService.HtmlOutput} an HtmlOutput object
 */
function doGet() {
    return HtmlService.createTemplateFromFile('Page').evaluate();
}

/**
 * Imports the specified file into the current file.
 * 
 * @param {string} filename the name of the file to import
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const sheet = SpreadsheetApp.openById(id).getSheetByName(name);

/**
 * Inserts data into the sheet.
 * 
 * @param {string} text a JSON string that represents a one-dimensional array with three elements
 * @return {string} an HTML string
 */
function insertData(text) {
    /**
     * @type {string[]}
     */
    const rowContents = JSON.parse(text);
    if (rowContents[0] == String()) return '<p>A may not be empty.</p>';
    const range = sheet.getRange('A:A').createTextFinder(rowContents[0]).findNext();
    if (range != null) return '<p>A must be unique.</p>';
    sheet.appendRow(rowContents);
    return `<p>The data were inserted into the sheet:</p>${toHtmlString(rowContents)}`;
}

/**
 * Selects data from the sheet.
 * 
 * @param {string} a the value of A
 * @return {string} an HTML string
 */
function selectData(a) {
    if (a == String()) return '<p>A may not be empty.</p>';
    const values = sheet.getDataRange().getValues().find(values => values[0] == a);
    if (values == undefined) return '<p>The data were not found in the sheet.</p>';
    return `<p>The data were selected from the sheet:</p>${toHtmlString(values)}`;
}

/**
 * Updates data in the sheet.
 * 
 * @param {string} text a JSON string that represents a one-dimensional array with three elements
 * @return {string} an HTML string
 */
function updateData(text) {
    /**
     * @type {string[]}
     */
    const values = JSON.parse(text);
    if (values[0] == String()) return '<p>A may not be empty.</p>';
    const range = sheet.getRange('A:A').createTextFinder(values[0]).matchEntireCell(true).findNext();
    if (range == null) return '<p>The data were not found in the sheet.</p>';
    const row = range.getRowIndex();
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    return `<p>The data were updated:</p>${toHtmlString(values)}`;
}

/**
 * Deletes data from the sheet.
 * 
 * @param {string} a the value of A
 * @return {string} an HTML string
 */
function deleteData(a) {
    if (a == String()) return '<p>A may not be empty.</p>';
    const range = sheet.getRange('A:A').createTextFinder(a).matchEntireCell(true).findNext();
    if (range == null) return '<p>The data were not found in the sheet.</p>';
    const rowPosition = range.getRowIndex();
    const values = sheet.getDataRange().getValues().find(values => values[0] == a);
    const htmlString = toHtmlString(values);
    sheet.deleteRow(rowPosition);
    return `<p>The data were deleted from the sheet:</p>${htmlString}`;
}

/**
 * Converts the specified data into an HTML string.
 * 
 * @param {string[]} rowContents a one-dimensional array with three elements
 * @return {string} an HTML string that contains a table element with the given values
 */
function toHtmlString(rowContents) {
    return `<table>
        <thead>
          <tr>
            <th>A</th>
            <th>B</th>
            <th>C</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            ${rowContents.map(rowContent => `<td>${rowContent}</td>`).reduce((td1, td2) => td1 + td2)}
          </tr>
        </tbody>
      </table>`;
}
