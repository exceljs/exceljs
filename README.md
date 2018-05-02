# ExcelJS

[![Build Status](https://travis-ci.org/guyonroche/exceljs.svg?branch=master)](https://travis-ci.org/guyonroche/exceljs)

Read, manipulate and write spreadsheet data and styles to XLSX and JSON.

Reverse engineered from Excel spreadsheet files as a project.

# Installation

```shell
npm install exceljs
```

# New Features!

<ul>
  <li>
    Merged <a href="https://github.com/guyonroche/exceljs/pull/556">Add test case with a huge file #556</a>.
    Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> and
    <a href="https://github.com/holm">Christian Holm</a> for this contribution.
  </li>
</ul>

# Contributions

Contributions are very welcome! It helps me know what features are desired or what bugs are causing the most pain.

I have just one request; If you submit a pull request for a bugfix, please add a unit-test or integration-test (in the spec folder) that catches the problem.
 Even a PR that just has a failing test is fine - I can analyse what the test is doing and fix the code from that.

To be clear, all contributions added to this library will be included in the library's MIT licence.

# Notice

I have to admit I'm still really snowed under at the moment and will not be able to
pursue active development of exceljs in the coming months.
I will however keep an eye on the pull requests and endeavour to merge and release
them as they come.

# Backlog

<ul>
  <li>ESLint - slowly turn on (justifyable) rules which should, I hope, help make contributions easier.</li>
  <li>Conditional Formatting.</li>
  <li>There are still more print-settings to add; Fixed rows/cols, etc.</li>
  <li>XLSX Streaming Reader.</li>
  <li>Parsing CSV with Headers</li>
</ul>

# Contents

<ul>
  <li>
    <a href="#interface">Interface</a>
    <ul>
      <li><a href="#create-a-workbook">Create a Workbook</a></li>
      <li><a href="#set-workbook-properties">Set Workbook Properties</a></li>
      <li><a href="#workbook-views">Workbook Views</a></li>
      <li><a href="#add-a-worksheet">Add a Worksheet</a></li>
      <li><a href="#remove-a-worksheet">Remove a Worksheet</a></li>
      <li><a href="#access-worksheets">Access Worksheets</a></li>
      <li><a href="#worksheet-properties">Worksheet Properties</a></li>
      <li><a href="#page-setup">Page Setup</a></li>
      <li>
        <a href="#worksheet-views">Worksheet Views</a>
        <ul>
          <li><a href="#frozen-views">Frozen Views</a></li>
          <li><a href="#split-views">Split Views</a></li>
        </ul>
      </li>
      <li><a href="#auto-filters">Auto Filters</a></li>
      <li><a href="#columns">Columns</a></li>
      <li><a href="#rows">Rows</a></li>
      <li><a href="#handling-individual-cells">Handling Individual Cells</a></li>
      <li><a href="#merged-cells">Merged Cells</a></li>
      <li><a href="#defined-names">Defined Names</a></li>
      <li><a href="#data-validations">Data Validations</a></li>
      <li><a href="#styles">Styles</a>
        <ul>
          <li><a href="#number-formats">Number Formats</a></li>
          <li><a href="#fonts">Fonts</a></li>
          <li><a href="#alignment">Alignment</a></li>
          <li><a href="#borders">Borders</a></li>
          <li><a href="#fills">Fills</a></li>
          <li><a href="#rich-text">Rich Text</a></li>
        </ul>
      </li>
      <li><a href="#outline-levels">Outline Levels</a></li>
      <li><a href="#images">Images</a></li>
      <li><a href="#file-io">File I/O</a>
        <ul>
          <li><a href="#xlsx">XLSX</a>
            <ul>
              <li><a href="#reading-xlsx">Reading XLSX</a></li>
              <li><a href="#writing-xlsx">Writing XLSX</a></li>
            </ul>
          </li>
          <li><a href="#csv">CSV</a>
            <ul>
              <li><a href="#reading-csv">Reading CSV</a></li>
              <li><a href="#writing-csv">Writing CSV</a></li>
            </ul>
          </li>
          <li><a href="#streaming-io">Streaming I/O</a>
            <ul>
              <li><a href="#reading-csv">Streaming XLSX</a></li>
            </ul>
          </li>
        </ul>
      </li>
    </ul>
  </li>
  <li><a href="#browser">Browser</a></li>
  <li>
    <a href="#value-types">Value Types</a>
    <ul>
      <li><a href="#null-value">Null Value</a></li>
      <li><a href="#merge-cell">Merge Cell</a></li>
      <li><a href="#number-value">Number Value</a></li>
      <li><a href="#string-value">String Value</a></li>
      <li><a href="#date-value">Date Value</a></li>
      <li><a href="#hyperlink-value">Hyperlink Value</a></li>
      <li>
        <a href="#formula-value">Formula Value</a>
        <ul>
          <li><a href="#shared-formula">Shared Formula</a></li>
          <li><a href="#formula-type">Formula Type</a></li>
        </ul>
      </li>
      <li><a href="#rich-text-value">Rich Text Value</a></li>
      <li><a href="#boolean-value">Boolean Value</a></li>
      <li><a href="#error-value">Error Value</a></li>
    </ul>
  </li>
  <li><a href="#config">Config</a></li>
  <li><a href="#known-issues">Known Issues</a></li>
  <li><a href="#release-history">Release History</a></li>
</ul>

# Interface

```javascript
var Excel = require('exceljs');
```

## Create a Workbook

```javascript
var workbook = new Excel.Workbook();
```

## Set Workbook Properties

```javascript
workbook.creator = 'Me';
workbook.lastModifiedBy = 'Her';
workbook.created = new Date(1985, 8, 30);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2016, 9, 27);
```

```javascript
// Set workbook dates to 1904 date system
workbook.properties.date1904 = true;
```

## Workbook Views

The Workbook views controls how many separate windows Excel will open when viewing the workbook.

```javascript
workbook.views = [
  {
    x: 0, y: 0, width: 10000, height: 20000,
    firstSheet: 0, activeTab: 1, visibility: 'visible'
  }
]
```

## Add a Worksheet

```javascript
var sheet = workbook.addWorksheet('My Sheet');
```

Use the second parameter of the addWorksheet function to specify options for the worksheet.

For Example:

```javascript
// create a sheet with red tab colour
var sheet = workbook.addWorksheet('My Sheet', {properties:{tabColor:{argb:'FFC0000'}}});

// create a sheet where the grid lines are hidden
var sheet = workbook.addWorksheet('My Sheet', {properties: {showGridLines: false}});

// create a sheet with the first row and column frozen
var sheet = workbook.addWorksheet('My Sheet', {views:[{xSplit: 1, ySplit:1}]});
```

## Remove a Worksheet

Use the worksheet `id` to remove the sheet from workbook.

For Example:

```javascript
// Create a worksheet
var sheet = workbook.addWorksheet('My Sheet');

// Remove the worksheet using worksheet id
workbook.removeWorksheet(sheet.id)
```

## Access Worksheets
```javascript
// Iterate over all sheets
// Note: workbook.worksheets.forEach will still work but this is better
workbook.eachSheet(function(worksheet, sheetId) {
    // ...
});

// fetch sheet by name
var worksheet = workbook.getWorksheet('My Sheet');

// fetch sheet by id
var worksheet = workbook.getWorksheet(1);
```

## Worksheet Properties

Worksheets support a property bucket to allow control over some features of the worksheet.

```javascript
// create new sheet with properties
var worksheet = workbook.addWorksheet('sheet', {properties:{tabColor:{argb:'FF00FF00'}}});

// create a new sheet writer with properties
var worksheetWriter = workbookWriter.addSheet('sheet', {properties:{outlineLevelCol:1}});

// adjust properties afterwards (not supported by worksheet-writer)
worksheet.properties.outlineLevelCol = 2;
worksheet.properties.defaultRowHeight = 15;
```

**Supported Properties**

| Name             | Default    | Description |
| ---------------- | ---------- | ----------- |
| tabColor         | undefined  | Color of the tabs |
| outlineLevelCol  | 0          | The worksheet column outline level |
| outlineLevelRow  | 0          | The worksheet row outline level |
| defaultRowHeight | 15         | Default row height |
| dyDescent        | 55         | TBD |

### Worksheet Metrics

Some new metrics have been added to Worksheet...

| Name              | Description |
| ----------------- | ----------- |
| rowCount          | The total row size of the document. Equal to the row number of the last row that has values. |
| actualRowCount    | A count of the number of rows that have values. If a mid-document row is empty, it will not be included in the count. |
| columnCount       | The total column size of the document. Equal to the maximum cell count from all of the rows |
| actualColumnCount | A count of the number of columns that have values. |


## Page Setup

All properties that can affect the printing of a sheet are held in a pageSetup object on the sheet.

```javascript
// create new sheet with pageSetup settings for A4 - landscape
var worksheet =  workbook.addWorksheet('sheet', {
  pageSetup:{paperSize: 9, orientation:'landscape'}
});

// create a new sheet writer with pageSetup settings for fit-to-page
var worksheetWriter = workbookWriter.addSheet('sheet', {
  pageSetup:{fitToPage: true, fitToHeight: 5, fitToWidth: 7}
});

// adjust pageSetup settings afterwards
worksheet.pageSetup.margins = {
  left: 0.7, right: 0.7,
  top: 0.75, bottom: 0.75,
  header: 0.3, footer: 0.3
};

// Set Print Area for a sheet
worksheet.pageSetup.printArea = 'A1:G20';

// Repeat specific rows on every printed page
worksheet.pageSetup.printTitlesRow = '1:3';

```

**Supported pageSetup settings**

| Name                  | Default       | Description |
| --------------------- | ------------- | ----------- |
| margins               |               | Whitespace on the borders of the page. Units are inches. |
| orientation           | 'portrait'    | Orientation of the page - i.e. taller (portrait) or wider (landscape) |
| horizontalDpi         | 4294967295    | Horizontal Dots per Inch. Default value is -1 |
| verticalDpi           | 4294967295    | Vertical Dots per Inch. Default value is -1 |
| fitToPage             |               | Whether to use fitToWidth and fitToHeight or scale settings. Default is based on presence of these settings in the pageSetup object - if both are present, scale wins (i.e. default will be false) |
| pageOrder             | 'downThenOver'| Which order to print the pages - one of ['downThenOver', 'overThenDown'] |
| blackAndWhite         | false         | Print without colour |
| draft                 | false         | Print with less quality (and ink) |
| cellComments          | 'None'        | Where to place comments - one of ['atEnd', 'asDisplayed', 'None'] |
| errors                | 'displayed'   | Where to show errors - one of ['dash', 'blank', 'NA', 'displayed'] |
| scale                 | 100           | Percentage value to increase or reduce the size of the print. Active when fitToPage is false |
| fitToWidth            | 1             | How many pages wide the sheet should print on to. Active when fitToPage is true  |
| fitToHeight           | 1             | How many pages high the sheet should print on to. Active when fitToPage is true  |
| paperSize             |               | What paper size to use (see below) |
| showRowColHeaders     | false         | Whether to show the row numbers and column letters |
| showGridLines         | false         | Whether to show grid lines |
| firstPageNumber       |               | Which number to use for the first page |
| horizontalCentered    | false         | Whether to center the sheet data horizontally |
| verticalCentered      | false         | Whether to center the sheet data vertically |

**Example Paper Sizes**

| Name                          | Value     |
| ----------------------------- | --------- |
| Letter                        | undefined |
| Legal                         |  5        |
| Executive                     |  7        |
| A4                            |  9        |
| A5                            |  11       |
| B5 (JIS)                      |  13       |
| Envelope #10                  |  20       |
| Envelope DL                   |  27       |
| Envelope C5                   |  28       |
| Envelope B5                   |  34       |
| Envelope Monarch              |  37       |
| Double Japan Postcard Rotated |  82       |
| 16K 197x273 mm                |  119      |

## Worksheet Views

Worksheets now support a list of views, that control how Excel presents the sheet:

* frozen - where a number of rows and columns to the top and left are frozen in place. Only the bottom left section will scroll
* split - where the view is split into 4 sections, each semi-independently scrollable.

Each view also supports various properties:

| Name              | Default   | Description |
| ----------------- | --------- | ----------- |
| state             | 'normal'  | Controls the view state - one of normal, frozen or split |
| rightToLeft       | false     | Sets the worksheet view's orientation to right-to-left |
| activeCell        | undefined | The currently selected cell |
| showRuler         | true      | Shows or hides the ruler in Page Layout |
| showRowColHeaders | true      | Shows or hides the row and column headers (e.g. A1, B1 at the top and 1,2,3 on the left |
| showGridLines     | true      | Shows or hides the gridlines (shown for cells where borders have not been defined) |
| zoomScale         | 100       | Percentage zoom to use for the view |
| zoomScaleNormal   | 100       | Normal zoom for the view |
| style             | undefined | Presentation style - one of pageBreakPreview or pageLayout. Note pageLayout is not compatable with frozen views |

### Frozen Views

Frozen views support the following extra properties:

| Name              | Default   | Description |
| ----------------- | --------- | ----------- |
| xSplit            | 0         | How many columns to freeze. To freeze rows only, set this to 0 or undefined |
| ySplit            | 0         | How many rows to freeze. To freeze columns only, set this to 0 or undefined |
| topLeftCell       | special   | Which cell will be top-left in the bottom-right pane. Note: cannot be a frozen cell. Defaults to first unfrozen cell |

```javascript
worksheet.views = [
    {state: 'frozen', xSplit: 2, ySplit: 3, topLeftCell: 'G10', activeCell: 'A1'}
];
```

### Split Views

Split views support the following extra properties:

| Name              | Default   | Description |
| ----------------- | --------- | ----------- |
| xSplit            | 0         | How many points from the left to place the splitter. To split vertically, set this to 0 or undefined |
| ySplit            | 0         | How many points from the top to place the splitter. To split horizontally, set this to 0 or undefined  |
| topLeftCell       | undefined | Which cell will be top-left in the bottom-right pane. |
| activePane        | undefined | Which pane will be active - one of topLeft, topRight, bottomLeft and bottomRight |

```javascript
worksheet.views = [
    {state: 'split', xSplit: 2000, ySplit: 3000, topLeftCell: 'G10', activeCell: 'A1'}
];
```

## Auto filters

It is possible to apply an auto filter to your worksheet.

```javascript
worksheet.autoFilter = 'A1:C1';
```

While the range string is the standard form of the autoFilter, the worksheet will also support the
following values:

```javascript
// Set an auto filter from A1 to C1
worksheet.autoFilter = {
    from: 'A1',
    to: 'C1',
}

// Set an auto filter from the cell in row 3 and column 1
// to the cell in row 5 and column 12
worksheet.autoFilter = {
    from: {
        row: 3,
        column: 1
    },
    to: {
        row: 5,
        column: 12
    }
}

// Set an auto filter from D3 to the
// cell in row 7 and column 5
worksheet.autoFilter = {
    from: 'D3',
    to: {
        row: 7,
        column: 5
    }
}
```

## Columns

```javascript
// Add column headers and define column keys and widths
// Note: these column structures are a workbook-building convenience only,
// apart from the column width, they will not be fully persisted.
worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
    { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 }
];

// Access an individual columns by key, letter and 1-based column number
var idCol = worksheet.getColumn('id');
var nameCol = worksheet.getColumn('B');
var dobCol = worksheet.getColumn(3);

// set column properties

// Note: will overwrite cell value C1
dobCol.header = 'Date of Birth';

// Note: this will overwrite cell values C1:C2
dobCol.header = ['Date of Birth', 'A.K.A. D.O.B.'];

// from this point on, this column will be indexed by 'dob' and not 'DOB'
dobCol.key = 'dob';

dobCol.width = 15;

// Hide the column if you'd like
dobCol.hidden = true;

// set an outline level for columns
worksheet.getColumn(4).outlineLevel = 0;
worksheet.getColumn(5).outlineLevel = 1;

// columns support a readonly field to indicate the collapsed state based on outlineLevel
expect(worksheet.getColumn(4).collapsed).to.equal(false);
expect(worksheet.getColumn(5).collapsed).to.equal(true);

// iterate over all current cells in this column
dobCol.eachCell(function(cell, rowNumber) {
    // ...
});

// iterate over all current cells in this column including empty cells
dobCol.eachCell({ includeEmpty: true }, function(cell, rowNumber) {
    // ...
});

// add a column of new values
worksheet.getColumn(6).values = [1,2,3,4,5];

// add a sparse column of values
worksheet.getColumn(7).values = [,,2,3,,5,,7,,,,11];

// cut one or more columns (columns to the right are shifted left)
// If column properties have been definde, they will be cut or moved accordingly
// Known Issue: If a splice causes any merged cells to move, the results may be unpredictable
worksheet.spliceColumns(3,2);

// remove one column and insert two more.
// Note: columns 4 and above will be shifted right by 1 column.
// Also: If the worksheet has more rows than values in the colulmn inserts,
//  the rows will still be shifted as if the values existed
var newCol3Values = [1,2,3,4,5];
var newCol4Values = ['one', 'two', 'three', 'four', 'five'];
worksheet.spliceColumns(3, 1, newCol3Values, newCol4Values);

```

## Rows

```javascript
// Add a couple of Rows by key-value, after the last current row, using the column keys
worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970,1,1)});
worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965,1,7)});

// Add a row by contiguous Array (assign to columns A, B & C)
worksheet.addRow([3, 'Sam', new Date()]);

// Add a row by sparse Array (assign to columns A, E & I)
var rowValues = [];
rowValues[1] = 4;
rowValues[5] = 'Kyle';
rowValues[9] = new Date();
worksheet.addRow(rowValues);

// Add an array of rows
var rows = [
    [5,'Bob',new Date()], // row by array
    {id:6, name: 'Barbara', dob: new Date()}
];
worksheet.addRows(rows);

// Get a row object. If it doesn't already exist, a new empty one will be returned
var row = worksheet.getRow(5);

// Get the last editable row in a worksheet (or undefined if there are none)
var row = worksheet.lastRow;

// Set a specific row height
row.height = 42.5;

// make row hidden
row.hidden = true;

// set an outline level for rows
worksheet.getRow(4).outlineLevel = 0;
worksheet.getRow(5).outlineLevel = 1;

// rows support a readonly field to indicate the collapsed state based on outlineLevel
expect(worksheet.getRow(4).collapsed).to.equal(false);
expect(worksheet.getRow(5).collapsed).to.equal(true);


row.getCell(1).value = 5; // A5's value set to 5
row.getCell('name').value = 'Zeb'; // B5's value set to 'Zeb' - assuming column 2 is still keyed by name
row.getCell('C').value = new Date(); // C5's value set to now

// Get a row as a sparse array
// Note: interface change: worksheet.getRow(4) ==> worksheet.getRow(4).values
row = worksheet.getRow(4).values;
expect(row[5]).toEqual('Kyle');

// assign row values by contiguous array (where array element 0 has a value)
row.values = [1,2,3];
expect(row.getCell(1).value).toEqual(1);
expect(row.getCell(2).value).toEqual(2);
expect(row.getCell(3).value).toEqual(3);

// assign row values by sparse array  (where array element 0 is undefined)
var values = []
values[5] = 7;
values[10] = 'Hello, World!';
row.values = values;
expect(row.getCell(1).value).toBeNull();
expect(row.getCell(5).value).toEqual(7);
expect(row.getCell(10).value).toEqual('Hello, World!');

// assign row values by object, using column keys
row.values = {
    id: 13,
    name: 'Thing 1',
    dob: new Date()
};

// Insert a page break prior to the row
row.addPageBreak();

// Iterate over all rows that have values in a worksheet
worksheet.eachRow(function(row, rowNumber) {
    console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
});

// Iterate over all rows (including empty rows) in a worksheet
worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
});

// Iterate over all non-null cells in a row
row.eachCell(function(cell, colNumber) {
    console.log('Cell ' + colNumber + ' = ' + cell.value);
});

// Iterate over all cells in a row (including empty cells)
row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
    console.log('Cell ' + colNumber + ' = ' + cell.value);
});

// Cut one or more rows (rows below are shifted up)
// Known Issue: If a splice causes any merged cells to move, the results may be unpredictable
worksheet.spliceRows(4,3);

// remove one row and insert two more.
// Note: rows 4 and below will be shifted down by 1 row.
var newRow3Values = [1,2,3,4,5];
var newRow4Values = ['one', 'two', 'three', 'four', 'five'];
worksheet.spliceRows(3, 1, newRow3Values, newRow4Values);

// Cut one or more cells (cells to the right are shifted left)
// Note: this operation will not affect other rows
row.splice(3,2);

// remove one cell and insert two more (cells to the right of the cut cell will be shifted right)
row.splice(4,1,'new value 1', 'new value 2');

// Commit a completed row to stream
row.commit();

// row metrics
var rowSize = row.cellCount;
var numValues = row.actualCellCount;
```

## Handling Individual Cells

```javascript
// Modify/Add individual cell
worksheet.getCell('C3').value = new Date(1968, 5, 1);

// query a cell's type
expect(worksheet.getCell('C3').type).toEqual(Excel.ValueType.Date);
```

## Merged Cells

```javascript
// merge a range of cells
worksheet.mergeCells('A4:B5');

// ... merged cells are linked
worksheet.getCell('B5').value = 'Hello, World!';
expect(worksheet.getCell('B5').value).toBe(worksheet.getCell('A4').value);
expect(worksheet.getCell('B5').master).toBe(worksheet.getCell('A4'));

// ... merged cells share the same style object
expect(worksheet.getCell('B5').style).toBe(worksheet.getCell('A4').style);
worksheet.getCell('B5').style.font = myFonts.arial;
expect(worksheet.getCell('A4').style.font).toBe(myFonts.arial);

// unmerging the cells breaks the style links
worksheet.unMergeCells('A4');
expect(worksheet.getCell('B5').style).not.toBe(worksheet.getCell('A4').style);
expect(worksheet.getCell('B5').style.font).not.toBe(myFonts.arial);

// merge by top-left, bottom-right
worksheet.mergeCells('G10', 'H11');
worksheet.mergeCells(10,11,12,13); // top,left,bottom,right
```

## Defined Names

Individual cells (or multiple groups of cells) can have names assigned to them.
 The names can be used in formulas and data validation (and probably more).

```javascript
// assign (or get) a name for a cell (will overwrite any other names that cell had)
worksheet.getCell('A1').name = 'PI';
expect(worksheet.getCell('A1').name).to.equal('PI');

// assign (or get) an array of names for a cell (cells can have more than one name)
worksheet.getCell('A1').names = ['thing1', 'thing2'];
expect(worksheet.getCell('A1').names).to.have.members(['thing1', 'thing2']);

// remove a name from a cell
worksheet.getCell('A1').removeName('thing1');
expect(worksheet.getCell('A1').names).to.have.members(['thing2']);
```

## Data Validations

Cells can define what values are valid or not and provide prompting to the user to help guide them.

Validation types can be one of the following:

| Type       | Description |
| ---------- | ----------- |
| list       | Define a discrete set of valid values. Excel will offer these in a dropdown for easy entry |
| whole      | The value must be a whole number |
| decimal    | The value must be a decimal number |
| textLength | The value may be text but the length is controlled |
| custom     | A custom formula controls the valid values |

For types other than list or custom, the following operators affect the validation:

| Operator              | Description |
| --------------------  | ----------- |
| between               | Values must lie between formula results |
| notBetween            | Values must not lie between formula results |
| equal                 | Value must equal formula result |
| notEqual              | Value must not equal formula result |
| greaterThan           | Value must be greater than formula result |
| lessThan              | Value must be less than formula result |
| greaterThanOrEqual    | Value must be greater than or equal to formula result |
| lessThanOrEqual       | Value must be less than or equal to formula result |

```javascript
// Specify list of valid values (One, Two, Three, Four).
// Excel will provide a dropdown with these values.
worksheet.getCell('A1').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"One,Two,Three,Four"']
};

// Specify list of valid values from a range.
// Excel will provide a dropdown with these values.
    worksheet.getCell('A1').dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['$D$5:$F$5']
};

// Specify Cell must be a whole number that is not 5.
// Show the user an appropriate error message if they get it wrong
worksheet.getCell('A1').dataValidation = {
    type: 'whole',
    operator: 'notEqual',
    showErrorMessage: true,
    formulae: [5],
    errorStyle: 'error',
    errorTitle: 'Five',
    error: 'The value must not be Five'
};

// Specify Cell must be a decomal number between 1.5 and 7.
// Add 'tooltip' to help guid the user
worksheet.getCell('A1').dataValidation = {
    type: 'decimal',
    operator: 'between',
    allowBlank: true,
    showInputMessage: true,
    formulae: [1.5, 7],
    promptTitle: 'Decimal',
    prompt: 'The value must between 1.5 and 7'
};

// Specify Cell must be have a text length less than 15
worksheet.getCell('A1').dataValidation = {
    type: 'textLength',
    operator: 'lessThan',
    showErrorMessage: true,
    allowBlank: true,
    formulae: [15]
};

// Specify Cell must be have be a date before 1st Jan 2016
worksheet.getCell('A1').dataValidation = {
    type: 'date',
    operator: 'lessThan',
    showErrorMessage: true,
    allowBlank: true,
    formulae: [new Date(2016,0,1)]
};
```

## Styles

Cells, Rows and Columns each support a rich set of styles and formats that affect how the cells are displayed.

Styles are set by assigning the following properties:

* <a href="#number-formats">numFmt</a>
* <a href="#fonts">font</a>
* <a href="#alignment">alignment</a>
* <a href="#borders">border</a>
* <a href="#fills">fill</a>

```javascript
// assign a style to a cell
ws.getCell('A1').numFmt = '0.00%';

// Apply styles to worksheet columns
ws.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32, style: { font: { name: 'Arial Black' } } },
    { header: 'D.O.B.', key: 'DOB', width: 10, style: { numFmt: 'dd/mm/yyyy' } }
];

// Set Column 3 to Currency Format
ws.getColumn(3).numFmt = '"£"#,##0.00;[Red]\-"£"#,##0.00';

// Set Row 2 to Comic Sans.
ws.getRow(2).font = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true };
```

When a style is applied to a row or column, it will be applied to all currently existing cells in that row or column.
 Also, any new cell that is created will inherit its initial styles from the row and column it belongs to.

If a cell's row and column both define a specific style (e.g. font), the cell will use the row style over the column style.
 However if the row and column define different styles (e.g. column.numFmt and row.font), the cell will inherit the font from the row and the numFmt from the column.

Caveat: All the above properties (with the exception of numFmt, which is a string), are JS object structures.
 If the same style object is assigned to more than one spreadsheet entity, then each entity will share the same style object.
 If the style object is later modified before the spreadsheet is serialized, then all entities referencing that style object will be modified too.
 This behaviour is intended to prioritize performance by reducing the number of JS objects created.
 If you want the style objects to be independent, you will need to clone them before assigning them.
 Also, by default, when a document is read from file (or stream) if spreadsheet entities share similar styles, then they will reference the same style object too.

### Number Formats

```javascript
// display value as '1 3/5'
ws.getCell('A1').value = 1.6;
ws.getCell('A1').numFmt = '# ?/?';

// display value as '1.60%'
ws.getCell('B1').value = 0.016;
ws.getCell('B1').numFmt = '0.00%';
```

### Fonts

```javascript

// for the wannabe graphic designers out there
ws.getCell('A1').font = {
    name: 'Comic Sans MS',
    family: 4,
    size: 16,
    underline: true,
    bold: true
};

// for the graduate graphic designers...
ws.getCell('A2').font = {
    name: 'Arial Black',
    color: { argb: 'FF00FF00' },
    family: 2,
    size: 14,
    italic: true
};

// note: the cell will store a reference to the font object assigned.
// If the font object is changed afterwards, the cell font will change also...
var font = { name: 'Arial', size: 12 };
ws.getCell('A3').font = font;
font.size = 20; // Cell A3 now has font size 20!

// Cells that share similar fonts may reference the same font object after
// the workbook is read from file or stream
```

| Font Property | Description       | Example Value(s) |
| ------------- | ----------------- | ---------------- |
| name          | Font name. | 'Arial', 'Calibri', etc. |
| family        | Font family for fallback. An integer value. | 1 - Serif, 2 - Sans Serif, 3 - Mono, Others - unknown |
| scheme        | Font scheme. | 'minor', 'major', 'none' |
| charset       | Font charset. An integer value. | 1, 2, etc. |
| color         | Colour description, an object containing an ARGB value. | { argb: 'FFFF0000'} |
| bold          | Font **weight** | true, false |
| italic        | Font *slope* | true, false |
| underline     | Font <u>underline</u> style | true, false, 'none', 'single', 'double', 'singleAccounting', 'doubleAccounting' |
| strike        | Font <strike>strikethrough</strike> | true, false |
| outline       | Font outline | true, false |

### Alignment

```javascript
// set cell alignment to top-left, middle-center, bottom-right
ws.getCell('A1').alignment = { vertical: 'top', horizontal: 'left' };
ws.getCell('B1').alignment = { vertical: 'middle', horizontal: 'center' };
ws.getCell('C1').alignment = { vertical: 'bottom', horizontal: 'right' };

// set cell to wrap-text
ws.getCell('D1').alignment = { wrapText: true };

// set cell indent to 1
ws.getCell('E1').alignment = { indent: 1 };

// set cell text rotation to 30deg upwards, 45deg downwards and vertical text
ws.getCell('F1').alignment = { textRotation: 30 };
ws.getCell('G1').alignment = { textRotation: -45 };
ws.getCell('H1').alignment = { textRotation: 'vertical' };
```

**Valid Alignment Property Values**

| horizontal       | vertical    | wrapText | indent  | readingOrder | textRotation |
| ---------------- | ----------- | -------- | ------- | ------------ | ------------ |
| left             | top         | true     | integer | rtl          | 0 to 90      |
| center           | middle      | false    |         | ltr          | -1 to -90    |
| right            | bottom      |          |         |              | vertical     |
| fill             | distributed |          |         |              |              |
| justify          | justify     |          |         |              |              |
| centerContinuous |             |          |         |              |              |
| distributed      |             |          |         |              |              |


### Borders

```javascript
// set single thin border around A1
ws.getCell('A1').border = {
    top: {style:'thin'},
    left: {style:'thin'},
    bottom: {style:'thin'},
    right: {style:'thin'}
};

// set double thin green border around A3
ws.getCell('A3').border = {
    top: {style:'double', color: {argb:'FF00FF00'}},
    left: {style:'double', color: {argb:'FF00FF00'}},
    bottom: {style:'double', color: {argb:'FF00FF00'}},
    right: {style:'double', color: {argb:'FF00FF00'}}
};

// set thick red cross in A5
ws.getCell('A5').border = {
    diagonal: {up: true, down: true, style:'thick', color: {argb:'FFFF0000'}}
};
```

**Valid Border Styles**

* thin
* dotted
* dashDot
* hair
* dashDotDot
* slantDashDot
* mediumDashed
* mediumDashDotDot
* mediumDashDot
* medium
* double
* thick

### Fills

```javascript
// fill A1 with red darkVertical stripes
ws.getCell('A1').fill = {
    type: 'pattern',
    pattern:'darkVertical',
    fgColor:{argb:'FFFF0000'}
};

// fill A2 with yellow dark trellis and blue behind
ws.getCell('A2').fill = {
    type: 'pattern',
    pattern:'darkTrellis',
    fgColor:{argb:'FFFFFF00'},
    bgColor:{argb:'FF0000FF'}
};

// fill A3 with blue-white-blue gradient from left to right
ws.getCell('A3').fill = {
    type: 'gradient',
    gradient: 'angle',
    degree: 0,
    stops: [
        {position:0, color:{argb:'FF0000FF'}},
        {position:0.5, color:{argb:'FFFFFFFF'}},
        {position:1, color:{argb:'FF0000FF'}}
    ]
};


// fill A4 with red-green gradient from center
ws.getCell('A2').fill = {
    type: 'gradient',
    gradient: 'path',
    center:{left:0.5,top:0.5},
    stops: [
        {position:0, color:{argb:'FFFF0000'}},
        {position:1, color:{argb:'FF00FF00'}}
    ]
};
```

#### Pattern Fills

| Property | Required | Description |
| -------- | -------- | ----------- |
| type     | Y        | Value: 'pattern'<br/>Specifies this fill uses patterns |
| pattern  | Y        | Specifies type of pattern (see <a href="#valid-pattern-types">Valid Pattern Types</a> below) |
| fgColor  | N        | Specifies the pattern foreground color. Default is black. |
| bgColor  | N        | Specifies the pattern background color. Default is white. |

**Valid Pattern Types**

* none
* solid
* darkVertical
* darkGray
* mediumGray
* lightGray
* gray125
* gray0625
* darkHorizontal
* darkVertical
* darkDown
* darkUp
* darkGrid
* darkTrellis
* lightHorizontal
* lightVertical
* lightDown
* lightUp
* lightGrid
* lightTrellis
* lightGrid

#### Gradient Fills

| Property | Required | Description |
| -------- | -------- | ----------- |
| type     | Y        | Value: 'gradient'<br/>Specifies this fill uses gradients |
| gradient | Y        | Specifies gradient type. One of ['angle', 'path'] |
| degree   | angle    | For 'angle' gradient, specifies the direction of the gradient. 0 is from the left to the right. Values from 1 - 359 rotates the direction clockwise |
| center   | path     | For 'path' gradient. Specifies the relative coordinates for the start of the path. 'left' and 'top' values range from 0 to 1 |
| stops    | Y        | Specifies the gradient colour sequence. Is an array of objects containing position and color starting with position 0 and ending with position 1. Intermediary positions may be used to specify other colours on the path. |

**Caveats**

Using the interface above it may be possible to create gradient fill effects not possible using the XLSX editor program.
For example, Excel only supports angle gradients of 0, 45, 90 and 135.
Similarly the sequence of stops may also be limited by the UI with positions [0,1] or [0,0.5,1] as the only options.
Take care with this fill to be sure it is supported by the target XLSX viewers.

#### Rich Text

Individual cells now support rich text or in-cell formatting.
 Rich text values can control the font properties of any number of sub-strings within the text value.
 See <a href="font">Fonts</a> for a complete list of details on what font properties are supported.

```javascript

ws.getCell('A1').value = {
  'richText': [
     {'font': {'size': 12,'color': {'theme': 0},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'This is '},
     {'font': {'italic': true,'size': 12,'color': {'theme': 0},'name': 'Calibri','scheme': 'minor'},'text': 'a'},
     {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' '},
     {'font': {'size': 12,'color': {'argb': 'FFFF6600'},'name': 'Calibri','scheme': 'minor'},'text': 'colorful'},
     {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' text '},
     {'font': {'size': 12,'color': {'argb': 'FFCCFFCC'},'name': 'Calibri','scheme': 'minor'},'text': 'with'},
     {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' in-cell '},
     {'font': {'bold': true,'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'format'}
  ]
};

expect(ws.getCell('A1').text).to.equal('This is a colorful text with in-cell format');
expect(ws.getCell('A1').type).to.equal(Excel.ValueType.RichText);

```

## Outline Levels

Excel supports outlining; where rows or columns can be expanded or collapsed depending on what level of detail the user wishes to view.

Outline levels can be defined in column setup:
```javascript
worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
    { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 }
];
```

Or directly on the row or column
```javascript
worksheet.getColumn(3).outlineLevel = 1;
worksheet.getRow(3).outlineLevel = 1;
```

The sheet outline levels can be set on the worksheet
```javascript
// set column outline level
worksheet.properties.outlineLevelCol = 1;

// set row outline level
worksheet.properties.outlineLevelRow = 1;
```

Note: adjusting outline levels on rows or columns or the outline levels on the worksheet will incur a side effect of also modifying the collapsed property of all rows or columns affected by the property change. E.g.:
```javascript
worksheet.properties.outlineLevelCol = 1;

worksheet.getColumn(3).outlineLevel = 1;
expect(worksheet.getColumn(3).collapsed).to.be.true;

worksheet.properties.outlineLevelCol = 2;
expect(worksheet.getColumn(3).collapsed).to.be.false;
```

The outline properties can be set on the worksheet

```javascript
worksheet.properties.outlineProperties = {
  summaryBelow: false,
  summaryRight: false,
};
```

## Images

Adding images to a worksheet is a two-step process.
First, the image is added to the workbook via the addImage() function which will also return an imageId value.
Then, using the imageId, the image can be added to the worksheet either as a tiled background or covering a cell range.

Note: As of this version, adjusting or transforming the image is not supported.

### Add Image to Workbook

The Workbook.addImage function supports adding images by filename or by Buffer.
Note that in both cases, the extension must be specified.
Valid extension values include 'jpeg', 'png', 'gif'.

```javascript
// add image to workbook by filename
var imageId1 = workbook.addImage({
  filename: 'path/to/image.jpg',
  extension: 'jpeg',
});

// add image to workbook by buffer
var imageId2 = workbook.addImage({
  buffer: fs.readFileSync('path/to.image.png'),
  extension: 'png',
});

// add image to workbook by base64
var myBase64Image = "data:image/png;base64,iVBORw0KG...";
var imageId2 = workbook.addImage({
  base64: myBase64Image,
  extension: 'png',
});
```

### Add image background to worksheet

Using the image id from Workbook.addImage, the background to a worksheet can be set using the addBackgroundImage function

```javascript
// set background
worksheet.addBackgroundImage(imageId1);
```

### Add image over a range

Using the image id from Workbook.addImage, an image can be embedded within the worksheet to cover a range.
The coordinates calculated from the range will cover from the top-left of the first cell to the bottom right of the second.

```javascript
// insert an image over B2:D6
worksheet.addImage(imageId2, 'B2:D6');
```

Using a structure instead of a range string, it is possible to partially cover cells.

Note that the coordinate system used for this is zero based, so the top-left of A1 will be { col: 0, row: 0 }.
Fractions of cells can be specified by using floating point numbers, e.g. the midpoint of A1 is { col: 0.5, row: 0.5 }.

```javascript
// insert an image over part of B2:D6
worksheet.addImage(imageId2, {
  tl: { col: 1.5, row: 1.5 },
  br: { col: 3.5, row: 5.5 }
});
```

The cell range can also have the eproperty 'editAs' which will control how the image is anchored to the cell(s)
It can have one of the following values:

| Value     | Description |
| --------- | ----------- |
| undefined | This is the default. It specifies the image will be moved and sized with cells |
| oneCell   | Image will be moved with cells but not sized |
| absolute  | Image will not be moved or sized with cells |

```javascript
ws.addImage(imageId, {
  tl: { col: 0.1125, row: 0.4 },
  br: { col: 2.101046875, row: 3.4 },
  editAs: 'oneCell'
});
```

## File I/O

### XLSX

#### Reading XLSX

```javascript
// read from a file
var workbook = new Excel.Workbook();
workbook.xlsx.readFile(filename)
    .then(function() {
        // use workbook
    });

// pipe from stream
var workbook = new Excel.Workbook();
stream.pipe(workbook.xlsx.createInputStream());
```

#### Writing XLSX

```javascript
// write to a file
var workbook = createAndFillWorkbook();
workbook.xlsx.writeFile(filename)
    .then(function() {
        // done
    });

// write to a stream
workbook.xlsx.write(stream)
    .then(function() {
        // done
    });
```

### CSV

#### Reading CSV

```javascript
// read from a file
var workbook = new Excel.Workbook();
workbook.csv.readFile(filename)
    .then(function(worksheet) {
        // use workbook or worksheet
    });

// read from a stream
var workbook = new Excel.Workbook();
workbook.csv.read(stream)
    .then(function(worksheet) {
        // use workbook or worksheet
    });

// pipe from stream
var workbook = new Excel.Workbook();
stream.pipe(workbook.csv.createInputStream());

// read from a file with European Dates
var workbook = new Excel.Workbook();
var options = {
    dateFormats: ['DD/MM/YYYY']
};
workbook.csv.readFile(filename, options)
    .then(function(worksheet) {
        // use workbook or worksheet
    });

// read from a file with custom value parsing
var workbook = new Excel.Workbook();
var options = {
    map: function(value, index) {
        switch(index) {
            case 0:
                // column 1 is string
                return value;
            case 1:
                // column 2 is a date
                return new Date(value);
            case 2:
                // column 3 is JSON of a formula value
                return JSON.parse(value);
            default:
                // the rest are numbers
                return parseFloat(value);
        }
    }
};
workbook.csv.readFile(filename, options)
    .then(function(worksheet) {
        // use workbook or worksheet
    });
```

The CSV parser uses [fast-csv](https://www.npmjs.com/package/fast-csv) to read the CSV file.
 The options passed into the read functions above is also passed to fast-csv for parsing of the csv data.
 Please refer to the fast-csv README.md for details.

Dates are parsed using the npm module [moment](https://www.npmjs.com/package/moment).
 If no dateFormats are supplied, the following are used:

* moment.ISO_8601
* 'MM-DD-YYYY'
* 'YYYY-MM-DD'

#### Writing CSV

```javascript

// write to a file
var workbook = createAndFillWorkbook();
workbook.csv.writeFile(filename)
    .then(function() {
        // done
    });

// write to a stream
// Be careful that you need to provide sheetName or
// sheetId for correct import to csv.
workbook.csv.write(stream, { sheetName: 'Page name' })
    .then(function() {
        // done
    });

// read from a file with European Date-Times
var workbook = new Excel.Workbook();
var options = {
    dateFormat: 'DD/MM/YYYY HH:mm:ss',
    dateUTC: true, // use utc when rendering dates
};
workbook.csv.readFile(filename, options)
    .then(function(worksheet) {
        // use workbook or worksheet
    });


// read from a file with custom value formatting
var workbook = new Excel.Workbook();
var options = {
    map: function(value, index) {
        switch(index) {
            case 0:
                // column 1 is string
                return value;
            case 1:
                // column 2 is a date
                return moment(value).format('YYYY-MM-DD');
            case 2:
                // column 3 is a formula, write just the result
                return value.result;
            default:
                // the rest are numbers
                return value;
        }
    }
};
workbook.csv.readFile(filename, options)
    .then(function(worksheet) {
        // use workbook or worksheet
    });
```

The CSV parser uses [fast-csv](https://www.npmjs.com/package/fast-csv) to write the CSV file.
 The options passed into the write functions above is also passed to fast-csv for writing the csv data.
 Please refer to the fast-csv README.md for details.

Dates are formatted using the npm module [moment](https://www.npmjs.com/package/moment).
 If no dateFormat is supplied, moment.ISO_8601 is used.
 When writing a CSV you can supply the boolean dateUTC as true to have ExcelJS parse the date without automatically
 converting the timezone using `moment.utc()`.

### Streaming I/O

The File I/O documented above requires that an entire workbook is built up in memory before the file can be written.
 While convenient, it can limit the size of the document due to the amount of memory required.

A streaming writer (or reader) processes the workbook or worksheet data as it is generated,
 converting it into file form as it goes. Typically this is much more efficient on memory as the final
 memory footprint and even intermediate memory footprints are much more compact than with the document version,
 especially when you consider that the row and cell objects are disposed once they are committed.

The interface to the streaming workbook and worksheet is almost the same as the document versions with a few minor practical differences:

* Once a worksheet is added to a workbook, it cannot be removed.
* Once a row is committed, it is no longer accessible since it will have been dropped from the worksheet.
* unMergeCells() is not supported.

Note that it is possible to build the entire workbook without committing any rows.
 When the workbook is committed, all added worksheets (including all uncommitted rows) will be automatically committed.
 However in this case, little will have been gained over the Document version.

#### Streaming XLSX

##### Streaming XLSX Writer

The streaming XLSX writer is available in the ExcelJS.stream.xlsx namespace.

The constructor takes an optional options object with the following fields:

| Field            | Description |
| ---------------- | ----------- |
| stream           | Specifies a writable stream to write the XLSX workbook to. |
| filename         | If stream not specified, this field specifies the path to a file to write the XLSX workbook to. |
| useSharedStrings | Specifies whether to use shared strings in the workbook. Default is false |
| useStyles        | Specifies whether to add style information to the workbook. Styles can add some performance overhead. Default is false |

If neither stream nor filename is specified in the options, the workbook writer will create a StreamBuf object
 that will store the contents of the XLSX workbook in memory.
 This StreamBuf object, which can be accessed via the property workbook.stream, can be used to either
 access the bytes directly by stream.read() or to pipe the contents to another stream.

```javascript
// construct a streaming XLSX workbook writer with styles and shared strings
var options = {
    filename: './streamed-workbook.xlsx',
    useStyles: true,
    useSharedStrings: true
};
var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
```

In general, the interface to the streaming XLSX writer is the same as the Document workbook (and worksheets)
 described above, in fact the row, cell and style objects are the same.

However there are some differences...

**Construction**

As seen above, the WorkbookWriter will typically require the output stream or file to be specified in the constructor.

**Committing Data**

When a worksheet row is ready, it should be committed so that the row object and contents can be freed.
 Typically this would be done as each row is added...

```javascript
worksheet.addRow({
   id: i,
   name: theName,
   etc: someOtherDetail
}).commit();
```

The reason the WorksheetWriter does not commit rows as they are added is to allow cells to be merged across rows:
```javascript
worksheet.mergeCells('A1:B2');
worksheet.getCell('A1').value = 'I am merged';
worksheet.getCell('C1').value = 'I am not';
worksheet.getCell('C2').value = 'Neither am I';
worksheet.getRow(2).commit(); // now rows 1 and two are committed.
```

As each worksheet is completed, it must also be committed:

```javascript
// Finished adding data. Commit the worksheet
worksheet.commit();
```

To complete the XLSX document, the workbook must be committed. If any worksheet in a workbook are uncommitted,
 they will be committed automatically as part of the workbook commit.

```javascript
// Finished the workbook.
workbook.commit()
  .then(function() {
    // the stream has been written
  });
```

# Browser

A portion of this library has been isolated and tested for use within a browser environment.

Due to the streaming nature of the workbook reader and workbook writer, these have not been included.
Only the document based workbook may be used (see <a href="#create-a-workbook">Create a Worbook</a> for details).

For example code using ExcelJS in the browser take a look at the <a href="https://github.com/guyonroche/exceljs/tree/master/spec/browser">spec/browser</a> folder in the github repo.

## Prebundled

The following files are pre-bundled and included inside the dist folder.

* exceljs.js
* exceljs.min.js

# Value Types

The following value types are supported.

## Null Value

Enum: Excel.ValueType.Null

A null value indicates an absence of value and will typically not be stored when written to file (except for merged cells).
  It can be used to remove the value from a cell.

E.g.

```javascript
worksheet.getCell('A1').value = null;
```

## Merge Cell

Enum: Excel.ValueType.Merge

A merge cell is one that has its value bound to another 'master' cell.
  Assigning to a merge cell will cause the master's cell to be modified.

## Number Value

Enum: Excel.ValueType.Number

A numeric value.

E.g.

```javascript
worksheet.getCell('A1').value = 5;
worksheet.getCell('A2').value = 3.14159;
```

## String Value

Enum: Excel.ValueType.String

A simple text string.

E.g.

```javascript
worksheet.getCell('A1').value = 'Hello, World!';
```

## Date Value

Enum: Excel.ValueType.Date

A date value, represented by the JavaScript Date type.

E.g.

```javascript
worksheet.getCell('A1').value = new Date(2017, 2, 15);
```

## Hyperlink Value

Enum: Excel.ValueType.Hyperlink

A URL with both text and link value.

E.g.
```javascript
// link to web
worksheet.getCell('A1').value = { text: 'www.mylink.com', hyperlink: 'http://www.mylink.com' };

// internal link
worksheet.getCell('A1').value = { text: 'Sheet2', hyperlink: '#\\"Sheet2\\"!A1' };
```

## Formula Value

Enum: Excel.ValueType.Formula

An Excel formula for calculating values on the fly.
  Note that while the cell type will be Formula, the cell may have an effectiveType value that will
  be derived from the result value.

Note that ExcelJS cannot process the formula to generate a result, it must be supplied.

E.g.

```javascript
worksheet.getCell('A3').value = { formula: 'A1+A2', result: 7 };
```

Cells also support convenience getters to access the formula and result:

```javascript
worksheet.getCell('A3').formula === 'A1+A2';
worksheet.getCell('A3').result === 7;
```

### Shared Formula

Shared formulae enhance the compression of the xlsx document by increasing the repetition
of text within the worksheet xml.

A shared formula can be assigned to a cell using a new value form:

```javascript
worksheet.getCell('B3').value = { sharedFormula: 'A3', result: 10 };
```

This specifies that the cell B3 is a formula that will be derived from the formula in
A3 and its result is 10.

The formula convenience getter will translate the formula in A3 to what it should be in B3:

```javascript
worksheet.getCell('B3').formula === 'B1+B2';
```

### Formula Type

To distinguish between real and translated formula cells, use the formulaType getter:

```javascript
worksheet.getCell('A3').formulaType === Enums.FormulaType.Master;
worksheet.getCell('B3').formulaType === Enums.FormulaType.Shared;
```

Formula type has the following values:

| Name                       |  Value  |
| -------------------------- | ------- |
| Enums.FormulaType.None     |   0     |
| Enums.FormulaType.Master   |   1     |
| Enums.FormulaType.Shared   |   2     |

## Rich Text Value

Enum: Excel.ValueType.RichText

Rich, styled text.

E.g.
```javascript
worksheet.getCell('A1').value = {
  richText: [
    { text: 'This is '},
    {font: {italic: true}, text: 'italic'},
  ]
};
```

## Boolean Value

Enum: Excel.ValueType.Boolean

E.g.

```javascript
worksheet.getCell('A1').value = true;
worksheet.getCell('A2').value = false;
```

## Error Value

Enum: Excel.ValueType.Error

E.g.

```javascript
worksheet.getCell('A1').value = { error: '#N/A' };
worksheet.getCell('A2').value = { error: '#VALUE!' };
```

The current valid Error text values are:

| Name                           | Value       |
| ------------------------------ | ----------- |
| Excel.ErrorValue.NotApplicable | #N/A        |
| Excel.ErrorValue.Ref           | #REF!       |
| Excel.ErrorValue.Name          | #NAME?      |
| Excel.ErrorValue.DivZero       | #DIV/0!     |
| Excel.ErrorValue.Null          | #NULL!      |
| Excel.ErrorValue.Value         | #VALUE!     |
| Excel.ErrorValue.Num           | #NUM!       |

# Interface Changes

Every effort is made to make a good consistent interface that doesn't break through the versions but regrettably, now and then some things have to change for the greater good.

## 0.1.0

### Worksheet.eachRow

The arguments in the callback function to Worksheet.eachRow have been swapped and changed; it was function(rowNumber,rowValues), now it is function(row, rowNumber) which gives it a look and feel more like the underscore (_.each) function and prioritises the row object over the row number.

### Worksheet.getRow

This function has changed from returning a sparse array of cell values to returning a Row object. This enables accessing row properties and will facilitate managing row styles and so on.

The sparse array of cell values is still available via Worksheet.getRow(rowNumber).values;

## 0.1.1

### cell.model

cell.styles renamed to cell.style

## 0.2.44

Promises returned from functions switched from Bluebird to native node Promise which can break calling code
 if they rely on Bluebird's extra features.

To mitigate this the following two changes were added to 0.3.0:

* A more fully featured and still browser compatable promise lib is used by default. This lib supports many of the features of Bluebird but with a much lower footprint.
* An option to inject a different Promise implementation. See <a href="#config">Config</a> section for more details.

# Config

ExcelJS now supports dependency injection for the promise library.
 You can restore Bluebird promises by including the following code in your module...

```javascript
ExcelJS.config.setValue('promise', require('bluebird'));
```

Please note: I have tested ExcelJS with bluebird specifically (since up until recently this was the library it used).
 From the tests I have done it will not work with Q.

# Known Issues

## Testing with PhantomJS

You may need to install phantomjs globally before running the browser-test script.

It's also possible that phantomjs will not run (or can't be found). If this happens, try the following:

```bash
sudo apt-get install libfontconfig
```

## Splice vs Merge

If any splice operation affects a merged cell, the merge group will not be moved correctly

# Release History

| Version | Changes |
| ------- | ------- |
| 0.0.9   | <ul><li><a href="#number-formats">Number Formats</a></li></ul> |
| 0.1.0   | <ul><li>Bug Fixes<ul><li>"&lt;" and "&gt;" text characters properly rendered in xlsx</li></ul></li><li><a href="#columns">Better Column control</a></li><li><a href="#rows">Better Row control</a></li></ul> |
| 0.1.1   | <ul><li>Bug Fixes<ul><li>More textual data written properly to xml (including text, hyperlinks, formula results and format codes)</li><li>Better date format code recognition</li></ul></li><li><a href="#fonts">Cell Font Style</a></li></ul> |
| 0.1.2   | <ul><li>Fixed potential race condition on zip write</li></ul> |
| 0.1.3   | <ul><li><a href="#alignment">Cell Alignment Style</a></li><li><a href="#rows">Row Height</a></li><li>Some Internal Restructuring</li></ul> |
| 0.1.5   | <ul><li>Bug Fixes<ul><li>Now handles 10 or more worksheets in one workbook</li><li>theme1.xml file properly added and referenced</li></ul></li><li><a href="#borders">Cell Borders</a></li></ul> |
| 0.1.6   | <ul><li>Bug Fixes<ul><li>More compatable theme1.xml included in XLSX file</li></ul></li><li><a href="#fills">Cell Fills</a></li></ul> |
| 0.1.8   | <ul><li>Bug Fixes<ul><li>More compatable theme1.xml included in XLSX file</li><li>Fixed filename case issue</li></ul></li><li><a href="#fills">Cell Fills</a></li></ul> |
| 0.1.9   | <ul><li>Bug Fixes<ul><li>Added docProps files to satisfy Mac Excel users</li><li>Fixed filename case issue</li><li>Fixed worksheet id issue</li></ul></li><li><a href="#set-workbook-properties">Core Workbook Properties</a></li></ul> |
| 0.1.10  | <ul><li>Bug Fixes<ul><li>Handles File Not Found error</li></ul></li><li><a href="#csv">CSV Files</a></li></ul> |
| 0.1.11  | <ul><li>Bug Fixes<ul><li>Fixed Vertical Middle Alignment Issue</li></ul></li><li><a href="#styles">Row and Column Styles</a></li><li><a href="#rows">Worksheet.eachRow supports options</a></li><li><a href="#rows">Row.eachCell supports options</a></li><li><a href="#columns">New function Column.eachCell</a></li></ul> |
| 0.2.0   | <ul><li><a href="#streaming-xlxs-writer">Streaming XLSX Writer</a><ul><li>At long last ExcelJS can support writing massive XLSX files in a scalable memory efficient manner. Performance has been optimised and even smaller spreadsheets can be faster to write than the document writer. Options have been added to control the use of shared strings and styles as these can both have a considerable effect on performance</li></ul></li><li><a href="#rows">Worksheet.lastRow</a><ul><li>Access the last editable row in a worksheet.</li></ul></li><li><a href="#rows">Row.commit()</a><ul><li>For streaming writers, this method commits the row (and any previous rows) to the stream. Committed rows will no longer be editable (and are typically deleted from the worksheet object). For Document type workbooks, this method has no effect.</li></ul></li></ul> |
| 0.2.2   | <ul><li><a href="https://pbs.twimg.com/profile_images/2933552754/fc8c70829ee964c5542ae16453503d37.jpeg">One Billion Cells</a><ul><li>Achievement Unlocked: A simple test using ExcelJS has created a spreadsheet with 1,000,000,000 cells. Made using random data with 100,000,000 rows of 10 cells per row. I cannot validate the file yet as Excel will not open it and I have yet to implement the streaming reader but I have every confidence that it is good since 1,000,000 rows loads ok.</li></ul></li></ul> |
| 0.2.3   | <ul><li>Bug Fixes<ul><li><a href="https://github.com/guyonroche/exceljs/issues/18">Merge Cell Styles</a><ul><li>Merged cells now persist (and parse) their styles.</li></ul></li></ul></li><li><a href="#streaming-xlxs-writer">Streaming XLSX Writer</a><ul><li>At long last ExcelJS can support writing massive XLSX files in a scalable memory efficient manner. Performance has been optimised and even smaller spreadsheets can be faster to write than the document writer. Options have been added to control the use of shared strings and styles as these can both have a considerable effect on performance</li></ul></li><li><a href="#rows">Worksheet.lastRow</a><ul><li>Access the last editable row in a worksheet.</li></ul></li><li><a href="#rows">Row.commit()</a><ul><li>For streaming writers, this method commits the row (and any previous rows) to the stream. Committed rows will no longer be editable (and are typically deleted from the worksheet object). For Document type workbooks, this method has no effect.</li></ul></li></ul> |
| 0.2.4   | <ul><li>Bug Fixes<ul><li><a href="https://github.com/guyonroche/exceljs/issues/27">Worksheets with Ampersand Names</a><ul><li>Worksheet names are now xml-encoded and should work with all xml compatable characters</li></ul></li></ul></li><li><a href="#rows">Row.hidden</a> & <a href="#columns">Column.hidden</a><ul><li>Rows and Columns now support the hidden attribute.</li></ul></li><li><a href="#worksheet">Worksheet.addRows</a><ul><li>New function to add an array of rows (either array or object form) to the end of a worksheet.</li></ul></li></ul> |
| 0.2.6   | <ul><li>Bug Fixes<ul><li><a href="https://github.com/guyonroche/exceljs/issues/87">invalid signature: 0x80014</a>: Thanks to <a href="https://github.com/hasanlussa">hasanlussa</a> for the PR</li></ul></li><li><a href="#defined-names">Defined Names</a><ul><li>Cells can now have assigned names which may then be used in formulas.</li></ul></li><li>Converted Bluebird.defer() to new Bluebird(function(resolve, reject){}). Thanks to user <a href="https://github.com/Nishchit14">Nishchit</a> for the Pull Request</li></ul> |
| 0.2.7   | <ul><li><a href="#data-validations">Data Validations</a><ul><li>Cells can now define validations that controls the valid values the cell can have</li></ul></li></ul> |
| 0.2.8   | <ul><li><a href="rich-text">Rich Text Value</a><ul><li>Cells now support <b><i>in-cell</i></b> formatting - Thanks to <a href="https://github.com/pvadam">Peter ADAM</a></li></ul></li><li>Fixed typo in README - Thanks to <a href="https://github.com/MRdNk">MRdNk</a></li><li>Fixing emit in worksheet-reader - Thanks to <a href="https://github.com/alangunning">Alan Gunning</a></li><li>Clearer Docs - Thanks to <a href="https://github.com/miensol">miensol</a></li></ul> |
| 0.2.9   | <ul><li>Fixed "read property 'richText' of undefined error. Thanks to  <a href="https://github.com/james075">james075</a></li></ul> |
| 0.2.10  | <ul><li>Refactoring Complete. All unit and integration tests pass.</li></ul> |
| 0.2.11  | <ul><li><a href="#outline-level">Outline Levels</a>. Thanks to <a href="https://github.com/cricri">cricri</a> for the contribution.</li><li><a href="#worksheet-properties">Worksheet Properties</a></li><li>Further refactoring of worksheet writer.</li></ul> |
| 0.2.12  | <ul><li><a href="#worksheet-views">Sheet Views</a>. Thanks to <a href="https://github.com/cricri">cricri</a> again for the contribution.</li></ul> |
| 0.2.13  | <ul><li>Fix for <a href="https://github.com/guyonroche/exceljs/issues">exceljs might be vulnerable for regular expression denial of service</a>. Kudos to <a href="https://github.com/yonjah">yonjah</a> and <a href="https://www.youtube.com/watch?v=wCfE-9bhY2Y">Josh Emerson</a> for the resolution.</li><li>Fix for <a href="https://github.com/guyonroche/exceljs/issues/162">Multiple Sheets opens in 'Group' mode in Excel</a>. My bad - overzealous sheet view code.</li><li>Also fix for empty sheet generating invalid xlsx.</li></ul> |
| 0.2.14  | <ul><li>Fix for <a href="https://github.com/guyonroche/exceljs/issues">exceljs might be vulnerable for regular expression denial of service</a>. Kudos to <a href="https://github.com/yonjah">yonjah</a> and <a href="https://www.youtube.com/watch?v=wCfE-9bhY2Y">Josh Emerson</a> for the resolution.</li><li>Fixed <a href="https://github.com/guyonroche/exceljs/issues/162">Multiple Sheets opens in 'Group' mode in Excel</a> again. Added <a href="#workbook-views">Workbook views</a>.</li><li>Also fix for empty sheet generating invalid xlsx.</li></ul> |
| 0.2.15  | <ul><li>Added <a href="#page-setup">Page Setup Properties</a>. Thanks to <a href="https://github.com/jackkum">Jackkum</a> for the PR</li></ul> |
| 0.2.16  | <ul><li>New <a href="#page-setup">Page Setup</a> Property: Print Area</li></ul> |
| 0.2.17  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/114">Fix a bug on phonetic characters</a>. This fixes an issue related to reading workbooks with phonetic text in. Note phonetic text is not properly supported yet - just properly ignored. Thanks to <a href="https://github.com/zephyrrider">zephyrrider</a> and <a href="https://github.com/gen6033">gen6033</a> for the contribution.</li></ul> |
| 0.2.18  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/175">Fix regression #150: Stream API fails to write XLSX files</a>. Apologies for the regression! Thanks to <a href="https://github.com/danieleds">danieleds</a> for the fix.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/114">Fix a bug on phonetic characters</a>. This fixes an issue related to reading workbooks with phonetic text in. Note phonetic text is not properly supported yet - just properly ignored. Thanks to <a href="https://github.com/zephyrrider">zephyrrider</a> and <a href="https://github.com/gen6033">gen6033</a> for the contribution.</li></ul> |
| 0.2.19  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/119">Update xlsx.js #119</a>. This should make parsing more resilient to open-office documents. Thanks to <a href="https://github.com/nvitaterna">nvitaterna</a> for the contribution.</li></ul> |
| 0.2.20  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/179">Changes from guyonroche/exceljs#127 applied to latest version #179</a>. Fixes parsing of defined name values. Thanks to <a href="https://github.com/agdevbridge">agdevbridge</a> and <a href="https://github.com/priitliivak">priitliivak</a> for the contribution.</li></ul> |
| 0.2.21  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/135">color tabs for worksheet-writer #135</a>. Modified the behaviour to print deprecation warning as tabColor has moved into options.properties. Thanks to <a href="https://github.com/ethanlook">ethanlook</a> for the contribution.</li></ul> |
| 0.2.22  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/136">Throw legible error when failing Value.getType() #136</a>. Thanks to <a href="https://github.com/wulfsolter">wulfsolter</a> for the contribution.</li><li>Honourable mention to contributors whose PRs were fixed before I saw them:<ul><li><a href="https://github.com/haoliangyu">haoliangyu</a></li><li><a href="https://github.com/wulfsolter">wulfsolter</a></li></ul></li></ul> |
| 0.2.23  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/137">Fall back to JSON.stringify() if unknown Cell.Type #137</a> with some modification. If a cell value is assigned to an unrecognisable javascript object, the stored value in xlsx and csv files will  be JSON stringified. Note that if the file is read again, no attempt will be made to parse the stringified JSON text. Thanks to <a href="https://github.com/wulfsolter">wulfsolter</a> for the contribution.</li></ul> |
| 0.2.24  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/166">Protect cell fix #166</a>. This does not mean full support for protected cells merely that the parser is not confused by the extra xml. Thanks to <a href="https://github.com/jayflo">jayflo</a> for the contribution.</li></ul> |
| 0.2.25  | <ul><li>Added functions to delete cells, rows and columns from a worksheet. Modelled after the Array splice method, the functions allow cells, rows and columns to be deleted (and optionally inserted). See <a href="#columns">Columns</a> and <a href="#rows">Rows</a> for details.<br />Note: <a href="#splice-vs-merge">Not compatable with cell merges</a></li></ul> |
| 0.2.26  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/184">Update border-xform.js #184</a>Border edges without style will be parsed and rendered as no-border. Thanks to <a href="https://github.com/skumarnk2">skumarnk2</a> for the contribution.</li></ul> |
| 0.2.27  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/187">Pass views to worksheet-writer #187</a>. Now also passes views to worksheet-writer. Thanks to <a href="https://github.com/Temetz">Temetz</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/189">Do not escape xml characters when using shared strings #189</a>. Fixing bug in shared strings. Thanks to <a href="https://github.com/tkirda">tkirda</a> for the contribution.</li></ul> |
| 0.2.28  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/190">Fix tiny bug [Update hyperlink-map.js] #190</a>Thanks to <a href="https://github.com/lszlkss">lszlkss</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/196">fix typo on sheet view showGridLines option #196</a> "showGridlines" should have been "showGridLines". Thanks to <a href="https://github.com/gadiaz1">gadiaz1</a> for the contribution.</li></ul> |
| 0.2.29  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/199">Fire finish event instead of end event on write stream #199</a> and <a href="https://github.com/guyonroche/exceljs/pull/200">Listen for finish event on zip stream instead of middle stream #200</a>. Fixes issues with stream completion events. Thanks to <a href="https://github.com/junajan">junajan</a> for the contribution.</li></ul> |
| 0.2.30  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/201">Fix issue #178 #201</a>. Adds the following properties to workbook:<ul><li>title</li><li>subject</li><li>keywords</li><li>category</li><li>description</li><li>company</li><li>manager</li></ul>Thanks to <a href="https://github.com/stavenko">stavenko</a> for the contribution.</li></ul> |
| 0.2.31  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/203">Fix issue #163: the "spans" attribute of the row element is optional #203</a>. Now xlsx parsing will handle documents without row spans. Thanks to <a href="https://github.com/arturas-vitkauskas">arturas-vitkauskas</a> for the contribution.</li></ul> |
| 0.2.32  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/208">Fix issue 206 #208</a>. Fixes issue reading xlsx files that have been printed. Also adds "lastPrinted" property to Workbook. Thanks to <a href="https://github.com/arturas-vitkauskas">arturas-vitkauskas</a> for the contribution.</li></ul> |
| 0.2.33  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/210">Allow styling of cells with no value. #210</a>. Includes Null type cells with style in the rendering parsing. Thanks to <a href="https://github.com/oferns">oferns</a> for the contribution.</li></ul> |
| 0.2.34  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/212">Fix "Unexpected xml node in parseOpen" bug in LibreOffice documents for attributes dc:language and cp:revision #212</a>. Thanks to <a href="https://github.com/jessica-jordan">jessica-jordan</a> for the contribution.</li></ul> |
| 0.2.35  | <ul><li>Fixed <a href="https://github.com/guyonroche/exceljs/issues/74">Getting a column/row count #74</a>. <a href="#worksheet-metrics">Worksheet</a> now has rowCount and columnCount properties (and actual variants), <a href="row">Row</a> has cellCount.</li></ul> |
| 0.2.36  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/217">Stream reader fixes #217</a>. Thanks to <a href="https://github.com/kturney">kturney</a> for the contribution.</li></ul> |
| 0.2.37  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/225">Fix output order of Sheet Properties #225</a>. Thanks to <a href="https://github.com/keeneym">keeneym</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/231">remove empty worksheet[0] from _worksheets #231</a>. Thanks to <a href="https://github.com/pookong">pookong</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/232">do not skip empty string in shared strings so that indexes match #232</a>. Thanks again to <a href="https://github.com/pookong">pookong</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/233">use shared strings for streamed writes #233</a>. Thanks again to <a href="https://github.com/pookong">pookong</a> for the contribution.</li></ul> |
| 0.2.38  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/236">Add a comment for issue #216 #236</a>. Thanks to <a href="https://github.com/jsalwen">jsalwen</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/237">Start on support for 1904 based dates #237</a>. Fixed date handling in documents with the 1904 flag set. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li></ul> |
| 0.2.39  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/245">Stops Bluebird warning about unreturned promise #245</a>. Thanks to <a href="https://github.com/robinbullocks4rb">robinbullocks4rb</a> for the contribution. </li> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/247">Added missing dependency: col-cache.js #247</a>. Thanks to <a href="https://github.com/Manish2005">Manish2005</a> for the contribution. </li> </ul> |
| 0.2.42  | <ul><li>Browser Compatable!<ul><li>Well mostly. I have added a browser sub-folder that contains a browserified bundle and an index.js that can be used to generate another. See <a href="#browser">Browser</a> section for details.</li></ul></li><li>Fixed corrupted theme.xml. Apologies for letting that through.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/253">[BUGFIX] data validation formulae undefined #253</a>. Thanks to <a href="https://github.com/jayflo">jayflo</a> for the contribution.</li></ul> |
| 0.2.43  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/255">added a (maybe partial) solution to issue 99. i wasn't able to create an appropriate test #255</a>. This fixes <a href="https://github.com/guyonroche/exceljs/issues/99">Too few data or empty worksheet generate malformed excel file #99</a>. Thanks to <a href="https://github.com/mminuti">mminuti</a> for the contribution.</li></ul> |
| 0.2.44  | <ul><li>Reduced Dependencies.<ul><li>Goodbye lodash, goodbye bluebird. Minified bundle is now just over half what it was in the first version.</li></ul></li></ul> |
| 0.2.45  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/256">Sheets with hyperlinks and data validations are corrupted #256</a>. Thanks to <a href="https://github.com/simon-stoic">simon-stoic</a> for the contribution.</li></ul> |
| 0.2.46  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/259">Exclude character controls from XML output. Fixes #234 #262</a>. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/262">Add support for identifier #259</a>. This fixes <a href="https://github.com/guyonroche/exceljs/issues/234">Broken XLSX because of "vertical tab" ascii character in a cell #234</a>. Thanks to <a href="https://github.com/NOtherDev">NOtherDev</a> for the contribution.</li></ul> |
| 0.3.0   | <ul><li>Addressed <a href="https://github.com/guyonroche/exceljs/issues/266">Breaking change removing bluebird #266</a>. Appologies for any inconvenience.</li><li>Added Promise library dependency injection. See <a href="#config">Config</a> section for more details.</li></ul> |
| 0.3.1   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/279">Update dependencies #279</a>. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/267">Minor fixes for stream handling #267</a>. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li><li>Added automated tests in phantomjs for the browserified code.</li></ul> |
| 0.4.0   | <ul><li>Fixed issue <a href="https://github.com/guyonroche/exceljs/issues/278">Boolean cell with value ="true" is returned as 1 #278</a>. The fix involved adding two new Call Value types:<ul><li><a href="#boolean-value">Boolean Value</a></li><li><a href="#error-value">Error Value</a></li></ul>Note: Minor version has been bumped up to 4 as this release introduces a couple of interface changes:<ul><li>Boolean cells previously will have returned 1 or 0 will now return true or false</li><li>Error cells that previously returned a string value will now return an error structure</li></ul></li><li>Fixed issue <a href="https://github.com/guyonroche/exceljs/issues/280">Code correctness - setters don't return a value #280</a>.</li><li>Addressed issue <a href="https://github.com/guyonroche/exceljs/issues/288">v0.3.1 breaks meteor build #288</a>.</li></ul> |
| 0.4.1   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/285">Add support for cp:contentStatus #285</a>. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/286">Fix Valid characters in XML (allow \n and \r when saving) #286</a>. Thanks to <a href="https://github.com/Rycochet">Rycochet</a> for the contribution.</li><li>Fixed <a href="https://github.com/guyonroche/exceljs/issues/275">hyperlink with query arguments corrupts workbook #275</a>. The hyperlink target is not escaped before serialising in the xml.</li></ul> |
| 0.4.2   | <ul><li><p>Addressed the following issues:<ul><li><a href="https://github.com/guyonroche/exceljs/issues/290">White text and borders being changed to black #290</a></li><li><a href="https://github.com/guyonroche/exceljs/issues/261">Losing formatting/pivot table from loaded file #261</a></li><li><a href="https://github.com/guyonroche/exceljs/issues/272">Solid fill become black #272</a></li></ul>These issues are potentially caused by a bug that caused colours with zero themes, tints or indexes to be rendered and parsed incorrectly.</p><p>Regarding themes: the theme files stored inside the xlsx container hold important information regarding colours, styles etc and if the theme information from a loaded xlsx file is lost, the results can be unpredictable and undesirable. To address this, when an ExcelJS Workbook parses an XLSX file, it will preserve any theme files it finds and include them when writing to a new XLSX. If this behaviour is not desired, the Workbook class exposes a clearThemes() function which will drop the theme content. Note that this behaviour is only implemented in the document based Workbook class, not the streamed Reader and Writer.</p></li></ul> |
| 0.4.3   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/294">Support error references in cell ranges #294</a>. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li></ul> |
| 0.4.4   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/297">Issue with copied cells #297</a>. This merge adds support for shared formulas. Thanks to <a href="https://github.com/muscapades">muscapades</a> for the contribution.</li></ul> |
| 0.4.6   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/304">Correct spelling #304</a>. Thanks to <a href="https://github.com/toanalien">toanalien</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/304">Added support for auto filters #306</a>. This adds <a href="#auto-filters">Auto Filters</a> to the Worksheet. Thanks to <a href="https://github.com/C4rmond4i">C4rmond4i</a> for the contribution.</li><li>Restored NodeJS 4.0.0 compatability by removing the destructuring code. My apologies for any inconvenience.</li></ul> |
| 0.4.9   | <ul><li>Switching to transpiled code for distribution. This will ensure compatability with 4.0.0 and above from here on. And it will also allow use of much more expressive JS code in the lib folder!</li><li><a href="#images">Basic Image Support!</a>Images can now be added to worksheets either as a tiled background or stretched over a range. Note: other features like rotation, etc. are not supported yet and will reqeuire further work.</li></ul> |
| 0.4.10  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/319">Add missing Office Rels #319</a>. Thanks goes to <a href="https://github.com/mauriciovillalobos">mauriciovillalobos</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/320">Add printTitlesRow Support #320</a> Thanks goes to <a href="https://github.com/psellers89">psellers89</a> for the contribution.</li></ul> |
| 0.4.11  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/327">Avoid error on anchor with no media #327</a>. Thanks goes to <a href="https://github.com/holm">holm</a> for the contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/332">Assortment of fixes for streaming read #332</a>. Thanks goes to <a href="https://github.com/holm">holm</a> for the contribution.</li></ul> |
| 0.4.12  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/334">Don’t set address if hyperlink r:id is undefined #334</a>. Thanks goes to <a href="https://github.com/holm">holm</a> for the contribution.</li></ul> |
| 0.4.13  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/343">Issue 296 #343</a>. This fixes <a href="https://github.com/guyonroche/exceljs/issues/296">Issue with writing newlines #296</a>. Thanks goes to <a href="https://github.com/holly-weisser">holly-weisser</a> for the contribution.</li></ul> |
| 0.4.14  | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/350">Syntax highlighting added ✨ #350</a>. Thanks goes to <a href="https://github.com/rmariuzzo">rmariuzzo</a> for the contribution.</li></ul> |
| 0.5.0   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/356">Fix right to left issues #356</a>. Fixes <a href="https://github.com/guyonroche/exceljs/issues/72">Add option to RTL file #72</a> and <a href="https://github.com/guyonroche/exceljs/issues/126">Adding an option to set RTL worksheet #126</a>. Big thank you to <a href="https://github.com/alitaheri">alitaheri</a> for this contribution.</li></ul> |
| 0.5.1   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/364">Fix #345 TypeError: Cannot read property 'date1904' of undefined #364</a>. This fixes <a href="https://github.com/guyonroche/exceljs/issues/345">TypeError: Cannot read property 'date1904' of undefined #345</a>. Thanks to <a href="https://github.com/Diluka">Diluka</a> for this contribution.</li></ul>
| 0.6.0   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/389">Add rowBreaks feature. #389</a>. Thanks to <a href="https://github.com/brucejo75">brucejo75</a> for this contribution.</li></ul> |
| 0.6.1   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/403">Guard null model fields - fix and tests #403</a>. Thanks to <a href="https://github.com/shdd-cjharries">thecjharries</a> for this contribution. Also thanks to <a href="https://github.com/Rycochet">Ryc O'Chet</a> for help with reviewing.</li></ul> |
| 0.6.2   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/396">Add some comments in readme according csv importing #396</a>. Thanks to <a href="https://github.com/Imperat">Michael Lelyakin</a> for this contribution. Also thanks to <a href="https://github.com/planemar">planemar</a> for help with reviewing. This also closes <a href="https://github.com/guyonroche/exceljs/issues/395">csv to stream doesn't work #395</a>.</li></ul> |
| 0.7.0   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/407">Impl &lt;xdr:twoCellAnchor editAs=oneCell&gt; #407</a>. Thanks to <a href="https://github.com/Ockejanssen">Ocke Janssen</a> and <a href="https://github.com/kay-ramme">Kay Ramme</a> for this contribution. This change allows control on how images are anchored to cells.</li></ul> |
| 0.7.1   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/423">Don't break when attempting to import a zip file that's not an Excel file (eg. .numbers) #423</a>. Thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. This change makes exceljs more reslilient when opening non-excel files.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/434">Fixes #419 : Updates readme. #434</a>. Thanks to <a href="https://github.com/getsomecoke">Vishnu Kyatannawar</a> for this contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/436">Don't break when docProps/core.xml contains a &lt;cp:version&gt; tag #436</a>. Thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. This change handles core.xml files with empty version tags.</li></ul>
| 0.8.0   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/442">Add Base64 Image support for the .addImage() method #442</a>. Thanks to <a href="https://github.com/jwmann">James W Mann</a> for this contribution.</li><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/453">update moment to 2.19.3 #453</a>. Thanks to <a href="https://github.com/cooltoast">Markan Patel</a> for this contribution.</li></ul> |
| 0.8.1   | <ul><li> Merged <a href="https://github.com/guyonroche/exceljs/pull/457">Additional information about font family property #457</a>. Thanks to <a href="https://github.com/kayakyakr">kayakyakr</a> for this contribution. </li> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/459">Fixes #458 #459</a>. This fixes <a href="https://github.com/guyonroche/exceljs/issues/458">Add style to column causes it to be hidden #458</a>. Thanks to <a href="https://github.com/AJamesPhillips">Alexander James Phillips</a> for this contribution. </li> </ul> |
| 0.8.2   | <ul><li>Merged <a href="https://github.com/guyonroche/exceljs/pull/466">Don't break when loading an Excel file containing a chartsheet #466</a>. Thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/471">Hotfix/sheet order#257 #471</a>. This fixes <a href="https://github.com/guyonroche/exceljs/issues/257">Sheet Order #257</a>. Thanks to <a href="https://github.com/robbi">Robbi</a> for this contribution. </li> </ul> |
| 0.8.3   | <ul> <li> Assimilated <a href="https://github.com/guyonroche/exceljs/pull/463">fix #79 outdated dependencies in unzip2</a>. Thanks to <a href="https://github.com/jsamr">Jules Sam. Randolph</a> for starting this fix and a really big thanks to <a href="https://github.com/kachkaev">Alexander Kachkaev</a> for finding the final solution. </li> </ul> |
| 0.8.4   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/479">Round Excel date to nearest millisecond when converting to javascript date #479</a>. Thanks to <a href="https://github.com/bjet007">Benoit Jean</a> for this contribution. </li> </ul> |
| 0.8.5   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/485">Bug fix: wb.worksheets/wb.eachSheet caused getWorksheet(0) to return sheet #485</a>. Thanks to <a href="https://github.com/mah110020">mah110020</a> for this contribution. </li> </ul> |
| 0.9.0   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/489">Feature/issue 424 #489</a>. This fixes <a href="https://github.com/guyonroche/exceljs/issues/424">No way to control summaryBelow or summaryRight #424</a>. Many thanks to <a href="https://github.com/sarahdmsi">Sarah</a> for this contribution. </li> </ul>  |
| 0.9.1   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/490">add type definition #490</a>. This adds type definitions to ExcelJS! Many thanks to <a href="https://github.com/taoqf">taoqf</a> for this contribution. </li> </ul> |
| 1.0.0   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/494">Add Node 8 and Node 9 to continuous integration testing #494</a>. Many thanks to <a href="https://github.com/cooltoast">Markan Patel</a> for this contribution. </li> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/508">Small README fix #508</a>. Many thanks to <a href="https://github.com/lbguilherme">Guilherme Bernal</a> for this contribution. </li> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/501">Add support for inlineStr, including rich text #501</a>. Many thanks to <a href="https://github.com/linguamatics-pdenes">linguamatics-pdenes</a> and <a href="https://github.com/robscotts4rb">Rob Scott</a> for their efforts towards this contribution. Since this change is technically a breaking change (the rendered XML for inline strings will change) I'm making this a major release! </li> </ul> |
| 1.0.1   | <ul> <li> Fixed <a href="https://github.com/guyonroche/exceljs/issues/520">spliceColumns problem when the number of columns are important #520</a>. </li> </ul> |
| 1.0.2   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/524">Loosen exceljs's dependency requirements for moment #524</a>. Many thanks to <a href="https://github.com/nicoladefranceschi">nicoladefranceschi</a> for this contribution. This change addresses <a href="https://github.com/guyonroche/exceljs/issues/517">Ability to use external "moment" package #517</a>. </li> </ul> |
| 1.1.0   | <ul> <li> Addressed <a href="https://github.com/guyonroche/exceljs/issues/514">Is there a way inserting values in columns. #514</a>. Added a new getter/setter property to Column to get and set column values (see <a href="#columns">Columns</a> for details). </li> </ul> |
| 1.1.1   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/532">Include index.d.ts in published packages #532</a>. To fix <a href="https://github.com/guyonroche/exceljs/issues/525">TypeScript definitions missing from npm package #525</a>. Many thanks to <a href="https://github.com/saschanaz">Kagami Sascha Rosylight</a> for this contribution. </li> </ul> |
| 1.1.2   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/536">Don't break when docProps/core.xml contains <cp:contentType /> #536</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> (and reviewers) for this contribution. </li> </ul> |
| 1.1.3   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/537">Try to handle the case where a &lt;c&gt; element is missing an r attribute #537</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> </ul> |
| 1.2.0   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/544">Add dateUTC flag to CSV Writing #544</a>. Many thanks to <a href="https://github.com/zgriesinger">Zackery Griesinger</a> for this contribution. </li> </ul> |
| 1.2.1   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/547">worksheet name is writable #547</a>. Many thanks to <a href="https://github.com/f111fei">xzper</a> for this contribution. </li> </ul> |
| 1.3.0   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/549">Add CSV write buffer support #549</a>. Many thanks to <a href="https://github.com/jloveridge">Jarom Loveridge</a> for this contribution. </li> </ul> |
| 1.4.2   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/541">Discussion: Customizable row/cell limit #541</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> </ul> |
| 1.4.3   | <ul> <li> Merged <a href="https://github.com/guyonroche/exceljs/pull/552">Get the right text out of hyperlinked formula cells #552</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> and <a href="https://github.com/holm">Christian Holm</a> for this contribution. </li> </ul> |
