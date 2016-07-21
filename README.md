# ExcelJS

Read, manipulate and write spreadsheet data and styles to XLSX and JSON.

Reverse engineered from Excel spreadsheet files as a project.

# Installation

npm install exceljs

# New Features!

<ul>
    <li><a href="#outline-level">Outline Levels</a> thanks to <a href="https://github.com/cricri">cricri</a> for the contribution.</li>
    <li><a href="#worksheet-properties">Worksheet Properties</a></li>
    <li>Further refactoring of worksheet writer.</li>
</ul>

# Backlog

<ul>
    <li>With the refactoring done, the priority for me now is to work through the pull requests and issues to bring them down to a more respectable level.</li>
    <li>XLSX Streaming Reader.</li>
    <li>ES6ify - This module was originally built for NodeJS 0.12.4 but things have moved on since then and I really want to start taking advantage of the modern JS features.
        I would also like to take the time to look at transpilers to support the earlier JSs</li>
    <li>Parsing CSV with Headers</li>
</ul>

# Contents

<ul>
    <li>
        <a href="#interface">Interface</a>
        <ul>
            <li><a href="#create-a-workbook">Create a Workbook</a></li>
            <li><a href="#set-workbook-properties">Set Workbook Properties</a></li>
            <li><a href="#add-a-worksheet">Add a Worksheet</a></li>
            <li><a href="#access-worksheets">Access Worksheets</a></li>
            <li><a href="#columns">Columns</a></li>
            <li><a href="#rows">Rows</a></li>
            <li><a href="#handling-individual-cells">Handling Individual Cells</a></li>
            <li><a href="#merged-cells">Merged Cells</a></li>
            <li><a href="#styles">Styles</a>
                <ul>
                    <li><a href="#number-formats">Number Formats</a></li>
                    <li><a href="#fonts">Fonts</a></li>
                    <li><a href="#alignment">Alignment</a></li>
                    <li><a href="#borders">Borders</a></li>
                    <li><a href="#fills">Fills</a></li>
                    <li><a href="rich-text">Rich Text</a></li>
                </ul>
            </li>
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
                </ul>
            </li>
        </ul>
    </li>
    <li><a href="#value-types">Value Types</a></li>
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
```

## Add a Worksheet

```javascript
var sheet = workbook.addWorksheet('My Sheet');
```

Use the second parameter of the addWorksheet function to create a new sheet with a specific tab color.
To add a new one with a red tab color use this example:

```javascript
var sheet = workbook.addWorksheet('My Sheet', 'FFC0000');
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

| Name             | Default    | Description |
| ---------------- | ---------- | ----------- |
| tabColor         | undefined  | Color of the tabs |
| outlineLevelCol  | 0          | The worksheet column outline level |
| outlineLevelRow  | 0          | The worksheet row outline level |
| defaultRowHeight | 15         | Default row height |
| dyDescent        | 55         | TBD |

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

// Commit a completed row to stream
row.commit();
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
| ---------- | ---------- |
| list       | Define a discrete set of valid values. Excel will offer these in a dropdown for easy entry |
| whole      | The value must be a whole number |
| decimal    | The value must be a decimal number |
| textLength | The value may be text but the length is controlled |
| custom     | A custom formula controls the valid values |

For types other than list or custom, the following operators affect the validation:
| Operator              | Description |
| --------------------  | ---------- |
| between               | Values must lie between formula results |
| notBetween            | Values must not lie between formula results |
| equal                 | Value must equal formula result |
| notEqual              | Value must not equal formula result |
| greaterThan           | Value must be greater than formula result |
| lessThan              | Value must be less than formula result |
| greaterThanOrEqual    | Value must be greater than or equal to formula result |
| lessThanOrEqual       | Value must be less than or equal to formula result |

```javascript
// Specify list of valid values (One, Two, Three, Four). Excel will provide a dropdown with these values.
worksheet.getCell('A1').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"One,Two,Three,Four"']
};

// Specify list of valid values from a range. Excel will provide a dropdown with these values.
    worksheet.getCell('A1').dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['$D$5:$F$5']
};

// Specify Cell must be a whole number that is not 5. Show the user an appropriate error message if they get it wrong
worksheet.getCell('A1').dataValidation = {
    type: 'whole',
    operator: 'notEqual',
    showErrorMessage: true,
    formulae: [5],
    errorStyle: 'error',
    errorTitle: 'Five',
    error: 'The value must not be Five'
};

// Specify Cell must be a decomal number between 1.5 and 7. Add 'tooltip' to help guid the user
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
ws.getColumn(3).numFmt = '�#,##0;[Red]-�#,##0';

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

| Font Property             | Description       | Example Value(s) |
| ------------------------- | ----------------- | ---------------- |
| name | Font name. | 'Arial', 'Calibri', etc. |
| family | Font family. An integer value. | 1,2,3, etc. |
| scheme | Font scheme. | 'minor', 'major', 'none' |
| charset | Font charset. An integer value. | 1, 2, etc. |
| color | Colour description, an object containing an ARGB value. | { argb: 'FFFF0000'} |
| bold | Font **weight** | true, false |
| italic | Font *slope* | true, false |
| underline | Font <u>underline</u> style | true, false, 'none', 'single', 'double', 'singleAccounting', 'doubleAccounting' |
| strike | Font <strike>strikethrough</strike> | true, false |
| outline | Font outline | true, false |

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
| stops    | Y        | Specifies the gradient colour sequence. Is an array of objects containing position and color starting with position 0 and ending with position 1. Intermediatary positions may be used to specify other colours on the path. |

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

Excel supports outlining; where rows or columns can be expanded or collapsed dependingn on what level of detail the user wishes to view.

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
workbook.csv.write(stream)
    .then(function() {
        // done
    });


// read from a file with European Date-Times
var workbook = new Excel.Workbook();
var options = {
    dateFormat: 'DD/MM/YYYY HH:mm:ss'
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

| Field | Description |
| ----- | ----------- |
| stream | Specifies a writable stream to write the XLSX workbook to. |
| filename | If stream not specified, this field specifies the path to a file to write the XLSX workbook to. |
| useSharedStrings | Specifies whether to use shared strings in the workbook. Default is false |
| useStyles | Specifies whether to add style information to the workbook. Styles can add some performance overhead. Default is false |

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

# Value Types

The following value types are supported.

| Enum Name                 | Enum(*)   | Description       | Example Value |
| ------------------------- | --------- | ----------------- | ------------- |
| Excel.ValueType.Null      | 0         | No value.         | null |
| Excel.ValueType.Merge     | 1         | N/A               | N/A |
| Excel.ValueType.Number    | 2         | A numerical value | 3.14 |
| Excel.ValueType.String    | 3         | A text value      | 'Hello, World!' |
| Excel.ValueType.Date      | 4         | A Date value      | new Date()  |
| Excel.ValueType.Hyperlink | 5         | A hyperlink       | { text: 'www.mylink.com', hyperlink: 'http://www.mylink.com' } |
| Excel.ValueType.Formula   | 6         | A formula         | { formula: 'A1+A2', result: 7 } |

# Interface Changes

Every effort is made to make a good consistent interface that doesn't break through the versions but regrettably, now and then some things have to change for the greater good.

## Interface Breaks in 0.1.0

### Worksheet.eachRow

The arguments in the callback function to Worksheet.eachRow have been swapped and changed; it was function(rowNumber,rowValues), now it is function(row, rowNumber) which gives it a look and feel more like the underscore (_.each) function and prioritises the row object over the row number.

### Worksheet.getRow

This function has changed from returning a sparse array of cell values to returning a Row object. This enables accessing row properties and will facilitate managing row styles and so on.

The sparse array of cell values is still available via Worksheet.getRow(rowNumber).values;

## Interface Breaks in 0.1.1

### cell.model

cell.styles renamed to cell.style

# Known Issues

## Too Many Worksheets Results in Parse Error

There appears to be an issue in one of the dependent libraries (unzip) where too many files causes the following error to be emitted:

```javascript
    invalid signature: 0x80014
```

In practical terms, this error only seems to arise with over 98 sheets (or 49 sheets with hyperlinks) so it shouldn't affect that many. I will keep an eye on it though.

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

