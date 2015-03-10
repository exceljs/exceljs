# ExcelJS

Read, manipulate and write spreadsheet data and styles to XLSX and JSON.

Reverse engineered from Excel spreadsheet files as a project.

# Installation

npm install exceljs

# New Features!

<ul>
    <li>
        Bug Fixes
        <ul>
            <li>Handles File Not Found error</li>
        </ul>
    </li>
    <li><a href="#csv">CSV Files</a></li>
</ul>

# Coming Soon

<ul>
    <li>Column and Row Styles</li>
    <li>XLSX Streaming Writer</li>
    <li>XLSX Streaming Parser</li>
</ul>

# Contents

<ul>
    <li>
        <a href="#interface">Interface</a>
        <ul>
            <li><a href="#create-a-workbook">Create a Workbook</a></li>
            <li><a href="#add-a-worksheet">Add a Worksheet</a></li>
            <li><a href="#access-worksheets">Access Worksheets</a></li>
            <li><a href="#columns">Columns</a></li>
            <li><a href="#rows">Rows</a></li>
            <li><a href="#handling-individual-cells">Handling Individual Cells</a></li>
            <li><a href="#merged-cells">Merged Cells</a></li>
            <li><a href="#cell-styles">Cell Styles</a>
                <ul>
                    <li><a href="#number-formats">Number Formats</a></li>
                    <li><a href="#fonts">Fonts</a></li>
                    <li><a href="#alignment">Alignment</a></li>
                    <li><a href="#fills">Fills</a></li>
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
    <li><a href="#release-history">Release History</a></li>
    <li><a href="#known-issues">Known Issues</a></li>
</ul>

# Interface

```javascript
var Excel = require("exceljs");
```

## Create a Workbook

```javascript
var workbook = new Excel.Workbook();
```

## Set Workbook Properties

```javascript
workbook.creator = "Me";
workbook.lastModifiedBy = "Her";
workbook.created = new Date(1985, 8, 30);
workbook.modified = new Date();
```

## Add a Worksheet

```javascript
var sheet = workbook.addWorksheet("My Sheet");
```

## Access Worksheets
```javascript
// Iterate over all sheets
// Note: workbook.worksheets.forEach will still work but this is better
workbook.eachSheet(function(worksheet, sheetId) {
    // ...
});

// fetch sheet by name
var worksheet = workbook.getWorksheet("My Sheet");

// fetch sheet by id
var worksheet = workbook.getWorksheet(1);
```

## Columns

```javascript
// Add column headers and define column keys and widths
// Note: these column structures are a workbook-building convenience only,
// apart from the column width, they will not be fully persisted.
worksheet.columns = [
    { header: "Id", key: "id", width: 10 },
    { header: "Name", key: "name", width: 32 },
    { header: "D.O.B.", key: "DOB", width: 10 }
];

// Access an individual columns by key, letter and 1-based column number
var idCol = worksheet.getColumn("id");
var nameCol = worksheet.getColumn("B");
var dobCol = worksheet.getColumn(3);
    
// set column properties

// Note: will overwrite cell value C1
dobCol.header = "Date of Birth";

// Note: this will overwrite cell values C1:C2
dobCol.header = ["Date of Birth", "A.K.A. D.O.B."];

// from this point on, this column will be indexed by "dob" and not "DOB"
dobCol.key = "dob";

dobCol.width = 15;
```

## Rows

```javascript
// Add a couple of Rows by key-value, after the last current row, using the column keys
worksheet.addRow({id: 1, name: "John Doe", dob: new Date(1970,1,1)});
worksheet.addRow({id: 2, name: "Jane Doe", dob: new Date(1965,1,7)});

// Add a row by contiguous Array (assign to columns A, B & C)
worksheet.addRow([3, "Sam", new Date()]);

// Add a row by sparse Array (assign to columns A, E & I)
var rowValues = [];
rowValues[1] = 4;
rowValues[5] = "Kyle";
rowValues[9] = new Date();
worksheet.addRow(rowValues);

// Get a row object. If it doesn't already exist, a new empty one will be returned
var row = worksheet.getRow(5);

// Set a specific row height
row.height = 42.5;

row.getCell(1).value = 5; // A5's value set to 5
row.getCell("name").value = "Zeb"; // B5's value set to "Zeb" - assuming column 2 is still keyed by name
row.getCell("C").value = new Date(); // C5's value set to now

// Get a row as a sparse array
// Note: interface change: worksheet.getRow(4) ==> worksheet.getRow(4).values
row = worksheet.getRow(4).values;
expect(row[5]).toEqual("Kyle");

// assign row values by contiguous array (where array element 0 has a value)
row.values = [1,2,3];
expect(row.getCell(1).value).toEqual(1);
expect(row.getCell(2).value).toEqual(2);
expect(row.getCell(3).value).toEqual(3);

// assign row values by sparse array  (where array element 0 is undefined)
var values = []
values[5] = 7;
values[10] = "Hello, World!";
row.values = values;
expect(row.getCell(1).value).toBeNull();
expect(row.getCell(5).value).toEqual(7);
expect(row.getCell(10).value).toEqual("Hello, World!");

// assign row values by object, using column keys
row.values = {
    id: 13,
    name: "Thing 1",
    dob: new Date()
};

// Iterate over all rows that have values in a worksheet
// Note: interface change - argument order is now row, rowNumber
worksheet.eachRow(function(row, rowNumber) {
    console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
});

// Iterate over all non-null cells in a row
row.eachCell(function(cell, colNumber) {
    console.log("Cell " + colNumber + " = " + cell.value);
});
```

## Handling Individual Cells

```javascript
// Modify/Add individual cell
worksheet.getCell("C3").value = new Date(1968, 5, 1);

// query a cell's type
expect(worksheet.getCell("C3").type).toEqual(Excel.ValueType.Date);
```

## Merged Cells

```javascript
// merge a range of cells
worksheet.mergeCells("A4:B5");

// merge by top-left, bottom-right
worksheet.mergeCells("G10", "H11");
worksheet.mergeCells(10,11,12,13); // top,left,bottom,right

// ... merged cells are linked
worksheet.getCell("B5").value = "Hello, World!";
expect(worksheet.getCell("A4").value).toBe(worksheet.getCell("B5").value);
expect(worksheet.getCell("A4")).toBe(worksheet.getCell("B5").master);
```

## Cell Styles

### Number Formats

```javascript
// display value as "1 3/5"
ws.getCell("A1").value = 1.6;
ws.getCell("A1").numFmt = "# ?/?";

// display value as "1.60%"
ws.getCell("B1").value = 0.016;
ws.getCell("B1").numFmt = "0.00%";
```

### Fonts

```javascript

// for the wannabe graphic designers out there
ws.getCell("A1").font = {
    name: "Comic Sans MS",
    family: 4,
    size: 16,
    underline: true,
    bold: true
};

// for the graduate graphic designers...
ws.getCell("A2").font = {
    name: "Arial Black",
    color: { argb: "FF00FF00" },
    family: 2,
    size: 14,
    italic: true
};

// note: the cell will store a reference to the font object assigned.
// If the font object is changed afterwards, the cell font will change also...
var font = { name: "Arial", size: 12 };
ws.getCell("A3").font = font;
font.size = 20; // Cell A3 now has font size 20!

// Cells that share similar fonts may reference the same font object after
// the workbook is read from file or stream

```

| Font Property             | Description       | Example Value(s) |
| ------------------------- | ----------------- | ---------------- |
| name | Font name. | "Arial", "Calibri", etc. |
| family | Font family. An integer value. | 1,2,3, etc. |
| scheme | Font scheme. | "minor", "major", "none" |
| charset | Font charset. An integer value. | 1, 2, etc. |
| color | Colour description, an object containing an ARGB value. | { argb: "FFFF0000"} |
| bold | Font **weight** | true, false |
| italic | Font *slope* | true, false |
| underline | Font <u>underline</u> style | true, false, "none", "single", "double", "singleAccounting", "doubleAccounting" |
| strike | Font <strike>strikethrough</strike> | true, false |
| outline | Font outline | true, false |

### Alignment

```javascript
// set cell alignment to top-left, middle-center, bottom-right
ws.getCell("A1").alignment = { vertical: "top", horizontal: "left" };
ws.getCell("B1").alignment = { vertical: "middle", horizontal: "center" };
ws.getCell("C1").alignment = { vertical: "bottom", horizontal: "right" };

// set cell to wrap-text
ws.getCell("D1").alignment = { wrapText: true };

// set cell indent to 1
ws.getCell("E1").alignment = { indent: 1 };

// set cell text rotation to 30deg upwards, 45deg downwards and vertical text
ws.getCell("F1").alignment = { textRotation: 30 };
ws.getCell("G1").alignment = { textRotation: -45 };
ws.getCell("H1").alignment = { textRotation: "vertical" };

```

**Valid Alignment Property Values**

| horizontal | vertical    | wrapText | indent  | readingOrder | textRotation |
| ---------- | ----------- | -------- | ------- | ------------ | ------------ |
| left       | top         | true     | integer | rtl          | 0 to 90      |
| center     | middle      | false    |         | ltr          | -1 to -90    |
| right      | bottom      |          |         |              | vertical     |
| fill       | distributed |          |         |              |              |
| justify    | justify     |          |         |              |              |
| centerContinuous |       |          |         |              |              |
| distributed |            |          |         |              |              |


### Borders

```javascript
// set single thin border around A1
ws.getCell("A1").border = {
    top: {style:"thin"},
    left: {style:"thin"},
    bottom: {style:"thin"},
    right: {style:"thin"}
};

// set double thin green border around A3
ws.getCell("A3").border = {
    top: {style:"double", color: {argb:"FF00FF00"}},
    left: {style:"double", color: {argb:"FF00FF00"}},
    bottom: {style:"double", color: {argb:"FF00FF00"}},
    right: {style:"double", color: {argb:"FF00FF00"}}
};

// set thick red cross in A5
ws.getCell("A5").border = {
    diagonal: {up: true, down: true, style:"thick", color: {argb:"FFFF0000"}}
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
ws.getCell("A1").fill = {
    type: "pattern",
    pattern:"darkVertical",
    fgColor:{argb:"FFFF0000"}
};

// fill A2 with yellow dark trellis and blue behind
ws.getCell("A2").fill = {
    type: "pattern",
    pattern:"darkTrellis",
    fgColor:{argb:"FFFFFF00"},
    bgColor:{argb:"FF0000FF"}
};

// fill A3 with blue-white-blue gradient from left to right
ws.getCell("A3").fill = {
    type: "gradient",
    gradient: "angle",
    degree: 0,
    stops: [
        {position:0, color:{argb:"FF0000FF"}},
        {position:0.5, color:{argb:"FFFFFFFF"}},
        {position:1, color:{argb:"FF0000FF"}}
    ]
};


// fill A4 with red-green gradient from center
ws.getCell("A2").fill = {
    type: "gradient",
    gradient: "path",
    center:{left:0.5,top:0.5},
    stops: [
        {position:0, color:{argb:"FFFF0000"}},
        {position:1, color:{argb:"FF00FF00"}}
    ]
};

```

#### Pattern Fills

| Property | Required | Description |
| -------- | -------- | ----------- |
| type     | Y        | Value: "pattern"<br/>Specifies this fill uses patterns |
| pattern  | Y        | Specifies type of pattern (see <a href="#valid-pattern-types">Valid Pattern Types</a> below) |
| fgColor  | N        | Specifies the pattern foreground color. Default is black. |
| bgColor  | N        | Specifies the pattern background color. Default is white. |

##### Valid Pattern Types

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
| type     | Y        | Value: "gradient"<br/>Specifies this fill uses gradients |
| gradient | Y        | Specifies gradient type. One of ["angle", "path"] |
| degree   | angle    | For "angle" gradient, specifies the direction of the gradient. 0 is from the left to the right. Values from 1 - 359 rotates the direction clockwise |
| center   | path     | For "path" gradient. Specifies the relative coordinates for the start of the path. "left" and "top" values range from 0 to 1 |
| stops    | Y        | Specifies the gradient colour sequence. Is an array of objects containing position and color starting with position 0 and ending with position 1. Intermediatary positions may be used to specify other colours on the path. |

**Caveats**
Using the interface above it may be possible to create gradient fill effects not possible using the XLSX editor program.
For example, Excel only supports angle gradients of 0, 45, 90 and 135.
Similarly the sequence of stops may also be limited by the UI with positions [0,1] or [0,0.5,1] as the only options.
Take care with this fill to be sure it is supported by the target XLSX viewers.

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
    dateFormats: ["DD/MM/YYYY"]
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
* "MM-DD-YYYY"
* "YYYY-MM-DD"

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
    dateFormat: "DD/MM/YYYY HH:mm:ss"
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
                return moment(value).format("YYYY-MM-DD");
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

# Value Types

The following value types are supported.

| Enum Name                 | Enum(*)   | Description       | Example Value |
| ------------------------- | --------- | ----------------- | ------------- |
| Excel.ValueType.Null      | 0         | No value.         | null |
| Excel.ValueType.Merge     | 1         | N/A               | N/A |
| Excel.ValueType.Number    | 2         | A numerical value | 3.14 |
| Excel.ValueType.String    | 3         | A text value      | "Hello, World!" |
| Excel.ValueType.Date      | 4         | A Date value      | new Date()  |
| Excel.ValueType.Hyperlink | 5         | A hyperlink       | { text: "www.mylink.com", hyperlink: "http://www.mylink.com" } |
| Excel.ValueType.Formula   | 6         | A formula         | { formula: "A1+A2", result: 7 } |

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

