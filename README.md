# ExcelJS

Read, manipulate and write spreadsheet data to XLSX and JSON.

Reverse engineered from Excel spreadsheet files as a project.

# Installation

npm install exceljs

# New Features!

<ul>
    <li>Bug Fixes
        <ul><li>"&lt;" and "&gt;" text characters properly rendered in xlsx</li></ul>
    </li>
    <li><a href="#columns">Better Column control</a></li>
    <li><a href="#rows">Better Row control</a></li>
    <li><a href="#interface-changes">Some breaking interface changes</a> were required in the version. Follow link to learn more.</li>
</ul>

# Coming Soon

<ul>
    <li>Column and Row Styles</li>
    <li>Fonts</li>
    <li>Fills</li>
    <li>Borders</li>
</ul>

# Contents

<ul>
    <li>
        <a href="#interface">Interface</a>
        <ul>
            <li><a href="#create-a-workbook">Create a Workbook</a>
            <li><a href="#add-a-worksheet">Add a Worksheet</a>
            <li><a href="#access-worksheets">Access Worksheets</a>
            <li><a href="#columns">Columns</a>
            <li><a href="#rows">Rows</a>
            <li><a href="#handling-individual-cells">Handling Individual Cells</a>
            <li><a href="#merged-cells">Merged Cells</a>
            <li><a href="#number-formats">Number Formats</a>
            <li><a href="#reading-xlsx">Reading XLSX</a>
            <li><a href="#writing-xlsx">Writing XLSX</a>
        </ul>
    </li>
    <li><a href="#value-types">Value Types</a></li>
    <li><a href="#release-history">Release History</a></li>
</ul>

# Interface

```javascript
var Excel = require("exceljs");
```

## Create a Workbook

```javascript
var workbook = new Excel.Workbook();
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
dobCol.header = "Date of Birth"; // Note: will overwrite cell value C1
dobCol.header = ["Date of Birth", "A.K.A. D.O.B."]; // Note: this will overwrite cell values C1:C2
dobCol.key = "dob"; // from this point on, this column will be indexed by "dob" and not "DOB"
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
// Note: interface change - this used to return a sparse array, now returns row object
var row = worksheet.getRow(5);

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

## Number Formats

```javascript
// display value as "1 3/5"
ws.getCell("A1").value = 1.6;
ws.getCell("A1").numFmt = "# ?/?";

// display value as "1.60%"
ws.getCell("B1").value = 0.016;
ws.getCell("B1").numFmt = "0.00%";
```

## Reading XLSX

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

## Writing XLSX

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

# Release History

| Version | Changes |
| ------- | ------- |
| 0.0.9 | <ul><li><a href="#number-formats">Number Formats</a></li></ul> |
| 0.1.0 | <ul><li>Bug Fixes<ul><li>"&lt;" and "&gt;" text characters properly rendered in xlsx</li></ul></li><li><a href="#columns">Better Column control</a></li><li><a href="#rows">Better Row control</a></li></ul> |


# Interface Changes

Every effort is made to make a good consistent interface that doesn't break through the versions but regrettably, now and then some things have to change for the greater good.

## Interface Breaks in 0.1.0

### Worksheet.eachRow

The arguments in the callback function to Worksheet.eachRow have been swapped and changed; it was function(rowNumber,rowValues), now it is function(row, rowNumber) which gives it a look and feel more like the underscore (_.each) function and prioritises the row object over the row number. 

### Worksheet.getRow

This function has changed from returning a sparse array of cell values to returning a Row object. This enables accessing row properties and will facilitate managing row styles and so on.

The sparse array of cell values is still available via Worksheet.getRow(rowNumber).values;
