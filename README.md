# ExcelJS

Read, manipulate and write spreadsheet data to XLSX and JSON.

Reverse engineered from Excel spreadsheet files as a project.

# Installation

npm install exceljs

# New Features!

* [Number Formats](#number-formats)

# Contents

* [Interface](#interface)
    * [Create a Workbook](#create-a-workbook)
    * [Add a Worksheet](#add-a-worksheet)
    * [Access Worksheets](#access-worksheets)
    * [Fetching Sheets](#fetching-sheets)
    * [Column Headers](#column-headers)
    * [Adding and Handling Rows](#adding-and-handling-rows)
    * [Handling Individual Cells](#handling-individual-cells)
    * [Merged Cells](#merged-cells)
    * [Number Formats](#number-formats)
    * [Reading XLSX](#reading-xlsx)
    * [Writing XLSX](#writing-xlsx)
* [Value Types](#value-types)

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
workbook.worksheets.forEach(function(worksheet) {
    // ...
});
```

## Fetching Sheets

```javascript
// by name
var sheet = workbook.getWorksheet("My Sheet");

// by id
var sheet = workbook.getWorksheet(1);
```

## Column Headers

```javascript
// Add column headers and define column keys and widths
sheet.columns = [
    { header: "Id", key: "id", width: 10 },
    { header: "Name", key: "name", width: 32 },
    { header: "D.O.B.", key: "dob", width: 10 }
];
```

## Adding and Handling Rows

```javascript
// Add a couple of Rows by key-value (after the last current row)
sheet.addRow({id: 1, name: "John Doe", dob: new Date(1970,1,1)});
sheet.addRow({id: 2, name: "Jane Doe", dob: new Date(1965,1,7)});

// Add a row by contiguous Array (assign to columns A, B & C)
sheet.addRow([3, "Sam", new Date()]);

// Add a row by sparse Array (assign to columns A, E & I)
var row = [];
row[1] = 4;
row[5] = "Kyle";
row[9] = new Date();
sheet.addRow(row);

// get a row as a sparse array
row = sheet.getRow(4);
assert(row[5] == "Kyle");

// iterate over all rows (as sparse arrays)
sheet.eachRow(function(number, row) {
    console.log("Row " + number + " = " + JSON.stringify(row));
});

## Handling Individual Cells

```javascript
// Modify/Add individual cell
sheet.getCell("C3").value = new Date(1968, 5, 1);

// query a cell's type
assert(sheet.getCell("C3").type == Excel.ValueType.Date);
```

## Merged Cells

```javascript
// merge a range of cells
sheet.mergeCells("A4:B5");

// ... merged cells are linked
sheet.getCell("B5").value = "Hello, World!";
assert(sheet.getCell("A4").value === sheet.getCell("B5").value);
assert(sheet.getCell("A4") === sheet.getCell("B5").master);
```

## Number Formats

* Note: This feature currently only implemented for write(). Coming soon - read().

```javascript
// display value as "1 3/5"
ws.getCell("A1").value = 1.6;
ws.getCell("A1").numFmt = "# ?/?";

// display value as "1.60%"
ws.getCell("B1").value = 1.6;
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

