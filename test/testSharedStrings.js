var fs = require('fs');
var _ = require('underscore');

var ExcelJS = require('../excel');

var Workbook = ExcelJS.Workbook;

var filename = process.argv[2];

var wb = new Workbook();
var ws = wb.addWorksheet('Sheet1');
ws.columns = [
  { header: "FirstName", key: "firstname" },
  { header: "LastName", key: "lastname" },
  { header: "Other Name", key: "othername" }
];


wb.xlsx.writeFile(filename)
  .then(() => { console.log('done'); })
  .catch((e) => console.log(e));
