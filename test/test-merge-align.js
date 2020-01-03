var fs = require('fs');

var HrStopwatch = require('./utils/hr-stopwatch');

var Workbook = require('../lib/doc/workbook');

var filename = process.argv[2];

var wb = new Workbook();
var ws = wb.addWorksheet('blort');

ws.addRow([1,2,3,4]);
ws.addRow(['one', 'two', 'three', 'four']);
ws.addRow(['une', 'deux', 'trois', 'quatre']);
ws.addRow(['uno', 'due', 'tre', 'quatro']);

ws.mergeCells('B2:C3');
ws.getCell('B2').alignment = { horizontal: 'center', vertical: 'middle' };

var stopwatch = new HrStopwatch();
stopwatch.start();
wb.xlsx.writeFile(filename)
  .then(function() {
    var micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros)
  })
  .catch(function(error) {
     console.log(error.message);
  });
