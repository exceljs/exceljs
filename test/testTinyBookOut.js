var fs = require('fs');

var HrStopwatch = require('./utils/hr-stopwatch');
var Workbook = require('../lib/doc/workbook');

var filename = process.argv[2];

var wb = new Workbook();
var ws = wb.addWorksheet('blort');

ws.getCell('A1').value = 7;
ws.getCell('B1').value = 'Hello, World!';

var stopwatch = new HrStopwatch();
stopwatch.start();
wb.xlsx.writeFile(filename)
  .then(function() {
    var micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros)
  });
// .catch(function(error) {
//    console.log(error.message);
// })
